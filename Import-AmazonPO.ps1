<#
.SYNOPSIS
    Creates OpenProd purchase orders from an Amazon Business reconciliation report.

.DESCRIPTION
    Reads the reconciliation report (.csv or .xlsx), groups the item rows by Amazon
    order, resolves the OpenProd ids it needs, then creates one purchase order per
    Amazon order with all its lines in a single API call.

    It talks to the Odoo-style "native" API of OpenProd. That API does NOT run the
    model onchanges on create, which is exactly what we want here: the Amazon unit
    price is written as-is instead of being recomputed from the product's purchase
    price. The flip side is that nothing is filled in automatically either, so every
    value this script does not send stays empty.

    Nothing is written without -Apply. The default run is a dry run: it performs the
    read-only lookups, prints exactly what would be created, and stops there.

.PARAMETER Path
    The Amazon Business reconciliation report, .csv or .xlsx. The .xlsx is read
    natively, Excel does not need to be installed.

.PARAMETER Apply
    Actually create the purchase orders. Without it the script only simulates.

.EXAMPLE
    .\Import-AmazonPO.ps1 -Path .\reconciliation_from_20260901_to_20260914.csv
    Dry run: shows what would be created.

.EXAMPLE
    .\Import-AmazonPO.ps1 -Path .\report.xlsx -Apply
    Creates the purchase orders.

.NOTES
    Only works from inside the corporate network: erp.local.wandercraft.eu resolves
    on the internal DNS only, and its certificate is issued by the internal IT CA.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string] $Path,

    [string] $Server         = 'https://erp.local.wandercraft.eu',
    [string] $Database       = 'Wandercraft_US',

    [string] $SupplierCode   = 'PA000016',
    [string] $ProductCode    = 'US - Office Expenses',
    [string] $AffairCode     = 'AF0006',
    [string] $UnitName       = 'Unit(s)',
    [string] $TaxDescription = '',

    [ValidateSet('Ref', 'FirstItem', 'Count')]
    [string] $HeaderDescription = 'Ref',
    [string] $SupplierName      = 'Amazon',

    # Reference of the purchase order used as a template for the supplier settings
    # (currency, addresses, payment term, invoicing method). Defaults to the most
    # recent order of that supplier.
    [string] $TemplateOrder,

    # How many recent purchase orders are read to find the template and to detect
    # the orders already imported. Raise it if an Amazon order older than that
    # window must still be recognised.
    [int] $ScanDepth = 1000,

    # Dumps an existing purchase order and its lines field by field, then stops.
    # Diagnostic only: nothing is written.
    [string] $InspectOrder,

    [System.Management.Automation.PSCredential] $Credential,

    # Process only the first N orders. Meant for a cautious first run: -Limit 1
    # creates a single purchase order that you can check before doing the batch.
    [int] $Limit = 0,

    # Connects, lists the mandatory fields of the two models, and stops.
    # Diagnostic only: nothing is written.
    [switch] $ShowRequired,

    [switch] $Apply,
    [switch] $ParseOnly,
    [switch] $SkipCertificateCheck
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

if ($SkipCertificateCheck) {
    # Only for a machine that does not carry the internal IT CA. A domain-joined
    # Wandercraft laptop trusts it already and never needs this.
    Add-Type -TypeDefinition @'
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class WdcTrustAll : ICertificatePolicy {
    public bool CheckValidationResult(ServicePoint sp, X509Certificate cert, WebRequest req, int problem) { return true; }
}
'@
    [Net.ServicePointManager]::CertificatePolicy = New-Object WdcTrustAll
}

#region ─── Reading the report ────────────────────────────────────────────────

# Header names are matched on a normalised form: Amazon shifts its columns between
# exports and some labels carry a stray trailing space ("Regulatory fees excl. tax ").
function ConvertTo-NormalKey([string] $s) {
    if ($null -eq $s) { return '' }
    $t = $s.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object Text.StringBuilder
    foreach ($c in $t.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($c) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void] $sb.Append($c)
        }
    }
    return ($sb.ToString().ToLowerInvariant() -replace '[^a-z0-9]', '')
}

# Excel column reference to a 0-based index: "A" -> 0, "AB" -> 27. Cells are omitted
# from the XML when empty, so values must be placed by reference and never by rank.
function Convert-ColumnRefToIndex([string] $ref) {
    $letters = ($ref -replace '[0-9]', '').ToUpperInvariant()
    $n = 0
    foreach ($c in $letters.ToCharArray()) { $n = $n * 26 + ([int][char]$c - 64) }
    return $n - 1
}

function Read-XlsxRows([string] $file) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead((Resolve-Path $file))
    try {
        $entryText = {
            param($name)
            $e = $zip.Entries | Where-Object { $_.FullName -eq $name }
            if (-not $e) { return $null }
            $sr = New-Object IO.StreamReader($e.Open(), [Text.Encoding]::UTF8)
            try { $sr.ReadToEnd() } finally { $sr.Close() }
        }

        # Resolve the first sheet through the workbook relationships rather than
        # assuming sheet1.xml: some exporters number their parts differently.
        [xml] $wbXml = & $entryText 'xl/workbook.xml'
        [xml] $relXml = & $entryText 'xl/_rels/workbook.xml.rels'
        $firstSheet = @($wbXml.workbook.sheets.sheet)[0]
        $rid = $firstSheet.id
        if (-not $rid) { $rid = $firstSheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships') }
        $target = (@($relXml.Relationships.Relationship) | Where-Object { $_.Id -eq $rid }).Target
        if (-not $target) { $target = 'worksheets/sheet1.xml' }
        $sheetPath = 'xl/' + ($target -replace '^/xl/', '' -replace '^/', '')

        $shared = @()
        $ssText = & $entryText 'xl/sharedStrings.xml'
        if ($ssText) {
            [xml] $ss = $ssText
            foreach ($si in @($ss.sst.si)) {
                # A cell can be split into several runs (<r><t>), concatenate them all.
                $shared += (($si.SelectNodes('.//*[local-name()="t"]') | ForEach-Object { $_.InnerText }) -join '')
            }
        }

        [xml] $sheet = & $entryText $sheetPath
        $rows = New-Object Collections.Generic.List[object]
        foreach ($row in @($sheet.worksheet.sheetData.row)) {
            $cells = @{}
            $maxIdx = -1
            foreach ($c in @($row.c)) {
                $idx = Convert-ColumnRefToIndex $c.r
                $val = $null
                switch ($c.t) {
                    's'         { $i = [int] $c.v; if ($i -lt $shared.Count) { $val = $shared[$i] } }
                    'inlineStr' { $val = ($c.SelectNodes('.//*[local-name()="t"]') | ForEach-Object { $_.InnerText }) -join '' }
                    'str'       { $val = $c.v }
                    default     { $val = $c.v }
                }
                $cells[$idx] = $val
                if ($idx -gt $maxIdx) { $maxIdx = $idx }
            }
            $arr = @()
            for ($i = 0; $i -le $maxIdx; $i++) { $arr += $(if ($cells.ContainsKey($i)) { $cells[$i] } else { '' }) }
            $rows.Add($arr)
        }
        return $rows
    }
    finally { $zip.Dispose() }
}

function Read-CsvRows([string] $file) {
    # Import-Csv handles RFC 4180 quoting, including the doubled quotes Amazon puts
    # inside item titles (3"" x 4""), which a naive split would cut through.
    $data = Import-Csv -Path $file
    if (-not $data) { return @() }
    $headers = $data[0].PSObject.Properties.Name
    $rows = New-Object Collections.Generic.List[object]
    $rows.Add(@($headers))
    foreach ($r in $data) { $rows.Add(@($headers | ForEach-Object { [string] $r.$_ })) }
    return $rows
}

function Read-ReportRows([string] $file) {
    if (-not (Test-Path -LiteralPath $file)) { throw "File not found: $file" }
    switch -Regex ([IO.Path]::GetExtension($file)) {
        '\.xlsx$' { return Read-XlsxRows $file }
        '\.csv$'  { return Read-CsvRows  $file }
        '\.xls$'  { throw 'The legacy .xls format is not supported. Open it in Excel and save it as .xlsx or .csv.' }
        default   { throw "Unsupported extension: $([IO.Path]::GetExtension($file)). Use .csv or .xlsx." }
    }
}

#endregion

#region ─── Parsing the Amazon rows ───────────────────────────────────────────

$AmzHeaders = [ordered]@{
    OrderDate = 'Order Date'
    OrderId   = 'Order ID'
    User      = 'Account User'
    TxType    = 'Transaction Type'
    Currency  = 'Currency'
    Qty       = 'Shipment Quantity'
    UnitPrice = 'Unit price excl. tax'
    TaxAmount = 'Total tax amount'
    NetTotal  = 'Net total'
    Asin      = 'ASIN'
    Title     = 'Title'
}
# Without an id, a date, a quantity, a price or a label no order line can be built:
# better to refuse the file than to push a silently truncated import.
$AmzRequired = @('OrderId', 'OrderDate', 'Qty', 'UnitPrice', 'Title')

function Resolve-AmzColumns([object[]] $headerRow) {
    $byName = @{}
    for ($i = 0; $i -lt $headerRow.Count; $i++) {
        $k = ConvertTo-NormalKey ([string] $headerRow[$i])
        if ($k -and -not $byName.ContainsKey($k)) { $byName[$k] = $i }
    }
    $cols = @{}
    $missing = @()
    foreach ($key in $AmzHeaders.Keys) {
        $k = ConvertTo-NormalKey $AmzHeaders[$key]
        if ($byName.ContainsKey($k)) { $cols[$key] = $byName[$k] } else { $cols[$key] = -1 }
        if ($cols[$key] -lt 0 -and $AmzRequired -contains $key) { $missing += $AmzHeaders[$key] }
    }
    if ($missing.Count) { throw "Columns missing from the report: $($missing -join ', ')" }
    return $cols
}

# Always invariant culture: on a French-locale machine [double]'79.00' would
# otherwise read as 7900.
function ConvertTo-Number($v) {
    $s = ([string] $v).Trim() -replace '\s', ''
    if (-not $s) { return $null }
    $s = $s -replace ',', '.'
    $out = 0.0
    if ([double]::TryParse($s, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref] $out)) { return $out }
    return $null
}

# Amazon Business US writes MM/DD/YYYY. An .xlsx may instead hold a real date cell,
# which arrives here as an Excel serial number.
function ConvertTo-OrderDate($v) {
    $s = ([string] $v).Trim()
    if (-not $s) { return $null }
    if ($s -match '^\d{1,2}[/\-.]\d{1,2}[/\-.]\d{4}$') {
        $p = $s -split '[/\-.]'
        $mon = [int] $p[0]; $day = [int] $p[1]; $year = [int] $p[2]
        if ($mon -gt 12 -and $day -le 12) { $t = $mon; $mon = $day; $day = $t }
        if ($mon -lt 1 -or $mon -gt 12 -or $day -lt 1 -or $day -gt 31) { return $null }
        return (Get-Date -Year $year -Month $mon -Day $day -Hour 0 -Minute 0 -Second 0)
    }
    $serial = ConvertTo-Number $s
    if ($null -ne $serial -and $serial -gt 1 -and $serial -lt 80000) {
        return [DateTime]::FromOADate($serial)
    }
    return $null
}

function Get-AmazonOrders([object[]] $rows) {
    $cols = Resolve-AmzColumns $rows[0]
    $get = {
        param($r, $key)
        $i = $cols[$key]
        if ($i -lt 0 -or $i -ge $r.Count) { return '' }
        return ([string] $r[$i]).Trim()
    }

    $orders = [ordered]@{}
    $skipped = New-Object Collections.Generic.List[object]

    for ($n = 1; $n -lt $rows.Count; $n++) {
        $r = $rows[$n]
        $rowNo = $n + 1
        $orderId = & $get $r 'OrderId'
        if (-not $orderId) { continue }

        # Refunds and rows without a quantity produce no purchase order line. They
        # are set aside and listed, never silently absorbed.
        $txType = & $get $r 'TxType'
        if ($txType -and $txType.ToLowerInvariant() -ne 'charge') {
            $skipped.Add([pscustomobject]@{ Row = $rowNo; OrderId = $orderId; Reason = "transaction type '$txType'" }); continue
        }
        $qty = ConvertTo-Number (& $get $r 'Qty')
        if ($null -eq $qty -or $qty -le 0) {
            $skipped.Add([pscustomobject]@{ Row = $rowNo; OrderId = $orderId; Reason = 'quantity missing or zero' }); continue
        }
        $price = ConvertTo-Number (& $get $r 'UnitPrice')
        if ($null -eq $price) {
            $skipped.Add([pscustomobject]@{ Row = $rowNo; OrderId = $orderId; Reason = 'unit price missing' }); continue
        }
        $date = ConvertTo-OrderDate (& $get $r 'OrderDate')
        if ($null -eq $date) {
            $skipped.Add([pscustomobject]@{ Row = $rowNo; OrderId = $orderId; Reason = "unreadable date '$(& $get $r 'OrderDate')'" }); continue
        }

        if (-not $orders.Contains($orderId)) {
            $orders[$orderId] = [pscustomobject]@{
                OrderId   = $orderId
                Date      = $date
                Purchaser = (& $get $r 'User')
                Currency  = $(if ((& $get $r 'Currency')) { & $get $r 'Currency' } else { 'USD' })
                Lines     = New-Object Collections.Generic.List[object]
            }
        }
        else {
            # Amazon dates the order, not the line. A gap means a recomposed export,
            # so the earliest is kept as it is the order date.
            if ($date -lt $orders[$orderId].Date) { $orders[$orderId].Date = $date }
        }

        $tax = ConvertTo-Number (& $get $r 'TaxAmount'); if ($null -eq $tax) { $tax = 0 }
        $net = ConvertTo-Number (& $get $r 'NetTotal');  if ($null -eq $net) { $net = 0 }
        $o = $orders[$orderId]
        $o.Lines.Add([pscustomobject]@{
            Sequence  = ($o.Lines.Count + 1) * 10
            Title     = (& $get $r 'Title')
            Asin      = (& $get $r 'Asin')
            Qty       = $qty
            UnitPrice = $price
            TaxAmount = $tax
            NetTotal  = $net
        })
    }

    return [pscustomobject]@{ Orders = @($orders.Values); Skipped = $skipped }
}

function Get-HeaderDescription($order) {
    switch ($HeaderDescription) {
        'FirstItem' { return $order.Lines[0].Title }
        'Count'     {
            if ($order.Lines.Count -eq 1) { return $order.Lines[0].Title }
            return "$SupplierName - $($order.Lines.Count) items"
        }
        default     { return '{0} - {1} - {2}' -f $order.Date.ToString('MM/dd/yyyy'), $SupplierName, $order.OrderId }
    }
}

#endregion

#region ─── OpenProd API ──────────────────────────────────────────────────────

# The session cookie is carried by hand in this container, shared by every call.
$script:Cookies = New-Object Net.CookieContainer

# HttpWebRequest rather than Invoke-RestMethod, and that is not a style choice:
# on a non-2xx status Invoke-RestMethod reads the body itself before throwing, so
# the stream reachable from the exception is already exhausted and comes back
# empty. Every failure would then read "500 Internal Server Error" and nothing
# else, while OpenProd puts its real message and its traceback in that very body.
function Invoke-OpenProd([string] $Route, [hashtable] $Params) {
    $json = @{ jsonrpc = '2.0'; method = 'call'; params = $Params } | ConvertTo-Json -Depth 20 -Compress
    $bytes = [Text.Encoding]::UTF8.GetBytes($json)

    $req = [Net.HttpWebRequest]::Create("$Server$Route")
    $req.Method = 'POST'
    $req.ContentType = 'application/json'
    $req.Accept = 'application/json'
    $req.Timeout = 180000
    $req.CookieContainer = $script:Cookies
    $req.ContentLength = $bytes.Length

    $rs = $req.GetRequestStream()
    try { $rs.Write($bytes, 0, $bytes.Length) } finally { $rs.Close() }

    $status = 0
    $text = ''
    try {
        $resp = $req.GetResponse()
    }
    catch [Net.WebException] {
        $resp = $_.Exception.Response
        if (-not $resp) { throw }          # network level: DNS, TLS, timeout
    }
    try {
        $status = [int] $resp.StatusCode
        $sr = New-Object IO.StreamReader($resp.GetResponseStream(), [Text.Encoding]::UTF8)
        try { $text = $sr.ReadToEnd() } finally { $sr.Close() }
    }
    finally { $resp.Close() }

    $parsed = $null
    try { $parsed = $text | ConvertFrom-Json } catch { }

    if (-not $parsed) {
        throw "OpenProd returned a non-JSON response on $Route (HTTP $status): $($text.Substring(0, [Math]::Min(500, $text.Length)))"
    }

    # JSON-RPC reports business errors inside the body, sometimes with a 200 and
    # sometimes with a 500. The status alone never says whether the call worked.
    if ($parsed.PSObject.Properties.Name -contains 'error' -and $parsed.error) {
        $msg = $parsed.error.message
        if ($parsed.error.data) {
            if ($parsed.error.data.message) {
                $msg = "$msg : $($parsed.error.data.message)"
            }
            elseif ($parsed.error.data.debug) {
                # The useful part of a Python traceback is its last lines.
                $tail = ($parsed.error.data.debug -split "`n" | Where-Object { $_.Trim() } | Select-Object -Last 5) -join ' | '
                $msg = "$msg : $tail"
            }
        }
        throw "OpenProd API error on $Route (HTTP $status) -> $msg"
    }
    return $parsed.result
}

function Connect-OpenProd([System.Management.Automation.PSCredential] $Cred) {
    $script:Cookies = New-Object Net.CookieContainer
    $plain = $Cred.GetNetworkCredential().Password
    $res = Invoke-OpenProd '/web/session/authenticate' @{
        db = $Database; login = $Cred.UserName; password = $plain
    }
    if (-not $res.uid) { throw "Authentication refused for '$($Cred.UserName)' on database '$Database'." }
    return $res
}

function Search-OpenProd([string] $Model, [array] $Domain, [array] $Fields, $Limit = 1, $Sort = $null) {
    $res = Invoke-OpenProd '/web/dataset/search_read' @{
        model = $Model; domain = $Domain; fields = $Fields; limit = $Limit; sort = $Sort
    }
    return @($res.records)
}

# search_read returns a many2one as [id, "label"]. Only the id can be written back.
function ConvertTo-Writable($v) {
    if ($null -eq $v) { return $null }
    if ($v -is [bool] -and -not $v) { return $null }
    if ($v -is [array] -or $v -is [Collections.IList]) {
        if ($v.Count -ge 1) { return $v[0] }
        return $null
    }
    return $v
}

# Relations are read from the model itself instead of being hard-coded: it is the
# only way to know which model 'affair_id' or 'purchaser_id' actually point to.
$script:Relations = @{}
function Get-Fields([string] $Model) {
    if (-not $script:Relations.ContainsKey($Model)) {
        $script:Relations[$Model] = Invoke-OpenProd '/web/dataset/call_kw' @{
            model = $Model; method = 'fields_get'; args = @(); kwargs = @{}
        }
    }
    return $script:Relations[$Model]
}

function Get-Relation([string] $Model, [string] $Field) {
    $key = "$Model.$Field"
    if ($script:Relations.ContainsKey($key)) { return $script:Relations[$key] }
    $meta = (Get-Fields $Model).$Field
    if (-not $meta) { throw "Field '$Field' does not exist on model '$Model'." }
    $script:Relations[$key] = $meta.relation
    return $meta.relation
}

# The native API does not run the onchanges, which is what keeps the Amazon price
# intact, but it also means nothing is prefilled: company, currency, type,
# location and the rest stay empty and the record fails its NOT NULL constraints.
# default_get returns exactly what the web client puts in a blank form, so the
# record starts from the same base as a manual creation.
function Get-ModelDefaults([string] $Model) {
    $meta = Get-Fields $Model
    # Scalars only: default_get on a x2many hands back command lists that would
    # fight with the lines built here.
    $names = @($meta.PSObject.Properties |
        Where-Object { @('one2many', 'many2many') -notcontains $_.Value.type -and -not $_.Value.readonly } |
        ForEach-Object { $_.Name })
    $def = Invoke-OpenProd '/web/dataset/call_kw' @{
        model = $Model; method = 'default_get'; args = @(, $names); kwargs = @{}
    }
    $out = [ordered]@{}
    if ($def) {
        foreach ($p in $def.PSObject.Properties) {
            $v = $p.Value
            if ($null -eq $v) { continue }
            if ($v -is [bool] -and -not $v) { continue }   # Odoo writes an unset value as false
            $out[$p.Name] = $v
        }
    }
    return $out
}

# Lists the mandatory fields the payload leaves unset. Odoo only ever says "a
# mandatory field is not correctly set" without naming it, so this does the naming.
function Get-MissingMandatory([string] $Model, $Values) {
    $missing = @()
    foreach ($p in (Get-Fields $Model).PSObject.Properties) {
        if ($p.Value.required -and -not $Values.Contains($p.Name)) {
            $missing += "$($p.Name) [$($p.Value.type)] $($p.Value.string)"
        }
    }
    return $missing
}

function Resolve-Id([string] $Model, [string] $Field, [string] $Value, [string] $Label) {
    if (-not $Value) { return $null }
    $hits = Search-OpenProd $Model @(, @($Field, '=', $Value)) @('id', 'name') 2
    if ($hits.Count -eq 0) { throw "$Label : nothing found with $Field = '$Value' on $Model." }
    if ($hits.Count -gt 1) { throw "$Label : '$Value' is ambiguous on $Model, $($hits.Count) records match." }
    return [int] $hits[0].id
}

#endregion

#region ─── Run ───────────────────────────────────────────────────────────────

Write-Host ''
Write-Host "Amazon -> OpenProd purchase orders" -ForegroundColor Cyan
Write-Host ("  report   : {0}" -f (Split-Path -Leaf $Path))
Write-Host ("  server   : {0}   database: {1}" -f $Server, $Database)
Write-Host ("  mode     : {0}" -f $(if ($Apply) { 'APPLY - purchase orders will be created' } else { 'DRY RUN - nothing will be written' })) `
    -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'Green' })
Write-Host ''

$rows = Read-ReportRows $Path
if ($rows.Count -lt 2) { throw 'The report has no data row.' }
$parsed = Get-AmazonOrders $rows
$orders = $parsed.Orders
if (-not $orders.Count) { throw 'No usable order in this report.' }

$htTotal = 0.0; $ttcTotal = 0.0; $lineCount = 0
foreach ($o in $orders) {
    foreach ($l in $o.Lines) { $htTotal += $l.Qty * $l.UnitPrice; $ttcTotal += $l.NetTotal; $lineCount++ }
}
Write-Host ("Read {0} row(s): {1} order(s), {2} line(s), {3:N2} excl. tax, {4:N2} incl. tax, {5} row(s) set aside." -f `
    ($rows.Count - 1), $orders.Count, $lineCount, $htTotal, $ttcTotal, $parsed.Skipped.Count)
foreach ($s in $parsed.Skipped) { Write-Host ("  set aside  row {0}  {1}  {2}" -f $s.Row, $s.OrderId, $s.Reason) -ForegroundColor DarkYellow }

if ($Limit -gt 0 -and $Limit -lt $orders.Count) {
    Write-Host ("Limited to the first {0} order(s) of {1} by -Limit." -f $Limit, $orders.Count) -ForegroundColor Yellow
    $orders = @($orders | Select-Object -First $Limit)
}
Write-Host ''

if ($ParseOnly) {
    # Reading the report is checked on its own, without ever contacting OpenProd.
    $orders | ForEach-Object {
        [pscustomobject]@{
            Order       = $_.OrderId
            Date        = $_.Date.ToString('yyyy-MM-dd')
            Purchaser   = $_.Purchaser
            Lines       = $_.Lines.Count
            Amount      = [math]::Round((($_.Lines | ForEach-Object { $_.Qty * $_.UnitPrice } | Measure-Object -Sum).Sum), 2)
            Description = Get-HeaderDescription $_
        }
    } | Format-Table -AutoSize
    Write-Host 'Parse only: OpenProd was not contacted.' -ForegroundColor Green
    return
}

if (-not $Credential) {
    # Prompted in the console rather than through Get-Credential: its dialog is a
    # separate window that regularly opens off-screen on a multi-monitor setup, and
    # then cannot be brought back. A console prompt is always where the user is
    # looking, and it also works over a remote session.
    $login = Read-Host "OpenProd login on $Database"
    if (-not $login) { throw 'No login given.' }
    $secret = Read-Host 'Password' -AsSecureString
    $Credential = New-Object System.Management.Automation.PSCredential($login, $secret)
}
$me = Connect-OpenProd $Credential
Write-Host ("Connected as {0} (uid {1})." -f $me.username, $me.uid) -ForegroundColor Green

if ($ShowRequired) {
    foreach ($m in 'purchase.order', 'purchase.order.line') {
        $meta = Invoke-OpenProd '/web/dataset/call_kw' @{ model = $m; method = 'fields_get'; args = @(); kwargs = @{} }
        Write-Host ''
        Write-Host "Mandatory fields on $m :" -ForegroundColor Cyan
        foreach ($p in $meta.PSObject.Properties) {
            if ($p.Value.required) {
                Write-Host ("  {0,-30} {1,-12} {2}" -f $p.Name, $p.Value.type, $p.Value.string)
            }
        }
    }
    Write-Host ''
    return
}

# Every lookup happens once, before anything is written: a wrong code must fail on
# the first order and not halfway through the batch.
$partnerModel = Get-Relation 'purchase.order' 'partner_id'
$userModel    = Get-Relation 'purchase.order' 'purchaser_id'
$affairModel  = Get-Relation 'purchase.order' 'affair_id'
$productModel = Get-Relation 'purchase.order.line' 'product_id'
$uomModel     = Get-Relation 'purchase.order.line' 'uom_id'
$taxModel     = Get-Relation 'purchase.order.line' 'taxes_ids'

$partnerId = Resolve-Id $partnerModel 'reference' $SupplierCode 'Supplier'
$productId = Resolve-Id $productModel 'code'      $ProductCode  'Product'
$uomId     = Resolve-Id $uomModel     'name'      $UnitName     'Unit'
$affairId  = $null
if ($AffairCode) { $affairId = Resolve-Id $affairModel 'code' $AffairCode 'Affair' }
$taxIds = @()
if ($TaxDescription) { $taxIds = @((Resolve-Id $taxModel 'description' $TaxDescription 'Tax')) }

Write-Host ("Resolved: supplier {0}={1}, product {2}={3}, unit {4}={5}, affair {6}={7}" -f `
    $SupplierCode, $partnerId, $ProductCode, $productId, $UnitName, $uomId, `
    $(if ($AffairCode) { $AffairCode } else { '(none)' }), $(if ($affairId) { $affairId } else { '-' }))

$purchaserIds = @{}
foreach ($name in @($orders | ForEach-Object { $_.Purchaser } | Where-Object { $_ } | Sort-Object -Unique)) {
    $purchaserIds[$name] = Resolve-Id $userModel 'name' $name 'Purchaser'
    Write-Host ("Resolved: purchaser {0}={1}" -f $name, $purchaserIds[$name])
}
Write-Host ''

$orderDefaults = Get-ModelDefaults 'purchase.order'
$lineDefaults  = Get-ModelDefaults 'purchase.order.line'
Write-Host ("Defaults from the model: {0} field(s) on the order, {1} on the line." -f $orderDefaults.Count, $lineDefaults.Count)

# What default_get does not cover is everything the supplier onchange computes:
# currency, addresses, payment term, invoicing method, account system. Rather than
# guessing them, they are copied from an existing purchase order of the same
# supplier, whose values the ERP itself produced and accepted. 'name' is left out
# on purpose: the model assigns it from its own sequence.
$ourKeys = @('partner_id', 'ref_order', 'default_expected_date', 'source_document',
             'purchase_order_line_ids', 'purchaser_id', 'affair_id', 'name')
$needed = @()
foreach ($p in (Get-Fields 'purchase.order').PSObject.Properties) {
    if ($p.Value.required -and -not $orderDefaults.Contains($p.Name) -and $ourKeys -notcontains $p.Name) {
        $needed += $p.Name
    }
}

# One unfiltered read, then all the filtering in PowerShell.
#
# A domain on purchase.order comes back empty on this server, even for a record
# the very same query returns when the domain is removed: searching partner_id = 7
# found nothing while the unfiltered search listed that exact order as its first
# row. The filter is honoured on res.partner, product.product and the others, so
# it is specific to this model. That matters well beyond the template: the
# duplicate check runs on ref_order through the same mechanism, and a false
# "not found" would recreate every order a second time. Neither is trusted here.
$scanFields = @(@('id', 'name', 'partner_id', 'ref_order') + $needed | Select-Object -Unique)
$recent = Search-OpenProd 'purchase.order' @() $scanFields $ScanDepth 'id desc'
Write-Host ("Scanned the last {0} purchase order(s) for the template and the duplicate check." -f $recent.Count)
if ($recent.Count -ge $ScanDepth) {
    Write-Host ("  the scan is full at {0} rows: raise -ScanDepth if an older order must be seen." -f $ScanDepth) -ForegroundColor DarkYellow
}

if ($InspectOrder) {
    # Dumps a purchase order and its lines field by field. Meant to compare a
    # correct order with one this script produced, rather than guessing which
    # field an amount is computed from.
    $row = $recent | Where-Object { $_.name -eq $InspectOrder } | Select-Object -First 1
    if (-not $row) { throw "Purchase order '$InspectOrder' not found in the last $($recent.Count) orders." }
    $full = Invoke-OpenProd '/web/dataset/call_kw' @{
        model = 'purchase.order'; method = 'read'; args = @(@([int] $row.id), @()); kwargs = @{}
    }
    Write-Host ''
    Write-Host "purchase.order $InspectOrder (id $($row.id)) - non-empty fields:" -ForegroundColor Cyan
    foreach ($p in ($full[0].PSObject.Properties | Sort-Object Name)) {
        $v = $p.Value
        if ($null -eq $v) { continue }
        if ($v -is [bool] -and -not $v) { continue }
        if ($v -is [array] -and $v.Count -eq 0) { continue }
        Write-Host ("  {0,-34} {1}" -f $p.Name, (($v | Out-String).Trim() -replace '\s+', ' '))
    }
    $lineIds = @($full[0].purchase_order_line_ids)
    if ($lineIds.Count) {
        $lines = Invoke-OpenProd '/web/dataset/call_kw' @{
            model = 'purchase.order.line'; method = 'read'; args = @($lineIds, @()); kwargs = @{}
        }
        foreach ($ln in $lines) {
            Write-Host ''
            Write-Host "purchase.order.line id $($ln.id) - non-empty fields:" -ForegroundColor Cyan
            foreach ($p in ($ln.PSObject.Properties | Sort-Object Name)) {
                $v = $p.Value
                if ($null -eq $v) { continue }
                if ($v -is [bool] -and -not $v) { continue }
                if ($v -is [array] -and $v.Count -eq 0) { continue }
                Write-Host ("  {0,-34} {1}" -f $p.Name, (($v | Out-String).Trim() -replace '\s+', ' '))
            }
        }
    }
    Write-Host ''
    return
}

# ref_order -> name, for the duplicate check further down.
$seenRefs = @{}
foreach ($r in $recent) {
    $ref = [string] $r.ref_order
    if ($ref -and -not $seenRefs.ContainsKey($ref)) { $seenRefs[$ref] = $r.name }
}

if ($needed.Count) {
    if ($TemplateOrder) {
        $tplRow = $recent | Where-Object { $_.name -eq $TemplateOrder } | Select-Object -First 1
    }
    else {
        $tplRow = $recent | Where-Object { (ConvertTo-Writable $_.partner_id) -eq $partnerId } | Select-Object -First 1
    }

    if (-not $tplRow) {
        Write-Host ''
        Write-Host 'No template found. The most recent purchase orders are:' -ForegroundColor Yellow
        foreach ($a in ($recent | Select-Object -First 5)) {
            $pn = $a.partner_id
            $lbl = $(if ($pn -is [array] -and $pn.Count -gt 1) { "$($pn[1]) (id $($pn[0]))" } else { "$pn" })
            Write-Host ("  {0,-16} {1}" -f $a.name, $lbl)
        }
        Write-Host ''
        throw ("No purchase order to copy the supplier settings from " +
               "($(if ($TemplateOrder) { "name = $TemplateOrder" } else { "partner_id = $partnerId, $SupplierCode" })). " +
               "Create one draft purchase order by hand in OpenProd for this supplier, then run this again: " +
               "its currency, addresses, payment term and invoicing method will be reused. " +
               "Fields still needed: $($needed -join ', ')")
    }

    $copied = @()
    foreach ($f in $needed) {
        $v = ConvertTo-Writable $tplRow.$f
        if ($null -ne $v -and $v -ne '') { $orderDefaults[$f] = $v; $copied += $f }
    }
    Write-Host ("Copied {0} supplier field(s) from {1}: {2}" -f $copied.Count, $tplRow.name, ($copied -join ', '))
    $stillMissing = @($needed | Where-Object { -not $orderDefaults.Contains($_) })
    if ($stillMissing.Count) {
        Write-Host ("  still empty on the template: {0}" -f ($stillMissing -join ', ')) -ForegroundColor DarkYellow
    }
}
Write-Verbose ("order defaults : " + (($orderDefaults.Keys | Sort-Object) -join ', '))
Write-Verbose ("line defaults  : " + (($lineDefaults.Keys | Sort-Object) -join ', '))
Write-Host ''

$report = New-Object Collections.Generic.List[object]
foreach ($o in $orders) {
    if ($seenRefs.ContainsKey($o.OrderId)) {
        $report.Add([pscustomobject]@{
            Order = $o.OrderId; Lines = $o.Lines.Count
            Amount = [math]::Round((($o.Lines | ForEach-Object { $_.Qty * $_.UnitPrice } | Measure-Object -Sum).Sum), 2)
            Action = 'skipped'; Result = "already exists as $($seenRefs[$o.OrderId])"
        })
        continue
    }

    $lineCmds = @()
    foreach ($l in $o.Lines) {
        $vals = [ordered]@{}
        foreach ($k in $lineDefaults.Keys) { $vals[$k] = $lineDefaults[$k] }
        $vals['sequence']      = $l.Sequence
        $vals['product_id']    = $productId
        # A line carries several quantity fields: the one shown in the grid
        # (uom_qty), the one the amount is computed from (quantity), and the ones
        # in the purchase and price units. The onchange normally propagates one to
        # the others; nothing does here, so a quantity set only on uom_qty leaves
        # the subtotal at zero. All three units are the same here, so the
        # conversion factor is 1 and the value is simply repeated.
        $vals['uom_qty']       = $l.Qty
        $vals['quantity']      = $l.Qty
        $vals['sec_uom_qty']   = $l.Qty
        $vals['uoi_qty']       = $l.Qty
        $vals['price_unit']    = $l.UnitPrice
        $vals['uom_id']        = $uomId
        $vals['sec_uom_id']    = $uomId
        $vals['uoi_id']        = $uomId
        $vals['expected_date'] = $o.Date.ToString('yyyy-MM-dd')
        $vals['name']          = $l.Title
        if ($taxIds.Count) { $vals['taxes_ids'] = @(, @(6, 0, $taxIds)) }
        $lineCmds += , @(0, 0, $vals)
    }

    $values = [ordered]@{}
    foreach ($k in $orderDefaults.Keys) { $values[$k] = $orderDefaults[$k] }
    $values['partner_id']            = $partnerId
    $values['ref_order']             = $o.OrderId
    $values['default_expected_date'] = $o.Date.ToString('yyyy-MM-dd')
    $values['source_document']       = Get-HeaderDescription $o
    $values['purchase_order_line_ids'] = $lineCmds
    if ($purchaserIds.ContainsKey($o.Purchaser)) { $values['purchaser_id'] = $purchaserIds[$o.Purchaser] }
    if ($affairId) { $values['affair_id'] = $affairId }

    $amount = [math]::Round((($o.Lines | ForEach-Object { $_.Qty * $_.UnitPrice } | Measure-Object -Sum).Sum), 2)

    if (-not $Apply) {
        $report.Add([pscustomobject]@{
            Order = $o.OrderId; Lines = $o.Lines.Count; Amount = $amount
            Action = 'would create'; Result = $values.source_document
        })
        Write-Verbose ($values | ConvertTo-Json -Depth 20)
        continue
    }

    try {
        $newId = Invoke-OpenProd '/web/dataset/call_kw' @{
            model = 'purchase.order'; method = 'create'; args = @($values); kwargs = @{}
        }
        # read() by id rather than a search domain, which this server ignores on
        # this model.
        $back = Invoke-OpenProd '/web/dataset/call_kw' @{
            model = 'purchase.order'; method = 'read'; args = @(@([int] $newId), @('name')); kwargs = @{}
        }
        $label = $(if ($back -and $back[0].name) { $back[0].name } else { "id $newId" })
        $seenRefs[$o.OrderId] = $label          # guards against a repeat inside this same run
        $report.Add([pscustomobject]@{
            Order = $o.OrderId; Lines = $o.Lines.Count; Amount = $amount; Action = 'created'; Result = $label
        })
        Write-Host ("  created  {0}  ->  {1}" -f $o.OrderId, $label) -ForegroundColor Green
    }
    catch {
        $msg = $_.Exception.Message
        $report.Add([pscustomobject]@{
            Order = $o.OrderId; Lines = $o.Lines.Count; Amount = $amount; Action = 'FAILED'; Result = $msg
        })
        Write-Host ("  FAILED   {0}  ->  {1}" -f $o.OrderId, $msg) -ForegroundColor Red

        # Odoo only ever says "a mandatory field is not correctly set" without
        # naming it. Name it here, once, on the first failure.
        if ($msg -match 'mandatory field' -and -not $script:MandatoryReported) {
            $script:MandatoryReported = $true
            foreach ($pair in @(@('purchase.order', $values), @('purchase.order.line', $lineCmds[0][2]))) {
                $gaps = Get-MissingMandatory $pair[0] $pair[1]
                Write-Host ''
                if ($gaps.Count) {
                    Write-Host ("Mandatory fields left unset on {0} :" -f $pair[0]) -ForegroundColor Yellow
                    $gaps | ForEach-Object { Write-Host "    $_" -ForegroundColor Yellow }
                }
                else {
                    Write-Host ("No mandatory field is missing on {0}." -f $pair[0]) -ForegroundColor DarkGray
                }
            }
            Write-Host ''
        }
    }
}

Write-Host ''
$report | Format-Table -AutoSize
$created = @($report | Where-Object { $_.Action -eq 'created' })
$failed  = @($report | Where-Object { $_.Action -eq 'FAILED' })
$skippedExisting = @($report | Where-Object { $_.Action -eq 'skipped' })

if ($Apply) {
    Write-Host ("{0} created, {1} already present, {2} failed." -f $created.Count, $skippedExisting.Count, $failed.Count) `
        -ForegroundColor $(if ($failed.Count) { 'Red' } else { 'Green' })
    Write-Host 'The purchase orders are in draft. Check them in OpenProd before requesting validation.'
}
else {
    Write-Host ("Dry run finished. {0} order(s) would be created, {1} already present." -f `
        @($report | Where-Object { $_.Action -eq 'would create' }).Count, $skippedExisting.Count) -ForegroundColor Green
    Write-Host 'Re-run with -Apply to actually create them.'
}
Write-Host ''

#endregion
