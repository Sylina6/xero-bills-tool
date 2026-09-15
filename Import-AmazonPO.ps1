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

    [System.Management.Automation.PSCredential] $Credential,

    # Process only the first N orders. Meant for a cautious first run: -Limit 1
    # creates a single purchase order that you can check before doing the batch.
    [int] $Limit = 0,

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

$script:Session = $null

function Invoke-OpenProd([string] $Route, [hashtable] $Params) {
    $body = @{ jsonrpc = '2.0'; method = 'call'; params = $Params } | ConvertTo-Json -Depth 20 -Compress
    $args = @{
        Uri         = "$Server$Route"
        Method      = 'Post'
        ContentType = 'application/json'
        Body        = [Text.Encoding]::UTF8.GetBytes($body)
        TimeoutSec  = 120
    }
    if ($script:Session) { $args.WebSession = $script:Session } else { $args.SessionVariable = 'newSession' }

    $resp = Invoke-RestMethod @args
    if (-not $script:Session) { $script:Session = $newSession }

    # JSON-RPC reports business errors inside a 200 response, so the HTTP status
    # alone never tells you whether the call worked.
    if ($resp.PSObject.Properties.Name -contains 'error' -and $resp.error) {
        $msg = $resp.error.message
        if ($resp.error.data -and $resp.error.data.message) { $msg = "$msg : $($resp.error.data.message)" }
        throw "OpenProd API error on $Route -> $msg"
    }
    return $resp.result
}

function Connect-OpenProd([System.Management.Automation.PSCredential] $Cred) {
    $script:Session = $null
    $plain = $Cred.GetNetworkCredential().Password
    $res = Invoke-OpenProd '/web/session/authenticate' @{
        db = $Database; login = $Cred.UserName; password = $plain
    }
    if (-not $res.uid) { throw "Authentication refused for '$($Cred.UserName)' on database '$Database'." }
    return $res
}

function Search-OpenProd([string] $Model, [array] $Domain, [array] $Fields, $Limit = 1) {
    $res = Invoke-OpenProd '/web/dataset/search_read' @{
        model = $Model; domain = $Domain; fields = $Fields; limit = $Limit; sort = $null
    }
    return @($res.records)
}

# Relations are read from the model itself instead of being hard-coded: it is the
# only way to know which model 'affair_id' or 'purchaser_id' actually point to.
$script:Relations = @{}
function Get-Relation([string] $Model, [string] $Field) {
    $key = "$Model.$Field"
    if ($script:Relations.ContainsKey($key)) { return $script:Relations[$key] }
    if (-not $script:Relations.ContainsKey($Model)) {
        $script:Relations[$Model] = Invoke-OpenProd '/web/dataset/call_kw' @{
            model = $Model; method = 'fields_get'; args = @(); kwargs = @{}
        }
    }
    $meta = $script:Relations[$Model].$Field
    if (-not $meta) { throw "Field '$Field' does not exist on model '$Model'." }
    $script:Relations[$key] = $meta.relation
    return $meta.relation
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

if (-not $Credential) { $Credential = Get-Credential -Message "OpenProd account on $Database" }
$me = Connect-OpenProd $Credential
Write-Host ("Connected as {0} (uid {1})." -f $me.username, $me.uid) -ForegroundColor Green

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

$report = New-Object Collections.Generic.List[object]
foreach ($o in $orders) {
    $existing = Search-OpenProd 'purchase.order' @(, @('ref_order', '=', $o.OrderId)) @('id', 'name') 1
    if ($existing.Count) {
        $report.Add([pscustomobject]@{
            Order = $o.OrderId; Lines = $o.Lines.Count
            Amount = [math]::Round((($o.Lines | Measure-Object -Property { $_.Qty * $_.UnitPrice } -Sum).Sum), 2)
            Action = 'skipped'; Result = "already exists as $($existing[0].name)"
        })
        continue
    }

    $lineCmds = @()
    foreach ($l in $o.Lines) {
        $vals = [ordered]@{
            sequence      = $l.Sequence
            product_id    = $productId
            uom_qty       = $l.Qty
            price_unit    = $l.UnitPrice
            uom_id        = $uomId
            sec_uom_id    = $uomId
            uoi_id        = $uomId
            expected_date = $o.Date.ToString('yyyy-MM-dd')
            name          = $l.Title
        }
        if ($taxIds.Count) { $vals.taxes_ids = @(, @(6, 0, $taxIds)) }
        $lineCmds += , @(0, 0, $vals)
    }

    $values = [ordered]@{
        partner_id             = $partnerId
        ref_order              = $o.OrderId
        default_expected_date  = $o.Date.ToString('yyyy-MM-dd')
        source_document        = Get-HeaderDescription $o
        purchase_order_line_ids = $lineCmds
    }
    if ($purchaserIds.ContainsKey($o.Purchaser)) { $values.purchaser_id = $purchaserIds[$o.Purchaser] }
    if ($affairId) { $values.affair_id = $affairId }

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
        $created = Search-OpenProd 'purchase.order' @(, @('id', '=', [int] $newId)) @('id', 'name') 1
        $label = $(if ($created.Count) { $created[0].name } else { "id $newId" })
        $report.Add([pscustomobject]@{
            Order = $o.OrderId; Lines = $o.Lines.Count; Amount = $amount; Action = 'created'; Result = $label
        })
        Write-Host ("  created  {0}  ->  {1}" -f $o.OrderId, $label) -ForegroundColor Green
    }
    catch {
        $report.Add([pscustomobject]@{
            Order = $o.OrderId; Lines = $o.Lines.Count; Amount = $amount; Action = 'FAILED'; Result = $_.Exception.Message
        })
        Write-Host ("  FAILED   {0}  ->  {1}" -f $o.OrderId, $_.Exception.Message) -ForegroundColor Red
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
