param([string]$Workbook = (Join-Path $PSScriptRoot '../templates/SRAM Extra BLK Failure  Analysis_v13.4.xlsm'))
$ErrorActionPreference = 'Stop'
$app = $null
$book = $null

function Assert($Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}

function New-SourceFixture([string]$SheetName, [string]$Header) {
    $sheet = $book.Worksheets.Add()
    $sheet.Name = $SheetName
    $sheet.Range('A1').Value2 = 'Other value'
    $sheet.Range('B1').Value2 = $Header
    $sheet.Range('A2').Value2 = [double]3
    $sheet.Range('B2').Value2 = [double]3
    $sheet.Range('B3').Value2 = [double]999999
    $sheet.Range('B4').Formula = '=NA()'
    $sheet.Range('B5').Value2 = [double]1
    $null = $sheet.ListObjects.Add(1, $sheet.Range('A1:B5'), $null, 1)
    return $sheet
}

try {
    $app = New-Object -ComObject Excel.Application
    $app.Visible = $false
    $app.DisplayAlerts = $false
    $app.EnableEvents = $false
    $book = $app.Workbooks.Open((Resolve-Path $Workbook).Path, 0, $true)
    $source = $book.Worksheets.Item('Fail Address')
    $source.Visible = -1
    $source.Range('A3').Value2 = 'Source row test'
    $detail = New-SourceFixture 'Chip Fail Details' 'Source Row'
    $analysis = New-SourceFixture 'Fail Analysis' 'Source Address Row'
    $otherSheet = New-SourceFixture 'Source Row Fixture' 'Source Row'
    $app.EnableEvents = $true

    foreach ($sheet in @($detail, $analysis)) {
        $sheet.Activate()
        foreach ($address in @('B1', 'A2', 'B3', 'B4', 'B5', 'B2:B3')) {
            $sheet.Range($address).Select()
            Assert ($book.ActiveSheet.Name -eq $sheet.Name) "Unexpected navigation from $($sheet.Name)!$address."
        }
        $sheet.Range('B2').Select()
        Assert ($book.ActiveSheet.Name -eq $source.Name) "Source Row selection did not navigate from $($sheet.Name)."
        Assert ($app.ActiveCell.Row -eq 3 -and $app.ActiveCell.Column -eq 1) 'Selection landed on the wrong source cell.'
        Assert $app.EnableEvents 'Source navigation did not restore Excel events.'
    }

    $otherSheet.Activate()
    $otherSheet.Range('B2').Select()
    Assert ($book.ActiveSheet.Name -eq $otherSheet.Name) 'An unrelated worksheet triggered source navigation.'
    Write-Output 'PASS native Excel Source Row selection for both detail sheets, valid and invalid cells, and event restoration.'
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $app) { $app.Quit() }
}
