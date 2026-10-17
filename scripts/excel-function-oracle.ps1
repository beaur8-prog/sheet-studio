param([Parameter(Mandatory=$true)][string]$InputPath,[Parameter(Mandatory=$true)][string]$OutputPath,[string]$ReadWorkbook)
$ErrorActionPreference='Stop'
$excel=$null
$workbook=$null
try {
  $cases=Get-Content -LiteralPath $InputPath -Raw -Encoding UTF8 | ConvertFrom-Json
  $excel=New-Object -ComObject Excel.Application
  $excel.Visible=$false
  $excel.DisplayAlerts=$false
  $excel.AutomationSecurity=3
  if($ReadWorkbook){$workbook=$excel.Workbooks.Open([IO.Path]::GetFullPath($ReadWorkbook),0,$true)}else{$workbook=$excel.Workbooks.Add()}
  $sheet=$workbook.Worksheets.Item(1)
  if(-not $ReadWorkbook){
    $sheet.Name='Sheet1'
    for($i=1;$i -le 5;$i++){$sheet.Cells.Item($i,1).Value2=[double]$i}
    $sheet.Cells.Item(1,2).Formula='=SUM(A1:A5)'
  }
  $results=@()
  $i=1
  foreach($case in $cases){
    $cell=$sheet.Cells.Item($i,4)
    try {
      if(-not $ReadWorkbook){$cell.Formula=$case.formula}
      [void]$cell.Calculate()
      $value=$cell.Value2
      if($excel.WorksheetFunction.IsError($cell)){$value=$cell.Text}
      $results+=@{name=$case.name;formula=$case.formula;expected=$value}
    } catch {$results+=@{name=$case.name;formula=$case.formula;oracleError=$_.Exception.Message}}
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell)
    $i++
  }
  @{excelVersion=$excel.Version;excelBuild=$excel.Build;capturedAt=[DateTime]::UtcNow.ToString('o');cases=$results} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
  Write-Output ('Excel oracle recorded '+$results.Count+' formulas.')
} finally {
  if($workbook){$workbook.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)}
  if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
