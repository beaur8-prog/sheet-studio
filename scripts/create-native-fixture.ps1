param([Parameter(Mandatory=$true)][string]$Destination)
$ErrorActionPreference='Stop'
$excel=$null;$book=$null
try {
  $excel=New-Object -ComObject Excel.Application
  $excel.AutomationSecurity=3;$excel.EnableEvents=$false;$excel.DisplayAlerts=$false
  $book=$excel.Workbooks.Add()
  $sheet=$book.Worksheets.Item(1);$sheet.Name='Data'
  $sheet.Range('A1').Value2=7.0;$sheet.Range('A2').Formula='=A1*2'
  $sheet.Range('C1').Value2='preserve';$sheet.Range('C1').Interior.Color=255
  [void]$book.Names.Add('OriginalInput','=Data!$A$1')
  $chart=$sheet.ChartObjects().Add(150,30,300,180)
  $chart.Chart.SetSourceData($sheet.Range('A1:A2'))
  $hasMacro=$false;$macroReason=''
  # Exercise the current policy. Never enable AccessVBOM to create a fixture.
  try {
    $module=$book.VBProject.VBComponents.Add(1);$module.Name='SheetStudioTest'
    $module.CodeModule.AddFromString("Public Sub Smoke()`r`nThisWorkbook.Worksheets(""Data"").Range(""B2"").Value = ThisWorkbook.Worksheets(""Data"").Range(""A1"").Value * 3`r`nEnd Sub")
    $hasMacro=$true
  } catch {$macroReason=$_.Exception.Message}
  $book.SaveAs($Destination,52)
  @{path=$Destination;hasMacro=$hasMacro;macroReason=$macroReason}|ConvertTo-Json -Compress
} finally {
  if($book){$book.Close($false)}
  if($excel){$excel.Quit()}
  foreach($value in @($module,$chart,$sheet,$book,$excel)){if($value -and [Runtime.InteropServices.Marshal]::IsComObject($value)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($value)}}
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
