param([Parameter(Mandatory=$true)][string]$Manifest)
$ErrorActionPreference='Stop'
$excel=$null;$book=$null;$results=@()
try {
  $excel=New-Object -ComObject Excel.Application
  $excel.AutomationSecurity=3;$excel.EnableEvents=$false;$excel.DisplayAlerts=$false
  $cases=Get-Content -LiteralPath $Manifest -Raw -Encoding UTF8|ConvertFrom-Json
  foreach($case in $cases){
    $book=$excel.Workbooks.Open([string]$case.path,0,$true)
    $sheet=$book.Worksheets.Item('Summary');$pivots=$sheet.PivotTables()
    if($case.empty){if($pivots.Count -ne 0){throw 'Empty report must not use an Excel filter that selects an arbitrary item'}}
    else {
      if($pivots.Count -ne 1){throw ('Expected one native pivot: '+$case.name)}
      $pivot=$pivots.Item(1)
      if(!$pivot.RefreshTable()){throw ('Pivot refresh failed: '+$case.name)}
      $range=$pivot.TableRange2
      if($null -ne $case.total){$last=$range.Cells.Item($range.Rows.Count,$range.Columns.Count).Value2;if([double]$last -ne [double]$case.total){throw ('Wrong total: '+$case.name+' = '+$last)}}
    }
    $saved=[IO.Path]::ChangeExtension([string]$case.path,'.resaved.xlsx')
    $book.SaveCopyAs($saved);$book.Close($false);$book=$null
    $results+=@{name=$case.name;saved=$saved;ok=$true}
  }
  ConvertTo-Json -InputObject $results -Compress
} finally {
  if($book){$book.Close($false)}
  if($excel){$excel.Quit()}
  foreach($value in @($range,$pivot,$pivots,$sheet,$book,$excel)){if($value -and [Runtime.InteropServices.Marshal]::IsComObject($value)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($value)}}
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
