param([Parameter(Mandatory=$true)][string]$Manifest)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=New-Object System.Text.UTF8Encoding($false)
$excel=$null;$book=$null;$results=@()
try {
  $excel=New-Object -ComObject Excel.Application
  $excel.AutomationSecurity=3;$excel.EnableEvents=$false;$excel.DisplayAlerts=$false
  $cases=Get-Content -LiteralPath $Manifest -Raw -Encoding UTF8|ConvertFrom-Json
  foreach($case in $cases){
    $book=$excel.Workbooks.Open([string]$case.path,0,$true)
    $sheet=$book.Worksheets.Item(1);$charts=$sheet.ChartObjects()
    if($charts.Count -ne $case.count){throw ('Chart count mismatch: '+$case.name+' = '+$charts.Count)}
    $items=@()
    foreach($object in $charts){
      $chart=$object.Chart;$series=$chart.SeriesCollection()
      $seriesTypes=@();$axisGroups=@();foreach($item in $series){$seriesTypes+=([int]$item.ChartType);$axisGroups+=([int]$item.AxisGroup)}
      $items+=@{name=[string]$object.Name;title=$(if($chart.HasTitle){[string]$chart.ChartTitle.Text}else{''});type=[int]$chart.ChartType;series=$series.Count;seriesTypes=$seriesTypes;axisGroups=$axisGroups;firstValues=@($series.Item(1).Values);width=[double]$object.Width;height=[double]$object.Height}
    }
    $saved=[IO.Path]::ChangeExtension([string]$case.path,'.resaved.xlsx')
    $book.SaveCopyAs($saved);$book.Close($false);$book=$null
    $results+=@{name=$case.name;saved=$saved;items=$items;ok=$true}
  }
  ConvertTo-Json -InputObject $results -Depth 8 -Compress
} finally {
  if($book){$book.Close($false)}
  if($excel){$excel.Quit()}
  foreach($value in @($series,$chart,$object,$charts,$sheet,$book,$excel)){if($value -and [Runtime.InteropServices.Marshal]::IsComObject($value)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($value)}}
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
