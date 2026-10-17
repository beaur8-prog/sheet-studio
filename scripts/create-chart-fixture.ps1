param([Parameter(Mandatory=$true)][string]$Destination)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=New-Object System.Text.UTF8Encoding($false)
$excel=$null;$book=$null
try {
  $excel=New-Object -ComObject Excel.Application
  $excel.AutomationSecurity=3;$excel.EnableEvents=$false;$excel.DisplayAlerts=$false
  $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1);$sheet.Name='Source'
  $rows=@(@('Category','Revenue','Expense'),@('A',40,22),@('B',25,18),@('C',10,5),@('D',35,10))
  for($r=0;$r -lt $rows.Count;$r++){for($c=0;$c -lt $rows[$r].Count;$c++){if($r -eq 0 -or $c -eq 0){$sheet.Cells.Item($r+1,$c+1).Value2=[string]$rows[$r][$c]}else{$sheet.Cells.Item($r+1,$c+1).Value2=[double]$rows[$r][$c]}}}
  $types=@(51,52,53,57,4,1,5,-4120,-4169,15,-4151,54,88,83)
  $results=@();$i=0
  foreach($type in $types){
    $object=$sheet.ChartObjects().Add(180+($i%3)*350,20+[Math]::Floor($i/3)*250,330,220)
    $chart=$object.Chart;$chart.SetSourceData($sheet.Range('A1:C5'),2)
    try{$chart.ChartType=$type;$chart.HasTitle=$true;$chart.ChartTitle.Text='Original '+$type;$results+=@{type=$type;created=$true}}catch{$object.Delete();$results+=@{type=$type;created=$false;reason=$_.Exception.Message}}
    $i++
  }
  # Modern chart types must be created through AddChart2 on recent Excel versions.
  foreach($type in @(117,118,119,120,121,122,123)){
    try{$shape=$sheet.Shapes.AddChart2(-1,$type,180,1400+($i-14)*250,330,220);$chart=$shape.Chart;$chart.SetSourceData($sheet.Range('A1:C5'));$chart.HasTitle=$true;$chart.ChartTitle.Text='Modern '+$type;$results+=@{type=$type;created=$true}}catch{if($shape){try{$shape.Delete()}catch{}};$results+=@{type=$type;created=$false;reason=$_.Exception.Message}};$shape=$null;$i++
  }
  $book.SaveAs($Destination,51)
  ConvertTo-Json -InputObject $results -Depth 5 -Compress
}finally{
  if($book){$book.Close($false)}
  if($excel){$excel.Quit()}
  foreach($value in @($chart,$object,$shape,$sheet,$book,$excel)){if($value -and [Runtime.InteropServices.Marshal]::IsComObject($value)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($value)}}
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
