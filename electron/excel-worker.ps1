param([Parameter(Mandatory=$true)][string]$JobDir,[switch]$CancelOwned)
$ErrorActionPreference = 'Stop'
$utf8 = New-Object System.Text.UTF8Encoding($false)
function Write-State([string]$state,[string]$message) {
  $value = @{state=$state;message=$message} | ConvertTo-Json -Compress
  $temporary = Join-Path $JobDir 'status-next.json'
  [IO.File]::WriteAllText($temporary,$value,$utf8)
  Move-Item -LiteralPath $temporary -Destination (Join-Path $JobDir 'status.json') -Force
}
function Release-Com($object) { if ($null -ne $object -and [Runtime.InteropServices.Marshal]::IsComObject($object)) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($object) } }
function Check-Cancel { if(Test-Path -LiteralPath (Join-Path $JobDir 'cancel.flag')){throw [OperationCanceledException]::new('Excel 작업을 취소했습니다.')} }
function Preserve-Origin([string]$file,[string]$zone) {
  if($zone){Set-Content -LiteralPath $file -Stream Zone.Identifier -Encoding Byte -Value $utf8.GetBytes($zone)}
}
function Save-Result($workbook,$zone) {
  $output=Join-Path $JobDir 'result.xlsm'
  [void]$workbook.SaveCopyAs($output)
  Preserve-Origin $output $zone
}
if ($CancelOwned) {
  $ownerPath = Join-Path $JobDir 'owner.json'
  if (Test-Path -LiteralPath $ownerPath) {
    $owner = Get-Content -LiteralPath $ownerPath -Raw | ConvertFrom-Json
    $owned = Get-Process -Id $owner.processId -ErrorAction SilentlyContinue
    if ($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks.ToString() -eq $owner.started) { Stop-Process -Id $owned.Id -Force }
  }
  exit
}
function Has-Field($object,[string]$name) { return $null -ne $object.PSObject.Properties[$name] }
function Color-Number([string]$value) {
  if ($value -match '^#([a-fA-F0-9]{3})$') { $hex=($value.Substring(1).ToCharArray() | ForEach-Object { "$_$_" }) -join ''; $value='#'+$hex }
  if ($value -notmatch '^#[a-fA-F0-9]{6}$') { throw "Excel color must be #RRGGBB: $value" }
  return [Convert]::ToInt32($value.Substring(1,2),16) + 256*[Convert]::ToInt32($value.Substring(3,2),16) + 65536*[Convert]::ToInt32($value.Substring(5,2),16)
}
function Set-Cell($sheet,$patch) {
  $cell=$sheet.Range($patch.address)
  try {
    if (Has-Field $patch 'raw') {
      $raw=[string]$patch.raw
      if ($raw -eq '') { [void]$cell.ClearContents() }
      elseif ($raw.StartsWith('=')) { try {$cell.Formula2=$raw} catch {$cell.Formula=$raw} }
      elseif ($raw.StartsWith("'")) {$cell.Value2=[string]$raw}
      elseif ($raw -match '^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$') { $cell.Value2=[double]::Parse($raw,[Globalization.CultureInfo]::InvariantCulture) }
      elseif ($raw -match '^[+-]?(?:\d+(?:\.\d*)?|\.\d+)%$') {$cell.Value2=[double]::Parse($raw.TrimEnd('%'),[Globalization.CultureInfo]::InvariantCulture)/100}
      elseif ($raw -match '^(?i:TRUE|FALSE)$') {$cell.Value2=($raw.ToUpperInvariant() -eq 'TRUE')}
      else {$cell.Value2=[string]("'"+$raw)}
    }
    if (Has-Field $patch 'style') {
      $style=$patch.style
      foreach($property in $style.PSObject.Properties) {
        $value=$property.Value
        switch($property.Name) {
          'bold' {$cell.Font.Bold=[bool]$value}
          'italic' {$cell.Font.Italic=[bool]$value}
          'underline' {if($value){$cell.Font.Underline=2}else{$cell.Font.Underline=-4142}}
          'fontFamily' {if($value){$cell.Font.Name=[string]$value}else{$cell.Font.Name='Calibri'}}
          'fontSize' {if($value){$cell.Font.Size=[double]$value}else{$cell.Font.Size=11}}
          'color' {if($value){$cell.Font.Color=Color-Number $value}else{$cell.Font.ColorIndex=-4105}}
          'background' {if($value){$cell.Interior.Color=Color-Number $value}else{$cell.Interior.ColorIndex=-4142}}
          'wrap' {$cell.WrapText=[bool]$value}
          'align' {$align=@{left=-4131;center=-4108;right=-4152;fill=5;justify=-4130;distributed=-4117;centerContinuous=7};if($align.ContainsKey([string]$value)){$cell.HorizontalAlignment=$align[[string]$value]}else{$cell.HorizontalAlignment=1}}
          'verticalAlign' {$align=@{top=-4160;middle=-4108;bottom=-4107};if($align.ContainsKey([string]$value)){$cell.VerticalAlignment=$align[[string]$value]}else{$cell.VerticalAlignment=-4107}}
          'border' {if($value){$cell.Borders.LineStyle=1;$cell.Borders.Weight=2}else{$cell.Borders.LineStyle=-4142}}
          'format' {$formats=@{general=[string]$sheet.Application.International(26);number='#,##0.00';currency='₩#,##0.00';accounting='₩#,##0.00';percent='0.00%';date='yyyy-mm-dd'};if($formats.ContainsKey([string]$value)){$cell.NumberFormat=[string]$formats[[string]$value]}else{$cell.NumberFormat=[string]$sheet.Application.International(26)}}
        }
      }
      if (Has-Field $style 'borders') {
        # Apply per-edge settings after the legacy all-borders flag, regardless of JSON key order.
        $indices=@{top=8;right=10;bottom=9;left=7}
        foreach($side in $indices.Keys) {
          if($null -ne $style.borders -and !(Has-Field $style.borders $side)){continue}
          $edge=$null;if($null -ne $style.borders){$edge=$style.borders.$side}
          $border=$cell.Borders.Item($indices[$side])
          try {
            if($null -eq $edge){$border.LineStyle=-4142;continue}
            $lines=@{dotted=-4118;dashed=-4115;mediumDashed=-4115;double=-4119;dashDot=4;mediumDashDot=4;dashDotDot=5;mediumDashDotDot=5;slantDashDot=13}
            $line=1;if($lines.ContainsKey([string]$edge.style)){$line=$lines[[string]$edge.style]}
            $weight=2;if($edge.style -eq 'hair'){$weight=1}elseif($edge.style -like 'medium*'){$weight=-4138}elseif($edge.style -eq 'thick'){$weight=4}
            $border.Weight=$weight;$border.LineStyle=$line;$border.Color=Color-Number $edge.color
          } finally {Release-Com $border}
        }
      }
      if ((Has-Field $style 'numberFormat') -and $style.numberFormat) {if($style.numberFormat -eq 'General'){$cell.NumberFormat=[string]$sheet.Application.International(26)}else{$cell.NumberFormat=[string]$style.numberFormat}}
      if ((Has-Field $style 'decimals') -and $null -ne $style.decimals -and $style.format -in @('number','currency','accounting','percent')) {
        $places=[Math]::Min(12,[Math]::Max(0,[int]$style.decimals));$fraction='';if($places){$fraction='.'+('0'*$places)}
        if($style.format -eq 'percent'){$cell.NumberFormat='0'+$fraction+'%'}elseif($style.format -in @('currency','accounting')){$cell.NumberFormat='₩#,##0'+$fraction}else{$cell.NumberFormat='#,##0'+$fraction}
      }
    }
    if (Has-Field $patch 'note') { if($cell.Comment){[void]$cell.Comment.Delete()};if($patch.note){[void]$cell.AddComment([string]$patch.note)} }
  } finally {Release-Com $cell}
}
function Set-Chart($sheet,$patch) {
  $object=$null;$chart=$null
  try {
    if($patch.name){$object=$sheet.ChartObjects().Item([string]$patch.name)}else{$object=$sheet.ChartObjects().Add([double]$patch.x*0.75,[double]$patch.y*0.75,[double]$patch.width*0.75,[double]$patch.height*0.75)}
    $chart=$object.Chart
    if(Has-Field $patch 'range'){$chart.SetSourceData($sheet.Range([string]$patch.range),2)}
    if(Has-Field $patch 'chartType'){$chart.ChartType=[int]$patch.chartType}
    if(Has-Field $patch 'series'){
      $collection=$chart.SeriesCollection()
      try {
        while($collection.Count -gt 0){$oldSeries=$collection.Item(1);try{[void]$oldSeries.Delete()}finally{Release-Com $oldSeries}}
        $index=0
        foreach($spec in $patch.series){
          $item=$collection.NewSeries();$values=$null;$categories=$null;$nameRange=$null;$sizes=$null
          try {
            $values=$sheet.Range([string]$spec.valuesRange);$item.Values=$values
            if(Has-Field $spec 'categoriesRange'){$categories=$sheet.Range([string]$spec.categoriesRange);$item.XValues=$categories}else{$item.XValues=[object[]](1..([int]$values.Count))}
            $prefix="='"+([string]$sheet.Name).Replace("'","''")+"'!"
            if(Has-Field $spec 'nameRange'){$nameRange=$sheet.Range([string]$spec.nameRange);$item.Name=$prefix+'R'+$nameRange.Row+'C'+$nameRange.Column}else{$item.Name='계열 '+($index+1)}
            if(Has-Field $spec 'sizesRange'){$sizes=$sheet.Range([string]$spec.sizesRange);$item.BubbleSizes=$prefix+$sizes.Address($true,$true,-4150)}
          }finally{Release-Com $sizes;Release-Com $nameRange;Release-Com $categories;Release-Com $values;Release-Com $item}
          $index++
        }
      }finally{Release-Com $collection}
    }
    if($patch.combo){$series=$chart.SeriesCollection();if($series.Count -gt 1){$last=$series.Item($series.Count);$last.ChartType=4;if($patch.secondaryAxis){$last.AxisGroup=2};Release-Com $last};Release-Com $series}
    foreach($key in @('x','y','width','height')){if(Has-Field $patch $key){$value=[double]$patch.$key*0.75;switch($key){'x'{$object.Left=$value}'y'{$object.Top=$value}'width'{$object.Width=$value}'height'{$object.Height=$value}}}}
    if(Has-Field $patch 'title'){$chart.HasTitle=[bool]$patch.title;if($patch.title){$chart.ChartTitle.Text=[string]$patch.title}}
    if(Has-Field $patch 'legend'){$chart.HasLegend=[bool]$patch.legend}
    if(Has-Field $patch 'plotVisibleOnly'){$chart.PlotVisibleOnly=[bool]$patch.plotVisibleOnly}
    if(Has-Field $patch 'labels'){if($patch.labels){[void]$chart.ApplyDataLabels();if($patch.percentLabels){foreach($series in $chart.SeriesCollection()){$labels=$null;try{$labels=$series.DataLabels();$labels.ShowValue=$false;$labels.ShowPercentage=$true}finally{Release-Com $labels;Release-Com $series}}}}else{foreach($series in $chart.SeriesCollection()){$series.HasDataLabels=$false;Release-Com $series}}}
    foreach($field in @('xTitle','yTitle')){if((Has-Field $patch $field) -and $chart.ChartType -notin @(5,-4120,-4151,81,82)){$axisType=1;if($field -eq 'yTitle'){$axisType=2};if($chart.HasAxis($axisType,1)){$axis=$chart.Axes($axisType,1);$axis.HasTitle=[bool]$patch.$field;if($patch.$field){$axis.AxisTitle.Text=[string]$patch.$field};Release-Com $axis}}}
  }finally{Release-Com $chart;Release-Com $object}
}
function Set-AutoFilter($sheet,$filter) {
  if($sheet.AutoFilterMode){$sheet.AutoFilterMode=$false}
  if($null -eq $filter){return}
  $range=$sheet.Range([string]$filter.range)
  try {
    [void]$range.AutoFilter()
    foreach($rule in $filter.filters){
      $v=[Type]::Missing;if($null -ne $rule.criteria1){$v=$rule.criteria1};if($v -is [array]){$v=[object[]]$v}
      if($rule.dateValues){
        $dates=@{};foreach($day in $rule.dateValues){$dates[[string]$day]=$true}
        $labels=@();if($rule.criteria1){$labels+=@($rule.criteria1)}
        for($i=2;$i -le $range.Rows.Count;$i++){$cell=$range.Cells.Item($i,[int]$rule.field);try{$value=$cell.Value2;if($value -is [double] -and $dates.ContainsKey([string][Math]::Floor($value))){$labels+=[string]$cell.Text}}finally{Release-Com $cell}}
        if(!$labels.Count){$labels=@($rule.dateValues | ForEach-Object {[string]$_})};$v=[string[]]@($labels | Select-Object -Unique)
      }
      $op=[Type]::Missing;if($null -ne $rule.operator){$op=[int]$rule.operator}
      $v2=[Type]::Missing;if($null -ne $rule.criteria2){$v2=$rule.criteria2}
      [void]$range.AutoFilter([int]$rule.field,$v,$op,$v2)
    }
    # Keep the user's applied snapshot until Reapply, as in Excel after cell edits.
    $range.EntireRow.Hidden=$false
    foreach($row in $filter.hiddenRows){$r=$sheet.Rows.Item([int]$row+1);try{$r.Hidden=$true}finally{Release-Com $r}}
  }finally{Release-Com $range}
}
function Set-ConditionalFormats($sheet,$rules) {
  [void]$sheet.Cells.FormatConditions.Delete()
  $priority=1
  foreach($entry in $rules){
    $range=$sheet.Range([string]$entry.range);$rule=$null
    try {
      $op=[Type]::Missing;if($entry.type -eq 1){$op=[int]$entry.operator}
      $f2=[Type]::Missing;if($entry.formula2){$f2=[string]$entry.formula2}
      $f1=[string]$entry.formula1;if($entry.type -eq 2 -and !$f1.StartsWith('=')){$f1='='+$f1}
      $rule=$range.FormatConditions.Add([int]$entry.type,$op,$f1,$f2)
      $rule.Priority=$priority;$priority++;$rule.StopIfTrue=[bool]$entry.stopIfTrue
      if($entry.style.background){$rule.Interior.Color=Color-Number $entry.style.background}
      if($entry.style.color){$rule.Font.Color=Color-Number $entry.style.color}
      if(Has-Field $entry.style 'bold'){$rule.Font.Bold=[bool]$entry.style.bold}
      if(Has-Field $entry.style 'italic'){$rule.Font.Italic=[bool]$entry.style.italic}
    }finally{Release-Com $rule;Release-Com $range}
  }
}
function Apply-Plan($book,$plan) {
  $map=@{}
  foreach($entry in $plan.sheets) {
    if($entry.sourceName){$sheet=$book.Worksheets.Item([string]$entry.sourceName)}else{$sheet=$book.Worksheets.Add([Type]::Missing,$book.Worksheets.Item($book.Worksheets.Count))}
    $map[[string]$entry.id]=$sheet
  }
  try {
    foreach($name in $plan.deleted){$sheet=$book.Worksheets.Item([string]$name);try{[void]$sheet.Delete()}finally{Release-Com $sheet}}
    foreach($entry in $plan.sheets){$sheet=$map[[string]$entry.id];if($sheet.Name -ne $entry.name){$sheet.Name='ss_'+[Guid]::NewGuid().ToString('N').Substring(0,20)}}
    foreach($entry in $plan.sheets){$map[[string]$entry.id].Name=[string]$entry.name}
    foreach($op in $plan.operations){$sheet=$map[[string]$op.sheetId];if(!$sheet){continue};if($op.axis -eq 'r'){$range=$sheet.Rows.Item([int]$op.index+1)}else{$range=$sheet.Columns.Item([int]$op.index+1)};try{if($op.delta -eq 1){[void]$range.Insert()}else{[void]$range.Delete()}}finally{Release-Com $range}}
    foreach($entry in $plan.sheets) {
      $sheet=$map[[string]$entry.id]
      foreach($range in $entry.unmerge){$r=$sheet.Range([string]$range);try{[void]$r.UnMerge()}finally{Release-Com $r}}
      foreach($patch in $entry.cells){Set-Cell $sheet $patch}
      foreach($p in $entry.columnWidths.PSObject.Properties){$r=$sheet.Columns.Item([int]$p.Name+1);try{$r.Hidden=($null -ne $p.Value -and $p.Value -eq 0);if($null -eq $p.Value){$r.ColumnWidth=$sheet.StandardWidth}elseif($p.Value -gt 0){$r.ColumnWidth=[Math]::Max(1,([double]$p.Value-5)/7)}}finally{Release-Com $r}}
      foreach($p in $entry.rowHeights.PSObject.Properties){$r=$sheet.Rows.Item([int]$p.Name+1);try{$r.Hidden=($null -ne $p.Value -and $p.Value -eq 0);if($null -eq $p.Value){$r.RowHeight=$sheet.StandardHeight}elseif($p.Value -gt 0){$r.RowHeight=[Math]::Max(1,[double]$p.Value*3/4)}}finally{Release-Com $r}}
      foreach($range in $entry.merge){$r=$sheet.Range([string]$range);try{[void]$r.Merge()}finally{Release-Com $r}}
      foreach($name in $entry.deletedCharts){$object=$sheet.ChartObjects().Item([string]$name);try{[void]$object.Delete()}finally{Release-Com $object}}
      foreach($patch in $entry.charts){Set-Chart $sheet $patch}
      if(Has-Field $entry 'conditionalFormats'){Set-ConditionalFormats $sheet $entry.conditionalFormats}
      if($entry.PSObject.Properties.Name -contains 'autoFilter'){Set-AutoFilter $sheet $entry.autoFilter}
      foreach($index in $entry.manualHiddenRows){$sheet.Rows.Item([int]$index+1).Hidden=$true}
      foreach($p in $entry.rowHeights.PSObject.Properties){if($null -ne $p.Value -and $p.Value -eq 0){$sheet.Rows.Item([int]$p.Name+1).Hidden=$true}}
      if((Has-Field $entry 'freeze') -and $sheet.Visible -eq -1){[void]$sheet.Activate();$window=$book.Windows.Item(1);try{$window.FreezePanes=$false;$window.SplitRow=[int]$entry.freeze.rows;$window.SplitColumn=[int]$entry.freeze.cols;$window.FreezePanes=($entry.freeze.rows -gt 0 -or $entry.freeze.cols -gt 0)}finally{Release-Com $window}}
    }
    if($map.ContainsKey([string]$plan.activeSheetId) -and $map[[string]$plan.activeSheetId].Visible -eq -1){[void]$map[[string]$plan.activeSheetId].Activate()}
  }finally{foreach($s in $map.Values){Release-Com $s}}
}
$excel=$null;$book=$null;$owned=$false
try {
  $request=Get-Content -LiteralPath (Join-Path $JobDir 'request.json') -Raw -Encoding UTF8 | ConvertFrom-Json
  $sourcePath=Join-Path $JobDir $request.filename
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $zip=[IO.Compression.ZipFile]::OpenRead($sourcePath)
  try {if($zip.Entries | Where-Object {$_.FullName -match '(?i)(^|/)macrosheets/'}){throw 'Excel 4.0 (XLM) macro sheets are not supported.'}}finally{$zip.Dispose()}
  Write-State 'starting' 'Excel 작업 사본을 준비하고 있습니다.'
  Check-Cancel
  $before=@(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
  $excel=New-Object -ComObject Excel.Application
  Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class SheetStudioWindowOwner { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId); }'
  [uint32]$excelProcessId=0;[void][SheetStudioWindowOwner]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$excelProcessId)
  if(!$excelProcessId -or $before -contains [int]$excelProcessId){throw 'A separate Excel instance could not be verified. Existing Excel windows were not modified.'}
  $process=Get-Process -Id $excelProcessId;$owned=$true
  [IO.File]::WriteAllText((Join-Path $JobDir 'owner-next.json'),(@{processId=$excelProcessId;started=$process.StartTime.ToUniversalTime().Ticks.ToString()}|ConvertTo-Json -Compress),$utf8)
  Move-Item -LiteralPath (Join-Path $JobDir 'owner-next.json') -Destination (Join-Path $JobDir 'owner.json')
  Check-Cancel
  # Never lower macro security or enable VBA project access. All initial edits
  # occur with macros and events disabled, using the original OOXML package.
  $excel.AutomationSecurity=3;$excel.EnableEvents=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false
  $book=$excel.Workbooks.Open($sourcePath,0,$false)
  if($book.ReadOnly){throw 'Excel opened the working copy read-only. Check workbook protection or trust settings in Excel.'}
  Apply-Plan $book $request.plan
  # Preserve app pivot settings alongside the materialized XLSM result. Only our namespace is replaced.
  if($null -ne $request.plan.pivots){
    $parts=$book.CustomXMLParts.SelectByNamespace('urn:sheet-studio:pivots:1')
    for($i=$parts.Count;$i -ge 1;$i--){$part=$parts.Item($i);[void]$part.Delete();Release-Com $part};Release-Com $parts
    if(@($request.plan.pivots).Count -gt 0){
      $json=ConvertTo-Json -InputObject @($request.plan.pivots) -Depth 20 -Compress
      $xml='<pivots xmlns="urn:sheet-studio:pivots:1">'+[Security.SecurityElement]::Escape($json)+'</pivots>'
      $part=$book.CustomXMLParts.Add($xml);Release-Com $part
    }
  }
  [void]$book.Save()
  Preserve-Origin $sourcePath $request.zone
  if($request.mode -eq 'save'){
    Save-Result $book $request.zone
    Write-State 'completed' 'VBA 프로젝트를 유지한 XLSM을 저장했습니다.'
  } else {
    [void]$book.Close($false);Release-Com $book;$book=$null
    $excel.Visible=$true;$excel.DisplayAlerts=$true;$excel.EnableEvents=$true
    # ByUI respects the user's Trust Center/MOTW decisions on this explicit run.
    $excel.AutomationSecurity=2
    $book=$excel.Workbooks.Open($sourcePath,0,$false)
    Write-State 'ready' 'Excel 작업 사본을 열었습니다. 매크로를 실행한 뒤 결과 가져오기를 누르세요.'
    $finished=$false
    while(!$finished){
      Check-Cancel
      $commandPath=Join-Path $JobDir 'command.json'
      if(!(Test-Path -LiteralPath $commandPath)){Start-Sleep -Milliseconds 200;continue}
      $command=Get-Content -LiteralPath $commandPath -Raw -Encoding UTF8|ConvertFrom-Json
      Remove-Item -LiteralPath $commandPath
      try {
        switch($command.type){
          'run' {
            Write-State 'running' 'Excel에서 VBA 매크로를 실행하고 있습니다.'
            [void]$book.Activate()
            $qualified="'"+$book.Name.Replace("'","''")+"'!"+[string]$command.name
            [void]$excel.Run($qualified)
            Write-State 'ready' '매크로가 반환되었습니다. 결과 가져오기를 누르면 Sheet Studio에 반영됩니다.'
          }
          'dialog' {Write-State 'running' 'Excel의 매크로 창에서 실행할 매크로를 선택하세요.';[void]$book.Activate();[void]$excel.Dialogs.Item(17).Show();Write-State 'ready' 'Excel 매크로 창을 닫았습니다. 결과를 가져올 수 있습니다.'}
          'finish' {$excel.EnableEvents=$false;Save-Result $book $request.zone;Write-State 'completed' 'Excel의 실행 결과를 가져옵니다.';$finished=$true}
          'cancel' {Write-State 'cancelled' '작업 사본을 닫았습니다.';$finished=$true}
        }
      } catch {Write-State 'ready' ('Excel: '+$_.Exception.Message+' 매크로 이름과 Excel 보안 설정을 확인하세요. 보안 설정은 자동 변경하지 않습니다.')}
    }
  }
}catch{if(Test-Path -LiteralPath (Join-Path $JobDir 'cancel.flag')){Write-State 'cancelled' 'Excel 작업을 취소했습니다.'}else{[IO.File]::WriteAllText((Join-Path $JobDir 'error.log'),($_.Exception.ToString()+[Environment]::NewLine+$_.ScriptStackTrace),$utf8);Write-State 'error' $_.Exception.Message}}
finally {
  if($owned -and $excel){try{$excel.EnableEvents=$false;$excel.DisplayAlerts=$false;if($book){[void]$book.Close($false)};[void]$excel.Quit()}catch{}}
  Release-Com $book;Release-Com $excel
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
