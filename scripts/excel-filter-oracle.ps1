param([Parameter(Mandatory=$true)][string]$InputPath,[Parameter(Mandatory=$true)][string]$OutputPath,[string]$ReadWorkbook,[string]$SaveWorkbook)
$ErrorActionPreference='Stop';$excel=$null;$book=$null
try {
  $data=Get-Content -LiteralPath $InputPath -Raw -Encoding UTF8 | ConvertFrom-Json
  $excel=New-Object -ComObject Excel.Application
  $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AutomationSecurity=3;$excel.EnableEvents=$false
  if($ReadWorkbook){$book=$excel.Workbooks.Open([IO.Path]::GetFullPath($ReadWorkbook),0,$true)}else{$book=$excel.Workbooks.Add()}
  $s=$book.Worksheets.Item(1)
  if(!$ReadWorkbook){
    $s.Name='Sheet1'
    for($r=0;$r -lt $data.rows.Count;$r++){for($c=0;$c -lt $data.rows[$r].Count;$c++){
      $v=$data.rows[$r][$c];if($null -ne $v){$cell=$s.Cells.Item($r+3,$c+2);if($v -is [string]){$cell.Value2=[string]$v}else{$cell.Value2=[double]$v}}
    }}
    $s.Range('D15').Formula='=SUBTOTAL(9,C4:C12)'
  }
  $results=@();if($data.range){$range=$s.Range([string]$data.range)}else{$range=$s.Range('B3:D12')}
  foreach($case in $data.cases){
    if(!$ReadWorkbook){
      if($s.AutoFilterMode){$s.AutoFilterMode=$false};[void]$range.AutoFilter()
      foreach($f in $case.filters){
        $v=$f.criteria1;if($v -is [array]){$v=[object[]]$v}
        $op=[Type]::Missing;if($null -ne $f.operator){$op=[int]$f.operator}
        $v2=[Type]::Missing;if($null -ne $f.criteria2){$v2=$f.criteria2}
        try{[void]$range.AutoFilter([int]$f.field,$v,$op,$v2)}catch{throw ($case.name+': '+$_.Exception.Message)}
      }
    }elseif($case.reapply){[void]$s.AutoFilter.ApplyFilter()}
    if($case.sort){$sort=$s.Sort;[void]$sort.SortFields.Clear();[void]$sort.SortFields.Add($s.Range('C4:C12'),0,2);[void]$sort.SetRange($range);$sort.Header=1;[void]$sort.Apply()}
    [void]$s.Calculate();$visible=@();for($r=4;$r -le 12;$r++){if(!$s.Rows.Item($r).Hidden){$visible+=($r-1)}}
    $records=@();for($r=4;$r -le 12;$r++){$records+=,@($s.Cells.Item($r,2).Value2,$s.Cells.Item($r,3).Value2,$s.Cells.Item($r,4).Value2)}
    $results+=@{name=$case.name;visible=$visible;records=$records;range=$s.AutoFilter.Range.Address();subtotal=$s.Range('D15').Value2;arrows=$s.AutoFilterMode;filterMode=$s.FilterMode}
  }
  @{excelVersion=$excel.Version;excelBuild=$excel.Build;cases=$results} | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
  if($SaveWorkbook){[void]$book.SaveAs([IO.Path]::GetFullPath($SaveWorkbook),51)}
  Write-Output ('Excel AutoFilter oracle recorded '+$results.Count+' cases.')
}finally{
  if($book){$book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
  if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
