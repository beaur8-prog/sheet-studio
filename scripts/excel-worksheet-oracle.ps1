param([Parameter(Mandatory=$true)][string]$WorkbookPath,[Parameter(Mandatory=$true)][string]$OutputPath)
$ErrorActionPreference='Stop';$excel=$null;$book=$null
try {
  $excel=New-Object -ComObject Excel.Application
  $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AutomationSecurity=3;$excel.EnableEvents=$false;$excel.AskToUpdateLinks=$false
  $book=$excel.Workbooks.Open([IO.Path]::GetFullPath($WorkbookPath),0,$true)
  $s=$book.Worksheets.Item(1);[void]$s.Activate();[void]$s.Calculate()
  $cells=@{};foreach($key in @('A2','A3','B4','C4','D3','D6')){$cell=$s.Range($key);$cells[$key]=@{value=$cell.Value2;background=$cell.DisplayFormat.Interior.Color;color=$cell.DisplayFormat.Font.Color}}
  $view=@{rows=$excel.ActiveWindow.SplitRow;cols=$excel.ActiveWindow.SplitColumn;frozen=$excel.ActiveWindow.FreezePanes;hiddenRow=[bool]$s.Rows.Item(9).Hidden;hiddenCol=[bool]$s.Columns.Item(6).Hidden}
  $s.Range('B3').Value2=1.0;$s.Range('B4').Value2=5.0;[void]$s.Calculate()
  $after=@{};foreach($key in @('B4','C4','D3')){$cell=$s.Range($key);$after[$key]=@{background=$cell.DisplayFormat.Interior.Color;color=$cell.DisplayFormat.Font.Color}}
  # Controlled Excel paste-special and relative-range assignment observations.
  [void]$s.Range('B2:C3').Copy();[void]$s.Range('J2').PasteSpecial(-4163,-4142,$false,$true)
  $paste=@($s.Range('J2').Value2,$s.Range('K2').Value2,$s.Range('J3').Value2,$s.Range('K3').Value2)
  $s.Range('G2:H3').FormulaR1C1='=RC[-5]*2';$fill=@($s.Range('G2').Formula,$s.Range('H3').Formula)
  @{version=$excel.Version;build=$excel.Build;view=$view;cells=$cells;after=$after;paste=$paste;fill=$fill} | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
  Write-Output 'Excel worksheet behavior oracle recorded.'
} finally {
  if($book){$book.Close($false);[void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)}
  if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)}
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
