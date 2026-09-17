param([string]$OutputPath='work/KC_Produktionsplan_dev.xlsm',[string]$SourcePath='work/source/KC_KUECHE_KOMPLETT_16-09-2026/MASTER_KW38_FUNKTIONIERT_16-09-2026.xlsm')
$ErrorActionPreference='Stop'
if ([IO.Path]::IsPathRooted($SourcePath)) { $taskSource=[IO.Path]::GetFullPath($SourcePath) } else { $taskSource=[IO.Path]::GetFullPath((Join-Path (Get-Location) $SourcePath)) }
if ([IO.Path]::IsPathRooted($OutputPath)) { $taskOut=[IO.Path]::GetFullPath($OutputPath) } else { $taskOut=[IO.Path]::GetFullPath((Join-Path (Get-Location) $OutputPath)) }
if ([string]::Equals($taskSource,$taskOut,[StringComparison]::OrdinalIgnoreCase)) { throw "MASTER-Quelle darf niemals Ziel des Umbaus sein" }
Copy-Item -LiteralPath $taskSource -Destination $taskOut -Force
$taskExcel=New-Object -ComObject Excel.Application
$taskExcel.Visible=$false
$taskExcel.DisplayAlerts=$false
$taskExcel.EnableEvents=$false
$taskExcel.AutomationSecurity=3
try {
 $taskBook=$taskExcel.Workbooks.Open($taskOut,0,$false)
 $taskProject=$taskBook.VBProject
 foreach($taskFile in Get-ChildItem (Join-Path $PSScriptRoot "../src/vba") -File){
  $taskCode=Get-Content -LiteralPath $taskFile.FullName -Raw -Encoding utf8
  $taskCode=($taskCode -split "`r?`n" | Where-Object {$_ -notmatch '^Attribute '}) -join "`r`n"
  $taskName=$taskFile.BaseName
  $taskComponent=$null
  try{$taskComponent=$taskProject.VBComponents.Item($taskName)}catch{}
  if($null -eq $taskComponent){$taskComponent=$taskProject.VBComponents.Add(1);$taskComponent.Name=$taskName}
  if($taskComponent.CodeModule.CountOfLines -gt 0){$taskComponent.CodeModule.DeleteLines(1,$taskComponent.CodeModule.CountOfLines)}
  $taskComponent.CodeModule.AddFromString($taskCode)
 }
 $taskBook.Save()
 $taskBook.Close($false)
 $taskExcel.AutomationSecurity=1
 $taskBook=$taskExcel.Workbooks.Open($taskOut,0,$false)
 $taskExcel.Run("'"+$taskBook.Name+"'!KC_Initialize")
 $taskBook.Worksheets('_KC_Config').Range('B2').Value2='AUS'
 $taskExcel.Run("'"+$taskBook.Name+"'!KC_Refresh")
 $taskBook.Save()
 Write-Output ('BUILT '+$taskOut)
 Write-Output ('Components='+$taskBook.VBProject.VBComponents.Count)
 $taskBook.Close($false)
} finally { $taskExcel.Quit();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($taskExcel) }



