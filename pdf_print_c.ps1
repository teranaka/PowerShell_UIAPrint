Param(
	[parameter(mandatory)][string]$PRINTERICONNAME,
	[parameter(mandatory)][string]$DATAPATH,
	[parameter(mandatory)][string]$TESTSCRIPTPATH
)
. ".\util\acrobat_print.ps1"
. ".\util\util.ps1"

# UIハイライト無効化
[UIAutomation.Preferences]::Highlight = $false

# 準備
$LASTEXITCODE = 99
$DATAPATH = Convert-Path $DATAPATH
$TESTSCRIPTPATH = Convert-Path $TESTSCRIPTPATH

New-Variable process 
New-Variable window
New-Variable acrobat_dialog
New-Variable dialog
$result = $False

# 開始コメント
Write-Host "----------------------" -ForegroundColor Cyan
Write-Host "$PSCommandPath" -ForegroundColor Cyan
Write-Host "$PRINTERICONNAME" -ForegroundColor Cyan
Write-Host "$DATAPATH" -ForegroundColor Cyan
Write-Host "$TESTSCRIPTPATH" -ForegroundColor Cyan
Write-Host "----------------------" -ForegroundColor Cyan


# ACROBATを開始し、印刷設定ダイアログを立ち上げる
try{
	$result = $False
	$result = (ACROBAT_Start $DATAPATH $PRINTERICONNAME ([ref]$process) ([ref]$window) ([ref]$acrobat_dialog) ([ref]$dialog))
}catch{
	$result = $False
}
if ($False -eq $result) {
	Write-Host "[FAILED] start Application" -ForegroundColor Red
	ACROBAT_Abort $process $window $acrobat_dialog $dialog
	exit 1
}

# 印刷設定にテストパターンの設定値を設定する
try{
	Write-Host "印刷設定"
	$result = $False
	$result = (&$TESTSCRIPTPATH $dialog)
} catch {
	$result = $False
}
if ($False -eq $result){
	Write-Host "[FAILED] call TEST_SCRIPT($TESTSCRIPTPATH)" -ForegroundColor Red
	ACROBAT_Abort $process $window $acrobat_dialog $dialog
	exit 1
}

# OKボタンを押下
try{
	$result = $False
	$result = (Click_ButtonOK $dialog)
}catch{
	$result = $False
}
if ($False -eq $result){
	Write-Host "[FAILED] click OK-BUTTON" -ForegroundColor Red
	ACROBAT_Abort $process $window $acrobat_dialog $dialog
	exit 1
}

# 印刷ボタンを押下
try{
	Write-Host "印刷する"
	$result = $False
	$result = (ACROBAT_Print $window $acrobat_dialog)
}catch{
	$result = $False
}
if ($False -eq $result){
	Write-Host "[FAILED] click PRINT-BUTTON" -ForegroundColor Red
	ACROBAT_Abort $process $window $acrobat_dialog $dialog
	exit 1
}

# アプリケーションを終了
ACROBAT_Exit $window

#プロセス終了待ち
$process.WaitForExit()

exit 0
