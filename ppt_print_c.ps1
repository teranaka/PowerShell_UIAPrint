Param(
	[parameter(mandatory)][string]$PRINTERICONNAME,
	[parameter(mandatory)][string]$DATAPATH,
	[parameter(mandatory)][string]$TESTSCRIPTPATH
)
. ".\util\powerpoint_print.ps1"
. ".\util\util.ps1"

# 準備
$LASTEXITCODE = 99
$DATAPATH = Convert-Path $DATAPATH
$TESTSCRIPTPATH = Convert-Path $TESTSCRIPTPATH

New-Variable process 
New-Variable window
New-Variable dialog
$result = $False

# 開始コメント
Write-Host "----------------------" -ForegroundColor Cyan
Write-Host "$PSCommandPath" -ForegroundColor Cyan
Write-Host "$PRINTERICONNAME" -ForegroundColor Cyan
Write-Host "$DATAPATH" -ForegroundColor Cyan
Write-Host "$TESTSCRIPTPATH" -ForegroundColor Cyan
Write-Host "----------------------" -ForegroundColor Cyan

# POWERPOINTを開始し、印刷設定ダイアログを立ち上げる
try{
	$result = $False
	$result = (POWERPOINT_Start $DATAPATH $PRINTERICONNAME ([ref]$process) ([ref]$window) ([ref]$dialog))
} catch {
	$result = $False
}
if ($False -eq $result) {
	Write-Host "[FAILED] start Application" -ForegroundColor Red
	POWERPOINT_Abort $process $window $dialog
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
	POWERPOINT_Abort $process $window $dialog
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
	POWERPOINT_Abort $process $window $dialog
	exit 1
}

# 印刷する
try{
	Write-Host "印刷する"
	$result = $False
	$result = (POWERPOINT_Print $window)
}catch{
	$result = $False
}
if ($False -eq $result){
	Write-Host "[FAILED] click PRINT-BUTTON" -ForegroundColor Red
	POWERPOINT_Abort $process $window $dialog
	exit 1
}

# アプリケーションを終了
POWERPOINT_Exit ($window)

# 印刷中断ダイアログが閉じるのを待つ
$process.WaitForExit()

exit 0
