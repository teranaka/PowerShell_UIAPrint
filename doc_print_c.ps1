Param(
	[parameter(mandatory)][string]$PRINTERICONNAME,
	[parameter(mandatory)][string]$DATAPATH,
	[parameter(mandatory)][string]$TESTSCRIPTPATH
)

Import-Module "C:\devel\UIAutomation\UIAutomation.dll"
. ".\util\word_print.ps1"
. ".\util\util.ps1"

# UIハイライト無効化
[UIAutomation.Preferences]::Highlight = $false

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


# WINWORDを開始し、印刷設定ダイアログを立ち上げる
try{
	$result = $False
	$result = WINWORD_Start $DATAPATH $PRINTERICONNAME ([ref]$process) ([ref]$window) ([ref]$dialog)
} catch {
	$result = $False
}
if ($False -eq $result ) {
	Write-Host "[FAILED] start Application" -ForegroundColor Red
	WINWORD_Abort $process $window $dialog
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
	WINWORD_Abort $process $window $dialog
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
	WINWORD_Abort $process $window $dialog
	exit 1
}

# 印刷ボタンを押下(Alt+p,p)
try{
	Write-Host "印刷する"
	$result = $False
	$result = (WINWORD_Print $window)
}catch{
	$result = $False
}
if ($False -eq $result){
	Write-Host "[FAILED] click PRINT-BUTTON" -ForegroundColor Red
	WINWORD_Abort $process $window $dialog
	exit 1
}

# アプリケーションを終了(印刷中に終了させると印刷中断ダイアログが起動する)
WINWORD_Exit($window)

# Wordが終了するのを待つ
$process.WaitForExit()

exit 0
