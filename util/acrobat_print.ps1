Import-Module "..\UIAutomation\UIAutomation.dll"

# ACROBATを開始し、印刷設定ダイアログを立ち上げる
# $FILEPATH				[in]	PDFファイルへのパス
# $DRVNAME				[in]	ドライバ名（プリンタアイコン名）
# result_process		[out]	AcrobatのProcessオブジェクトを取得
# result_window			[out]	AcrobatのメインウィンドウのWindowオブジェクトを取得
# result_app_dialog		[out]	Acrobatの印刷設定ダイアログのWindowオブジェクトを取得
# result_dialog			[out]	ドライバの印刷設定ダイアログのWindowオブジェクトを取得
function ACROBAT_Start($FILEPATH, $DRVNAME, [ref]$result_process, [ref]$result_window, [ref]$result_app_dialog, [ref]$result_dialog){

	# 拡張子pdfの関連アプリケーションのパスを取得
	$assoc_val = cmd /c assoc .pdf
	$assoc_command = cmd /c ftype $assoc_val.Split("=")[1]
	$app_path = $assoc_command.Split('"*"')[1]
	if ($app_path -notmatch "AcroRd32.EXE\s*$"){
		Write-Warning -Message "拡張子'pdf'に'AcroRd32.EXE'が関連つけられていません" 
		return $False
	}

	# ファイルのパスの確認
	if ( -not (Test-Path $FILEPATH)){
		Write-Warning -Message "$FILEPATHが見つかりません"
		return $False
	}

	# pdfのアプリケーション（Acrobat）を起動
	Write-Verbose -Message"'$FILEPATH'を起動する"
	$app_process = Start-Process -FilePath $app_path -PassThru -ArgumentList $FILEPATH
	if ($app_process -eq $null){
		Write-Warning -Message "AcrobatのProcessを取得できません"
		return $False
	}
	$result_process.value = $app_process

	# 最大10秒待機
	if (!$app_process.WaitForInputIdle(10000)) {
		Write-Warning -Message "入力可能状態待ちタイムアウト"
		return $False
	}
	Start-Sleep -Millisecond 200
	
	# Acrobatはプロセスが2つ起動し、2つ目がWindowを持つため、子プロセスを取得する
	try {
		$app_process = Get-Process -pid (Get-WmiObject -Class Win32_Process | Where {$_.ParentProcessId -eq $app_process.Id}).ProcessId
	} catch {
		return $False
	}
	$result_process.value = $app_process

	# 何故か立ち上がるダイアログを閉じる（たまにショートカットキーがフックされて邪魔されるので）
	try{
		$app_process.WaitForInputIdle(10000) >$null
		$temp = Get-UiaWindow -ProcessId $app_process.Id -Title 'タグ付けされていない文書の読み上げ*'
		$temp.Close()
	} catch {}

	# Acrobatのメインウィンドウ取得（
	$window = Get-UiaWindow -ProcessId $app_process.Id
	if ($window -eq $null){
		Write-Warning -Message "AcrobatのWindowを取得できません"
		return $False
	}
	$result_window.value = $window

	# 最大10秒待機
	if (!$app_process.WaitForInputIdle(10000)) {
		Write-Warning -Message "入力可能状態待ちタイムアウト"
		return $False
	}

	# Ctrl+Pで印刷メニューにジャンプ
	Start-Sleep -Millisecond 200
	$window.Keyboard.KeyDown([WindowsInput.Native.VirtualKeyCode]::CONTROL) >$null
	Start-Sleep -Millisecond 200
	$window.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::VK_P) >$null
	Start-Sleep -Millisecond 200
	$window.Keyboard.KeyUp([WindowsInput.Native.VirtualKeyCode]::CONTROL) >$null

	# Acrobatの印刷ダイアログ
	Start-Sleep -Millisecond 200
	$acro_dialog = Get-UiaWindow -Class '#32770' -Name '印刷'
	if($acro_dialog -eq $null){
		Write-Warning -Message "Acrobatの印刷ダイアログを取得できません"
		return $False
	}
	$result_app_dialog.value = $acro_dialog
	
  	# プリンターの選択コンボボックスの取得 (なぜかコンボボックス名が部数)
	$combobox = Get-UiaComboBox -InputObject $acro_dialog -Class 'ComboBox' -Name '部数(C) :'
	if ($combobox -eq $null){
 		Write-Warning -Message "プリンター設定が見つからない。処理を中断する"
 		return $False
	}
	if ($combobox.Value -ne $DRVNAME) {
		Write-Verbose -Message"対象のプリンタに変更: $DRVNAME"
		try{
			($combobox | Invoke-UiaComboBoxExpand | Get-UiaListItem -Name $DRVNAME -ErrorAction Stop | Invoke-UiaListItemClick) >$null
		} catch {
	 		Write-Warning -Message "対象のプリンター名が見つからない。処理を中断する"
	 		return $False
		}
	}


	# 印刷設定ダイアログを起動する(Alt+p)
	Write-Verbose -Message"印刷設定ダイアログを起動する"
	Start-Sleep -Millisecond 200
	$acro_dialog.Keyboard.KeyDown([WindowsInput.Native.VirtualKeyCode]::MENU) >$null
	Start-Sleep -Millisecond 200
	$acro_dialog.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::VK_P) >$null
	Start-Sleep -Millisecond 200
	$acro_dialog.Keyboard.KeyUp([WindowsInput.Native.VirtualKeyCode]::MENU) >$null

	# 印刷設定ダイアログの取得
	Start-Sleep -Millisecond 200
	try{
		($print_dialog = Get-UiaWindow -Name "印刷" -ErrorAction Stop) >$null
	} catch {
		Write-Warning -Message "印刷設定ダイアログが取得できません。処理を中断します"
		return $False
	}
	$result_dialog.value = $print_dialog
	return $True
}


# 印刷ボタンを押下する（印刷ダイアログが開いている状態で使用すること）
# $window		[in]	AcrobatのメインウィンドウのWindowオブジェクト
# $app_dialog	[in]	Acrobatの印刷ダイアログのWindowオブジェクト
function ACROBAT_Print($window, $app_dialog){
	if ($window -eq $null){
		return
	}
	if ($app_dialog -eq $null){
		return
	}
	
	# 印刷ボタンを押下
	Start-Sleep -Millisecond 200
	try{
		$app_dialog | Get-UiaButton -Class 'Button' -Name '印刷' | Invoke-UiaButtonClick > $null
	} catch {
		Write-Warning -Message "印刷ボタンが取得できません。処理を中断します"
		return
	}

	# 印刷ダイアログが出現するまでのウエイト
	Start-Sleep -Millisecond 5000

	# 印刷ダイアログクローズ待ち
	Write-Verbose -Message"印刷ダイアログクローズ待ち"
	do{
		Start-Sleep -Millisecond 200
	}while($window.WindowInteractionState -ne [System.Windows.Automation.WindowInteractionState]::ReadyForUserInteraction)
	Write-Verbose -Message"印刷ダイアログクローズ待ち終了"
}


# ウィンドウを終了する
function ACROBAT_Exit($window){
	if ($window -eq $null){
		return
	}

	# アプリケーションを閉じる(Alt+F4)
	Write-Verbose -Message"アプリケーションを閉じる"
	Start-Sleep -Millisecond 200
	$window.Close()
	#$window.Keyboard.KeyDown([WindowsInput.Native.VirtualKeyCode]::MENU) >$null
	#Start-Sleep -Millisecond 200
	#$window.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::F4) >$null
	#Start-Sleep -Millisecond 200
	#$window.Keyboard.KeyUp([WindowsInput.Native.VirtualKeyCode]::MENU) >$null
}

# アプリケーションを強制終了する
# process		[in]	AcrobatのProcessオブジェクトを取得
# window		[in]	AcrobatのメインウィンドウのWindowオブジェクトを取得
# app_dialog	[in]	Acrobatの印刷設定ダイアログのWindowオブジェクトを取得
# dialog		[in]	ドライバの印刷設定ダイアログのWindowオブジェクトを取得
function ACROBAT_Abort($process, $window, $app_dialog, $dialog){
	New-Variable tmp_process
	try{
		($tmp_process = Get-Process -Id $process.id -ErrorAction SilentlyContinue) >$null
	} catch {
		return
	}
	if($tmp_process -eq $null){
		return
	}

	if( ($dialog -ne $null) -and ($dialog.WindowInteractionState -ne [System.Windows.Automation.WindowInteractionState]::Closing) ){
		Write-Verbose -Message"ダイアログを閉じる"
		do{
			($dialog.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::ESCAPE)) >$null
			Start-Sleep -Millisecond 200
		}while($dialog.WindowInteractionState -ne [System.Windows.Automation.WindowInteractionState]::Closing)

		Write-Verbose -Message"アプリケーションの印刷ダイアログを閉じる"
		do{
			($app_dialog.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::ESCAPE)) >$null
			Start-Sleep -Millisecond 200
		}while($app_dialog.WindowInteractionState -ne [System.Windows.Automation.WindowInteractionState]::Closing)
		
		ACROBAT_Exit($window)
	}
	elseif( ($app_dialog -ne $null) -and ($app_dialog.WindowInteractionState -ne [System.Windows.Automation.WindowInteractionState]::Closing) ){
		Write-Verbose -Message"アプリケーションの印刷ダイアログを閉じる"
		do{
			($app_dialog.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::ESCAPE)) >$null
			Start-Sleep -Millisecond 200
		}while($app_dialog.WindowInteractionState -ne [System.Windows.Automation.WindowInteractionState]::Closing)

		ACROBAT_Exit($window)
	}
	elseif($window -ne $null -and $window.WindowInteractionState -ne [System.Windows.Automation.WindowInteractionState]::Closing){
		ACROBAT_Exit($window)
	}
	else{
		$tmp_process.Kill()
	}

	# 印刷中断ダイアログが閉じるのを待つ
	if($process.WaitForExit(30000)){
		return
	}

	$tmp_process.Kill()
	$process.WaitForExit()
}
