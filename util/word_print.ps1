Import-Module "..\UIAutomation\UIAutomation.dll"

# WINWORDを開始し、印刷設定ダイアログを立ち上げる
# $FILEPATH			[in]	DOCXファイルへのパス
# $DRVNAME			[in]	ドライバ名（プリンタアイコン名）
# result_process	[out]	WINWORDのProcessオブジェクトを取得
# result_window		[out]	WINWORDのメインウィンドウのWindowオブジェクトを取得
# result_dialog		[out]	WINWORDの印刷設定ダイアログのWindowオブジェクトを取得
function WINWORD_Start($FILEPATH, $DRVNAME, [ref]$result_process, [ref]$result_window, [ref]$result_dialog){

	Write-Verbose -Message"'Word'を起動する"

	# 拡張子docxの関連アプリケーションのパスを取得
	$assoc_val = cmd /c assoc .docx
	$assoc_command = cmd /c ftype $assoc_val.Split("=")[1]
	$app_path = $assoc_command.Split('"*"')[1]
	if ($app_path -notmatch "WINWORD.EXE\s*$"){
		Write-Warning -Message "拡張子'docx'に'WINWORD.EXE'が関連つけられていません"
		return $False
	}

	# ファイルのパスの確認
	if ( -not (Test-Path $FILEPATH)){
		Write-Warning -Message "$FILEPATHが見つかりません"
		return $False
	}

	# docxのアプリケーション（WORD）を起動
	Write-Verbose -Message"'$FILEPATH'を起動する"
	$app_process = Start-Process -FilePath $app_path -PassThru -ArgumentList "/q", $FILEPATH
	if ($app_process -eq $null){
		Write-Warning -Message "Wordを起動できませんでした"
		return $False
	}
	$result_process.value = $app_process
	
	# 最大10秒待機
	if (!$app_process.WaitForInputIdle(10000)) {
		Write-Warning -Message "入力可能状態待ちタイムアウト"
		return $False
	}

	# メインウィンドウを取得
	$window = Get-UIAWindow -Class 'OpusApp' -ProcessId $app_process.Id
	if ($window -eq $null){
		Write-Warning -Message "Windowを取得できない。処理を中断する"
		return $False
	}

	# 保護モードを解除
	Write-Verbose -Message"保護モードを解除"
	try {
		($window | Get-UiaButton -Class 'NetUISimpleButton' -Name '編集を有効にする(E)' -ErrorAction Stop | Invoke-UiaButtonClick) >$null
	} catch {}

	# Ctrl+Pで印刷メニューにジャンプ
	Start-Sleep -Millisecond 200
	$window.Keyboard.KeyDown([WindowsInput.Native.VirtualKeyCode]::CONTROL) >$null
	Start-Sleep -Millisecond 200
	$window.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::VK_P) >$null
	Start-Sleep -Millisecond 200
	$window.Keyboard.KeyUp([WindowsInput.Native.VirtualKeyCode]::CONTROL) >$null
	Start-Sleep -Millisecond 200

	# メインウィンドウを取得
	$window = Get-UiaWindow  -Class 'OpusApp' -ProcessId $app_process.Id
	if ($window -eq $null){
		Write-Warning -Message "Windowを取得できない。処理を中断する"
		return $False
	}
	$result_window.value = $window
	
	# 印刷タブを取得
	$tab = $window | Get-UiaGroup -Class 'NetUISlabContainer' -Name '印刷'
	if ($tab -eq $null){
		Write-Warning -Message "印刷タブを取得できない。処理を中断する"
		return $False
	}

	# 対象のプリンタドライバを選択
	Write-Verbose -Message"対象のプリンタドライバを選択"
	$group = $tab | Get-UiaGroup -Class 'NetUIElement' -Name 'プリンター'
	$combobox = $group | Get-UiaComboBox -Class 'NetUIDropdownAnchor' -Name '使用するプリンター'
	if ($combobox.Value -ne $DRVNAME) {
		Write-Verbose -Message"対象のプリンタに変更: $DRVNAME"
		try{
			$combobox | Invoke-UiaComboBoxExpand | Get-UiaListItem -Name $DRVNAME -ErrorAction Stop | Invoke-UiaListItemClick
		} catch {
	 		Write-Warning -Message "対象のプリンター名が見つからない。処理を中断する"
	 		return $False
		}
	}

	# 印刷設定ダイアログを起動する(Alt+p,r)
	Write-Verbose -Message"印刷設定ダイアログを起動する"
	$window.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::MENU) >$null
	Start-Sleep -Millisecond 200
	$window.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::VK_P) >$null
	Start-Sleep -Millisecond 200
	$window.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::VK_R) >$null
	Start-Sleep -Millisecond 200
	

	# 印刷設定ダイアログの取得
	try{
		($print_dialog = $window | Get-UiaChildWindow -Name "$DRVNAME*" -ErrorAction Stop) >$null
	} catch {
		Write-Warning -Message "印刷設定ダイアログが取得できません。処理を中断します"
		return $False
	}
	$result_dialog.value = $print_dialog
	return $True
}


# 印刷ボタンを押下する（印刷タブが開いている状態で使用すること）
# $window	[in]	WINWORDのメインウィンドウのWindowオブジェクト
function WINWORD_Print($window){
	if ($window -eq $null){
		return $False
	}
	
	# 印刷ボタンを押下(Alt+p,p)
	$window.Keyboard.KeyDown([WindowsInput.Native.VirtualKeyCode]::MENU) >$null
	Start-Sleep -Millisecond 200
	$window.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::VK_P) >$null
	Start-Sleep -Millisecond 200
	$window.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::VK_P) >$null
	Start-Sleep -Millisecond 200
	$window.Keyboard.KeyUp([WindowsInput.Native.VirtualKeyCode]::MENU) >$null
	
	# 印刷ボタン押下後の待ち
	if (!$window.WaitForInputIdle(10000)) {
		Write-Warning -Message "印刷ボタン押下後の入力可能状態待ちタイムアウト"
	}
	return $True
}


# ウィンドウを終了する
# $window	[in]	WINWORDのメインウィンドウのWindowオブジェクト
function WINWORD_Exit($window){
	if ($window -eq $null){
		return
	}
	
	# 入力可能状態待ち
	$window.WaitForInputIdle(10000) > $null
	
	# アプリケーションを閉じる(Ctrl+F4)
	Write-Verbose -Message"アプリケーションを閉じる"
	$window.Keyboard.KeyDown([WindowsInput.Native.VirtualKeyCode]::MENU) >$null
	$window.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::F4) >$null
	$window.Keyboard.KeyUp([WindowsInput.Native.VirtualKeyCode]::MENU) >$null
}


# WINWORDを強制終了する
# $process	[in]	WINWORDのProcessオブジェクト
# $window	[in]	WINWORDのメインウィンドウのWindowオブジェクト
# $dialog	[in]	WINWORDの印刷設定ダイアログのWindowオブジェクト
function WINWORD_Abort($process, $window, $dialog){
	Write-Verbose -Message"処理を中断し、アプリケーションを閉じる"
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

		WINWORD_Exit($window)
	}
	elseif($window -ne $null -and $window.WindowInteractionState -ne [System.Windows.Automation.WindowInteractionState]::Closing){
		WINWORD_Exit($window)
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
