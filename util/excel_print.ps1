Import-Module "..\UIAutomation\UIAutomation.dll"
add-type -AssemblyName microsoft.VisualBasic

# EXCELを開始し、印刷設定ダイアログを立ち上げる
# $FILEPATH			[in]	XLSXファイルへのパス
# $DRVNAME			[in]	ドライバ名（プリンタアイコン名）
# result_process	[out]	EXCELのProcessオブジェクトを取得
# result_window		[out]	EXCELのメインウィンドウのWindowオブジェクトを取得
# result_dialog		[out]	EXCELの印刷設定ダイアログのWindowオブジェクトを取得
function EXCEL_Start($FILEPATH, $DRVNAME, [ref]$result_process, [ref]$result_window, [ref]$result_dialog){

	Write-Verbose -Message"'Excel'を起動する"

	# 拡張子docxの関連アプリケーションのパスを取得
	$assoc_val = cmd /c assoc .xlsx
	$assoc_command = cmd /c ftype $assoc_val.Split("=")[1]
	$app_path = $assoc_command.Split('"*"')[1]
	if ($app_path -notmatch "EXCEL.EXE\s*$"){
		Write-Warning -Message "拡張子'xlsx'に'EXCEL.EXE'が関連つけられていません"
		return $False
	}

	# ファイルのパスの確認
	if ( -not (Test-Path $FILEPATH)){
		Write-Warning -Message "$FILEPATHが見つかりません"
		return $False
	}

	# アプリケーション（EXCEL）を起動
	Write-Verbose -Message"'$FILEPATH'を起動する"
	$app_process = Start-Process -FilePath $app_path -PassThru -ArgumentList "/e", "/r", $FILEPATH
	$result_process.value = $app_process
	
	# 最大10秒待機
	if ($False -eq $app_process.WaitForInputIdle(10000)) {
		Write-Warning -Message "入力可能状態待ちタイムアウト"
		return $False
	}
	
	# mainWindowの取得
	$window = Get-UIAWindow -Class 'XLMAIN' -ProcessId $app_process.Id
	if ($window -eq $null){
		Write-Warning -Message "Windowを取得できない。処理を中断する"
		return $False
	}

	# 保護モードを解除
	Write-Verbose -Message"保護モードを解除"
	try{
		$protect_button = Get-UiaButton -InputObject $window -Class 'NetUISimpleButton' -Name '編集を有効にする(E)' -ErrorAction Stop
		if ($protect_button -ne $null){
			try {
				$protect_button | Invoke-UiaButtonClick
			} catch {}
			
			# 保護モード解除するとWindowを作り直すため、Windowを再取得
			Start-Sleep -Millisecond 200
			$window = Get-UiaWindow -Class 'XLMAIN' -ProcessId $app_process.Id
			if ($window -eq $null){
				Write-Warning -Message "Windowを取得できない。処理を中断する"
				return $False
			}
		}
	} catch {}

	# 最大10秒待機
	if ($False -eq $window.WaitForInputIdle(30000)) {
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
	Start-Sleep -Millisecond 200

	# 印刷タブを取得
	$window = Get-UiaWindow  -Class 'XLMAIN' -ProcessId $app_process.Id
	if ($window -eq $null){
		Write-Warning -Message "Windowを取得できない。処理を中断する"
		return ＄False
	}
	$result_window.value = $window

	# 対象のプリンタドライバを選択
	Write-Verbose -Message"対象のプリンタドライバを選択"
	$tab = $window | Get-UiaGroup -Class 'NetUISlabContainer' -Name '印刷'
	$group = $tab | Get-UiaGroup -Class 'NetUIElement' -Name 'プリンター'
	$combobox = $group | Get-UiaComboBox -Class 'NetUIDropdownAnchor' -Name '使用するプリンター'
	if ($combobox.Value -ne $TARGET_DRV_NAME) {
		Write-Verbose -Message"対象のプリンタに変更: $TARGET_DRV_NAME"
		try{
			$combobox | Invoke-UiaComboBoxExpand | Get-UiaListItem -Name $DRVNAME -ErrorAction Stop | Invoke-UiaListItemClick
		} catch {
	 		Write-Warning -Message "対象のプリンター名が見つからない。処理を中断する"
	 		return $False
		}
	}

	# 対象のプリンタドライバを選択
	Write-Verbose -Message"'ブック全体を印刷'を選択"
	$tab = $window | Get-UiaGroup -Class 'NetUISlabContainer' -Name '印刷'
	$group = $tab | Get-UiaGroup -Class 'NetUIElement' -Name '設定'
	$combobox = $group | Get-UiaComboBox -Class 'NetUIDropdownAnchor' -Name '印刷対象'
	if ($combobox.Value -ne 'ブック全体を印刷') {
		try{
			$combobox | Invoke-UiaComboBoxExpand | Get-UiaListItem -Name 'ブック全体を印刷' -ErrorAction Stop | Invoke-UiaListItemClick
		} catch {
	 		Write-Warning -Message "'ブック全体を印刷'に設定できない。処理を中断する"
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
		($print_dialog = $window | Get-UiaChildWindow -Name "$TARGET_DRV_NAME*" -ErrorAction Stop) >$null
	} catch {
		Write-Warning -Message "印刷設定ダイアログが取得できません。処理を中断します"
		return $False
	}
	$result_dialog.value = $print_dialog
	return $True
}


# 印刷ボタンを押下する（印刷タブが開いている状態で使用すること）
# $window	[in]	EXCELのメインウィンドウのWindowオブジェクト
function EXCEL_Print($window){
	if ($window -eq $null){
		return
	}
	
	# 印刷ボタンを押下(Alt+p,p)
	$window.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::MENU) >$null
	Start-Sleep -Millisecond 200
	$window.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::VK_P) >$null
	Start-Sleep -Millisecond 200
	$window.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::VK_P) >$null
	Start-Sleep -Millisecond 200
	
	# 印刷ダイアログが出現するまでのウエイト
	Start-Sleep -Millisecond 5000

	# 印刷ダイアログのクローズ待ち
	Write-Verbose -Message"印刷ダイアログクローズ待ち"
	do{
		Start-Sleep -Millisecond 200
	}while($window.WindowInteractionState -ne [System.Windows.Automation.WindowInteractionState]::ReadyForUserInteraction)
	Write-Verbose -Message"印刷ダイアログクローズ待ち終了"
}


# ウィンドウを終了する
# $window	[in]	EXCELのメインウィンドウのWindowオブジェクト
function EXCEL_Exit($window){
	if ($window -eq $null){
		return
	}
	
	# アプリケーションを閉じる(Ctrl+F4)
	Write-Verbose -Message"アプリケーションを閉じる"
	$window.Close()
	
	Start-Sleep -Millisecond 1000
	
	# 保存ダイアログの確認（取得できなかったら無視）
	if ($window.WindowInteractionState -eq [System.Windows.Automation.WindowInteractionState]::BlockedByModalWindow){
		($cw = Get-UiaChildWindow -InputObject $window -Class 'NUIDialog') >$null
		if($cw -ne $null){
			# 保存しないボタン（ALT+N）の取得
			$bt = $cw.Children.Buttons | Where-Object {$_.Current.AccessKey -eq "ALT+N"}
			if($bt -ne $null){
				for([int]$i=0; $i -lt 10; $i++){
					# Windowをアクティブにする
					[Microsoft.VisualBasic.Interaction]::AppActivate($window.Current.ProcessId)
					
					# 保存しないボタン（ALT+N）を押下する
					(Invoke-UiaButtonClick -InputObject $bt) > $null
					
					if($window.WindowInteractionState -ne [System.Windows.Automation.WindowInteractionState]::BlockedByModalWindow){
						break;
					}
					Start-Sleep 200
				}
			}
		}
	}
}


# EXCELを強制終了する
# $process	[in]	EXCELのProcessオブジェクト
# $window	[in]	EXCELのメインウィンドウのWindowオブジェクト
# $dialog	[in]	EXCELの印刷設定ダイアログのWindowオブジェクト
function EXCEL_Abort($process, $window, $dialog){
	New-Variable tmp_process
	
	# Processが有効か確認（再取得）
	try{
		($tmp_process = Get-Process -Id $process.id -ErrorAction SilentlyContinue) >$null
	} catch {
		return
	}
	if($tmp_process -eq $null){
		return
	}

	# ダイアログが起動中の場合
	if( ($dialog -ne $null) -and ($dialog.WindowInteractionState -ne [System.Windows.Automation.WindowInteractionState]::Closing) ){
		Write-Verbose -Message"ダイアログを閉じる"
		do{
			($dialog.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::ESCAPE)) >$null
			Start-Sleep -Millisecond 200
		}while($dialog.WindowInteractionState -ne [System.Windows.Automation.WindowInteractionState]::Closing)

		EXCEL_Exit($window)
	}
	# Excelウィンドウが起動中の場合
	elseif($window -ne $null -and $window.WindowInteractionState -ne [System.Windows.Automation.WindowInteractionState]::Closing){
		EXCEL_Exit($window)
	}
	else{
		# 異常事態なので終了させる
		$tmp_process.Kill()
	}

	# プロセス終了待ち（Windowやダイアログを正常ルートで終了）
	if ($process.WaitForExit(60000)){
		return
	}

	# プロセスを強制終了（正常ルートで終了しない場合）
	$tmp_process.Kill()
	$process.WaitForExit()
}
