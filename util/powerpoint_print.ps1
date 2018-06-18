Import-Module "..\UIAutomation\UIAutomation.dll"
add-type -AssemblyName microsoft.VisualBasic

# POWERPOINTを開始し、印刷設定ダイアログを立ち上げる
# $FILEPATH			[in]	XLSXファイルへのパス
# $DRVNAME			[in]	ドライバ名（プリンタアイコン名）
# result_process	[out]	POWERPOINTのProcessオブジェクトを取得
# result_window		[out]	POWERPOINTのメインウィンドウのWindowオブジェクトを取得
# result_dialog		[out]	POWERPOINTの印刷設定ダイアログのWindowオブジェクトを取得
function POWERPOINT_Start($FILEPATH, $DRVNAME, [ref]$result_process, [ref]$result_window, [ref]$result_dialog){

	Write-Verbose -Message"'PowerPoint'を起動する"

	# 拡張子docxの関連アプリケーションのパスを取得
	$assoc_val = cmd /c assoc .pptx
	$assoc_command = cmd /c ftype $assoc_val.Split("=")[1]
	$app_path = $assoc_command.Split('"*"')[1]
	if ($app_path -notmatch "POWERPNT.EXE\s*$"){
		Write-Warning -Message "拡張子'docx'に'POWERPNT.EXE'が関連つけられていません"
		return $False
	}
	
	# ファイルのパスの確認
	if ( -not (Test-Path $FILEPATH)){
		Write-Warning -Message "$FILEPATHが見つかりません"
		return $False
	}
	
	# アプリケーション（POWERPOINT）を起動
	Write-Verbose -Message"'$FILEPATH'を起動する"
	$app_process = Start-Process -FilePath $app_path -PassThru -ArgumentList $FILEPATH
	$result_process.value = $app_process
	
	# 最大10秒待機
	if ($False -eq $app_process.WaitForInputIdle(10000)) {
		Write-Warning -Message "入力可能状態待ちタイムアウト"
		return $False
	}
	
	# mainWindowの取得
	$window = Get-UIAWindow -Class 'PPTFrameClass' -ProcessId $app_process.Id
	if ($window -eq $null){
		Write-Warning -Message "Windowを取得できない。処理を中断する"
		return
	}
	$result_window.value = $window

	# オプションをバッググラウンド印刷に変更
	# オプションを開く
	Write-Verbose -Message"オプション設定変更"
	Start-Sleep -Millisecond 200
	$window.Keyboard.KeyDown([WindowsInput.Native.VirtualKeyCode]::MENU) >$null
	Start-Sleep -Millisecond 200
	$window.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::VK_F) >$null
	Start-Sleep -Millisecond 200
	$window.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::VK_T) >$null
	Start-Sleep -Millisecond 200
	$window.Keyboard.KeyUp([WindowsInput.Native.VirtualKeyCode]::MENU) >$null
	Start-Sleep -Millisecond 200

	# オプションダイアログを取得
	$option_dialog = Get-UIAChildWindow -InputObject $window -Class 'NUIDialog' -Name 'PowerPoint のオプション'
	if ($option_dialog -ne $null){
		# 詳細設定画面に切り替え
		try{
			(Get-UiaList -InputObject $option_dialog | Get-UiaListItem -Class 'NetUIListViewItem' -Name '詳細設定' -ErrorAction Stop| Invoke-UiaListItemClick -ErrorAction Stop) >$null
		} catch {
			Write-Warning -Message "オプション設定を変更できませんでした"
			$option_dialog.Close()
			return $False
		}
		
		# バックグラウンド印刷の設定状態を取得
		($checkbox = Get-UiaPane -InputObject $option_dialog -Name '詳細設定' -ErrorAction SilentlyContinue | Get-UiaCheckBox -Name 'バックグラウンドで印刷する' -ErrorAction SilentlyContinue) >$null
		if ($checkbox -eq $null){
			Write-Warning -Message "オプション設定を変更できませんでした"
			$option_dialog.Close()
			return $False
		}

		# バックグラウンドで印刷をOffにする
		While ($checkbox.ToggleState -ne 'Off'){
			Write-Verbose -Message"オプション設定を'バックグラウンド印刷しない'に変更"
			(Set-UiaCheckBoxToggleState -InputObject $checkbox $False) >$null
		}
		
		# OKボタンを押下してクローズ
		(Get-UiaButton -InputObject $option_dialog -Name 'OK' | Invoke-UiaButtonClick) >$null
		
		# もしクローズされていなかったら、強制終了
		Start-Sleep -Millisecond 200
		if ($option_dialog.WindowInteractionState -ne [System.Windows.Automation.WindowInteractionState]::Closing){
			$option_dialog.Close()
		}
	}
	
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
	$tab = $window | Get-UiaGroup -Class 'NetUISlabContainer' -Name '印刷'
	if ($tab -eq $null){
		Write-Warning -Message "印刷タブを取得できない。処理を中断する"
		return ＄False
	}

	# 対象のプリンタドライバを選択
	Write-Verbose -Message"対象のプリンタドライバを選択"
	$group = $tab | Get-UiaGroup -Class 'NetUIElement' -Name 'プリンター'
	$combobox = $group | Get-UiaComboBox -Class 'NetUIDropdownAnchor' -Name '使用するプリンター'
	if ($combobox.Value -ne $TARGET_DRV_NAME) {
		Write-Verbose -Message"対象のプリンタに変更: $TARGET_DRV_NAME"
		try{
			($combobox | Invoke-UiaComboBoxExpand | Get-UiaListItem -Name $DRVNAME -ErrorAction Stop | Invoke-UiaListItemClick) >$null
		} catch {
	 		Write-Warning -Message "対象のプリンター名が見つからない。処理を中断する" -ForegroundColor Red >$null
	 		return $False
		}
	}

	# 印刷対象を選択
	Write-Verbose -Message"'すべてのスライドを印刷'を選択"
	$group = Get-UiaGroup -InputObject $tab -Class 'NetUIElement' -Name '設定'
	$combobox = Get-UiaComboBox -InputObject $group -Class 'NetUIDropdownAnchor' -Name '印刷対象'
	if ($combobox.Value -ne 'ブック全体を印刷') {
		try{
			($combobox | Invoke-UiaComboBoxExpand | Get-UiaListItem -Name 'すべてのスライドを印刷' -ErrorAction Stop | Invoke-UiaListItemClick) >$null
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
# $window	[in]	POWERPOINTのメインウィンドウのWindowオブジェクト
function POWERPOINT_Print($window){
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
# $window	[in]	POWERPOINTのメインウィンドウのWindowオブジェクト
function POWERPOINT_Exit($window){
	if ($window -eq $null){
		return
	}
	
	# アプリケーションを閉じる(Ctrl+F4)
	Write-Verbose -Message"アプリケーションを閉じる"
	$window.Close()

	# 保存ダイアログ立ち上がり待ち
	Start-Sleep -Millisecond 2000
	($window.WaitForInputIdle(10000)) >$null
	
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


# POWERPOINTを強制終了する
# $process	[in]	POWERPOINTのProcessオブジェクト
# $window	[in]	POWERPOINTのメインウィンドウのWindowオブジェクト
# $dialog	[in]	POWERPOINTの印刷設定ダイアログのWindowオブジェクト
function POWERPOINT_Abort($process, $window, $dialog){
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
		
		POWERPOINT_Exit($window)
	}
	# POWERPOINTウィンドウが起動中の場合
	elseif($window -ne $null -and $window.WindowInteractionState -ne [System.Windows.Automation.WindowInteractionState]::Closing){
		POWERPOINT_Exit($window)
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
