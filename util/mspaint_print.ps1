add-type -AssemblyName microsoft.VisualBasic

# MSPAINTを開始し、印刷設定ダイアログを立ち上げる
# $FILEPATH			[in]	XLSXファイルへのパス
# $DRVNAME			[in]	ドライバ名（プリンタアイコン名）
# result_process	[out]	MSPAINTのProcessオブジェクトを取得
# result_window		[out]	MSPAINTのメインウィンドウのWindowオブジェクトを取得
# result_dialog		[out]	MSPAINTの印刷ダイアログのWindowオブジェクトを取得
# result_dialog		[out]	ドライバ印刷設定ダイアログのWindowオブジェクトを取得
function MSPAINT_Start($FILEPATH, $DRVNAME, [ref]$result_process, [ref]$result_window, [ref]$result_app_dialog, [ref]$result_dialog){

	Write-Verbose -Message"'MSPAINT'を起動する"

	# アプリケーションのパスを取得
	#$app_path = [System.Environment]::ExpandEnvironmentVariables("%systemroot%\system32\mspaint.exe")
	$app_path = "mspaint.exe"

	# ファイルのパスの確認
	if ( -not (Test-Path $FILEPATH)){
		Write-Warning -Message "$FILEPATHが見つかりません"
		return $False
	}

	# アプリケーション（MSPAINT）を起動
	Write-Verbose -Message"'$FILEPATH'を起動する"
	$app_process = Start-Process -FilePath $app_path -PassThru -ArgumentList $FILEPATH
	if ($app_process -eq $null){
		return $False
	}
	$result_process.value = $app_process
	
	# 最大10秒待機
	if ($False -eq $app_process.WaitForInputIdle(10000)) {
		Write-Warning -Message "入力可能状態待ちタイムアウト"
		return $False
	}
	
	# mainWindowの取得
	$window = Get-UIAWindow -Class 'MSPaintApp' -ProcessId $app_process.Id
	if ($window -eq $null){
		Write-Warning -Message "Windowを取得できない。処理を中断する"
		return $False
	}
	$result_window.value = $window

	# Ctrl+Pで印刷メニューにジャンプ
	Start-Sleep -Millisecond 200
	$window.Keyboard.KeyDown([WindowsInput.Native.VirtualKeyCode]::CONTROL) >$null
	Start-Sleep -Millisecond 200
	$window.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::VK_P) >$null
	Start-Sleep -Millisecond 200
	$window.Keyboard.KeyUp([WindowsInput.Native.VirtualKeyCode]::CONTROL) >$null
	Start-Sleep -Millisecond 200
	
	# 最大10秒待機
	if ($False -eq $app_process.WaitForInputIdle(30000)) {
		Write-Warning -Message "入力可能状態待ちタイムアウト"
		return $False
	}

	# 印刷ダイアログを取得
	$dialog = Get-UiaChildWindow -InputObject $window -Class '#32770'
	if ($dialog -eq $null){
		Write-Warning -Message "印刷ダイアログが見つからない。処理を中断する" >$null
		return $False
	}
	$result_app_dialog.value = $dialog
	
	# 対象のプリンタドライバを選択
	Write-Verbose -Message"対象のプリンタドライバを選択"
	$list = Get-UiaList -InputObject $dialog -Name 'プリンターの選択'
	$selected_listitem = Get-UiaListSelection -InputObject $list
	if ($selected_listitem.Value -ne $DRVNAME){
		try{
			(Get-UiaListItem -InputObject $list -Name $DRVNAME -ErrorAction Stop| Invoke-UiaListItemSelectItem) >$null
		}catch{
			Write-Warning -Message "対象のプリンター名が見つからない。処理を中断する" >$null
			return $False
		}
	}

	# 印刷設定ダイアログを起動する(Alt+r)
	Write-Verbose -Message"印刷設定ダイアログを起動する"
	Start-Sleep -Millisecond 200
	$dialog.Keyboard.KeyDown([WindowsInput.Native.VirtualKeyCode]::MENU) >$null
	$dialog.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::VK_R) >$null
	$dialog.Keyboard.KeyUp([WindowsInput.Native.VirtualKeyCode]::MENU) >$null
	Start-Sleep -Millisecond 200
	

	# 印刷設定ダイアログの取得
	try{
		($driver_dialog = (Get-UiaChildWindow -InputObject $dialog -Regex -Name '印刷設定' -ErrorAction Stop)) >$null
	} catch {
		Write-Warning -Message "印刷設定ダイアログが取得できません。処理を中断します"
		return $False
	}
	if ($driver_dialog -eq $null){
		Write-Warning -Message "印刷設定ダイアログが取得できません。処理を中断します"
		return $False
	}
	$result_dialog.value = $driver_dialog

	return $True
}


# 印刷ボタンを押下する（印刷タブが開いている状態で使用すること）
# $window		[in]	MSPAINTのメインウィンドウのWindowオブジェクト
# $app_dialog	[in]	MSPAINTの印刷ダイアログオブジェクト
function MSPAINT_Print($window, $app_dialog){
	if ($app_dialog -eq $null){
		return
	}
	
	# 印刷ボタンを押下(Alt+p)
	Start-Sleep -Millisecond 200
	$app_dialog.Keyboard.KeyDown([WindowsInput.Native.VirtualKeyCode]::MENU) >$null
	$app_dialog.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::VK_P) >$null
	$app_dialog.Keyboard.KeyUp([WindowsInput.Native.VirtualKeyCode]::MENU) >$null
	Start-Sleep -Millisecond 200
	
	# 印刷中ダイアログが出現するまでのウエイト
	Start-Sleep -Millisecond 5000

	# 印刷中ダイアログのクローズ待ち
	Write-Verbose -Message"印刷中ダイアログクローズ待ち"
	do{
		Start-Sleep -Millisecond 200
	}while($window.WindowInteractionState -ne [System.Windows.Automation.WindowInteractionState]::ReadyForUserInteraction)
	Write-Verbose -Message"印刷中ダイアログクローズ待ち終了"
}


# ウィンドウを終了する
# $window	[in]	MSPAINTのメインウィンドウのWindowオブジェクト
function MSPAINT_Exit($window){
	if ($window -eq $null){
		return
	}
	
	# アプリケーションを閉じる(Ctrl+F4)
	Write-Verbose -Message"アプリケーションを閉じる"
	$window.Close()

	# 保存ダイアログ立ち上がり待ち
	Start-Sleep -Millisecond 200
	($window.WaitForInputIdle(3000)) >$null
	
	# 保存ダイアログの確認（取得できなかったら無視）
	if ($window.WindowInteractionState -eq [System.Windows.Automation.WindowInteractionState]::BlockedByModalWindow){
		($cw = Get-UiaChildWindow -InputObject $window -Class 'NUIDialog' -ErrorAction SilentlyContinue) >$null
		if($cw -ne $null){
			# 保存しない
			[Microsoft.VisualBasic.Interaction]::AppActivate($window.Current.ProcessId)
			$cw.Keyboard.KeyDown([WindowsInput.Native.VirtualKeyCode]::MENU) >$null
			$cw.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::VK_N) >$null
			$cw.Keyboard.KeyUp([WindowsInput.Native.VirtualKeyCode]::MENU) >$null
		}
	}
}


# MSPAINTを強制終了する
# $process	[in]	MSPAINTのProcessオブジェクト
# $window	[in]	MSPAINTのメインウィンドウのWindowオブジェクト
# $dialog	[in]	MSPAINTの印刷設定ダイアログのWindowオブジェクト
function MSPAINT_Abort($process, $window, $app_dialog, $dialog){
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

		Write-Verbose -Message"アプリケーションの印刷ダイアログを閉じる"
		do{
			($app_dialog.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::ESCAPE)) >$null
			Start-Sleep -Millisecond 200
		}while($app_dialog.WindowInteractionState -ne [System.Windows.Automation.WindowInteractionState]::Closing)

		MSPAINT_Exit($window)
	}
	# MSPAINTの印刷ダイアログが起動中の場合
	elseif( ($app_dialog -ne $null) -and ($app_dialog.WindowInteractionState -ne [System.Windows.Automation.WindowInteractionState]::Closing) ){
		Write-Verbose -Message"アプリケーションの印刷ダイアログを閉じる"
		do{
			($app_dialog.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::ESCAPE)) >$null
			Start-Sleep -Millisecond 200
		}while($app_dialog.WindowInteractionState -ne [System.Windows.Automation.WindowInteractionState]::Closing)

		MSPAINT_Exit($window)
	}
	# MSPAINTウィンドウが起動中の場合
	elseif($window -ne $null -and $window.WindowInteractionState -ne [System.Windows.Automation.WindowInteractionState]::Closing){
		MSPAINT_Exit($window)
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
