# 指定されたコンボボックスに値を設定する
# $dialog			[in]	WINWORDの印刷設定ダイアログのWindowオブジェクト
# $combobox_name	[in]	設定するコンボボックス名の文字列
# $value			[in]	コンボボックスに設定する値の文字列
function Set_ComboBox($dialog, $combobox_name, $value){
	New-Variable combobox
	try{
		$combobox = Get-UiaComboBox -InputObject $dialog -Class 'ComboBox' -Name $combobox_name -ErrorAction Stop
	} catch {
		Write-Warning -Message "'$combobox_name' コンボボックスが見つかりません"
		return $False
	}

	if ($value -eq $combobox.Value){
		return $True
	}

	try{
		($combobox | Invoke-UiaComboBoxExpand | Get-UiaListItem -Name $value -ErrorAction Stop | Invoke-UiaListItemClick) >$null
	} catch {
		Write-Warning -Message "'$combobox_name' コンボボックスに、値 '$value' が存在しません"
		return $False
	}
	return $True
}


# 現在のダイアログのOKボタンを押下する
# $dialog	[in]	ダイアログのWindowオブジェクト
function Click_ButtonOK($dialog){
	try{
		($button = Get-UiaButton -InputObject $dialog -AutomationId '1' -Class 'Button' -Name 'OK' -ErrorAction Stop ) >$null
		Invoke-UiaButtonClick -InputObject $button -ErrorAction Stop >$null
	} catch {
		return $False
	}
	
	return $True
}