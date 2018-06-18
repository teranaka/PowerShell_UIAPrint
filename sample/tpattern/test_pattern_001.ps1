Param( [parameter(mandatory)]$DIALOG )

# [原稿サイズ]設定
Write-Host "[原稿サイズ]設定"
if (-not (Set_ComboBox $DIALOG "原稿サイズ(D):" "A3 (297 x 420 mm)")){
	return $False
}

return $True
