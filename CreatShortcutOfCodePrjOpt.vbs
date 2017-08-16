Const ID_DIV_SHORTCUT_L1 = "div_shortcut_l1"
Const ID_DIV_SHORTCUT_KK = "div_shortcut_kk"

Sub creatShortcutForL1()
	Dim sCurrentCode, sCurrentPrj, sCurrentOpt
	sCurrentCode = getElementValue(ID_INPUT_CODE_PATH_L1)
	sCurrentPrj = getElementValue(ID_INPUT_PROJECT_L1)
	sCurrentOpt = getElementValue(ID_INPUT_OPTION_L1)

	Call addShortcutButtonForL1(sCurrentCode, sCurrentPrj, sCurrentOpt, ID_DIV_SHORTCUT_L1)
End Sub

Sub applyShortcutForL1(sCurrentCode, sCurrentPrj, sCurrentOpt)
	Call setElementValue(ID_INPUT_CODE_PATH_L1, sCurrentCode)
	Call onloadPrj()

	Call setElementValue(ID_INPUT_PROJECT_L1, sCurrentPrj)
	Call onloadOpt()
	
	Call setElementValue(ID_INPUT_OPTION_L1, sCurrentOpt)
End Sub

Sub creatShortcutForKK()
	Dim sCurrentCode, sCurrentPrj, sCurrentOpt
	sCurrentCode = getElementValue(ID_INPUT_CODE_PATH_KK)
	sCurrentPrj = getElementValue(ID_INPUT_PROJECT_KK)

	Call addShortcutButtonForKK(sCurrentCode, sCurrentPrj, ID_DIV_SHORTCUT_KK)
End Sub

Sub applyShortcutForKK(sCurrentCode, sCurrentPrj)
	Call setElementValue(ID_INPUT_CODE_PATH_KK, sCurrentCode)
	Call setElementValue(ID_INPUT_PROJECT_KK, sCurrentPrj)
End Sub