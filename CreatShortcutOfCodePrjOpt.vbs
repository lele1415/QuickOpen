Const ID_DIV_SHORTCUT = "div_shortcut"

Sub creatShortcut()
	Dim sCurrentCode, sCurrentPrj, sCurrentOpt
	sCurrentCode = getElementValue(ID_INPUT_CODE_PATH)
	sCurrentPrj = getElementValue(ID_INPUT_PROJECT)
	sCurrentOpt = getElementValue(ID_INPUT_OPTION)

	Call addShortcutButton(sCurrentCode, sCurrentPrj, sCurrentOpt, ID_DIV_SHORTCUT)
End Sub

Sub applyShortcut(sCurrentCode, sCurrentPrj, sCurrentOpt)
	Call setElementValue(ID_INPUT_CODE_PATH, sCurrentCode)
	Call onloadPrj()

	Call setElementValue(ID_INPUT_PROJECT, sCurrentPrj)
	Call onloadOpt()
	
	Call setElementValue(ID_INPUT_OPTION, sCurrentOpt)
End Sub