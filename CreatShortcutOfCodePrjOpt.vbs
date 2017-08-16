Const ID_DIV_SHORTCUT = "div_shortcut"

Sub creatShortcut()
	Dim sCurrentCode, sCurrentPrj, sCurrentOpt
	sCurrentCode = getElementValue(ID_INPUT_CODE_PATH)
	sCurrentPrj = getElementValue(ID_INPUT_PROJECT)

	Call addShortcutButton(sCurrentCode, sCurrentPrj, ID_DIV_SHORTCUT)
End Sub

Sub applyShortcut(sCurrentCode, sCurrentPrj)
	Call setElementValue(ID_INPUT_CODE_PATH, sCurrentCode)
	Call setElementValue(ID_INPUT_PROJECT, sCurrentPrj)
End Sub