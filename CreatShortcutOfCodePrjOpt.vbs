Sub creatShortcut()
	Dim sCurrentCode, sCurrentPrj, sCurrentOpt
	sCurrentCode = getElementValue(ID_INPUT_CODE_PATH_L1)
	sCurrentPrj = getElementValue(ID_INPUT_PROJECT_L1)
	sCurrentOpt = getElementValue(ID_INPUT_OPTION_L1)

	Call addShortcutButton(sCurrentCode, sCurrentPrj, sCurrentOpt)
End Sub

Sub applyShortcut(sCurrentCode, sCurrentPrj, sCurrentOpt)
	Call setElementValue(ID_INPUT_CODE_PATH_L1, sCurrentCode)
	Call onloadPrj()

	Call setElementValue(ID_INPUT_PROJECT_L1, sCurrentPrj)
	Call onloadOpt()
	
	Call setElementValue(ID_INPUT_OPTION_L1, sCurrentOpt)
End Sub