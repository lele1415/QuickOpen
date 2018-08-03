Const ID_COMMAND_ENG = "command_eng"
Const ID_COMMAND_USERDEBUG = "command_userdebug"
Const ID_COMMAND_USER = "command_user"

Const ID_COMMAND_RM_OUT = "command_rm_out"
Const ID_COMMAND_BUILD_OTA = "command_build_ota"

Sub CommandOfMake()
    Dim commandFinal, commandOta
    commandFinal = "make -j12 2>&1 | tee build.log"
    commandOta = "make -j12 otapackage 2>&1 | tee build_ota.log"

    If element_isChecked(ID_COMMAND_RM_OUT) Then commandFinal = "rm -rf out/ && " & commandFinal
    If element_isChecked(ID_COMMAND_BUILD_OTA) Then commandFinal = commandFinal & " && " & commandOta

    Call CopyString(commandFinal)
End Sub

Sub CommandOfLunch()
	Dim projectName, optionName, comboName, buildType, commandFinal

	projectName = getElementValue(ID_INPUT_PROJECT)
	optionName = getElementValue(ID_INPUT_OPTION)

	Select Case True
		Case element_isChecked(ID_COMMAND_ENG)
		    buildType = "eng"
		Case element_isChecked(ID_COMMAND_USERDEBUG)
		    buildType = "userdebug"
		Case element_isChecked(ID_COMMAND_USER)
		    buildType = "user"
	End Select

	comboName = "full_" & projectName & "-" & buildType

	commandFinal = "source build/envsetup.sh ; lunch " & comboName & " " & optionName
	Call CopyString(commandFinal)
End Sub

Sub CommandOfOut()
	Dim codePath, projectName, commandFinal

	codePath = getElementValue(ID_INPUT_CODE_PATH)
	projectName = getElementValue(ID_INPUT_PROJECT)

	commandFinal = codePath & "\out\target\product\" & projectName
	Call CopyString(commandFinal)
End Sub

Sub CopyString(str)
    Clipboard="MsHta vbscript:ClipBoardData.setData(""Text"","""&str&""")(Window.Close)"  
    oWs.Run(Clipboard)
End Sub