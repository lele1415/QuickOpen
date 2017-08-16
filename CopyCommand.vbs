Const ID_CHECKBOX_RM_OUT = "checkbox_rm_out"
Const ID_CHECKBOX_USER = "checkbox_user"
Const ID_CHECKBOX_BUILD_OTA = "checkbox_build_ota"

Sub CommandOfBuild()
    Dim userOrEng, optionName, commandOta, commandFinal
    userOrEng = "user"
    If Not element_isChecked(ID_CHECKBOX_USER) Then userOrEng = "eng"
    optionName = getElementValue(ID_INPUT_PROJECT)
    commandFinal = "./mk -o=TARGET_BUILD_VARIANT=" & userOrEng & " " & optionName & " n"
    commandOta = "./mk -o=TARGET_BUILD_VARIANT=" & userOrEng & " " & optionName & " otapackage"

    If element_isChecked(ID_CHECKBOX_RM_OUT) Then commandFinal = "rm -rf out/ && " & commandFinal
    If element_isChecked(ID_CHECKBOX_BUILD_OTA) Then commandFinal = commandFinal & " && " & commandOta

    Call CopyString(commandFinal)
End Sub

Sub CopyString(str)
    Clipboard="MsHta vbscript:ClipBoardData.setData(""Text"","""&str&""")(Window.Close)"  
    oWs.Run(Clipboard)
End Sub