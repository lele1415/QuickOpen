Const ID_CHECKBOX_RM_OUT = "checkbox_rm_out"
Const ID_CHECKBOX_BUILD_OTA = "checkbox_build_ota"

Sub CommandOfBuild()
    Dim commandFinal, commandOta
    commandFinal = "make -j12 2>&1 | tee build.log"
    commandOta = "make -j12 otapackage 2>&1 | tee build_ota.log"

    If element_isChecked(ID_CHECKBOX_RM_OUT) Then commandFinal = "rm -rf out/ && " & commandFinal
    If element_isChecked(ID_CHECKBOX_BUILD_OTA) Then commandFinal = commandFinal & " && " & commandOta

    Call CopyString(commandFinal)
End Sub

Sub CopyString(str)
    Clipboard="MsHta vbscript:ClipBoardData.setData(""Text"","""&str&""")(Window.Close)"  
    oWs.Run(Clipboard)
End Sub