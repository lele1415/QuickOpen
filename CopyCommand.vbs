Const ID_CHECKBOX_RM_OUT_L1 = "checkbox_rm_out_l1"
Const ID_CHECKBOX_BUILD_OTA_L1 = "checkbox_build_ota_l1"
Const ID_CHECKBOX_RM_OUT_KK = "checkbox_rm_out_kk"
Const ID_CHECKBOX_USER_KK = "checkbox_user_kk"
Const ID_CHECKBOX_BUILD_OTA_KK = "checkbox_build_ota_kk"

Sub CommandOfBuildL1()
    Dim commandFinal, commandOta
    commandFinal = "make -j12 2>&1 | tee build.log"
    commandOta = "make -j12 otapackage 2>&1 | tee build_ota.log"

    If element_isChecked(ID_CHECKBOX_RM_OUT_L1) Then commandFinal = "rm -rf out/ && " & commandFinal
    If element_isChecked(ID_CHECKBOX_BUILD_OTA_L1) Then commandFinal = commandFinal & " && " & commandOta

    Call CopyString(commandFinal)
End Sub

Sub CommandOfBuildKK()
    Dim userOrEng, optionName, commandOta, commandFinal
    userOrEng = "user"
    If Not element_isChecked(ID_CHECKBOX_USER_KK) Then userOrEng = "eng"
    optionName = getElementValue(ID_INPUT_PROJECT_KK)
    commandFinal = "./mk -o=TARGET_BUILD_VARIANT=" & userOrEng & " " & optionName & " n"
    commandOta = "./mk -o=TARGET_BUILD_VARIANT=" & userOrEng & " " & optionName & " otapackage"

    If element_isChecked(ID_CHECKBOX_RM_OUT_KK) Then commandFinal = "rm -rf out/ && " & commandFinal
    If element_isChecked(ID_CHECKBOX_BUILD_OTA_KK) Then commandFinal = commandFinal & " && " & commandOta

    Call CopyString(commandFinal)
End Sub

Sub CopyString(str)
    Clipboard="MsHta vbscript:ClipBoardData.setData(""Text"","""&str&""")(Window.Close)"  
    oWs.Run(Clipboard)
End Sub