

Const ID_INPUT_PROJECT_INFOS = "input_project_infos"
Const ID_BUTTON_SHOW_PROJECT_INFOS = "button_show_project_infos"
Const ID_BUTTON_CLOSE_PROJECT_INFOS = "button_close_project_infos"

Dim pMMIProjectConfigMk
Dim pDriverProjectConfigMk
Dim pDeviceProjectConfigMk

Function getDriverProjectConfigValue(keyStr)
    Dim value
    value = readTextAndGetValue(keyStr, pDriverProjectConfigMk)
    If value = "" Then
        value = readTextAndGetValue(keyStr, pDeviceProjectConfigMk)
    End If
    getDriverProjectConfigValue = value
End Function

Function getMMIProjectConfigValue(keyStr)
    Dim value
    value = readTextAndGetValue(keyStr, pMMIProjectConfigMk)
    If value = "" Then
        value = getDriverProjectConfigValue(keyStr)
    End If
    getMMIProjectConfigValue = value
End Function

Function getBootLogo()
    Call getProjectConfigMk()
    getBootLogo = getDriverProjectConfigValue("BOOT_LOGO=")
End Function

Sub getProjectConfigMk()
    pMMIProjectConfigMk = getProductPath & "/" & getProject() & "/config/ProjectConfig.mk"
    pDriverProjectConfigMk = getProductPath & "/" & Replace(getProject(), "-MMI", "") & "/config/ProjectConfig.mk"
    pDeviceProjectConfigMk = getSdkPath() & "/device/mediateksample/" & getProduct() & "/ProjectConfig.mk"
End Sub

Sub getProjectInfos()
    Call getProjectConfigMk()

    Dim str
    str = str & "BOOT_LOGO = " & getDriverProjectConfigValue("BOOT_LOGO=") & VbLf
    str = str & "LCM_WIDTH = " & getDriverProjectConfigValue("LCM_WIDTH=") & VbLf
    str = str & "LCM_HEIGHT = " & getDriverProjectConfigValue("LCM_HEIGHT=") & VbLf
    str = str & "CUSTOM_MODEM = " & getDriverProjectConfigValue("CUSTOM_MODEM=") & VbLf
    str = str & "BUILD_GMS = " & getMMIProjectConfigValue("BUILD_GMS=") & VbLf
    str = str & "BUILD_AGO_GMS = " & getMMIProjectConfigValue("BUILD_AGO_GMS=")

    Call hideElement(ID_BUTTON_SHOW_PROJECT_INFOS)
    Call showElement(ID_BUTTON_CLOSE_PROJECT_INFOS)

    Call showElement(ID_INPUT_PROJECT_INFOS)
    Call setElementValue(ID_INPUT_PROJECT_INFOS, str)
End Sub

Sub closeProjectInfos()
    Call hideElement(ID_INPUT_PROJECT_INFOS)
    Call hideElement(ID_BUTTON_CLOSE_PROJECT_INFOS)
    Call showElement(ID_BUTTON_SHOW_PROJECT_INFOS)
End Sub


