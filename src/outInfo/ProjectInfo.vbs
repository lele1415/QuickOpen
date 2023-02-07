Option Explicit

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

Sub getProjectConfigMk()
    If mIp.hasProjectAlps() Then
        pMMIProjectConfigMk = mIp.Infos.ProjectPath & "/config/ProjectConfig.mk"
        If Not isFileExists(pMMIProjectConfigMk) Then pMMIProjectConfigMk = ""

        pDriverProjectConfigMk = mIp.Infos.ProductPath & "/" & mIp.Infos.DriverProject & "/config/ProjectConfig.mk"
        If Not isFileExists(pDriverProjectConfigMk) Then pDriverProjectConfigMk = ""
    Else
        pMMIProjectConfigMk = mIp.Infos.getOverlayPath("device/mediateksample/" & mIp.Infos.Product & "/ProjectConfig.mk")
        If Not isFileExists(pMMIProjectConfigMk) Then pMMIProjectConfigMk = ""

        pDriverProjectConfigMk = mIp.Infos.getDriverOverlayPath("device/mediateksample/" & mIp.Infos.Product & "/ProjectConfig.mk")
        If Not isFileExists(pDriverProjectConfigMk) Then pDriverProjectConfigMk = ""
    End If
    
    pDeviceProjectConfigMk = "device/mediateksample/" & mIp.Infos.Product & "/ProjectConfig.mk"
    If Not isFileExists(pDeviceProjectConfigMk) Then pDeviceProjectConfigMk = ""
End Sub

Sub getProjectInfos()
    If Not mIp.hasProjectInfos() Then
        pMMIProjectConfigMk = ""
        pDriverProjectConfigMk = ""
        pDeviceProjectConfigMk = ""
        Exit Sub
    End If

    'Call getProjectConfigMk()

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

Function getPlatform()
    getPlatform = readTextAndGetValue("CUSTOM_HAL_COMBO", pDeviceProjectConfigMk)
End Function


