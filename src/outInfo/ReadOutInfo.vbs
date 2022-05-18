

Const ID_INPUT_OUT_INFOS = "input_out_infos"
Const ID_BUTTON_SHOW_OUT_INFOS = "button_show_out_infos"
Const ID_BUTTON_CLOSE_OUT_INFOS = "button_close_out_infos"

Dim pSystemBuildProp
Dim pVendorBuildProp
Dim pProductBuildProp

Sub getOutInfos()
    If Not mIp.hasProjectInfos() Then
        pSystemBuildProp = ""
        pVendorBuildProp = ""
        pProductBuildProp = ""
        Exit Sub
    End If
    
    Dim outProductPath : outProductPath = getOutProductPath()

    pSystemBuildProp = outProductPath & "/system/build.prop"
    pVendorBuildProp = outProductPath & "/vendor/build.prop"
    pProductBuildProp = outProductPath & "/product/build.prop"
    If Not oFso.FileExists(pProductBuildProp) Then
        pProductBuildProp = outProductPath & "/product/etc/build.prop"
    End If

    Dim str
    str = str & "display_id=" & readTextAndGetValue("ro.build.display.id", pSystemBuildProp) & VbLf
    str = str & "fingerprint=" & readTextAndGetValue("ro.system.build.fingerprint", pSystemBuildProp) & VbLf
    str = str & "incremental=" & readTextAndGetValue("ro.build.version.incremental", pSystemBuildProp) & VbLf
    str = str & "build_type=" & readTextAndGetValue("ro.build.type", pSystemBuildProp) & VbLf
    str = str & "build_date=" & readTextAndGetValue("ro.build.date", pSystemBuildProp) & VbLf
    str = str & "build_date_utc=" & readTextAndGetValue("ro.build.date.utc", pSystemBuildProp) & VbLf
    str = str & "locale=" & readTextAndGetValue("ro.product.locale", pSystemBuildProp) & VbLf
    str = str & "brand=" & readTextAndGetValue("ro.product.system.brand", pSystemBuildProp) & VbLf
    str = str & "model=" & readTextAndGetValue("ro.product.system.model", pSystemBuildProp) & VbLf
    str = str & "device=" & readTextAndGetValue("ro.product.system.device", pSystemBuildProp) & VbLf
    str = str & "product=" & readTextAndGetValue("ro.product.system.name", pSystemBuildProp) & VbLf
    str = str & "manufacturer=" & readTextAndGetValue("ro.product.system.manufacturer", pSystemBuildProp) & VbLf
    str = str & "platform=" & readTextAndGetValue("ro.board.platform", pVendorBuildProp) & VbLf
    str = str & "base_os=" & readTextAndGetValue("ro.build.version.base_os", pSystemBuildProp) & VbLf
    str = str & "gmsversion=" & readTextAndGetValue("ro.com.google.gmsversion", pProductBuildProp) & VbLf
    str = str & "security_path=" & readTextAndGetValue("ro.build.version.security_patch", pSystemBuildProp) & VbLf
    str = str & "client_id=" & readTextAndGetValue("ro.com.google.clientidbase", pProductBuildProp) & VbLf
    str = str & "density=" & readTextAndGetValue("ro.sf.lcd_density", pVendorBuildProp) & VbLf
    str = str & "ringtone=" & readTextAndGetValue("ro.config.ringtone", pVendorBuildProp) & VbLf
    str = str & "notification_sound=" & readTextAndGetValue("ro.config.notification_sound", pVendorBuildProp) & VbLf
    str = str & "alarm_alert=" & readTextAndGetValue("ro.config.alarm_alert", pVendorBuildProp)

    Call hideElement(ID_BUTTON_SHOW_OUT_INFOS)
    Call showElement(ID_BUTTON_CLOSE_OUT_INFOS)

    Call showElement(ID_INPUT_OUT_INFOS)
    Call setElementValue(ID_INPUT_OUT_INFOS, str)
End Sub

Sub closeOutInfos()
    Call hideElement(ID_INPUT_OUT_INFOS)
    Call hideElement(ID_BUTTON_CLOSE_OUT_INFOS)
    Call showElement(ID_BUTTON_SHOW_OUT_INFOS)
End Sub


