Option Explicit

Sub handleCmdInput()
	If HandleFolderPathCmd() Then Call mCmdInput.setText("") : Exit Sub
	If HandleFilePathCmd() Then Call mCmdInput.setText("") : Exit Sub
	If handleProp() Then Exit Sub
    If handleEditTextCmd() Then Call mCmdInput.setText("") : Exit Sub
    If handleProjectCmd() Then Call mCmdInput.setText("") : Exit Sub
	If handleCurrentDictCmd() Then Call mCmdInput.setText("") : Exit Sub
End Sub

Function setPathFromCmd(cmd, path)
	If mCmdInput.text = cmd Then
	    Call setOpenPath(path)
	    Call onOpenPathChange()
	    setPathFromCmd = True
	Else
	    setPathFromCmd = False
	End If
End Function

Function setPathFromCmdAndCopyKey(cmd, key, path)
	If mCmdInput.text = cmd Then
	    Call setOpenPath(path)
	    Call onOpenPathChange()
		Call CopyString(key)
	    setPathFromCmdAndCopyKey = True
	Else
	    setPathFromCmdAndCopyKey = False
	End If
End Function

Function runPathFromCmd(cmd, path)
	If mCmdInput.text = cmd Then
	    Call runPath(path)
	    runPathFromCmd = True
	Else
	    runPathFromCmd = False
	End If
End Function

Function showPropString(cmd, prop)
    If mCmdInput.text = cmd Then
	    Call mCmdInput.setText(prop)
		showPropString = True
	Else
	    showPropString = False
	End If
End Function

Function HandleFolderPathCmd()
	HandleFolderPathCmd = True
	If runPathFromCmd("m", mIp.Infos.ProjectSdkPath) Then Exit Function
	If runPathFromCmd("d", mIp.Infos.DriverProjectSdkPath) Then Exit Function
	If runPathFromCmd("rom", Left(mIp.Infos.Sdk, InStr(mIp.Infos.Sdk, "alps") - 1) & "ROM") Then Exit Function
	If runPathFromCmd("out", mIp.Infos.OutSdkPath) Then Exit Function
	If runPathFromCmd("oa", mIp.Infos.OutSdkPath & "/obj/APPS") Then Exit Function
	If runPathFromCmd("os", mIp.Infos.OutSdkPath & "/system/system_ext/priv-app") Then Exit Function
	If runPathFromCmd("tf", mIp.Infos.OutSdkPath & "/obj/PACKAGING/target_files_intermediates") Then Exit Function
	If setPathFromCmd("st", "vendor/mediatek/proprietary/packages/apps/MtkSettings") Then Exit Function
	If setPathFromCmd("su", "vendor/mediatek/proprietary/packages/apps/SystemUI") Then Exit Function
	If setPathFromCmd("ft", "vendor/mediatek/proprietary/packages/apps/FactoryTest") Then Exit Function
	If setPathFromCmd("fm", "vendor/mediatek/proprietary/packages/apps/FactoryMode") Then Exit Function
	If setPathFromCmd("gms", "vendor/partner_gms") Then Exit Function
	If setPathFromCmd("fwa", "frameworks/base/core/java/android") Then Exit Function
	'If setPathFromCmd("fws", "frameworks/base/services/core/java/com/android/server") Then Exit Function
	If setPathFromCmd("fwv", "frameworks/base/core/res/res/values") Then Exit Function
	If setPathFromCmd("vp", "vendor/mediatek/proprietary/packages/apps") Then Exit Function
	If setPathFromCmd("lot", "vendor/partner_gms/apps/GmsSampleIntegration") Then Exit Function
	If setPathFromCmd("lg", "vendor/mediatek/proprietary/bootable/bootloader/lk/dev/logo/[boot_logo]") Then Exit Function
	If setPathFromCmd("md", "vendor/weibu_sz/media") Then Exit Function
	If setPathFromCmd("tee", "vendor/mediatek/proprietary/trustzone/trustkernel/source/build/[product]") Then Exit Function
	HandleFolderPathCmd = False
End Function

Function HandleFilePathCmd()
	HandleFilePathCmd = True
	If runPathFromCmd("b", mIp.Infos.Sdk & "/build.log") Then Exit Function
	If runPathFromCmd("sb", mIp.Infos.OutSdkPath & "/system/build.prop") Then Exit Function
	If runPathFromCmd("vb", mIp.Infos.OutSdkPath & "/vendor/build.prop") Then Exit Function
	If runPathFromCmd("pb", mIp.Infos.OutSdkPath & "/product/etc/build.prop") Then Exit Function
	If setPathFromCmd("bi", "build/make/tools/buildinfo.sh") Then Exit Function
	If setPathFromCmd("mf", "build/make/core/Makefile") Then Exit Function
	If setPathFromCmd("pc", "device/mediateksample/[product]/ProjectConfig.mk") Then Exit Function
	If setPathFromCmd("sc", "device/mediatek/system/[sys_target_project]/SystemConfig.mk") Then Exit Function
	If setPathFromCmd("full", "device/mediateksample/[product]/full_[product].mk") Then Exit Function
	If setPathFromCmd("sys", "device/mediatek/system/[sys_target_project]/sys_[sys_target_project].mk") Then Exit Function
	If setPathFromCmd("vnd", "device/mediateksample/[product]/vnd_[product].mk") Then Exit Function
	If setPathFromCmd("bc", "device/mediatek/system/common/BoardConfig.mk") Then Exit Function
	If setPathFromCmd("sp", "device/mediatek/system/common/system.prop") Then Exit Function
	If setPathFromCmd("apn", "device/mediatek/config/apns-conf.xml") Then Exit Function
	If setPathFromCmd("cc", "device/mediatek/vendor/common/custom.conf") Then Exit Function
	If setPathFromCmd("fwc", "frameworks/base/core/res/res/values/config.xml") Then Exit Function
	If setPathFromCmd("fws", "frameworks/base/core/res/res/values/strings.xml") Then Exit Function
	If runPathFromCmd("tz", "frameworks/base/packages/SettingsLib/res/xml/timezones.xml") Then Exit Function
	If setPathFromCmd("dc", "[kernel_version]/arch/[target_arch]/configs/[product]_defconfig") Then Exit Function
	If setPathFromCmd("ddc", "[kernel_version]/arch/[target_arch]/configs/[product]_debug_defconfig") Then Exit Function
	If setPathFromCmdAndCopyKey("mtp", "getDeviceProperty", "frameworks/base/media/java/android/mtp/MtpDatabase.java") Then Exit Function
	If setPathFromCmdAndCopyKey("wifiap", "getDefaultApConfiguration", "packages/modules/Wifi/service/java/com/android/server/wifi/WifiApConfigStore.java") Then Exit Function
	If setPathFromCmdAndCopyKey("wifidrt", "getPersistedDeviceName", "packages/modules/Wifi/service/java/com/android/server/wifi/p2p/WifiP2pServiceImpl.java") Then Exit Function
	If setPathFromCmdAndCopyKey("bt", "btif_default_local_name", "system/bt/btif/src/btif_dm.cc") Then Exit Function
	If setPathFromCmdAndCopyKey("st-device", "initializeDeviceName", "vendor/mediatek/proprietary/packages/apps/MtkSettings/src/com/android/settings/deviceinfo/DeviceNamePreferenceController.java") Then Exit Function
	If setPathFromCmd("suc", "vendor/mediatek/proprietary/packages/apps/SystemUI/res/values/config.xml") Then Exit Function
	If setPathFromCmd("spdf", "vendor/mediatek/proprietary/packages/apps/SettingsProvider/res/values/defaults.xml") Then Exit Function
	If setPathFromCmd("spdb", "vendor/mediatek/proprietary/packages/apps/SettingsProvider/src/com/android/providers/settings/DatabaseHelper.java") Then Exit Function
	If setPathFromCmd("brt", "vendor/mediatek/proprietary/packages/overlay/vendor/FrameworkResOverlay/res/values/config.xml") Then Exit Function
	If setPathFromCmd("lgu", "vendor/mediatek/proprietary/bootable/bootloader/lk/dev/logo/[boot_logo]/[boot_logo]_uboot.bmp") Then Exit Function
	If setPathFromCmd("lgk", "vendor/mediatek/proprietary/bootable/bootloader/lk/dev/logo/[boot_logo]/[boot_logo]_kernel.bmp") Then Exit Function
	If setPathFromCmd("ani", "vendor/weibu_sz/media/bootanimation.zip") Then Exit Function
	If setPathFromCmd("label", "vendor/mediatek/proprietary/buildinfo_sys/label.ini") Then Exit Function
	HandleFilePathCmd = False
End Function

Function handleProp()
    handleProp = True
	If showPropString("sample", "persist.sys.sample.device.name") Then Exit Function
	If showPropString("locale", "persist.sys.locale") Then Exit Function
	If showPropString("timezone", "persist.sys.timezone") Then Exit Function
	If showPropString("vol_media", "ro.config.media_vol_default") Then Exit Function
	If showPropString("vol_alarm", "ro.config.alarm_vol_default") Then Exit Function
	If showPropString("vol_system", "ro.config.system_vol_default") Then Exit Function
	If showPropString("vol_notif", "ro.config.notification_vol_default") Then Exit Function
	If showPropString("vol_call", "ro.config.vc_call_vol_default") Then Exit Function
	If showPropString("sku", "ro.boot.hardware.sku") Then Exit Function
	If showPropString("hardware", "ro.boot.hardware.revision") Then Exit Function
	handleProp = False
End Function

Function handleEditTextCmd()
    handleEditTextCmd = True
    If InStr(mCmdInput.text, "bn=") > 0 Then Call modBuildNumber(Replace(mCmdInput.text, "bn=", "")) : Exit Function
    If InStr(mCmdInput.text, "-ota") > 0 Then Call modDisplayIdForOtaTest() : Exit Function
    handleEditTextCmd = False
End Function

Function handleProjectCmd()
	handleProjectCmd = True
	If isNumeric(mCmdInput.text) And Len(mCmdInput.text) < 4 Then
		Dim i, obj : For i = vaWorksInfo.Bound To 0 Step -1
		    Set obj = vaWorksInfo.V(i)
		    If mCmdInput.text = obj.TaskNum Then
		    	Call applyProjectInfo(obj)
		    	Exit Function
		    End If
		Next
	End If
	handleProjectCmd = False
End Function

Function handleCurrentDictCmd()
	handleCurrentDictCmd = True
	If runPathFromCmd("s", pSdkPathText) Then Exit Function
	If runPathFromCmd("p", pPathText) Then Exit Function
	If runPathFromCmd("c", pConfigText) Then Exit Function
	If runPathFromCmd("op", oWs.CurrentDirectory) Then Exit Function
	handleCurrentDictCmd = False
End Function

Sub modBuildNumber(number)
	Dim sysPath, vndPath, sysExist, vndExist, sedStr, R_bnStr, S_bnStr, bnStr, commandStr
	sysPath = "device/mediatek/system/common/BoardConfig.mk"
	vndPath = "device/mediatek/vendor/common/BoardConfig.mk"
	sysExist = False
	vndExist = False
	R_bnStr = "BUILD_NUMBER_WEIBU"
	S_bnStr = "WEIBU_BUILD_NUMBER"

	If oFso.FileExists(mIp.Infos.getOverlaySdkPath(sysPath)) Then sysExist = True
	If oFso.FileExists(mIp.Infos.getOverlaySdkPath(vndPath)) Then vndExist = True
	If Not sysExist And Not vndExist Then MsgBox("Not found BoardConfig.mk overlay") : Exit Sub
    If InStr(mIp.Infos.Sdk, "_r") Then bnStr = R_bnStr
    If InStr(mIp.Infos.Sdk, "_s") Then bnStr = S_bnStr
    If bnStr = "" Then MsgBox("Not found _r OR _s in SDK name") : Exit Sub

	sedStr = "sed -i '/" & bnStr & "/s/[0-9]\+/" & number & "/' "
    If sysExist Then commandStr = sedStr & mIp.Infos.getOverlayPath(sysPath)
    If vndExist Then commandStr = commandStr & " " & mIp.Infos.getOverlayPath(vndPath)
    
    Call CopyString(commandStr)
End Sub

Sub modDisplayIdForOtaTest()
	Dim buildinfo, keyStr, sedStr
	buildinfo = mIp.Infos.ProjectPath & "/config/buildinfo.sh"
	If Not oFso.FileExists(mIp.Infos.Sdk & "/" & buildinfo) Then buildinfo = mIp.Infos.DriverProjectPath & "/config/buildinfo.sh"
	If Not oFso.FileExists(mIp.Infos.Sdk & "/" & buildinfo) Then buildinfo = mIp.Infos.getOverlayPath("build/make/tools/buildinfo.sh")
	If Not oFso.FileExists(mIp.Infos.Sdk & "/" & buildinfo) Then MsgBox("No buildinfo.sh found in overlay") : Exit Sub

	keyStr = "ro.build.display.id"
	If InStr(buildinfo, "/config/") Then
	    sedStr = "sed -i '/" & keyStr & "/s/$/-OTA_test/' " & buildinfo
	    Call CopyString(sedStr)
	Else
	    sedStr = """sed -i '/" & keyStr & "/s/\""&Chr(34)&""$/-OTA_test\""&Chr(34)&""/' " & buildinfo & """"
	    Call CopyQuoteString(sedStr)
	End If
End Sub
