Option Explicit

Sub handleCmdInput()
	If HandleFolderPathCmd() Then Exit Sub
	If HandleFilePathCmd() Then Exit Sub
    If handleEditTextCmd() Then Exit Sub
    If handleProjectCmd() Then Exit Sub
	If handleCurrentDictCmd() Then Exit Sub
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

Function runPathFromCmd(cmd, path)
	If mCmdInput.text = cmd Then
	    Call runPath(path)
	    runPathFromCmd = True
	Else
	    runPathFromCmd = False
	End If
End Function

Function HandleFolderPathCmd()
	HandleFolderPathCmd = True
	If runPathFromCmd("m", mIp.Infos.ProjectSdkPath) Then Exit Function
	If runPathFromCmd("d", mIp.Infos.DriverProjectSdkPath) Then Exit Function
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
	If setPathFromCmd("vnd", "device/mediateksample/[product]/vnd_[product].mk") Then Exit Function
	If setPathFromCmd("bc", "device/mediatek/system/common/BoardConfig.mk") Then Exit Function
	If setPathFromCmd("sp", "device/mediatek/system/common/system.prop") Then Exit Function
	If setPathFromCmd("apn", "device/mediatek/config/apns-conf.xml") Then Exit Function
	If setPathFromCmd("cc", "device/mediatek/vendor/common/custom.conf") Then Exit Function
	If setPathFromCmd("fwc", "frameworks/base/core/res/res/values/config.xml") Then Exit Function
	If setPathFromCmd("fws", "frameworks/base/core/res/res/values/strings.xml") Then Exit Function
	If setPathFromCmd("suc", "vendor/mediatek/proprietary/packages/apps/SystemUI/res/values/config.xml") Then Exit Function
	If setPathFromCmd("spdf", "vendor/mediatek/proprietary/packages/apps/SettingsProvider/res/values/defaults.xml") Then Exit Function
	If setPathFromCmd("spdb", "vendor/mediatek/proprietary/packages/apps/SettingsProvider/src/com/android/providers/settings/DatabaseHelper.java") Then Exit Function
	If setPathFromCmd("lgu", "vendor/mediatek/proprietary/bootable/bootloader/lk/dev/logo/[boot_logo]/[boot_logo]_uboot.bmp") Then Exit Function
	If setPathFromCmd("lgk", "vendor/mediatek/proprietary/bootable/bootloader/lk/dev/logo/[boot_logo]/[boot_logo]_kernel.bmp") Then Exit Function
	If setPathFromCmd("ani", "vendor/weibu_sz/media/bootanimation.zip") Then Exit Function
	If setPathFromCmd("label", "vendor/mediatek/proprietary/buildinfo_sys/label.ini") Then Exit Function
	HandleFilePathCmd = False
End Function

Function handleEditTextCmd()
	handleEditTextCmd = True
	If InStr(mCmdInput.text, "bn=") > 0 Then
        Call modBuildNumber(Replace(mCmdInput.text, "bn=", ""))
        Exit Function
    End If
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