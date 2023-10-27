Option Explicit

Sub handleCmdInput()
	If HandleFolderPathCmd() Then Call mCmdInput.setText("") : Exit Sub
	If HandleFilePathCmd() Then Call mCmdInput.setText("") : Exit Sub
	If handleProp() Then Exit Sub
    If handleGetInfo() Then Call mCmdInput.setText("") : Exit Sub
    If handleEditTextCmd() Then Call mCmdInput.setText("") : Exit Sub
    If handleLinuxCmd() Then Call mCmdInput.setText("") : Exit Sub
    If handleMultiMkdirCmd() Then Call mCmdInput.setText("") : Exit Sub
    If handleOpenPathCmd() Then Call mCmdInput.setText("") : Exit Sub
    If handleCopyCommandCmd() Then Call mCmdInput.setText("") : Exit Sub
    If handleProjectCmd() Then Call mCmdInput.setText("") : Exit Sub
	If handleCurrentDictCmd() Then Call mCmdInput.setText("") : Exit Sub
End Sub

Function HandleFolderPathCmd()
	HandleFolderPathCmd = True
	If mCmdInput.text = "m" Then Call runPath(mIp.Infos.ProjectPath) : Exit Function
	If mCmdInput.text = "d" Then Call runPath(mIp.Infos.DriverProjectPath) : Exit Function
	If mCmdInput.text = "rom" Then Call runPath(getRomPath()) : Exit Function
	If mCmdInput.text = "out" Then Call runPath(mIp.Infos.OutPath) : Exit Function
	If mCmdInput.text = "oa" Then Call runPath(mIp.Infos.OutPath & "/obj/APPS") : Exit Function
	If mCmdInput.text = "os" Then Call runPath(mIp.Infos.OutPath & "/system/system_ext/priv-app") : Exit Function
	If mCmdInput.text = "tf" Then Call runPath(mIp.Infos.OutPath & "/obj/PACKAGING/target_files_intermediates") : Exit Function
	If mCmdInput.text = "lc" Then Call setPathFromCmd("packages/apps/Launcher3") : Exit Function
	If mCmdInput.text = "vlc" Then Call setPathFromCmd("vendor/mediatek/proprietary/packages/apps/Launcher3") : Exit Function
	If mCmdInput.text = "st" Then Call setPathFromCmd("vendor/mediatek/proprietary/packages/apps/MtkSettings") : Exit Function
	If mCmdInput.text = "su" Then Call setPathFromCmd("vendor/mediatek/proprietary/packages/apps/SystemUI") : Exit Function
	If mCmdInput.text = "cam" Then Call setPathFromCmd("vendor/mediatek/proprietary/packages/apps/Camera2") : Exit Function
	If mCmdInput.text = "ft" Then Call setPathFromCmd("vendor/mediatek/proprietary/packages/apps/FactoryTest") : Exit Function
	If mCmdInput.text = "fm" Then Call setPathFromCmd("vendor/mediatek/proprietary/packages/apps/FactoryMode") : Exit Function
	If mCmdInput.text = "gms" Then Call setPathFromCmd("vendor/partner_gms") : Exit Function
	If mCmdInput.text = "fwa" Then Call setPathFromCmd("frameworks/base/core/java/android") : Exit Function
	'If mCmdInput.text = "fws" Then Call setPathFromCmd("frameworks/base/services/core/java/com/android/server") : Exit Function
	If mCmdInput.text = "fwv" Then Call setPathFromCmd("frameworks/base/core/res/res/values") : Exit Function
	If mCmdInput.text = "vp" Then Call setPathFromCmd("vendor/mediatek/proprietary/packages/apps") : Exit Function
	If mCmdInput.text = "lot" Then Call setPathFromCmd("vendor/partner_gms/apps/GmsSampleIntegration") : Exit Function
	If mCmdInput.text = "lg" Then Call setPathFromCmd(getLogoPath()) : Exit Function
	'If mCmdInput.text = "md" Then Call setPathFromCmd("vendor/weibu_sz/media") : Exit Function
	If mCmdInput.text = "tee" Then Call setPathFromCmd("vendor/mediatek/proprietary/trustzone/trustkernel/source/build/[product]") : Exit Function
	HandleFolderPathCmd = False
End Function

Function HandleFilePathCmd()
	HandleFilePathCmd = True
	If mCmdInput.text = "b" Then Call runPath("build.log") : Exit Function
	If mCmdInput.text = "sb" Then Call runPath(mIp.Infos.OutPath & "/system/build.prop") : Exit Function
	If mCmdInput.text = "vb" Then Call runPath(mIp.Infos.OutPath & "/vendor/build.prop") : Exit Function
	If mCmdInput.text = "pb" Then Call runPath(mIp.Infos.OutPath & "/product/etc/build.prop") : Exit Function
	If mCmdInput.text = "bi" Then Call setPathFromCmd("build/make/tools/buildinfo.sh") : Exit Function
	If mCmdInput.text = "mf" Then Call setPathFromCmd("build/make/core/Makefile") : Exit Function
	If mCmdInput.text = "pc" Then Call setPathFromCmd("device/mediateksample/[product]/ProjectConfig.mk") : Exit Function
	If mCmdInput.text = "ci" Then Call setPathFromCmd("weibu/[product]/[project-driver]/config/csci.ini") : Exit Function
	If mCmdInput.text = "dvc" Then Call setPathFromCmd("device/mediatek/system/common/device.mk") : Exit Function
	If mCmdInput.text = "sc" Then Call setPathFromCmd("device/mediatek/system/[sys_target_project]/SystemConfig.mk") : Exit Function
	If mCmdInput.text = "full" Then
		If isT0SdkSys() Then Call setT0SdkVnd()
		Call setPathFromCmd("device/mediateksample/[product]/full_[product].mk") : Exit Function
	End If
	If mCmdInput.text = "sys" Then Call setPathFromCmd("device/mediatek/system/[sys_target_project]/sys_[sys_target_project].mk") : Exit Function
	If mCmdInput.text = "vnd" Then
	    If isT08781() Then
			Call setPathFromCmd("device/mediateksample/[product]/vext_[product].mk") : Exit Function
		Else
			Call setPathFromCmd("device/mediateksample/[product]/vnd_[product].mk") : Exit Function
		End If
	End If
	If mCmdInput.text = "bc" Then Call setPathFromCmd("device/mediatek/system/common/BoardConfig.mk") : Exit Function
	If mCmdInput.text = "sp" Then Call setPathFromCmd("device/mediatek/system/common/system.prop") : Exit Function
	If mCmdInput.text = "apn" Then Call setPathFromCmd("device/mediatek/config/apns-conf.xml") : Exit Function
	'If mCmdInput.text = "cc" Then Call setPathFromCmd("device/mediatek/vendor/common/custom.conf") : Exit Function
	If mCmdInput.text = "fwc" Then Call setPathFromCmd("frameworks/base/core/res/res/values/config.xml") : Exit Function
	If mCmdInput.text = "fws" Then Call setPathFromCmd("frameworks/base/core/res/res/values/strings.xml") : Exit Function
	If mCmdInput.text = "tz" Then Call runPath("frameworks/base/packages/SettingsLib/res/xml/timezones.xml") : Exit Function
	If mCmdInput.text = "tz2" Then Call runPath("system/timezone/output_data/android/tzlookup.xml") : Exit Function
	If mCmdInput.text = "dc" Then Call setPathFromCmd("[kernel_version]/arch/[target_arch]/configs/[product]_defconfig") : Exit Function
	If mCmdInput.text = "ddc" Then Call setPathFromCmd("[kernel_version]/arch/[target_arch]/configs/[product]_debug_defconfig") : Exit Function
	If mCmdInput.text = "mtp" Then Call setPathFromCmdAndCopyKey("getDeviceProperty", "frameworks/base/media/java/android/mtp/MtpDatabase.java") : Exit Function
	If mCmdInput.text = "wfap" Then Call setPathFromCmdAndCopyKey("getDefaultApConfiguration", "packages/modules/Wifi/service/java/com/android/server/wifi/WifiApConfigStore.java") : Exit Function
	If mCmdInput.text = "wfdrt" Then Call setPathFromCmdAndCopyKey("getPersistedDeviceName", "packages/modules/Wifi/service/java/com/android/server/wifi/p2p/WifiP2pServiceImpl.java") : Exit Function
	If mCmdInput.text = "bt" Then Call setPathFromCmdAndCopyKey("btif_default_local_name", "system/bt/btif/src/btif_dm.cc") : Exit Function
	If mCmdInput.text = "devicename" Then Call setPathFromCmdAndCopyKey("initializeDeviceName", "vendor/mediatek/proprietary/packages/apps/MtkSettings/src/com/android/settings/deviceinfo/DeviceNamePreferenceController.java") : Exit Function
	If mCmdInput.text = "bat" Then Call setPathFromCmdAndCopyKey("battery.capacity", "vendor/mediatek/proprietary/packages/overlay/vendor/FrameworkResOverlay/power/res/xml/power_profile.xml") : Exit Function
	If mCmdInput.text = "suc" Then Call setPathFromCmd("vendor/mediatek/proprietary/packages/apps/SystemUI/res/values/config.xml") : Exit Function
	If mCmdInput.text = "spdf" Then Call setPathFromCmd("vendor/mediatek/proprietary/packages/apps/SettingsProvider/res/values/defaults.xml") : Exit Function
	If mCmdInput.text = "spdb" Then Call setPathFromCmd("vendor/mediatek/proprietary/packages/apps/SettingsProvider/src/com/android/providers/settings/DatabaseHelper.java") : Exit Function
	If mCmdInput.text = "brt" Then Call setPathFromCmd("vendor/mediatek/proprietary/packages/overlay/vendor/FrameworkResOverlay/res/values/config.xml") : Exit Function
	If mCmdInput.text = "lgu" Then Call setPathFromCmd("vendor/mediatek/proprietary/bootable/bootloader/lk/dev/logo/[boot_logo]/[boot_logo]_uboot.bmp") : Exit Function
	If mCmdInput.text = "lgk" Then Call setPathFromCmd("vendor/mediatek/proprietary/bootable/bootloader/lk/dev/logo/[boot_logo]/[boot_logo]_kernel.bmp") : Exit Function
	If mCmdInput.text = "wb" Then Call setPathFromCmd("vendor/weibu_sz") : Exit Function
	If mCmdInput.text = "pdt" Then Call setPathFromCmd("vendor/weibu_sz/products/products.mk") : Exit Function
	If mCmdInput.text = "label" Then Call runPath("vendor/mediatek/proprietary/buildinfo_sys/label.ini") : Exit Function
	If mCmdInput.text = "ftn" Then Call runPath("vendor/mediatek/proprietary/packages/apps/FactoryTest/res/xml/factory.xml") : Exit Function
	If mCmdInput.text = "calc" Then Call setPathFromCmd("vendor/mediatek/proprietary/packages/apps/ExactCalculator/src/com/android/calculator2/Calculator.java") : Exit Function
	If mCmdInput.text = "sp1" Then Call setPathFromCmd("build/make/core/version_defaults.mk") : Exit Function
	If mCmdInput.text = "sp2" Then Call setPathFromCmd("vendor/mediatek/proprietary/buildinfo_vnd/device.mk") : Exit Function
	If mCmdInput.text = "bn1" Then Call setPathFromCmd("build/make/core/weibu_config.mk") : Exit Function
	If mCmdInput.text = "bn2" Then Call setPathFromCmd("device/mediatek/system/common/BoardConfig.mk") : Exit Function
	HandleFilePathCmd = False
End Function

Function handleProp()
    handleProp = True
	If mCmdInput.text = "sample" Then Call mCmdInput.setText("persist.sys.sample.device.name") : Exit Function
	If mCmdInput.text = "locale" Then Call mCmdInput.setText("persist.sys.locale") : Exit Function
	If mCmdInput.text = "timezone" Then Call mCmdInput.setText("persist.sys.timezone") : Exit Function
	If mCmdInput.text = "vol_media" Then Call mCmdInput.setText("ro.config.media_vol_default") : Exit Function
	If mCmdInput.text = "vol_alarm" Then Call mCmdInput.setText("ro.config.alarm_vol_default") : Exit Function
	If mCmdInput.text = "vol_system" Then Call mCmdInput.setText("ro.config.system_vol_default") : Exit Function
	If mCmdInput.text = "vol_notif" Then Call mCmdInput.setText("ro.config.notification_vol_default") : Exit Function
	If mCmdInput.text = "vol_call" Then Call mCmdInput.setText("ro.config.vc_call_vol_default") : Exit Function
	If mCmdInput.text = "sku" Then Call mCmdInput.setText("ro.boot.hardware.sku") : Exit Function
	If mCmdInput.text = "hardware" Then Call mCmdInput.setText("ro.boot.hardware.revision") : Exit Function
	If mCmdInput.text = "date" Then Call mCmdInput.setText("$(date +%Y%m%d)") : Exit Function
	handleProp = False
End Function

Function handleGetInfo()
    handleGetInfo = True
	If mCmdInput.text = "getid" Then Call setOpenPath(getOutInfo("ro.build.display.inner.id")) : Exit Function
	If mCmdInput.text = "getfp" Then Call setOpenPath(getOutInfo("ro.system.build.fingerprint")) : Exit Function
	If mCmdInput.text = "getsp" Then Call setOpenPath(getOutInfo("ro.build.version.security_patch")) : Exit Function
	If mCmdInput.text = "getbo" Then Call setOpenPath(getOutInfo("ro.build.version.base_os")) : Exit Function
	If mCmdInput.text = "getgmsv" Then Call setOpenPath(getOutInfo("ro.com.google.gmsversion")) : Exit Function
	If mCmdInput.text = "plf" Then Call setOpenPath(getPlatform()) : Exit Function
	If mCmdInput.text = "gmsv" Then Call setOpenPath(getGmsVersion()) : Exit Function
	If mCmdInput.text = "spc" Then Call setOpenPath(getSecurityPatch()) : Exit Function
    handleGetInfo = False
End Function

Function handleEditTextCmd()
    handleEditTextCmd = True
    If InStr(mCmdInput.text, "bn=") > 0 Then Call modBuildNumber(Replace(mCmdInput.text, "bn=", "")) : Exit Function
    If InStr(mCmdInput.text, "-ota") > 0 Then Call modDisplayIdForOtaTest() : Exit Function
    If InStr(mCmdInput.text, "tz=") > 0 Then Call modSystemprop(Split(mCmdInput.text, "=")) : Exit Function
    If InStr(mCmdInput.text, "loc=") > 0 Then Call modSystemprop(Split(mCmdInput.text, "=")) : Exit Function
    If InStr(mCmdInput.text, "ftd=") > 0 Then Call modSystemprop(Split(mCmdInput.text, "=")) : Exit Function
    If InStr(mCmdInput.text, "=") > 0 Then Call cpFileAndSetValue(Split(mCmdInput.text, "=")) : Exit Function
    handleEditTextCmd = False
End Function

Function handleLinuxCmd()
    handleLinuxCmd = True
    If mCmdInput.text = "cd" Then Call copyCdSdkCommand() : Exit Function
    handleLinuxCmd = False
End Function

Function handleMultiMkdirCmd()
    handleMultiMkdirCmd = True
	If mCmdInput.text = "md-lg" Then Call mkdirLogo() : Exit Function
	If mCmdInput.text = "md-ani" Then Call mkdirBootAnimation() : Exit Function
	If mCmdInput.text = "md-wp" Then Call mkdirWallpaper(False) : Exit Function
	If mCmdInput.text = "md-wp-go" Then Call mkdirWallpaper(True) : Exit Function
	If mCmdInput.text = "md-tee" Then Call mkdirTee() : Exit Function
	If mCmdInput.text = "md-sys" Then Call mkdirProductInfo("sys") : Exit Function
	If mCmdInput.text = "md-vnd" Then Call mkdirProductInfo("vnd") : Exit Function
    handleMultiMkdirCmd = False
End Function

Function handleOpenPathCmd()
    handleOpenPathCmd = True
	If mCmdInput.text = "fjava" Then Call findFrameworksJavaFile() : Exit Function
	If mCmdInput.text = "java" Then Call findJavaFile() : Exit Function
	If mCmdInput.text = "xml" Then Call findXmlFile() : Exit Function
	If mCmdInput.text = "app" Then Call findAppFolder() : Exit Function
	If mCmdInput.text = "cl" Then Call setOpenPath("") : Exit Function
	If mCmdInput.text = "addp" Then Call addProjectPath() : Exit Function
	If mCmdInput.text = "addd" Then Call addDriverProjectPath() : Exit Function
	If mCmdInput.text = "cuts" Then Call cutSdkPath() : Exit Function
	If mCmdInput.text = "cutp" Then Call cutProjectPath() : Exit Function
	If mCmdInput.text = "cp" Then Call compareForProject() : Exit Function
	If mCmdInput.text = "cs" Then Call selectForCompare() : Exit Function
	If mCmdInput.text = "ct" Then Call compareTo() : Exit Function
	If mCmdInput.text = "fmw" Then Call openFirmwareFolder() : Exit Function
	If mCmdInput.text = "req" Then Call openRequirementsFolder() : Exit Function
	If mCmdInput.text = "zt" Then Call openZentao() : Exit Function
    handleOpenPathCmd = False
End Function

Function handleCopyCommandCmd()
    handleCopyCommandCmd = True
	If mCmdInput.text = "lcu" Then Call getLunchCommand("user") : Exit Function
	If mCmdInput.text = "lcd" Then Call getLunchCommand("userdebug") : Exit Function
	If mCmdInput.text = "lce" Then Call getLunchCommand("eng") : Exit Function
	If mCmdInput.text = "lcus" Then Call getT0SysLunchCommand("user") : Exit Function
	If mCmdInput.text = "lcds" Then Call getT0SysLunchCommand("userdebug") : Exit Function
	If mCmdInput.text = "mk" Then Call getMakeCommand(False, False, False) : Exit Function
	If mCmdInput.text = "bmk" Then Call getMakeCommand(False, True, False) : Exit Function
	If mCmdInput.text = "omk" Then Call getMakeCommand(True, False, False) : Exit Function
	If mCmdInput.text = "mko" Then Call getMakeCommand(False, False, True) : Exit Function
	If mCmdInput.text = "bmko" Then Call getMakeCommand(False, True, True) : Exit Function
	If mCmdInput.text = "omko" Then Call getMakeCommand(True, False, True) : Exit Function
	If mCmdInput.text = "md" Then Call MkdirWeibuFolderPath() : Exit Function
	If mCmdInput.text = "cm" Then Call CopyCommitInfo("") : Exit Function
	If InStr(mCmdInput.text, "cm-") = 1 Then Call CopyCommitInfo(Replace(mCmdInput.text, "cm-", "")) : Exit Function
	If mCmdInput.text = "dcm" Then Call CopyDriverCommitInfo() : Exit Function
	If mCmdInput.text = "ota" Then Call CopyBuildOtaUpdate() : Exit Function
	If mCmdInput.text = "cc" Then Call CopyCleanCommand(False) : Exit Function
	If mCmdInput.text = "ccsv" Then Call CopyCleanCommand(True) : Exit Function
	If mCmdInput.text = "outp" Then Call CommandOfOut() : Exit Function
	If mCmdInput.text = "winp" Then Call CopyPathInWindows() : Exit Function
	If mCmdInput.text = "lnxp" Then Call CopyPathInLinux() : Exit Function
	If InStr(mCmdInput.text, "ps-") = 1 Then Call CopyAdbPushCmd(Replace(mCmdInput.text, "ps-", "")) : Exit Function
	If InStr(mCmdInput.text, "cl-") = 1 Then Call CopyAdbClearCmd(Replace(mCmdInput.text, "cl-", "")) : Exit Function
	If InStr(mCmdInput.text, "st-") = 1 Then Call CopyAdbStartCmd(Replace(mCmdInput.text, "st-", "")) : Exit Function
	If InStr(mCmdInput.text, "dp-") = 1 Then Call CopyAdbDumpsysCmd(Replace(mCmdInput.text, "dp-", "")) : Exit Function
	If InStr(mCmdInput.text, "sts-") = 1 Then Call CopyAdbSettingsCmd(Replace(mCmdInput.text, "sts-", "")) : Exit Function
	If mCmdInput.text = "exp" Then Call copyExportToolsPathCmd() : Exit Function
	If mCmdInput.text = "cmd" Then Call startCmdMode() : Exit Function
	If mCmdInput.text = "exit" Then Call exitCmdMode() : Exit Function
	If mCmdInput.text = "ss" Then Call mSaveString.copy() : Call searchStrInVSCode() : Exit Function
	If mCmdInput.text = "muo" Then Call mvOut("user", "out") : Exit Function
	If mCmdInput.text = "mui" Then Call mvOut("user", "in") : Exit Function
	If mCmdInput.text = "mdo" Then Call mvOut("debug", "out") : Exit Function
	If mCmdInput.text = "mdi" Then Call mvOut("debug", "in") : Exit Function
	If InStr(mCmdInput.text, "qm-") = 1 Then Call CopyQmakeCmd(Replace(mCmdInput.text, "qm-", "")) : Exit Function
	If mCmdInput.text = "huq" Then Call sendWeiXinMsg("huq") : Exit Function
	If mCmdInput.text = "zhh" Then Call sendWeiXinMsg("zhh") : Exit Function
	If mCmdInput.text = "luo" Then Call sendWeiXinMsg("luo") : Exit Function
	If mCmdInput.text = "getl" Then Call getCommitMsgList() : Exit Function
	If mCmdInput.text = "ps" Then Call copyStrAndPasteInXshell("git pull -r origin master && git push origin master") : Exit Function
	If mCmdInput.text = "update" Then Call copyStrAndPasteInXshell("git remote update origin --prune") : Exit Function
    handleCopyCommandCmd = False
End Function

Sub sendWeiXinMsg(who)
	Call CopyOpenPathAllText()
    idTimer = window.setTimeout("Call openWeiXin(""" & who & """)", 150, "VBScript")
End Sub

Sub openWeiXin(who)
    Select Case who
	    Case "huq" : who = "huqipeng"
	    Case "zhh" : who = "zhonghongqiang"
	    Case "luo" : who = "luoqingjun"
	End Select

    Call oWs.appactivate("ÆóÒµÎ¢ÐÅ")
    Call oWs.sendkeys("^f")
    Call oWs.sendkeys(who & "xiangmuqun")
	idTimer = window.setTimeout("Call sendEnterKeyInWeiXin(""" & who & """)", 500, "VBScript")
End Sub

Sub sendEnterKeyInWeiXin(who)
    window.clearTimeout(idTimer)
    Call oWs.sendkeys("{ENTER}")
	idTimer = window.setTimeout("Call writeEnterKeyInWeiXin(""" & who & """)", 500, "VBScript")
End Sub

Sub writeEnterKeyInWeiXin(who)
    window.clearTimeout(idTimer)
	Call oWs.sendkeys("@"&who&"{ENTER}@chengtingrong{ENTER}")
    Call oWs.sendkeys("{ENTER}")
	Call oWs.sendkeys("^v")
End Sub

Sub getCommitMsgList()
    Dim arr, i, listStr, count
	listStr = VbLf
	count = 0
	arr = Split(getOpenPath(), VbLf)
	For i = UBound(arr) To 0 Step -1
	    If InStr(arr(i), "] : ") > 0 Then
		    count = count + 1
			listStr = listStr & count & ". " & Right(arr(i), Len(arr(i)) - InStr(arr(i), "] : ") - Len("] : ") + 1) & VbLf
		End If
	Next
	Call setOpenPath(listStr)
End Sub

Function handleProjectCmd()
	handleProjectCmd = True
	If isNumeric(mCmdInput.text) And Len(mCmdInput.text) < 5 Then
		Dim i, obj : For i = vaWorksInfo.Bound To 0 Step -1
		    Set obj = vaWorksInfo.V(i)
		    If mCmdInput.text = obj.TaskNum Then
		    	Call applyShortcutInfos(obj)
		    	Exit Function
		    End If
		Next
		If vbOK = MsgBox("search project?", 1) Then
			Call findProjectWithTaskNum(mCmdInput.text)
		End If
		Exit Function
	ElseIf mCmdInput.text = "z6" Or mCmdInput.text = "x1" Or mCmdInput.text = "x2" Then
	    Call setDrive(mCmdInput.text)
		Exit Function
	ElseIf mCmdInput.text = "s" And isT0SdkVnd() Then
	    Call setT0SdkSys()
		Exit Function
	ElseIf mCmdInput.text = "v" And isT0SdkSys() Then
	    Call setT0SdkVnd()
		Exit Function
	ElseIf mCmdInput.text = "8766s" Or mCmdInput.text = "8168s" Or mCmdInput.text = "8766r" Or mCmdInput.text = "8168r" Then
	    Call setSdk(mCmdInput.text)
		Exit Function
	ElseIf mCmdInput.text = "pp" Then
	    Call applyProjectPath()
		Exit Function
	ElseIf mCmdInput.text = "pl" Then
	    Call updateProductList()
		Exit Function
	End If
	handleProjectCmd = False
End Function

Function handleCurrentDictCmd()
	handleCurrentDictCmd = True
	If mCmdInput.text = "sdk" Then Call runPath(pSdkPathText) : Exit Function
	If mCmdInput.text = "path" Then Call runPath(pPathText) : Exit Function
	If mCmdInput.text = "config" Then Call runPath(pConfigText) : Exit Function
	If mCmdInput.text = "op" Then Call runPath(oWs.CurrentDirectory) : Exit Function
	handleCurrentDictCmd = False
End Function

Sub setPathFromCmd(path)
	Call checkT0Path(path)
	Call setOpenPath(path)
	Call onOpenPathChange()
End Sub

Sub setPathFromCmdAndCopyKey(key, path)
	Call checkT0Path(path)
	Call setOpenPath(path)
	Call onOpenPathChange()
	mSaveString.str = key
End Sub

Sub modBuildNumber(number)
    If InStr(mIp.Infos.Sdk, "8168_s") > 0 Then
	    Call cpFileAndSetValue(Array("bn2", number))
	Else
	    Call cpFileAndSetValue(Array("bn", number))
	End If
End Sub

Function getGmsVersion()
    getGmsVersion = readTextAndGetValue("GMS_PACKAGE_VERSION_ID", "vendor/partner_gms/products/gms_package_version.mk")
End Function

Function getSecurityPatch()
    getSecurityPatch = readTextAndGetValue("VENDOR_SECURITY_PATCH", "vendor/mediatek/proprietary/buildinfo_vnd/device.mk")
End Function

Sub checkT0Path(path)
    If isT0SdkSys() Then
		If InStr(path, "bootable") > 0 Or _
				InStr(path, "vnd") > 0 Or _
				InStr(path, "FrameworkResOverlay") > 0 Or _
				InStr(path, "trustzone") > 0 Then
			Call setT0SdkVnd()
		ElseIf InStr(path, "btif_dm.cc") > 0 Then
			path = "vendor/mediatek/proprietary/packages/modules/Bluetooth/system/btif/src/btif_dm.cc"
		End If
	End If
End Sub

Const TITLE_XSHELL = "Xshell 5 (Free for Home/School)"
Const TITLE_POWERSHELL = "Windows PowerShell"
Const TITLE_VSCODE = "Visual Studio Code"
Sub pasteCmdInXshell()
    idTimer = window.setTimeout("Call appactivateAndPaste(" & """" & TITLE_XSHELL & """)", 500, "VBScript")
End Sub

Sub pasteCmdInPowerShell()
    idTimer = window.setTimeout("Call appactivateAndPaste(" & """" & TITLE_POWERSHELL & """)", 500, "VBScript")
End Sub

Sub appactivateAndPaste(title)
    window.clearTimeout(idTimer)
    Call oWs.appactivate(title)
	Call oWs.sendkeys("+{INSERT}")
End Sub

Sub searchStrInVSCode()
    Call oWs.appactivate(TITLE_VSCODE)
	Call oWs.sendkeys("^f")
	Call oWs.sendkeys("^v")
End Sub
