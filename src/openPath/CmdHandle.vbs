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
	If mCmdInput.text = "rom" Then Call runPath(Left(mIp.Infos.DriveSdk, InStr(mIp.Infos.DriveSdk, "alps") - 1) & "ROM") : Exit Function
	If mCmdInput.text = "out" Then Call runPath(mIp.Infos.OutPath) : Exit Function
	If mCmdInput.text = "oa" Then Call runPath(mIp.Infos.OutPath & "/obj/APPS") : Exit Function
	If mCmdInput.text = "os" Then Call runPath(mIp.Infos.OutPath & "/system/system_ext/priv-app") : Exit Function
	If mCmdInput.text = "tf" Then Call runPath(mIp.Infos.OutPath & "/obj/PACKAGING/target_files_intermediates") : Exit Function
	If mCmdInput.text = "st" Then Call setPathFromCmd("vendor/mediatek/proprietary/packages/apps/MtkSettings") : Exit Function
	If mCmdInput.text = "su" Then Call setPathFromCmd("vendor/mediatek/proprietary/packages/apps/SystemUI") : Exit Function
	If mCmdInput.text = "ft" Then Call setPathFromCmd("vendor/mediatek/proprietary/packages/apps/FactoryTest") : Exit Function
	If mCmdInput.text = "fm" Then Call setPathFromCmd("vendor/mediatek/proprietary/packages/apps/FactoryMode") : Exit Function
	If mCmdInput.text = "gms" Then Call setPathFromCmd("vendor/partner_gms") : Exit Function
	If mCmdInput.text = "fwa" Then Call setPathFromCmd("frameworks/base/core/java/android") : Exit Function
	'If mCmdInput.text = "fws" Then Call setPathFromCmd("frameworks/base/services/core/java/com/android/server") : Exit Function
	If mCmdInput.text = "fwv" Then Call setPathFromCmd("frameworks/base/core/res/res/values") : Exit Function
	If mCmdInput.text = "vp" Then Call setPathFromCmd("vendor/mediatek/proprietary/packages/apps") : Exit Function
	If mCmdInput.text = "lot" Then Call setPathFromCmd("vendor/partner_gms/apps/GmsSampleIntegration") : Exit Function
	If mCmdInput.text = "lg" Then Call setPathFromCmd("vendor/mediatek/proprietary/bootable/bootloader/lk/dev/logo/[boot_logo]") : Exit Function
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
	If mCmdInput.text = "dvc" Then Call setPathFromCmd("device/mediatek/system/common/device.mk") : Exit Function
	If mCmdInput.text = "sc" Then Call setPathFromCmd("device/mediatek/system/[sys_target_project]/SystemConfig.mk") : Exit Function
	If mCmdInput.text = "full" Then Call setPathFromCmd("device/mediateksample/[product]/full_[product].mk") : Exit Function
	If mCmdInput.text = "sys" Then Call setPathFromCmd("device/mediatek/system/[sys_target_project]/sys_[sys_target_project].mk") : Exit Function
	If mCmdInput.text = "vnd" Then Call setPathFromCmd("device/mediateksample/[product]/vnd_[product].mk") : Exit Function
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
	If mCmdInput.text = "st-device" Then Call setPathFromCmdAndCopyKey("initializeDeviceName", "vendor/mediatek/proprietary/packages/apps/MtkSettings/src/com/android/settings/deviceinfo/DeviceNamePreferenceController.java") : Exit Function
	If mCmdInput.text = "bat" Then Call setPathFromCmdAndCopyKey("battery.capacity", "vendor/mediatek/proprietary/packages/overlay/vendor/FrameworkResOverlay/power/res/xml/power_profile.xml") : Exit Function
	If mCmdInput.text = "suc" Then Call setPathFromCmd("vendor/mediatek/proprietary/packages/apps/SystemUI/res/values/config.xml") : Exit Function
	If mCmdInput.text = "spdf" Then Call setPathFromCmd("vendor/mediatek/proprietary/packages/apps/SettingsProvider/res/values/defaults.xml") : Exit Function
	If mCmdInput.text = "spdb" Then Call setPathFromCmd("vendor/mediatek/proprietary/packages/apps/SettingsProvider/src/com/android/providers/settings/DatabaseHelper.java") : Exit Function
	If mCmdInput.text = "brt" Then Call setPathFromCmd("vendor/mediatek/proprietary/packages/overlay/vendor/FrameworkResOverlay/res/values/config.xml") : Exit Function
	If mCmdInput.text = "lgu" Then Call setPathFromCmd("vendor/mediatek/proprietary/bootable/bootloader/lk/dev/logo/[boot_logo]/[boot_logo]_uboot.bmp") : Exit Function
	If mCmdInput.text = "lgk" Then Call setPathFromCmd("vendor/mediatek/proprietary/bootable/bootloader/lk/dev/logo/[boot_logo]/[boot_logo]_kernel.bmp") : Exit Function
	If mCmdInput.text = "ani" Then Call setPathFromCmd("vendor/weibu_sz/media") : Exit Function
	If mCmdInput.text = "pdt" Then Call setPathFromCmd("vendor/weibu_sz/products/products.mk") : Exit Function
	If mCmdInput.text = "label" Then Call runPath("vendor/mediatek/proprietary/buildinfo_sys/label.ini") : Exit Function
	If mCmdInput.text = "ftn" Then Call runPath("vendor/mediatek/proprietary/packages/apps/FactoryTest/res/xml/factory.xml") : Exit Function
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
	handleProp = False
End Function

Function handleGetInfo()
    handleGetInfo = True
	If mCmdInput.text = "getid" Then Call setOpenPath(getOutInfo("ro.build.display.inner.id")) : Exit Function
	If mCmdInput.text = "getfp" Then Call setOpenPath(getOutInfo("ro.system.build.fingerprint")) : Exit Function
	If mCmdInput.text = "getsp" Then Call setOpenPath(getOutInfo("ro.build.version.security_patch")) : Exit Function
	If mCmdInput.text = "getbo" Then Call setOpenPath(getOutInfo("ro.build.version.base_os")) : Exit Function
	If mCmdInput.text = "getplf" Then Call setOpenPath(getPlatform()) : Exit Function
	If mCmdInput.text = "getgmsv" Then Call setOpenPath(getGmsVersion()) : Exit Function
    handleGetInfo = False
End Function

Function handleEditTextCmd()
    handleEditTextCmd = True
    If InStr(mCmdInput.text, "bn=") > 0 Then Call modBuildNumber(Replace(mCmdInput.text, "bn=", "")) : Exit Function
    If InStr(mCmdInput.text, "-ota") > 0 Then Call modDisplayIdForOtaTest() : Exit Function
    If InStr(mCmdInput.text, "tz=") > 0 Then Call modSystemprop(Split(mCmdInput.text, "=")) : Exit Function
    If InStr(mCmdInput.text, "loc=") > 0 Then Call modSystemprop(Split(mCmdInput.text, "=")) : Exit Function
    If InStr(mCmdInput.text, "ftd=") > 0 Then Call modSystemprop(Split(mCmdInput.text, "=")) : Exit Function
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
	If mCmdInput.text = "cc" Then Call CopyCleanCommand() : Exit Function
	If mCmdInput.text = "outp" Then Call CommandOfOut() : Exit Function
	If mCmdInput.text = "winp" Then Call CopyPathInWindows() : Exit Function
	If mCmdInput.text = "lnxp" Then Call CopyPathInLinux() : Exit Function
	If mCmdInput.text = "ps-su" Then Call CopyAdbPushCmd("su") : Exit Function
	If mCmdInput.text = "ps-st" Then Call CopyAdbPushCmd("st") : Exit Function
	If mCmdInput.text = "ps-sl" Then Call CopyAdbPushCmd("sl") : Exit Function
	If mCmdInput.text = "ps-cam" Then Call CopyAdbPushCmd("cam") : Exit Function
	If mCmdInput.text = "exp" Then Call copyExportToolsPathCmd() : Exit Function
	If mCmdInput.text = "cmd" Then Call startCmdMode() : Exit Function
	If mCmdInput.text = "exit" Then Call exitCmdMode() : Exit Function
	If mCmdInput.text = "ss" Then Call mSaveString.copy() : Exit Function
    handleCopyCommandCmd = False
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
	ElseIf mCmdInput.text = "z" Or mCmdInput.text = "z6" Or mCmdInput.text = "x" Then
	    Call setDrive(mCmdInput.text)
		Exit Function
	ElseIf mCmdInput.text = "8766s" Or mCmdInput.text = "8168s" Or mCmdInput.text = "8766r" Or mCmdInput.text = "8168r" Then
	    Call setSdk(mCmdInput.text)
		Exit Function
	ElseIf mCmdInput.text = "pp" Then
	    Call applyProjectPath()
		Exit Function
	End If
	handleProjectCmd = False
End Function

Function handleCurrentDictCmd()
	handleCurrentDictCmd = True
	If mCmdInput.text = "s" Then Call runPath(pSdkPathText) : Exit Function
	If mCmdInput.text = "p" Then Call runPath(pPathText) : Exit Function
	If mCmdInput.text = "c" Then Call runPath(pConfigText) : Exit Function
	If mCmdInput.text = "op" Then Call runPath(oWs.CurrentDirectory) : Exit Function
	handleCurrentDictCmd = False
End Function

Sub setPathFromCmd(path)
	Call setOpenPath(path)
	Call onOpenPathChange()
End Sub

Sub setPathFromCmdAndCopyKey(key, path)
	Call setOpenPath(path)
	Call onOpenPathChange()
	mSaveString.str = key
End Sub

Function getMultiMkdirStr(arr, what)
    Dim str, path, ovlFolder, ovlFile
    For Each path In arr
	    ovlFolder = mIp.Infos.getOverlayPath(getFolderPath(path))
	    ovlFile = mIp.Infos.getOverlayPath(path)
	    If (Not (what = "lg" And InStr(path, "_kernel.bmp") > 0)) And (Not isFolderExists(ovlFolder)) Then
	        str =  str & "mkdir -p " & ovlFolder & ";"
		End If

	    If what = "lg" Then
	        str =  str & "cp ../File/logo.bmp " & ovlFile & ";"
		ElseIf what = "ani" And InStr(path, "bootanimation.zip") > 0 Then
	        str =  str & "cp ../File/bootanimation.zip " & ovlFile & ";"
		ElseIf what = "wp" Then
		    If InStr(path, "default_wallpaper.png") > 0 Then
	            str =  str & "cp ../File/default_wallpaper.png " & ovlFile & ";"
			ElseIf InStr(path, "default_wallpaper.jpg") > 0 Then
	            str =  str & "cp ../File/default_wallpaper.jpg " & ovlFile & ";"
			End If
		Else
		    str =  str & "cp " & path & " " & ovlFile & ";"
		End If
	Next
	getMultiMkdirStr = str
End Function

Sub modBuildNumber(number)
	Dim sysPath, vndPath, sysExist, vndExist, sedStr, R_bnStr, S_bnStr, bnStr, commandStr
	sysPath = "device/mediatek/system/common/BoardConfig.mk"
	vndPath = "device/mediatek/vendor/common/BoardConfig.mk"
	sysExist = False
	vndExist = False
	R_bnStr = "BUILD_NUMBER_WEIBU"
	S_bnStr = "WEIBU_BUILD_NUMBER"

	If isFileExists(mIp.Infos.getOverlayPath(sysPath)) Then sysExist = True
	If isFileExists(mIp.Infos.getOverlayPath(vndPath)) Then vndExist = True
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
	If Not isFileExists(buildinfo) Then buildinfo = mIp.Infos.DriverProjectPath & "/config/buildinfo.sh"
	If Not isFileExists(buildinfo) Then buildinfo = mIp.Infos.getOverlayPath("build/make/tools/buildinfo.sh")
	If Not isFileExists(buildinfo) Then MsgBox("No buildinfo.sh found in overlay") : Exit Sub

	keyStr = "ro.build.display.id"
	If InStr(buildinfo, "/config/") Then
	    sedStr = "sed -i '/" & keyStr & "/s/$/-OTA_test/' " & buildinfo
	    Call CopyString(sedStr)
	Else
	    sedStr = """sed -i '/" & keyStr & "/s/\""&Chr(34)&""$/-OTA_test\""&Chr(34)&""/' " & buildinfo & """"
	    Call CopyQuoteString(sedStr)
	End If
End Sub

Sub modSystemprop(whatArr)
    Dim systempropPath, cmdStr, keyStr, valueStr
	systempropPath = mIp.Infos.ProjectPath & "/config/system.prop"

    valueStr = whatArr(1)
	If whatArr(0) = "tz" Then
	    keyStr = "persist.sys.timezone"
		valueStr = Replace(valueStr, "/", "\/")
	ElseIf whatArr(0) = "loc" Then
	    keyStr = "persist.sys.locale"
	ElseIf whatArr(0) = "ftd" Then
	    keyStr = "ro.weibu.factorytest.disable_" & whatArr(1)
		valueStr = "1"
	Else
	    Exit Sub
	End If

    If Not isFileExists(systempropPath) Then
	    cmdStr = """echo -e ""&Chr(34)&""" & keyStr & "=" & valueStr & """&Chr(34)&""\n > " & systempropPath
		cmdStr = cmdStr & ";git diff " & systempropPath
	    Call CopyQuoteString(cmdStr)
	Else
	    If strExistInFile(systempropPath, keyStr) Then
		    cmdStr = "sed -i '/" & keyStr & "/s/.*/" & keyStr & "=" & valueStr & "/' " & systempropPath
		Else
	        cmdStr = "sed -i '$a " & keyStr & "=" & valueStr & "' " & systempropPath
		End If
		cmdStr = cmdStr & ";git diff " & systempropPath
	    Call CopyString(cmdStr)
	End If
End Sub

Sub mkdirLogo()
    Dim lg_fd, lg_u, lg_k
    lg_fd = "vendor/mediatek/proprietary/bootable/bootloader/lk/dev/logo/[boot_logo]/"
    lg_u = Replace(lg_fd & "[boot_logo]_uboot.bmp", "[boot_logo]", mIp.Infos.BootLogo)
    lg_k = Replace(lg_fd & "[boot_logo]_kernel.bmp", "[boot_logo]", mIp.Infos.BootLogo)

    Dim arr, finalStr
    arr = Array(lg_u, lg_k)
    finalStr = getMultiMkdirStr(arr, "lg")
    Call CopyString(finalStr)
End Sub

Sub mkdirBootAnimation()
    Dim ani_media, ani_product
    ani_media = "vendor/weibu_sz/media/bootanimation.zip"
    ani_product = "vendor/weibu_sz/products/products.mk"

    Dim arr, finalStr
    arr = Array(ani_media, ani_product)
    finalStr = getMultiMkdirStr(arr, "ani")
    Call CopyString(finalStr)
End Sub

Sub mkdirWallpaper(go)
    Dim wp_gms, wp_go1, wp_go2, wp1, wp2, wp3
    wp_gms = "vendor/partner_gms/overlay/AndroidSGmsBetaOverlay/res/drawable-nodpi/default_wallpaper.png"
    wp_go1 = "device/mediatek/common/overlay/ago/frameworks/base/core/res/res/drawable-nodpi/default_wallpaper.jpg"
    wp_go2 = "device/mediatek/system/common/overlay/ago/frameworks/base/core/res/res/drawable-nodpi/default_wallpaper.jpg"
    wp1 = "frameworks/base/core/res/res/drawable-nodpi/default_wallpaper.png"
    wp2 = "frameworks/base/core/res/res/drawable-sw600dp-nodpi/default_wallpaper.png"
    wp3 = "frameworks/base/core/res/res/drawable-sw720dp-nodpi/default_wallpaper.png"

    Dim arr, finalStr
    If Not go Then
        arr = Array(wp_gms, wp1, wp2, wp3)
    Else
        arr = Array(wp_go1, wp_go2, wp1, wp2, wp3)
    End If
    finalStr = getMultiMkdirStr(arr, "wp")
    Call CopyString(finalStr)
End Sub

Sub CopyAdbPushCmd(which)
    Dim outPath, sourcePath, targetPath, finalStr
	outPath = mIp.Infos.getPathWithDriveSdk(mIp.Infos.OutPath)
    If which = "su" Then
	    sourcePath = outPath & "\system\system_ext\priv-app\MtkSystemUI"
		targetPath = "/system/system_ext/priv-app/"
		finalStr = "adb push " & sourcePath & " " & targetPath
	ElseIf which = "st" Then
	    sourcePath = outPath & "\system\system_ext\priv-app\MtkSettings"
		targetPath = "/system/system_ext/priv-app/"
		finalStr = "adb push " & sourcePath & " " & targetPath
	ElseIf which = "sl" Then
	    sourcePath = outPath & "\system\system_ext\priv-app\SearchLauncherQuickStep"
		targetPath = "/system/system_ext/priv-app/"
		finalStr = "adb push " & sourcePath & " " & targetPath
	ElseIf which = "cam" Then
	    sourcePath = outPath & "\system\system_ext\app\Camera"
		targetPath = "/system/system_ext/app/"
		finalStr = "adb push " & sourcePath & " " & targetPath
	End If
	Call CopyString(finalStr)
End Sub

Function getGmsVersion()
    getGmsVersion = readTextAndGetValue("GMS_PACKAGE_VERSION_ID", "vendor/partner_gms/products/gms_package_version.mk")
End Function
