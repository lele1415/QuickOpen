Option Explicit

Const TITLE_XSHELL = "Xshell 5 (Free for Home/School)"
Const TITLE_POWERSHELL = "Windows PowerShell"
Const TITLE_VSCODE = "Visual Studio Code"
Const TITLE_SUBLIME = "Sublime Text (UNREGISTERED)"

Dim mSaveString : Set mSaveString = New SaveString
Dim mLeftComparePath, mRightComparePath
Dim idTimer

Sub copyStrAndPasteInXshell(cmdStr)
    If cmdStr = "" Then MsgBox("Empty!") : Exit Sub
    Call CopyString(cmdStr)
    idTimer = window.setTimeout("Call appactivateAndPaste(" & """" & TITLE_XSHELL & """)", 500, "VBScript")
End Sub

Sub copyStrAndPasteInPowerShell(cmdStr)
    If cmdStr = "" Then MsgBox("Empty!") : Exit Sub
    Call CopyString(cmdStr)
    idTimer = window.setTimeout("Call appactivateAndPaste(" & """" & TITLE_POWERSHELL & """)", 500, "VBScript")
End Sub

Sub setPathFromCmd(path)
    Call setOpenPath(path)
    Call onOpenPathChange()
End Sub

Sub setPathFromCmdAndCopyKey(key, path)
    Call setOpenPath(path)
    Call onOpenPathChange()
    mSaveString.str = key
End Sub

Sub appactivateAndPaste(title)
    window.clearTimeout(idTimer)
    Call oWs.appactivate(title)
    Call oWs.sendkeys("+{INSERT}")
End Sub

Sub copyStrAndPasteInCodeEditor()
    mSaveString.copy()
    idTimer = window.setTimeout("Call appactivateCodeEditor(" & """" & TITLE_SUBLIME & """)", 500, "VBScript")
End Sub

Sub appactivateCodeEditor(title)
    window.clearTimeout(idTimer)
    Call oWs.appactivate(title)
    idTimer = window.setTimeout("Call searchInCodeEditor()", 300, "VBScript")
End Sub

Sub searchInCodeEditor()
    window.clearTimeout(idTimer)
    Call oWs.sendkeys("^f")
    idTimer = window.setTimeout("Call pasteCodeEditor()", 300, "VBScript")
End Sub

Sub pasteCodeEditor()
    window.clearTimeout(idTimer)
    Call oWs.sendkeys("^v")
End Sub

Sub handleCmdInput()
    If HandleFolderPathCmd() Then Call setCmdText("") : Exit Sub
    If HandleFilePathCmd() Then Call setCmdText("") : Exit Sub
    If handleProp() Then Call setCmdText("") : Exit Sub
    If handleGetInfo() Then Call setCmdText("") : Exit Sub
    If handleLinuxCmd() Then Call setCmdText("") : Exit Sub
    If handleMultiMkdirCmd() Then Call setCmdText("") : Exit Sub
    If handleOpenPathCmd() Then Call setCmdText("") : Exit Sub
    If handleCopyCommandCmd() Then Call setCmdText("") : Exit Sub
    If handleEditTextCmd() Then Call setCmdText("") : Exit Sub
    If handleProjectCmd() Then Call setCmdText("") : Exit Sub
    If handleCurrentDictCmd() Then Call setCmdText("") : Exit Sub
End Sub

Function HandleFolderPathCmd()
    HandleFolderPathCmd = True
    If mCmdInput.text = "pp" Then Call runPath(mBuild.Infos.ProjectPath) : Exit Function
    If mCmdInput.text = "rom" Then Call runPath("../ROM") : Exit Function
    If mCmdInput.text = "out" Then Call runPath(mBuild.Infos.OutPath) : Exit Function
    If mCmdInput.text = "oa" Then Call runPath(mBuild.Infos.OutPath & "/obj/APPS") : Exit Function
    If mCmdInput.text = "os" Then Call runPath(mBuild.Infos.OutSystemExtPath) : Exit Function
    If mCmdInput.text = "tf" Then Call runPath(mBuild.Infos.OutPath & "/obj/PACKAGING/target_files_intermediates") : Exit Function
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
    If mCmdInput.text = "lg" Then Call setVndBuild() : Call setPathFromCmd(mTask.Vnd.Infos.LogoPath) : Exit Function
    'If mCmdInput.text = "md" Then Call setPathFromCmd("vendor/weibu_sz/media") : Exit Function
    If mCmdInput.text = "tee" Then Call setVndBuild() : Call setPathFromCmd("vendor/mediatek/proprietary/trustzone/trustkernel/source/build/" & mBuild.Product) : Exit Function
    HandleFolderPathCmd = False
End Function

Function HandleFilePathCmd()
    HandleFilePathCmd = True
    If mCmdInput.text = "b" Then Call runPath("build.log") : Exit Function
    If mCmdInput.text = "sbl" Then Call setSysBuild() : Call runPath("sys_build.log") : Exit Function
    If mCmdInput.text = "vbl" Then Call runPath("[vnd]_build.log") : Exit Function
    If mCmdInput.text = "kbl" Then Call runPath("krn_build.log") : Exit Function
    If mCmdInput.text = "hbl" Then Call runPath("hal_build.log") : Exit Function
    If mCmdInput.text = "sb" Then Call runPath(mBuild.Infos.OutPath & "/system/build.prop") : Exit Function
    If mCmdInput.text = "vb" Then Call runPath(mBuild.Infos.OutPath & "/vendor/build.prop") : Exit Function
    If mCmdInput.text = "pb" Then Call runPath(mBuild.Infos.OutPath & "/product/etc/build.prop") : Exit Function
    If mCmdInput.text = "gtxt" Then Call runPath(mBuild.Infos.OutPath & "/system/data/misc/git.txt") : Exit Function
    If mCmdInput.text = "dtxt" Then Call runPath(mBuild.Infos.OutPath & "/system/data/misc/diff.txt") : Exit Function
    If mCmdInput.text = "bi" Then Call setPathFromCmd("build/make/tools/buildinfo.sh") : Exit Function
    If mCmdInput.text = "mf" Then Call setPathFromCmd("build/make/core/Makefile") : Exit Function
    If mCmdInput.text = "pc" Then Call setVndBuild() : Call setPathFromCmd("device/mediateksample/[product]/ProjectConfig.mk") : Exit Function
    If mCmdInput.text = "ci" Then Call setVndBuild() : Call setPathFromCmd("weibu/[product]/[project]/config/csci.ini") : Exit Function
    If mCmdInput.text = "sdc" Then Call setPathFromCmd("device/mediatek/system/common/device.mk") : Exit Function
    If mCmdInput.text = "vdc" Then Call setPathFromCmd("device/mediatek/system/common/device.mk") : Exit Function
    If mCmdInput.text = "cdc" Then Call setPathFromCmd("device/mediatek/common/device.mk") : Exit Function
    If mCmdInput.text = "sc" Then Call setSysBuild() : Call setPathFromCmd("device/mediatek/system/[product]/SystemConfig.mk") : Exit Function
    If mCmdInput.text = "full" Then Call setVndBuild() : Call setPathFromCmd("device/mediateksample/[product]/full_[product].mk") : Exit Function
    If mCmdInput.text = "sys" Then Call setSysBuild() : Call setPathFromCmd("device/mediatek/system/[product]/sys_[product].mk") : Exit Function
    If mCmdInput.text = "vnd" Then Call setVndBuild() : Call setPathFromCmd("device/mediateksample/[product]/[vnd]_[product].mk") : Exit Function
    If mCmdInput.text = "bc" Then Call setPathFromCmd("device/mediatek/system/common/BoardConfig.mk") : Exit Function
    If mCmdInput.text = "sp" Then Call setPathFromCmd("device/mediatek/system/common/system.prop") : Exit Function
    If mCmdInput.text = "apn" Then Call setPathFromCmd("device/mediatek/config/apns-conf.xml") : Exit Function
    'If mCmdInput.text = "cc" Then Call setPathFromCmd("device/mediatek/vendor/common/custom.conf") : Exit Function
    If mCmdInput.text = "fwc" Then Call setPathFromCmd("frameworks/base/core/res/res/values/config.xml") : Exit Function
    If mCmdInput.text = "fws" Then Call setPathFromCmd("frameworks/base/core/res/res/values/strings.xml") : Exit Function
    If mCmdInput.text = "tv" Then Call setPathFromCmd("frameworks/base/core/java/android/widget/TextView.java") : Exit Function
    If mCmdInput.text = "tz" Then Call runPath("frameworks/base/packages/SettingsLib/res/xml/timezones.xml") : Exit Function
    If mCmdInput.text = "tz2" Then Call runPath("system/timezone/output_data/android/tzlookup.xml") : Exit Function
    If mCmdInput.text = "mtp" Then Call setPathFromCmdAndCopyKey("getDeviceProperty", "frameworks/base/media/java/android/mtp/MtpDatabase.java") : Exit Function
    If mCmdInput.text = "wfap" Then Call setPathFromCmdAndCopyKey("getDefaultApConfiguration", "packages/modules/Wifi/service/java/com/android/server/wifi/WifiApConfigStore.java") : Exit Function
    If mCmdInput.text = "wfdrt" Then Call setPathFromCmdAndCopyKey("getPersistedDeviceName", "packages/modules/Wifi/service/java/com/android/server/wifi/p2p/WifiP2pServiceImpl.java") : Exit Function
    If mCmdInput.text = "bt" Then Call setPathFromCmdAndCopyKey("btif_default_local_name", mBuild.Infos.getBluetoothfilePath()) : Exit Function
    If mCmdInput.text = "bat" Then
        If mBuild.Infos.Version < 13 Then
            Call setVndBuild()
        Else
            Call setSysBuild()
        End If
        Call setPathFromCmdAndCopyKey("battery.capacity", mBuild.Infos.getPowerProfilePath()) : Exit Function
    End If
    If mCmdInput.text = "-c" Then Call setPathFromCmd(getOpenPath() & "/res/values/config.xml") : Exit Function
    If mCmdInput.text = "-s" Then Call setPathFromCmd(getOpenPath() & "/res/values/strings.xml") : Exit Function
    If mCmdInput.text = "-zs" Then Call setPathFromCmd(getOpenPath() & "/res/values-zh-rCN/strings.xml") : Exit Function
    If mCmdInput.text = "-js" Then Call setPathFromCmd(getOpenPath() & "/res/values-ja/strings.xml") : Exit Function
    If mCmdInput.text = "-rs" Then Call setPathFromCmd(getOpenPath() & "/res/values-ru/strings.xml") : Exit Function
    If mCmdInput.text = "spdf" Then Call setPathFromCmd("vendor/mediatek/proprietary/packages/apps/SettingsProvider/res/values/defaults.xml") : Exit Function
    If mCmdInput.text = "spdb" Then Call setPathFromCmd("vendor/mediatek/proprietary/packages/apps/SettingsProvider/src/com/android/providers/settings/DatabaseHelper.java") : Exit Function
    If mCmdInput.text = "lot" Then Call setPathFromCmd("vendor/partner_gms/apps/GmsSampleIntegration/res_dhs_full/xml/partner_default_layout.xml") : Exit Function
    If mCmdInput.text = "brt" Then Call setVndBuild() : Call setPathFromCmd("vendor/mediatek/proprietary/packages/overlay/vendor/FrameworkResOverlay/res/values/config.xml") : Exit Function
    If mCmdInput.text = "pdt" Then Call setPathFromCmd("vendor/weibu_sz/products/products.mk") : Exit Function
    If mCmdInput.text = "label" Then Call runPath("vendor/mediatek/proprietary/buildinfo_sys/label.ini") : Exit Function
    If mCmdInput.text = "ftx" Then Call runPath("vendor/mediatek/proprietary/packages/apps/FactoryTest/res/xml/factory.xml") : Exit Function
    If mCmdInput.text = "calc" Then Call setPathFromCmd("vendor/mediatek/proprietary/packages/apps/ExactCalculator/src/com/android/calculator2/Calculator.java") : Exit Function
    If mCmdInput.text = "sp1" Then Call setPathFromCmd("build/make/core/version_defaults.mk") : Exit Function
    If mCmdInput.text = "sp2" Then Call setPathFromCmd("vendor/mediatek/proprietary/buildinfo_vnd/device.mk") : Exit Function
    If mCmdInput.text = "bn1" Then Call setPathFromCmd("build/make/core/weibu_config.mk") : Exit Function
    If mCmdInput.text = "bn2" Then Call setPathFromCmd("device/mediatek/system/common/BoardConfig.mk") : Exit Function
    HandleFilePathCmd = False
End Function

Function handleProp()
    handleProp = True
    If mCmdInput.text = "sample" Then Call setOpenPath("persist.sys.sample.device.name") : Exit Function
    If mCmdInput.text = "locale" Then Call setOpenPath("ro.product.locale" & VbLf & "persist.sys.locale") : Exit Function
    If mCmdInput.text = "timezone" Then Call setOpenPath("persist.sys.timezone") : Exit Function
    If mCmdInput.text = "vol" Then Call setOpenPath("ro.config.media_vol_default=15" & VbLf &_
                                                                "ro.config.alarm_vol_default=15" & VbLf &_
                                                                "ro.config.ring_vol_default=15" & VbLf &_
                                                                "ro.config.system_vol_default=15" & VbLf &_
                                                                "ro.config.notification_vol_default=15" & VbLf &_
                                                                "ro.config.vc_call_vol_default=7") : Exit Function
    If mCmdInput.text = "sku" Then Call setOpenPath("ro.boot.hardware.sku") : Exit Function
    If mCmdInput.text = "hardware" Then Call setOpenPath("ro.boot.hardware.revision") : Exit Function
    If mCmdInput.text = "date" Then Call setOpenPath("$(date +%Y%m%d)") : Exit Function
    handleProp = False
End Function

Function handleGetInfo()
    handleGetInfo = True
    If mCmdInput.text = "getdi" Then Call setOpenPath(mBuild.Infos.getOutProp("ro.build.display.id")) : Exit Function
    If mCmdInput.text = "getid" Then Call setOpenPath(mBuild.Infos.getOutProp("ro.build.display.inner.id")) : Exit Function
    If mCmdInput.text = "getfp" Then Call setOpenPath(mBuild.Infos.getOutProp("ro.system.build.fingerprint")) : Exit Function
    If mCmdInput.text = "getsp" Then Call setOpenPath(mBuild.Infos.getOutProp("ro.build.version.security_patch")) : Exit Function
    If mCmdInput.text = "getbo" Then Call setOpenPath(mBuild.Infos.getOutProp("ro.build.version.base_os")) : Exit Function
    If mCmdInput.text = "getgmsv" Then Call setOpenPath(mBuild.Infos.getOutProp("ro.com.google.gmsversion")) : Exit Function
    If mCmdInput.text = "plf" Then Call setOpenPath(mBuild.Infos.getPlatform()) : Exit Function
    If mCmdInput.text = "gmsv" Then Call setOpenPath(readTextAndGetValue("GMS_PACKAGE_VERSION_ID", "vendor/partner_gms/products/gms_package_version.mk")) : Exit Function
    If mCmdInput.text = "spc" Then Call setOpenPath(readTextAndGetValue("PLATFORM_SECURITY_PATCH", "build/make/core/version_defaults.mk")) : Exit Function
    If mCmdInput.text = "pbd" Then Call setOpenPath(readTextAndGetValue("PRODUCT_BRAND", mBuild.Infos.getOverlayPath(mBuild.Infos.ProductFile))) : Exit Function
    If mCmdInput.text = "pmd" Then Call setOpenPath(readTextAndGetValue("PRODUCT_MODEL", mBuild.Infos.getOverlayPath(mBuild.Infos.ProductFile))) : Exit Function
    If mCmdInput.text = "pdc" Then Call setOpenPath(readTextAndGetValue("PRODUCT_DEVICE", mBuild.Infos.getOverlayPath(mBuild.Infos.ProductFile))) : Exit Function
    If mCmdInput.text = "pnm" Then Call setOpenPath(readTextAndGetValue("PRODUCT_NAME", mBuild.Infos.getOverlayPath(mBuild.Infos.ProductFile))) : Exit Function
    If mCmdInput.text = "pmf" Then Call setOpenPath(readTextAndGetValue("PRODUCT_MANUFACTURER", mBuild.Infos.getOverlayPath(mBuild.Infos.ProductFile))) : Exit Function
    handleGetInfo = False
End Function

Function handleLinuxCmd()
    handleLinuxCmd = True
    If mCmdInput.text = "cd" Then Call copyCdSdkCommand() : Exit Function
    If mCmdInput.text = "lcu" Then Call getLunchCommand("user", mTask.Infos.TaskNum) : Exit Function
    If mCmdInput.text = "lcd" Then Call getLunchCommand("userdebug", mTask.Infos.TaskNum) : Exit Function
    If mCmdInput.text = "lce" Then Call getLunchCommand("eng", mTask.Infos.TaskNum) : Exit Function
    If InStr(mCmdInput.text, "lcu-") = 1 Then Call getLunchCommand("user", Replace(mCmdInput.text, "lcu-", "")) : Exit Function
    If InStr(mCmdInput.text, "lcd-") = 1 Then Call getLunchCommand("userdebug", Replace(mCmdInput.text, "lcd-", "")) : Exit Function
    If InStr(mCmdInput.text, "lce-") = 1 Then Call getLunchCommand("eng", Replace(mCmdInput.text, "lce-", "")) : Exit Function
    If mCmdInput.text = "lcus" Then Call getSysLunchCommand("user") : Exit Function
    If mCmdInput.text = "lcds" Then Call getSysLunchCommand("userdebug") : Exit Function
    If mCmdInput.text = "lcuv" Then Call getVndLunchCommand("user") : Exit Function
    If mCmdInput.text = "lcdv" Then Call getVndLunchCommand("userdebug") : Exit Function
    If mCmdInput.text = "mk" Then Call getMakeCommand(False, False, False) : Exit Function
    If mCmdInput.text = "bmk" Then Call getMakeCommand(False, True, False) : Exit Function
    If mCmdInput.text = "omk" Then Call getMakeCommand(True, False, False) : Exit Function
    If mCmdInput.text = "mko" Then Call getMakeCommand(False, False, True) : Exit Function
    If mCmdInput.text = "bmko" Then Call getMakeCommand(False, True, True) : Exit Function
    If mCmdInput.text = "omko" Then Call getMakeCommand(True, False, True) : Exit Function
    If InStr(mCmdInput.text, "smk") = 1 Then Call copySplitBuildCommand(Replace(mCmdInput.text, "smk", "")) : Exit Function
    If InStr(mCmdInput.text, "smo") = 1 Then Call getSplitTestOTABuildCommand(Replace(mCmdInput.text, "smo", "")) : Exit Function
    If mCmdInput.text = "md" Then Call MkdirWeibuFolderPath() : Exit Function
    If mCmdInput.text = "cm" Then Call CopyCommitInfo("") : Exit Function
    If InStr(mCmdInput.text, "cm-") = 1 Then Call CopyCommitInfo(Replace(mCmdInput.text, "cm-", "")) : Exit Function
    If mCmdInput.text = "ota" Then Call CopyBuildOtaUpdate() : Exit Function
    If mCmdInput.text = "cc" Then Call copyStrAndPasteInXshell("git checkout .;git clean -df") : Exit Function
    If InStr(mCmdInput.text, "qm-") = 1 Then Call CopyQmakeCmd(Replace(mCmdInput.text, "qm-", "")) : Exit Function
    If mCmdInput.text = "exp" Then Call copyStrAndPasteInXshell("export PATH=$HOME/Tools:$PATH") : Exit Function
    If mCmdInput.text = "mo" Then Call moveOutFoldersOut("") : Exit Function
    If InStr(mCmdInput.text, "mo-") = 1 Then Call moveOutFoldersOut(Replace(mCmdInput.text, "mo-", "")) : Exit Function
    If mCmdInput.text = "mui" Then Call moveOutFoldersIn("user", False) : Exit Function
    If mCmdInput.text = "mdi" Then Call moveOutFoldersIn("debug", False) : Exit Function
    If mCmdInput.text = "mui-f" Then Call moveOutFoldersIn("user", True) : Exit Function
    If mCmdInput.text = "mdi-f" Then Call moveOutFoldersIn("debug", True) : Exit Function
    If mCmdInput.text = "ps" Then Call copyStrAndPasteInXshell("git pull -r origin master && git push origin master") : Exit Function
    If mCmdInput.text = "ccps" Then Call copyStrAndPasteInXshell("git checkout .;git clean -df;git pull -r origin master && git push origin master") : Exit Function
    If mCmdInput.text = "update" Then Call copyStrAndPasteInXshell("git remote update origin --prune") : Exit Function
    If mCmdInput.text = "df" Then Call GetDiffCmdFromOverlayPath() : Exit Function
    If InStr(mCmdInput.text, "fd-") = 1 Then Call findProjectPathWithTaskNum(Replace(mCmdInput.text, "fd-", "")) : Exit Function
    handleLinuxCmd = False
End Function

Function handleMultiMkdirCmd()
    handleMultiMkdirCmd = True
    If mCmdInput.text = "md-lg" Then Call mkdirLogo() : Exit Function
    If mCmdInput.text = "md-ani" Then Call mkdirBootAnimation() : Exit Function
    If mCmdInput.text = "md-wp" Then Call mkdirWallpaper(False) : Exit Function
    If mCmdInput.text = "md-wp-go" Then Call mkdirWallpaper(True) : Exit Function
    If mCmdInput.text = "md-tee" Then Call mkdirTee() : Exit Function
    If mCmdInput.text = "md-spi" Then Call mkdirProductInfo("sys") : Exit Function
    If mCmdInput.text = "md-vpi" Then Call mkdirProductInfo("vnd") : Exit Function
    handleMultiMkdirCmd = False
End Function

Function handleOpenPathCmd()
    handleOpenPathCmd = True
    If mCmdInput.text = "fjava" Then Call findFileInList("f-java.txt", ".java", "find frameworks/ -type f -name *.java > f-java.txt") : Exit Function
    If mCmdInput.text = "java" Then Call findFileInList("java.txt", ".java", "find -type f -name *.java > java.txt") : Exit Function
    If mCmdInput.text = "kt" Then Call findFileInList("kt.txt", ".kt", "find -type f -name *.kt > kt.txt") : Exit Function
    If mCmdInput.text = "xml" Then Call findFileInList("xml.txt", ".xml", "find -type f -name *.xml > xml.txt") : Exit Function
    If mCmdInput.text = "app" Then Call findFileInList("app.txt", "app", "find -type f -name Android.* > app.txt") : Exit Function
    If mCmdInput.text = "cl" Then Call setOpenPath("") : Exit Function
    If mCmdInput.text = "addp" Then Call setOpenPath(mBuild.Infos.getOverlayPath(getOpenPath())) : Exit Function
    If mCmdInput.text = "cutp" Then Call setOpenPath(Split(getOpenPath(), "/alps/")(1)) : Exit Function
    If mCmdInput.text = "cp" Then Call compareForProject() : Exit Function
    If mCmdInput.text = "cs" Then Call selectForCompare() : Exit Function
    If mCmdInput.text = "ct" Then Call compareTo() : Exit Function
    If mCmdInput.text = "fmw" Then Call runPath(Replace("\\192.168.0.248\安卓固件文件1\" & mTask.Infos.CustomName, "纳斯达", "纳思达")) : Exit Function
    If mCmdInput.text = "req" Then Call runPath("\\192.168.0.24\wbshare\客户需求\" & mTask.Infos.CustomName) : Exit Function
    If mCmdInput.text = "zt" Then Call runWebsite("http://192.168.0.29:3000/zentao/task-view-" & mTask.Infos.TaskNum & ".html") : Exit Function
    handleOpenPathCmd = False
End Function

Function handleCopyCommandCmd()
    handleCopyCommandCmd = True
    If mCmdInput.text = "spp" Then Call getProjectPathWithTaskNum(mTask.Infos.TaskNum, "s") : Exit Function
    If mCmdInput.text = "vpp" Then Call getProjectPathWithTaskNum(mTask.Infos.TaskNum, "v") : Exit Function
    If InStr(mCmdInput.text, "spp-") = 1 Then Call getProjectPathWithTaskNum(Replace(mCmdInput.text, "spp-", ""), "s") : Exit Function
    If InStr(mCmdInput.text, "vpp-") = 1 Then Call getProjectPathWithTaskNum(Replace(mCmdInput.text, "vpp-", ""), "v") : Exit Function
    If mCmdInput.text = "outp" Then Call CopyString(mTask.Sys.Infos.DownloadOutPath) : Exit Function
    If mCmdInput.text = "winp" Then Call CopyString(mBuild.Infos.getPathWithDriveSdk(getOpenPath())) : Exit Function
    If mCmdInput.text = "lnxp" Then Call CopyString(mSdk & "\" & getOpenPath()) : Exit Function
    If InStr(mCmdInput.text, "ps-") = 1 Then Call CopyAdbPushCmd(Replace(mCmdInput.text, "ps-", "")) : Exit Function
    If InStr(mCmdInput.text, "cl-") = 1 Then Call CopyAdbClearCmd(Replace(mCmdInput.text, "cl-", "")) : Exit Function
    If InStr(mCmdInput.text, "st-") = 1 Then Call CopyAdbStartCmd(Replace(mCmdInput.text, "st-", "")) : Exit Function
    If InStr(mCmdInput.text, "dp-") = 1 Then Call CopyAdbDumpsysCmd(Replace(mCmdInput.text, "dp-", "")) : Exit Function
    If InStr(mCmdInput.text, "lg-") = 1 Then Call CopyAdbLogcatCmd(Replace(mCmdInput.text, "lg-", "")) : Exit Function
    If InStr(mCmdInput.text, "sts-") = 1 Then Call CopyAdbSettingsCmd(Replace(mCmdInput.text, "sts-", "")) : Exit Function
    If InStr(mCmdInput.text, "ins-") = 1 Then Call CopyAdbInstallCmd(Replace(mCmdInput.text, "ins-", "")) : Exit Function
    If mCmdInput.text = "gmsp" Then Call CopyAdbGetGmsPropCmd() : Exit Function
    If mCmdInput.text = "ss" Then Call copyStrAndPasteInCodeEditor() : Exit Function
    If mCmdInput.text = "hqp" Then Call sendWeiXinMsg("hqp") : Exit Function
    If mCmdInput.text = "zhq" Then Call sendWeiXinMsg("zhq") : Exit Function
    If mCmdInput.text = "lqj" Then Call sendWeiXinMsg("lqj") : Exit Function
    If mCmdInput.text = "lyh" Then Call sendWeiXinMsg("lyh") : Exit Function
    If mCmdInput.text = "wj" Then Call sendWeiXinMsg("wj") : Exit Function
    If mCmdInput.text = "getcl" Then Call getCommitMsgList() : Exit Function
    If mCmdInput.text = "getrl" Then Call getReleaseNoteList() : Exit Function
    handleCopyCommandCmd = False
End Function

Function handleEditTextCmd()
    handleEditTextCmd = True
    If InStr(mCmdInput.text, "bn=") > 0 Then Call modBuildNumber(Replace(mCmdInput.text, "bn=", "")) : Exit Function
    If InStr(mCmdInput.text, "-ota") > 0 Then Call copyStrAndPasteInXshell(getOTATestSedStr(True)) : Exit Function
    If InStr(mCmdInput.text, "tz=") > 0 Then Call modSystemprop(Split(mCmdInput.text, "=")) : Exit Function
    If InStr(mCmdInput.text, "loc=") > 0 Then Call modSystemprop(Split(mCmdInput.text, "=")) : Exit Function
    If InStr(mCmdInput.text, "ftd=") > 0 Then Call modSystemprop(Split(mCmdInput.text, "=")) : Exit Function
    If InStr(mCmdInput.text, "spdf-") > 0 Then Call modSettingProviderDefault(Split(mCmdInput.text, "-")(1)) : Exit Function
    If InStr(mCmdInput.text, "=") > 0 Then Call cpFileAndSetValue(Split(mCmdInput.text, "=")) : Exit Function
    handleEditTextCmd = False
End Function

Function handleProjectCmd()
    handleProjectCmd = True
    If isTaskNum(mCmdInput.text) Then
        If getTmpTaskWithNum(mCmdInput.text) Then Call setCurrentTask(mTmpTask)
        Exit Function
    ElseIf mCmdInput.text = "z6" Or mCmdInput.text = "x1" Or mCmdInput.text = "x2" Then
        Call setCurrentDrive(mCmdInput.text)
        Exit Function
    ElseIf mCmdInput.text = "s" Then
        Call setSysBuild()
        Exit Function
    ElseIf mCmdInput.text = "v" Then
        Call setVndBuild()
        Exit Function
    ElseIf mCmdInput.text = "v0" Then Call setCommonTask("v0") : Exit Function
    ElseIf mCmdInput.text = "u0" Then Call setCommonTask("u0") : Exit Function
    ElseIf mCmdInput.text = "t0" Then Call setCommonTask("t0") : Exit Function
    ElseIf mCmdInput.text = "s0" Then Call setCommonTask("s0") : Exit Function
    ElseIf InStr(mCmdInput.text, "show-") = 1 Then
        Call showWorkInfo(Replace(mCmdInput.text, "show-", ""))
        Exit Function
    ElseIf mCmdInput.text = "tpi" Then
        Call setOpenPath("mtk_sp_t0/vnd/" & VbLf & "mtk_sp_t0/sys/")
        Exit Function
    ElseIf mCmdInput.text = "upi" Then
        Call setOpenPath("mtk_sp_t0/vnd/" & VbLf & "mtk_sp_t0/u_sys/")
        Exit Function
    ElseIf mCmdInput.text = "vpi" Then
        Call setOpenPath("mtk_sp_t0/v_sys/" & VbLf & "mtk_sp_t0/v_sys/")
        Exit Function
    ElseIf mCmdInput.text = "svpi" Then
        Call setOpenPath("mtk_sp_t0/vnd/" & VbLf & "mtk_sp_t0/v_sys/")
        Exit Function
    ElseIf mCmdInput.text = "getpi" Then
        Call getProjectInfosFromOpenPath()
        Exit Function
    ElseIf mCmdInput.text = "save" Then
        Call handleForWorksInfo(Replace(getOpenPath(), VbLf, " | "))
        Call updateTaskList()
        Exit Function
    ElseIf mCmdInput.text = "tl" Then
        Call updateTaskList()
        Exit Function
    End If
    handleProjectCmd = False
End Function

Function handleCurrentDictCmd()
    handleCurrentDictCmd = True
    If mCmdInput.text = "config" Then Call runPath(PATH_CONFIG) : Exit Function
    If mCmdInput.text = "tl" Then Call runPath(PATH_TASK_LIST) : Exit Function
    If mCmdInput.text = "op" Then Call runPath(oWs.CurrentDirectory) : Exit Function
    handleCurrentDictCmd = False
End Function

Sub copyCdSdkCommand()
    Dim path, arr
    arr = Split(mDrive & mBuild.Sdk, ":\")
    path = getParentPath(relpaceSlashInPath(arr(1)))
    Call copyStrAndPasteInXshell("cd " & path)
End Sub

Function getSedCmd(cmdStr, searchStr, replaceStr, newStr, filePath)
    If isArray(filePath) Then
        Dim i, str
        str = cmdStr & "sed -i '/" & checkBackslash(searchStr) & "/s/" & checkBackslash(replaceStr) & "/" & checkBackslash(newStr) & "/'"
        For i = 0 To UBound(filePath)
            str = str & " " & filePath(i)
        Next
        getSedCmd = str & ";"
    Else
        getSedCmd = cmdStr & "sed -i '/" & checkBackslash(searchStr) & "/s/" & checkBackslash(replaceStr) & "/" & checkBackslash(newStr) & "/' " & filePath & ";"
    End If
End Function

Function getSedAddCmd(cmdStr, searchStr, addStr, filePath)
    getSedAddCmd = cmdStr & "sed -i '/" & checkBackslash(searchStr) & "/a\" & checkBackslash(addStr) & "' " & filePath & ";"
End Function

Function getGitDiffCmd(cmdStr, filePath)
    If isFileExists(mBuild.Infos.getOverlayPath(filePath)) Then
        getGitDiffCmd = cmdStr & "git diff " & mBuild.Infos.getOverlayPath(filePath) & ";"
    Else
        getGitDiffCmd = cmdStr & "git diff --no-index " & filePath & " " & mBuild.Infos.getOverlayPath(filePath) & ";"
    End If
End Function

Function getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, mode)
    Dim folderPath, cmdStr
    folderPath = getParentPath(filePath)

    If Not isFileExists(mBuild.Infos.getOverlayPath(filePath)) Then
        cmdStr = cmdStr & "mkdir -p " & mBuild.Infos.getOverlayPath(folderPath) & ";"
        cmdStr = cmdStr & "cp " & filePath & " " & mBuild.Infos.getOverlayPath(folderPath) & ";"
    End If
    cmdStr = getSedStr(cmdStr, mBuild.Infos.getOverlayPath(filePath), searchStr, startStr, valueStr, mode)
    cmdStr = getGitDiffCmd(cmdStr, filePath)

    getCpAndSedCmdStr = cmdStr
End Function

Function getSedStr(cmdStr, filePath, searchStr, startStr, valueStr, mode)
    If mode = "s" Then
        getSedStr = getSedCmd(cmdStr, searchStr, startStr & ".*$", startStr & valueStr, filePath)
    ElseIf mode = "ss" Then
        getSedStr = getSedCmd(cmdStr, searchStr, startStr, valueStr, filePath)
    ElseIf mode = "a" Then
        getSedStr = getSedAddCmd(cmdStr, searchStr, valueStr, filePath)
    End If
End Function

Function getMultiMkdirStr(arr, what)
    Dim str, path, ovlFolder, ovlFile
    For Each path In arr
        ovlFolder = mBuild.Infos.getOverlayPath(getParentPath(path))
        ovlFile = mBuild.Infos.getOverlayPath(path)
        If (Not (what = "lg" And InStr(path, "_kernel.bmp") > 0)) And (Not isFolderExists(ovlFolder)) Then
            str =  str & "mkdir -p " & ovlFolder & ";"
        End If

        If what = "lg" Then
            str =  str & "cp ../File/logo.bmp " & ovlFile & ";"
        ElseIf what = "ani" Then
            If InStr(path, "bootanimation.zip") > 0 Then
                str =  str & "cp ../File/bootanimation.zip " & ovlFolder & ";"
            ElseIf InStr(path, "products.mk") > 0 Then
                If Not isFileExists(ovlFile) Then str =  str & "cp " & path & " " & ovlFolder & ";"
                str =  getSedCmd(str, "bootanimation", "#", "", mBuild.Infos.getOverlayPath(path))
                str =  getGitDiffCmd(str, path)
            End If
        ElseIf what = "wp" Then
            If InStr(path, "default_wallpaper.png") > 0 Then
                str =  str & "cp ../File/default_wallpaper.png " & ovlFolder & ";"
            ElseIf InStr(path, "default_wallpaper.jpg") > 0 Then
                str =  str & "cp ../File/default_wallpaper.jpg " & ovlFolder & ";"
            End If
        Else
            str =  str & "cp " & path & " " & ovlFolder & ";"
        End If
    Next
    getMultiMkdirStr = str
End Function

Sub cpFileAndSetValue(whatArr)
    Dim cmdStr
    cmdStr = getCmdStrForCpFileAndSetValue(whatArr)
    If cmdStr = "" Then Exit Sub
    Call copyStrAndPasteInXshell(cmdStr)
End Sub

Function getCmdStrForCpFileAndSetValue(whatArr)
    Dim filePath, folderPath, keyStr, startStr, searchStr, valueStr, cmdStr
    If whatArr(0) = "gmsv" Then
        filePath = "vendor/partner_gms/products/gms_package_version.mk"
        keyStr = "GMS_PACKAGE_VERSION_ID"
        startStr = " := "
        searchStr = keyStr & startStr
        valueStr = whatArr(1)
        cmdStr = getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "s")

    ElseIf whatArr(0) = "sp" Then
        If mBuild.Infos.isV0() Then
            filePath = "build/make/core/version_util.mk"
        Else
            filePath = "build/make/core/version_defaults.mk"
        End If
        keyStr = "PLATFORM_SECURITY_PATCH"
        startStr = " := "
        searchStr = keyStr & startStr
        valueStr = whatArr(1)
        cmdStr = getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "s")

        filePath = "vendor/mediatek/proprietary/buildinfo_vnd/device.mk"
        keyStr = "VENDOR_SECURITY_PATCH"
        searchStr = keyStr & startStr
        cmdStr = cmdStr & getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "s")

    ElseIf whatArr(0) = "bn" Then
        Dim weibuConfig : weibuConfig = "build/make/core/weibu_config.mk"
        keyStr = "WEIBU_BUILD_NUMBER"
        startStr = " := "
        If isFileExists(weibuConfig) And mBuild.Infos.isVnd() Then
            filePath = weibuConfig
            startStr = " ?= "
        Else
            filePath = "device/mediatek/system/common/BoardConfig.mk"
            If InStr(mBuild.Sdk, "_r") > 0 Then
                keyStr = "BUILD_NUMBER_WEIBU"
            End If
        End If
        searchStr = keyStr & startStr
        valueStr = whatArr(1)
        cmdStr = getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "s")

    ElseIf whatArr(0) = "bn2" Then
        filePath = "device/mediatek/system/common/BoardConfig.mk"
        keyStr = "WEIBU_BUILD_NUMBER"
        startStr = " := "
        searchStr = keyStr & startStr
        valueStr = whatArr(1)
        cmdStr = getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "s")

        filePath = "device/mediatek/vendor/common/BoardConfig.mk"
        cmdStr = cmdStr & getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "s")
    
    ElseIf whatArr(0) = "fp" Then
        Dim version, buildId
        version = whatArr(1)
        If version = "15" Then
            buildId="AP3A.240905.015.A2"
        ElseIf version = "14" Then
            buildId="UP1A.231005.007"
        ElseIf version = "13" Then
            buildId="TP1A.220624.014"
        Else
            MsgBox("Unknown fp version: " & version)
            getCmdStrForCpFileAndSetValue = ""
            Exit Function
        End If
        cmdStr = getCpAndSedCmdStr("build/make/core/build_id.mk", "BUILD_ID", "=", buildId, "s")
        cmdStr = cmdStr & getCpAndSedCmdStr("build/make/core/sysprop.mk", "BUILD_FINGERPRINT := $(PRODUCT_BRAND)", "$(PLATFORM_VERSION)", version, "ss")

    ElseIf whatArr(0) = "bt" Then
        filePath = mBuild.Infos.getBluetoothfilePath()
        keyStr = "static char btif_default_local_name"
        startStr = " = "
        searchStr = keyStr
        valueStr = """&Chr(34)&""" & whatArr(1) & """&Chr(34)&"";"
        cmdStr = getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "s")

    ElseIf whatArr(0) = "mtp" Then
        filePath = "frameworks/base/media/java/android/mtp/MtpDatabase.java"
        startStr = " = "
        searchStr = "mDeviceProperties.getString"
        valueStr = """&Chr(34)&""" & whatArr(1) & """&Chr(34)&"";"
        cmdStr = getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "s")
        searchStr = "Build.MODEL"
        cmdStr = cmdStr & getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "s")
    
    ElseIf whatArr(0) = "wfap" Then
        filePath = "packages/modules/Wifi/service/java/com/android/server/wifi/WifiApConfigStore.java"
        startStr = "("
        searchStr = "configBuilder.setSsid(Build.MODEL)"
        valueStr = """&Chr(34)&""" & whatArr(1) & """&Chr(34)&"");"
        cmdStr = getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "s")

    ElseIf whatArr(0) = "wfdrt" Then
        filePath = "packages/modules/Wifi/service/java/com/android/server/wifi/p2p/WifiP2pServiceImpl.java"
        searchStr = "String getPersistedDeviceName()"
        valueStr = "            if (true) return ""&Chr(34)&""" & whatArr(1) & """&Chr(34)&"";"
        cmdStr = getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "a")

    ElseIf whatArr(0) = "brand" Or whatArr(0) = "model" Or whatArr(0) = "manufacturer" Then
        filePath = mBuild.Infos.ProductFile
        keyStr = "PRODUCT_" & UCase(whatArr(0))
        startStr = " := "
        searchStr = keyStr & startStr
        valueStr = whatArr(1)
        cmdStr = getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "s")

    ElseIf whatArr(0) = "name" Or whatArr(0) = "device" Then
        filePath = mBuild.Infos.ProductFile
        keyStr = "PRODUCT_SYSTEM_" & UCase(whatArr(0))
        startStr = " := "
        If mBuild.Product = "tb8765ap1_bsp_1g_k419" Or _
                mBuild.Product = "tb8766p1_64_bsp" Or _
                mBuild.Product = "tb8788p1_64_bsp_k419" Or _
                mBuild.Product = "tb8321p3_bsp" Or _
                mBuild.Product = "tb8768p1_64_bsp"  Then
            searchStr = keyStr & startStr
            valueStr = whatArr(1)
            cmdStr = getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "s")
        Else
            searchStr = "PRODUCT_BRAND"
            valueStr = "PRODUCT_SYSTEM_" & UCase(whatArr(0)) & startStr & whatArr(1)
            cmdStr = getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "a")
        End If

    ElseIf whatArr(0) = "brt" Then
        If isSplitSdkSys() Then Call setT0SdkVnd()
        filePath = "vendor/mediatek/proprietary/packages/overlay/vendor/FrameworkResOverlay/res/values/config.xml"
        startStr = ">"
        searchStr = "config_screenBrightnessSettingDefaultFloat"
        valueStr = whatArr(1) & "</item>"
        cmdStr = getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "s")

    ElseIf whatArr(0) = "bat" Then
        If mBuild.Infos.Version < 13 Then
            Call setVndBuild()
        Else
            Call setSysBuild()
        End If
        filePath = mBuild.Infos.getPowerProfilePath()
        startStr = ">"
        searchStr = "battery.capacity"
        valueStr = whatArr(1) & "</item>"
        cmdStr = getCpAndSedCmdStr(filePath, searchStr, startStr, valueStr, "s")

    End If
    getCmdStrForCpFileAndSetValue = cmdStr
End Function

Sub mkdirLogo()
    Call setVndBuild()
    Dim lg_fd, lg_u, lg_k, boot_logo
    lg_fd = mBuild.Infos.LogoPath & "/"
    boot_logo = mBuild.Infos.BootLogo
    lg_u = lg_fd & boot_logo & "_uboot.bmp"
    lg_k = lg_fd & boot_logo & "_kernel.bmp"

    Dim arr, finalStr
    arr = Array(lg_u, lg_k)
    finalStr = getMultiMkdirStr(arr, "lg")
    Call copyStrAndPasteInXshell(finalStr)
End Sub

Sub mkdirBootAnimation()
    Dim ani_media, ani_product
    ani_media = "vendor/weibu_sz/media/bootanimation.zip"
    ani_product = "vendor/weibu_sz/products/products.mk"

    Dim arr, finalStr
    arr = Array(ani_media, ani_product)
    finalStr = getMultiMkdirStr(arr, "ani")
    Call copyStrAndPasteInXshell(finalStr)
End Sub

Sub mkdirWallpaper(go)
    Dim wp_gms, wp_go1, wp_go2, wp1, wp2, wp3
    If mBuild.Infos.Version > 12 Then
        wp_gms = "vendor/partner_gms/overlay/AndroidGmsBetaOverlay/res/drawable-nodpi/default_wallpaper.png"
    Else
        wp_gms = "vendor/partner_gms/overlay/AndroidSGmsBetaOverlay/res/drawable-nodpi/default_wallpaper.png"
    End If
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
    Call copyStrAndPasteInXshell(finalStr)
End Sub

Sub mkdirTee()
    Call setVndBuild()
    Dim teeOverlayPath, finalStr, teeOverlayPath_8781
    teeOverlayPath = mBuild.Infos.getOverlayPath("vendor/mediatek/proprietary/trustzone/trustkernel/source/build/" & mBuild.Product)
    teeOverlayPath_8781 = mBuild.Infos.getOverlayPath("vendor/mediatek/proprietary/trustzone/trustkernel/source/build/hal_mgvi_t_64_armv82")

    If Not isFolderExists(teeOverlayPath) Then finalStr = finalStr & "mkdir -p " & teeOverlayPath & ";"
    finalStr = finalStr & "cp ../File/cert.dat " & teeOverlayPath & ";"
    finalStr = finalStr & "cp ../File/array.c " & teeOverlayPath & ";"

    If mBuild.Infos.is8781() Then
        If Not isFolderExists(teeOverlayPath_8781) Then finalStr = finalStr & "mkdir -p " & teeOverlayPath_8781 & ";"
        finalStr = finalStr & "cp ../File/cert.dat " & teeOverlayPath_8781 & ";"
        finalStr = finalStr & "cp ../File/array.c " & teeOverlayPath_8781 & ";"
    End If
    Call copyStrAndPasteInXshell(finalStr)
End Sub

Sub mkdirProductInfo(where)
    If Not isFileExists("../File/product.txt") Then MsgBox("product.txt does not exist!") : Exit Sub
    If where = "sys" Then Call setSysBuild()
    If where = "vnd" Then Call setVndBuild()
    Dim info, infoArr, infoDict, cmdStr, finalStr
    infoArr = Array("brand", "manufacturer", "model", "name", "device")
    Set infoDict = CreateObject("Scripting.Dictionary")
    For Each info In infoArr
        Call infoDict.Add(info, readTextAndGetValue(info, "../File/product.txt"))
    Next
    For Each info In infoArr
        If infoDict.Item(info) <> "" Then
            cmdStr = cmdStr & getCmdStrForCpFileAndSetValue(Array(info, infoDict.Item(info)))
        End If
    Next
    Dim cmd, cmdArr, mkdirResult, cpResult, diffStr
    cmdArr = Split(cmdStr, ";")
    mkdirResult = False
    cpResult = False
    For Each cmd In cmdArr
        If InStr(cmd, "mkdir") = 1 Then
            If Not mkdirResult Then
                finalStr = finalStr & cmd & ";"
                mkdirResult = True
            End If
        ElseIf InStr(cmd, "cp") = 1 Then
            If Not cpResult Then
                finalStr = finalStr & cmd & ";"
                cpResult = True
            End If
        ElseIf InStr(cmd, "git diff") = 1 Then
            diffStr = cmd
        Else
            If cmd <> "" Then finalStr = finalStr & cmd & ";"
        End If
    Next
    finalStr = finalStr & diffStr
    Call copyStrAndPasteInXshell(finalStr)
End Sub

Function getFileListPathFromRes(name)
    getFileListPathFromRes = relpaceBackSlashInPath(oWs.CurrentDirectory & "\res\filelist\" & Replace(mBuild.Sdk, "alps", "") & "\" & name)
End Function

Sub addFileList()
    If mFileButtonList.VaArray.Bound = 0 Then
        Call setOpenPath(mFileButtonList.VaArray.V(0))
        Call mFileButtonList.VaArray.ResetArray()
    ElseIf mFileButtonList.VaArray.Bound > 0 Then
        Call mFileButtonList.addList()
        Call mFileButtonList.toggleButtonList()
    End If
End Sub

Sub makeFileList(fileListPath, suffix)
    If Trim(getOpenPath()) = "" Or InStr(getOpenPath(), "/") > 0 Or InStr(getOpenPath(), "\") > 0 Then Exit Sub

    If suffix <> "app" Then
        Call findFileInListText(getOpenPath(), suffix, fileListPath)
    Else
        Call findAppFolderInListText(getOpenPath(), fileListPath)
    End If

    Call addFileList()
End Sub

Sub findFileInListText(input, suffix, path)
    Dim oText, sReadLine, keyStr, count
    If Not InStr(input, ".") > 0 Then
        keyStr = input & suffix
    Else
        keyStr = input
    End If
    keyStr = "/" & keyStr
    Set oText = oFso.OpenTextFile(path, FOR_READING)
    count = 0

    Do Until oText.AtEndOfStream
        sReadLine = oText.ReadLine
        If count > 10 Then Exit Do
        If Right(sReadLine, Len(keyStr)) = keyStr Then
            Call mFileButtonList.VaArray.append(Replace(sReadLine, "./", ""))
            count = count + 1
        End If
    Loop

    oText.Close
    Set oText = Nothing
End Sub

Sub findAppFolderInListText(input, path)
    Dim oText, sReadLine, keyStr, count
    keyStr = "/" & input & "/Android."
    Set oText = oFso.OpenTextFile(path, FOR_READING)
    count = 0

    Do Until oText.AtEndOfStream
        If count > 10 Then Exit Do
        sReadLine = oText.ReadLine
        If InStr(sReadLine, keyStr) > 0 Then
            path = Left(sReadLine, InStr(sReadLine, keyStr) + Len(input))
            Call mFileButtonList.VaArray.append(Replace(path, "./", ""))
            count = count + 1
        End If
    Loop
End Sub

Sub findFileInList(fileList, fileType, cmdStr)
    If mFileButtonList.hideListIfShowing() Then Exit Sub

    Dim fileListPath : fileListPath = getFileListPathFromRes(fileList)
    If isFileExists(fileListPath) Then
        Call makeFileList(getFileListPathFromRes(fileList), fileType)
    Else
        Call CopyString(cmdStr)
        MsgBox(fileListPath & " not exist!")
    End If
End Sub

Sub compareForProject()
    Dim inputPath, wholePath
    inputPath = getOpenPath()

    If isFileExists(inputPath) Or isFolderExists(inputPath) Then
        If InStr(inputPath, "weibu") > 0 Then
            mLeftComparePath = inputPath
            mRightComparePath = Split(inputPath, "/alps/")(1)
        Else
            mLeftComparePath = mBuild.Infos.getOverlayPath(inputPath)
            mRightComparePath = inputPath
        End If

        mLeftComparePath = """" & mLeftComparePath & """"
        mRightComparePath = """" & mRightComparePath & """"

        Call runBeyondCompare(mLeftComparePath, mRightComparePath)
    Else
        MsgBox("Not found :" & Vblf & inputPath)
    End If
End Sub

Sub selectForCompare()
    Dim inputPath
    inputPath = getOpenPath()

    If isFileExists(inputPath) Or isFolderExists(inputPath) Then
        mLeftComparePath = """" & inputPath & """"
    Else
        MsgBox("Not found :" & Vblf & inputPath)
    End If
End Sub

Sub compareTo()
    Dim inputPath
    inputPath = getOpenPath()

    If isFileExists(inputPath) Or isFolderExists(inputPath) Then
        mRightComparePath = """" & inputPath & """"
        Call runBeyondCompare(mLeftComparePath, mRightComparePath)
    Else
        MsgBox("Not found :" & Vblf & inputPath)
    End If
End Sub

Sub findProjectPathWithTaskNum(taskNum)
    Dim commandFinal
    If isTaskNum(taskNum) Then
        commandFinal = "find weibu -maxdepth 2 -name ""&Chr(34)&""*" & taskNum & "*""&Chr(34)&"""
    End If
    Call copyStrAndPasteInXshell(commandFinal)
End Sub

Sub getProjectPathWithTaskNum(taskNum, which)
    If Not getTmpTaskWithNum(taskNum) Then Exit Sub
    If which = "s" Then
        Call setOpenPath("weibu/" & mTmpTask.Sys.Product & "/" & mTmpTask.Sys.Project)
    Else
        Call setOpenPath("weibu/" & mTmpTask.Vnd.Product & "/" & mTmpTask.Vnd.Project)
    End If
End Sub

Function getVVLunchStr(task, buildType)
    If task.Vnd.Product <> "" And task.Vnd.Project <> "" And task.Sys.Product <> "" And task.Sys.Project <> "" Then
        getVVLunchStr = "vnd_" & task.Vnd.Product & "-next-" & buildType & " " & task.Vnd.Project &_
                " sys_" & task.Sys.Product & "-next-" & buildType & " " & task.Sys.Project
    Else
        getVVLunchStr = ""
    End If
End Function

Function getSULunchStr(task, buildType, androidVer)
    If task.Vnd.Product <> "" And task.Vnd.Project <> "" And task.Sys.Product <> "" And task.Sys.Project <> "" Then
        getSULunchStr = " vnd_" & task.Vnd.Product & "-" & buildType & " " & task.Vnd.Project &_
                " sys_" & task.Sys.Product & "-" & buildType & " " & task.Sys.Project &_
                androidVer
    Else
        getSULunchStr = ""
    End If
End Function

Function get8781LunchStr(task, buildType, releseStr, androidVer)
    If task.Vnd.Product <> "" And task.Vnd.Project <> "" And task.Sys.Product <> "" And task.Sys.Project <> "" Then
        get8781LunchStr = "hal_mgvi_t_64_armv82-" & buildType &_
                " krn_mgk_64_entry_level_k510-" & buildType &_
                " vext_" & task.Vnd.Product & "-" & buildType & " " & task.Vnd.Project &_
                " sys_" & task.Sys.Product & releseStr & buildType & " " & task.Sys.Project &_
                androidVer
    Else
        get8781LunchStr = ""
    End If
End Function

Function getLunchStrFromWSavedWork(buildType, task)
    Dim lunchStr
    If task.Sys.Infos.isV0() Then
        If Not task.Vnd.Infos.is8781() Then
            lunchStr = getVVLunchStr(task, buildType)
        Else
            lunchStr = get8781LunchStr(task, buildType, "-next-", " V")
        End If
    ElseIf task.Sys.Infos.isU0() Then
        If Not task.Vnd.Infos.is8781() Then
            lunchStr = getSULunchStr(task, buildType, " U")
        Else
            lunchStr = get8781LunchStr(task, buildType, "-", " U")
        End If
    ElseIf task.Sys.Infos.isT0() Then
        If Not task.Vnd.Infos.is8781() Then
            lunchStr = getSULunchStr(task, buildType, " T")
        Else
            lunchStr = get8781LunchStr(task, buildType, "-", " T")
        End If
    ElseIf task.Sys.Infos.is8168() Then
        Dim sysStr, vndStr
        sysStr = "sys_" & mTask.Sys.Product & "-" & buildType
        vndStr = "vnd_" & mTask.Vnd.Product & "-" & buildType
        lunchStr = sysStr & " " & vndStr & " " & mTask.Sys.Project
        lunchStr = "lunch_item=""&Chr(34)&""" & lunchStr & """&Chr(34)&"""
    Else
        lunchStr = ""
    End If

    getLunchStrFromWSavedWork = lunchStr
End Function

Function getLunchCommandInSplitBuild(buildType, task)
    Dim lunchStr, commandStr
    If Not task.Sys.Infos.is8168() Then
        lunchStr = getLunchStrFromWSavedWork(buildType, task)
        If lunchStr = "" Then getLunchCommandInSplitBuild = "" : Exit Function
        commandStr = "sed -i 's/^.*$/" & lunchStr & "/' lunch_item"
        If InStr(task.Vnd.Product, "tb8781") Then commandStr = commandStr & "_v2"
    Else
        lunchStr = getLunchStrFromWSavedWork(buildType, task)
        If lunchStr = "" Then getLunchCommandInSplitBuild = "" : Exit Function
        Dim keyStr
        keyStr = "##Cusomer Settings"
        commandStr = "sed -i '/" & keyStr & "/i\" & lunchStr & "' split_build.sh;git diff split_build.sh"
    End If
    getLunchCommandInSplitBuild = commandStr
End Function

Sub getLunchCommand(buildType, taskNum)
    If Not getTmpTaskWithNum(taskNum) Then Exit Sub
    Dim commandFinal, comboName
    If mTmpTask.Sys.Infos.Version > 12 Or mTmpTask.Sys.Infos.is8168() Then
        commandFinal = getLunchCommandInSplitBuild(buildType, mTmpTask)
        If mTmpTask.Vnd.Infos.isV0() Then
            commandFinal = "cd ~" & relpaceSlashInPath(Split(mDrive, ":")(1)) & mTmpTask.Vnd.Sdk & " && " & commandFinal
        End If
    Else
        comboName = "full_" & mBuild.Product & "-" & buildType
        commandFinal = "source build/envsetup.sh ; lunch " & comboName & " " & mBuild.Project
    End If
    Call copyStrAndPasteInXshell(commandFinal)
End Sub

Sub getSysLunchCommand(buildType)
    Dim commandFinal
    If mTask.Sys.Infos.isV0() THen
        commandFinal = "source build/envsetup.sh && lunch sys_" & mTask.Sys.Product & "-next-" & buildType & " " & mTask.Sys.Project
    Else
        commandFinal = "source build/envsetup.sh && lunch sys_" & mTask.Sys.Product & "-" & buildType & " " & mTask.Sys.Project
    End If
    Call copyStrAndPasteInXshell(commandFinal)
End Sub

Sub getVndLunchCommand(buildType)
    Dim commandFinal
    If mTask.Vnd.Infos.is8781() Then
        commandFinal = "source build/envsetup.sh && lunch vext_" & mTask.Vnd.Product & "-" & buildType & " " & mTask.Vnd.Project
    ELse
        commandFinal = "source build/envsetup.sh && lunch vnd_" & mTask.Vnd.Product & "-" & buildType & " " & mTask.Vnd.Project
    End If
    Call copyStrAndPasteInXshell(commandFinal)
End Sub

Sub getMakeCommand(rmOut, rmBuildprop, ota)
    Dim commandOta, commandFinal
    commandFinal = "make -j36 2>&1 | tee build.log"
    commandOta = "make -j36 otapackage 2>&1 | tee build_ota.log"
    
    If rmOut Then
        commandFinal = "rm -rf out/ && " & commandFinal
    ElseIf rmBuildprop Then
        commandFinal = "find " & mBuild.Infos.OutPath & " -type f -name build*.prop | xargs rm -v && " & commandFinal
    End If

    If ota Then commandFinal = commandFinal & " && " & commandOta

    Call copyStrAndPasteInXshell(commandFinal)
End Sub

Function getCustomModemSedStr()
    If Not isFileExists(mTask.Vnd.Infos.ProjectConfigMk) Then getCustomModemSedStr = "" : Exit Function
    Dim customModem, deviceModem
    customModem = readTextAndGetValue("CUSTOM_MODEM", mTask.Vnd.Infos.ProjectConfigMk)
    deviceModem = readTextAndGetValue("CUSTOM_MODEM", mTask.Vnd.Infos.OriginProjectConfigMk)
    if customModem <> "" And customModem <> deviceModem Then
        If mTask.Vnd.Infos.isV0() Then
            getCustomModemSedStr = getSedCmd("", "CUSTOM_MODEM", "=.*$", "= " & customModem, mTask.Vnd.Infos.OriginProjectConfigMk)
        Else
            getCustomModemSedStr = getSedCmd("", "CUSTOM_MODEM", "=.*$", "= " & customModem, "vnd/" & mTask.Vnd.Infos.OriginProjectConfigMk)
        End If
    Else
        getCustomModemSedStr = ""
    End If
End Function

Function checkBuildTask()
    If Not isFolderExists(mBuild.Infos.Out) Then checkBuildTask = True : Exit Function
    Dim outDisplayId, lunchItemPath, lunchStr, index
    Call setSysBuild()
    outDisplayId = mBuild.Infos.getOutProp("ro.build.display.inner.id")
    If mBuild.Infos.is8781() Then
        lunchItemPath = "../lunch_item_v2"
    ElseIf mBuild.Infos.isV0() Then
        lunchItemPath = "lunch_item"
    Else
        lunchItemPath = "../lunch_item"
    End If
    lunchStr = readLineOfTextFile(1, lunchItemPath)
    If mBuild.Infos.is8781() Then
        index = 3
    Else
        index = 1
    End If
    If InStr(outDisplayId, Replace(Split(lunchStr, " ")(index), "_", ".")) > 0 Then
        checkBuildTask = True
    Else
        checkBuildTask = False
        MsgBox("Diff lunch item and out!")
    End If
End Function

Function checkBuildType()
    If Not isFolderExists(mBuild.Infos.Out) Then checkBuildType = True : Exit Function
    Dim outBuildType, lunchItemPath, lunchStr
    Call setSysBuild()
    outBuildType = mBuild.Infos.getOutProp("ro.build.type")
    If mBuild.Infos.is8781() Then
        lunchItemPath = "../lunch_item_v2"
    ElseIf mBuild.Infos.isV0() Then
        lunchItemPath = "lunch_item"
    Else
        lunchItemPath = "../lunch_item"
    End If
    lunchStr = readLineOfTextFile(1, lunchItemPath)
    If InStr(lunchStr, outBuildType & " ") Then
        checkBuildType = True
    Else
        checkBuildType = False
        MsgBox("Wrong build type!" & VbLf & "out=" & outBuildType & VbLf & "lunch_item=" & lunchStr)
    End If
End Function

Function getSplitBuildCommand(opts)
    Call checkBuildTask()
    If Not checkBuildType() Then getSplitBuildCommand = "" : Exit Function
    Dim buildsh, params, commandStr
    If mBuild.Infos.is8781() Then
        buildsh = "./split_build_v2.sh"
    Else
        buildsh = "./split_build.sh"
    End If
    If opts = "a" Then
        If mBuild.Infos.is8781() Then
            params = " vext krn hal sys m p"
        ElseIf mTask.Vnd.Infos.isV0() Then
            params = " krn vnd sys m p"
        Else
            params = " vnd krn sys m p"
        End If
    Else
        If InStr(opts, "v") Then
            If mBuild.Infos.is8781() Then
                params = params & " vext"
            Else
                params = params & " vnd"
            End If
        End If
        If InStr(opts, "k") Then params = params & " krn"
        If mBuild.Infos.is8781() And isInStr(opts, "h") Then params = params & " hal"
        If InStr(opts, "s") Then params = params & " sys"
        If InStr(opts, "m") Then params = params & " m"
        If InStr(opts, "p") Then params = params & " p"
    End if

    commandStr = buildsh & params

    if InStr(params, " p") > 0 And InStr(params, " vnd") = 0 And InStr(params, " vext") = 0 And InStr(params, " krn") = 0 And InStr(params, " hal") = 0 Then
        Call setVndBuild()
        commandStr = getCustomModemSedStr() & commandStr
    End If
    getSplitBuildCommand = commandStr & ";"
End Function

Sub copySplitBuildCommand(opts)
    Call copyStrAndPasteInXshell(getSplitBuildCommand(opts))
End Sub

Function getOTATestSedStr(showDiff)
    Dim buildinfo, keyStr, sedStr
    buildinfo = getMultiOverlayPath("build/make/tools/buildinfo.sh")

    keyStr = "ro.build.display.id"
    If InStr(buildinfo, "/config/") Then
        sedStr = "sed -i '/" & keyStr & "/s/$/-OTA_test/' " & buildinfo
    Else
        sedStr = "sed -i '/" & keyStr & "/s/""&Chr(34)&""$/-OTA_test""&Chr(34)&""/' " & buildinfo
    End If
    If showDiff Then
        sedStr = sedStr & "; git diff " & buildinfo
    End If
    getOTATestSedStr = sedStr
End Function

Sub getSplitTestOTABuildCommand(opts)
    Dim cmdStr
    cmdStr = getSplitBuildCommand(opts)
    
    If mTask.Vnd.Infos.isV0() Then
        cmdStr = cmdStr & "mkdir -p ../OTA/" & mTask.Infos.TaskNum & ";"
        cmdStr = cmdStr & "mv merged/target_files.zip ../OTA/" & mTask.Infos.TaskNum & "/target_files_s.zip;"
        cmdStr = cmdStr & getOTATestSedStr(False) & ";"
        cmdStr = cmdStr & getSplitBuildCommand("sm")
    Else
        cmdStr = cmdStr & "mkdir -p OTA/" & mTask.Infos.TaskNum & ";"
        cmdStr = cmdStr & "mv merged/target_files.zip OTA/" & mTask.Infos.TaskNum & "/target_files_s.zip;"
        If isSplitSdkVnd() Then Call setT0SdkSys()
        cmdStr = cmdStr & "cd " & getFileNameFromPath(mTask.Sys.Sdk) & ";"
        cmdStr = cmdStr & getOTATestSedStr(False) & ";"
        cmdStr = cmdStr & "cd ..;"
        cmdStr = cmdStr & getSplitBuildCommand("sm")
    End If
    Call copyStrAndPasteInXshell(cmdStr)
End Sub

Sub MkdirWeibuFolderPath()
    Dim commandFinal
    If Not isFileExists(getOpenPath()) Then
        If Not isFolderExists(getOpenPath()) Then
            MsgBox("File or Folder not exist! " & getOpenPath())
            Exit Sub
        Else
            commandFinal = "mkdir -p " & mBuild.Infos.getOverlayPath(getOriginPathFromOverlayPath(getOpenPath())) & ";"
        End If
    Else
        Dim filePath : filePath = getOriginPathFromOverlayPath(getOpenPath())
        'If Not isFileExists(filePath) Then MsgBox("File not exist! " & filePath) : Exit Sub

        Dim folderPath
        Dim overlayFilePath
        Dim overlayFolderPath
        Dim mkdirCmd, cpCmd

        folderPath = getParentPath(filePath)
        overlayFilePath = mBuild.Infos.getOverlayPath(filePath)
        overlayFolderPath = mBuild.Infos.getOverlayPath(folderPath)

        If isFileExists(overlayFilePath) Then MsgBox("File exist! " & overlayFilePath) : Exit Sub
        If Not isFolderExists(overlayFolderPath) Then
            mkdirCmd = "mkdir -p " & overlayFolderPath & ";"
        End If

        Dim multiOverlayFile
        multiOverlayFile = getMultiOverlayPath(filePath)
        If multiOverlayFile <> filePath Then
            filePath = multiOverlayFile
        ElseIf getOpenPath() <> filePath Then
            filePath = getOpenPath()
        End If 

        cpCmd = "cp " & filePath & " " & mBuild.Infos.getOverlayPath(folderPath)
        commandFinal = mkdirCmd & cpCmd
    End If

    commandFinal = relpaceSlashInPath(commandFinal)
    Call setOpenPath(overlayFilePath)
    Call copyStrAndPasteInXshell(commandFinal)
End Sub

Sub CopyCommitInfo(what)
     Dim commandFinal
    If what = "" Then
        commandFinal = "[" & mBuild.Project & "] : "
        Call copyStrAndPasteInXshell(commandFinal)
        Exit Sub

    ElseIf what = "lg" Then
        commandFinal = "Logo [" & mBuild.Project & "] : 客制开机logo"
    ElseIf what = "ani" Then
        commandFinal = "BootAnimation [" & mBuild.Project & "] : 客制开机动画"
    ElseIf what = "wp" Then
        commandFinal = "Wallpaper [" & mBuild.Project & "] : 客制默认壁纸"
    ElseIf what = "loc" Then
        commandFinal = "Locale [" & mBuild.Project & "] : 默认语言"
    ElseIf what = "tz" Then
        commandFinal = "Timezone [" & mBuild.Project & "] : 默认时区"
    ElseIf what = "di" Then
        commandFinal = "DisplayId [" & mBuild.Project & "] : 版本号"
    ElseIf InStr(what, "bn=") = 1 Then
        commandFinal = "BuildNumber [" & mBuild.Project & "] : 固定指纹信息（build number " & Replace(what, "bn=", "") & "）"
    ElseIf InStr(what, "sp=") = 1 Then
        If InStr(what, "-bn=") > 0 Then
            commandFinal = "GMS [" & mBuild.Project & "] : 固定GMS信息（安全补丁日期 " & Replace(Split(what, "-bn=")(0), "sp=", "") & "、build number " & Split(what, "-bn=")(1) & "）"
        Else
            commandFinal = "GMS [" & mBuild.Project & "] : 固定安全补丁日期 " & Replace(what, "sp=", "")
        End If
    ElseIf what = "bm" Then
        commandFinal = "MMI [" & mBuild.Project & "] : 品牌，型号"
    ElseIf what = "bwm" Then
        commandFinal = "MMI [" & mBuild.Project & "] : 蓝牙、WiFi热点、WiFi直连、盘符"
    ElseIf what = "mmi" Then
        commandFinal = "MMI [" & mBuild.Project & "] : "
    ElseIf what = "st" Then
        commandFinal = "Settings [" & mBuild.Project & "] : "    
    ElseIf what = "su" Then
        commandFinal = "SystemUI [" & mBuild.Project & "] : "
    ElseIf what = "lc" Then
        commandFinal = "Launcher [" & mBuild.Project & "] : "
    ElseIf what = "cam" Then
        commandFinal = "Camera [" & mBuild.Project & "] : "
    ElseIf what = "bt" Then
        commandFinal = "Bluetooth [" & mBuild.Project & "] : 默认蓝牙名称"
    ElseIf what = "wfap" Then
        commandFinal = "WiFi [" & mBuild.Project & "] : 默认WiFi热点名称"
    ElseIf what = "wfdrt" Then
        commandFinal = "WiFi [" & mBuild.Project & "] : 默认WiFi直连名称"
    ElseIf what = "mtp" Then
        commandFinal = "MTP [" & mBuild.Project & "] : 默认盘符名称"
    ElseIf what = "brt" Then
        commandFinal = "Brightness [" & mBuild.Project & "] : 默认亮度%"
    ElseIf what = "ad" Then
        commandFinal = "Audio [" & mBuild.Project & "] : 默认音量%"
    ElseIf what = "slp" Then
        commandFinal = "Settings [" & mBuild.Project & "] : 默认休眠时间"
    ElseIf what = "bat" Then
        commandFinal = "Battery [" & mBuild.Project & "] : 电池检测容量mAh"
    ElseIf what = "ft" Then
        commandFinal = "FactoryTest [" & mBuild.Project & "] : "
    ElseIf what = "tv" Then
        commandFinal = "TextView [" & mBuild.Project & "] : "
    Else
        commandFinal = what & " [" & mBuild.Project & "] : "
    End If
    Call copyStrAndPasteInXshell("git add weibu;git commit -m ""&Chr(34)&""" & commandFinal & """&Chr(34)&""")
End Sub

Sub CopyBuildOtaUpdate()
    Dim commandFinal
    If mTask.Sys.Infos.Version > 11 Then
        commandFinal = "./" & mTask.Sys.Infos.Out & "/host/linux-x86/bin/ota_from_target_files -i target_files_.zip target_files.zip update_.zip"
    Else
        commandFinal = "./build/tools/releasetools/ota_from_target_files -i old.zip new.zip update.zip"
    End If
    Call copyStrAndPasteInXshell(commandFinal)
End Sub

Sub CopyAdbPushCmd(which)
    Dim outPath, sourcePath, targetPath, finalStr
    outPath = mBuild.Infos.getPathWithDriveSdk(mBuild.Infos.OutPath)
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
    ElseIf which = "fws" Then
        sourcePath = outPath & "\system\framework\services.jar"
        targetPath = "/system/framework"
        finalStr = "adb push " & sourcePath & " " & targetPath
    ElseIf which = "ft" Then
        sourcePath = outPath & "\system\priv-app\FactoryTest"
        targetPath = "/system/priv-app/"
        finalStr = "adb push " & sourcePath & " " & targetPath
    ElseIf which = "sr" Then
        sourcePath = outPath & "\system\app\SoundRecorder_old"
        targetPath = "/system/app/"
        finalStr = "adb push " & sourcePath & " " & targetPath
    ElseIf which = "fwr" Then
        sourcePath = outPath & "\system\framework\framework-res.apk"
        targetPath = "/system/framework/"
        finalStr = "adb push " & sourcePath & " " & targetPath
    End If

    If mTask.Sys.Infos.Version > 13 Then
        finalStr = Replace(finalStr, "\system\system_ext\", "\system_ext\")
        finalStr = Replace(finalStr, "/system/system_ext/", "/system_ext/")
    End If
    Call copyStrAndPasteInPowerShell(finalStr)
End Sub

Sub CopyAdbClearCmd(which)
    Dim finalStr
    If which = "su" Then
        finalStr = "adb shell stop;adb shell start"
    ElseIf which = "st" Then
        finalStr = "adb shell pm clear com.android.settings"
    ElseIf which = "lc" Then
        finalStr = "adb shell pm clear com.android.launcher3"
    ElseIf which = "cam" Then
        finalStr = "adb shell pm clear com.mediatek.camera"
    ElseIf which = "ft" Then
        finalStr = "adb shell pm clear com.weibu.factorytest"
    ElseIf which = "sr" Then
        finalStr = "adb shell pm clear com.android.soundrecorder"
    End If
    Call copyStrAndPasteInPowerShell(finalStr)
End Sub

Sub CopyAdbStartCmd(which)
    Dim finalStr
    If which = "ab" Then
        finalStr = "adb shell am start -a android.settings.DEVICE_INFO_SETTINGS"
    ElseIf which = "log" Then
        finalStr = "adb shell am start -n com.debug.loggerui/.MainActivity"
    ElseIf which = "ft" Then
        finalStr = "adb shell am start -n com.weibu.factorytest/.FactoryTest"
    End If
    Call copyStrAndPasteInPowerShell(finalStr)
End Sub

Sub CopyAdbDumpsysCmd(which)
    Dim finalStr
    If which = "a" Then
        finalStr = "adb shell ""&Chr(34)&""dumpsys activity top | grep ACTIVITY | tail -n 1""&Chr(34)&"""
    ElseIf which = "f" Then
        finalStr = "adb shell ""&Chr(34)&""dumpsys activity top | grep '#[0-9]: ' | tail -n 1""&Chr(34)&"""
    ElseIf which = "r" Then
        finalStr = "adb shell ""&Chr(34)&""dumpsys activity activities | grep '* ActivityRecord{'""&Chr(34)&"""
    ElseIf which = "temp" Then
        finalStr = "adb shell dumpsys battery set temp"
    ElseIf which = "level" Then
        finalStr = "adb shell dumpsys battery set level"
    ElseIf which = "su" Then
        finalStr = "adb shell ""&Chr(34)&""dumpsys activity service com.android.systemui | grep --color""&Chr(34)&"""
    End If
    Call copyStrAndPasteInPowerShell(finalStr)
End Sub

Sub CopyAdbLogcatCmd(which)
    Dim finalStr
    If which = "as" Then
        finalStr = "adb shell ""&Chr(34)&""logcat -s  ActivityTaskManager | grep START""&Chr(34)&"""
    End If
    Call copyStrAndPasteInPowerShell(finalStr)
End Sub

Sub CopyAdbSettingsCmd(which)
    Dim finalStr
    If which = "sec" Then
        finalStr = "adb shell settings put secure clock_seconds 1"
    ElseIf which = "brt" Then
        finalStr = "adb shell settings get system screen_brightness"
    End If
    Call copyStrAndPasteInPowerShell(finalStr)
End Sub

Sub CopyAdbInstallCmd(which)
    Dim finalStr
    If which = "att" Then
        finalStr = "adb install D:\APK\Antutu\antutu-benchmark-v10.apk"
    ElseIf which = "aida" Then
        finalStr = "adb install D:\APK\Antutu\aida64-v198.apk"
    ElseIf which = "dvc" Then
        finalStr = "adb install D:\APK\Antutu\DevCheck_v5.11_Mod.apk"
    ElseIf which = "z" Then
        finalStr = "adb install D:\APK\Antutu\CPU-Z-1.43.apk"
    ElseIf which = "hw" Then
        finalStr = "adb install D:\APK\Antutu\DeviceInfoHW-v5.20.1.apk"
    End If
    Call copyStrAndPasteInPowerShell(finalStr)
End Sub

Sub CopyAdbGetGmsPropCmd()
    Dim finalStr
    finalStr = "adb shell ""&Chr(34)&""getprop | grep fingerprint | grep -v ro.bootimage.build.fingerprint | grep -v preview_sdk_fingerprint""&Chr(34)&"";adb shell ""&Chr(34)&""getprop | grep -E 'security_patch|gmsversion|base_os|first_api_level|clientidbase'""&Chr(34)&"""
    Call copyStrAndPasteInPowerShell(finalStr)
End Sub

Sub CopyQmakeCmd(which)
    Dim cmdStr
    If which = "sl" Then
        cmdStr = "qmake SearchLauncherQuickStep"
    ElseIf which = "st" Then
        cmdStr = "qmake MtkSettings"
    ElseIf which = "su" Then
        cmdStr = "qmake MtkSystemUI"
    ElseIf which = "ft" Then
        cmdStr = "qmake FactoryTest"
    ElseIf which = "fws" Then
        cmdStr = "mmm -j32 frameworks/base/services:services"
    ElseIf which = "fwr" Then
        cmdStr = "mmm -j32 frameworks/base/core/res"
    ElseIf which = "lot" Then
        cmdStr = "qmake GmsSampleIntegration"
    End If
    Call copyStrAndPasteInXshell(cmdStr)
End Sub

Function checkMvOut(outPath, folders)
    Dim cmdStr, folder, outFolder, parentFolder, tmpFolder
    For Each folder In folders
        If mBuild.Infos.isSdkT0() Then folder = "../" & folder
        If Not isFolderExists(folder) Then
            MsgBox(folder & " does not exist!")
            checkMvOut = ""
            Exit Function
        End If
    Next

    For Each folder In folders
        outFolder = outPath & "/" & folder
        If isFolderExists(outFolder) Then
            MsgBox(outFolder & " already exist!")
            checkMvOut = ""
            Exit Function
        Else
            tmpFolder = getParentPath(outFolder)
            If parentFolder <> tmpFolder Then
                parentFolder = tmpFolder
                IF Not isFolderExists(parentFolder) Then
                    cmdStr = cmdStr & "mkdir -p " & parentFolder & ";"
                End If
            End If
            cmdStr = cmdStr & "mv " & folder & " " & parentFolder & ";"
        End If
    Next

    if mBuild.Infos.isSdkT0() Then cmdStr = Replace(cmdStr, "../", "")
    checkMvOut = cmdStr
End Function

Function checkMvIn(outPath, folders, force)
    Dim cmdStr, folder, outFolder, parentFolder
    For Each folder In folders
        outFolder = outPath & "/" & folder
        If Not isFolderExists(outFolder) Then
            MsgBox(outFolder & " does not exist!")
            checkMvIn = ""
            Exit Function
        End If
    Next

    For Each folder In folders
        Dim checkFd : checkFd = folder
        If mBuild.Infos.isSdkT0() Then checkFd = "../" & folder
        If (Not force) And isFolderExists(checkFd) Then
            MsgBox(folder & " already exist!")
            checkMvIn = ""
            Exit Function
        Else
            outFolder = outPath & "/" & folder
            parentFolder = getParentPath(folder)
            if parentFolder = "" Then parentFolder = "./"
            cmdStr = cmdStr & "mv " & outFolder & " " & parentFolder & ";"
        End If
    Next

    if mBuild.Infos.isSdkT0() Then cmdStr = Replace(cmdStr, "../", "")
    checkMvIn = cmdStr
End Function

Function getCurrentTaskOutFolders()
    If mBuild.Infos.isSdkT0() Then
        If mBuild.Infos.is8168() Then
            getCurrentTaskOutFolders = Array("merged", "sys/out", "vnd/out")
        ElseIf mBuild.Infos.is8781() Then
            If mTask.Sys.Infos.isV0() Then
                getCurrentTaskOutFolders = Array("merged", "v_sys/out_sys", "vnd/out", "vnd/out_hal", "vnd/out_krn")
            ElseIf mTask.Sys.Infos.isU0() Then
                getCurrentTaskOutFolders = Array("merged", "u_sys/out_sys", "vnd/out", "vnd/out_hal", "vnd/out_krn")
            Else
                getCurrentTaskOutFolders = Array("merged", "sys/out", "vnd/out", "vnd/out_hal", "vnd/out_krn")
            End If
        Else
            If mTask.Sys.Infos.isV0() Then
                If mTask.Vnd.Infos.is8791() Then
                    getCurrentTaskOutFolders = Array("v_sys/merged", "v_sys/out_sys", "v_sys/out")
                Else
                    getCurrentTaskOutFolders = Array("v_sys/merged", "v_sys/out_sys", "v_sys/out", "v_sys/out_krn")
                End If
            ElseIf mTask.Sys.Infos.isU0() Then
                getCurrentTaskOutFolders = Array("merged", "u_sys/out", "vnd/out")
            ElseIf mTask.Sys.Infos.isT0() Then
                getCurrentTaskOutFolders = Array("merged", "sys/out", "vnd/out")
            Else
                getCurrentTaskOutFolders = Array("vnd/out")
            End If
        End If
    ElseIf mBuild.Infos.is8168() Then
        getCurrentTaskOutFolders = Array("out", "out_sys")
    Else
        getCurrentTaskOutFolders = Array("out")
    End If
End Function

Function findOutFoldersForMvOut()
    If mBuild.Infos.isSdkT0() Then
        'v+v
        If mTask.Sys.Infos.isV0() And isFolderExists("../v_sys/merged") And isFolderExists("../v_sys/out_sys") And isFolderExists("../v_sys/out") Then
            If isFolderExists("../v_sys/out_krn") Then
                findOutFoldersForMvOut = Array("v_sys/merged", "v_sys/out_sys", "v_sys/out", "v_sys/out_krn")
            Else
                findOutFoldersForMvOut = Array("v_sys/merged", "v_sys/out_sys", "v_sys/out")
            End If
        '8781
        ElseIf isFolderExists("../vnd/out_hal") Then
            '8781 s+v
            If mTask.Sys.Infos.isV0() And isFolderExists("../v_sys/out_sys") Then
                If isFolderExists("../v_sys/out") Then
                    MsgBox("There are two sys out folders: out/ out_sys/")
                    findOutFoldersForMvOut = Array("")
                Else
                    findOutFoldersForMvOut = Array("merged", "v_sys/out_sys", "vnd/out", "vnd/out_hal", "vnd/out_krn")
                End If
            '8781 s+u
            ElseIf mTask.Sys.Infos.isU0() And isFolderExists("../u_sys/out_sys") Then
                If isFolderExists("../u_sys/out") Then
                    MsgBox("There are two sys out folders: out/ out_sys/")
                    findOutFoldersForMvOut = Array("")
                Else
                    findOutFoldersForMvOut = Array("merged", "u_sys/out_sys", "vnd/out", "vnd/out_hal", "vnd/out_krn")
                End If
            '8781 s+t
            ElseIf isFolderExists("../sys/out_sys") Then
                findOutFoldersForMvOut = Array("merged", "sys/out_sys", "vnd/out", "vnd/out_hal", "vnd/out_krn")
            Else
                MsgBox("No sys out_sys!")
                findOutFoldersForMvOut = Array("")
            End If
        ElseIf isFolderExists("../vnd/out") Then
            If mTask.Sys.Infos.isU0() And isFolderExists("../u_sys/out") Then
                findOutFoldersForMvOut = Array("merged", "u_sys/out", "vnd/out")
            ElseIf isFolderExists("../sys/out") Then
                findOutFoldersForMvOut = Array("merged", "sys/out", "vnd/out")
            Else
                findOutFoldersForMvOut = Array("vnd/out")
            End If
        Else
            MsgBox("No vnd out!")
            findOutFoldersForMvOut = Array("")
        End If
    Else
        If isFolderExists("out_sys") And isFolderExists("out") Then
            findOutFoldersForMvOut = Array("out", "out_sys")
        ElseIf isFolderExists("out") Then
            findOutFoldersForMvOut = Array("out")
        Else
            MsgBox("No out!")
            findOutFoldersForMvOut = Array("")
        End If
    End If
End Function

Sub moveOutFoldersOut(taskNumInput)
    Dim taskNum
    If taskNumInput <> "" Then
        taskNum = taskNumInput
    Else
        Call setSysBuild()
        Dim innerId
        innerId = mBuild.Infos.getOutProp("ro.build.display.inner.id")
        If innerId = "" Then MsgBox("Empty inner id!") : Exit Sub
        If InStr(innerId, ".") = 0 Then MsgBox("Invalid inner id!" & VbLf & innerId) : Exit Sub
        Dim idArr
        idArr = Split(innerId, ".")
        If UBound(idArr) < 2 Then MsgBox("Invalid inner id!" & VbLf & innerId) : Exit Sub
        Dim numArr, i
        numArr = Split(idArr(2), "-")
        For i = UBound(numArr) To 0 Step -1
            If isTaskNum(numArr(i)) Then
                taskNum = numArr(i)
                Exit For
            End If
        Next
        If taskNum = "" Then MsgBox("Empty taskNum! ") : Exit Sub
    End If

    Dim buildType, outPath, outFolders, cmdStr
    If Not getTmpTaskWithNum(taskNum) Then Exit Sub
    buildType = mBuild.Infos.getOutProp("ro.build.type")
    If buildType = "userdebug" Then buildType = "debug"
    outPath = "../OUT/" & mTmpTask.Infos.TaskName & "_" & buildType
    If mTask.Infos.TaskNum = taskNum Then
        outFolders = getCurrentTaskOutFolders()
    Else
        outFolders = findOutFoldersForMvOut()
    End If
    If outFolders(0) <> "" Then
        cmdStr = checkMvOut(outPath, outFolders)
        If cmdStr <> "" And mTmpTask.Vnd.Infos.isV0() Then
            cmdStr = "cd ~" & relpaceSlashInPath(Split(mDrive, ":")(1)) & getParentPath(mTmpTask.Vnd.Sdk) & " && " & cmdStr
        End If
        Call copyStrAndPasteInXshell(cmdStr)
    End If
End Sub

Sub moveOutFoldersIn(buildType, force)
    Dim outPath, outFolders, cmdStr
    outPath = "../OUT/" & mTask.Infos.TaskName & "_" & buildType
    outFolders = getCurrentTaskOutFolders()
    cmdStr = checkMvIn(outPath, outFolders, force)
    If cmdStr <> "" And mTask.Vnd.Infos.isV0() Then
        cmdStr = "cd ~" & relpaceSlashInPath(Split(mDrive, ":")(1)) & getParentPath(mTask.Vnd.Sdk) & " && " & cmdStr
    End If
    Call copyStrAndPasteInXshell(cmdStr)
End Sub

Sub sendWeiXinMsg(who)
    'Call CopyOpenPathAllText()
    idTimer = window.setTimeout("Call appactivateWeiXin(""" & who & """)", 100, "VBScript")
End Sub

Sub appactivateWeiXin(who)
    window.clearTimeout(idTimer)
    Select Case who
        Case "hqp" : who = "huqipeng"
        Case "zhq" : who = "zhonghongqiang"
        Case "lqj" : who = "luoqingjun"
        Case "lyh" : who = "laiyuhui"
        Case "wj" : who = "weijuan"
    End Select

    Call oWs.appactivate("企业微信")
    idTimer = window.setTimeout("Call searchInWeiXin(""" & who & """)", 200, "VBScript")
End Sub

Sub searchInWeiXin(who)
    window.clearTimeout(idTimer)
    Call oWs.sendkeys("^f")
    idTimer = window.setTimeout("Call writeSearchStrInWeiXin(""" & who & """)", 200, "VBScript")
End Sub

Sub writeSearchStrInWeiXin(who)
    window.clearTimeout(idTimer)
    Call oWs.sendkeys(who & "xiangmuqun")
    'idTimer = window.setTimeout("Call enterSearchInWeiXin(""" & who & """)", 500, "VBScript")
End Sub

Sub enterSearchInWeiXin(who)
    window.clearTimeout(idTimer)
    Call oWs.sendkeys("{ENTER}")
    idTimer = window.setTimeout("Call writeWhoInWeiXin(""" & who & """)", 200, "VBScript")
End Sub

Sub writeWhoInWeiXin(who)
    window.clearTimeout(idTimer)
    Call oWs.sendkeys("@"&who&"")
    idTimer = window.setTimeout("Call enterWhoInWeiXin()", 300, "VBScript")
End Sub

Sub enterWhoInWeiXin()
    window.clearTimeout(idTimer)
    Call oWs.sendkeys("{ENTER}")
End Sub

Sub getCommitMsgList()
    Dim arr, i, listStr, splitWord
    splitWord = "]:"
    arr = Split(getOpenPath(), VbLf)
    For i = UBound(arr) To 0 Step -1
        if InStr(arr(i), "]") > 0 Then
            arr(i) = Replace(arr(i), "] : ", splitWord)
            arr(i) = Replace(arr(i), "] :", splitWord)
            arr(i) = Replace(arr(i), "]:", splitWord)
        End If
        If InStr(arr(i), splitWord) > 0 Then
            listStr = listStr & Right(arr(i), Len(arr(i)) - InStr(arr(i), splitWord) - Len(splitWord) + 1) & VbLf
        End If
    Next
    Call setOpenPath(listStr)
End Sub

Sub getReleaseNoteList()
    Dim arr, i, listStr, count
    count = 0
    arr = Split(getOpenPath(), VbLf)
    For i = 0 To UBound(arr)
        count = count + 1
        listStr = listStr & count & ". " & arr(i) & VbLf
    Next
    Call setOpenPath(listStr)
End Sub

Sub GetDiffCmdFromOverlayPath()
    Dim overlayPath, originPath
    overlayPath = getOpenPath()
    If InStr(overlayPath, "weibu/") = 1 Then
        originPath = getOriginPathFromOverlayPath(getOpenPath())
    Else
        originPath = overlayPath
        overlayPath = mBuild.Infos.getOverlayPath(originPath)
        If Not isFileExists(overlayPath) Then MsgBox("Overlay path not exist!" & VbLf & overlayPath) : Exit Sub
    End If
    Call copyStrAndPasteInXshell("git diff --no-index " & originPath & " " & overlayPath)
End Sub

Sub modBuildNumber(number)
    If InStr(mBuild.Sdk, "8168_s") Then
        Call cpFileAndSetValue(Array("bn2", number))
    Else
        Call cpFileAndSetValue(Array("bn", number))
    End If
End Sub

Function getModSystempropCmdStr(systempropPath, cmdStr, keyStr, valueStr)
    If strExistInFile(systempropPath, keyStr) Then
        cmdStr = cmdStr & "sed -i '/" & keyStr & "/s/.*/" & keyStr & "=" & valueStr & "/' " & systempropPath & ";"
    Else
        cmdStr = cmdStr & "sed -i '$a " & keyStr & "=" & valueStr & "' " & systempropPath & ";"
    End If
    getModSystempropCmdStr = cmdStr
End Function

Sub modSystemprop(whatArr)
    Dim systempropPath, cmdStr, keyStr, valueStr
    systempropPath = mBuild.Infos.ProjectPath & "/config/system.prop"

    valueStr = whatArr(1)
    If whatArr(0) = "tz" Then
        keyStr = "persist.sys.timezone"
        valueStr = Replace(valueStr, "/", "\/")
    ElseIf whatArr(0) = "loc" Then
        keyStr = "ro.product.locale" & "##" & "persist.sys.locale"
    ElseIf whatArr(0) = "ftd" Then
        keyStr = "ro.weibu.factorytest.disable_" & whatArr(1)
        valueStr = "1"
    Else
        Exit Sub
    End If

    If Not isFolderExists(mBuild.Infos.ProjectPath & "/config") Then
        cmdStr = cmdStr & "mkdir -p " & mBuild.Infos.ProjectPath & "/config" & ";"
    End If

    If Not isFileExists(systempropPath) Then
        cmdStr = cmdStr & "touch " & systempropPath & ";"
    End If

    If InStr(keyStr, "##") Then
        Dim i, ub, keyArr
        keyArr = Split(keyStr, "##")
        ub = UBound(keyArr)
        For i = 0 To ub
            cmdStr = getModSystempropCmdStr(systempropPath, cmdStr, keyArr(i), valueStr)
        Next
    Else
        cmdStr = getModSystempropCmdStr(systempropPath, cmdStr, keyStr, valueStr)
    End If

    cmdStr = cmdStr & "git diff " & systempropPath

    Call copyStrAndPasteInXshell(cmdStr)
End Sub

Sub modSettingProviderDefault(what)
    Dim filePath, cmdStr
    filePath = "vendor/mediatek/proprietary/packages/apps/SettingsProvider/res/values/defaults.xml"
    If Not isFileExists(mBuild.Infos.getOverlayPath(filePath)) Then
        cmdStr = "mkdir -p " & mBuild.Infos.getOverlayPath(getParentPath(filePath)) & ";"
        cmdStr = cmdStr & "cp " & filePath & " " & mBuild.Infos.getOverlayPath(getParentPath(filePath)) & ";"
    End If

    If what = "tz" Then
        cmdStr = getSedCmd(cmdStr, "def_auto_time_zone", "true", "false", mBuild.Infos.getOverlayPath(filePath))
    ElseIf what = "24" Then
        cmdStr = getSedCmd(cmdStr, "def_time_12_24", ">12<", ">24<", mBuild.Infos.getOverlayPath(filePath))
    ElseIf isNumeric(what) Then
        cmdStr = getSedCmd(cmdStr, "def_screen_off_timeout", "60000", what & "000", mBuild.Infos.getOverlayPath(filePath))
    ElseIf what = "rota" Then
        cmdStr = getSedCmd(cmdStr, "def_accelerometer_rotation", "false", "true", mBuild.Infos.getOverlayPath(filePath))
    ElseIf what = "bt" Then
        cmdStr = getSedCmd(cmdStr, "def_bluetooth_on", "false", "true", mBuild.Infos.getOverlayPath(filePath))
    ElseIf what = "wifi" Then
        cmdStr = getSedCmd(cmdStr, "def_wifi_on", "false", "true", mBuild.Infos.getOverlayPath(filePath))
    End If
    cmdStr = getGitDiffCmd(cmdStr, filePath)
    Call copyStrAndPasteInXshell(cmdStr)
End Sub

Sub showWorkInfo(taskNum)
    If isTaskNum(taskNum) Then
        Dim infos
        If Not getTmpTaskWithNum(taskNum) Then Exit Sub
        infos = mTmpTask.Infos.TaskNum & VbLf &_
                mTmpTask.Infos.TaskName & VbLf &_
                mTmpTask.Infos.CustomName & VbLf &_
                mTmpTask.Vnd.Sdk & VbLf &_
                mTmpTask.Vnd.Product & VbLf &_
                mTmpTask.Vnd.Project & VbLf &_
                mTmpTask.Sys.Sdk & VbLf &_
                mTmpTask.Sys.Product & VbLf &_
                mTmpTask.Sys.Project

        Call setOpenPath(infos)
    End If
End Sub

Sub getProjectInfosFromOpenPath()
    Dim Infos, inputArray
    inputArray = Split(getOpenPath(), VbLf)

    Infos = "[TaskNum]" & VbLf & "[TaskName]" & VbLf & "[CustomName]" & VbLf &_
            relpaceBackSlashInPath(Split(inputArray(0), "/weibu/")(0)) & VbLf &_
            Split(Split(inputArray(0), "/weibu/")(1), "/")(0) & VbLf &_
            Split(Split(inputArray(0), "/weibu/")(1), "/")(1) & VbLf &_
            relpaceBackSlashInPath(Split(inputArray(1), "/weibu/")(0)) & VbLf &_
            Split(Split(inputArray(1), "/weibu/")(1), "/")(0) & VbLf &_
            Split(Split(inputArray(1), "/weibu/")(1), "/")(1)
    Call setOpenPath(Infos)
End Sub

Sub setCommonTask(sdk)
    Dim taskInfos, vndBuild, sysBuild
    If sdk = "v0" Then
        Set taskInfos = (New TaskInfos)("0000", "COMMON_V0", "COMMON_V0")
        Set vndBuild = (New BaseBuild)("vnd", "mtk_sp_t0/v_sys", "tb8786p1_64_k66", "COMMON")
        Set sysBuild = (New BaseBuild)("sys", "mtk_sp_t0/v_sys", "mssi_64_cn", "COMMON")
        Set mTask = (New TaskBuild)(taskInfos, vndBuild, sysBuild)
        Call setCurrentBuild(mTask.Sys)
    ElseIf sdk = "u0" Then
        Set taskInfos = (New TaskInfos)("0000", "COMMON_U0", "COMMON_U0")
        Set vndBuild = (New BaseBuild)("vnd", "mtk_sp_t0/vnd", "tb8781p1_64", "M100TB_CS_625_WIFI")
        Set sysBuild = (New BaseBuild)("sys", "mtk_sp_t0/u_sys", "mssi_t_64_cn_armv82", "COMMON")
        Set mTask = (New TaskBuild)(taskInfos, vndBuild, sysBuild)
        Call setCurrentBuild(mTask.Sys)
    ElseIf sdk = "t0" Then
        Set taskInfos = (New TaskInfos)("0000", "COMMON_T0", "COMMON_T0")
        Set vndBuild = (New BaseBuild)("vnd", "mtk_sp_t0/vnd", "tb8781p1_64", "M100TB_CS_625_WIFI")
        Set sysBuild = (New BaseBuild)("sys", "mtk_sp_t0/sys", "mssi_t_64_cn_armv82", "COMMON")
        Set mTask = (New TaskBuild)(taskInfos, vndBuild, sysBuild)
        Call setCurrentBuild(mTask.Sys)
    ElseIf sdk = "s0" Then
        Set taskInfos = (New TaskInfos)("0000", "COMMON_S0", "COMMON_S0")
        Set vndBuild = (New BaseBuild)("vnd", "mtk_sp_t0/vnd", "tb8781p1_64", "M100TB_CS_625_WIFI")
        Set sysBuild = (New BaseBuild)("sys", "mtk_sp_t0/vnd", "tb8781p1_64", "M100TB_CS_625_WIFI")
        Set mTask = (New TaskBuild)(taskInfos, vndBuild, sysBuild)
        Call setCurrentBuild(mTask.Sys)
    End If
End Sub
