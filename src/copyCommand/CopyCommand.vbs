Option Explicit

Const ID_COMMAND_ENG = "command_eng"
Const ID_COMMAND_USERDEBUG = "command_userdebug"
Const ID_COMMAND_USER = "command_user"

Const ID_COMMAND_RM_OUT = "command_rm_out"
Const ID_COMMAND_RM_BUILDPROP = "command_rm_buildprop"
Const ID_COMMAND_BUILD_OTA = "command_build_ota"

Dim commandFinal

Sub CommandOfMake()
    Dim commandOta
    commandFinal = "make -j36 2>&1 | tee build.log"
    commandOta = "make -j36 otapackage 2>&1 | tee build_ota.log"

    If element_isChecked(ID_COMMAND_RM_OUT) Then
        commandFinal = "rm -rf out/ && " & commandFinal
    ElseIf element_isChecked(ID_COMMAND_RM_BUILDPROP) Then
        commandFinal = "find " & mIp.Infos.OutPath & " -type f -name build*.prop | xargs rm -v && " & commandFinal
    End If

    If element_isChecked(ID_COMMAND_BUILD_OTA) Then commandFinal = commandFinal & " && " & commandOta

    Call CopyString(commandFinal)
End Sub

Sub CommandOfLunch()
    If Not mIp.hasProjectInfos() Then Exit Sub
    Dim comboName, buildType

    Select Case True
        Case element_isChecked(ID_COMMAND_ENG)
            buildType = "eng"
        Case element_isChecked(ID_COMMAND_USERDEBUG)
            buildType = "userdebug"
        Case element_isChecked(ID_COMMAND_USER)
            buildType = "user"
    End Select

    If InStr(mIp.Infos.Sdk, "8168") > 0 Then
        commandFinal = getLuncnItemInSplitBuild(buildType)
        Call CopyQuoteString(commandFinal)
        Exit Sub
    Else
        comboName = "full_" & mIp.Infos.Product & "-" & buildType
        commandFinal = "source build/envsetup.sh ; lunch " & comboName & " " & mIp.Infos.Project
    End If
    Call CopyString(commandFinal)
End Sub

Function getLuncnItemInSplitBuild(buildType)
    Dim sysStr, vndStr, commandStr
    sysStr = "sys_" & mIp.Infos.SysTarget & "-" & buildType
    vndStr = "vnd_" & mIp.Infos.Product & "-" & buildType
    commandStr = sysStr & " " & vndStr & " " & mIp.Infos.Project
    commandStr = """lunch_item=""&Chr(34)&""" & commandStr & """&Chr(34)"

    Dim keyStr
    keyStr = "##Cusomer Settings"
    commandStr = """sed -i '/" & keyStr & "/i\""&" & commandStr & "&""' split_build.sh"""
    getLuncnItemInSplitBuild = commandStr
End Function

Sub CommandOfOut()
    If Not mIp.hasProjectInfos() Then Exit Sub
    Call CopyString(mIp.Infos.DownloadOutPath)
End Sub

Sub CopyCleanCommand()
	commandFinal = "git checkout .;git clean -df"
	Call CopyString(commandFinal)
End Sub

Sub CopyCommitInfo()
    If Not mIp.hasProjectInfos() Then Exit Sub
	commandFinal = "[" & mIp.Infos.Project & "] : "
	Call CopyString(commandFinal)
End Sub

Sub CopyBuildOtaUpdate()
    If InStr(mIp.Infos.Sdk, "_s") > 0 Then
        commandFinal = "./out/host/linux-x86/bin/ota_from_target_files -i old.zip new.zip update.zip"
    Else
        commandFinal = "./build/tools/releasetools/ota_from_target_files -i old.zip new.zip update.zip"
    End If
    Call CopyString(commandFinal)
End Sub

Sub MkdirWeibuFolderPath()
    If Not mIp.hasProjectInfos() Then Exit Sub
    Dim path : path = getOpenPath()
    Dim wholePath : wholePath = mIp.Infos.Sdk & "/" & path
    Dim folderPartPath, folderWholePath
    Dim mkdirCmd, cpCmd
    commandFinal = ""

    If oFso.FileExists(wholePath) Or oFso.FolderExists(wholePath) Then
        Dim index : index = InStrRev(path, "/")
        folderPartPath = Left(path, index)
        folderWholePath = mIp.Infos.Sdk & "/" & folderPartPath

        If oFso.FolderExists(folderWholePath) Then

            If Not oFso.FolderExists(mIp.Infos.getOverlaySdkPath(folderPartPath)) Then
                mkdirCmd = "mkdir -p " & mIp.Infos.getOverlayPath(folderPartPath) & ";"
            End If

            commandFinal = mkdirCmd
        End If
    End If

    If oFso.FileExists(wholePath) Then
        If Not oFso.FileExists(mIp.Infos.getOverlaySdkPath(path)) Then
            cpCmd = "cp " & path & " " & mIp.Infos.getOverlayPath(folderPartPath)
        Else
            MsgBox("File exist!")
        End If
            commandFinal = mkdirCmd & cpCmd
    End If

    commandFinal = Replace(commandFinal, "\", "/")
    Call CopyString(commandFinal)
End Sub

Sub copyExportToolsPathCmd()
    commandFinal = "export PATH=$HOME/Tools:$PATH"
    Call CopyString(commandFinal)
End Sub
