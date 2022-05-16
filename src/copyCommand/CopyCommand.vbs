Const ID_COMMAND_ENG = "command_eng"
Const ID_COMMAND_USERDEBUG = "command_userdebug"
Const ID_COMMAND_USER = "command_user"

Const ID_COMMAND_RM_OUT = "command_rm_out"
Const ID_COMMAND_BUILD_OTA = "command_build_ota"

Dim commandFinal

Sub CommandOfMake()
    Dim commandOta
    commandFinal = "make -j36 2>&1 | tee build.log"
    commandOta = "make -j36 otapackage 2>&1 | tee build_ota.log"

    If element_isChecked(ID_COMMAND_RM_OUT) Then commandFinal = "rm -rf out/ && " & commandFinal
    If element_isChecked(ID_COMMAND_BUILD_OTA) Then commandFinal = commandFinal & " && " & commandOta

    Call CopyString(commandFinal)
End Sub

Sub CommandOfLunch()
    Dim comboName, buildType

    Select Case True
        Case element_isChecked(ID_COMMAND_ENG)
            buildType = "eng"
        Case element_isChecked(ID_COMMAND_USERDEBUG)
            buildType = "userdebug"
        Case element_isChecked(ID_COMMAND_USER)
            buildType = "user"
    End Select

    comboName = "full_" & mIp.Infos.Product & "-" & buildType

    commandFinal = "source build/envsetup.sh ; lunch " & comboName & " " & mIp.Infos.Project
    Call CopyString(commandFinal)
End Sub

Sub CommandOfOut()
    Call CopyString(mIp.Infos.OutSdkPath)
End Sub

Sub CopyCleanCommand()
	commandFinal = "git checkout .;git clean -df"
	Call CopyString(commandFinal)
End Sub

Sub CopyCommitInfo()
	commandFinal = "[" & mIp.Infos.Project & "] : "
	Call CopyString(commandFinal)
End Sub

Sub CopyWeibuFolderPath()
    commandFinal = mIp.Infos.ProjectPath
    Call CopyString(commandFinal)
    'Call setOpenPath(commandFinal)
End Sub

Sub CopyWeibuDriverFolderPath()
    commandFinal = mIp.Infos.ProjectPath
    If (InStr(commandFinal, "-MMI") > 0) Then
        commandFinal = Replace(commandFinal, "-MMI", "")
    End If
    Call CopyString(commandFinal)
    'Call setOpenPath(commandFinal)
End Sub

Sub AddWeibuFolderPath()
    commandFinal = mIp.Infos.ProjectPath & "/" & getOpenPath()
    Call setOpenPath(commandFinal)
End Sub

Sub DelWeibuFolderPath()
    commandFinal = Replace(getOpenPath(), mIp.Infos.ProjectPath, "")
    If InStr(commandFinal, "/") = 1 Then
        commandFinal = Replace(commandFinal, "/", "", 1, 1)
    End If
    Call setOpenPath(commandFinal)
End Sub

Sub MkdirWeibuFolderPath()
    Dim filePartPath : filePartPath = Replace(getOpenPath(), "\", "/")
    Dim fileWholePath : fileWholePath = Replace(mIp.Infos.Sdk, "\", "/") & "/" & filePartPath
    Dim folderPartPath, folderWholePath
    Dim mkdirCmd, cpCmd
    commandFinal = ""

    If oFso.FileExists(fileWholePath) Then
        Dim index : index = InStrRev(filePartPath, "/")
        folderPartPath = Left(filePartPath, index)
        folderWholePath = mIp.Infos.Sdk & "/" & folderPartPath

        If oFso.FolderExists(folderWholePath) Then

            If Not oFso.FolderExists(mIp.Infos.ProjectSdkPath & "/" & folderPartPath) Then
                mkdirCmd = "mkdir -p " & mIp.Infos.ProjectPath & "/" & folderPartPath & ";"
            End If

            If Not oFso.FileExists(mIp.Infos.ProjectSdkPath & "/" & filePartPath) Then
                cpCmd = "cp " & filePartPath & " " & mIp.Infos.ProjectPath & "/" & folderPartPath
            Else
                MsgBox("File exist!")
            End If

            commandFinal = mkdirCmd & cpCmd
            'If commandFinal = "" Then
            '    MsgBox("File exist!")
            'End If
        End If
    End If
    commandFinal = Replace(commandFinal, "\", "/")
    Call CopyString(commandFinal)
End Sub

Sub CopyString(str)
    Clipboard="MsHta vbscript:ClipBoardData.setData(""Text"","""&str&""")(Window.Close)"  
    oWs.Run(Clipboard)
End Sub