Const ID_COMMAND_ENG = "command_eng"
Const ID_COMMAND_USERDEBUG = "command_userdebug"
Const ID_COMMAND_USER = "command_user"

Const ID_COMMAND_RM_OUT = "command_rm_out"
Const ID_COMMAND_BUILD_OTA = "command_build_ota"

Sub CommandOfMake()
    Dim commandFinal, commandOta
    commandFinal = "make -j36 2>&1 | tee build.log"
    commandOta = "make -j36 otapackage 2>&1 | tee build_ota.log"

    If element_isChecked(ID_COMMAND_RM_OUT) Then commandFinal = "rm -rf out/ && " & commandFinal
    If element_isChecked(ID_COMMAND_BUILD_OTA) Then commandFinal = commandFinal & " && " & commandOta

    Call CopyString(commandFinal)
End Sub

Sub CommandOfLunch()
    Dim targetProduct, customProject, comboName, buildType, commandFinal

    targetProduct = getProduct()
    customProject = getProject()

    Select Case True
        Case element_isChecked(ID_COMMAND_ENG)
            buildType = "eng"
        Case element_isChecked(ID_COMMAND_USERDEBUG)
            buildType = "userdebug"
        Case element_isChecked(ID_COMMAND_USER)
            buildType = "user"
    End Select

    comboName = "full_" & targetProduct & "-" & buildType

    commandFinal = "source build/envsetup.sh ; lunch " & comboName & " " & customProject
    Call CopyString(commandFinal)
End Sub

Sub CommandOfOut()
    Call CopyString(getOutPath())
End Sub

Sub CopyCleanCommand()
	commandFinal = "git checkout .;git clean -df"
	Call CopyString(commandFinal)
End Sub

Sub CopyCommitInfo()
	commandFinal = "[" & getProject() & "] : "
	Call CopyString(commandFinal)
End Sub

Sub CopyWeibuFolderPath()
    commandFinal = getProjectPathWithoutSdk()
    Call CopyString(commandFinal)
    Call setOpenPath(commandFinal)
End Sub

Sub CopyWeibuDriverFolderPath()
    commandFinal = getProjectPathWithoutSdk()
    If (InStr(commandFinal, "-MMI") > 0) Then
        commandFinal = Replace(commandFinal, "-MMI", "")
    End If
    Call CopyString(commandFinal)
    Call setOpenPath(commandFinal)
End Sub

Sub AddWeibuFolderPath()
    commandFinal = getProjectPathWithoutSdk() & "/" & getOpenPath()
    Call setOpenPath(commandFinal)
End Sub

Sub DelWeibuFolderPath()
    commandFinal = Replace(getOpenPath(), getProjectPathWithoutSdk(), "")
    If InStr(commandFinal, "/") = 1 Then
        commandFinal = Replace(commandFinal, "/", "", 1, 1)
    End If
    Call setOpenPath(commandFinal)
End Sub

Sub MkdirWeibuFolderPath()
    Dim filePartPath : filePartPath = Replace(getOpenPath(), "\", "/")
    Dim fileWholePath : fileWholePath = Replace(getSdkPath(), "\", "/") & "/" & filePartPath
    Dim folderPartPath, folderWholePath
    Dim mkdirCmd, cpCmd
    commandFinal = ""

    If oFso.FileExists(fileWholePath) Then
        Dim index : index = InStrRev(filePartPath, "/")
        folderPartPath = Left(filePartPath, index)
        folderWholePath = getSdkPath() & "/" & folderPartPath

        If oFso.FolderExists(folderWholePath) Then

            If Not oFso.FolderExists(getProjectPath() & "/" & folderPartPath) Then
                mkdirCmd = "mkdir -p " & getProjectPathWithoutSdk() & "/" & folderPartPath & ";"
            End If

            'If Not oFso.FileExists(getProjectPath() & "/" & filePartPath) Then
                cpCmd = "cp " & filePartPath & " " & getProjectPathWithoutSdk() & "/" & folderPartPath
            'End If

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