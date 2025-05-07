Option Explicit

Const DRIVE_WORK_1 = "X:\work1\"
Const DRIVE_WORK_2 = "X:\work2\"
Const DRIVE_WORK_06 = "Z:\work06\"
Dim mDrive : mDrive = DRIVE_WORK_1
Dim mBuild : Set mBuild = New BaseBuild
Dim mTask : Set mTask = New TaskBuild
Dim mTaskList : Set mTaskList = New VariableArray
Dim mLastTaskNum

Sub updateCurrentTaskTitle()
    Dim buildType
    If mBuild.Infos.isVnd() Then
        buildType = "VND"
    Else
        buildType = "SYS"
    End If
    Call updateTitle(mTask.Infos.TaskName & " | " & buildType & " | " & mDrive  & " | " & relpaceBackSlashInPath(mBuild.Sdk))
End Sub

Sub setCurrentDrive(drive)
    If (drive = "x1") Then
        mDrive = DRIVE_WORK_1
    ElseIf (drive = "x2") Then
        mDrive = DRIVE_WORK_2
    ElseIf (drive = "z6") Then
        mDrive = DRIVE_WORK_06
    End If
    Call updateCurrentTaskTitle()
End Sub

Sub findDrive()
    Dim sdkP, userOut, debugOut
    sdkP = getParentPath(mBuild.Sdk)
    userOut = mTask.Infos.TaskName & "_user"
    debugOut = mTask.Infos.TaskName & "_debug"
    If isFolderExists(mBuild.Infos.Out) And isInStr(mBuild.Infos.getOutProp("ro.build.display.inner.id"), Replace(mTask.Vnd.Project, "_", ".")) Then
        Exit Sub
    Else
        If mDrive = DRIVE_WORK_1 Then
            Call setCurrentDrive("x2")
        Else
            Call setCurrentDrive("x1")
        End If
        If isFolderExists(mBuild.Infos.Out) And isInStr(mBuild.Infos.getOutProp("ro.build.display.inner.id"), Replace(mTask.Vnd.Project, "_", ".")) Then Exit Sub
    End If

    If oFso.FolderExists(DRIVE_WORK_1 & sdkP & "\OUT\" & userOut) Or oFso.FolderExists(DRIVE_WORK_1 & sdkP & "\OUT\" & debugOut) Then
        Call setCurrentDrive("x1")
        Exit Sub
    ElseIf oFso.FolderExists(DRIVE_WORK_2 & sdkP & "\OUT\" & userOut) Or oFso.FolderExists(DRIVE_WORK_2 & sdkP & "\OUT\" & debugOut) Then
        Call setCurrentDrive("x2")
        Exit Sub
    End If

    If isFolderExists(mBuild.Infos.ProjectPath) Then
        Exit Sub
    Else
        If mDrive = DRIVE_WORK_1 Then
            Call setCurrentDrive("x2")
        Else
            Call setCurrentDrive("x1")
        End If
        If isFolderExists(mBuild.Infos.ProjectPath) Then Exit Sub
    End If
End Sub

Sub setCurrentBuild(build)
    Set mBuild = build
    Call updateCurrentTaskTitle()
End Sub

Sub setCurrentTask(task)
    Set mTask = task
    Call setCurrentBuild(mTask.Sys)
    Call findDrive()
    If mTask.Infos.TaskNum <> mLastTaskNum Then
        Call saveLastTaskNum()
        mLastTaskNum = mTask.Infos.TaskNum
    End If
End Sub

Sub setVndBuild()
    Set mBuild = mTask.Vnd
    Call updateCurrentTaskTitle()
End Sub

Sub setSysBuild()
    Set mBuild = mTask.Sys
    Call updateCurrentTaskTitle()
End Sub

Sub getTaskList()
    If Not isFileExists(PATH_TASK_LIST) Then Exit Sub
    
    Dim oText, sReadLine
    Set oText = oFso.OpenTextFile(PATH_TASK_LIST, FOR_READING, False, True)

    Do Until oText.AtEndOfStream
        sReadLine = oText.ReadLine
        If InStr(sReadLine, " | ") Then
            Call handleForWorksInfo(sReadLine)
        End If
    Loop

    oText.Close
    Set oText = Nothing
End Sub

Sub handleForWorksInfo(sReadLine)
    Dim arr, taskInfos, vndBuild, sysBuild, taskBuild
    arr = Split(sReadLine, " | ")
    If UBound(arr) < 8 Or Not isTaskNum(trimStr(arr(0))) Then MsgBox("Invalid task infos!" & VbLf & sReadLine) : Exit Sub

    'taskNum, taskName, customName
    Set taskInfos = (New TaskInfos)(trimStr(arr(0)), trimStr(arr(1)), trimStr(arr(2)))
    'type, sdk, product, project
    Set vndBuild = (New BaseBuild)("vnd", trimStr(arr(3)), trimStr(arr(4)), trimStr(arr(5)))
    'type, sdk, product, project
    Set sysBuild = (New BaseBuild)("sys", trimStr(arr(6)), trimStr(arr(7)), trimStr(arr(8)))

    Set taskBuild = (New TaskBuild)(taskInfos, vndBuild, sysBuild)

    Call mTaskList.append(taskBuild)
End Sub

Sub loadTaskWithNum(taskNum)
    Dim i
    For i = mTaskList.Bound To 0 Step -1
        If mTaskList.v(i).Infos.TaskNum = taskNum Then
            Call setCurrentTask(mTaskList.v(i))
            Exit For
        End If
    Next
End Sub

Function getLastTaskNum()
    If Not isFileExists(PATH_LAST_TASK) Then
        getLastTaskNum = ""
        Exit Function
    End If
    
    Dim oText, sReadLine
    Set oText = oFso.OpenTextFile(PATH_LAST_TASK, FOR_READING, False, True)

    Do Until oText.AtEndOfStream
        sReadLine = oText.ReadLine
        If isTaskNum(sReadLine) Then
            Exit Do
        End If
    Loop

    oText.Close
    Set oText = Nothing

    getLastTaskNum = sReadLine
End Function

Sub saveLastTaskNum()
    Call initTxtFile(PATH_LAST_TASK)
    Dim oTxt
    Set oTxt = oFso.OpenTextFile(PATH_LAST_TASK, FOR_APPENDING, False, True)
    oTxt.WriteLine(mTask.Infos.TaskNum)
    oTxt.Close
    Set oTxt = Nothing
End Sub

Sub loadLastTask()
    Dim task, taskNum
    taskNum = getLastTaskNum()
    If isTaskNum(taskNum) Then
        mLastTaskNum = taskNum
        Call loadTaskWithNum(taskNum)
    End If
End Sub

Sub removeTask(taskNum)
    Dim i, task, seq
    seq = -1
    For i = mTaskList.Bound To 0 Step -1
        Set task = mTaskList.v(i)
        If task.Infos.TaskNum = taskNum Then
            seq = i
            Exit For
        End If
    Next
    If seq > -1 Then
        Call mTaskList.popBySeq(seq)
    End If
    Call updateTaskList()
End Sub

Sub updateTaskList()
    Call initTxtFile(PATH_TASK_LIST)
    Dim oTxt, i, task, sp
    Set oTxt = oFso.OpenTextFile(PATH_TASK_LIST, FOR_APPENDING, False, True)
    sp = " | "

    For i = 0 To mTaskList.Bound
        Set task = mTaskList.v(i)
        oTxt.WriteLine(task.Infos.toString(sp) & sp & task.Vnd.toString(sp) & sp & task.Sys.toString(sp))
    Next

    oTxt.Close
    Set oTxt = Nothing
End Sub

Dim mTmpTask
Function getTmpTaskWithNum(taskNum)
    If Not isTaskNum(taskNum) Then getTmpTaskWithNum = False : Exit Function
    If taskNum = mTask.Infos.TaskNum Then
        Set mTmpTask = mTask
        getTmpTaskWithNum = True
        Exit Function
    End If

    Dim task, i, result
    result = False
    For i = mTaskList.Bound To 0 Step -1
        Set task = mTaskList.v(i)
        If task.Infos.TaskNum = taskNum Then
            Set mTmpTask = task
            result = True
            Exit For
        End If
    Next
    getTmpTaskWithNum = result
End Function
