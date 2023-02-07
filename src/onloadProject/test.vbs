Option Explicit

'Dim str, resultStr
'projectName = Array("m863ur200_64-SBYH_A8005A_Call-Nitro_8_MMI", _
'	                "m863ur200_64-SBYH_A8005A-Nitro_8_MMI", _
'	                "m863ur200_64-RK_029-J013_Masstel_Tab_83_MMI", _
'	                "m863ur200_64-JR_039_Call-T8Plus_MMI", _
'	                "M862P_JR_065-MMI", _
'	                "m863ur200_64-SBYH_A8005A-Nitro_8_MX_N81WW_MMI", _
'	                "m863ur200_64_asb-XHB_072-MMI", _
'	                "M863Y_RK_104-MMI", _
'	                "m863ur200_64-RK_108-J013_Masstel_Tab_83_MMI", _
'	                "M866Y_WZX_101-MMI", _
'	                "M863Y_YUKE_066-MMI")
'For i = 0 To UBound(projectName)
'	str = Replace(projectName(i), "-MMI", "")
'	str = Replace(str, "_MMI", "")
'	If InStr(str, "-") > 0 Then
'	    str = Replace(str, Left(str, InStr(str, "-")), "")
'	Else
'	    str = Replace(str, Left(str, InStr(str, "_")), "")
'	End If
'	resultStr = resultStr & VbLf & str
'Next
'MsgBox(resultStr)


'str = Replace("Z:\work05\mt876w6_r\mt8766_r\alps1", "/", "\")
'str = Left(str, InStrRev(str, "\alps") - 1)
'str = Replace(str, Left(str, InStrRev(str, "\")), "")
'MsgBox(str)

'path = "weibu/tb8766p1_64_bsp/M863Y_YUKE_066/config"
'str = relpaceSlashInPath(path)
'If InStr(str, "/") > 0 Then
'    str = Replace(str, Left(str, InStrRev(str, "/")), "")
'Else
'    str = path
'End If
'MsgBox(str)

Function findStr(str, key)
    Dim i, num
    num = 0
    For i = 1 To (Len(str) - Len(key) + 1)
        If Mid(str, i, Len(key)) = key Then
            num = num + 1
        End If
    Next
    findStr = num
End Function

Function getDriverProjectName(mmiFolderName)
    Dim str : str = mmiFolderName

    'M863Y_YUKE_066-MMI
    If InStr(str, "-MMI") > 0 And findStr(str, "-") = 1 Then
        str = Replace(str, "-MMI", "")

    'm863ur200_64-SBYH_A8005A-Nitro_8_MMI
    ElseIf findStr(str, "-") > 1 Then
        Do Until findStr(str, "-") = 1
            str = Left(str, InStrRev(str, "-") - 1)
        Loop
    Else
        str = ""
    End If
    getDriverProjectName = str
End Function

'MsgBox(getDriverProjectName("m863ur200_64-SBYH_A8009_MasstelTab8Edu-apk-MMI"))
Dim value, mTaskNum
value="http://192.168.0.29:3000/zentao/task-view-181.html"
mTaskNum = Replace(Right(value, Len(value) - InStr(value, "task-view-") - Len("task-view-") + 1), ".html", "")
MsgBox(mTaskNum)
