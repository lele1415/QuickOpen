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

'Function findStr(str, key)
'    Dim i, num
'    num = 0
'    For i = 1 To (Len(str) - Len(key) + 1)
'        If Mid(str, i, Len(key)) = key Then
'            num = num + 1
'        End If
'    Next
'    findStr = num
'End Function
'
'Function getDriverProjectName(mmiFolderName)
'    Dim str : str = mmiFolderName
'
'    'M863Y_YUKE_066-MMI
'    If InStr(str, "-MMI") > 0 And findStr(str, "-") = 1 Then
'        str = Replace(str, "-MMI", "")
'
'    'm863ur200_64-SBYH_A8005A-Nitro_8_MMI
'    ElseIf findStr(str, "-") > 1 Then
'        Do Until findStr(str, "-") = 1
'            str = Left(str, InStrRev(str, "-") - 1)
'        Loop
'    Else
'        str = ""
'    End If
'    getDriverProjectName = str
'End Function
'
''MsgBox(getDriverProjectName("m863ur200_64-SBYH_A8009_MasstelTab8Edu-apk-MMI"))
'Dim value, mTaskNum
'value="http://192.168.0.29:3000/zentao/task-view-181.html"
'mTaskNum = Replace(Right(value, Len(value) - InStr(value, "task-view-") - Len("task-view-") + 1), ".html", "")
'MsgBox(mTaskNum)

'MsgBox(InStr("123/", "/"))
'MsgBox(Len("123/"))

'If True Then Execute("If False Then MsgBox(""1"") : Else MsgBox(""2"")") : Else MsgBox("3")
'MsgBox(Split("a-b", "-")(1))
'MsgBox(Right("12345367", Len("12345367") - InStrRev("12345367", "3")))

'Function getParentPath(path)
'    Dim str, index
'    str = path
'    index = InStrRev(str, "/")
'    If index > 0 And index < Len(str) Then
'        str = Left(str, index)
'    End If
'    getParentPath = str
'End Function
'
'MsgBox(getParentPath("123/456/789"))

'MsgBox(InStr("123", ""))

'MsgBox(Replace("1" & VbLf & "2" & VbLf & "3", VbLf, " | "))

Class test
    Public Property Get aaa : aaa = "aaa" : End Property
    Public Function b() : b = "b" : End Function
    Public cc
End Class

Dim t : Set t = New test
t.cc = "cc"
MsgBox t.cc