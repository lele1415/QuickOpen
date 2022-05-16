
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


str = Replace("Z:\work05\mt876w6_r\mt8766_r\alps1", "/", "\")
str = Left(str, InStrRev(str, "\alps") - 1)
str = Replace(str, Left(str, InStrRev(str, "\")), "")
MsgBox(str)
