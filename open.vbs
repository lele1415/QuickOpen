Dim width,height
width=CreateObject("HtmlFile").ParentWindow.Screen.AvailWidth
Window.MoveTo width-500,100
Window.ResizeTo 500,800

Set ws=CreateObject("wscript.shell")
Set Fso=CreateObject("Scripting.FileSystemObject")

Dim textEditorPath, codePathTxt
textEditorPath = "F:\tools\Sublime_Text_3\sublime_text.exe"
codePathTxt_KK = "codePath_KK.txt"
codePathTxt_L1 = "codePath_L1.txt"
codePathListId_KK = "codePath_KK"
codePathListId_L1 = "codePath_L1"

ReadCodePath codePathTxt_KK, codePathListId_KK
ReadCodePath codePathTxt_L1, codePathListId_L1

ReadHistory "Dict_PN","project_name","PN_ul_id","KK"
ReadHistory "Dict_OpenFile","FilePath_id","Openfile_ul_id","L1"
ReadHistory "Dict_OpenFile","FilePath_l1_id","Openfile_ul_l1_id","L1"
'onloadL1Project1()

folder_deep_count = 1

'Function CopyString(s) 
'    Dim forms, textbox
'
'    Set forms=CreateObject("forms.form.1") 
'    Set textbox=forms.Controls.Add("forms.textbox.1").Object 
'    With textbox 
'        .multiline=True 
'        .text=s 
'        .selstart=0 
'        .sellength=Len(.text) 
'        .copy 
'    End With 
'End Function

Function ReadCodePath(codePathTxtName, codePathListId)
    Dim txtPath, oTxt, codePathCount, sTmp, continue
    txtPath = ws.CurrentDirectory & "\" & codePathTxtName
    codePathCount = 0
    continue = True
    If Fso.FileExists(txtPath) Then
        Set oTxt = fso.OpenTextFile(txtPath,1)
    Else
        continue = False
    End If

    If continue Then
        Do Until oTxt.AtEndOfStream
            sTmp = oTxt.ReadLine
            If sTmp <> "" Then
                addOption codePathListId, sTmp
            End If
        Loop
    End If
End Function

Function CopyString(str)
    Clipboard="MsHta vbscript:ClipBoardData.setData(""Text"","""&str&""")(Window.Close)"  
    ws.Run(Clipboard)
End Function

Function getPath(Name)
    Dim PN, SC, LK
    PN=element_getValue("project_name")
    SC=element_getValue("codePath_KK")
    Select Case Name
        Case "alps"
            LK = SC
        Case "binary"
            LK = SC&"\mediatek\binary\packages\"&PN
        Case "config"
            LK = SC&"\mediatek\config\"&PN
        Case "config_only"
            LK = SC&"\mediatek\config\"
        Case "custom"
            LK = SC&"\mediatek\custom\"&PN
        Case "overlay"
            PN = Replace(PN,"[","-")
            PN = Replace(PN,"]","")
            LK = SC&"\mediatek\custom\common\resource_overlay\roco\resandroid\"&PN
        Case "out"
            LK = SC&"\out\target\product"
        Case "ProjectConfig.mk"
            LK = SC&"\mediatek\config\"&PN&"\ProjectConfig.mk"
        Case "system.prop"
            LK = SC&"\mediatek\config\"&PN&"\system.prop"
        Case "product_roco.mk"
            LK = SC&"\mediatek\config\"&PN&"\product_roco.mk"
        Case "codePath_KK"
            LK = SC
    End Select

    getPath = LK

End Function

Sub openProjectFile(Name)
    Dim LK
    LK=getPath(Name)
    If Fso.FileExists(LK) Then
        ws.run textEditorPath&" "&LK
    Else
        msgbox(Name&"不存在")
    End If
End Sub

Sub openProjectFolder(Name)
    Dim LK
    LK=getPath(Name)

    Select Case Name
        Case "overlay"
            If (Not Fso.FolderExists(LK)) Then
                LK = SC&"\mediatek\custom\common\resource_overlay\roco\reslight\"&PN 
            End If
        Case "out"
            If Fso.FolderExists(LK) Then
                Dim MyFolder_PN, Folders, FDN, PN1
                Set MyFolder_PN = Fso.GetFolder(LK)
                Set Folders = MyFolder_PN.SubFolders
                For Each Folder in Folders
                    FDN=Folder.Name
                    If Fso.FolderExists(LK&"\"&FDN&"\system") Then
                        PN1=Folder.Name
                        Exit For 
                    End If
                Next
                LK=LK&"\"&PN1
            End If
    End Select

    If Fso.FolderExists(LK) Then
        ws.run "explorer.exe "&LK
    Else
        msgbox(Name&"不存在")
    End If
End Sub

Function OpenMore(ParentFolderName,path,FolderDeepCount)
    Dim ButtonOfMoreId, DivId, ArrayFolderName, ArrayFolderNameSorted, ArrayFolderNameCount, ArrayFileName, ArrayFileNameSorted, ArrayFileNameCount
    ReDim ArrayFolderName(0)
    ReDim ArrayFileName(0)
    ArrayFolderNameCount = 0
    ArrayFileNameCount = 0

    ButtonOfMoreId = ParentFolderName&"_More_id"
    DivId = ParentFolderName&"MoreDiv_id"
    If (element_getValue(ButtonOfMoreId)="-") Then
        node_removeNode(DivId)
        input_changeInfo ButtonOfMoreId, "+", 0
    Else
        If Fso.FolderExists(path) Then
            addDiv ParentFolderName,DivId

            addButtonOfFolderNameOfOpenMore ParentFolderName,path,FolderDeepCount
            addButtonOfFileNameOfOpenMore ParentFolderName,path,FolderDeepCount

            input_changeInfo ButtonOfMoreId, "-", 0
        End If
    End If

End Function

Function addButtonOfFolderNameOfOpenMore(ParentFolderName,path,FolderDeepCount)
    Dim fso, oFolder, oSubFolders
    Set fso=CreateObject("Scripting.FileSystemObject")
    Set oFolder = fso.GetFolder(path&"\")
    Set oSubFolders = oFolder.SubFolders

    Dim ArrayFolderName, ArrayFolderNameSorted, ArrayFolderNameCount
    ReDim ArrayFolderName(0)
    ArrayFolderNameCount = 0

    For Each Folder in oSubFolders
        ReDim Preserve ArrayFolderName(ArrayFolderNameCount)
        ArrayFolderName(ArrayFolderNameCount) = Folder.Name
        ArrayFolderNameCount = ArrayFolderNameCount + 1
    Next

    If ArrayFolderNameCount > 0 Then
        ArrayFolderNameSorted = sortArray(ArrayFolderName)
        Dim i
        For i=0 to UBound(ArrayFolderNameSorted)
            addButtonOfFolderName path+"\",ParentFolderName,ArrayFolderNameSorted(i),FolderDeepCount,1
        Next
    End If
End Function

Function addButtonOfFileNameOfOpenMore(ParentFolderName,path,FolderDeepCount)
    Dim fso, oFolder, oSubFiles
    Set fso=CreateObject("Scripting.FileSystemObject")
    Set oFolder = fso.GetFolder(path&"\")
    Set oSubFiles = oFolder.Files

    Dim ArrayFileName, ArrayFileNameSorted, ArrayFileNameCount
    ReDim ArrayFileName(0)
    ArrayFileNameCount = 0

    For Each File in oSubFiles
        ReDim Preserve ArrayFileName(ArrayFileNameCount)
        ArrayFileName(ArrayFileNameCount) = File.Name
        ArrayFileNameCount = ArrayFileNameCount + 1
        ''addButtonOfFolderName LK+"\",ParentFolderName,File.Name,FolderDeepCount+1,0
    Next

    If ArrayFileNameCount > 0 Then
        ArrayFileNameSorted = sortArray(ArrayFileName)
        Dim i
        For i=0 to UBound(ArrayFileNameSorted)
            addButtonOfFolderName path+"\",ParentFolderName,ArrayFileNameSorted(i),FolderDeepCount,0
        Next
    End If
End Function



Function removeMoreButton(ButtonName)
    If (element_getValue(ButtonName&"_More_id")="-") Then
        OpenMore ButtonName,getPath(ButtonName),1
    End If
End Function

Function removeAllMoreButton(sVersion)
    If sVersion = "KK" Then
        removeMoreButton("binary")
        removeMoreButton("config")
        removeMoreButton("custom")
        removeMoreButton("overlay")
        removeMoreButton("out")
    ElseIf sVersion = "L1" Then
        removeMoreButton("alps_L1")
        removeMoreButton("device1_L1")
        removeMoreButton("device2_L1")
        removeMoreButton("custom_L1")
        removeMoreButton("out_L1")
    End If
End Function

Function runApplication(path)
    ws.run path
End Function

Function OpenByPath(path)
    If Fso.FolderExists(path) Then
        ws.run "explorer.exe "&path
    ElseIf Fso.FileExists(path) Then
        ws.run "F:\tools\Sublime_Text_3\sublime_text.exe "&path
    Else
        msgbox("路径不存在")
    End If
End Function

Function OpenFile(AndroidVersion)
    If AndroidVersion = "KK" Then
    	SC=element_getValue("codePath_KK") 
        FP=element_getValue("FilePath_id")
    Else
        SC=element_getValue("codePath_L1")
        FP=element_getValue("FilePath_l1_id")
    End If

    If Len(FP) > 0 Then
        FP=Replace(FP,"/","\")
        LK=SC&"\"&FP
    Else
        LK = SC
    End If    

    OpenByPath(LK)
End Function

Sub CommandOfBuildKK()
    Dim PN, mRmOut, mOTA, str1, str2, FinalCommand
    PN = element_getValue("project_name")
    mRmOut = element_isChecked("rm_out_kk_id")
    mBuildMode = element_isChecked("is_user_kk_id")
    mOTA = element_isChecked("build_ota_kk_id")
    str1 = ""
    str3 = ""

    If mRmOut Then
        str1 = "rm -rf out/ && "
    End If

    If mBuildMode Then
        buildModeStr = "user"
    Else
        buildModeStr = "eng"
    End If

    str2 = "./mk -o=TARGET_BUILD_VARIANT="&buildModeStr&" "&PN&" n"

    If mOTA Then
        str3 = " && ./mk -o=TARGET_BUILD_VARIANT="&buildModeStr&" "&PN&" otapackage"
    End If

    FinalCommand = str1 & str2 & str3
    CopyString(FinalCommand)
End Sub

Sub CommandOfBuildL1()
    Dim mRmOut, mOTA, str1, str2, FinalCommand
    mRmOut = element_isChecked("rm_out_l1_id")
    mOTA = element_isChecked("build_ota_l1_id")
    str1 = ""
    str2 = "make -j24 2>&1 | tee build.log"
    str3 = ""

    If mRmOut Then
        str1 = "rm -rf out/ && "
    End If
    If mOTA Then
        str3 = " && make -j24 otapackage 2>&1 | tee build_ota.log"
    End If

    FinalCommand = str1 & str2 & str3
    CopyString(FinalCommand)
End Sub

Function getProjectName()
    Dim path, FileName, PJValue
    path = element_getValue("getProjectName_id")
    FileName = getFileNameByStrInFolder(path, "checklist")

    If FileName <> "" Then
        PJValue = getPJValueByStrInExcel(path&"\"&FileName)
        If PJValue <> "" Then
            element_setValue "project_name", PJValue
        End If
    Else
        msgbox("软件路径不存在checklist，请确认后重试")
    End If
End Function

Function getFileNameByStrInFolder(path, str)
    Dim Folder, File, FileName
    Set Folder = Fso.GetFolder(path&"\")
    Set File = Folder.Files
    FileName = ""

    If Fso.FolderExists(path) Then
        For Each File in File
            If InStr(File.Name,str) > 0 Then
                FileName = File.Name
                Exit For
            End If
        Next
    ElseIf Len(path) = 0 Then
        msgbox("请输入路径")
    Else
        msgbox("路径不存在，请确认后重试")
    End If

    Set Folder = Nothing
    Set File = Nothing

    getFileNameByStrInFolder = FileName
End Function

Function getPJValueByStrInExcel(path)
    Dim ExcelApp, ExcelBook, ExcelSheet, RowCount, PJValue, CellValue, Count
    PJValue = ""
    Set ExcelApp = CreateObject("Excel.Application")
    Set ExcelBook= ExcelApp.Workbooks.Open(path)
    Set ExcelSheet = ExcelBook.Sheets(1)

    For RowCount = 1 To 1000
        If Not ExcelSheet.Rows(RowCount).Hidden Then
            CellValue = ExcelSheet.Cells(RowCount,2).value
            Count = InStr(CellValue, "project")
            If Count > 0 Then
                PJValue = Trim(Mid(CellValue,Count+8))
                Exit For
            End If
        End If
    Next

    ExcelBook.Close
    ExcelApp.Quit
    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelApp = Nothing

    getPJValueByStrInExcel = PJValue
End Function



Function SearchStr(fread,Str,LK,flag)
    Dim LineStr, fread_new
    Do Until fread.AtEndOfStream
        LineStr=fread.readline
        If InStr(LineStr,Str) > 0 Then
            If InStr(LineStr,"#"&Str) = 0 Then
                SearchStr = Mid(LineStr,Len(Str)+1)
                flag = 1
                Exit Do
            End If
        End If
    Loop
    If flag = 0 Then
        Set fread_new = Fso.opentextfile(LK,1)
        SearchStr = SearchStr(fread_new,Str,LK,1)
    End If   
End Function

Function checkGPS()
    GPS_flag = ""
    LK = getPath("ProjectConfig.mk")
    If Fso.FileExists(LK) Then
        Set fread = Fso.opentextfile(LK,1)

        str_Version = Trim(SearchStr(fread,"ROCO_SW_VERSION=",LK,0))
        str_Modem = Trim(SearchStr(fread,"CUSTOM_MODEM=",LK,0))

        If InStr(str_Version,"LOW") > 0 Then
            GPS_flag = GPS_flag&"1"
        End If

        If InStr(str_Modem,"LOW_GPS") > 0 Then
            GPS_flag = GPS_flag&"2"
        End If

        LK = getPath("custom")
        If Fso.FolderExists(LK) Then
            If checkStrExistInFile(LK&"\cgen\cfgdefault\CFG_GPS_Default.h", "0xFE") = "yes" Then
                GPS_flag = GPS_flag&"3"
            End If

            If checkStrExistInFile(LK&"\hal\ant\mt6582_ant_m1\WMT_SOC.cfg", "flag=1") = "yes" Then
                GPS_flag = GPS_flag&"4"
            End If

            msgboxInfoOfCheckGPS(GPS_flag)
        Else
            msgbox("custom目录不存在")
        End If
    Else
        msgbox("ProjectConfig.mk文件不存在")
    End If
End Function

Function checkStrExistInFile(path, str)
    Dim flag
    flag = ""

    If Fso.FileExists(path) Then
        Set oFile = Fso.OpenTextFile(path,1)
        Do Until oFile.AtEndOfStream
            If InStr(oFile.ReadLine, str) > 0 Then
                flag = "yes"
                Exit Do
            End If
        Loop
        Set oFile = Nothing
    Else
        msgbox("不存在"&path)
    End If

    checkStrExistInFile = flag
End Function

Function msgboxInfoOfCheckGPS(flag)
    If flag = "" Then
        msgbox "GPS不降成本"
    ElseIf flag = "1234" Then
        msgbox "GPS降成本"
    Else
        str_msgbox = "改动不完全:"
        if Not (InStr(flag, "1") > 0)  Then
            str_msgbox = str_msgbox&vbLf&"Version"
        End If
        if Not (InStr(flag, "2") > 0) Then
            str_msgbox = str_msgbox&vbLf&"Modem"
        End If
        if Not (InStr(flag, "3") > 0) Then
            str_msgbox = str_msgbox&vbLf&"CFG_GPS_Default.h"
        End If
        if Not (InStr(flag, "4") > 0) Then
            str_msgbox = str_msgbox&vbLf&"WMT_SOC.cfg"
        End If
        msgbox str_msgbox
    End If
End Function



Function checkOTA()
    OTA_flag = ""
    LK = getPath("ProjectConfig.mk")
    Set fso = createobject("scripting.filesystemobject")
    Set fread = fso.opentextfile(LK,1)

    str_SHULIAN = Trim(SearchStr(fread,"ROCO_SHULIAN_FOTA_SUPPORT=",LK,0))
    str_FOTA1 = Trim(SearchStr(fread,"ADUPS_FOTA_SUPPORT=",LK,0))
    str_FOTA2 = Trim(SearchStr(fread,"ROCO_DISPLAY_ID=",LK,0))

    If str_SHULIAN = "yes" Then
        If str_FOTA1 <> "yes" Then
            msgbox("数联OTA，注意修改版本号哦！")
        Else
            msgbox("数联OTA和广升OTA同时存在，请检查！")
        End If
    ElseIf str_FOTA1 = "yes" Then
        If Len(str_FOTA2) > 0 Then
            msgbox("广升OTA，注意修改DISPLAY_ID哦！")
        Else
            msgbox("广升OTA，未写DISPLAY_ID，请检查！")
        End If
    Else
        msgbox("不带OTA")
    End If
End Function


Function checkStrExistInFileWithPound(path, str)
    Dim flag, tmpStr
    flag = ""

    If Fso.FileExists(path) Then
        Set oFile = Fso.OpenTextFile(path,1)
        Do Until oFile.AtEndOfStream
            tmpStr = oFile.ReadLine
            If (InStr(tmpStr, str) > 0) AND (NOT InStr(tmpStr,"#") > 0) Then
                flag = "yes"
                Exit Do
            End If
        Loop
        Set oFile = Nothing
    Else
        msgbox("不存在"&path)
    End If

    checkStrExistInFileWithPound = flag
End Function

Function checkFileExist(path, fileName)
    Dim flag
    flag = ""

    If Fso.FolderExists(path) Then
        If Fso.FileExists(path&"\"&fileName) Then
            flag = "yes"
        End If
    Else
        msgbox("不存在"&path)
    End If

    checkFileExist = flag
End Function

Function getProductInfo(fileName)
    LK = getPath("product_roco.mk")
    getProductInfo = checkStrExistInFileWithPound(LK, fileName)
End Function

Function getCustomInfo(fileName)
    LK = getPath("custom")&"\prebuilts\system\media"
    getCustomInfo = checkFileExist(LK, fileName)
End Function

Function getSystemPropInfoForShut()
    LK = getPath("system.prop")
    getSystemPropInfoForShut = checkStrExistInFileWithPound(LK, "ro.operator.optr=CUST")
End Function

Function checkProductEqualsCustom(fileName)
    Dim flag
    flag = ""

    If getProductInfo(fileName) = getCustomInfo(fileName) Then
        flag = "yes"
    Else
        msgbox(fileName&"有问题，请检查"&vbLf&"getProductInfo="&getProductInfo(fileName)&vbLf&"getCustomInfo="&getCustomInfo(fileName))
    End If

    If InStr(fileName, "shut") Then
        If getProductInfo(fileName) = "yes" AND getSystemPropInfoForShut() = "" Then
            msgbox("system.prop文件中缺少ro.operator.optr=CUST")
            flag = ""
        End If
    End If

    checkProductEqualsCustom = flag
End Function

Function checkProductFinal()
    Dim flag
    flag = "ok"
    If checkProductEqualsCustom("bootanimation.zip") = "" Then
        flag = ""
    End If
    If checkProductEqualsCustom("bootanimation1.zip") = "" Then
        flag = ""
    End If
    If checkProductEqualsCustom("shutanimation.zip") = "" Then
        flag = ""
    End If
    If checkProductEqualsCustom("shutanimation1.zip") = "" Then
        flag = ""
    End If
    If checkProductEqualsCustom("bootaudio.mp3") = "" Then
        flag = ""
    End If
    If checkProductEqualsCustom("bootaudio1.mp3") = "" Then
        flag = ""
    End If
    If checkProductEqualsCustom("shutaudio.mp3") = "" Then
        flag = ""
    End If
    If checkProductEqualsCustom("shutaudio1.mp3") = "" Then
        flag = ""
    End If

    If flag = "ok" Then
        msgbox("检查OK")
    End If
End Function

Function checkGsensor()
    Dim flag, checkStr1, checkStr2, checkFolder
    flag = "ok"
    checkStr1 = "ROCO_ACCELEROMETER=bma2xx_auto mir3da_auto mc3xxx_auto"
    checkStr2 = "MTK_AUTO_DETECT_ACCELEROMETER=yes"
    checkFolder = getPath("custom")
    If checkStrExistInFileWithPound(getPath("ProjectConfig.mk"), checkStr1) = "" Then
        flag = ""
    End If
    If checkStrExistInFileWithPound(getPath("ProjectConfig.mk"), checkStr2) = "" Then
        flag = ""
    End If
    If flag = "ok" Then
        msgbox("检查OK")
    Else
        msgbox("未兼容")
    End If
End Function

Function checkBattery()
    Dim flag, checkFile
    flag = "ok"
    checkFile1 = getPath("codePath_KK") & "\mediatek/platform/mt6572/lk/platform.c"
    checkFile2 = getPath("codePath_KK") & "\mediatek/platform/mt6572/lk/mt_battery.c"
    
    If Not (checkStrExistInFile(checkFile1, "//    mt65xx_backlight_on();") = "yes" And checkStrExistInFile(checkFile1, "//    mt65xx_bat_init();") = "yes") Then
        flag = ""
    ElseIf Not (checkStrExistInFile(checkFile2, "pchr_turn_off_charging") = "yes" And checkStrExistInFile(checkFile2, "check_bat_protect_status") = "yes") Then
        flag = ""
    End If

    If flag = "ok" Then
        msgbox("检查OK")
    Else
        msgbox("未修改")
    End If
End Function

Function WriteHistory(DictName,inputId,ulId,sVersion)
    Dim str, fso
    str = element_getValue(inputId)
    Set fso = createobject("scripting.filesystemobject")
    If Len(str) = 0 Then
        msgbox("请输入")
    Else
        Dim DictPath, oDict
        DictPath = ws.CurrentDirectory
        If fso.FileExists(DictPath&"\"&DictName&".txt")=False Then
            Set oDict = fso.CreateTextFile(DictPath&"\"&DictName&".txt", True)
        End If    
        Set oDict = fso.OpenTextFile(DictPath&"\"&DictName&".txt",1)

        Dim strArray(9), lineCount
        lineCount = 0
        Do Until oDict.AtEndOfStream
            strArray(lineCount) = oDict.ReadLine
            lineCount = lineCount + 1
        Loop
        
        Dim i, j, equalsFlag, countFlag
        equalsFlag = -1 
        countFlag = lineCount - 1    

        For i = 0 To countFlag
            If strArray(i) = str Then
                equalsFlag = i     
                Exit For
            End If
        Next

        If equalsFlag = -1 Then
            If countFlag < 9 Then
                Set oDict = fso.OpenTextFile(DictPath&"\"&DictName&".txt",8)
                oDict.WriteLine(str)
                addBeforeLi str,inputId,ulId,sVersion
            Else
                For i = 0 To countFlag - 1
                    strArray(i) = strArray(i + 1)
                Next
                strArray(countFlag) = str
                oDict.Close
                fso.DeleteFile(DictPath&"\"&DictName&".txt")
                Set oDict = fso.CreateTextFile(DictPath&"\"&DictName&".txt", True)
                For i = 0 To countFlag
                    oDict.WriteLine(strArray(i))
                Next
                parentNode_removeChild "PN_ul_id", countFlag
                addBeforeLi str,inputId,ulId,sVersion
            End If
        ElseIf equalsFlag < countFlag Then
            For j = equalsFlag To countFlag - 1
                strArray(j) = strArray(j + 1)
            Next
            strArray(countFlag) = str
            oDict.Close
            fso.DeleteFile(DictPath&"\"&DictName&".txt")
            Set oDict = fso.CreateTextFile(DictPath&"\"&DictName&".txt", True)
            For i = 0 To countFlag
                oDict.WriteLine(strArray(i))
            Next
            parentNode_removeChild "PN_ul_id", countFlag-equalsFlag
            addBeforeLi str,inputId,ulId,sVersion
        End If

        oDict.Close
    End If
    Set fso = Nothing
End Function
    
Function ReadHistory(DictName,inputId,ulId,sVersion)
    Dim DictPath, oDict
    DictPath = ws.CurrentDirectory
    If Fso.FileExists(DictPath&"\"&DictName&".txt")=False Then
        Set oDict = fso.CreateTextFile(DictPath&"\"&DictName&".txt", True)
        Set oDict = fso.OpenTextFile(DictPath&"\"&DictName&".txt",1)
    Else
        Set oDict = fso.OpenTextFile(DictPath&"\"&DictName&".txt",1)
    End If    

    Dim strArray(9), lineCount
    lineCount = 0
    Do Until oDict.AtEndOfStream
        strArray(lineCount) = oDict.ReadLine
        lineCount = lineCount + 1
    Loop

    Dim i, countFlag
    countFlag = lineCount - 1   

    For i = 0 To countFlag
        addBeforeLi strArray(i),inputId,ulId,sVersion
    Next
End Function

Function onloadL1Project1()
    removeAllMoreButton("L1")
    
    Dim SC, LK, flag, OnchangeFunName
    SC = element_getValue("codePath_L1")
    LK = SC&"\device\joya_sz\"
    flag = 0
    If Fso.FolderExists(LK) Then
        addSelectOfL1Project1()
        removeAllOption "projectFolder_1"

        Set MyFolder = Fso.GetFolder(LK)
        Set Folders = MyFolder.SubFolders
        For Each Folder in Folders
            If Fso.FolderExists(LK&Folder.Name&"\roco\") Then
                addOption "projectFolder_1",Folder.Name
                flag = 1
            End If
        Next
        Set MyFolder = Nothing
        Set Folders = Nothing
        
        If flag = 1 Then
            onloadL1Project2()
        End If
    Else
        msgbox("not found device\joya_sz")
    End If
End Function

Function onloadL1Project2()
    removeAllMoreButton("L1")

    Dim PJ1, LK, SC, flag
    SC = element_getValue("codePath_L1")
    PJ1 = element_getValue("projectFolder_1")
    LK = SC&"\device\joya_sz\"&PJ1&"\roco\"

    addSelectOfL1Project2()
    removeAllOption "projectFolder_2"

    Set MyFolder = Fso.GetFolder(LK)
    Set Folders = MyFolder.SubFolders
    For Each Folder in Folders
        addOption "projectFolder_2",Folder.Name
    Next
    Set MyFolder = Nothing
    Set Folders = Nothing
End Function

Function getPathL1(Name)
    SC=element_getValue("codePath_L1")
    PJ1 = element_getValue("projectFolder_1")
    PJ2 = element_getValue("projectFolder_2")

    Select Case Name
        Case "alps_L1"
            LK = SC
        Case "device1_L1"
            LK = SC&"\device\joya_sz\"&PJ1
        Case "device2_L1"
            LK = SC&"\device\joya_sz\"&PJ1&"\roco\"&PJ2
        Case "custom_L1"
            LK = SC&"\vendor\mediatek\proprietary\custom\"&PJ1
        Case "out_L1"
            LK = SC&"\out\target\product"
        Case "ProjectConfig.mk"
            LK = SC&"\device\joya_sz\"&PJ1&"\ProjectConfig.mk"
        Case "system.prop"
            LK = SC&"\device\joya_sz\"&PJ1&"\roco\"&PJ2&"\system.prop"
        Case "items.ini"
            LK = SC&"\device\joya_sz\"&PJ1&"\roco\"&PJ2&"\items.ini"
    End Select

    getPathL1 = LK

End Function

Sub openProjectFileL1(Name)
    LK=getPathL1(Name)
    If Fso.FileExists(LK) Then
        ws.run textEditorPath&" "&LK
    Else
        msgbox(Name&"不存在")
    End If
End Sub

Sub openProjectFolderL1(Name)
    LK=getPathL1(Name)

    Select Case Name
        Case "out_L1"
            If Fso.FolderExists(LK) Then
                Set MyFolder_PN = Fso.GetFolder(LK)
                Set Folders = MyFolder_PN.SubFolders
                For Each Folder in Folders
                    FDN=Folder.Name
                    If Fso.FolderExists(LK&"\"&FDN&"\system") Then
                        PN1=Folder.Name
                        Exit For 
                    End If
                Next
                LK=LK&"\"&PN1
            End If
    End Select

    If Fso.FolderExists(LK) Then
        ws.run "explorer.exe "&LK
    Else
        msgbox(Name&"不存在")
    End If
End Sub

Function sortArray(ArrayString)
    Dim i, j, length, tmp
    length = UBound(ArrayString)
    For i=0 to length-1
        For j=i+1 to length
            If (StrComp(ArrayString(i),ArrayString(j)) > 0) Then
                tmp = ArrayString(i)
                ArrayString(i) = ArrayString(j)
                ArrayString(j) = tmp
            End If
        Next
    Next
    sortArray = ArrayString
End Function

Function onloadFolderNameList(FolderPath,SelectId)
    Dim count, ArrayString(), i, ArrayStringGet
    ReDim ArrayString(0)
    count = 0

    removeAllOption SelectId
    Set MyFolder = Fso.GetFolder(FolderPath)
    Set Folders = MyFolder.SubFolders
    For Each Folder in Folders
        ReDim Preserve ArrayString(count)
        ArrayString(count) = Folder.Name
        ''addOption SelectModemId,Folder.Name
        count = count + 1
    Next
    Set MyFolder = Nothing
    Set Folders = Nothing

    ArrayStringGet = sortArray(ArrayString)
    For i=1 to count
        addOption SelectId,ArrayStringGet(i-1)
    Next
End Function

Function getModemPath(AndroidVersion)
    if AndroidVersion = "KK" Then
        getModemPath = element_getValue("codePath_KK")&"\mediatek\custom\common\modem"
    Elseif AndroidVersion = "L1" Then
        Dim PJ1
        PJ1 = element_getValue("projectFolder_1")
        getModemPath = element_getValue("codePath_L1")&"\vendor\mediatek\proprietary\custom\"&PJ1&"\modem"
    End If
End Function

Function getSelectModemId(AndroidVersion)
    if AndroidVersion = "KK" Then
        getSelectModemId = "modem_kk_list_select"
    ElseIf AndroidVersion = "L1" Then
        getSelectModemId = "modem_l1_list_select"
    End If
End Function

Function copyModemName(AndroidVersion)
    Dim SelectModemId, ModemName
    SelectModemId = getSelectModemId(AndroidVersion)
    ModemName = element_getValue(SelectModemId)
    CopyString(ModemName)
End Function

Function applyProjectName(SelectId)
    Dim SelectValue
    SelectValue = element_getValue(SelectId)
    element_setValue "project_name",SelectValue
    removeAllMoreButton("KK")
End Function

Function changeButtonStatus(Name)
    If Name = "project" Then
        addFolderNameList Name, "Apply", getPath("config_only")
    ElseIf Name = "modem_kk" Then
        addFolderNameList Name, "Copy", getModemPath("KK")
    ElseIf Name = "modem_l1" Then
        addFolderNameList Name, "Copy", getModemPath("L1")
    End If
End Function