Option Explicit

Dim mCmdInput : Set mCmdInput = (New InputText)(getCmdInputId())
Dim mOpenPathInput : Set mOpenPathInput = (New InputText)(getOpenPathInputId())
Dim mDrive : mDrive = "X:\work2\"
Dim mIp : Set mIp = New ProjectInputs
Dim mProductList : Set mProductList = (New InputWithOneLayerList)(getProductParentId(), getProductInputId(), "product")
Dim mProjectList : Set mProjectList = (New InputWithOneLayerList)(getProjectParentId(), getProjectInputId(), "project")
Dim mSdkPathList : Set mSdkPathList = (New InputWithTwoLayerList)(getSdkPathParentId(), getSdkPathInputId(), "sdkpath")
Dim mOpenPathList : Set mOpenPathList = (New InputWithTwoLayerList)(getOpenPathParentId(), getOpenPathInputId(), "openpath")
Dim mOutFileList : Set mOutFileList = (New ButtonWithOneLayerList)(getOutButtonParentId(), "outfile")
Dim mOpenButtonList : Set mOpenButtonList = (New ButtonWithOneLayerList)(getOpenButtonParentId(), "openbutton")
Dim mFileButtonList : Set mFileButtonList = (New ButtonWithOneLayerList)(getFileButtonParentId(), "filebutton")

Dim vaPathHistory : Set vaPathHistory = New VariableArray
Dim mCurrentPath
Dim mSaveString : Set mSaveString = New SaveString



Class VariableArray
    Private mName, mPreBound, mBound, mArray()

    Private Sub Class_Initialize
        mName = ""
        mPreBound = -1
        mBound = -1
    End Sub

    Public Property Get Name
        Name = mName
    End Property

    Public Property Let Name(sValue)
        mName = sValue
    End Property

    Public Sub SetPreBound(sValue)
        If isNumeric(sValue) Then
            If sValue > mBound Then
                ReDim Preserve mArray(sValue)
                mPreBound = sValue
            End If
        End If
    End Sub

    Public Property Get Bound
        Bound = mBound
    End Property

    Public Property Get V(seq)
        If Not isNumeric(seq) Then
            MsgBox("Error: Get V(seq) seq is not a number")
            Exit Property
        ELse
            seq = Cint(seq)
        End If

        'MsgBox("seq="&seq&" mBound="&mBound)
        If seq < 0 Or seq > mBound Then
            MsgBox("Error: Get V(seq) seq out of bound! seq="&seq&" mBound="&mBound)
            Exit Property
        End If

        If isObject(mArray(seq)) Then
            Set V = mArray(seq)
        Else
            V = mArray(seq)
        End If
    End Property

    Public Property Let V(seq, sValue)
        If Not isNumeric(seq) Then
            MsgBox("Error: Let V(seq) seq is not a number")
            Exit Property
        ELse
            seq = Cint(seq)
        End If

        If seq < 0 Or seq > mBound Then
            MsgBox("Error: Let V(seq) seq out of bound")
            Exit Property
        End If

        If isObject(sValue) Then
            Set mArray(seq) = sValue
        Else
            mArray(seq) = sValue
        End If
    End Property

    Public Sub Append(value)
        mBound = mBound + 1
        If mBound > mPreBound Then
            ReDim Preserve mArray(mBound)
        End If

        If isObject(value) Then
            Set mArray(mBound) = value
        ELse
            mArray(mBound) = value
        End If
    End Sub

    Public Sub ResetArray()
        mBound = -1
    End Sub

    Public Property Get InnerArray
        InnerArray = mArray
    End Property

    Public Property Let InnerArray(newArray)
        If Not isArray(newArray) Then
            MsgBox("Error: Set InnerArray(newArray) newArray is not array")
            Exit Property
        End If
        Call ResetArray()
        Dim i
        For i = 0 To UBound(newArray)
            mBound = mBound + 1
            ReDim Preserve mArray(mBound)
            mArray(mBound) = newArray(i)
        Next
    End Property

    Public Sub SwapTwoValues(seq1, seq2)
        If Not isNumeric(seq1) Or Not isNumeric(seq2) Then
            MsgBox("Error: SwapTwoValues(seq1, seq2) seq1 or seq2 is not a number")
            Exit Sub
        ELse
            seq1 = Cint(seq1)
            seq2 = Cint(seq2)
        End If

        If (seq1 < 0 Or seq1 > mBound) Or (seq2 < 0 Or seq2 > mBound) Then
            MsgBox("Error: SwapTwoValues(seq1, seq2) seq1 or seq2 out of bound")
            Exit Sub
        End If

        If isObject(mArray(seq1)) And isObject(mArray(seq2)) Then
            Dim oTmp1, oTmp2
            Set oTmp1 = mArray(seq1)
            Set oTmp2 = mArray(seq2)

            Set mArray(seq1) = Nothing
            Set mArray(seq2) = Nothing

            Set mArray(seq1) = oTmp2
            Set mArray(seq2) = oTmp1
        Else
            Dim sTmp
            sTmp = mArray(seq1)
            mArray(seq1) = mArray(seq2)
            mArray(seq2) = sTmp
        End If
    End Sub

    Public Sub InsertBySeq(seq, value)
        If Not isNumeric(seq) Then
            MsgBox("Error: insertBySeq(seq) seq is not a number")
            Exit Sub
        ELse
            seq = Cint(seq)
        End If

        If seq < 0 Or seq > mBound Then
            MsgBox("Error: insertBySeq(seq) seq out of bound")
            Exit Sub
        End If

        mBound = mBound + 1
        ReDim Preserve mArray(mBound + 1)

        Dim i
        If isObject(value) Then
            If seq <> mBound - 1 Then
                For i = mBound To seq + 2 Step -1
                    Set mArray(i) = mArray(i - 1)
                Next
            End If

            Set mArray(seq + 1) = value
        Else
            If seq <> mBound - 1 Then
                For i = mBound To seq + 2 Step -1
                    mArray(i) = mArray(i - 1)
                Next
            End If

            mArray(seq + 1) = value
        End If
    End Sub
    Public Sub PopBySeq(seq)
        If Not isNumeric(seq) Then
            MsgBox("Error: PopBySeq(seq) seq is not a number")
            Exit Sub
        ELse
            seq = Cint(seq)
        End If

        If seq < 0 Or seq > mBound Then
            MsgBox("Error: PopBySeq(seq) seq out of bound")
            Exit Sub
        End If

        If seq <> mBound Then
            Dim i
            If isObject(mArray(seq)) Then
                For i = seq To mBound - 1
                    Set mArray(i) = mArray(i + 1)
                Next
            Else
                For i = seq To mBound - 1
                    mArray(i) = mArray(i + 1)
                Next
            End If
        End If

        mBound = mBound - 1
        ReDim Preserve mArray(mBound)
    End Sub

    Public Sub MoveToTop(seq)
        If Not isNumeric(seq) Then
            MsgBox("Error: MoveToTop(seq) seq is not a number")
            Exit Sub
        ELse
            seq = Cint(seq)
        End If

        If seq < 0 Or seq > mBound Then
            MsgBox("Error: MoveToTop(seq) seq out of bound")
            Exit Sub
        End If

        If seq = 0 Then Exit Sub

        Dim i, sValueToBeMove
        If isObject(mArray(seq)) Then
            Set sValueToBeMove = mArray(seq)
            For i = seq To 1 Step -1
                Set mArray(i) = mArray(i - 1)
                Set mArray(0) = sValueToBeMove
            Next
        Else
            sValueToBeMove = mArray(seq)
            For i = seq To 1 Step -1
                mArray(i) = mArray(i - 1)
                mArray(0) = sValueToBeMove
            Next
        End If
    End Sub

    Public Sub MoveToEnd(seq)
        If Not isNumeric(seq) Then
            MsgBox("Error: MoveToEnd(seq) seq is not a number")
            Exit Sub
        ELse
            seq = Cint(seq)
        End If

        If seq < 0 Or seq > mBound Then
            MsgBox("Error: MoveToEnd(seq) seq out of bound")
            Exit Sub
        End If

        If seq = 0 Then Exit Sub

        Dim i, sValueToBeMove
        If isObject(mArray(seq)) Then
            Set sValueToBeMove = mArray(seq)
            For i = seq To mBound - 1
                Set mArray(i) = mArray(i + 1)
            Next
            Set mArray(mBound) = sValueToBeMove
        Else
            sValueToBeMove = mArray(seq)
            For i = seq To mBound - 1
                mArray(i) = mArray(i + 1)
            Next
            mArray(mBound) = sValueToBeMove
        End If
    End Sub

    Public Function GetIndexIfExist(value)
        If mBound = -1 Then
            GetIndexIfExist = -1
            Exit Function
        End If

        Dim i
        For i = 0 To mBound
            If StrComp(mArray(i), value) = 0 Then
                GetIndexIfExist = i
                Exit Function
            End If
        Next
        GetIndexIfExist = -1
    End Function

    Public Function IsExistInObject(value, seq)
        If mBound = -1 Then
            IsExistInObject = False
            Exit Function
        End If

        Dim i
        For i = 0 To mBound
            If StrComp(mArray(i).V(seq), value) = 0 Then
                IsExistInObject = True
                Exit Function
            End If
        Next
        IsExistInObject = False
    End Function

    Public Sub SortArray()
        If mBound <= 0 Then
            'MsgBox("Error: SortArray() mBound <= 0, no need to sort")
            Exit Sub
        End If

        Dim i, j
        For i = 0 To mBound - 1
            For j = i + 1 To mBound
                If StrComp(mArray(i), mArray(j)) > 0 Then
                    Dim sTmp : sTmp = mArray(i) : mArray(i) = mArray(j) : mArray(j) = sTmp
                End If
            Next
        Next
    End Sub

    Public Function ToStringWithSpace()
        If mBound = -1 Then ToStringWithSpace = "" : Exit Function

        Dim i, sTmp
        For i = 0 To mBound
            sTmp = sTmp & " " & mArray(i)
        Next

        ToStringWithSpace = Trim(sTmp)
    End Function

    Public Function ToString()
        If mBound <> -1 Then
            Dim i, sTmp
            sTmp = "v(0) = " & mArray(0)
            If mBound > 0 Then
                For i = 1 To mBound
                    If isArray(mArray(i)) Then
                        sTmp = sTmp & Vblf & "v(" & i & ") = " & join(mArray(i))
                    ElseIf isObject(mArray(i)) Then
                        sTmp = sTmp & Vblf & "v(" & i & ") = [Object]"
                    Else
                        sTmp = sTmp & Vblf & "v(" & i & ") = " & mArray(i)
                    End If
                Next
            End If
            ToString = sTmp
        Else
            MsgBox("Error: ToString() mArray has no element")
        End If
    End Function
End Class



Class InputText
    Private mInputId

    Public Default Function Constructor(inputId)
        mInputId = inputId
        Set Constructor = Me
    End Function

    Private Function checkElement()
        If isElementIdExist(mInputId) Then
            checkElement = True
        Else
            MsgBox("Element not exist!" & VbLf & "Id: " & mInputId)
            checkElement = False
        End If
    End Function

    Public Property Get Text
        If checkElement() Then
            Text = document.getElementById(mInputId).value
        Else
            Text = ""
        End If
    End Property

    Public Sub setText(text)
        If checkElement() Then
            document.getElementById(mInputId).value = text
        End If
    End Sub
End Class



Class InputWithOneLayerList
    Private mParentId, mInputId, mListDivId, mListUlId, mVaArray

    Public Default Function Constructor(parentId, inputId, name)
        mParentId = parentId
        mInputId = inputId
        mListDivId = "list_div_" & name
        mListUlId = "list_ul_" & name
        Set Constructor = Me
    End Function

    Public Sub addList(vaArray)
        Set mVaArray = vaArray
        If mVaArray.Bound = -1 Then Exit Sub

        Call removeLi(mListUlId)
        Call setInputClickFun(mParentId, mInputId, mListDivId)
        Call addListUL(mParentId, mListDivId, mListUlId)
        Dim i : For i = 0 To mVaArray.Bound
            Call addListLi(mParentId, mInputId, mListDivId, mListUlId, mVaArray.V(i), True)
        Next
    End Sub
End Class

Class ButtonWithOneLayerList
    Private mParentId, mListDivId, mListUlId, mVaArray, mFocusLiIndex

    Public Default Function Constructor(parentId, name)
        mParentId = parentId
        mListDivId = "list_div_" & name
        mListUlId = "list_ul_" & name
        Set mVaArray = New VariableArray
        mFocusLiIndex = -1
        Set Constructor = Me
    End Function

    Public Property Get VaArray
        Set VaArray = mVaArray
    End Property

    Public Sub addList()
        If mVaArray.Bound = -1 Then Exit Sub

        Call removeLi(mListUlId)
        Call addListUL(mParentId, mListDivId, mListUlId)
        Dim i : For i = 0 To mVaArray.Bound
            Call addListLi(mParentId, "", mListDivId, mListUlId, mVaArray.V(i), False)
        Next
    End Sub

    Public Sub removeList()
        Call removeLi(mListUlId)
        Call mVaArray.ResetArray()
    End Sub

    Public Sub toggleButtonList()
        Call toggleListDiv(mParentId, mListDivId)
    End Sub

    Public Function isShowing()
        isShowing = isDivShowing(mListDivId)
    End Function

    Public Function hideListIfShowing()
        If isDivShowing(mListDivId) Then
            Call removeList()
            Call toggleButtonList()
            hideListIfShowing = True
        Else
            hideListIfShowing = False
        End If
    End Function

    Public Function changeListFocus(keyCode)
        If isDivShowing(mListDivId) Then
            If keyCode = KEYCODE_UP Then
                mFocusLiIndex =  changeLiFocusUp(mListUlId)
            ElseIf keyCode = KEYCODE_DOWN Then
                mFocusLiIndex = changeLiFocusDown(mListUlId)
            End If
            changeListFocus = True : Exit Function
        End If
        changeListFocus = False
    End Function

    Public Sub clickFocusedLi()
        If mFocusLiIndex > -1 Then
            Call clickListLi(mListUlId, mFocusLiIndex)
        End if
    End Sub
End Class



Class InputWithTwoLayerList
    Private mParentId, mInputId, mDirDivId, mDirUlId, mListDivId, mListUlId, mVaArray

    Public Default Function Constructor(parentId, inputId, name)
        mParentId = parentId
        mInputId = inputId
        mDirDivId = "list_dir_div_" & name
        mDirUlId = "list_dir_ul_" & name
        mListDivId = "list_div_"
        mListUlId = "list_ul_"
        Set Constructor = Me
    End Function

    Public Sub addList(vaArray)
        Set mVaArray = vaArray
        If mVaArray.Bound = -1 Then Exit Sub

        Call addListUL(mParentId, mDirDivId, mDirUlId)
        Dim i, j, category

        For i = 0 To mVaArray.Bound
            category = mVaArray.V(i).Name
            Call addListDirectoryLi(mParentId, mDirDivId, mDirUlId, mListDivId & LCase(category), category)

            Call addListUL(mParentId, mListDivId & LCase(category), mListUlId & LCase(category))

            if mVaArray.V(i).Bound <> -1 Then
                For j = 0 To mVaArray.V(i).Bound
                    Call addListLi(mParentId, mInputId, mListDivId & LCase(category), mListUlId & LCase(category), mVaArray.V(i).V(j), True)
                Next
            End If
        Next
    End Sub

    Public Sub toggleList()
        If mVaArray.Bound = -1 Then Exit Sub

        If isDivShowing(mDirDivId) Then
            Call hideListDiv(mParentId, mDirDivId)
        Else
            Dim i, category, isShow
            isShow = False
            For i = 0 To mVaArray.Bound
                category = mVaArray.V(i).Name
                If isDivShowing(mListDivId & LCase(category)) Then
                    isShow = True
                    Exit For
                End If
            Next
            If isShow Then
                Call hideListDiv(mParentId, mListDivId & LCase(category))
            Else
                Call showListDiv(mParentId, mDirDivId)
            End If
        End If
    End Sub
End Class



Function getOutPath(Product)
    getOutPath = "out/target/product/" & Product
End Function

Function getDownloadOutPath(Product)
    Dim outPath : outPath = mIp.Infos.getPathWithDriveSdk(getOutPath(Product))
    If isFolderExists(outPath & "\merged") Then
        getDownloadOutPath = outPath & "\merged"
    Else
        getDownloadOutPath = outPath
    End If
End Function

Function getProductPath(Product)
    getProductPath = "weibu/" & Product
End Function

Function getProjectPath(Product, Project)
    getProjectPath = "weibu/" & Product & "/" & Project
End Function


Sub msgboxPathNotExist(path)
    If Trim(path) <> "" Then
        MsgBox("path not exist! " & VbLf & path)
    End If
End Sub

Class ProjectInfos
    Private mWork, mSdk, mProduct, mProject, mFirmware, mRequirements, mZentao, mTaskNum
    Private mProjectAlps, mBootLogo, mSysTarget, mKernelVer, mTargetArch

    Public Property Get Work
        Work = mWork
    End Property

    Public Property Get Sdk
        Sdk = mSdk
    End Property

    Public Property Get Product
        Product = mProduct
    End Property

    Public Property Get BootLogo
        BootLogo = mBootLogo
    End Property

    Public Property Get SysTarget
        SysTarget = mSysTarget
    End Property

    Public Property Get KernelVer
        KernelVer = mKernelVer
    End Property

    Public Property Get TargetArch
        TargetArch = mTargetArch
    End Property

    Public Property Get Project
        Project = mProject
    End Property

    Public Property Get Firmware
        Firmware = mFirmware
    End Property

    Public Property Get Requirements
        Requirements = mRequirements
    End Property

    Public Property Get Zentao
        Zentao = mZentao
    End Property

    Public Property Get TaskNum
        TaskNum = mTaskNum
    End Property

    Public Property Get DriveSdk
        DriveSdk = mDrive & mSdk
    End Property

    Public Property Get OutPath
        OutPath = getOutPath(Product)
    End Property

    Public Property Get DownloadOutPath
        DownloadOutPath = getDownloadOutPath(Product)
    End Property

    Public Property Get ProductPath
        ProductPath = getProductPath(Product)
    End Property

    Public Property Get ProjectPath
        ProjectPath = getProjectPath(Product, Project)
    End Property

    Public Property Get ProjectAlps
        ProjectAlps = mProjectAlps
    End Property

    Public Property Get DriverProject
        DriverProject = getDriverProjectName(Project)
    End Property

    Public Property Get DriverProjectPath
        DriverProjectPath = getProjectPath(Product, DriverProject)
    End Property

    Public Function getPathWithDriveSdk(path)
        getPathWithDriveSdk = DriveSdk & "/" & path
    End Function

    Public Function getOverlayPath(path)
        getOverlayPath = ProjectPath & ProjectAlps & "/" & path
    End Function

    Public Function getDriverOverlayPath(path)
        getDriverOverlayPath = DriverProjectPath & ProjectAlps & "/" & path
    End Function

    Sub getBootLogo()
        mBootLogo = getDriverProjectConfigValue("BOOT_LOGO")
        Call getFakeOrientation() 
    End Sub

    Sub getFakeOrientation()
        Dim csciPath
        csciPath = checkDriveSdkPath(DriverProjectPath & "/config/csci.ini")
        If Not isFileExists(csciPath) Then : Exit Sub

        Dim oText, sReadLine
        Set oText = oFso.OpenTextFile(csciPath, FOR_READING)

        Do Until oText.AtEndOfStream
            sReadLine = oText.ReadLine
            If InStr(sReadLine, "ro.vendor.fake.orientation") > 0 And InStr(sReadLine, " 1 ") > 0 Then
                mBootLogo = mBootLogo & "nl"
                Exit Do
            End If
        Loop
    End Sub

    Sub getSysTargetProject()
        Dim fullMkPath
        fullMkPath = "device/mediateksample/" & Product & "/full_" & Product & ".mk"
        If Not isFileExists(fullMkPath) Then Exit Sub
        mSysTarget = readTextAndGetValue("SYS_TARGET_PROJECT", fullMkPath)
    End Sub

    Sub getKernelInfos()
        Dim projectConfigPath, k64Support
        projectConfigPath = "/device/mediateksample/" & Product & "/ProjectConfig.mk"
        If Not isFileExists(projectConfigPath) Then MsgBox(projectConfigPath) : Exit Sub

        mKernelVer = readTextAndGetValue("LINUX_KERNEL_VERSION", projectConfigPath)
        k64Support = readTextAndGetValue("MTK_K64_SUPPORT", projectConfigPath)

        If k64Support = "yes" Then
            mTargetArch = "arm64"
        ElseIf k64Support = "no" Then
            mTargetArch = "arm"
        End If
        
    End Sub

    Public Property Let Work(value)
        mWork = value
    End Property

    Public Property Let Sdk(value)
        mSdk = value
    End Property

    Public Property Let Product(value)
        mProduct = value
    End Property

    Public Property Let Project(value)
        mProject = value
    End Property

    Public Property Let Firmware(value)
        mFirmware = value
    End Property

    Public Property Let Requirements(value)
        mRequirements = value
    End Property

    Public Property Let Zentao(value)
        mZentao = value
        If value <> "" And InStr(value, "task-view-") > 0 Then
            mTaskNum = Replace(Right(value, Len(value) - InStr(value, "task-view-") - Len("task-view-") + 1), ".html", "")
        End If
    End Property

    Public Property Let ProjectAlps(value)
        mProjectAlps = value
    End Property
End Class



Class ProjectInputs
    Private mInfos

    Public Sub Class_Initialize
        Set mInfos = New ProjectInfos
    End Sub

    Public Property Get Infos
        Set Infos = mInfos
    End Property

    Public Property Get Work
        Work = document.getElementById(getWorkInputId()).value
    End Property

    Public Property Get Sdk
        Sdk = document.getElementById(getSdkPathInputId()).value
    End Property

    Public Property Get Product
        Product = document.getElementById(getProductInputId()).value
    End Property

    Public Property Get Project
        Project = document.getElementById(getProjectInputId()).value
    End Property

    Public Property Get Firmware
        Firmware = document.getElementById(getFirmwareInputId()).value
    End Property

    Public Property Get Requirements
        Requirements = document.getElementById(getRequirementsInputId()).value
    End Property

    Public Property Get Zentao
        Zentao = document.getElementById(getZentaoInputId()).value
    End Property

    Public Property Let Work(value)
        Call setElementValue(getWorkInputId(), value)
        document.title = value & " " & mDrive
        Call onWorkChange()
    End Property

    Public Property Let Sdk(value)
        Call setElementValue(getSdkPathInputId(), value)
        Call onSdkChange()
    End Property

    Public Property Let Product(value)
        Call setElementValue(getProductInputId(), value)
        Call onProductChange()
    End Property

    Public Property Let Project(value)
        Call setElementValue(getProjectInputId(), value)
        Call onProjectChange()
    End Property

    Public Property Let Firmware(value)
        Call setElementValue(getFirmwareInputId(), value)
        Call onFirmwareChange()
    End Property

    Public Property Let Requirements(value)
        Call setElementValue(getRequirementsInputId(), value)
        Call onRequirementsChange()
    End Property

    Public Property Let Zentao(value)
        Call setElementValue(getZentaoInputId(), value)
        Call onZentaoChange()
    End Property

    Public Function hasProjectInfos()
        If mInfos.Sdk <> "" And mInfos.Product <> "" And mInfos.Project <> "" Then
            hasProjectInfos = True
        Else
            'MsgBox("No project infos!")
            hasProjectInfos = False
        End If
    End Function

    Public Function hasProjectAlps()
        If mInfos.ProjectAlps = "/alps" Then
            hasProjectAlps = True
        Else
            hasProjectAlps = False
        End If
    End Function

    Public Sub clearWorkInfos()
        Work = ""
        Firmware = "\\192.168.0.248\安卓固件文件\"
        Requirements = "\\192.168.0.24\wbshare\客户需求\"
        Zentao = "http://192.168.0.29:3000/zentao/task-view-1.html"
    End Sub

    Public Sub clearSdkInfos()
        Sdk = ""
        Product = ""
        Project = ""
    End Sub

    Public Sub cutSdkInOpenPath()
        If Trim(getOpenPath()) = "" Then Exit Sub
        If hasProjectInfos() Then
            Call replaceSlash()
            Call setOpenPath(Replace(getOpenPath(), relpaceSlashInPath(mInfos.DriveSdk) & "/", ""))
        End If
    End Sub

    Public Sub cutProjectInOpenPath()
        If Trim(getOpenPath()) = "" Then Exit Sub
        Call cutSdkInOpenPath()
        If hasProjectInfos() Then
            Call setOpenPath(Replace(getOpenPath(), mInfos.ProjectPath & mInfos.ProjectAlps & "/", ""))
            Call setOpenPath(Replace(getOpenPath(), mInfos.DriverProjectPath & mInfos.ProjectAlps & "/", ""))
        End If
    End Sub

    Public Function cutProject(path)
        path = relpaceSlashInPath(path)
        path = Replace(path, relpaceSlashInPath(mInfos.DriveSdk) & "/", "")
        path = Replace(path, mInfos.ProjectPath & mInfos.ProjectAlps & "/", "")
        path = Replace(path, mInfos.DriverProjectPath & mInfos.ProjectAlps & "/", "")
        cutProject = path
    End Function

    Public Sub onWorkChange()
        mInfos.Work = Work
    End Sub

    Public Function onSdkChange()
        'Call mIp.cutSdkInOpenPath()

        mInfos.Sdk = Sdk
        If isFolderExists(mInfos.DriveSdk) Then
            onSdkChange = True
        Else
            msgboxPathNotExist(mInfos.DriveSdk)
            Call setElementValue(getSdkPathInputId(), "")
            mInfos.Sdk = ""
            onSdkChange = False
        End If
    End Function

    Public Function onProductChange()
        mInfos.Product = Product
        If isFolderExists(mInfos.ProductPath) Then
            Call mInfos.getSysTargetProject()
            Call mInfos.getKernelInfos()
            onProductChange = True
        Else
            msgboxPathNotExist(mInfos.ProductPath)
            Call setElementValue(getProductInputId(), "")
            mInfos.Product = ""
            onProductChange = False
        End If
    End Function

    Public Function onProjectChange()
        'Call mIp.cutProjectInOpenPath()

        mInfos.Project = Project
        mInfos.ProjectAlps = ""
        If isFolderExists(mInfos.ProjectPath) Then
            If InStr(mInfos.Sdk, "_r") = 0 _
                    Or isFolderExists(mInfos.ProjectPath & "/alps") _
                    Or isFolderExists(mInfos.ProjectPath & "/config") Then
                mInfos.ProjectAlps = "/alps"
            Else
                mInfos.ProjectAlps = ""
            End If
            Call getProjectConfigMk()
            Call mInfos.getBootLogo()
            onProjectChange = True
        Else
            msgboxPathNotExist(mInfos.ProjectPath)
            Call setElementValue(getProjectInputId(), "")
            mInfos.Project = ""
            onProjectChange = False
        End If
    End Function

    Public Sub onFirmwareChange()
        mInfos.Firmware = Firmware
    End Sub

    Public Sub onRequirementsChange()
        mInfos.Requirements = Requirements
    End Sub

    Public Sub onZentaoChange()
        mInfos.Zentao = Zentao
    End Sub

End Class



Class SaveString
    Private mStr

    Public Property Let Str(value)
        mStr = value
    End Property

    Public Sub copy()
        Call CopyString(mStr)
    End Sub
End Class
