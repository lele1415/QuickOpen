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
    Private mParentId, mInputId, mListDivId, mListUlId, mSetValue

    Public Default Function Constructor(parentId, inputId, name, setValue)
        mParentId = parentId
        mInputId = inputId
        mListDivId = "list_div_" & name
        mListUlId = "list_ul_" & name
        mSetValue = setValue
        Set Constructor = Me
    End Function

    Public Sub addList(vaArray)
        If vaArray.Bound <> -1 Then
            Call removeLi(mListUlId)
            Call setInputClickFun(mParentId, mInputId, mListDivId)
            Call addListUL(mParentId, mListDivId, mListUlId)
            Dim i : For i = 0 To vaArray.Bound
                Call addListLi(mInputId, mListDivId, mListUlId, vaArray.V(i), mSetValue)
            Next
        End If
    End Sub

    Public Sub removeList()
        Call removeLi(mListUlId)
        Call resetInputOnClick(mInputId)
    End Sub
End Class



Class InputWithTwoLayerList
    Private mParentId, mInputId, mDirDivId, mDirUlId, mListDivId, mListUlId, mSetValue

    Public Default Function Constructor(parentId, inputId, name, setValue)
        mParentId = parentId
        mInputId = inputId
        mDirDivId = "list_dir_div_" & name
        mDirUlId = "list_dir_ul_" & name
        mListDivId = "list_div_"
        mListUlId = "list_ul_"
        mSetValue = setValue
        Set Constructor = Me
    End Function

    Public Sub addList(vaArray)
        If vaArray.Bound <> -1 Then
            Call setInputClickFun(mParentId, mInputId, mDirDivId)
            Call addListUL(mParentId, mDirDivId, mDirUlId)
            Dim i, j, category

            For i = 0 To vaArray.Bound
                category = vaArray.V(i).Name
                Call addListDirectoryLi(mParentId, mDirDivId, mDirUlId, mListDivId & LCase(category), category)

                Call addListUL(mParentId, mListDivId & LCase(category), mListUlId & LCase(category))

                if vaArray.V(i).Bound <> -1 Then
                    For j = 0 To vaArray.V(i).Bound
                        Call addListLi(mInputId, mListDivId & LCase(category), mListUlId & LCase(category), vaArray.V(i).V(j), mSetValue)
                    Next
                End If
            Next
        End If
    End Sub
End Class



Function getWeibuSdkPath(Sdk)
    getWeibuSdkPath = Sdk & "/weibu"
End Function

Function getOutSdkPath(Sdk, Product)
    getOutSdkPath = Sdk & "\out\target\product\" & Product
End Function

Function getProductPath(Product)
    getProductPath = "weibu/" & Product
End Function

Function getProductSdkPath(Sdk, Product)
    getProductSdkPath = Sdk & "/" & getProductPath(Product)
End Function

Function getProjectPath(Product, Project)
    getProjectPath = "weibu/" & Product & "/" & Project
End Function

Function getProjectSdkPath(Sdk, Product, Project)
    getProjectSdkPath = Sdk & "/" & getProjectPath(Product, Project)
End Function

Sub msgboxPathNotExist(path)
    If Trim(path) <> "" Then
        MsgBox("path not exist! " & VbLf & path)
    End If
End Sub

Class ProjectInfos
    Private mWork, mSdk, mProduct, mProject, mFirmware, mRequirements, mZentao
    Private mProjectAlps

    Public Property Get Work
        Work = mWork
    End Property

    Public Property Get Sdk
        Sdk = mSdk
    End Property

    Public Property Get Product
        Product = mProduct
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

    Public Property Get WeibuSdkPath
        WeibuSdkPath = getWeibuSdkPath(Sdk)
    End Property

    Public Property Get OutSdkPath
        OutSdkPath = getOutSdkPath(Sdk, Product)
    End Property

    Public Property Get ProductPath
        ProductPath = getProductPath(Product)
    End Property

    Public Property Get ProductSdkPath
        ProductSdkPath = getProductSdkPath(Sdk, Product)
    End Property

    Public Property Get ProjectPath
        ProjectPath = getProjectPath(Product, Project)
    End Property

    Public Property Get ProjectSdkPath
        ProjectSdkPath = getProjectSdkPath(Sdk, Product, Project)
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

    Public Property Get DriverProjectSdkPath
        DriverProjectSdkPath = getProjectSdkPath(Sdk, Product, DriverProject)
    End Property

    Public Function getOverlayPath(path)
        getOverlayPath = ProjectPath & ProjectAlps & "/" & path
    End Function

    Public Function getOverlaySdkPath(path)
        getOverlaySdkPath = ProjectSdkPath & ProjectAlps & "/" & path
    End Function

    Public Function getDriverOverlayPath(path)
        getDriverOverlayPath = DriverProjectPath & ProjectAlps & "/" & path
    End Function

    Public Function getDriverOverlaySdkPath(path)
        getDriverOverlaySdkPath = DriverProjectSdkPath & ProjectAlps & "/" & path
    End Function

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
            MsgBox("No project infos!")
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
        Firmware = ""
        Requirements = ""
        Zentao = ""
    End Sub

    Public Sub clearSdkInfos()
        Sdk = ""
        Product = ""
        Project = ""
    End Sub

    Public Sub onWorkChange()
        mInfos.Work = Work
    End Sub

    Public Function onSdkChange()
        mInfos.Sdk = Sdk
        If oFso.FolderExists(mInfos.Sdk) Then
            onSdkChange = True
        Else
            msgboxPathNotExist(mInfos.Sdk)
            Call setElementValue(getSdkPathInputId(), "")
            mInfos.Sdk = ""
            onSdkChange = False
        End If
    End Function

    Public Function onProductChange()
        mInfos.Product = Product
        If oFso.FolderExists(mInfos.ProductSdkPath) Then
            onProductChange = True
        Else
            msgboxPathNotExist(mInfos.ProductSdkPath)
            Call setElementValue(getProductInputId(), "")
            mInfos.Product = ""
            onProductChange = False
        End If
    End Function

    Public Function onProjectChange()
        mInfos.Project = Project
        mInfos.ProjectAlps = ""
        If oFso.FolderExists(mInfos.ProjectSdkPath) Then
            If oFso.FolderExists(mInfos.ProjectSdkPath & "/alps") Then
                mInfos.ProjectAlps = "/alps"
            Else
                mInfos.ProjectAlps = ""
            End If
            onProjectChange = True
        Else
            msgboxPathNotExist(mInfos.ProjectSdkPath)
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
