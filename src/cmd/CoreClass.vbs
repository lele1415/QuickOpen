Option Explicit

Class VariableArray
    Private mPreBound, mBound, mArray()

    Private Sub Class_Initialize
        mPreBound = -1
        mBound = -1
    End Sub

    Public Sub setPreBound(sValue)
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

    Public Property Get v(seq)
        If Not isNumeric(seq) Then
            MsgBox("Error: Get v(seq) seq is not a number")
            Exit Property
        ELse
            seq = Cint(seq)
        End If

        'MsgBox("seq="&seq&" mBound="&mBound)
        If seq < 0 Or seq > mBound Then
            MsgBox("Error: Get v(seq) seq out of bound! seq="&seq&" mBound="&mBound)
            Exit Property
        End If

        If isObject(mArray(seq)) Then
            Set V = mArray(seq)
        Else
            V = mArray(seq)
        End If
    End Property

    Public Property Let v(seq, sValue)
        If Not isNumeric(seq) Then
            MsgBox("Error: Let v(seq) seq is not a number")
            Exit Property
        ELse
            seq = Cint(seq)
        End If

        If seq < 0 Or seq > mBound Then
            MsgBox("Error: Let v(seq) seq out of bound")
            Exit Property
        End If

        If isObject(sValue) Then
            Set mArray(seq) = sValue
        Else
            mArray(seq) = sValue
        End If
    End Property

    Public Sub append(value)
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

    Public Sub resetArray()
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
        Call resetArray()
        Dim i
        For i = 0 To UBound(newArray)
            mBound = mBound + 1
            ReDim Preserve mArray(mBound)
            mArray(mBound) = newArray(i)
        Next
    End Property

    Public Sub swapTwoValues(seq1, seq2)
        If Not isNumeric(seq1) Or Not isNumeric(seq2) Then
            MsgBox("Error: swapTwoValues(seq1, seq2) seq1 or seq2 is not a number")
            Exit Sub
        ELse
            seq1 = Cint(seq1)
            seq2 = Cint(seq2)
        End If

        If (seq1 < 0 Or seq1 > mBound) Or (seq2 < 0 Or seq2 > mBound) Then
            MsgBox("Error: swapTwoValues(seq1, seq2) seq1 or seq2 out of bound")
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

    Public Sub insertBySeq(seq, value)
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
    Public Sub popBySeq(seq)
        If Not isNumeric(seq) Then
            MsgBox("Error: popBySeq(seq) seq is not a number")
            Exit Sub
        ELse
            seq = Cint(seq)
        End If

        If seq < 0 Or seq > mBound Then
            MsgBox("Error: popBySeq(seq) seq out of bound")
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

    Public Sub moveToTop(seq)
        If Not isNumeric(seq) Then
            MsgBox("Error: moveToTop(seq) seq is not a number")
            Exit Sub
        ELse
            seq = Cint(seq)
        End If

        If seq < 0 Or seq > mBound Then
            MsgBox("Error: moveToTop(seq) seq out of bound")
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

    Public Sub moveToEnd(seq)
        If Not isNumeric(seq) Then
            MsgBox("Error: moveToEnd(seq) seq is not a number")
            Exit Sub
        ELse
            seq = Cint(seq)
        End If

        If seq < 0 Or seq > mBound Then
            MsgBox("Error: moveToEnd(seq) seq out of bound")
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

    Public Function getIndexIfExist(value)
        If mBound = -1 Then
            getIndexIfExist = -1
            Exit Function
        End If

        Dim i
        For i = 0 To mBound
            If StrComp(mArray(i), value) = 0 Then
                getIndexIfExist = i
                Exit Function
            End If
        Next
        getIndexIfExist = -1
    End Function

    Public Function isExistInObject(value, seq)
        If mBound = -1 Then
            isExistInObject = False
            Exit Function
        End If

        Dim i
        For i = 0 To mBound
            If StrComp(mArray(i).V(seq), value) = 0 Then
                isExistInObject = True
                Exit Function
            End If
        Next
        isExistInObject = False
    End Function

    Public Sub sortArray()
        If mBound <= 0 Then
            'MsgBox("Error: sortArray() mBound <= 0, no need to sort")
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

    Public Function toStringWithSpace()
        If mBound = -1 Then toStringWithSpace = "" : Exit Function

        Dim i, sTmp
        For i = 0 To mBound
            sTmp = sTmp & " " & mArray(i)
        Next

        toStringWithSpace = Trim(sTmp)
    End Function

    Public Function toString()
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
            toString = sTmp
        Else
            MsgBox("Error: toString() mArray has no element")
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
            Text = getElementValue(mInputId)
        Else
            Text = ""
        End If
    End Property

    Public Sub setText(text)
        If checkElement() Then
            Call setElementValue(mInputId, text)
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

    Public Property Let VaArray(value)
        Set mVaArray = value
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
        Call mVaArray.resetArray()
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

    Public Function changeFocus(keyCode)
        If isDivShowing(mListDivId) Then
            If keyCode = KEYCODE_UP Then
                mFocusLiIndex =  changeLiFocusUp(mListUlId)
            ElseIf keyCode = KEYCODE_DOWN Then
                mFocusLiIndex = changeLiFocusDown(mListUlId)
            End If
            changeFocus = True : Exit Function
        End If
        changeFocus = False
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

Class SaveString
    Private mStr

    Public Property Let Str(value)
        mStr = value
    End Property

    Public Sub copy()
        Call CopyString(mStr)
    End Sub
End Class
