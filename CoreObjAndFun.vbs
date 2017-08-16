Const WINDOW_WIDTH = 400
Const WINDOW_HEIGHT = 800
Sub Window_OnLoad
    Dim ScreenWidth : ScreenWidth = CreateObject("HtmlFile").ParentWindow.Screen.AvailWidth
    Dim ScreenHeight : ScreenHeight = CreateObject("HtmlFile").ParentWindow.Screen.AvailHeight
    Window.MoveTo ScreenWidth - WINDOW_WIDTH ,(ScreenHeight - WINDOW_HEIGHT) \ 3
    Window.ResizeTo WINDOW_WIDTH, WINDOW_HEIGHT
End Sub

Set oWs=CreateObject("wscript.shell")
Set oFso=CreateObject("Scripting.FileSystemObject")

Const FOR_READING = 1
Const FOR_APPENDING = 8

Class VariableArray
    Private mLength, mArray()

    Private Sub Class_Initialize
        mLength = -1
    End Sub

    Public Property Get Length
        Length = mLength
    End Property

    Public Property Get Value(seq)
        If Not isNumeric(seq) Then
            MsgBox("Error: Get Value(seq) seq is not a number")
            Exit Property
        ELse
            seq = Cint(seq)
        End If

        'MsgBox("seq="&seq&" mLength="&mLength)
        If seq < 0 Or seq > mLength Then
            MsgBox("Error: Get Value(seq) seq out of bound")
            Exit Property
        End If

        If isObject(mArray(seq)) Then
            Set Value = mArray(seq)
        Else
            Value = mArray(seq)
        End If
    End Property

    Public Property Let Value(seq, sValue)
        If Not isNumeric(seq) Then
            MsgBox("Error: Let Value(seq) seq is not a number")
            Exit Property
        ELse
            seq = Cint(seq)
        End If

        If seq < 0 Or seq > mLength Then
            MsgBox("Error: Let Value(seq) seq out of bound")
            Exit Property
        End If

        mArray(seq) = sValue
    End Property

    Public Function Append(value)
        mLength = mLength + 1
        ReDim Preserve mArray(mLength)

        If isObject(value) Then
            Set mArray(mLength) = value
        ELse
            mArray(mLength) = value
        End If
    End Function

    Public Function ResetArray()
        mLength = -1
    End Function

    Public Property Let InnerArray(newArray)
        If Not isArray(newArray) Then
            MsgBox("Error: Set InnerArray(newArray) newArray is not array")
            Exit Property
        End If

        Dim i
        For i = 0 To UBound(newArray)
            mLength = mLength + 1
            ReDim Preserve mArray(mLength)
            mArray(mLength) = newArray(i)
        Next
    End Property

    Public Function PopBySeq(seq)
        If Not isNumeric(seq) Then
            MsgBox("Error: PopBySeq(seq) seq is not a number")
            Exit Function
        ELse
            seq = Cint(seq)
        End If

        If seq < 0 Or seq > mLength Then
            MsgBox("Error: PopBySeq(seq) seq out of bound")
            Exit Function
        End If

        If seq <> mLength Then
            Dim i
            For i = seq To mLength - 1
                mArray(i) = mArray(i + 1)
            Next
        End If

        mLength = mLength - 1
        ReDim Preserve mArray(mLength)
    End Function

    Public Function MoveToTop(seq)
        If Not isNumeric(seq) Then
            MsgBox("Error: MoveToTop(seq) seq is not a number")
            Exit Function
        ELse
            seq = Cint(seq)
        End If

        If seq < 0 Or seq > mLength Then
            MsgBox("Error: MoveToTop(seq) seq out of bound")
            Exit Function
        End If

        If seq = 0 Then Exit Function

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
    End Function

    Public Function MoveToEnd(seq)
        If Not isNumeric(seq) Then
            MsgBox("Error: MoveToEnd(seq) seq is not a number")
            Exit Function
        ELse
            seq = Cint(seq)
        End If

        If seq < 0 Or seq > mLength Then
            MsgBox("Error: MoveToEnd(seq) seq out of bound")
            Exit Function
        End If

        If seq = 0 Then Exit Function

        Dim i, sValueToBeMove
        If isObject(mArray(seq)) Then
            Set sValueToBeMove = mArray(seq)
            For i = seq To mLength - 1
                Set mArray(i) = mArray(i + 1)
            Next
            Set mArray(mLength) = sValueToBeMove
        Else
            sValueToBeMove = mArray(seq)
            For i = seq To mLength - 1
                mArray(i) = mArray(i + 1)
            Next
            mArray(mLength) = sValueToBeMove
        End If
    End Function

    Public Function IsExistInArray(value)
        If mLength = -1 Then
            IsExistInArray = -1
            Exit Function
        End If

        Dim i
        For i = 0 To mLength
            If StrComp(mArray(i), value) = 0 Then
                IsExistInArray = i
                Exit Function
            End If
        Next
        IsExistInArray = -1
    End Function

    Public Function SortArray()
        If mLength = -1 Then
            'MsgBox("Error: SortArray() mLength <= 0, no need to sort")
            Exit Function
        End If

        Dim i, j
        For i = 0 To mLength - 1
            For j = i + 1 To mLength
                If StrComp(mArray(i), mArray(j)) > 0 Then
                    Dim sTmp : sTmp = mArray(i) : mArray(i) = mArray(j) : mArray(j) = sTmp
                End If
            Next
        Next
    End Function

    Public Function ToString()
        If mLength <> -1 Then
            Dim i, sTmp
            sTmp = "v(0) = " & mArray(0)
            If mLength > 0 Then
                For i = 1 To mLength
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


Const STATUS_INVALID = 0
Const STATUS_VALID = 1
Const STATUS_WAIT = 2

Class StatusHolder
    Private mValue, mStatus, mInvalidMsg

    Private Sub Class_Initialize
        mValue = ""
        mStatus = STATUS_WAIT
    End Sub

    Public Property Get Value
        Value = mValue
    End Property

    Public Property Get Status
        Status = mStatus
    End Property

    Public Property Let Status(value)
        mStatus = value
    End Property

    Public Property Get InvalidMsg
        InvalidMsg = mInvalidMsg
    End Property

    Public Property Let InvalidMsg(value)
        mInvalidMsg = value
    End Property

    Public Sub Reset()
        mValue = ""
        mStatus = STATUS_WAIT
    End Sub

    Public Sub SetValue(status, value)
        mStatus = status
        mValue = value
    End Sub

    Public Sub SetStatusAndMsg(status, msg, showMsg)
        mStatus = status
        mInvalidMsg = msg
        If showMsg Then MsgBox(mInvalidMsg)
    End Sub

    Public Function checkInvalidAndShowMsg()
        If mStatus = STATUS_INVALID Then
            MsgBox(mInvalidMsg)
            checkInvalidAndShowMsg = True
            Exit Function
        End If
        checkInvalidAndShowMsg = False
    End Function

    Public Function checkStatusAndDoSomething(waitFun, invalidFun)
        If mStatus = STATUS_WAIT Then
            Call Execute(waitFun)
            If mStatus = STATUS_INVALID Then
                Call Execute(invalidFun)
                checkStatusAndDoSomething = True
            End If
        ElseIf checkInvalidAndShowMsg() Then
            checkStatusAndDoSomething = True
        End If
    End Function
End Class

Const SEARCH_FILE = 0
Const SEARCH_FOLDER = 1
Const SEARCH_ROOT = 0
Const SEARCH_SUB = 1
Const SEARCH_WHOLE_NAME = 0
Const SEARCH_PART_NAME = 1
Const SEARCH_ONE = 0
Const SEARCH_ALL = 1
Const SEARCH_RETURN_PATH = 0
Const SEARCH_RETURN_NAME = 1

Function searchFolder(pRootFolder, str, searchType, searchWhere, searchMode, searchTimes, returnType)
    If Not oFso.FolderExists(pRootFolder) Then searchFolder = "" : Exit Function
    If searchMode = SEARCH_WHOLE_NAME Then searchTimes = SEARCH_ONE

    Dim oRootFolder : Set oRootFolder = oFso.GetFolder(pRootFolder)

    Dim Folder, sTmp
    Select Case True
        Case searchType = SEARCH_FILE And searchWhere = SEARCH_ROOT
            If searchTimes = SEARCH_ALL Then
                Set searchFolder = startSearch(oRootFolder.Files, pRootFolder, str, searchMode, SEARCH_ALL, returnType)
            Else
                searchFolder = startSearch(oRootFolder.Files, pRootFolder, str, searchMode, SEARCH_ONE, returnType)
            End If

        Case searchType = SEARCH_FOLDER And searchWhere = SEARCH_ROOT
            If searchTimes = SEARCH_ALL Then
                Set searchFolder = startSearch(oRootFolder.SubFolders, pRootFolder, str, searchMode, SEARCH_ALL, returnType)
            Else
                searchFolder = startSearch(oRootFolder.SubFolders, pRootFolder, str, searchMode, SEARCH_ONE, returnType)
            End If

        Case searchType = SEARCH_FILE And searchWhere = SEARCH_SUB
            For Each Folder In oRootFolder.SubFolders
                sTmp = startSearch(Folder.Files, pRootFolder & "\" & Folder.Name, str, searchMode, SEARCH_ONE, returnType)
                If sTmp <> "" Then
                    searchFolder = sTmp
                    Exit Function
                End If
            Next
            searchFolder = ""

        Case searchType = SEARCH_FILE And searchWhere = SEARCH_SUB
            For Each Folder In oRootFolder.SubFolders
                sTmp = startSearch(Folder.SubFolders, pRootFolder & "\" & Folder.Name, str, searchMode, SEARCH_ONE, returnType)
                If sTmp <> "" Then
                    searchFolder = sTmp
                    Exit Function
                End If
            Next
            searchFolder = ""
    End Select
End Function

        Function startSearch(oAll, pRootFolder, str, searchMode, searchTimes, returnType)
            Dim oSingle

            If searchTimes = SEARCH_ALL Then
                Dim vaStr : Set vaStr = New VariableArray
                For Each oSingle In oAll
                    If checkSearchName(oSingle.Name, str, searchMode) Then
                        If returnType = SEARCH_RETURN_PATH Then
                            vaStr.Append(pRootFolder & "\" & oSingle.Name)
                        Else
                            vaStr.Append(oSingle.Name)
                        End If
                    End If
                Next
                Set startSearch = vaStr
                Exit Function
            Else
                For Each oSingle In oAll
                    If checkSearchName(oSingle.Name, str, searchMode) Then
                        If returnType = SEARCH_RETURN_PATH Then
                            startSearch = pRootFolder & "\" & oSingle.Name
                        Else
                            startSearch = oSingle.Name
                        End If
                        Exit Function
                    End If
                Next
            End If
            startSearch = ""
        End Function

        Function checkSearchName(name, str, searchMode)
            If searchMode = SEARCH_WHOLE_NAME Then
                If StrComp(name ,str) = 0 Then
                    checkSearchName = True
                Else
                    checkSearchName = False
                End If
            ELseIf searchMode = SEARCH_PART_NAME Then
                If InStr(name ,str) > 0 Then
                    checkSearchName = True
                Else
                    checkSearchName = False
                End If
            End If
        End Function

Function getElementValue(elementId)
    getElementValue = document.getElementById(elementId).value
End Function

Sub setElementValue(elementId, str)
    document.getElementById(elementId).value = str
End Sub