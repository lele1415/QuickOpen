Option Explicit

Const STR_JOYA_O1 = "mediateksample"
Const STR_JOYA_MID = "mid"
Const STR_JOYA_JOYA = "joya_sz"
Const STR_ROCO_MID = "mid"
Const STR_ROCO_ROCO = "ROCO"

Dim vaTargetProduct, vaCustomProject
Dim mPdtPath, mPjtPath
Dim mRocoStr, mJoyaStr
Dim mProductName, mProductIndex
Dim mLoadOptInAllPrj
Dim idTimer



Sub onWorkChange()
    Call mIp.onWorkChange()
End Sub

Sub onFirmwareChange()
    Call mIp.onFirmwareChange()
End Sub

Sub onRequirementsChange()
    Call mIp.onRequirementsChange()
End Sub

Sub onZentaoChange()
    Call mIp.onZentaoChange()
End Sub

Sub onSdkPathChange()
    Call mIp.clearWorkInfos()
    If mIp.onSdkChange() Then
        Call updateProductList()
    End If
End Sub

Sub onProductChange()
	Call mIp.clearWorkInfos()
    If mIp.onProductChange() Then
        Call updateProjectList()
    End If
End Sub

Sub onProjectChange()
	Call mIp.clearWorkInfos()
    If mIp.onProjectChange() Then
        Call createWorkName()
    End If
End Sub

Sub updateProductList()
	Call freezeAllInput()
	idTimer = window.setTimeout( _
			"Call findProduct()", 0, "VBScript")
End Sub

Sub findProduct()
	window.clearTimeout(idTimer)
	If Not oFso.FolderExists(mIp.Infos.WeibuSdkPath) Then
		MsgBox "Not found: " & mIp.Infos.WeibuSdkPath
		Exit Sub
	End If
	Set vaTargetProduct = searchFolder(mIp.Infos.WeibuSdkPath, "_bsp", _
		    SEARCH_FOLDER, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ALL, SEARCH_RETURN_NAME)
	If vaTargetProduct.Bound = -1 Then MsgBox("No product found!") : Exit Sub

	Call vaTargetProduct.SortArray()
	Call mProductList.addList(vaTargetProduct)

    If vaTargetProduct.GetIndexIfExist(mIp.Infos.Product) = -1 Then
        mIp.Product = vaTargetProduct.V(0)
    End If

	Call updateProjectList()
End Sub


Sub updateProjectList()
	idTimer = window.setTimeout( _
			"findProject()", 0, "VBScript")
End Sub

Sub findProject()
	window.clearTimeout(idTimer)
	Set vaCustomProject = searchFolder(mIp.Infos.ProductSdkPath, "", _
		    SEARCH_FOLDER, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ALL, SEARCH_RETURN_NAME)
	If vaCustomProject.Bound = -1 Then MsgBox("No project found!") : Exit Sub

	Call vaCustomProject.SortArray()
	Call mProjectList.addList(vaCustomProject)

	If vaCustomProject.GetIndexIfExist(mIp.Infos.Project) = -1 Then
        mIp.Project = vaCustomProject.V(0)
    End If

    Call createWorkName()

	Call unfreezeAllInput()
End Sub

Sub createWorkName()
	If Not checkProjectInfosExist() Then
        mIp.Work = getProjectSimpleName() & " " & getSdkSimpleName()
    End If
End Sub

Function checkProjectInfosExist()
	Dim i, oInfos
	For i = 0 To vaWorksInfo.Bound
		Set oInfos = vaWorksInfo.V(i)
		If oInfos.Sdk = mIp.Infos.Sdk And _
	    	    oInfos.Product = mIp.Infos.Product And _
	    	    oInfos.Project = mIp.Infos.Project Then
	    	mIp.Work = oInfos.Work
	    	checkProjectInfosExist = True
	    	Exit Function
		End If
	Next
	checkProjectInfosExist = False
End Function

Function getProjectSimpleName()
	Dim str
	str = Replace(mIp.Infos.Project, "-MMI", "")
	str = Replace(str, "_MMI", "")
	If InStr(str, "-") > 0 Then
	    str = Replace(str, Left(str, InStr(str, "-")), "")
	ElseIf InStr(str, "_") > 0 Then
	    str = Replace(str, Left(str, InStr(str, "_")), "")
	End If
	getProjectSimpleName = str
End Function

Function getSdkSimpleName()
	Dim str
	str = Replace(mIp.Infos.Sdk, "/", "\")
	If InStr(str, "\alps") > 0 Then str = Left(str, InStrRev(str, "\alps") - 1)
	str = Replace(str, Left(str, InStrRev(str, "\")), "")
	getSdkSimpleName = str
End Function
