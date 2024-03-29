Option Explicit

Const STR_JOYA_O1 = "mediateksample"
Const STR_JOYA_MID = "mid"
Const STR_JOYA_JOYA = "joya_sz"
Const STR_ROCO_MID = "mid"
Const STR_ROCO_ROCO = "ROCO"

Dim vaTargetProduct, vaCustomProject, vaSysProduct
Dim mPdtPath, mPjtPath
Dim mRocoStr, mJoyaStr
Dim mProductName, mProductIndex
Dim mLoadOptInAllPrj
Dim idTimer



Sub onInputWorkChange()
    Call mIp.onWorkChange()
End Sub

Sub onInputFirmwareChange()
    Call mIp.onFirmwareChange()
End Sub

Sub onInputRequirementsChange()
    Call mIp.onRequirementsChange()
End Sub

Sub onInputZentaoChange()
    Call mIp.onZentaoChange()
End Sub

Sub onInputSdkChange()
    Call mIp.onSdkChange()
End Sub

Sub onInputProductChange()
    Call mIp.onProductChange()
End Sub

Sub onInputProjectChange()
    Call mIp.onProjectChange()
End Sub

Sub updateProductList()
	Call freezeAllInput()
	idTimer = window.setTimeout( _
			"Call findProduct()", 0, "VBScript")
End Sub

Sub findProduct()
	window.clearTimeout(idTimer)
	Call getProductList()
	Call mProductList.addList(vaTargetProduct)

    If vaTargetProduct.GetIndexIfExist(mIp.Infos.Product) = -1 Then
        mIp.Product = vaTargetProduct.V(0)
    End If

	Call updateProjectList()
End Sub

Sub getProductList()
    If Not isFolderExists("weibu") Then
		MsgBox "Not found: weibu/"
		Exit Sub
	End If

	Set vaTargetProduct = searchFolder("weibu", "_", _
			SEARCH_FOLDER, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ALL, SEARCH_RETURN_NAME)

	If isT0Sdk() Then
	    Call setT0SdkSys()
	    Set vaSysProduct = searchFolder("weibu", "mssi_", _
		        SEARCH_FOLDER, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ALL, SEARCH_RETURN_NAME)
		Call setT0SdkVnd()
		Call vaSysProduct.SortArray()
	End If

	If vaTargetProduct.Bound = -1 Then MsgBox("No product found!") : Exit Sub

	Call vaTargetProduct.SortArray()
End Sub

Sub updateProjectList()
	idTimer = window.setTimeout( _
			"findProject()", 0, "VBScript")
End Sub

Sub findProject()
	window.clearTimeout(idTimer)
	Set vaCustomProject = searchFolder(mIp.Infos.ProductPath, "", _
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
		If oInfos.isSameProject(mIp.Infos) Then
	    	mIp.Work = oInfos.Work
	    	checkProjectInfosExist = True
	    	Exit Function
		End If
	Next
	checkProjectInfosExist = False
End Function

Function getProjectSimpleName()
	Dim str
	If isT0Sdk() Then
		str = Replace(mIp.Infos.SysProject, "-MMI", "")
	Else
		str = Replace(mIp.Infos.Project, "-MMI", "")
	End If
	str = Replace(str, "_MMI", "")
	If InStr(mIp.Infos.Sdk, "_r") > 0 And InStr(str, "-") > 0 Then
	    str = Replace(str, Left(str, InStr(str, "-")), "")
	ElseIf InStr(str, "_") > 0 Then
	    str = Replace(str, Left(str, InStr(str, "_")), "")
	End If
	getProjectSimpleName = str
End Function

Function getSdkSimpleName()
	Dim str
	str = mIp.Infos.Sdk
	If InStr(str, "alps") > 0 Then str = getParentPath(str)
	getSdkSimpleName = str
End Function
