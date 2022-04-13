


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



Sub findProduct()
	Call freezeAllInput()
	idTimer = window.setTimeout( _
			"Call getAllWeibuProduct()", 0, "VBScript")
End Sub

Sub getAllWeibuProduct()
	window.clearTimeout(idTimer)
	Set vaTargetProduct = searchFolder(getWeibuPath(), "_bsp", _
		    SEARCH_FOLDER, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ALL, SEARCH_RETURN_NAME)
	If vaTargetProduct.Bound = -1 Then MsgBox("No product found!") : Exit Sub

	Call vaTargetProduct.SortArray()
	Call addProductLi()

	Call setProduct(vaTargetProduct.V(0))

	Call findProject()
End Sub

Sub addProductLi()
	Call removeLi(getProductULId())
	Call setListParentAndInputIds(getProductParentId(), getProductInputId())
	Call setListDivIds(getProductDivId(), getProductULId())
    Call addListUL()

	Dim i : For i = 0 To vaTargetProduct.Bound
		Call addListLi(vaTargetProduct.V(i))
	Next
End Sub


Sub findProject()
	idTimer = window.setTimeout( _
			"getAllWeibuProject()", 0, "VBScript")
End Sub

Sub getAllWeibuProject()
	window.clearTimeout(idTimer)
	Set vaCustomProject = searchFolder(getProductPath(), "", _
		    SEARCH_FOLDER, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ALL, SEARCH_RETURN_NAME)
	If vaCustomProject.Bound = -1 Then MsgBox("No project found!") : Exit Sub

	Call vaCustomProject.SortArray()
	'MsgBox(vaCustomProject.ToString())

	Call setProject(vaCustomProject.V(0))
	Call addProjectLi()

	Call unfreezeAllInput()
End Sub

Sub addProjectLi()
	Call removeLi(getProjectULId())
	Call setListParentAndInputIds(getProjectParentId(), getProjectInputId())
	Call setListDivIds(getProjectDivId(), getProjectULId())
    Call addListUL()

	Dim i : For i = 0 To vaCustomProject.Bound
		Call addListLi(vaCustomProject.V(i))
	Next
End Sub
