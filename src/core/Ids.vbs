Const ID_PARENT_SDK_PATH = "parent_sdk_path"
Const ID_INPUT_SDK_PATH = "input_sdk_path"

Const ID_DIV_SDK_PATH_DIRECTORY = "div_sdk_path_directory"
Const ID_UL_SDK_PATH_DIRECTORY = "ul_sdk_path_directory"
Const ID_DIV_SDK_PATH_ = "div_sdk_path_"
Const ID_UL_SDK_PATH_ = "ul_sdk_path_"

Const ID_PARENT_PRODUCT = "parent_product"
Const ID_INPUT_PRODUCT = "input_product"
Const ID_DIV_PRODUCT = "list_target_product"
Const ID_UL_PRODUCT = "ul_target_product"

Const ID_PARENT_PROJECT = "parent_project"
Const ID_INPUT_PROJECT = "input_project"
Const ID_DIV_PROJECT = "list_custom_project"
Const ID_UL_PROJECT = "ul_custom_project"

Const ID_PARENT_OPEN_PATH = "parent_open_path"
Const ID_INPUT_OPEN_PATH = "input_open_path"

Const ID_DIV_OPEN_PATH_DIRECTORY = "div_open_path_directory"
Const ID_UL_OPEN_PATH_DIRECTORY = "ul_open_path_directory"

Const ID_DIV_OPEN_PATH_ = "div_open_path_"
Const ID_UL_OPEN_PATH_ = "ul_open_path_"

Function getParentSdkPathId()
    getParentSdkPathId = ID_PARENT_SDK_PATH
End Function

Function getSdkPathInputId()
    getSdkPathInputId = ID_INPUT_SDK_PATH
End Function

Function getSdkPathDirectoryDivId()
    getSdkPathDirectoryDivId = ID_DIV_SDK_PATH_DIRECTORY
End Function

Function getSdkPathDirectoryULId()
    getSdkPathDirectoryULId = ID_UL_SDK_PATH_DIRECTORY
End Function

Function getSdkPathDivId()
    getSdkPathDivId = ID_DIV_SDK_PATH_
End Function

Function getSdkPathULId()
    getSdkPathULId = ID_UL_SDK_PATH_
End Function

'Product
Function getProductParentId()
    getProductParentId = ID_PARENT_PRODUCT
End Function

Function getProductInputId()
    getProductInputId = ID_INPUT_PRODUCT
End Function

Function getProductDivId()
    getProductDivId = ID_DIV_PRODUCT
End Function

Function getProductULId()
    getProductULId = ID_UL_PRODUCT
End Function

'Project
Function getProjectParentId()
    getProjectParentId = ID_PARENT_PROJECT
End Function

Function getProjectInputId()
    getProjectInputId = ID_INPUT_PROJECT
End Function

Function getProjectDivId()
    getProjectDivId = ID_DIV_PROJECT
End Function

Function getProjectULId()
    getProjectULId = ID_UL_PROJECT
End Function

'Open path
Function getParentOpenPathId()
    getParentOpenPathId = ID_PARENT_OPEN_PATH
End Function

Function getOpenPathInputId()
    getOpenPathInputId = ID_INPUT_OPEN_PATH
End Function

Function getOpenPathDirectoryDivId()
    getOpenPathDirectoryDivId = ID_DIV_OPEN_PATH_DIRECTORY
End Function

Function getOpenPathDirectoryULId()
    getOpenPathDirectoryULId = ID_UL_OPEN_PATH_DIRECTORY
End Function

Function getOpenPathDivId()
    getOpenPathDivId = ID_DIV_OPEN_PATH_
End Function

Function getOpenPathULId()
    getOpenPathULId = ID_UL_OPEN_PATH_
End Function


Dim mSdkPathInput : Set mSdkPathInput = (New InputText)(getSdkPathInputId())
Dim mProductInput : Set mProductInput = (New InputWithOneLayerList)(getProductInputId(), getProductDivId())
Dim mProjectInput : Set mProjectInput = (New InputWithOneLayerList)(getProjectInputId(), getProjectDivId())
Dim mOpenPathInput : Set mOpenPathInput = (New InputText)(getOpenPathInputId())

'SDK path
Sub onSdkPathChange()
    Dim path : path = getSdkPath()
    If oFso.FolderExists(path) Then
        Call setSdkPath(path)
        Call findProduct()
    Else
        Call invalidSdkPath(path)
    End If
End Sub

Function getSdkPath()
    getSdkPath = mSdkPathInput.text
End Function

Function getWeibuPath()
    getWeibuPath = getSdkPath() & "/weibu"
End Function

Sub setSdkPath(path)
    If oFso.FolderExists(path) Then
        mSdkPathInput.setText(path)
    Else
        Call invalidSdkPath(path)
    End If
End Sub

Sub invalidSdkPath(path)
    mSdkPathInput.setText("")
    mProductInput.setText("")
    mProjectInput.setText("")
    MsgBox("Code path not exist!" & VbLf & path)
End Sub

'Product
Sub onProductChange()
    Dim path : path = getProductPath()
    If oFso.FolderExists(path) Then
        Call findProject()
    Else
        Call invalidProduct(path)
    End If
End Sub

Function getProduct()
    getProduct = mProductInput.text
End Function

Function getProductPath()
    getProductPath = getSdkPath() & "/weibu/" & mProductInput.text
End Function

Function getProductPathWithoutSdk()
    getProductPathWithoutSdk = "weibu/" & mProductInput.text
End Function

Sub setProduct(product)
    Dim path : path = getSdkPath() & "/weibu/" & product
    If oFso.FolderExists(path) Then
        mProductInput.setText(product)
    Else
        Call invalidProduct(path)
    End If
End Sub

Sub invalidProduct(path)
    mProductInput.setText("")
    mProjectInput.setText("")
    MsgBox("Product folder not exist!" & VbLf & path)
End Sub

'Project
Sub onProjectChange()
    Dim path : path = getProjectPath()
    If oFso.FolderExists(path) Then
        Call setProject(mProjectInput.Text)
    Else
        Call invalidProject(path)
    End If
End Sub

Function getProject()
    getProject = mProjectInput.text
End Function

Function getProjectPath()
    If checkProjectAlps() Then
        getProjectPath = getProductPath() & "/" & mProjectInput.text & "/alps"
    Else
        getProjectPath = getProductPath() & "/" & mProjectInput.text
    End If
End Function

Function getProjectPathWithoutSdk()
    If checkProjectAlps() Then
        getProjectPathWithoutSdk = getProductPathWithoutSdk() & "/" & mProjectInput.text & "/alps"
    Else
        getProjectPathWithoutSdk = getProductPathWithoutSdk() & "/" & mProjectInput.text
    End If
End Function

Function checkProjectAlps()
    If oFso.FolderExists(getProductPath() & "/" & mProjectInput.text & "/alps") Then
        checkProjectAlps = True
    Else
        checkProjectAlps = False
    End If
End Function

Sub setProject(project)

    Dim path : path = getSdkPath() & "/weibu/" & getProduct() & "/" & project
    If oFso.FolderExists(path) Then
        mProjectInput.setText(project)
    Else
        Call invalidProject(path)
    End If
End Sub

Sub invalidProject(path)
    mProjectInput.setText("")
    MsgBox("Project folder not exist!" & VbLf & path)
End Sub

'Open path
Sub onOpenPathChange()
    Call replaceOpenPath()
End Sub

Function getOpenPath()
    getOpenPath = mOpenPathInput.text
End Function

Sub setOpenPath(path)
    If oFso.FolderExists(path) Then
        mOpenPathInput.setText(path)
    Else
        mOpenPathInput.setText("")
        MsgBox("Path not exist!" & VbLf & path)
    End If
End Sub

'Out Path
Function getOutPath()
    getOutPath = getSdkPath() & "\out\target\product\" & getProduct()
End Function
