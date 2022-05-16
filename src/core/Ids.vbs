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

Const ID_WORK_NAME = "work_name"
Const ID_DIV_SHORTCUT = "div_shortcut"

Const ID_PARENT_OPEN_PATH = "parent_open_path"
Const ID_INPUT_OPEN_PATH = "input_open_path"

Const ID_DIV_OPEN_PATH_DIRECTORY = "div_open_path_directory"
Const ID_UL_OPEN_PATH_DIRECTORY = "ul_open_path_directory"

Const ID_DIV_OPEN_PATH_ = "div_open_path_"
Const ID_UL_OPEN_PATH_ = "ul_open_path_"

Function getParentSdkPathId()
    getParentSdkPathId = ID_PARENT_SDK_PATH
End Function

Function getWorkInputId()
    getWorkInputId = ID_WORK_NAME
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


Dim mOpenPathInput : Set mOpenPathInput = (New InputText)(getOpenPathInputId())
Dim mIp : Set mIp = New ProjectInputs
Dim mIf : Set mIf = New ProjectInfos
