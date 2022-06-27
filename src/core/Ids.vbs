Option Explicit

Const FOR_READING = 1
Const FOR_APPENDING = 8

Dim oWs, oFso
Set oWs=CreateObject("wscript.shell")
Set oFso=CreateObject("Scripting.FileSystemObject")

Dim pConfigText : pConfigText = oWs.CurrentDirectory & "\res\config.ini"
Dim pSdkPathText : pSdkPathText = oWs.CurrentDirectory & "\res\sdk.ini"



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
Const ID_INPUT_FIRMWARE = "input_firmware"
Const ID_INPUT_REQUIREMENTS = "input_requirements"
Const ID_INPUT_ZENTAO = "input_zentao"

Const ID_PARENT_OPEN_PATH = "parent_open_path"
Const ID_INPUT_OPEN_PATH = "input_open_path"

Const ID_DIV_OPEN_PATH_DIRECTORY = "div_open_path_directory"
Const ID_UL_OPEN_PATH_DIRECTORY = "ul_open_path_directory"

Const ID_DIV_OPEN_PATH_ = "div_open_path_"
Const ID_UL_OPEN_PATH_ = "ul_open_path_"

Const ID_PARENT_OUT_BUTTON = "parent_out_button"

Const ID_PARENT_OPEN_BUTTON = "parent_open_button"

Const ID_PARENT_FILE_BUTTON = "parent_file_button"

Const ID_SELECT_FOR_COMPARE = "select_for_compare"
Const ID_COMPARE_TO = "compare_to"



Function getWorkInputId()
    getWorkInputId = ID_WORK_NAME
End Function

Function getFirmwareInputId()
    getFirmwareInputId = ID_INPUT_FIRMWARE
End Function

Function getRequirementsInputId()
    getRequirementsInputId = ID_INPUT_REQUIREMENTS
End Function

Function getZentaoInputId()
    getZentaoInputId = ID_INPUT_ZENTAO
End Function

Function getSdkPathParentId()
    getSdkPathParentId = ID_PARENT_SDK_PATH
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
Function getOpenPathParentId()
    getOpenPathParentId = ID_PARENT_OPEN_PATH
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

Function getOutButtonParentId()
    getOutButtonParentId = ID_PARENT_OUT_BUTTON
End Function

Function getOpenButtonParentId()
    getOpenButtonParentId = ID_PARENT_OPEN_BUTTON
End Function

Function getFileButtonParentId()
    getFileButtonParentId = ID_PARENT_FILE_BUTTON
End Function
