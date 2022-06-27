Option Explicit

Const WINDOW_WIDTH = 460
Const WINDOW_HEIGHT = 850
Sub Window_OnLoad
    Dim ScreenWidth : ScreenWidth = CreateObject("HtmlFile").ParentWindow.Screen.AvailWidth
    Dim ScreenHeight : ScreenHeight = CreateObject("HtmlFile").ParentWindow.Screen.AvailHeight
    Window.MoveTo ScreenWidth - WINDOW_WIDTH ,(ScreenHeight - WINDOW_HEIGHT) \ 2
    Window.ResizeTo WINDOW_WIDTH, WINDOW_HEIGHT
End Sub

Call readConfigText(pConfigText)
Call readSdkPathText(pSdkPathText)
Call readWorksInfoText()

Call mSdkPathList.addList(vaAndroidVer)
Call addOpenPathList()
Call addOutFileList()

Call applyLastWorkInfo()
