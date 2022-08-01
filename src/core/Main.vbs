Option Explicit

Const WINDOW_WIDTH = 460
Const WINDOW_HEIGHT = 850
Sub Window_OnLoad
    Dim ScreenWidth : ScreenWidth = CreateObject("HtmlFile").ParentWindow.Screen.AvailWidth
    Dim ScreenHeight : ScreenHeight = CreateObject("HtmlFile").ParentWindow.Screen.AvailHeight
    Window.MoveTo ScreenWidth - WINDOW_WIDTH ,(ScreenHeight - WINDOW_HEIGHT) \ 2
    Window.ResizeTo WINDOW_WIDTH, WINDOW_HEIGHT
    
    Call runInitFuns()
End Sub

Sub runInitFuns()
    Call readConfigText()
    Call readWorksInfoText()

    Call addSdkPathList()
    Call addOpenPathList()
    Call addOutFileList()

    Call applyLastWorkInfo()
    Call checkConfigInfos()
End Sub
