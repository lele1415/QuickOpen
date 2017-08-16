Const ID_OPEN_PATH_L1 = "open_path_l1"
Const ID_OPEN_PATH_KK = "open_path_kk"

Dim aIdCodePath : aIdCodePath = Array(ID_OPEN_PATH_L1, ID_OPEN_PATH_KK)
Function openCodePath(where)
	oWs.Run "explorer.exe " & getElementValue(aIdCodePath(where))
End Function