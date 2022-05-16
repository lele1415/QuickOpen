var mSdkPathList = "";

function showOrHideSdkPathList(listId, types)
{ 
    //var listIdArray = listId.split("$")
    //for (var i = 0; i < listIdArray.length; i++) {
        var Layer=window.document.getElementById(listId); 
        switch (types) { 
            case "show": 
                if (mSdkPathList == "") {
                    //alert("show listId="+listId);
                    Layer.style.display="block"; 
                    mSdkPathList = listId;
                }
                break;
            case "hide": 
                //alert("hide listId="+listId);
                Layer.style.display="none";
                if (mSdkPathList == listId) {
                    mSdkPathList = "";
                }
                break; 
        }
    //}
}

function HideSdkPathList() {
    //alert("HideOpenedList mSdkPathList="+mSdkPathList);
    if (mSdkPathList != "") {
        showOrHideSdkPathList(mSdkPathList, "hide");
    }
}

function addAfterLiForSdkPath(str, inputId, listId, ulId)
{
    var obj = document.getElementById(ulId);
    var li = document.createElement("li");
    li.onmousedown = function(){setListValueForSdkPath(inputId, listId, str)};
    li.innerHTML = str;
    li.style.fontSize = "x-small";
    
    obj.appendChild(li);
}
