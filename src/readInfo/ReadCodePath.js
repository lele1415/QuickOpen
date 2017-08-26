var mCodePathList = "";

function showOrHideCodePathList(listId, types)
{ 
    //var listIdArray = listId.split("$")
    //for (var i = 0; i < listIdArray.length; i++) {
        var Layer=window.document.getElementById(listId); 
        switch (types) { 
            case "show": 
                if (mCodePathList == "") {
                    //alert("show listId="+listId);
                    Layer.style.display="block"; 
                    mCodePathList = listId;
                }
                break;
            case "hide": 
                //alert("hide listId="+listId);
                Layer.style.display="none";
                if (mCodePathList == listId) {
                    mCodePathList = "";
                }
                break; 
        }
    //}
}

function HideCodePathList() {
    //alert("HideOpenedList mCodePathList="+mCodePathList);
    if (mCodePathList != "") {
        showOrHideCodePathList(mCodePathList, "hide");
    }
}

function addAfterLiForCodePath(str, inputId, listId, ulId)
{
    var obj = document.getElementById(ulId);
    var li = document.createElement("li");
    li.onmousedown = function(){setListValueForCodePath(inputId, listId, str)};
    li.innerHTML = str;
    li.style.fontSize = "x-small";
    
    obj.appendChild(li);
}
