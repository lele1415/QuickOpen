var mOpenPathList = "";

function showOrHideOpenPathList(listId, types)
{ 
    //var listIdArray = listId.split("$")
    //for (var i = 0; i < listIdArray.length; i++) {
        var Layer=window.document.getElementById(listId); 
        switch (types) { 
            case "show": 
                if (mOpenPathList == "") {
                    //alert("show listId="+listId);
                    Layer.style.display="block"; 
                    mOpenPathList = listId;
                }
                break;
            case "hide": 
                //alert("hide listId="+listId);
                Layer.style.display="none";
                if (mOpenPathList == listId) {
                    mOpenPathList = "";
                }
                break; 
        }
    //}
}

function HideOpenPathList() {
    //alert("HideOpenedList mOpenPathList="+mOpenPathList);
    if (mOpenPathList != "") {
        showOrHideOpenPathList(mOpenPathList, "hide");
    }
}

function addAfterLiForOpenPath(str, inputId, listId, ulId)
{
    var obj = document.getElementById(ulId);
    var li = document.createElement("li");
    li.onmousedown = function(){setListValueForOpenPath(inputId, listId, str)};
    li.innerHTML = str;
    li.style.fontSize = "x-small";
    
    obj.appendChild(li);
}
