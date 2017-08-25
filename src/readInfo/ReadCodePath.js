function showAndHide(listId, types)
{ 
    //var listIdArray = listId.split("$")
    //for (var i = 0; i < listIdArray.length; i++) {
        var Layer=window.document.getElementById(listId); 
        switch (types) { 
            case "show": 
                Layer.style.display="block"; 
                break; 
            case "hide": 
                Layer.style.display="none";
                break; 
        }
    //}
}

function addAfterLiForCodePath(str, inputId, listId, ulId)
{
    var obj = document.getElementById(ulId);
    var li = document.createElement("li");
    li.onmousedown = function(){setListValue(inputId, listId, str)};
    li.innerHTML = str;
    li.style.fontSize = "x-small";
    
    obj.appendChild(li);
}
