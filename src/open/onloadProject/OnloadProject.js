function addOption(selectId, optionName) 
{
    var option = option_creat(optionName, optionName);
    parentNode_appendChild(selectId, option);
}

function removeAllChild(parentId)
{
    var i = parentNode_getChildNodesLength(parentId);
    if (i>0) {
        for(var j=0; j<i; j++) {
            parentNode_removeChild(parentId, 0);
        }   
    }
}

function showOrHidePrjList(listId, types)
{ 
    var Layer=window.document.getElementById(listId); 
    switch (types) { 
        case "show": 
            Layer.style.display="block"; 
            break;
        case "hide": 
            Layer.style.display="none";
            break; 
    }
}

function addAfterLiForOnloadPrj(str, inputId, listId, ulId)
{
    var obj = document.getElementById(ulId);
    var li = document.createElement("li");
    li.onmousedown = function(){setListValueForOnloadPrj(inputId, listId, str)};
    li.innerHTML = str;
    li.style.fontSize = "x-small";
    
    obj.appendChild(li);
}
