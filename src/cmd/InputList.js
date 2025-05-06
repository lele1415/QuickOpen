function setInputClickFun(parentId, inputId, divId) {
    var input = document.getElementById(inputId);
    //if (input.onclick == undefined) {
        input.onclick = function(){toggleListDiv(parentId, divId)};
    //}
}

function removeLi(ulId) {
    var i = parentNode_getChildNodesLength(ulId);
    if (i > 0) {
        for(var j=0; j<i; j++) {
            parentNode_removeChild(ulId, 0);
        }   
    }
}

function addListUL(parentId, divId, ulId) {
    var parent = document.getElementById(parentId);

    var div1 = document.createElement("div");
    div1.className = "Menu_sdkPath";
    div1.id = divId;
    parent.appendChild(div1);

    var div2 = document.createElement("div");
    div2.className = "Menu2";
    div1.appendChild(div2);

    var ul = document.createElement("ul");
    ul.id = ulId;
    div2.appendChild(ul);
}

function addListLi(parentId, inputId, divId, ulId, str, setValue)
{
    var li = document.createElement("li");
    li.onmousedown = function(){onListLiClick(parentId, inputId, divId, str, setValue)};
    li.innerHTML = str;
    li.style.fontSize = "x-small";
    li.tabIndex = 1;
    
    var ul = document.getElementById(ulId);
    ul.appendChild(li);
}

function onListLiClick(parentId, inputId, divId, str, setValue) {
    hideListDiv(parentId, divId);

    if (!setValue) {
        onInputListClick(divId, str);
        return;
    }

    var input = document.getElementById(inputId);
    if (input.value != str) {
        input.value = str;
    }

    if (input.onchange != undefined) {
        input.onchange();
    }

    input.focus();
}

function addListDirectoryLi(parentId, dirDivId, dirUlId, listDivId, str)
{
    var ul = document.getElementById(dirUlId);
    var li = document.createElement("li");
    li.onmousedown = function(){onDirectoryLiClick(parentId, dirDivId, listDivId)};
    li.innerHTML = str;
    li.style.fontSize = "x-small";
    
    ul.appendChild(li);
}

function onDirectoryLiClick(parentId, dirDivId, listDivId) {
    hideListDiv(parentId, dirDivId)
    showListDiv(parentId, listDivId);
}

function isDivShowing(divId) {
    var layer = window.document.getElementById(divId);
    return layer != null && layer.style.display == "block";
}

function showListDiv(parentId, divId) {
    var parent = document.getElementById(parentId);
    //parent.onmouseleave = function(){hideListDiv(parentId, divId)};
    parent.style.display="block";

    var layer=window.document.getElementById(divId);
    layer.style.display="block";
}

function hideListDiv(parentId, divId) {
    var parent = document.getElementById(parentId);
    parent.style.display="none";
    var layer=window.document.getElementById(divId); 
    layer.style.display="none";
}

function toggleListDiv(parentId, divId) {
    var layer=window.document.getElementById(divId);
    if (layer.style.display == "block") {
        hideListDiv(parentId, divId);
    } else {
        showListDiv(parentId, divId);
    }
}

function changeLiFocusDown(ulId) {
    var i = parentNode_getChildNodesLength(ulId);
    var parentNode = document.getElementById(ulId);
    var focusIndex = -1;
    if (i > 0) {
        for(var j=0; j<i; j++) {
            if (parentNode.childNodes[j].style.background == "#66d9ef") {
                parentNode.childNodes[j].blur();
                focusIndex = j;
                break;
            }
        }
        if (focusIndex > -1 && focusIndex < i - 1) {
            parentNode.childNodes[focusIndex + 1].focus();
            return focusIndex + 1;
        } else if (focusIndex == -1) {
            parentNode.childNodes[0].focus();
            return 0;
        } else if (focusIndex == i - 1) {
            parentNode.childNodes[i - 1].focus();
            return i - 1;
        } else {
            return -1;
        }
    }
}

function changeLiFocusUp(ulId) {
    var i = parentNode_getChildNodesLength(ulId);
    var parentNode = document.getElementById(ulId);
    var focusIndex = -1;
    if (i > 0) {
        for(var j=0; j<i; j++) {
            if (parentNode.childNodes[j].style.background == "#66d9ef") {
                parentNode.childNodes[j].blur();
                focusIndex = j;
                break;
            }
        }
        if (focusIndex > 0) {
            parentNode.childNodes[focusIndex - 1].focus();
            return focusIndex - 1;
        } else if (focusIndex == 0) {
            parentNode.childNodes[0].focus();
            return 0;
        } else {
            return -1;
        }
    }
}

function clickListLi(ulId, index) {
    var parentNode = document.getElementById(ulId);
    parentNode.childNodes[index].onmousedown();
}
