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

function resetInputOnClick(inputId) {
    var input = document.getElementById(inputId);
    input.onclick = function(){onInputElementClick(inputId)};
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

function addListLi(inputId, divId, ulId, str, setValue)
{
    var li = document.createElement("li");
    li.onmousedown = function(){onListLiClick(inputId, divId, str, setValue)};
    li.innerHTML = str;
    li.style.fontSize = "x-small";
    
    var ul = document.getElementById(ulId);
    ul.appendChild(li);
}

function onListLiClick(inputId, divId, str, setValue) {
    hideListDiv(divId);

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

function onDirectoryLiClick(parentId, directoryDivId, secondDivId) {
    hideListDiv(directoryDivId)
    showListDiv(parentId, secondDivId);
}

function showListDiv(parentId, divId) {
    var parent = document.getElementById(parentId);
    parent.onmouseleave = function(){hideListDiv(divId)};

    var layer=window.document.getElementById(divId);
    layer.style.display="block";
}

function hideListDiv(divId) {
    var layer=window.document.getElementById(divId); 
    layer.style.display="none";
}

function toggleListDiv(parentId, divId) {
    var layer=window.document.getElementById(divId);
    if (layer.style.display == "block") {
        hideListDiv(divId);
    } else {
        showListDiv(parentId, divId);
    }
}
