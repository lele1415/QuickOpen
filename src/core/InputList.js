var mInputId, mDirectoryDivId, mDivId, mUlId;
var mParent;

function setListParentAndInputIds(parentId, inputId) {
    mInputId = inputId;
    mParent = document.getElementById(parentId);
}

function setListDirectoryDivId(directoryDivId) {
    var input = document.getElementById(mInputId);
    input.onclick = function(){toggleListDiv(directoryDivId)};

    mDirectoryDivId = directoryDivId;
}

function setListDivIds(divId, ulId) {
    var input = document.getElementById(mInputId);
    if (input.onclick == undefined) {
        input.onclick = function(){toggleListDiv(divId)};
    }

    mDivId = divId;
    mUlId = ulId;
}

function removeLi(ulId) {
    var i = parentNode_getChildNodesLength(ulId);
    if (i > 0) {
        for(var j=0; j<i; j++) {
            parentNode_removeChild(ulId, 0);
        }   
    }
}

function addListUL() {
    var div1 = document.createElement("div");
    div1.className = "Menu_sdkPath";
    div1.id = mDivId;
    mParent.appendChild(div1);

    var div2 = document.createElement("div");
    div2.className = "Menu2";
    div1.appendChild(div2);

    var ul = document.createElement("ul");
    ul.id = mUlId;
    div2.appendChild(ul);
}

function addListLi(str)
{
    var li = document.createElement("li");
    var inputId = mInputId;
    var divId = mDivId;
    li.onmousedown = function(){onListLiClick(inputId, divId, str)};
    li.innerHTML = str;
    li.style.fontSize = "x-small";
    
    var ul = document.getElementById(mUlId);
    ul.appendChild(li);
}

function onListLiClick(inputId, divId, str) {
    hideListDiv(divId);
    var input = document.getElementById(inputId);
    if (input.value != str) {
        input.value = str;
    }

    if (input.onchange != undefined) {
        input.onchange();
    }
}

function addListDirectoryLi(str, secondDivId)
{
    var obj = document.getElementById(mUlId);
    var li = document.createElement("li");
    var directoryDivId = mDirectoryDivId;
    li.onmousedown = function(){onDirectoryLiClick(directoryDivId, secondDivId)};
    li.innerHTML = str;
    li.style.fontSize = "x-small";
    
    obj.appendChild(li);
}

function onDirectoryLiClick(directoryDivId, secondDivId) {
    hideListDiv(directoryDivId)
    showListDiv(secondDivId);
}

function showListDiv(divId) {
    var div = document.getElementById(divId);
    div.onmouseleave = function(){hideListDiv(divId)};

    var layer=window.document.getElementById(divId);
    layer.style.display="block";
}

function hideListDiv(divId) {
    var layer=window.document.getElementById(divId); 
    layer.style.display="none";
}

function toggleListDiv(divId) {
    var layer=window.document.getElementById(divId);
    if (layer.style.display == "block") {
        hideListDiv(divId);
    } else {
        showListDiv(divId);
    }
}
