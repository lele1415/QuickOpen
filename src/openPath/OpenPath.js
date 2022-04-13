var mParentId, mButtonId, mInputId, mValuePartId, mDivId, mUlId

function setParentIds(parentId, buttonId, inputId, valuePartId) {
    mParentId = parentId;
    mButtonId = buttonId;
    mInputId = inputId;
    mValuePartId = valuePartId;
}

function setListIds(divId, ulId) {
    mDivId = divId;
    mUlId = ulId;
}

function showList(listId) {
    setCurrentListId(listId)
    var layer=window.document.getElementById(listId); 
    layer.style.display="block";
}

function hideList(listId) {
    var layer=window.document.getElementById(listId); 
    layer.style.display="none";
}

function addDirectoryList(str, connectDivId)
{
    var obj = document.getElementById(mUlId);
    var li = document.createElement("li");
    li.onmousedown = function(){onDirectoryListClick(mDivId, connectDivId)};
    li.innerHTML = str;
    li.style.fontSize = "x-small";
    
    obj.appendChild(li);
}

function addList(str)
{
    var obj = document.getElementById(mUlId);
    var li = document.createElement("li");
    li.onmousedown = function(){onListClick(str, mDivId, mInputId)};
    li.innerHTML = str;
    li.style.fontSize = "x-small";
    
    obj.appendChild(li);
}

function addUL() {
    var divParent = document.getElementById(mParentId);

    var div1 = document.createElement("div");
    div1.className = "Menu_codePath";
    div1.id = mDivId;
    divParent.appendChild(div1);

    var div2 = document.createElement("div");
    div2.className = "Menu2";
    div1.appendChild(div2);

    var ul = document.createElement("ul");
    ul.id = mUlId;
    div2.appendChild(ul);
}

function onDirectoryListClick(divId, connectDivId) {
    hideList(divId);
    showList(connectDivId);

}

function onListClick(str, divId, inputId) {
    hideList(divId);
    setInputValueFromList(str, inputId);
}
