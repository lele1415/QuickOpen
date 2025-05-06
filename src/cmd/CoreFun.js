function isElementIdExist(id) {
    var element = document.getElementById(id);
    if (element) {
        return true;
    } else {
        return false;
    }
}

function input_creat(inputType,inputId,inputValue,inputOnclickFun,inputSize) {
    var input = document.createElement("input");
    input.type = inputType;
    input.id = inputId;
    if (inputValue) {
        input.value = inputValue;
    }
    if (inputOnclickFun) {
        input.onclick = function(){inputOnclickFun()};
    }
    if (inputSize){
        input.style.width = inputSize;
    }

    return input;
}

function input_changeInfo(inputId,inputValue,inputOnclickFun) {
    var input = document.getElementById(inputId);
    if(inputValue){
        input.value = inputValue;
    }
    if(inputOnclickFun){
        input.onclick = function(){inputOnclickFun()};
 }
}

function select_creat(selectId,selectOnchangeFun,selectWidth) {
    var select = document.createElement("select");
    select.id = selectId;
    if(selectOnchangeFun){
        select.onchange = function(){selectOnchangeFun()};
 }
    if(selectWidth){
        select.style.width = selectWidth;
    }

    return select;
}

function option_creat(optionValue,optionInnerHTML) {
    var option = document.createElement("option");
    option.value = optionValue;
    option.innerHTML = optionInnerHTML;

    return option;
}

function parentNode_appendChild(parentNodeId,node) {
    var parentNode = document.getElementById(parentNodeId);
    parentNode.appendChild(node);
}

function parentNode_insertBefore(newNode,oldNodeId) {
    var oldNode = document.getElementById(oldNodeId);
    oldNode.parentNode.insertBefore(newNode,oldNode);
}

function parentNode_getChildNodesLength(parentNodeId) {
    var parentNode = document.getElementById(parentNodeId);
    if (parentNode && parentNode.childNodes) {
        return parentNode.childNodes.length;
    } else {
        return 0;
    }
}

function parentNode_removeChild(parentNodeId,childNodeCount) {
    var parentNode = document.getElementById(parentNodeId);
    if(parentNode.childNodes[childNodeCount]){
        parentNode.removeChild(parentNode.childNodes[childNodeCount]);
    }
}

function parentNode_removeAllChilds(parentNodeId) {
    var pathLength = parentNode_getChildNodesLength(parentNodeId);
    for (var i = pathLength - 1; i > - 1; i--) {
        parentNode_removeChild(parentNodeId, i);
    }
}

function node_removeNode(nodeId) {
    var node = document.getElementById(nodeId);
    if(node){
        node.removeNode(true);
    }
}

function element_getValue(elementId) {
    var elementValue = document.getElementById(elementId).value;
    return elementValue;
}

function element_setValue(elementId,elementValue) {
    document.getElementById(elementId).value = elementValue;
}

function element_isChecked(elementId) {
    var isChecked = document.getElementById(elementId).checked;
    return isChecked;
}

function hideElement(elementId) {
    document.getElementById(elementId).style.display="none";
}

function showElement(elementId) {
    document.getElementById(elementId).style.display="";
}

document.onkeydown=documentKeydown;

function documentKeydown(e) {
    var currKey = 0, e = e || event;
    currKey = e.keyCode || e.which || e.charCode;
    //alert(currKey)
    if (!e.shiftKey) {
        return onKeyDown(currKey);
    }
}

function focusElement(elementId) {
    document.getElementById(elementId).focus();
}

function setCmdTextClass(id1, id2, className) {
    var e1 = document.getElementById(id1);
    var e2 = document.getElementById(id2);
    e1.className = className;
    e2.className = className;
}
