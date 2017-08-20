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
