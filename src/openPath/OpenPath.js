function addAfterLiForOpenPath(str, inputId, listId, ulId)
{
    var obj = document.getElementById(ulId);
    var li = document.createElement("li");
    li.onmousedown = function(){setListValueForOpenPath(inputId, listId, str)};
    li.innerHTML = str;
    li.style.fontSize = "x-small";
    
    obj.appendChild(li);
}
