function showAndHideForTwoList(listId1, listId2, types)
{ 
    var Layer1=window.document.getElementById(listId1); 
    var Layer2=window.document.getElementById(listId2); 
    switch(types){ 
        case "show": 
            Layer1.style.display="block"; 
            Layer2.style.display="block"; 
            break; 
        case "hide": 
            Layer1.style.display="none"; 
            Layer2.style.display="none"; 
    } 
} 

function addAfterLiForTwoList(str, inputId, listId, ulId)
{
    var obj = document.getElementById(ulId);
    var li = document.createElement("li");
    li.onmousedown = function(){setListValue(inputId, listId, str)};
    li.innerHTML = str;
    li.style.fontSize = "x-small";
    
    obj.appendChild(li);
}
