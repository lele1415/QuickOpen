addFirstButtonOfMoreTextInput();
var input_count = 1;

function addFirstButtonOfMoreTextInput()
{
    var inputFirst = input_creat("button","input_1","+",addButtonOfMoreTextInput,0);
    parentNode_appendChild("input_text",inputFirst);
}

function addButtonOfMoreTextInput() 
{
    var count0 = input_count;
    var count1 = input_count + 1;
    var count2 = input_count + 2;
    var count3 = input_count + 3;

    var input_child = input_creat("text","input_"+(count1),0,0,0);

    var br = document.createElement("br");
    br.id = "input_"+(count2);

    var button_child = input_creat("button","input_"+(count3),"+",addButtonOfMoreTextInput,0);

    parentNode_appendChild("input_text",input_child);
    parentNode_appendChild("input_text",br);
    parentNode_appendChild("input_text",button_child);

    input_changeInfo("input_"+(count0),"-",function(){removeButtonOfMoreTextInput(count0)});

    input_count = input_count + 3;
}

function removeButtonOfMoreTextInput(i)
{
    for(var j=i;j<i+3;j++){
        node_removeNode("input_"+j);
    }
}

function addButtonOfMore(ChildName,Path,FolderDeepCount) 
{
    var input = input_creat("button",ChildName+"_More_id","+",function(){OpenMore(ChildName,Path,FolderDeepCount+1)},0);
    parentNode_insertBefore(input,ChildName)
}

function addDiv(ButtonId,DivId) 
{
    var Button = document.getElementById(ButtonId);
    Button.insertAdjacentHTML("afterEnd","<div id="+DivId+" ></div>");
}

function addButtonOfFolderName(path,ParentFolderName,ChildName,FolderDeepCount,Folderflag) 
{
    var input = input_creat("button", ParentFolderName+ChildName, ChildName, function(){OpenByPath(path+ChildName)},0)

    var a = document.createElement("a");
    for (var i=0;i<FolderDeepCount;i++){
        a.innerHTML += "&nbsp;&nbsp;&nbsp;&nbsp;"
    }

    var br = document.createElement("br");

    parentNode_appendChild(ParentFolderName+"MoreDiv_id", a);
    parentNode_appendChild(ParentFolderName+"MoreDiv_id", input);
    parentNode_appendChild(ParentFolderName+"MoreDiv_id", br);

    if(Folderflag){
        addButtonOfMore(ParentFolderName+ChildName,path+ChildName,FolderDeepCount);
    }
}

function showAndHide(obj,types)
{ 
    var Layer=window.document.getElementById(obj); 
    switch(types){ 
        case "show": 
            Layer.style.display="block"; 
            break; 
        case "hide": 
            Layer.style.display="none"; 
    } 
} 

function getValue(obj,str,sVersion)
{ 
    var input=window.document.getElementById(obj); 
    input.value=str; 
    removeAllMoreButton(sVersion);
}

/*function addLi(str,inputId,ulId)
{
    var obj = document.getElementById(ulId);
    var li = document.createElement("li");
    //li.onmousedown="showAndHide 'List1','hide'";
    li.onmousedown=function(){getValue(inputId,str)};
    li.innerHTML = str;
    li.style.fontSize = "x-small";
         
    obj.appendChild(li);
}*/

function addBeforeLi(str,inputId,ulId,sVersion)
{
    var obj = document.getElementById(ulId);
        var li = document.createElement("li");
        li.onmousedown=function(){getValue(inputId,str,sVersion)};
        li.innerHTML = str;
        li.style.fontSize = "x-small";
    if(obj.childNodes.length > 0){
        var node = obj.childNodes[0];
        obj.insertBefore(li,node);
    }
    else{
    obj.appendChild(li);
    }
}

function addSelectOfL1Project(DivId,SelectId,SelectOnchangeFun)
{
    if (parentNode_getChildNodesLength(DivId)<2){
        var select = select_creat(SelectId, SelectOnchangeFun, 0);
        parentNode_appendChild(DivId, select);
    }
}

function addSelectOfL1Project1()
{
    addSelectOfL1Project("project_l1_div", "projectFolder_1", function(){onloadL1Project2()});
}

function addSelectOfL1Project2()
{
    addSelectOfL1Project("project_l1_div", "projectFolder_2", 0);
}

function addOption(SelectId,OptionName) 
{
    var option = option_creat(OptionName,OptionName);
    parentNode_appendChild(SelectId,option);
}

function removeAllOption(SelectId)
{
    var i = parentNode_getChildNodesLength(SelectId);
    if (i>0){
        for(var j=0;j<i;j++){
            parentNode_removeChild(SelectId,0);
        }   
    }
}

function addFolderNameList(Name,ButtonValue,Path)
{
    var DivId = Name+"_list_div";
    var SelectId = Name+"_list_select";
    var ButtonId = Name+"_list_button";
    var HideAndShowId = Name+"_list_has"

    var button2 = document.getElementById(HideAndShowId);
    if (element_getValue(HideAndShowId) == "+") {
        element_setValue(HideAndShowId, "-");

        var select = select_creat(SelectId, 0, 250);

        if (Name == "project") {
            var button1 = input_creat("button", ButtonId, ButtonValue, function(){applyProjectName(SelectId)}, 0);
        } else if (Name == "modem_kk") {
            var button1 = input_creat("button", ButtonId, ButtonValue, function(){copyModemName("KK")}, 0);
        } else if (Name == "modem_l1") {
            var button1 = input_creat("button", ButtonId, ButtonValue, function(){copyModemName("L1")}, 0);
        }

        parentNode_appendChild(DivId, select);
        parentNode_appendChild(DivId, button1);

        onloadFolderNameList(Path,SelectId);
    } else {
        element_setValue(HideAndShowId, "+");
        node_removeNode(SelectId);
        node_removeNode(ButtonId);
    }
}

function addShortcutOfProjectName(version)
{
    if (version == "KK") {
        var source = element_getValue("codePath_KK");
        var pjname = element_getValue("project_name");
    } else {
        var source = element_getValue("codePath_L1");
        var pjfolder1 = element_getValue("projectFolder_1");
        var pjfolder2 = element_getValue("projectFolder_2");
    }
    if (version == "KK") {
        if (pjname != "") {
            var clickfun = function(){clickFunOfShortcut(source, pjname)};
            var button = input_creat("button", pjname, pjname, clickfun, 300);
            var removefun = function(){removeShortcut(pjname, pjname+"_remove", pjname+"_br")};
            var button1 = input_creat("button", pjname+"_remove", "-", removefun, 0);
            var br = document.createElement("br");
            br.id = pjname+"_br"

            parentNode_appendChild("shortcut_id", button1);
            parentNode_appendChild("shortcut_id", button);
            parentNode_appendChild("shortcut_id", br);
        }
    } else {
        var clickfun = function(){clickFunOfShortcutL1(source, pjfolder1, pjfolder2)};
        var button = input_creat("button", source+"/"+pjfolder1+"/"+pjfolder2, source+"\n"+pjfolder1+"/"+pjfolder2, clickfun, 300);
        var removefun = function(){removeShortcut(source+"/"+pjfolder1+"/"+pjfolder2, source+"/"+pjfolder1+"/"+pjfolder2+"_remove", source+"/"+pjfolder1+"/"+pjfolder2+"_br")};
        var button1 = input_creat("button", source+"/"+pjfolder1+"/"+pjfolder2+"_remove", "-", removefun, 0);
        var br = document.createElement("br");
        br.id = source+"/"+pjfolder1+"/"+pjfolder2+"_br"

        parentNode_appendChild("shortcut_l1_id", button1);
        parentNode_appendChild("shortcut_l1_id", button);
        parentNode_appendChild("shortcut_l1_id", br);
}
}

function clickFunOfShortcut(sourceValue, pjnameValue)
{
    element_setValue("codePath_KK", sourceValue);
    element_setValue("project_name", pjnameValue);
    removeAllMoreButton('KK');
}

function clickFunOfShortcutL1(sourceValue, pjfolder1, pjfolder2)
{
    element_setValue("codePath_L1", sourceValue);
    onloadL1Project1();
    element_setValue("projectFolder_1", pjfolder1);
    element_setValue("projectFolder_2", pjfolder2);
}

function removeShortcut(buttonId, button1Id, brId)
{
    node_removeNode(buttonId);
    node_removeNode(button1Id);
    node_removeNode(brId);
}