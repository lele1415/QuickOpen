function removeAllButton(divId) {
	var i = parentNode_getChildNodesLength(divId);
    if (i>0) {
        for(var j=0; j<i; j++){
            parentNode_removeChild(divId, 0);
        }
    }
}

function addButtonOfFolder(divAreaId, name, divPathId, where) {
	var btn_more = input_creat("button", "", "+", function(){AddSubFolder(name, 0, divAreaId, divPathId, where)}, 0);
	var btn_folder = input_creat("button", "", " " + name + " ", function(){OpenFolder(name, where)}, 0);
	var btn_br = document.createElement("br");

	btn_more = setButtonStyle(btn_more);
	btn_folder = setButtonStyle(btn_folder);

    parentNode_appendChild(divAreaId, btn_more);
    parentNode_appendChild(divAreaId, btn_folder);
    parentNode_appendChild(divAreaId, btn_br);
}

function addButtonOfFile(divAreaId, name, where) {
	var btn_file = input_creat("button", "", " " + name + " ", function(){OpenFolder(name, where)}, 0);
	var btn_br = document.createElement("br");

	btn_file = setButtonStyle(btn_file);

    parentNode_appendChild(divAreaId, btn_file);
    parentNode_appendChild(divAreaId, btn_br);
}

function addButtonOfCurrentPath(divPathId, name, pathDeepCount, divAreaId, where) {
	var btn_Path = input_creat("button", "", name, function(){AddSubFolder(name, pathDeepCount, divAreaId, divPathId, where)}, 0);
	btn_Path = setButtonStyle(btn_Path);
	parentNode_appendChild(divPathId, btn_Path);
}

function removeButtonOfCurrentPath(divAreaId, currentDeepCount, nextDeepCount) {
	for (var i=currentDeepCount-1; i>nextDeepCount-1; i--) {
		parentNode_removeChild(divAreaId, i);
	}
}
function setButtonStyle(btn) {
	btn.style.width = "auto";
	btn.style.overflow = "visible";
	btn.style.border = "1px solid";
	btn.style.margin = "1px";
	btn.style.solid = "#f8f8f2";
	btn.style.background = "none";
	btn.style.color = "#f8f8f2";
	btn.onmouseover = function(){ 
            this.style.background="#66d9ef";}
    btn.onmouseout = function(){ 
            this.style.background="";}
	return btn;
}