function addButtonOfFolder(divAreaId, folderName) {
	var btn_more = input_creat("button", "", "+", function(){clickPlus(folderName)}, 0);
	var btn_folder = input_creat("button", "", " " + folderName + " ", function(){clickFolder(folderName)}, 0);
	var btn_br = document.createElement("br");

	btn_more = setButtonStyle(btn_more);
	btn_folder = setButtonStyle(btn_folder);

    parentNode_appendChild(divAreaId, btn_more);
    parentNode_appendChild(divAreaId, btn_folder);
    parentNode_appendChild(divAreaId, btn_br);
}

function addButtonOfFile(divAreaId, fileName) {
	var btn_file = input_creat("button", "", " " + fileName + " ", function(){clickFile(fileName)}, 0);
	var btn_br = document.createElement("br");

	btn_file = setButtonStyle(btn_file);

    parentNode_appendChild(divAreaId, btn_file);
    parentNode_appendChild(divAreaId, btn_br);
}

function addExpPath(divPathId, folderName, pathLength) {
	var btn_Path = input_creat("button", "", folderName, function(){clickPath(pathLength)}, 0);
	btn_Path = setButtonStyle(btn_Path);
	parentNode_appendChild(divPathId, btn_Path);
}
