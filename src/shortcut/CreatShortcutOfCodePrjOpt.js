function addShortcutButton(sWorkName, sWorkCode, sWorkPrj, sWorkOpt, divId) 
{
	var baseId = sWorkName + "/" + sWorkCode + "/" + sWorkPrj + "/" + sWorkOpt
	var shortcutId = baseId + "_shortcut";
	var brId = baseId + "_br";
    var upId = baseId + "_up";
    var removeId = baseId + "_remove";

    var shortcutFun = function(){applyShortcut(sWorkName, sWorkCode, sWorkPrj, sWorkOpt)};
    var upFun = function(){upShortcut(sWorkName)}
	var removeFun = function(){removeShortcutBtn(upId, removeId, shortcutId, brId)};

	var shortcutValue = " " + sWorkName + " ";

    parentNode_appendChild(divId, createShortcutBtn(upId, " ↑ ", upFun));
    parentNode_appendChild(divId, createShortcutBtn(shortcutId, shortcutValue, shortcutFun));
    parentNode_appendChild(divId, createShortcutBtn(removeId, " × ", removeFun));
    parentNode_appendChild(divId, createBr(brId));
}

function createShortcutBtn(id, value, fun) {
    var btn = input_creat("button", id, value, fun, 0);
    btn = setButtonStyle(btn);
    return btn;
}

function createBr(id) {
	var btn = document.createElement("br");
    btn.id = id
    return btn;
}

function removeShortcutBtn(upId, removeId, shortcutId, brId)
{
    node_removeNode(upId);
    node_removeNode(removeId);
    node_removeNode(shortcutId);
    node_removeNode(brId);
    removeShortcut(shortcutId);
}
