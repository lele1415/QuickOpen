function addShortcutButton(sCurrentCode, sCurrentPrj, divId) 
{
    var baseId = sCurrentCode + "/" + sCurrentPrj
    var shortcutId = baseId + "_shortcut";
    var brId = baseId + "_br";
    var removeId = baseId + "_remove";

    var shortcutFun = function(){applyShortcut(sCurrentCode, sCurrentPrj)};
    var removeFun = function(){removeShortcut(removeId, shortcutId, brId)};

    var shortcutValue = " " + sCurrentCode + "\n" + sCurrentPrj + " "

    parentNode_appendChild(divId, createRemoveBtn(removeId, removeFun));
    parentNode_appendChild(divId, createShortcutBtn(shortcutId, shortcutValue, shortcutFun));
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

function createRemoveBtn(id, fun) {
    var btn = input_creat("button", id, " - ", fun, 0);
    btn = setButtonStyle(btn);
    return btn;
}

function removeShortcut(buttonId, button1Id, brId)
{
    node_removeNode(buttonId);
    node_removeNode(button1Id);
    node_removeNode(brId);
}
