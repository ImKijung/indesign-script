function functionHideLayer() {
    //arguments: handler, mustSinge, needSave, needVisibleOpen
    progressTitle = '레이어 숨기는 중: ';
    if (GetLayerNames()) {
        if (SelectVisibleLayer('다음 레이어 중 숨길 레이어를 선택해 주세요', false)) {
            return function_core(HideLayer, false, true, false);
        } else {
            isStopped = true;
            return false;
        }
    } else {
        isStopped = true;
        return false;
    }
}

function functionShowLayer() {
    //arguments: handler, mustSinge, needSave, needVisibleOpen
    progressTitle = '레이어 표시 중: ';
    if (GetLayerNames()) {
        if (SelectVisibleLayer('다음 레이어 중 표시할 레이어를 선택해 주세요', true)) {
            return function_core(ShowLayer, false, true, false);
        } else {
            isStopped = true;
            return false;
        }
    } else {
        isStopped = true;
        return false;
    }
}

function functionAddLayer() {
    //arguments: handler, mustSinge, needSave, needVisibleOpen
    progressTitle = '레이어 추가 중: ';
    layerName = prompt('추가할 레이어 이름을 입력하십시오.', '신규 레이어 1', '레이어 이름 입력');
    if (layerName == null || layerName == '') {
        isStopped = true;
        return false;
    } else
        return function_core(AddLayer, false, true, false);
}

function functionDeleteLayer() {
    //arguments: handler, mustSinge, needSave, needVisibleOpen
    progressTitle = '레이어 삭제 중: ';
    if (confirm ('선택한 모든 레이어가 삭제됩니다.\r\n진행하시겠습니까?', true, '삭제 확인')) {
        if (GetLayerNames()) {
            if (SelectVisibleLayer('다음 레이어 중 삭제할 레이어를 선택해 주세요', true, true)) {

                return function_core(DeleteLayer, false, true, false);

            } else {
                isStopped = true;
                return false;
            }
        } else {
            isStopped = true;
            return false;
        }
    } else {
        isStopped = true;
        return false;
    }
}

function HideLayer(myDoc) {
    try {
        var result = true;
        SetLayers(myDoc, false);
        return result;
    } catch (ex) {
        errorMsgs.push('HideLayer 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function ShowLayer(myDoc) {
    try {
        var result = true;
        SetLayers(myDoc, true);
        return result;
    } catch (ex) {
        errorMsgs.push('ShowLayer 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function SetLayer(myDoc, visible) {
    try {
        var layers = myDoc.layers;
        for (var i = 0 ; i < layers.length ; i++) {
            var layer = layers[i];
            if (layer.name == '#레이어1') {
                layer.visible = true;
                continue;
            } else if (layer.name == '규격') {
                layer.visible = false;
                continue;
            } else {
                layer.visible = visible;
            }
        }
        return true;
    } catch (ex) {
        errorMsgs.push('SetLayer 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function SetLayers(myDoc, visible) {
    try {
        //code
        var layers = myDoc.layers;
        for (var i = 0 ; i < layers.length ; i++) {
            var layer = layers[i];
            var name = layer.name;
            layer.visible = false;
            for (var j = 0 ; j < layerNames.length ; j++) {
                if (name == layerNames[j]) {
                    layer.visible = layerVisibles[j];
                    break;
                }
            }
        }
        return true;
    } catch (ex) {
        errorMsgs.push('SetLayers 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function GetLayerNames() {
    var close = false;
    try {
        var documents = [];
        if (app.documents.length > 0) {
            try {
                documents[0] = app.activeDocument;
            } catch (ex) {
                for (var i = app.documents.length - 1 ; i >= 0 ; i--) {
                    var doc = app.documents[i];
                    if (doc.visible == false)
                        doc.close(SaveOptions.NO);
                }
                return GetLayerNames();
            }
        } else if (app.books.length > 0) {
            if (app.books.length > 1) {
                var sb = select_book();
                app.activeBook = sb;
            }
            var book = app.activeBook;
            var contents = book.bookContents;
            //set disable contents codes....
            for (var i = 0 ; i < contents.length ; i++) {
                documents[i] = app.open(contents[i].fullName, false);
            }
            close = true;
        } else {
            errorMsgs.push('필요한 문서나 책이 열려있지 않습니다.');
            isStopped = true;
            return false;
        }

        if (documents.length < 1) {
            throw new Error('작업대상인 문서나 책이 없습니다.', null, 41);
        } else {
            for (var i = 0 ; i < documents.length ; i++) {
                var doc =  documents[i];
                var layers = doc.layers;
                for (var j = 0 ; j < layers.length ; j++) {
                    var layer = layers[j];
                    var name = layer.name;
                    var exist = false;
                    for (var k = 0 ; k < layerNames.length ; k++) {
                        if (name == layerNames[k]) {
                            exist = true;
                            break;
                        }
                    }
                    if (exist == false) {
                        if (layerNames.length > 0)
                            layerNames[layerNames.length] = name;
                        else
                            layerNames[0] = name;
                    }
                }
                doc.close();
            }
        }

        return true;
    } catch(ex) {
        errorMsgs.push('GetLayerNames 함수 오류: Line:' + ex.line + ':: ' + ex);
        hasError = true;
        try {
            if (close) {

            }
        }
        catch (ex) { }
        throw ex;
    }
}

function SelectVisibleLayer(msg, visible, isRemove) {
    try {
        var vmsg = '표시할';
        if (visible == false) {
            vmsg = '숨길'
        }
        if (msg == null || msg == '') {
            msg = '다음 레이어 중 ' + vmsg + ' 레이어를 선택해 주세요';
        }
        var res =
            "dialog { alignChildren: 'fill', \
                Alert: Panel { orientation: 'column', alignChildren:'left', \
                    text: '" + msg + "', \
                }, \
                Layers: Panel { orientation: 'column', alignChildren:'left', \
                    text: '레이어 이름', \
                }, \
                buttons: Group { orientation: 'row', alignment: 'right', \
                    cancelBtn: Button { text:'Cancel', properties:{name:'cancel'} } \
                    okBtn: Button { text:'OK', properties:{name:'ok'} }, \
                } \
            }";
        var select = new Window( res, '레이어 선택창' );
        select.buttons.okBtn.onClick = function() {
            var selected = 0;
            for (var i = 0 ; i < layerNames.length ; i++) {
                if (visible == true || isRemove == true)
                    layerVisibles[i] = layerChecked[i].value;
                else
                    layerVisibles[i] = !layerChecked[i].value;

                if (layerChecked[i].value)
                    selected++;
            }
            if (selected > 0)
                select.close(1);
            else
                alert('최소 1개 이상의 레이어를 선택해 주세요.', '레이어 선택 확인');
        }
        var layerChecked = [];
        for (var i = 0 ; i < layerNames.length ; i++) {
            layerChecked[i] = select.Layers.add('checkbox', undefined, layerNames[i]);
            if (isRemove == true || !visible)
                layerChecked[i].value = false;
            else
                layerChecked[i].value = true;

            layerChecked[i].onClick = function() {
                if (isRemove == true && this.value == true) {
                    var panel = this.parent;
                    var checkboxs = panel.children;
                    var selected = 0;
                    for (var j = 0 ; j < checkboxs.length ; j++) {
                        if (checkboxs[j].value == true)
                            selected++;
                    }
                    if (selected == checkboxs.length) {
                        alert('최소 1개 이상의 레이어가 남아있어야 합니다.', '레이어 선택 확인');
                        this.value = false;
                    }
                }
            }
            //layerVisibles[i] = layerChecked[i].value;
        }
        var result = select.show();
        if (result == 1) {
            return true;
        } else {
            isStopped = true;
            return false;
        }
    } catch(ex) {
        errorMsgs.push('SelectVisibleLayer 함수 오류: Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function SetSepcificLayer(myDoc, visible) {
    try {
        var layers = myDoc.layers;
        try {
            var layer = layers.itemByName(layerName);
            if (layer == null) {
                errorMsgs.push("'" + myDoc.name + "' 문서에 '" + layerName + "' 레이어가 없습니다.");
                hasError = true;
            } else
                layer.visible = visible;
        } catch (ex) {
            throw ex;
        }
        return true;
    } catch (ex) {
        errorMsgs.push('SetLayer 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function DeleteLayer(myDoc) {
    try {
        var layers = myDoc.layers;
        for (var i = layers.length - 1 ; i >= 0 ; i--) {
            var layer = layers[i];
            var name = layer.name;
            for (var j = 0 ; j < layerNames.length ; j++) {
                if (name == layerNames[j]) {
                    if (layerVisibles[j]) {
                        layer.remove();
                    }
                    break;
                }
            }
        }

        /*
        var l1 = layers.itemByName('#레이어1');
        var l2 = layers.itemByName('규격');
        if (l1 == null && l2 == null) {
            errorMsgs.push("'" + myDoc.name + "' 문서에 '#레이어1', '규격' 레이어가 없습니다.");
            hasError = true;
        } else {
            for (var i = layers.length - 1 ; i >= 0 ; i--) {
                var layer = layers[i];
                if (layer.name == '#레이어1' || layer.name == '규격') {
                    continue;
                } else {
                    layer.remove();
                }
            }
        }
        */
        return true;
    } catch (ex) {
        errorMsgs.push('DeleteLayer 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function DeleteSecificLayer(myDoc) {
    try {
        var layers = myDoc.layers;
        var layer = layers.itemByName(layerName);
        if (layer != null)
            layer.remove();
        return true;
    } catch (ex) {
        errorMsgs.push('DeleteLayer 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function AddLayer(myDoc) {
    try {
        var layers = myDoc.layers;
        var layer = layers.itemByName(layerName);
        if (layer == null) {
            layer = layers.add();
            layer.name = layerName;
        } else {
            errorMsgs.push("'" + myDoc.name + "' 문서에 '" + layerName + "' 레이어가 이미 있습니다.");
            hasError = true;
        }
        return true;
    } catch (ex) {
        errorMsgs.push('AddLayer 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}