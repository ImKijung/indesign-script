function functionDeleteIndexes() {
    //arguments: handler, mustSinge, needSave, needVisibleOpen
    progressTitle = '인덱스 정보 제거 중: ';
    if (confirm ('책 혹은 현재 문서의 모든 색인을 삭제합니다.', true, '삭제 확인')) {
        return function_core(DeleteIndexes, false, true, false);
    } else {
        isStopped = true;
        return CloseWithoutMsg;
    }
}

//색인
function SaveIndexToXML(myDoc) {
    try {
        var result = true;

        var indexes = myDoc.indexes;
        for (var i = 0 ; i < indexes.length ; i++) {
            var index = indexes[i];
            var topics = index.allTopics;
            SetTopics(myDoc, topics);
        }

        return result;
    } catch (ex) {
        errorMsgs.push('SaveIndexToXML 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function SetTopics(myDoc, topics) {
    var tempId = 1;
    var errorCount = 0;
    for (var j = 0 ; j <topics.length ; j++) {
        try {
            var topic = topics[j];
            var topicName = topic.name;
            var pageReference = topic.pageReferences;
            var ancientName, parentName
            for (var k=0; k<pageReference.length; k++) {
                if (pageReference[k] != null) {
                    var sourceText = pageReference[k].sourceText;
                    var sourceXML = sourceText.associatedXMLElements[0];
                    if (topic.name == '') { 
                        tempId += 1;
                        topicName = 'topic' + tempId; 
                    }
                    var newitem = sourceXML.xmlElements.add('topicItem', sourceText);
                    newitem.xmlAttributes.add('topicName', topicName);
                    newitem.xmlAttributes.add('pageReferenceType', Number(pageReference[k].pageReferenceType) + '');
    
                    if (topic.parent.toString() == '[object Topic]') {
                        if (topic.parent.name == '') {
                            tempId += 1;
                            parentName = 'topic' + tempId; 
                        } else
                            parentName = topic.parent.name;
                        newitem.xmlAttributes.add('parentName', parentName);
                        if (topic.parent.parent.toString() == '[object Topic]') {
                            if (topic.parent.parent.name == '') {
                                tempId += 1;
                                ancientName = 'topic' + tempId; 
                            } else
                                ancientName = topic.parent.parent.name;
                            newitem.xmlAttributes.add('ancientName', ancientName);
                        }
                    }
                } else {
                    if (topic.name == '') {
                        tempId += 1;
                        topicName = 'topic' + tempId; 
                    }
                    var newitem = myDoc.xmlElements[0].xmlElements.add('topicItem', topicName);
                    newitem.xmlAttributes.add('topicName', topicName);
    
                    if (topic.parent.toString() == '[object Topic]') {
                        parentName = topic.parent.name;
                        if (topic.parent.name == '') {
                            tempId += 1;
                            parentName = 'topic' + tempId; 
                        }
                        newitem.xmlAttributes.add('parentName', parentName);
                        if (topic.parent.parent.toString() == '[object Topic]') {
                            ancientName = topic.parent.parent.name;
                            if (topic.parent.parent.name == '') {
                                tempId += 1;
                                ancientName = 'topic' + tempId; 
                            }
                            newitem.xmlAttributes.add('ancientName', ancientName);
                        }
                    }
                }
            }
        } catch (ex) {
            errorCount++;
        }
    }
    if (errorCount > 0) {
        errorMsgs.push('인덱스 저장 오류: ' + myDoc.name + '문서에서 ' + errorCount + '개의 인덱스를 XML에 저장하지 못했습니다.');
    }
}

function SetIndexFromXML(myDoc) {
    try {
        var result = true;

        DeleteIndexes(myDoc);
        
        var errorCount = 0;
        var index = null;
        var indexes = myDoc.indexes;
        if (indexes == null || indexes.length == 0) {
            index = indexes.add();
        } else {
            index = indexes[0];
        }
        var topics = index.topics;
        var pageItems = myDoc.xmlElements[0].evaluateXPathExpression('//*[@topicName][not(@parentName)]');
        for (var i = 0 ; i < pageItems.length ; i++) {
            try {
                var pageItem = pageItems[i];
                var topic = topics.add(pageItem.xmlAttributes.itemByName('topicName').value);
                if (pageItem.xmlAttributes.itemByName('pageReferenceType') != null) {
                    var lastText = pageItem.parent.texts[pageItem.parent.texts.length - 1];
                    var lastIndex = lastText.insertionPoints.length - 2;
                    if (lastIndex < 0) {
                        lastIndex = 0;
                    }
                    // 11111
                    var tempFrame = myDoc.textFrames.add();
                    var pageReference = topic.pageReferences.add(tempFrame.parentStory.texts[0]);
                    pageReference.pageReferenceType = Number(pageItem.xmlAttributes.itemByName('pageReferenceType').value);
                    tempFrame.parentStory.texts[0].move(LocationOptions.AT_BEGINNING, lastText.insertionPoints[lastIndex])
                    // var pageReference = topic.pageReferences.add(lastText.insertionPoints[lastIndex]);
                    tempFrame.remove();
                }
            } catch (ex) {
                // errorMsgs.push('SetIndexFromXML 함수 오류3: ' + myDoc.name + '문서 - ' + lastText + ":" + ex);
                errorCount++;
            }
        }

        topics = index.topics;
        pageItems = myDoc.xmlElements[0].evaluateXPathExpression('//*[@topicName][@parentName][not(@ancientName)]');
        for (var i = 0 ; i < pageItems.length ; i++) {
            try {
                var pageItem = pageItems[i];
                var parent = topics.itemByName(pageItem.xmlAttributes.itemByName('parentName').value)
                var topic = parent.topics.add(pageItem.xmlAttributes.itemByName('topicName').value);
                if (pageItem.xmlAttributes.itemByName('pageReferenceType') != null) {
                    var lastText = pageItem.parent.texts[pageItem.parent.texts.length - 1];
                    var lastIndex = lastText.insertionPoints.length - 2;
                    if (lastIndex < 0) {
                        lastIndex = 0;
                    }
                    // 2222
                    var tempFrame = myDoc.textFrames.add();
                    var pageReference = topic.pageReferences.add(tempFrame.parentStory.texts[0]);
                    pageReference.pageReferenceType = Number(pageItem.xmlAttributes.itemByName('pageReferenceType').value);
                    tempFrame.parentStory.texts[0].move(LocationOptions.AT_BEGINNING, lastText.insertionPoints[lastIndex])
                    // var pageReference = topic.pageReferences.add(lastText.insertionPoints[lastIndex]);
                    tempFrame.remove();
                }
            } catch (ex) {
                // errorMsgs.push('SetIndexFromXML 함수 오류3: ' + myDoc.name + '문서 - ' + lastText + ":" + ex);
                errorCount++;
            }
        }

        topics = index.topics;
        pageItems = myDoc.xmlElements[0].evaluateXPathExpression('//*[@topicName][@ancientName]');
        for (var i = 0 ; i < pageItems.length ; i++) {
            try {
                var pageItem = pageItems[i];
                var ancient = topics.itemByName(pageItem.xmlAttributes.itemByName('ancientName').value)
                var parent = ancient.topics.itemByName(pageItem.xmlAttributes.itemByName('parentName').value)
                var topic = parent.topics.add(pageItem.xmlAttributes.itemByName('topicName').value);
                if (pageItem.xmlAttributes.itemByName('pageReferenceType') != null) {
                    var lastText = pageItem.parent.texts[pageItem.parent.texts.length - 1];
                    var lastIndex = lastText.insertionPoints.length - 2;
                    if (lastIndex < 0) {
                        lastIndex = 0;
                    }
                    // 3333
                    var tempFrame = myDoc.textFrames.add();
                    var pageReference = topic.pageReferences.add(tempFrame.parentStory.texts[0]);
                    pageReference.pageReferenceType = Number(pageItem.xmlAttributes.itemByName('pageReferenceType').value);
                    tempFrame.parentStory.texts[0].move(LocationOptions.AT_BEGINNING, lastText.insertionPoints[lastIndex])
                    // var pageReference = topic.pageReferences.add(lastText.insertionPoints[lastIndex]);
                    tempFrame.remove();
                }
            } catch (ex) {
                // errorMsgs.push('SetIndexFromXML 함수 오류3: ' + myDoc.name + '문서 - ' + lastText + ":" + ex);
                errorCount++;
            }
        }
        if (errorCount > 0) {
            errorMsgs.push('인덱스 재설정 오류: ' + myDoc.name + '문서에서 ' + errorCount + '개의 인덱스를 복구하지 못했습니다.');
        }

        return result;
    } catch (ex) {
        errorMsgs.push('SetIndexFromXML 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function DeleteIndexes(myDoc) {
    try {
        var result = true;

        var indexes = myDoc.indexes;
        for (var i = indexes.length - 1 ; i >= 0 ; i--) {
            var index = indexes[i];
            var topics = index.allTopics;
            for (var j = topics.length - 1 ; j >= 0 ; j--) {
                try {
                    topics[j].remove();
                } catch (ex) {
                    ex;
                }
            }
            try {
                index.removeUnusedTopics();
            } catch (ex) {
                ex;
            }
        }
        return result;
    } catch (ex) {
        errorMsgs.push('DeleteIndexes 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}