function DeleteHyperlink(myDoc) {
    try {
        var result = true;

        var hyper_count = myDoc.hyperlinkURLDestinations.count();
        for (var i = 0 ; i < hyper_count ; i++)
            myDoc.hyperlinkURLDestinations[0].remove();

        hyper_count = myDoc.hyperlinkTextSources.count();
        for (var i = 0 ; i < hyper_count ; i++)
            myDoc.hyperlinkTextSources[0].remove();

        hyper_count = myDoc.crossReferenceSources.count();
        for (var i = 0 ; i < hyper_count ; i++)
            myDoc.crossReferenceSources[0].remove();

        hyper_count = myDoc.hyperlinkTextDestinations.count();
        for (var i = 0 ; i < hyper_count ; i++)
            myDoc.hyperlinkTextDestinations[0].remove();

        hyper_count = myDoc.hyperlinks.count();
        for (var i = 0 ; i < hyper_count ; i++)
            myDoc.hyperlinks[0].remove();

        return true;
    } catch (ex) {
        errorMsgs.push('DeleteHyperlink 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

//문서 1개로 작업할때 하는 추가 작업
function SaveDestinationToMemory(myDoc) {
    try {
        var result = true;
        
        var docs = [];
        linked = [];
        var files = new Folder(myDoc.filePath).getFiles('*.indd');
        for (var i = 0 ; i < files.length ; i++) {
            if (files[i].toString() != myDoc.fullName)
                docs.push(app.open(files[i], false));
        }
        for (var i = 0 ; i < docs.length ; i++) {
            var item = docs[i];
            for (var j = 0 ; j < item.hyperlinks.count() ; j++) {
                var hyperlink = item.hyperlinks[j];
                try {
                    if (hyperlink.source.constructor == CrossReferenceSource && hyperlink.destination.parent.id == myDoc.id)
                        linked.push([item.id, hyperlink, hyperlink.destination.name]);
                } catch (ex) {
                    errorMsgs.push('SaveDestinationToMemory 함수 오류: Line:' + ex.line + ':: '
                     + ex.toString() + '\r\n오류가 난 문서 : ' + item.name + '\r\n오류가 난 링크 : ' + hyperlink);
                    hasError = true;
                    result = false;
                }
            }
        }

        return result;
    } catch (ex) {
        errorMsgs.push('SaveDestinationToMemory 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

//문서 1개로 작업할때 하는 추가 작업
function ReconnectDestinationFromMemory(myDoc) {
    try {
        var result = true;

        for (var i = 0 ; i < linked.length ; i++) {
            linked[i][1].destination = myDoc.hyperlinkTextDestinations.itemByName(linked[i][2].toString());
            app.documents.itemByID(linked[i][0] * 1).save();
        }

        for (var i = 0 ; i < app.documents.count() ; i++) {
            if (!app.documents[i].visible) {
                app.documents[i].close(SaveOptions.NO);
                i = -1;
            }
        }

        return result;
    } catch (ex) {
        errorMsgs.push('ReconnectDestinationFromMemory 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

//XML에 하이퍼링크 정보 저장
function SaveHyperlinkToXML(myDoc) {
    var link_item = null;
    var source_item = null;
    try {
        var result = true;
        var link_count = myDoc.hyperlinks.count();
        var destination_item = null;
        var newitem = null;
        var stringtype = '';
        for (var i = 0 ; i < link_count ; i++){
            link_item = myDoc.hyperlinks[i];
            if (link_item.name.replace(/[^0-9]/g,'') == '.')
                continue;
            source_item = link_item.source;
            destination_item = link_item.destination;
            
            if (source_item == null) {
                errorMsgs.push('SaveHyperlinkToXML 함수 오류: ' + myDoc.name + ' 파일에 잘못된 하이퍼 링크가 있습니다(source값 없음): ' + link_item.name);
                hasError = true;
                result = false;
            } else if (destination_item == null) {
                errorMsgs.push('SaveHyperlinkToXML 함수 오류: ' + myDoc.name + ' 파일에 잘못된 하이퍼 링크가 있습니다(destination값 없음): ' + link_item.name
                + ' (sourceText: ' + source_item.sourceText.contents + ')');
                hasError = true;
                result = false;
            } else if (source_item.constructor == CrossReferenceSource) {
                newitem = source_item.sourceText.associatedXMLElements[0].xmlElements.add('xref', source_item.sourceText);
                newitem.xmlAttributes.add('hyperlinkInform', 'CrossReferenceSource');
                newitem.xmlAttributes.add('name', link_item.name + stringtype);
                newitem.xmlAttributes.add('label', link_item.label + stringtype);
                newitem.xmlAttributes.add('crossReferenceFormatIndex', source_item.appliedFormat.index + stringtype);
                newitem.xmlAttributes.add('destinationDocumentName', destination_item.parent.name + stringtype);
                newitem.xmlAttributes.add('destinationName', destination_item.name + stringtype);
            } else if (source_item.constructor == HyperlinkTextSource) {
                if (destination_item.constructor == HyperlinkURLDestination) {
                    newitem = source_item.sourceText.associatedXMLElements[0];//;.xmlElements.add('xref', source_item.sourceText);
                    try {
                        newitem.xmlAttributes.add('xref', source_item.sourceText.contents);
                        newitem.xmlAttributes.add('hyperlinkInform', 'HyperlinkTextSource');
                        newitem.xmlAttributes.add('name', link_item.name + stringtype);
                        newitem.xmlAttributes.add('label', link_item.label + stringtype);
                        newitem.xmlAttributes.add('destinationName', destination_item.name + stringtype);
                        newitem.xmlAttributes.add('destinationURL', destination_item.destinationURL + stringtype);
                    } catch (ex) {
                        errorMsgs.push('SaveHyperlinkToXML xref 속성 추가 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
                        hasError = true;
                        result = false;
                    }
                }
            }
        }

        var anchor = null;
        var anchor_count = myDoc.hyperlinkTextDestinations.count();
        for (var i = 0; i < anchor_count ; i++) {
            anchor = myDoc.hyperlinkTextDestinations[i];
            newitem = null;
            newitem = anchor.destinationText.associatedXMLElements[0].xmlElements.add('xref', anchor.destinationText);

            try {
                newitem.xmlAttributes.add('hyperlinkInform', 'TextAnchor');
                newitem.xmlAttributes.add('name', anchor.name)
                newitem.xmlAttributes.add('label', anchor.label)
            } catch (ex) {
                errorMsgs.push('SaveHyperlinkToXML TextAnchor 속성 추가 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
                hasError = true;
                result = false;
            }
        }

        return result;
    } catch (ex) {
        var item = '';
        try {
            if (link_item != null && link_item.name != null) {
                item = ':' + link_item.name;
                if (source_item != null) {
                    if (source_item.toString() == '[object HyperlinkPageItemSource]') {
                        if (source_item.sourcePageItem != null) {
                            item = item + ', ' + source_item.sourcePageItem.parentPage.name + 'page';
                        }
                    } else {
                        if (source_item.sourceText != null && source_item.sourceText.pageItems.length > 0) {
                            item = item + ', ' + source_item.sourceText.pageItems[0].parentPage.name + 'page';
                        }
                    }
                }
            }
        } catch (ex2) { }
        errorMsgs.push('SaveHyperlinkToXML 함수 오류(' + myDoc.name + item + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

// XML정보에서 텍스트 앵커 새로 생성
function SetHyperlinkDestinationFromXML(myDoc) {
    try {
        //if (result == null)
        var result = true;

        var hyper_count = myDoc.hyperlinkTextDestinations.count();
        for (var i = 0 ; i < hyper_count ; i++)
            myDoc.hyperlinkTextDestinations[0].remove();

        var anchors = myDoc.xmlElements[0].evaluateXPathExpression('//*[@hyperlinkInform = \'TextAnchor\']');
        var anchor_count = anchors.length;
        for (var i = 0; i < anchor_count ; i++) {
            var anchor = anchors[i];
            var item = myDoc.hyperlinkTextDestinations.add(anchor.texts[0]);
            item.name = anchor.xmlAttributes.itemByName('name').value;
            item.label = anchor.xmlAttributes.itemByName('label').value + '';
        }

        return result;
    } catch (ex) {
        errorMsgs.push('SetHyperlinkDestinationFromXML 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

// XML정보에서 하이퍼링크 생성
function SetHyperlinkFromXML(myDoc) {
    var layerVisibles = [];
    try {
        var result = true;

        var xmllist = myDoc.xmlElements[0].evaluateXPathExpression('//*[@hyperlinkInform]');
        if (xmllist.length < 1)
            return result;

        SetLayerVisible(myDoc, layerVisibles, true);
        var xml_count = xmllist.length;
        var cross_item = null;
        var hyper_item = null;
        var newitem = null;
        var dest_item = null;
        var dest_doc = null;
        var linktype = null;
        var url = null;
        var dest_file = null;
        for (var i = 0 ; i < xml_count ; i++) {
            var item = xmllist[i];
            linktype = item.xmlAttributes.itemByName('hyperlinkInform').value;
            if (linktype == 'CrossReferenceSource') {
                cross_item = myDoc.crossReferenceSources.add(item.texts[0], myDoc.crossReferenceFormats[item.xmlAttributes.itemByName('crossReferenceFormatIndex').value * 1]);
                dest_file = new File(myDoc.filePath + '/' + item.xmlAttributes.itemByName('destinationDocumentName').value);
                if (!dest_file.exists) {
                    errorMsgs.push('SetHyperlinkFromXML 함수 오류: ' + dest_file + ' 문서가 없습니다.');
                    hasError = true;
                    result = false;
                } else {
                    dest_doc = app.open(new File(myDoc.filePath + '/' + item.xmlAttributes.itemByName('destinationDocumentName').value), false);
                    dest_item = dest_doc.hyperlinkTextDestinations.itemByName(item.xmlAttributes.itemByName('destinationName').value);
                    if (!dest_item.isValid) {
                        errorMsgs.push('SetHyperlinkFromXML 함수 오류: ' + item.xmlAttributes.itemByName('destinationDocumentName').value
                                + '문서의 ' + item.xmlAttributes.itemByName('destinationName').value + ' hyperlinkTextDestination이 올바르지 않습니다.');
                        hasError = true;
                        result = false;
                    } else {
                        newitem = myDoc.hyperlinks.add(cross_item, dest_item);
                        newitem.highlight = HyperlinkAppearanceHighlight.INVERT;
                        try {
                            newitem.name = item.xmlAttributes.itemByName('name').value;
                            newitem.label = item.xmlAttributes.itemByName('label').value;
                        } catch (ex) {
                            //need log
                        }
                    }
                }
            } else if (linktype == 'HyperlinkTextSource') {
                hyper_item = myDoc.hyperlinkTextSources.add(item.texts[0]);
                url = myDoc.hyperlinkURLDestinations.itemByName(item.xmlAttributes.itemByName('name').value);
                if (!url.isValid || url.destinationURL != item.xmlAttributes.itemByName('destinationURL').value)
                    url = myDoc.hyperlinkURLDestinations.add(item.xmlAttributes.itemByName('destinationURL').value, {hidden:true});
                newitem = myDoc.hyperlinks.add(hyper_item, url, {visible:false, hidden:false});
                try {
                    newitem.name = item.xmlAttributes.itemByName('name').value;
                    newitem.label = item.xmlAttributes.itemByName('label').value;
                } catch (ex) {
                    //errorMsgs.push('SetHyperlinkFromXML 함수 오류: ' + ex);
                    //hasError = true;
                    //result = false;
                }
            }
        }

        result = result && RemoveStructureAndTags(myDoc);
        SetLayerVisible(myDoc, layerVisibles, false);

        return result;
    } catch (ex) {
        errorMsgs.push('SetHyperlinkFromXML 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        CloseAllInvisibleDocuments();
        SetLayerVisible(myDoc, layerVisibles, false);
        throw ex;
    }
}