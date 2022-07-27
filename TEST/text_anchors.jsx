//DESCRIPTION: Show all text anchors, select unused ones.
//Peter Kahrel

#targetengine text_anchors;

(function () {

	var styleName = 'anchor display';
	
	function scriptPath () {
		try {
			return app.activeScript;
		} catch (e) {
			return File (e.fileName);
		}
	}

	function saveData (obj) {
		var f = File (scriptPath().fullName.replace (/\.jsx?(bin)?$/, '.txt'));
		f.open ('w');
		f.write (obj.toSource());
		f.close ();
	}

	function getPrevious () {
		var f = File (scriptPath().fullName.replace (/\.jsx?(bin)?$/, '.txt'));
		var obj = {};
		if (f.exists) {
			obj = $.evalFile (f);
		}
		return obj;
	}


	function unaccent (s) {
		return s.replace (/[-_\x20]/g, '').
		replace (/[ÁÀÂÄÅĀĄĂÆ]/g, 'A').
		replace (/[ÇĆČĊ]/g, 'C').
		replace (/[ĎĐ]/g, 'D').
		replace (/[ÉÈÊËĘĒĔĖĚ]/g, 'E').
		replace (/[ĢĜĞĠ]/g, 'G').
		replace (/[ĤĦ]/g, 'H').
		replace (/[ÍÌÎÏĪĨĬĮİ]/g, 'I').
		replace (/[Ĵ]/g, 'J').
		replace (/[Ķ]/g, 'K').
		replace (/[ŁĹĻĽ]/g, 'L').
		replace (/[ÑŃŇŅŊ]/g, 'N').
		replace (/[ÓÒÔÖŌŎŐØŒ]/g, 'O').
		replace (/[ŔŘŖ]/g, 'R').
		replace (/[ŚŠŜŞȘSS]/g, 'S').
		replace (/[ŢȚŤŦ]/g, 'T').
		replace (/[ÚÙÛÜŮŪŲŨŬŰŲ]/g, 'U').
		replace (/[Ŵ]/g, 'W').
		replace (/[ŸÝŶ]/g, 'Y').
		replace (/[ŹŻŽ]/g, 'Z');
	}

	function check_styles (s) {
			
		function getColour (s, val) {
			if (!app.documents[0].swatches.item(s).isValid) {
				app.documents[0].colors.add({name: s, space: ColorSpace.CMYK, colorValue: val});
			}
			return app.documents[0].swatches.item(s);
		}

		if (app.documents[0].paragraphStyles.item(s) == null) {
			var pstyle = app.documents[0].paragraphStyles.add ({name: s});
			try {
				pstyle.appliedFont = app.fonts.item ('Myriad Pro\tCondensed')
			} catch (_) {
			};
			pstyle.properties = {
				pointSize: 10,
				leading: 11,
				fillColor: getColour (s + ' text', [0, 100, 100, 0]),
			}
		}


		if (app.documents[0].objectStyles.item (s) == null) {
			app.documents[0].objectStyles.add ({
				name: s,
				basedOn: app.documents[0].objectStyles[0],
				enableParagraphStyle: true,
					appliedParagraphStyle: s,
				enableStroke: true,
					strokeWeight: 0.25,
				enableFill: true,
					fillColor: getColour (s+' background', [0, 0, 100, 0]),
					fillTint: 30,
				enableAnchoredObjectOptions: true,
					anchoredObjectSettings: {
						spineRelative: true,
						anchoredPosition: AnchorPosition.ANCHORED,
						anchorPoint: AnchorPoint.BOTTOM_RIGHT_ANCHOR,  // Anchored OBJECT > reference point
						horizontalReferencePoint: AnchoredRelativeTo.TEXT_FRAME, // columnEdge,
						anchorXoffset: '0',
						horizontalAlignment: HorizontalAlignment.LEFT_ALIGN,  // Anchored POSITION > reference point
						pinPosition: true,
					},
				enableTextFrameAutoSizingOptions: true,
				enableTextFrameGeneralOptions: true,
					textFramePreferences: {
						autoSizingReferencePoint: AutoSizingReferenceEnum.CENTER_POINT,
						autoSizingType: AutoSizingTypeEnum.HEIGHT_ONLY,
						useFixedColumnWidth: true,
						//textFramePreferences.textColumnFixedWidth: app.documents[0].pages[0].marginPreferences.right - 12,
						textColumnFixedWidth: '160 pt',
						insetSpacing: '3 pt',
					}
				}
			);
		return app.documents[0].objectStyles.item (s);
		}
	}


	function pad (n) {
		switch (n.length) {
			case 1: return '00'; break;
			case 2: return '0'; break;
			default: return '';
		}
	}


	function getAnchors () {
		var anchors = {};
		var ret = [];
		
		function stripString (str) {
			return unaccent (str.toUpperCase()).replace(/(\d+)/, function () {
				return (pad(arguments[1]))+arguments[1]
			})
		}

		function bmarkIDs () {
			var name;
			var bmarks = app.documents[0].bookmarks.everyItem().getElements();
			for (var i = bmarks.length-1; i >= 0; i--) {
				if (bmarks[i].destination instanceof HyperlinkTextDestination) {
					name = bmarks[i].destination.name;
					anchors[name] = {
						sortKey: stripString (name),
						destinationName: [bmarks[i].name],
						type: ['Bookmark']
					};
				}
			}
		}

		function getDocNames () {
			var o = {};
			var n = app.documents.everyItem().name;
			for (var i = 0; i < n.length; i++) {
				o[n[i]] = true;
			}
			return o;
		}
	

		function hlinkIDs (doc) {
			var docCount = app.documents.length;
			var hlinks = doc.hyperlinks.everyItem().getElements();
			var name;
			var o = {};
			for (var i = hlinks.length-1; i > -1; i--) {
				try {
					hlinks[i].destination;
				} catch (_) {
					continue;
				}
				if (!(hlinks[i].destination instanceof HyperlinkTextDestination)) {
					continue;
				}
				var name = hlinks[i].destination.name;
				
				var suffix = hlinks[i].destination.parent == doc ? '' : '*';
				if (app.documents.length > docCount) {
					app.documents[0].close (SaveOptions.NO);
					continue;
				}
				
				if (anchors[name]) {
					anchors[name].destinationName.push (hlinks[i].name+suffix);
					switch (hlinks[i].source.constructor.name) {
						case 'CrossReferenceSource' : anchors[name].type.push('Xref'+suffix); break;
						case 'HyperlinkTextSource': anchors[name].type.push('Hyperlink'+suffix);
					}
				} else {
					o = {
						sortKey: stripString (name),
						destinationName: [hlinks[i].name],
					}
					switch (hlinks[i].source.constructor.name) {
						case 'CrossReferenceSource' : o.type = ['Xref'+suffix]; break;
						case 'HyperlinkTextSource': o.type = ['Hyperlink'+suffix];
					}
					anchors[name] = o;
				}
			}
		}



		function independentIDs () {
			var tDest = app.documents[0].hyperlinkTextDestinations.everyItem().getElements();
			for (var i = tDest.length-1; i >= 0; i--) {
				if (!anchors[tDest[i].name]) {
					anchors[tDest[i].name] = {
						sortKey: stripString (tDest[i].name),
						destinationName: [],
						type: []
					};
				}
			}
		}


		bmarkIDs();
		hlinkIDs(app.documents[0]);
		independentIDs();
		
		for (var i in anchors) {
			ret.push ({
				name: i,
				sortKey: anchors[i].sortKey,
				destinationName: anchors[i].destinationName,
				type: anchors[i].type,
			});
		}
		return ret.sort (function (a,b) {
			return a.sortKey > b.sortKey;
		});
	}

	//=====================================================================================


	function inSelection (anchor) {
		var d = anchor.destinationText;
		if (d.parentStory != app.selection[0].parentStory) return false;
		if (d.index < app.selection[0].insertionPoints[0].index) return false;
		if (d.index > app.selection[0].insertionPoints[-1].index) return false;
		return true;
	}


	function showAnchors () {

		var anchors = app.documents[0].hyperlinkTextDestinations.everyItem().getElements();
		if (anchors.length == 0) return;
		
		var ostyle = check_styles (styleName);

		if (app.selection.length > 0 && app.selection[0].hasOwnProperty ('parentStory')) {
			for (var i = anchors.length-1; i >= 0; i--) {
				if (inSelection (anchors[i])) {
					anchors[i].destinationText.textFrames.add ({appliedObjectStyle: ostyle, contents: anchors[i].name});
				}
			}
		} else {
			for (var i = anchors.length-1; i >= 0; i--) {
				anchors[i].destinationText.textFrames.add ({appliedObjectStyle: ostyle, contents: anchors[i].name});
			}
		}
	}

	//-------------------------------------------------------------------------------------------------------------------------------
	
	function renameAnchor (list, name) {
		var w = new Window ('dialog', 'Rename text destination');
			w.group = w.add ('group');
				w.group.add ('statictext {text: "New name:"}');
				w.input = w.group.add ('edittext {characters: 20, active: true, text: "' + name + '"}');
				w.layout.layout();
				w.input.active = true;
				w.buttons = w.add ('group {alignment: "right"}');
					w.ok = w.buttons.add ('button {text: "OK", enabled: false}');
					w.buttons.add ('button {text: "Cancel"}');

			w.input.onChanging = function () {
				w.ok.enabled = list.find (w.input.text) == null;
			}

		if (w.show() == 1) {
			app.activeDocument.hyperlinkTextDestinations.item(name).name = w.input.text;
			list.selection[0].text = w.input.text;
		}
	}

	//===============================================================================================


	var columnProperties = {
		numberOfColumns: 3, 
		columnWidths: [150, 150, 150], 
		columnTitles: ['Text anchor', 'Source type', 'Source name'],
		showHeaders: true, 
		multiselect: true
	}

	

	function main () {
		var list_item;
		var names = getAnchors (app.activeDocument);
		var w = new Window ('palette', 'Text anchors', undefined, {resizeable: true});
			w.alignChildren = 'top';
			w.orientation = 'row';
			w.listContainer = w.add ('group {alignChildren: "top"}');
			var list = w.listContainer.add ('listbox', undefined, undefined, columnProperties);		
			list.minimumSize = w.listContainer.minimumSize = [400,400];
			list.alignment = w.listContainer.alignment = ['fill', 'fill'];
			list.maximumSize.height = w.maximumSize.height-150;

			var b = w.add ('group {orientation: "column", alignChildren: "fill"}');
				b.alignment = ['right', 'top']
				w.filter = b.add ('edittext {active: true, helpTip: "Filter by name"}');
				var go_to = b.add ('button', undefined, 'Go to selected');
				var rename_anchor = b.add ('button', undefined, 'Rename');
				var move_anchor = b.add ('button', undefined, 'Move to ip');
				var del_selected = b.add ('button', undefined, 'Delete selected');
				var sel_unused = b.add ('button', undefined, 'Select unused');
				var show_anchors = b.add ('button', undefined, 'Show anchors');
				var delete_frames = b.add ('button', undefined, 'Delete displayed anchors');

			for (var i = 0; i < names.length; i++) {
				list_item = list.add ('item', names[i].name);
				if (names[i].type && names[i].type.length) {
					list_item.subItems[0].text = names[i].type.join('/');
				}
				if (names[i].destinationName.length) {
					list_item.subItems[1].text = names[i].destinationName.join('/');
				}
			}
			
			list.maximumSize.height = w.maximumSize.height - 300;

			w.filter.onChange = filterList;
	
			function filterList () {
				var newlist = w.listContainer.add ('listbox', list.bounds, '', columnProperties);
				for (var i = 0; i < names.length; i++) {
					if (names[i].name.search (w.filter.text) > -1) {
						with (newlist.add ('item', names[i].name)) {
							subItems[0].text = names[i].type;
							subItems[1].text = names[i].destinationName;
						}
					}
				}
				w.remove(list);
				list = newlist;
				w.filter.active = true;
				w.text = 'Text anchors (' + String(list.children.length).replace (/(\d)(?=(\d\d\d)+$)/g, "$1,") + ')';
			}


			list.onChange = function () {
				if (list.selection == null) { // Can this happen?
					go_to.enabled = false;
					rename_anchor.enabled = false;
					move_anchor.enabled = false;
					del_selected.enabled = false;
					return;
				}

				go_to.enabled = true;
				rename_anchor.enabled = true;
				move_anchor.enabled = true;
				del_selected.enabled = true;

				if (list.selection.length == 1) {
					del_selected.enabled = list.selection[0].subItems[0].text == '';
					if (/\*$/.test (list.selection[0].subItems[0].text)) {
						go_to.enabled = false;
						rename_anchor.enabled = false;
						move_anchor.enabled = false;
						del_selected.enabled = false;
					}
					return;
				}

				for (var i = list.selection.length-1; i > -1; i--) {
					if (list.selection[i].subItems[0].text !== '') {
						del_selected.enabled = false;
						go_to.enabled = false;
						rename_anchor.enabled = false;
						move_anchor.enabled = false;
						break;
					}
				}
			}


			go_to.onClick = list.onDoubleClick = function () {
				var ip = app.activeDocument.hyperlinkTextDestinations.item (list.selection[0].text).destinationText;
				try {
					if (ip.parent instanceof Cell || ip.parent instanceof Footnote) {
						ip.parent.texts[0].characters[ip.index].select();
					} else {
						ip.parentStory.characters[ip.index].select();
					}
				} catch (_) {
					ip.select();
				}
				app.documents[0].layoutWindows[0].zoomPercentage = 200;
			}


			move_anchor.onClick = function () {
				if (app.selection.length > 0 && app.selection[0] instanceof InsertionPoint) {
					var d = app.documents[0].hyperlinkTextDestinations.item(list.selection[0].text);
					var name = d.name;
					d.remove()
					app.documents[0].hyperlinkTextDestinations.add (app.selection[0], {name: name});
				}
			}


			del_selected.onClick = function () {
				for (var i = list.selection.length-1; i > -1; i--) {
					if (list.selection[i].subItems[0].text == '') {
						app.activeDocument.hyperlinkTextDestinations.item (list.selection[i].text).remove();
						list.remove (list.selection[i]);
					}
				}
				if (list.items.length > 0) {
					list.selection = 0;
				}
				w.text = 'Text anchors (' + String(list.children.length).replace (/(\d)(?=(\d\d\d)+$)/g, "$1,") + ')';
			}


			sel_unused.onClick = function () {
				var i;
				var u = [];
				list.selection = null;
				for (i = list.items.length-1; i > -1; i--) {
					if (list.items[i].subItems[0].text == '') {
						u.push(i);
					}
				}
				list.selection = u;
			}


			rename_anchor.onClick = function () {
				renameAnchor (list, list.selection[0].text);
			}


			show_anchors.onClick = function () {
				showAnchors();
			}


			delete_frames.onClick = function () {
				if (app.documents[0].objectStyles.item('anchor display').isValid) {
					app.findObjectPreferences = null;
					app.findObjectPreferences.appliedObjectStyles = app.documents[0].objectStyles.item('anchor display');
					app.findChangeObjectOptions.objectType = ObjectTypes.TEXT_FRAMES_TYPE;
					var frames = app.documents[0].findObject();
					for (var i = frames.length-1; i >= 0; i--) {
						if (frames[i].isValid) {
							frames[i].remove();
						}
					}
					try {app.documents[0].objectStyles.item(styleName).remove();} catch(_){}
					try {app.documents[0].paragraphStyles.item(styleName).remove();} catch(_){}
					try {app.documents[0].swatches.item(styleName+' text').remove();} catch(_){}
					try {app.documents[0].swatches.item(styleName+' background').remove();} catch(_){}
				}
			}


		w.onResizing = w.onResize = function () {
			this.layout.resize();
		}


		w.onClose = function () {
			//$.writeln (list.preferredWidths)
			saveData ({
				location: [w.location[0], w.location[1]],
				size: [w.size.width, w.size.height],
				filter: w.filter.text
			});
		}


		w.onShow = function () {
			var previous = getPrevious();
			if (previous) {
				w.location = previous.location || [40, 40];
				w.size = previous.size || [600, 300];
				w.layout.resize();
				w.filter.text = previous.filter || '';
				if (w.filter.text !== '') {
					filterList();
				}
				list.selection = 0;
				list.notify();
			}
			w.text = 'Text anchors (' + String(list.children.length).replace (/(\d)(?=(\d\d\d)+$)/g, "$1,") + ')';
		}
			
		w.show ();
	}

	if (app.documents.length) {
		main();
	}

}());