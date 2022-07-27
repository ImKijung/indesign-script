// Show menu actions and their ids
// Peter Kahrel -- www.kahrel.plus.com
// See also http://kasyan.ho.com.ua/open_menu_item.html

#target indesign;

var action_object = create_object ();

var w = new Window ("dialog {text: 'Menu actions', orientation: 'row', alignChildren: 'top'}");
	var total = 0;
	var list = w.add ("listbox", undefined, "", {multiselect: true, 
																	numberOfColumns: 3, 
																	showHeaders: true, 
																	columnTitles: ["Name", "Area", "ID"], 
																	columnWidths: [250, 150, undefined]});
																
	list.maximumSize.height = w.maximumSize.height-100;
	
	var filter = w.add ('group {orientation: "column", alignChildren: "right"}');
		var name_group = filter.add ('group');
			name_group.add ('statictext {text: "Name:"}');
			var search_name = name_group.add ('edittext {preferredSize: [150, 20], active: true}');

		var area_group = filter.add ('group');
			area_group.add ('statictext {text: "Area:"}');
			var area_dropdown = area_group.add ('dropdownlist', undefined, action_object.areas);
				area_dropdown.preferredSize.width = 150;
				area_dropdown.selection = 0;

		var id_group = filter.add ('group');
			id_group.add ('statictext {text: "ID:"}');
			var search_id = id_group.add ('edittext {preferredSize: [150, 20]}');
		
		var sort_group = filter.add ('group');
			sort_group.add ('statictext {text: "Sort:"}');
			var sort_dropdown = sort_group.add ('dropdownlist', undefined, ['Name', 'Area', 'ID']);
				sort_dropdown.preferredSize.width = 150;
				sort_dropdown.selection = 1;
			

	// Populate the list
	
	for (var i in action_object.list)
		{
		with (list.add ('item', i))
			{
			subItems[0].text = action_object.list[i].area;
			subItems[1].text = action_object.list[i].id;
			}
		total++
		}
	w.text += "Menu actions ("+total+" items)";


	search_name.onChange = function()
		{
		total = 0;
		list.removeAll ();
		for (var i in action_object.list)
			{
			if (i.search(search_name.text) > -1)
				{
				with (list.add ('item', i))
					{
					subItems[0].text = action_object.list[i].area;
					subItems[1].text = action_object.list[i].id;
					}
				total++;
				}
			}
		w.text = "Menu actions ("+total+" items)";
		}
		

	area_dropdown.onChange = function ()
		{
		total = 0;
		list.removeAll ();
		for (var i in action_object.list)
			{
			if (action_object.list[i].area.indexOf(area_dropdown.selection.text) == 0 || area_dropdown.selection.text == '[All]')
				{
				with (list.add ('item', i))
					{
					subItems[0].text = action_object.list[i].area;
					subItems[1].text = action_object.list[i].id;
					}
				total++;
				}
			}
		w.text = "Menu actions  ("+total+" items)";
		}
	
	
	search_id.onChange = function ()
		{
		list.selection = null;
		var L = list.items.length;
		for (var i = 0; i < L; i++)
			{
			if (list.items[i].subItems[1].text == search_id.text)
				{
				list.selection = i;
				break;
				}
			}
		}
	
	list.onDoubleClick = function () {search_name.text = list.selection[0].text + "|" + action_object.list[list.selection[0].text].id;}
	
	sort_dropdown.onChange = function ()
		{
		var new_list = sort_multi_column_list (list, sort_dropdown.selection.text.toLowerCase());
		}
	
w.show ();


function create_object ()
	// We create two objects here: a key-value object for the dialog's listbox
	// and an array for the Area dropdown.
	{
	var list = {}, 
		areas = [], 
		known = {}, 
		actions = app.menuActions.everyItem().getElements();
		
	for(var i = 0; i < actions.length; i++)
		{
		if (!ignore (actions[i]))
			{
			list[actions[i].name] = {area: actions[i].area, id: actions[i].id};
			if (!known[actions[i].area])
				{
				areas.push (actions[i].area);
				known[actions[i].area] = true;
				}
			} // ignore
		}
	// Initially sort by area
	list = sort_object (list, "area");
	areas.sort();
	areas.unshift('[All]'); // add [All] at the beginning of the array
	return {list: list, areas: areas}
	}


function ignore (action)
	{
	return (Number(action.id) >= 57603 && Number(action.id) <= 60956)  // fonts and style names
				|| action.name.indexOf ('.indd') >= 0
				|| action.area.indexOf ('.js') >= 0
				|| action.area == 'Text Selection'
				|| action.area.indexOf ('Menu:Insert') > -1 ;
	}


function sort_object (list, key)
	{
	var temp = [], 
		temp2 = {};

	for (var i in list)
		temp.push ({name:i, area: list[i].area, id: list[i].id});
	temp.sort (function (a,b) {return a[key] > b[key]});
	for (var i = 0; i < temp.length; i++)
		temp2[temp[i].name] = {area: temp[i].area, id: temp[i].id};
	return temp2;
	} // sort_object


function sort_multi_column_list (list, key)
	{
	var temp = [], 
		L = list.items.length;
		
	for (var i = 0; i < L; i++)
		temp.push ({name: list.items[i].text, area: list.items[i].subItems[0].text, id: list.items[i].subItems[1].text});
	if (key == 'id')
		temp.sort (function (a,b) {return Number (a[key]) > Number (b[key])});
	else
		temp.sort (function (a,b) {return a[key] > b[key]});

	list.removeAll();
	for (var i = 0; i < L; i++)
		{
		with (list.add ('item', temp[i].name))
			{
			subItems[0].text = temp[i].area;
			subItems[1].text = temp[i].id;
			}
		}
	} // sort_multi_column_list


function gettime () {return new Date().getTime()}
function timediff (start) {return (new Date() - start) / 1000}
