var myDoc = app.activeDocument;
var indexes = myDoc.indexes;
for (var i = 0 ; i < indexes.length ; i++) {
	var index = indexes[i];
	var topics = index.allTopics;
	for (var j=0; j <topics.length; j++) {
		var topic = topics[j]
		var pageReference = topic.pageReferences
		for (var k=0; k < pageReference.length; k++) {
				var topicName = "topicName=" + topic.name;
				// parent Name 이 없으면 끝, 있으면 출력
				if (topic.parent.name == "") {
					$.writeln(j + "-" + k + ":" + topicName);
				} else {
					var parentName = "parentName=" + topic.parent.name;
					if (topic.parent.parent.name == "") {
						$.writeln(j + "-" + k + ":" + topicName + ";" + parentName);
					} else {
						var acientName = "acientName=" + topic.parent.parent.name;
						$.writeln(j + "-" + k + ":" + topicName + ";" + parentName + ";" + acientName);
					}
				}
			}
		}
}