var osdDB = "C:/MMI_AutoMatching/MA-MMI_List.txt";
var txtFile = File (osdDB);
txtFile.encoding = 'UTF-8';
txtFile.open("r");
var fileCotentsString = txtFile.read();
var mmiFindChangeArray = fileCotentsString.split("\n");
txtFile.close();
var lineArray = mmiFindChangeArray[0].split("\t");
$.writeln(lineArray[0].replace("ID:",""));
var allIDs = lineArray[0].replace("ID:","");
allIDs *= 1;
$.writeln(allIDs + 1);

function fillZero(width, str) {
	return str.length >= width ? str:new Array(width-str.length+1).join('0')+str;//남는 길이만큼 0으로 채움
}