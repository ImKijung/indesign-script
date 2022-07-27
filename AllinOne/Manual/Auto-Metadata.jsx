function AutoMetadata() {
	var win = new Window ("dialog");  
		win.alignChildren = "left";  
		win.KOR = win.add ("radiobutton", undefined, "KOR (한국)");  
		win.ENGUS = win.add ("radiobutton", undefined, "ENG-US (미주영어)");  
		win.ENG = win.add ("radiobutton", undefined, "ENG (영국영어)");  
		win.CFRA = win.add ("radiobutton", undefined, "FRA-US (미주프랑스)");  
		win.FRA = win.add ("radiobutton", undefined, "FRA (프랑스)");
		win.MSPA = win.add ("radiobutton", undefined, "SPA-US (미주스페인)");  
		win.SPA = win.add ("radiobutton", undefined, "SPA (스페인)");  
		win.BPOR = win.add ("radiobutton", undefined, "B-POR (브라질포르투칼)");  
		win.POR = win.add ("radiobutton", undefined, "POR (포르투칼)");  
		win.GER = win.add ("radiobutton", undefined, "DEU (독일)");  
		win.ITA = win.add ("radiobutton", undefined, "ITA (이탈리아)");  
		win.DUT = win.add ("radiobutton", undefined, "DUT (네덜란드)");  
		win.RUS = win.add ("radiobutton", undefined, "RUS (러시아)");  
		win.POL = win.add ("radiobutton", undefined, "POL (폴란드)");  
		win.DAN = win.add ("radiobutton", undefined, "DAN (덴마크)");  
		win.SWE = win.add ("radiobutton", undefined, "SWE (스웨덴)");  
		win.NOR = win.add ("radiobutton", undefined, "NOR (노르웨이");  
		win.FIN = win.add ("radiobutton", undefined, "FIN (핀란드)");  
		win.CHI = win.add ("radiobutton", undefined, "CHI (중국)");  
		win.quitBtn = win.add("button", undefined, "OK");    
		win.cancelBtn = win.add("button", undefined, "Cancel");    
		win.defaultElement = win.quitBtn;
		win.cancelElement = win.cancelBtn;

	if (win.show() == 1) {  
		if (win.KOR.value == true){
			var myDocument = app.activeDocument;  
			with (myDocument.metadataPreferences){ 
				description = "language=korea, voice1=참고, voice2=주의";  
				};
			}
		if (win.ENGUS.value == true){
			var myDocument = app.activeDocument;  
			with (myDocument.metadataPreferences){ 
				description = "language=engilsh-US, voice1=Note, voice2=Warning";  
				};
			}
		if (win.ENG.value == true){
			var myDocument = app.activeDocument;  
			with (myDocument.metadataPreferences){ 
				description = "language=engilsh, voice1=Note, voice2=Warning";  
				};
			}
			
		if (win.CFRA.value == true){
			var myDocument = app.activeDocument;  
			with (myDocument.metadataPreferences){ 
				description = "language=c-french, voice1=Remarque, voice2=Avertissement";  
				};
			}
		
		if (win.FRA.value == true){
			var myDocument = app.activeDocument;  
			with (myDocument.metadataPreferences){ 
				description = "language=french, voice1=Remarque, voice2=Avertissement";  
				};
			}
			
		if (win.MSPA.value == true){
			var myDocument = app.activeDocument;  
			with (myDocument.metadataPreferences){ 
				description = "language=m-spanish, voice1=Nota, voice2=Advertencia";  
				};
			}
		
		if (win.SPA.value == true){
			var myDocument = app.activeDocument;  
			with (myDocument.metadataPreferences){ 
				description = "language=spanish, voice1=Nota, voice2=Advertencia";  
				};
			}
			
		if (win.BPOR.value == true){
			var myDocument = app.activeDocument;  
			with (myDocument.metadataPreferences){ 
				description = "language=b-portuguese, voice1=Nota, voice2=Aviso";  
				};
			}
		
		if (win.POR.value == true){
			var myDocument = app.activeDocument;  
			with (myDocument.metadataPreferences){ 
				description = "language=portuguese, voice1=Nota, voice2=Aviso";  
				};
			}
			
		if (win.GER.value == true){
			var myDocument = app.activeDocument;  
			with (myDocument.metadataPreferences){ 
				description = "language=german, voice1=Hinweis, voice2=Warnung";  
				};
			}

		if (win.ITA.value == true){
			var myDocument = app.activeDocument;  
			with (myDocument.metadataPreferences){ 
				description = "language=italian, voice1=Nota, voice2=Avvertenza";  
				};
			}

		if (win.DUT.value == true){
			var myDocument = app.activeDocument;  
			with (myDocument.metadataPreferences){ 
				description = "language=dutch, voice1=Opmerking, voice2=Waarschuwing";  
				};
			}
		
		if (win.RUS.value == true){
			var myDocument = app.activeDocument;  
			with (myDocument.metadataPreferences){ 
				description = "language=russian, voice1=Примечание, voice2=Предупреждение";  
				};
			}
		
		if (win.POL.value == true){
			var myDocument = app.activeDocument;  
			with (myDocument.metadataPreferences){ 
				description = "language=polish, voice1=Uwaga, voice2=Ostrzeżenie";  
				};
			}
		
		if (win.DAN.value == true){
			var myDocument = app.activeDocument;  
			with (myDocument.metadataPreferences){ 
				description = "language=danish, voice1=Bemærk, voice2=Advarsel";  
				};
			}
		
		if (win.SWE.value == true){
			var myDocument = app.activeDocument;  
			with (myDocument.metadataPreferences){ 
				description = "language=swedish, voice1=Obs, voice2=Varning";  
				};
			}
		
		if (win.NOR.value == true){
			var myDocument = app.activeDocument;  
			with (myDocument.metadataPreferences){ 
				description = "language=norwegian, voice1=Merk, voice2=Advarsel";  
				};
			}
		
		if (win.FIN.value == true){
			var myDocument = app.activeDocument;  
			with (myDocument.metadataPreferences){ 
				description = "language=finnish, voice1=Huomautus, voice2=Varoitus";  
				};
			}
		
		if (win.CHI.value == true){
			var myDocument = app.activeDocument;  
			with (myDocument.metadataPreferences){ 
				description = "language=s-chinese, voice1=注意, voice2=警告";  
				};
			}
	}
		
	w = new Window ("dialog","OK", undefined, {closeButton: false});
	w.add ("statictext", undefined, "Complete");
	w.add ("button", undefined, "OK");
	w.show ( );
}