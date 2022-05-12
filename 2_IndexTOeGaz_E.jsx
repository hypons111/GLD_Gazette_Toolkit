nameAndTitle();

indexToEgaz();


function nameAndTitle(){
	var myObject;
	//Make certain that user interaction (display of dialogs, etc.) is turned on.
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
	if(app.documents.length > 0){
		if(app.selection.length > 0){
			switch(app.selection[0].constructor.name){
				case "InsertionPoint":
				case "Character":
				case "Word":
				case "TextStyleRange":
				case "Line":
				case "Paragraph":
				case "TextColumn":
				case "Text":
				case "Cell":
				case "Column":
				case "Row":
				case "Table":
					myDisplayDialog();
					break;
				default:
					//Something was selected, but it wasn't a text object, so search the document.
					myFindChangeByList(app.documents.item(0));
			}
		}
		else{
			//Nothing was selected, so simply search the document.
			myFindChangeByList(app.documents.item(0));
		}
	}
	else{
		alert("No documents are open. Please open a document and try again.");
	}
}
function myDisplayDialog(){
	var myObject;
	var myDialog = app.dialogs.add({name:"FindChangeByList"});
	with(myDialog.dialogColumns.add()){
		with(dialogRows.add()){
			with(dialogColumns.add()){
				staticTexts.add({staticLabel:"Search Range:"});
			}
			var myRangeButtons = radiobuttonGroups.add();
			with(myRangeButtons){
				radiobuttonControls.add({staticLabel:"Document", checkedState:true});
				radiobuttonControls.add({staticLabel:"Selected Story"});
				if(app.selection[0].contents != ""){
					radiobuttonControls.add({staticLabel:"Selection", checkedState:true});
				}
			}			
		}
	}
	var myResult = myDialog.show();
	if(myResult == true){
		switch(myRangeButtons.selectedButton){
			case 0:
				myObject = app.documents.item(0);
				break;
			case 1:
				myObject = app.selection[0].parentStory;
				break;
			case 2:
				myObject = app.selection[0];
				break;
		}
		myDialog.destroy();
		myFindChangeByList(myObject);
	}
	else{
		myDialog.destroy();
	}
}
function myFindChangeByList(myObject){
	var myScriptFileName, myFindChangeFile, myFindChangeFileName, myScriptFile, myResult;
	var myFindChangeArray, myFindPreferences, myChangePreferences, myFindLimit, myStory;
	var myStartCharacter, myEndCharacter;
	var myFindChangeFile = myFindFile("/FindChangeSupport/大報 - 英文Index.txt")
	if(myFindChangeFile != null){
		myFindChangeFile = File(myFindChangeFile);
		var myResult = myFindChangeFile.open("r", undefined, undefined);
		if(myResult == true){
			//Loop through the find/change operations.
			do{
				myLine = myFindChangeFile.readln();
				//Ignore comment lines and blank lines.
				if((myLine.substring(0,4)=="text")||(myLine.substring(0,4)=="grep")||(myLine.substring(0,5)=="glyph")){
					myFindChangeArray = myLine.split("\t");
					//The first field in the line is the findType string.
					myFindType = myFindChangeArray[0];
					//The second field in the line is the FindPreferences string.
					myFindPreferences = myFindChangeArray[1];
					//The second field in the line is the ChangePreferences string.
					myChangePreferences = myFindChangeArray[2];
					//The fourth field is the range--used only by text find/change.
					myFindChangeOptions = myFindChangeArray[3];
                    switch(myFindType){
						case "text":
							myFindText(myObject, myFindPreferences, myChangePreferences, myFindChangeOptions);
							break;
						case "grep":
							myFindGrep(myObject, myFindPreferences, myChangePreferences, myFindChangeOptions);
							break;
						case "glyph":
							myFindGlyph(myObject, myFindPreferences, myChangePreferences, myFindChangeOptions);
							break;
					}
				}
			} while(myFindChangeFile.eof == false);
			myFindChangeFile.close();
		}
	}
}
function myFindText(myObject, myFindPreferences, myChangePreferences, myFindChangeOptions){
	//Reset the find/change preferences before each search.
	app.changeTextPreferences = NothingEnum.nothing;
	app.findTextPreferences = NothingEnum.nothing;
	var myString = "app.findTextPreferences.properties = "+ myFindPreferences + ";";
	myString += "app.changeTextPreferences.properties = " + myChangePreferences + ";";
	myString += "app.findChangeTextOptions.properties = " + myFindChangeOptions + ";";
	app.doScript(myString, ScriptLanguage.javascript);
	myFoundItems = myObject.changeText();
	//Reset the find/change preferences after each search.
	app.changeTextPreferences = NothingEnum.nothing;
	app.findTextPreferences = NothingEnum.nothing;
}
function myFindGrep(myObject, myFindPreferences, myChangePreferences, myFindChangeOptions){
	//Reset the find/change grep preferences before each search.
	app.changeGrepPreferences = NothingEnum.nothing;
	app.findGrepPreferences = NothingEnum.nothing;
    app.findGrepPreferences.appliedParagraphStyle = "Item-Tail-E"
	var myString = "app.findGrepPreferences.properties = "+ myFindPreferences + ";";
	myString += "app.changeGrepPreferences.properties = " + myChangePreferences + ";";
	myString += "app.findChangeGrepOptions.properties = " + myFindChangeOptions + ";";
	app.doScript(myString, ScriptLanguage.javascript);
	var myFoundItems = myObject.changeGrep();
	//Reset the find/change grep preferences after each search.
	app.changeGrepPreferences = NothingEnum.nothing;
	app.findGrepPreferences = NothingEnum.nothing;
}
function myFindGlyph(myObject, myFindPreferences, myChangePreferences, myFindChangeOptions){
	//Reset the find/change glyph preferences before each search.
	app.changeGlyphPreferences = NothingEnum.nothing;
	app.findGlyphPreferences = NothingEnum.nothing;
	var myString = "app.findGlyphPreferences.properties = "+ myFindPreferences + ";";
	myString += "app.changeGlyphPreferences.properties = " + myChangePreferences + ";";
	myString += "app.findChangeGlyphOptions.properties = " + myFindChangeOptions + ";";
	app.doScript(myString, ScriptLanguage.javascript);
	var myFoundItems = myObject.changeGlyph();
	//Reset the find/change glyph preferences after each search.
	app.changeGlyphPreferences = NothingEnum.nothing;
	app.findGlyphPreferences = NothingEnum.nothing;
}
function myFindFile(myFilePath){
	var myScriptFile = myGetScriptPath();
	var myScriptFile = File(myScriptFile);
	var myScriptFolder = myScriptFile.path;
	myFilePath = myScriptFolder + myFilePath;
	if(File(myFilePath).exists == false){
		//Display a dialog.
		myFilePath = File.openDialog("Choose the file containing your find/change list");
	}
	return myFilePath;
}
function myGetScriptPath(){
	try{
		myFile = app.activeScript;
	}
	catch(myError){
		myFile = myError.fileName;
	}
	return myFile;
}





function indexToEgaz() {
	if (app.layoutWindows.length == 0) return;
	var changeObject = app.documents[0];
	if (changeObject.hasOwnProperty('characters') && changeObject.characters.length == 0) return;
	var doc = app.documents[0];
	var style;
	var scriptVersion = app.scriptPreferences.version;
	app.scriptPreferences.version = 8.0;
	var options = app.findChangeGrepOptions.properties;
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
    
    // HHAHAHAHAHAHAHAAAA
    try { 
        app.findGrepPreferences.findWhat = ".+"; 
        app.changeGrepPreferences.properties = ({appliedLanguage:"英文: 英國"});
        app.activeDocument.changeGrep()
    } catch (e) {alert(e + ' at line ' + e.line)}
    
    // change to titlecase 
    try {
        app.findGrepPreferences.appliedParagraphStyle = 'H--Caps-E'
        app.findGrepPreferences.findWhat = ""; 
        myFind = app.activeDocument.findGrep();
        for(i = 0; i < myFind.length; i++) {
            myFind[i].changecase (ChangecaseMode.titlecase);
        }
    } catch (e) {alert(e + ' at line ' + e.line)}
    
    try {
        app.findGrepPreferences.appliedParagraphStyle = 'H--Small Caps-E'
        app.findGrepPreferences.findWhat = ""; 
        myFind = app.activeDocument.findGrep();
        for(i = 0; i < myFind.length; i++) {
            myFind[i].changecase (ChangecaseMode.titlecase);
        }
    } catch (e) {alert(e + ' at line ' + e.line)}    

//Added but HOLDED, Not Sure if case changed is needed for this style (2-9-2021)
/*    try {
        app.findGrepPreferences.appliedParagraphStyle = 'H-Roman-Center2-E'
        app.findGrepPreferences.findWhat = ""; 
        myFind = app.activeDocument.findGrep();
        for(i = 0; i < myFind.length; i++) {
            myFind[i].changecase (ChangecaseMode.titlecase);
        }
    } catch (e) {alert(e + ' at line ' + e.line)}    
*/

// wait for the conclusion to replace line 107 - 114
/*
    try {
        app.findGrepPreferences.appliedParagraphStyle = 'H--Small Caps-E'
        app.findGrepPreferences.findWhat = "";    
        app.findGrepPreferences.appliedParagraphStyle =  style;
        app.changeGrepPreferences.appliedParagraphStyle =  style;
        myFind = app.activeDocument.findGrep();
        for(i = 0; i < myFind.length; i++) {
            myFind[i].changecase (ChangecaseMode.titlecase);
        }
    } catch (e) {alert(e + ' at line ' + e.line)}    
*/    
    
    
    
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
    
     // added on 3-9-2021, found that hyphen was missing between 2 lines at GN5393 of issue35
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"(?i)Index Number"});
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'H--Small Caps-E', 'paragraphStyles');
		if (!style.isValid) throw Error(localize(({en:"Missing change pagraphstyle [%1] for query [%2]", de:"Fehlendes Ersetze-Absatzsformat [%1] bei Abfrage [%2]", fr:"La requête [%2] invoque en remplacement un style de paragraphe manquant : [%1]", ja_JP:"クエリ[%2]の置換形式に設定された段落スタイル[%1]が見つかりませんでした", nl:"Vervangende alineastijl [%1] mist voor zoekopdracht [%2]"}), 'H--Small Caps-E', '~GazInE_3A_Small Caps-E') );
		app.changeGrepPreferences.appliedParagraphStyle =  style;
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
   
    
    
	// Query [[~GazInE_2A_Caps-E]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, kanaSensitive:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'H--Caps-E', 'paragraphStyles');
		if (!style.isValid) throw Error(localize(({en:"Missing find pagraphstyle [%1] for query [%2]", de:"Fehlendes Such-Absatzsformat [%1] bei Abfrage [%2]", fr:"La requête [%2] invoque en recherche un style de paragraphe manquant : [%1]", ja_JP:"クエリ[%2]の検索形式に設定された段落スタイル[%1]が見つかりませんでした", nl:"Gezochte alineastijl [%1] mist voor zoekopdracht [%2]"}), 'H--Caps-E', '~GazInE_2A_Caps-E') );
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'H--Caps-E', 'paragraphStyles');
		if (!style.isValid) throw Error(localize(({en:"Missing change pagraphstyle [%1] for query [%2]", de:"Fehlendes Ersetze-Absatzsformat [%1] bei Abfrage [%2]", fr:"La requête [%2] invoque en remplacement un style de paragraphe manquant : [%1]", ja_JP:"クエリ[%2]の置換形式に設定された段落スタイル[%1]が見つかりませんでした", nl:"Vervangende alineastijl [%1] mist voor zoekopdracht [%2]"}), 'H--Caps-E', '~GazInE_2A_Caps-E') );
		app.changeGrepPreferences.appliedParagraphStyle =  style;
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
    // Query [[~GazInE_3A_Small Caps-E]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, kanaSensitive:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'H--Small Caps-E', 'paragraphStyles');
		if (!style.isValid) throw Error(localize(({en:"Missing find pagraphstyle [%1] for query [%2]", de:"Fehlendes Such-Absatzsformat [%1] bei Abfrage [%2]", fr:"La requête [%2] invoque en recherche un style de paragraphe manquant : [%1]", ja_JP:"クエリ[%2]の検索形式に設定された段落スタイル[%1]が見つかりませんでした", nl:"Gezochte alineastijl [%1] mist voor zoekopdracht [%2]"}), 'H--Small Caps-E', '~GazInE_3A_Small Caps-E') );
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'H--Small Caps-E', 'paragraphStyles');
		if (!style.isValid) throw Error(localize(({en:"Missing change pagraphstyle [%1] for query [%2]", de:"Fehlendes Ersetze-Absatzsformat [%1] bei Abfrage [%2]", fr:"La requête [%2] invoque en remplacement un style de paragraphe manquant : [%1]", ja_JP:"クエリ[%2]の置換形式に設定された段落スタイル[%1]が見つかりませんでした", nl:"Vervangende alineastijl [%1] mist voor zoekopdracht [%2]"}), 'H--Small Caps-E', '~GazInE_3A_Small Caps-E') );
		app.changeGrepPreferences.appliedParagraphStyle =  style;
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// 
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, kanaSensitive:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\t+"});
		style = getStyleByString(doc, 'H--Caps-E', 'paragraphStyles');
		if (!style.isValid) throw Error(localize(({en:"Missing find pagraphstyle [%1] for query [%2]", de:"Fehlendes Such-Absatzsformat [%1] bei Abfrage [%2]", fr:"La requête [%2] invoque en recherche un style de paragraphe manquant : [%1]", ja_JP:"クエリ[%2]の検索形式に設定された段落スタイル[%1]が見つかりませんでした", nl:"Gezochte alineastijl [%1] mist voor zoekopdracht [%2]"}), 'H--Caps-E', '~GazInE_2A_Caps-E') );
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({changeTo:"\\s"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// 
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, kanaSensitive:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\t+"});
		style = getStyleByString(doc, 'H--Small Caps-E', 'paragraphStyles');
		if (!style.isValid) throw Error(localize(({en:"Missing find pagraphstyle [%1] for query [%2]", de:"Fehlendes Such-Absatzsformat [%1] bei Abfrage [%2]", fr:"La requête [%2] invoque en recherche un style de paragraphe manquant : [%1]", ja_JP:"クエリ[%2]の検索形式に設定された段落スタイル[%1]が見つかりませんでした", nl:"Gezochte alineastijl [%1] mist voor zoekopdracht [%2]"}), 'H--Caps-E', '~GazInE_2A_Caps-E') );
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({changeTo:"\\s"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_6A_Date]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"^\\d+\\s[\\l\\u]+\\s20\\d+\\s*(?=\\t|\\r)"});
		style = getStyleByString(doc, 'Item-Tail-E', 'paragraphStyles');
		if (!style.isValid) throw Error(localize(({en:"Missing find pagraphstyle [%1] for query [%2]", de:"Fehlendes Such-Absatzsformat [%1] bei Abfrage [%2]", fr:"La requête [%2] invoque en recherche un style de paragraphe manquant : [%1]", ja_JP:"クエリ[%2]の検索形式に設定された段落スタイル[%1]が見つかりませんでした", nl:"Gezochte alineastijl [%1] mist voor zoekopdracht [%2]"}), 'Item-Tail-E', '~GazInE_6A_Date') );
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_6B_Council Chamber]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"^Council Chamber(?=\\t)"});
		style = getStyleByString(doc, 'Item-Tail-E', 'paragraphStyles');
		if (!style.isValid) throw Error(localize(({en:"Missing find pagraphstyle [%1] for query [%2]", de:"Fehlendes Such-Absatzsformat [%1] bei Abfrage [%2]", fr:"La requête [%2] invoque en recherche un style de paragraphe manquant : [%1]", ja_JP:"クエリ[%2]の検索形式に設定された段落スタイル[%1]が見つかりませんでした", nl:"Gezochte alineastijl [%1] mist voor zoekopdracht [%2]"}), 'Item-Tail-E', '~GazInE_6A_Date') );
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_6C_post 2line]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\r\\t"});
		style = getStyleByString(doc, 'Item-Tail-E', 'paragraphStyles');
		if (!style.isValid) throw Error(localize(({en:"Missing find pagraphstyle [%1] for query [%2]", de:"Fehlendes Such-Absatzsformat [%1] bei Abfrage [%2]", fr:"La requête [%2] invoque en recherche un style de paragraphe manquant : [%1]", ja_JP:"クエリ[%2]の検索形式に設定された段落スタイル[%1]が見つかりませんでした", nl:"Gezochte alineastijl [%1] mist voor zoekopdracht [%2]"}), 'Item-Tail-E', '~GazInE_6A_Date') );
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[移除後置空格]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\s+$"});
		app.changeGrepPreferences.properties = ({});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_5A_--]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"~_\\s+|~_"});
		app.changeGrepPreferences.properties = ({changeTo:"--"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_5B_space1]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\n\\t|\\n"});
		app.changeGrepPreferences.properties = ({changeTo:"\\s"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_5C_space2]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"[~m~>~f~|~S~s~<~/~.~3~4~% ~(]{1,}"});
		app.changeGrepPreferences.properties = ({changeTo:"\\s"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_6C1_--]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, kanaSensitive:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\r"});
		style = getStyleByString(doc, 'H-Roman-Center2-E', 'paragraphStyles');
		if (!style.isValid) throw Error(localize(({en:"Missing find pagraphstyle [%1] for query [%2]", de:"Fehlendes Such-Absatzsformat [%1] bei Abfrage [%2]", fr:"La requête [%2] invoque en recherche un style de paragraphe manquant : [%1]", ja_JP:"クエリ[%2]の検索形式に設定された段落スタイル[%1]が見つかりませんでした", nl:"Gezochte alineastijl [%1] mist voor zoekopdracht [%2]"}), 'H-Roman-Center2-E', '~GazInE_6C1_--') );
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({changeTo:"--"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_6C_--]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, kanaSensitive:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\r"});
		style = getStyleByString(doc, 'H--Small Caps-E', 'paragraphStyles');
		if (!style.isValid) throw Error(localize(({en:"Missing find pagraphstyle [%1] for query [%2]", de:"Fehlendes Such-Absatzsformat [%1] bei Abfrage [%2]", fr:"La requête [%2] invoque en recherche un style de paragraphe manquant : [%1]", ja_JP:"クエリ[%2]の検索形式に設定された段落スタイル[%1]が見つかりませんでした", nl:"Gezochte alineastijl [%1] mist voor zoekopdracht [%2]"}), 'H--Small Caps-E', '~GazInE_6C_--') );
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({changeTo:"--"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_6D_GN]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, kanaSensitive:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"(?i)(G\\.N\\.\\s\\d+)\\r"});
		app.changeGrepPreferences.properties = ({changeTo:"$1\\t\\t"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_8]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\r"});
		app.changeGrepPreferences.properties = ({changeTo:"\\s"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_9]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"(?i)--G\\.N\\. |G\\.N\\. "});
		app.changeGrepPreferences.properties = ({changeTo:"\\r"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_11]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"--(?=\\t)"});
		app.changeGrepPreferences.properties = ({});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_12]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"(?i)--as applied by section 26 of the--"});
		app.changeGrepPreferences.properties = ({changeTo:" as applied by section 26 of the "});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInC_14_)--(]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"(?<=\\))--(?=\\()"});
		app.changeGrepPreferences.properties = ({changeTo:" "});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// 
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"[~m~>~f~|~S~s~<~/~.~3~4~% ~(]{1,}"});
		app.changeGrepPreferences.properties = ({changeTo:"\\s"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;


    // Query [[~GazInE_12]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\<Of "});
		app.changeGrepPreferences.properties = ({changeTo:"of "});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_12]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\<Or "});
		app.changeGrepPreferences.properties = ({changeTo:"or "});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_12]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\<Is "});
		app.changeGrepPreferences.properties = ({changeTo:"is "});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_12]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\<The "});
		app.changeGrepPreferences.properties = ({changeTo:"the "});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_12]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\<On "});
		app.changeGrepPreferences.properties = ({changeTo:"on "});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_12]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\<In "});
		app.changeGrepPreferences.properties = ({changeTo:"in "});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_12]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\<And "});
		app.changeGrepPreferences.properties = ({changeTo:"and "});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_12]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\<To "});
		app.changeGrepPreferences.properties = ({changeTo:"to "});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_12]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\<For "});
		app.changeGrepPreferences.properties = ({changeTo:"for "});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_12]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\<Under "});
		app.changeGrepPreferences.properties = ({changeTo:"under "});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_12]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\<A "});
		app.changeGrepPreferences.properties = ({changeTo:"a "});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_12]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\<An "});
		app.changeGrepPreferences.properties = ({changeTo:"an "});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_12]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\<As "});
		app.changeGrepPreferences.properties = ({changeTo:"as "});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_12]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\<At "});
		app.changeGrepPreferences.properties = ({changeTo:"at "});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_12]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\<By "});
		app.changeGrepPreferences.properties = ({changeTo:"by "});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_12]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"to Kwa Wan "});
		app.changeGrepPreferences.properties = ({changeTo:"To Kwa Wan "});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_12]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"Ma on Shan "});
		app.changeGrepPreferences.properties = ({changeTo:"Ma On Shan "});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_12]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\<Mtr "});
		app.changeGrepPreferences.properties = ({changeTo:"MTR "});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_12]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"Pwp "});
		app.changeGrepPreferences.properties = ({changeTo:"PWP "});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInE_12]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"Iiia "});
		app.changeGrepPreferences.properties = ({changeTo:"IIIA "});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;


    // change to titlecase
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"(?<=(?i)Subsidiary Legislation )\\w+?\\>"});
        myFind = app.activeDocument.findGrep();
        for(i = 0; i < myFind.length; i++) {
            myFind[i].changecase (ChangecaseMode.UPPERCASE);
        }
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// change to titlecase
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"(?<=(?i)No\\. )\\w+?\\>"});
        myFind = app.activeDocument.findGrep();
        for(i = 0; i < myFind.length; i++) {
            myFind[i].changecase (ChangecaseMode.UPPERCASE);
        }
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// change to titlecase
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"(?<=(?i)Plan No\\. ).+?\\s"});
        myFind = app.activeDocument.findGrep();
        for(i = 0; i < myFind.length; i++) {
            myFind[i].changecase (ChangecaseMode.UPPERCASE);
        }
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
    // change to titlecase
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"(?<=(?i)\\(Chapter )\\w+?\\>"});
        myFind = app.activeDocument.findGrep();
        for(i = 0; i < myFind.length; i++) {
            myFind[i].changecase (ChangecaseMode.UPPERCASE);
        }
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
   // Query [[移除後置空格]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\s+$"});
		app.changeGrepPreferences.properties = ({});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findChangeGrepOptions.properties = options;
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	app.scriptPreferences.version = scriptVersion;

};

function getStyleByString(docOrGroup, string, property) {
	if (string == '[No character style]') return docOrGroup[property][0];
	if (string == '[No paragraph style]') return docOrGroup[property][0];
	if (string == 'NormalParagraphStyle') return docOrGroup[property][1];
	stringResult = string.match (/^(.*?[^\\]):(.*)$/);
	var styleName = (stringResult) ? stringResult[1] : string;
	styleName = styleName.replace (/\\:/g, ':');
	remainingString = (stringResult) ? stringResult[2] : '';
	var newProperty = (stringResult) ? property.replace(/s$/, '') + 'Groups' : property;
	var styleOrGroup = docOrGroup[newProperty].itemByName(styleName);
	if (remainingString.length > 0 && styleOrGroup.isValid) styleOrGroup = getStyleByString (styleOrGroup, remainingString, property);
	return styleOrGroup;
};
