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
	var myFindChangeFile = myFindFile("/FindChangeSupport/大報 - 中文Index.txt")
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

//限段式

//app.findGrepPreferences.appliedParagraphStyle = "Item-Tail-E"
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
	// Query [[~GazInC_1_衞]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:""});
		app.changeGrepPreferences.properties = ({changeTo:"衞"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// 
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"~+"});
		app.changeGrepPreferences.properties = ({changeTo:"╱"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInC_2_(]] -- If you delete this comment you break the update function    
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\n\\t\\("});
		app.changeGrepPreferences.properties = ({changeTo:"("});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInC_2_-]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"~_ |~_"});
		app.changeGrepPreferences.properties = ({changeTo:"-"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInC_3_tab at title1]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\t+"});
		style = getStyleByString(doc, 'H-Bold-Center', 'paragraphStyles');
		if (!style.isValid) throw Error(localize(({en:"Missing find pagraphstyle [%1] for query [%2]", de:"Fehlendes Such-Absatzsformat [%1] bei Abfrage [%2]", fr:"La requête [%2] invoque en recherche un style de paragraphe manquant : [%1]", ja_JP:"クエリ[%2]の検索形式に設定された段落スタイル[%1]が見つかりませんでした", nl:"Gezochte alineastijl [%1] mist voor zoekopdracht [%2]"}), 'H-Bold-Center', '~GazInC_4_date') );
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInC_3_tab at title2]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\t+"});
		style = getStyleByString(doc, 'H-Roman-Center', 'paragraphStyles');
		if (!style.isValid) throw Error(localize(({en:"Missing find pagraphstyle [%1] for query [%2]", de:"Fehlendes Such-Absatzsformat [%1] bei Abfrage [%2]", fr:"La requête [%2] invoque en recherche un style de paragraphe manquant : [%1]", ja_JP:"クエリ[%2]の検索形式に設定された段落スタイル[%1]が見つかりませんでした", nl:"Gezochte alineastijl [%1] mist voor zoekopdracht [%2]"}), 'H-Roman-Center', '~GazInC_4_date') );
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInC_3_tab at title3]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\t+"});
		style = getStyleByString(doc, 'H-Roman-Center2', 'paragraphStyles');
		if (!style.isValid) throw Error(localize(({en:"Missing find pagraphstyle [%1] for query [%2]", de:"Fehlendes Such-Absatzsformat [%1] bei Abfrage [%2]", fr:"La requête [%2] invoque en recherche un style de paragraphe manquant : [%1]", ja_JP:"クエリ[%2]の検索形式に設定された段落スタイル[%1]が見つかりませんでした", nl:"Gezochte alineastijl [%1] mist voor zoekopdracht [%2]"}), 'H-Roman-Center2', '~GazInC_4_date') );
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInC_4_date]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"^20\\d+年\\d+月\\d+日\\s*(?=\\t|\\r)"});
		style = getStyleByString(doc, 'Item-Tail', 'paragraphStyles');
		if (!style.isValid) throw Error(localize(({en:"Missing find pagraphstyle [%1] for query [%2]", de:"Fehlendes Such-Absatzsformat [%1] bei Abfrage [%2]", fr:"La requête [%2] invoque en recherche un style de paragraphe manquant : [%1]", ja_JP:"クエリ[%2]の検索形式に設定された段落スタイル[%1]が見つかりませんでした", nl:"Gezochte alineastijl [%1] mist voor zoekopdracht [%2]"}), 'Item-Tail', '~GazInC_4_date') );
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInC_4B_行政會議廳]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"^行政會議廳(?=\\t)"});
		style = getStyleByString(doc, 'Item-Tail', 'paragraphStyles');
		if (!style.isValid) throw Error(localize(({en:"Missing find pagraphstyle [%1] for query [%2]", de:"Fehlendes Such-Absatzsformat [%1] bei Abfrage [%2]", fr:"La requête [%2] invoque en recherche un style de paragraphe manquant : [%1]", ja_JP:"クエリ[%2]の検索形式に設定された段落スタイル[%1]が見つかりませんでした", nl:"Gezochte alineastijl [%1] mist voor zoekopdracht [%2]"}), 'Item-Tail', '~GazInC_4_date') );
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInC_5]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\n\\t"});
		app.changeGrepPreferences.properties = ({});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInC_6]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\)\\n\\("});
		app.changeGrepPreferences.properties = ({changeTo:") ("});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInC_6]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\n"});
		app.changeGrepPreferences.properties = ({});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInC_7_Gn]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\r"});
		style = getStyleByString(doc, 'Item--Head', 'paragraphStyles');
		if (!style.isValid) throw Error(localize(({en:"Missing find pagraphstyle [%1] for query [%2]", de:"Fehlendes Such-Absatzsformat [%1] bei Abfrage [%2]", fr:"La requête [%2] invoque en recherche un style de paragraphe manquant : [%1]", ja_JP:"クエリ[%2]の検索形式に設定された段落スタイル[%1]が見つかりませんでした", nl:"Gezochte alineastijl [%1] mist voor zoekopdracht [%2]"}), 'Item--Head', '~GazInC_7_Gn') );
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({changeTo:"\\t\\t"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInC_8]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\r\\t"});
		app.changeGrepPreferences.properties = ({changeTo:"\\t"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInC_9_--]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\r"});
		app.changeGrepPreferences.properties = ({changeTo:"--"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInC_10]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"--(第\\d+號公告)"});
		app.changeGrepPreferences.properties = ({changeTo:"\\r$1"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInC_11]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\t(第\\d+號公告)"});
		app.changeGrepPreferences.properties = ({changeTo:"\\r$1"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInC_13_田土轉易處]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"(?<=地政總署)--(?=\\(法律諮詢及田土轉易處\\))"});
		app.changeGrepPreferences.properties = ({changeTo:" "});
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
	// Query [[~GazInC_14_--第26條引用--]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"--第26條引用--"});
		app.changeGrepPreferences.properties = ({changeTo:"第26條引用"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInC_]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"-+(?=\\r)"});
		app.changeGrepPreferences.properties = ({});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInC_15_(]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"~K(?=\\()"});
		app.changeGrepPreferences.properties = ({changeTo:"$0 "});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInC_15_)]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"(?<=\\))~K"});
		app.changeGrepPreferences.properties = ({changeTo:" $0"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findChangeGrepOptions.properties = options;
	app.findGrepPreferences = NothingEnum.NOTHING;
	//
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true, wholeWord:false});
		app.findGrepPreferences.properties = ({findWhat:"-{4}"});
		app.changeGrepPreferences.properties = ({changeTo:""});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findChangeGrepOptions.properties = options;
	app.findGrepPreferences = NothingEnum.NOTHING;
	// Query [[~GazInC_15_)]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"(?<=\\t\\t)(.+?)\\t(.+?)\\t(.+?)\\t(.+?)\\t"});
		app.changeGrepPreferences.properties = ({changeTo:"$1\\t$2$3$4\\t"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findChangeGrepOptions.properties = options;
	app.findGrepPreferences = NothingEnum.NOTHING;
   //
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
