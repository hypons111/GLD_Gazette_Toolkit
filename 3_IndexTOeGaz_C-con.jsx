﻿//This script was auto generated by chainGREP.jsx
//chainGREP.jsx is provided by Gregor Fellenz https://www.publishingx.de/
//Download at https://www.publishingx.de/download/chain-grep
const  activeDocument = app.activeDocument


main();


for(var x = 0; x < activeDocument.pages.length; x ++) {
    deleteBlankPage(activeDocument.pages[x])
}

for(var x = 0; x < activeDocument.pages.length; x ++) {
    for(var y = 0; y < activeDocument.pages[x].textFrames.length; y ++) {
        spreadGNs(activeDocument.pages[x].textFrames[y])
    }
}

for(var x = 0; x < activeDocument.pages.length; x ++) {
    for(var y = 0; y < activeDocument.pages[x].textFrames.length; y ++) {
        relocateTextFrames(activeDocument.pages[x].textFrames[y])
    }
}




function main() {
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
//
    try {
		app.findChangeGrepOptions.properties = ({includeHiddenLayers:true, includeMasterPages:true, includeFootnotes:true, kanaSensitive:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'C--Text-Page', 'paragraphStyles');
		if (!style.isValid) throw Error(localize(({en:"Missing find pagraphstyle [%1] for query [%2]", de:"Fehlendes Such-Absatzsformat [%1] bei Abfrage [%2]", fr:"La requête [%2] invoque en recherche un style de paragraphe manquant : [%1]", ja_JP:"クエリ[%2]の検索形式に設定された段落スタイル[%1]が見つかりませんでした", nl:"Gezochte alineastijl [%1] mist voor zoekopdracht [%2]"}), 'C--Text-Page', '11111') );
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findChangeGrepOptions.properties = options;
	app.findGrepPreferences = NothingEnum.NOTHING;
//
    try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, kanaSensitive:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"~a\\r"});
		app.changeGrepPreferences.properties = ({});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
//
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, kanaSensitive:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\n"});
		app.changeGrepPreferences.properties = ({});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
 //
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, kanaSensitive:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"再登的公告\\r((.+\\r)|\\r)+.+"});
		app.changeGrepPreferences.properties = ({});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findChangeGrepOptions.properties = options;
	app.findGrepPreferences = NothingEnum.NOTHING;
   // 
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"~+"});
		app.changeGrepPreferences.properties = ({changeTo:"╱"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[任何空格轉為空格]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"[~m~>~f~|~S~s~<~/~.~3~4~% ~(]{1,}"});
		app.changeGrepPreferences.properties = ({changeTo:"\\s"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
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
	// Query [[~GazInC_2_-]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"~_ |~_"});
		app.changeGrepPreferences.properties = ({changeTo:"-"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// 
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\n"});
		app.changeGrepPreferences.properties = ({});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~~GazInE-con_1_page]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, kanaSensitive:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\d+~=\\d+|\\d+"});
		style = getStyleByString(doc, 'Text-Page', 'paragraphStyles');
		if (!style.isValid) throw Error(localize(({en:"Missing find pagraphstyle [%1] for query [%2]", de:"Fehlendes Such-Absatzsformat [%1] bei Abfrage [%2]", fr:"La requête [%2] invoque en recherche un style de paragraphe manquant : [%1]", ja_JP:"クエリ[%2]の検索形式に設定された段落スタイル[%1]が見つかりませんでした", nl:"Gezochte alineastijl [%1] mist voor zoekopdracht [%2]"}), 'Text-Page', '~~GazInE-con_1_page') );
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	// Query [[~~GazInC-con_2]] -- If you delete this comment you break the update function
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, kanaSensitive:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:".+"});
		style = getStyleByString(doc, 'Heading-with space', 'paragraphStyles');
		if (!style.isValid) throw Error(localize(({en:"Missing find pagraphstyle [%1] for query [%2]", de:"Fehlendes Such-Absatzsformat [%1] bei Abfrage [%2]", fr:"La requête [%2] invoque en recherche un style de paragraphe manquant : [%1]", ja_JP:"クエリ[%2]の検索形式に設定された段落スタイル[%1]が見つかりませんでした", nl:"Gezochte alineastijl [%1] mist voor zoekopdracht [%2]"}), 'Heading-with space', '~~GazInC-con_2') );
		app.findGrepPreferences.appliedParagraphStyle =  style;
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
    //
    try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\r"});
		style = getStyleByString(doc, 'Text-Subject', 'paragraphStyles');
		if (!style.isValid) throw Error(localize(({en:"Missing find pagraphstyle [%1] for query [%2]", de:"Fehlendes Such-Absatzsformat [%1] bei Abfrage [%2]", fr:"La requête [%2] invoque en recherche un style de paragraphe manquant : [%1]", ja_JP:"クエリ[%2]の検索形式に設定された段落スタイル[%1]が見つかりませんでした", nl:"Gezochte alineastijl [%1] mist voor zoekopdracht [%2]"}), 'Item--Head', '~GazInC_7_Gn') );
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({changeTo:"<br>"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;	
  //
    try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"<br>(\\d+\\t|\\d+~=\\d+\\t)"});
		app.changeGrepPreferences.properties = ({changeTo:"\\r$1"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;	
  //
    try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"<br>(~K)"});
		app.changeGrepPreferences.properties = ({changeTo:"\\r$1"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;	
  //
    try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"(<br>)\\t"});
		app.changeGrepPreferences.properties = ({changeTo:"$1"});
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


function spreadGNs(textFrame) {
    if(textFrame.overflows) {
        var nextThreadingTextFrame = activeDocument.pages[x].textFrames.add()
        if(textFrame.contents.toString().slice(0, 1) !== "\\w" && activeDocument.pages[x].textFrames.length > 1) {
            y += 1
        }   
        nextThreadingTextFrame.geometricBounds = ["48", "28p0", "495", "0"];  
        textFrame.nextTextFrame = nextThreadingTextFrame
        activeDocument.pages.add(LocationOptions.AFTER, activeDocument.pages[x])
        nextThreadingTextFrame.move(activeDocument.pages[x + 1])
    }
}


function relocateTextFrames(textFrame) {
    textFrame.geometricBounds = ["48", "28p0", "495", "0",];    //["上", "左", "下", "右",]
}


function deleteBlankPage(page) {
    if(page.allGraphics.length === 0 && page.textFrames.length === 0 && page.groups.length === 0){
        page.remove()
    } else if(page.textFrames.length === 1 && page.textFrames[0].startTextFrame.contents === "") {
        page.remove()
    }
}