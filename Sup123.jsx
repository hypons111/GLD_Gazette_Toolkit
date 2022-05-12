const activeDocument = app.activeDocument

main();

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
	app.findChangeGrepOptions.properties = ({includeFootnotes:true, includeMasterPages:false, kanaSensitive:true, widthSensitive:true});
//

//
try {
    app.findTextPreferences = NothingEnum.NOTHING;
    app.changeTextPreferences = NothingEnum.NOTHING;
    app.findTextPreferences.findWhat = "^p<0016>";
    myFind = activeDocument.findText();
    if(myFind.length > 0){
        app.findTextPreferences.findWhat = "^p<0016>";
        var changed = 0
        for(i = 0; i < myFind.length; i ++) {
     //       alert("1")
            app.selection = myFind[i]
     //       alert("2")
            app.copy()
     //       alert("3")
            app.changeTextPreferences.properties = ({changeTo:"^C^p"});
     //       alert("4")
            myFind[i].changeText();
     //       alert("5")
            changed ++
        }
    }
} catch(e) {
} finally {
    alert("HAHAHAHAHHAHA")
}
//--
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true,  includeMasterPages:false, kanaSensitive:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\r+"});
		style = getStyleByString(doc, '_Text folder:_Part Line', 'paragraphStyles');
		if (!style.isValid) throw Error(localize(({en:"Missing find pagraphstyle [%1] for query [%2]", de:"Fehlendes Such-Absatzsformat [%1] bei Abfrage [%2]", fr:"La requête [%2] invoque en recherche un style de paragraphe manquant : [%1]", ja_JP:"クエリ[%2]の検索形式に設定された段落スタイル[%1]が見つかりませんでした", nl:"Gezochte alineastijl [%1] mist voor zoekopdracht [%2]"}), '_Text folder:_Part Line', '1234') );
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({changeTo:"\\r"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findChangeGrepOptions.properties = options;
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;    
//--
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true,  includeMasterPages:false, kanaSensitive:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"\\r+"});
		style = getStyleByString(doc, 'Text folder:Part Line', 'paragraphStyles');
		if (!style.isValid) throw Error(localize(({en:"Missing find pagraphstyle [%1] for query [%2]", de:"Fehlendes Such-Absatzsformat [%1] bei Abfrage [%2]", fr:"La requête [%2] invoque en recherche un style de paragraphe manquant : [%1]", ja_JP:"クエリ[%2]の検索形式に設定された段落スタイル[%1]が見つかりませんでした", nl:"Gezochte alineastijl [%1] mist voor zoekopdracht [%2]"}), '_Text folder:_Part Line', '1234') );
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({changeTo:"\\r"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findChangeGrepOptions.properties = options;
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;    
 //--
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true,  includeMasterPages:false, kanaSensitive:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"(?<![\\l\\u]|\\d)(\\))(\\()(?![\\l\\u]|\\d)"});
		app.changeGrepPreferences.properties = ({changeTo:"$1~4$2"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
 //--
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true,  includeMasterPages:false, kanaSensitive:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"(~K\\))(\\d)"});
		app.changeGrepPreferences.properties = ({changeTo:"$1 $2"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
 //--
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true,  includeMasterPages:false, kanaSensitive:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"(\\d)(\\(~K)"});
		app.changeGrepPreferences.properties = ({changeTo:"$1 $2"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
 //--
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true,  includeMasterPages:false, kanaSensitive:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:" ~_~_"});
		app.changeGrepPreferences.properties = ({changeTo:"~_~_"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
 //--
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true,  includeMasterPages:false, kanaSensitive:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"~P"});
		app.changeGrepPreferences.properties = ({changeTo:"\\r"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;


// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Heading:_Heading-Bold', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Heading:_Heading-Bold', 'paragraphStyles');
        app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Heading:_Heading-Part-Bold', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Heading:_Heading-Part-Bold', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Heading:_Heading-Part-Bold 2', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Heading:_Heading-Part-Bold 2', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Heading:_Heading-PDiv-Bold', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Heading:_Heading-PDiv-Bold', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Heading:_Heading-PDiv-Bold-2', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Heading:_Heading-PDiv-Bold-2', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Heading:_Heading-Sch Right', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Heading:_Heading-Sch Right', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Heading:_Heading-Schedule', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Heading:_Heading-Schedule', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Heading:_Heading-Schedule 2', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Heading:_Heading-Schedule 2', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Heading:_Heading-SchPart-B', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Heading:_Heading-SchPart-B', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Heading:_H-SubDiv-Bold', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Heading:_H-SubDiv-Bold', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Heading:Heading-Bold', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Heading:Heading-Bold', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Heading:Heading-Part-Bold', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Heading:Heading-Part-Bold', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Heading:Heading-Part-Bold 2', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Heading:Heading-Part-Bold 2', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Heading:Heading-PDiv-Bold', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Heading:Heading-PDiv-Bold', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Heading:Heading-PDiv-Bold-2', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Heading:Heading-PDiv-Bold-2', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Heading:Heading-Sch Right', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Heading:Heading-Sch Right', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Heading:Heading-Schedule', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Heading:Heading-Schedule', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Heading:Heading-Schedule 2', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Heading:Heading-Schedule 2', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Heading:Heading-SchPart-B', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Heading:Heading-SchPart-B', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Heading:H-SubDiv-Bold', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Heading:H-SubDiv-Bold', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Other:_Amend', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Other:_Amend', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Other:_Enacted by the Leg Co', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Other:_Enacted by the Leg Co', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Other:_Ex M-Text', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Other:_Ex M-Text', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Other:_Item-Council', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Other:_Item-Council', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Other:_Item-Date', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Other:_Item-Date', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Other:_Item-Explanatory', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Other:_Item-Explanatory', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Other:_Item-Title', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Other:_Item-Title', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Other:_Text_FP', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Other:_Text_FP', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Other:_Text_Resolved', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Other:_Text_Resolved', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Other:_Text-centre', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Other:_Text-centre', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Other:Amend', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Other:Amend', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Other:Enacted by the Leg Co', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Other:Enacted by the Leg Co', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Other:Ex M-Text', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Other:Ex M-Text', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Other:Item-Council', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Other:Item-Council', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Other:Item-Date', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Other:Item-Date', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Other:Item-Explanatory', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Other:Item-Explanatory', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Other:Item-Title', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Other:Item-Title', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Other:Text_FP', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Other:Text_FP', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Other:Text_Resolved', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Other:Text_Resolved', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Other:Text-centre', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Other:Text-centre', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '**Other:Title_Chi', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '**Other:Title_Chi', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Ind tab text:_Ind_Head-Part-Bold', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Ind tab text:_Ind_Head-Part-Bold', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Ind tab text:_Ind_Head-PDiv-Bold', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Ind tab text:_Ind_Head-PDiv-Bold', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Ind tab text:_Ind_Text-center', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Ind tab text:_Ind_Text-center', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Ind tab text:_Ind-11', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Ind tab text:_Ind-11', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Ind tab text:_Ind-13', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Ind tab text:_Ind-13', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Ind tab text:_Ind-15', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Ind tab text:_Ind-15', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Ind tab text:_Ind-17', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Ind tab text:_Ind-17', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Ind tab text:_Ind-5', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Ind tab text:_Ind-5', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Ind tab text:_Ind-7', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Ind tab text:_Ind-7', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Ind tab text:_Ind-9', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Ind tab text:_Ind-9', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Ind tab text:_Ind-H-SubDiv-Bold', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Ind tab text:_Ind-H-SubDiv-Bold', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Ind tab text:_Ind_Head-Part-Bold', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Ind tab text:_Ind_Head-Part-Bold', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Note:_Note-Bold', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Note:_Note-Bold', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Note:_Note-Text a2', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Note:_Note-Text a2', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Text folder:_Centre-HBold_Ind3', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Text folder:_Centre-HBold_Ind3', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Text folder:_Centre-HBold_Ind7', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Text folder:_Centre-HBold_Ind7', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Text folder:_hang-3', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Text folder:_hang-3', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Text folder:_hang-5', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Text folder:_hang-5', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Text folder:_hang-7', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Text folder:_hang-7', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Text folder:_Head-Bold', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Text folder:_Head-Bold', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Text folder:_Head-Bold_Ind-5', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Text folder:_Head-Bold_Ind-5', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Text folder:_Head-Bold_Ind-7', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Text folder:_Head-Bold_Ind-7', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Text folder:_Part Line', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Text folder:_Part Line', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Text folder:_Part Line_Ind-3', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Text folder:_Part Line_Ind-3', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Text folder:_Text', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Text folder:_Text', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Text folder:_Text_ind-p1', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Text folder:_Text_ind-p1', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Text folder:_Text-11', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Text folder:_Text-11', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Text folder:_Text-13', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Text folder:_Text-13', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Text folder:_Text-3', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Text folder:_Text-3', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Text folder:_Text-3_ind-p1', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Text folder:_Text-3_ind-p1', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Text folder:_Text-5', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Text folder:_Text-5', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Text folder:_Text-5 Bold', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Text folder:_Text-5 Bold', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Text folder:_Text-5_ind-p1', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Text folder:_Text-5_ind-p1', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Text folder:_Text-7', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Text folder:_Text-7', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Text folder:_Text-7 Bold', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Text folder:_Text-7 Bold', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Text folder:_Text-7_ind-p1', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Text folder:_Text-7_ind-p1', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Text folder:_Text-9', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Text folder:_Text-9', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, '_Text folder:_Text-Sch', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, '_Text folder:_Text-Sch', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Article:_Art-Centre Bold', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Article:_Art-Centre Bold', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Article:_Art-Text', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Article:_Art-Text', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Article:_Ind-5 (+9)', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Article:_Ind-5 (+9)', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Article:_Ind-7 (+9)', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Article:_Ind-7 (+9)', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Article:_Text-5 (+9)', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Article:_Text-5 (+9)', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Article:Art-Centre Bold', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Article:Art-Centre Bold', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Article:Art-Text', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Article:Art-Text', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Article:Ind-5 (+9)', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Article:Ind-5 (+9)', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Article:Ind-7 (+9)', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Article:Ind-7 (+9)', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Article:Text-5 (+9)', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Article:Text-5 (+9)', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Ind tab text:Ind_Head-Part-Bold', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Ind tab text:Ind_Head-Part-Bold', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Ind tab text:Ind_Head-PDiv-Bold', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Ind tab text:Ind_Head-PDiv-Bold', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Ind tab text:Ind_Text-center', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Ind tab text:Ind_Text-center', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Ind tab text:Ind-11', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Ind tab text:Ind-11', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Ind tab text:Ind-13', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Ind tab text:Ind-13', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Ind tab text:Ind-15', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Ind tab text:Ind-15', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Ind tab text:Ind-17', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Ind tab text:Ind-17', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Ind tab text:Ind-5', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Ind tab text:Ind-5', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Ind tab text:Ind-7', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Ind tab text:Ind-7', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Ind tab text:Ind-9', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Ind tab text:Ind-9', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Ind tab text:Ind-H-SubDiv-Bold', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Ind tab text:Ind-H-SubDiv-Bold', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Ind tab text:Ind_Head-Part-Bold', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Ind tab text:Ind_Head-Part-Bold', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:_N-Heading_3p', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:_N-Heading_3p', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:_N-Heading_5p', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:_N-Heading_5p', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:_N-Heading_7p', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:_N-Heading_7p', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:_N-Heading_9p', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:_N-Heading_9p', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:_N-text_(3a)', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:_N-text_(3a)', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:_N-text_(5a)', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:_N-text_(5a)', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:_N-text_(7a)', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:_N-text_(7a)', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:_N-text_(9a)', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:_N-text_(9a)', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:_N-text_11', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:_N-text_11', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:_N-text_13', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:_N-text_13', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:_N-text_3', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:_N-text_3', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:_N-text_5', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:_N-text_5', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:_N-text_7', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:_N-text_7', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:_N-text_9', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:_N-text_9', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:_N-text_Ind-11', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:_N-text_Ind-11', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:_N-text_Ind-13', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:_N-text_Ind-13', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:_N-text_Ind-15', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:_N-text_Ind-15', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:_N-text_Ind-7', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:_N-text_Ind-7', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:_N-text_Ind-9', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:_N-text_Ind-9', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:N-Heading_3p', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:N-Heading_3p', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:N-Heading_5p', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:N-Heading_5p', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:N-Heading_7p', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:N-Heading_7p', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:N-Heading_9p', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:N-Heading_9p', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:Note-Bold', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:Note-Bold', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:Note-Text a2', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:Note-Text a2', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:N-text_(3a)', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:N-text_(3a)', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:N-text_(5a)', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:N-text_(5a)', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:N-text_(7a)', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:N-text_(7a)', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:N-text_(9a)', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:N-text_(9a)', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:N-text_11', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:N-text_11', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:N-text_13', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:N-text_13', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:N-text_3', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:N-text_3', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:N-text_5', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:N-text_5', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:N-text_7', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:N-text_7', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:N-text_9', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:N-text_9', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:N-text_Ind-11', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:N-text_Ind-11', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:N-text_Ind-13', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:N-text_Ind-13', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:N-text_Ind-15', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:N-text_Ind-15', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:N-text_Ind-7', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:N-text_Ind-7', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Note:N-text_Ind-9', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Note:N-text_Ind-9', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Text folder:Centre-HBold_Ind3', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Text folder:Centre-HBold_Ind3', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Text folder:Centre-HBold_Ind7', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Text folder:Centre-HBold_Ind7', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Text folder:hang-3', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Text folder:hang-3', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Text folder:hang-5', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Text folder:hang-5', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Text folder:hang-7', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Text folder:hang-7', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Text folder:Head-Bold', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Text folder:Head-Bold', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Text folder:Head-Bold_Ind-5', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Text folder:Head-Bold_Ind-5', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Text folder:Head-Bold_Ind-7', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Text folder:Head-Bold_Ind-7', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Text folder:Part Line', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Text folder:Part Line', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Text folder:Part Line_Ind-3', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Text folder:Part Line_Ind-3', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Text folder:Text', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Text folder:Text', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Text folder:Text_ind-p1', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Text folder:Text_ind-p1', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Text folder:Text-11', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Text folder:Text-11', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Text folder:Text-13', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Text folder:Text-13', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Text folder:Text-3', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Text folder:Text-3', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Text folder:Text-3_ind-p1', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Text folder:Text-3_ind-p1', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Text folder:Text-5', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Text folder:Text-5', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Text folder:Text-5 Bold', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Text folder:Text-5 Bold', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Text folder:Text-5_ind-p1', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Text folder:Text-5_ind-p1', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Text folder:Text-7', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Text folder:Text-7', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Text folder:Text-7 Bold', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Text folder:Text-7 Bold', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Text folder:Text-7_ind-p1', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Text folder:Text-7_ind-p1', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Text folder:Text-9', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Text folder:Text-9', 'paragraphStyles');
app.changeGrepPreferences.appliedParagraphStyle =  style;
changeObject.changeGrep();
} catch (e) {}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// 
try {
app.findGrepPreferences.properties = ({});
style = getStyleByString(doc, 'Text folder:Text-Sch', 'paragraphStyles');
		app.findGrepPreferences.appliedParagraphStyle =  style;
		app.changeGrepPreferences.properties = ({});
		style = getStyleByString(doc, 'Text folder:Text-Sch', 'paragraphStyles');
		app.changeGrepPreferences.appliedParagraphStyle =  style;
		changeObject.changeGrep();
	} catch (e) {}
	app.findChangeGrepOptions.properties = options;
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
//end of styles changes

//--
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true,  includeMasterPages:false, kanaSensitive:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"(?<=註釋)\\r"});
		app.changeGrepPreferences.properties = ({changeTo:"\\r1.\\t", fillColor:"None"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
//--
	try {
		app.findChangeGrepOptions.properties = ({includeFootnotes:true,  includeMasterPages:false, kanaSensitive:true, widthSensitive:true});
		app.findGrepPreferences.properties = ({findWhat:"(?<=Explanatory Note)\\r"});
		app.changeGrepPreferences.properties = ({changeTo:"\\r1.\\t", fillColor:"None"});
		changeObject.changeGrep();
	} catch (e) {alert(e + ' at line ' + e.line)}
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
