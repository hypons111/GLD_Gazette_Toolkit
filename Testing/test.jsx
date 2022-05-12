var file1 = app.activeDocument
file1.pages.add()


/*
var activeDocument = app.activeDocument

//makeThreadTextFrames()

        var image = activeDocument.pages[0].textFrames[0].rectangles[0].allGraphics[0]
        alert(image.geometricBounds[0])
        var newBottom = image.geometricBounds[2] - image.geometricBounds[0] - 119
            image.geometricBounds = ["-119p", image.geometricBounds[1], image.geometricBounds[2], image.geometricBounds[3]];


function setTitle() {
    app.findGrepPreferences = NothingEnum.NOTHING;
    app.changeGrepPreferences = NothingEnum.NOTHING;
    app.findChangeGrepOptions.properties = ({includeFootnotes: false, includeMasterPages: false, includeHiddenLayers: false, wholeWord: false});
    style = getStyleByString(activeDocument, '**Other:Title_Chi', 'paragraphStyles');
    app.findGrepPreferences.appliedParagraphStyle =  style;
    app.selection = activeDocument.findGrep();
    app.copy()

	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
    app.findChangeGrepOptions.properties = ({includeFootnotes: true, includeMasterPages: true, includeHiddenLayers: true, wholeWord: true});
    app.findGrepPreferences.properties = ({findWhat:"立法會決議", pointSize:9});
    app.selection = activeDocument.findGrep();
    app.paste()
//    app.findGrepPreferences.properties = ({findWhat:"立法會決議"});
//    app.selection = activeDocument.findGrep();

//    activeDocument.masterSpreads.item(0).groups[0].ungroup()
//    activeDocument.masterSpreads.item(0).textFrames[2].contents

	app.findChangeGrepOptions.properties = options;
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	app.scriptPreferences.version = scriptVersion;

}
    */

function makeThreadTextFrames() {
    for(var x = 0; x < activeDocument.pages.length; x ++) {
        for(var y = 0; y < activeDocument.pages[x].textFrames.length; y ++) {
            var textFrame = activeDocument.pages[x].textFrames[y]
            if(textFrame.overflows) {
                var nextThreadingTextFrame = activeDocument.pages[x].textFrames.add()
                textFrame.nextTextFrame = nextThreadingTextFrame
                if(x % 2 === 0) {
                    activeDocument.pages.add()
                } else {
                    activeDocument.pages.add()
                    activeDocument.pages.add()
                }
                nextThreadingTextFrame.move(activeDocument.pages[x + 2])
                nextThreadingTextFrame.geometricBounds = ["0", "0", "37p4", "28p"]; 
            }
        }
    }
}





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









/*
兩 無
751(3)
798(2)

兩 有
751(1)
747(7)
746(2)
751(1)

三 無
744(3)
745(2)(b)
797(2)(b)

三 有
796(3)



moveLeft()
function moveLeft() {
    for(var x = 0; x < activeDocument.pages.length; x ++) {
        var image = activeDocument.pages[x].textFrames[0].rectangles[0].allGraphics[0]
        var newLeft = image.geometricBounds[3] - image.geometricBounds[1] - 1.66666666666666
        image.geometricBounds = [image.geometricBounds[0], "-1p8", image.geometricBounds[2], newLeft];        
    }    
}

*/






/*
    for(var x = 0; x < 4; x ++) {
        originalGeometricBoundsArray.push(image.geometricBounds[x])
    }
    for(var x = 0; x < 4; x ++) {
        newGeometricBoundsArray.push(originalGeometricBoundsArray[0])
        newGeometricBoundsArray.push("-1p8")
        newGeometricBoundsArray.push(originalGeometricBoundsArray[2] - originalGeometricBoundsArray[0])
        newGeometricBoundsArray.push(originalGeometricBoundsArray[3] - originalGeometricBoundsArray[1])
    }    
    
    
for(var x = 0; x < activeDocument.pages.length; x ++) {
    deleteHeader(activeDocument.pages[x])
}

function deleteHeader(page) {
    for(var y = 0; y < page.textFrames.length; y ++) {
        switch(page.textFrames[y].contents) {
            case "委任令" :
            case "公  告" :
            case "招  標" :
            case "再登的公告" :
            case "本期憲報全文完" :
            case "\t指數數值\r特選材料" :
            case "\t指數數值\r工資" :
            case "\t指數數值" :
            case "APPOINTMENTS" :
            case "NOTICES" :
            case "TENDERS" :
            case "NOTICES REPEATED" :
            case "End of Main Gazette of this issue." :
            case "\tIndex Numbers\rSelected materials" :
            case "\tIndex Numbers\rLabour" :
            case "\tIndex Numbers" :
                page.textFrames[y].remove()
        }
    }
}



const activeDocument = app.activeDocument
const issue = activeDocument.name.toString()[3] + activeDocument.name.toString()[4]
const path = activeDocument.fullName.toString().slice(0, 55)
const leftTable = activeDocument.textFrames[1].tables[0]


//const issue = leftTable.cells.item(0)
const date = leftTable.cells.item(1)


function mainGazette(fullName, name) {
    this.fullName = fullName
    this.name = name
    this.fileName = []
    this.numberOfPage = []
    this.getFileName = function (x) {
        for(; x <= 154; x += 4) {
            this.fileName.push(leftTable.cells.item(x))
        }
    }
    this.getNumberOfPage = function (x) {
        for(; x <= 154; x += 4) {
            this.numberOfPage.push(leftTable.cells.item(x))
        }
    }
}


const chinese = new mainGazette("Chinese", "C")
chinese.getFileName(6)
chinese.getNumberOfPage(7)

const english = new mainGazette("English", "E")
english.getFileName(8)
english.getNumberOfPage(9)


for(var x = 21252; x < 21254; x ++) {
    var maFile = File(path + x + '.indd')
    if(maFile.exists) {
        alert(maFile.pages.length)
    }
}







const rightTable = activeDocument.textFrames[0].tables[0]
const cover = rightTable.cells.item(2)
const imprint = rightTable.cells.item(5)
const chineseContants = rightTable.cells.item(8)
const chineseText = rightTable.cells.item(11)
const englishContants = rightTable.cells.item(14)
const englishText = rightTable.cells.item(17)


chinese.fileName[0].contents = "1"
chinese.numberOfPage[0].contents = "2"
english.fileName[0].contents = "3"
english.numberOfPage[0].contents = "4"

cover.contents = "Cover"
imprint.contents = "Imprint"
chineseContants.contents = "chineseContants"
chineseText.contents = "chineseText"
englishContants.contents = "englishContants"
englishText.contents = "englishText"
*/













/*
#targetengine "session"


app.activeDocument.eventListeners.everyItem().remove()
app.activeDocument.eventListeners.add("afterSave", changeXXX)
app.activeDocument.eventListeners.add("afterSaveAs", changeXXX)


function changeXXX() {
    alert("a")
    app.findChangeGrepOptions.properties = ({includeMasterPages: true, kanaSensitive:true, widthSensitive:true});
    app.findGrepPreferences.findWhat = "(?<=第)xxx";
    app.changeGrepPreferences.changeTo = app.activeDocument.name.slice(0, 3).toString()
    app.activeDocument.changeGrep();

    app.findGrepPreferences.findWhat = "(?<=G.N. )xxx";
    app.changeGrepPreferences.changeTo = app.activeDocument.name.slice(0, 3).toString()
    app.activeDocument.changeGrep();
    
    app.findGrepPreferences.findWhat = "xxx(?=\n-)";
    app.changeGrepPreferences.changeTo = app.activeDocument.name.slice(0, 3).toString()
    app.activeDocument.changeGrep();
    
    
    app.activeDocument.eventListeners.everyItem().remove()
}
*/







/*
function OBJ(name) {
    this.name = name
    this.getTextFrame = function () {
        return app.activeDocument.pages[0].textFrames[0]
    }
    this.stateSwitch = function () {
        app.findChangeGrepOptions.properties = ({includeMasterPages: false});
        app.findGrepPreferences.findWhat = "xxx";
        app.changeGrepPreferences.changeTo = app.activeDocument.name.slice(0, 3)
        app.activeDocument.changeGrep();
    }
}
*/

/*
#targetengine "session"

main()

function main() {
    app.activeDocument.eventListeners.everyItem().remove()
    app.activeDocument.eventListeners.add("beforeSave", execute)
}

function execute(myEvent) {
    app.activeDocument.pages[0].textFrames[0].contents = outer(app.activeDocument.pages[0].textFrames[0].contents)   
}

function outer(state) {
    if(state === "true") {
        return "false"
    }
    return "true"
}
*/







/*
function outer(state) {
    var documentState = state
    return function inner() {
        if(documentState = true) {
            return false
        }
        return true
    }
}
*/



/*
app.selection = app.activeDocument.findText();
app.copy()
app.changeTextPreferences.properties = ({changeTo:"^C^p"});
app.activeDocument.changeText();
*/


/* demo
app.findTextPreferences = NothingEnum.NOTHING;
app.changeTextPreferences = NothingEnum.NOTHING;
app.findTextPreferences.findWhat = "<0016>";
app.selection = app.activeDocument.findText();
app.copy()
app.changeTextPreferences.properties = ({changeTo:"^C^p"});
app.activeDocument.changeText();
*/



/*
myFindChangeByList()

function myFindChangeByList(myObject){
	var myFindChangeFile = File('~/Desktop/a.txt')
    myFindChangeFile = File(myFindChangeFile);
    var myResult = myFindChangeFile.open("r", undefined, undefined);
    do {
        myLine = myFindChangeFile.readln();
        myFindChangeArray = myLine.split('"');
        alert(myFindChangeArray)
    } while(myFindChangeFile.eof == false) {
        myFindChangeFile.close();
    }
}

const logFile = File('~/Desktop/a.txt')
logFile.open("r")

const a

const b = logFile.read()

const c = []

while(b.indexOf("：") !== -1) {
var a = b.splice(1, 2)
//    var end = b.indexOf("：") + 1
    alert(a)
   
}

for(var x = 0; x < 10; x ++) {   
    a = logFile.readln()
//    alert(a)
}
*/

//for(var x = 0; x < activeDocument.textFrames.length; x ++) {   
//    for(var y = 0; y < activeDocument.pages[x].pageItems.length; y ++) {
//        try{
//            if(activeDocument.pages[x].allGraphics[y].imageTypeName !== "TIFF") {
//                document.pages[x].allGraphics[y].associatedXMLElement.untag()   
//            }
//        } catch(e) {}
//    }
//}



/*

const activeDocument = app.activeDocument

activeDocument.metadataPreferences.documentTitle = "1"


function isTiff() {
    for(var x = 0; x < activeDocument.pages.length; x ++) {
        for(var y = 0; y < activeDocument.pages[x].allGraphics.length; y ++) {
            if(activeDocument.pages[x].allGraphics[y].imageTypeName !== "TIFF") {
                document.pages[x].allGraphics[y].associatedXMLElement.untag()   
            }
        }
    }

    for(var x = 0; x < activeDocument.pages.length; x ++) {
        for(var y = 0; y < activeDocument.pages[x].textFrames.length; y ++) {
            document.pages[x].textFrames[y].associatedXMLElement.untag()   
        }
    }
}
*/


/*
for(var x = 0; x < activeDocument.pages.length; x ++) {
    for(var y = 0; y < activeDocument.pages[x].allGraphics.length; y ++) {
        if(activeDocument.allGraphics[x].imageTypeName === "TIFF") {
            alert("Yes")
        }
    }
}
*/


/*
isThreadTextFrame(activeDocument)

function isThreadTextFrame(document) {
    var textFrameLength = 0
    var totalTextFrame = 0

    for(var y = 0; y < document.pages[0].textFrames.length; y ++) {
        alert(document.pages[0].textFrames[y].geometricBounds[3] - document.pages[0].textFrames[y].geometricBounds[1] >= 28)
        if(document.pages[0].textFrames[y].geometricBounds[3] - document.pages[0].textFrames[y].geometricBounds[1] === 28) {
            textFrameLength = document.pages[0].textFrames[y].parentStory.textContainers.length
        break
        }
    }

    for(var a = 0; a < document.pages.length; a ++) {
        for(var y = 0; y < document.pages[a].textFrames.length; y ++) {
            //alert(document.pages[a].textFrames[y].geometricBounds[3] - document.pages[a].textFrames[y].geometricBounds[1] === 28)
            if(document.pages[a].textFrames[y].geometricBounds[3] - document.pages[a].textFrames[y].geometricBounds[1] === 28) {
                totalTextFrame ++
            }
        }
    }

    if(textFrameLength === 0) {
        textFrameLength = 1
    }

    if(document.pages.length > document.pages[0].textFrames[0].parentStory.textContainers.length
    || document.pages.length !== textFrameLength
    || document.pages.length < totalTextFrame) {
//        alert('textFrameLength: ' + textFrameLength)
//        alert('totalTextFrame: ' + totalTextFrame)
//        alert('page: ' + document.pages.length)
        return false
    }
//        alert('textFrameLength: ' + textFrameLength)
//        alert('totalTextFrame: ' + totalTextFrame)
//        alert('page: ' + document.pages.length)
    return true
}
//a.save(File('~/Desktop/' + issue + '-list-' + new Date().getHours() + "_" + new Date().getMinutes() + '.txt'));

var activeDocument = app.activeDocument
const changeObject = app.activeDocument

try {
    app.findChangeGrepOptions.properties = ({includeFootnotes:true, widthSensitive:true});
    app.findGrepPreferences.properties = ({findWhat:"\\d{1, 5}"});
    app.changeGrepPreferences.properties = ({changeTo:"0"});
    changeObject.changeGrep();
} catch (e) {alert(e + ' at line ' + e.line)}
app.findGrepPreferences = NothingEnum.NOTHING;
app.changeGrepPreferences = NothingEnum.NOTHING;
// Query [[~GazInC_2_-]] -- If you delete this comment you break the update function






for(var y = app.activeDocument.pages[0].pageItems.length - 1; y > 0; y --) {
    if(app.activeDocument.pages[0].pageItems[y].geometricBounds[0] > "665.717"
    || app.activeDocument.pages[0].pageItems[y].geometricBounds[1] > "28"
    || app.activeDocument.pages[0].pageItems[y].geometricBounds[2] < "0"
    || app.activeDocument.pages[0].pageItems[y].geometricBounds[3] < "0"
    ){
        app.activeDocument.pages[0].pageItems[y].remove()
    }
}

///////////    Version 8    ///////////

/////////////////////////////////////////////////////////////////////////////////////////////////////    LOG    /////////////////////////////////////////////////////////////////////////////////////////////////////

///////////    Version 8    ///////////    2021-6-28    ///////////
//fixed exorted pdf has empty pages problem (testing)
//removed changeNameAndSaveAndDelete()

/////////////////////////////////////////////////////////////////////////////////////////////////////    LOG    /////////////////////////////////////////////////////////////////////////////////////////////////////

const activeDocument = app.activeDocument
const path = activeDocument.filePath
const issue = activeDocument.filePath.toString()[46] + activeDocument.filePath.toString()[47]
const numberOfPages = activeDocument.pages.length
const logFilePath = activeDocument.fullName.toString().slice(0, 46)
const logFile = File(logFilePath + issue + '/log.txt');

/////////////////////////////////////////////////////////////////////////////////////////////////////    SCRIPT    /////////////////////////////////////////////////////////////////////////////////////////////////////


main()


/////////////////////////////////////////////////////////////////////////////////////////////////////    FUNCTION    /////////////////////////////////////////////////////////////////////////////////////////////////////

function main() {
    var preflightProfile = app.preflightProfiles[1]
    var preflightProcess = app.preflightProcesses.add(activeDocument, preflightProfile)
    
    preflightProcess.waitForProcess()
    preflightResult = preflightProcess.processResults
    
    if(preflightResult.indexOf('None') === 0) {
        
    } else {
        for(var x = 0; x < preflightProcess.aggregatedResults[2].length; x ++) {
            if(preflightProcess.aggregatedResults[2][x].toString().slice(2, 4)) {
                
            }
        }
    }
}
*/
