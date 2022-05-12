// 基本語法
// 重複



// 新增檔案
app.documents.add();
// 新增頁面
document.pages.add();


////////////////////////////////////////////////////////////////    變數或常數    ////////////////////////////////////////////////////////////////
// var 變數
// const 常數
// 變數和常數分別: 變數可以更改值, 常數不可以更改值


////////////////////////////////////////////////////////////////    使用變數或常數控制各種物件    ////////////////////////////////////////////////////////////////
// 宣告 file1 為新增的檔案
const file1 = app.documents.add();
// 宣告 file1 為舊檔案
const file1 = app.open (File('~/Desktop/asdf.indd'));
// 宣告 file1 為現在的檔案
const file1 = app.activeDocument
// file1增加頁面
file1.pages.add();
// 另存 file1
file1.save(File('~/Desktop/asdffdsa.indd'));
// 關閉 file1
file1.close();
// 移動頁面
// 第一個數字代表要移動的頁面, 
// 第二個數字表目的地頁面,
//新增頁面
file1.pages.add()
// LocationOptions.AFTER 代表移動到目的地頁面之前或是之後
file1.pages[0].move(LocationOptions.AFTER, document.pages[3])


////////////////////////////////////////////////////////////////    色票    ////////////////////////////////////////////////////////////////
// 新增色票
// name:Bluee是在色票上顯示的名稱
document.colors.add({name:"Bluee", space:ColorSpace.CMYK, model:ColorModel.process, colorValue:[100, 0, 0, 0]});


////////////////////////////////////////////////////////////////    Master Page    ////////////////////////////////////////////////////////////////
var masterP1 = document.masterSpreads.add(1); //"document.masterSpreads.add(1)" 數字代表頁數
var masterP2 = document.masterSpreads.add(2);
var masterP3 = document.masterSpreads.add(3);
document.pages[0].appliedMaster = masterP2;


////////////////////////////////////////////////////////////////    文字    ////////////////////////////////////////////////////////////////
// 宣告常數 text1 為文字框
// 第一個數字代表要新增文字框的頁面
const text1 = document.pages.item(0).textFrames.add();
// text1 文字框內容
text1.contents = "P1";
// text1 文字大小
text1.parentStory.pointSize = 24;
// text1 字款
text1.parentStory.appliedFont = "Times New Roman";
// text1 套用字元樣式
text1.parentStory.appliedCharacterStyle = "字元樣式名稱";
// text1 套用段落樣式
text1.parentStory.appliedParagraphStyle = "段落樣式名稱";
// text1 內水平容齊行    
text1.texts.item(0).justification = Justification.centerAlign;  //leftAlign    centerAlign    rightAlign    ???Align
// text1 內容垂直齊行    
text1.textFramePreferences.verticalJustification = VerticalJustification.CENTER_ALIGN;  //TOP_ALIGN    CENTER_ALIGN    BOTTOM_ALIGN    justifyAlign
// 使 text1 文字框符合內容
text1.fit(FitOptions.frameToContent);
// text1 文字框座標
text1.geometricBounds = ["10", "10", "50", "50",];    //["上", "左", "下", "右",]
// text1 文字框線條粗幼
text1.strokeWeight = 2;
// text1 文字框線條顏色
text1.strokeColor = "色標名稱";
// text1 文字框填色
text1.fillColor = "色標名稱";
// text1 文字框旋轉角度
text1.transform(CoordinateSpaces.pasteboardCoordinates, AnchorPoint.centerAnchor, app.transformationMatrices.add({counterclockwiseRotationAngle: 180}));	
// text1 文字套用色標顏色
text1.fillColor = "色標名稱";
// text1 文字線條顏色
text1.strokeColor = "色標名稱";
// text1 文字線條粗幼
text1.strokeWeight = 1;


////////////////////////////////////////////////////////////////    矩形    ////////////////////////////////////////////////////////////////
// 宣告常數 retangle1 為新增的矩形
const retangle1 = document.rectangles.add();
// 新增矩形(連參數)
const retangle1 = document.rectangles.add({geometricBounds:[72, 72, 96, 96], strokeWeight:2, strokeColor:blue, fillColor:blue});
// 使 retangle1 線條粗幼
retangle1.strokeWeight = 2;
// 使 retangle1 線條顏色
retangle1.strokeColor = "色標名稱";
// 使 retangle1 框填色
retangle1.fillColor = "色標名稱";
//矩形旋轉角度
//將counterclockwiseRotationAngle:設定為需要的角度
//CoordinateSpaces、AnchorPoint，都可以選擇設定，但我無去研究)
retangle1.transform(CoordinateSpaces.pasteboardCoordinates, AnchorPoint.centerAnchor, app.transformationMatrices.add({counterclockwiseRotationAngle: 45}));	

//or 
 
//在rotate()的()中放入要旋轉的目標，例如: app.activeDocument.pages[0].textFrames[0]
rotate()
function rotate(target) {
    // 將counterclockwiseRotationAngle:設定為需要的角度
    var rotateDegree = app.transformationMatrices.add({counterclockwiseRotationAngle: 90});
    target.transform (CoordinateSpaces.pasteboardCoordinates, AnchorPoint.centerAnchor, rotateDegree);
}

////////////////////////////////////////////////////////////////    表格    ////////////////////////////////////////////////////////////////
// 宣告常數 table1 為新增的表格
// columnCount: = 欄數
// bodyRowCount: = 行數
const table1 = file1.textFrames.add().tables.add({columnCount:3, bodyRowCount:3});
// 使 table1 新增一行
table1.rows.add();
// 使 table1 新增一欄
table1.columns.add();

var tableFrame = document.masterSpreads.item(6).textFrames.add();
tableFrame.geometricBounds = ["20", "10", "277", "200"];    //["上", "左", "下", "右",]
var numberPageTable = tableFrame.tables.add({columnCount:11, bodyRowCount:32});
numberPageTable.rows.everyItem().height = 6;
numberPageTable.columns.everyItem().width = 6;

numberPageTable.rows[5].height = 30;
numberPageTable.columns[5].width = 40;
numberPageHeader.rows[0].bottomEdgeStrokeColor = "色標名稱";
numberPageHeader.rows[0].bottomEdgeStrokeWeight = 10;


////////////////////////////////////////////////////////////////    文件設定    ////////////////////////////////////////////////////////////////
//設定 x 軸單位為 points 
document.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.points;
//設定 x 軸單位為 millimeters 
document.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.millimeters;

//設定 y 軸單位為 points 
document.viewPreferences.verticalMeasurementUnits = MeasurementUnits.points;
//設定 y 軸單位為 millimeters 
document.viewPreferences.verticalMeasurementUnits = MeasurementUnits.millimeters;

//設定坐標
activeDocument.zeroPoint = [0, 0]


////////////////////////////////////////////////////////////////    物件    ////////////////////////////////////////////////////////////////
//ungroup物件 
activeDocument.pages[x].groups.everyItem().ungroup()

//全選頁面物件並刪除
var item = activeDocument.pages[0].allPageItems
for(var i = item.length - 1; i > 0; i --) {
    item[i].remove()
}

//全選頁面文字框 
var text1 = activeDocument.pages[0].textFrames.everyItem().select()
//檢查溢排文字 
if(text1.textFrame.overflows) {}

//使用變數儲存selection物件 
a = app.selection[0]
//使用變數儲存selection物件 
b = app.selection[1]
//使用變數儲存selection物件 
c = app.selection[2]
//a成為c的串聯文字框
c.nextTextFrame = a

//全選頁面 Rectangles
activeDocument.pages[0].rectangles.everyItem().select()
//全選頁面 Groups
activeDocument.pages[0].groups.everyItem().select()








/*
document.viewPreferences.rulerOrigin = 10;

master32.rectangles.add({geometricBounds:[45, 15, 71.516, 52.5], strokeWeight:1});

myDocument.pages[2].appliedMaster = document.masterSpreads.item("B-32");

myDocument.masterSpreads.add();
var b = myDocument.masterSpreads.item(1);
var myRightPage = b.pages.item(0);
var myRightFooter = myRightPage.rectangles.add();
myRightFooter.geometricBounds = [728, 70, 742, 528];

rectangle = document.rectangles.add({geometricBounds:[39.508, 20.492, 77.008, 47.008], strokeWeight:1, strokeColor:black});  //sample
dialog.pan4 = dialog.add('panel', [210,200,438,350], "類型：");
mapPages = Number(dialog.mapPages.value);
reverseOrder = Number(dialog.reverseOrder.value);
docStartPG = Number(dialog.docStartPG.text);

var master0 = document.masterSpreads.item(1);
var master32 = document.masterSpreads.add(1);
var master16 = document.masterSpreads.add(1);
var master12 = document.masterSpreads.add(1);
var master8 = document.masterSpreads.add(1);
var master4 = document.masterSpreads.add(1);


var myDocument = app.documents.add();
myDocument.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.points;    //millimeters
myDocument.viewPreferences.verticalMeasurementUnits = MeasurementUnits.points;
myDocument.viewPreferences.rulerOrigin = RulerOrigin.pageOrigin;


var master32 = myDocument.masterSpreads.item(0);
master32.rectangles.add({geometricBounds:[45, 15, 71.516, 52.5], strokeWeight:1});


myDocument.masterSpreads.add();
var b = myDocument.masterSpreads.item(1);
var myRightPage = b.pages.item(0);
var myRightFooter = myRightPage.rectangles.add();
myRightFooter.geometricBounds = [728, 70, 742, 528];









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
















isValid
signatureFields
textBoxes
radioButtons
listBoxes
comboBoxes
checkBoxes
multiStateObjects
buttons
formFields
epstexts
groups
guides
eventListeners
events
polygons
textFrames
graphicLines
rectangles
pageItems
splineItems
ovals
preferences
properties
parent
label
id
gridData
tabOrder
allGraphics
allPageItems
masterPageTransform
appliedMaster
masterPageItems
pageColor
appliedTrapPreset
bounds
documentOffset
index
appliedSection
name
side
marginPreferences
optionalPage
snapshotBlendingMode
layoutRule
appliedAlternateLayout





*/