///////////    Version 8    ///////////

/////////////////////////////////////////////////////////////////////////////////////////////////////    LOG    /////////////////////////////////////////////////////////////////////////////////////////////////////

///////////    Version 1    ///////////    2021-5-31    ///////////
//can get artwork path now

///////////    Version 2    ///////////    2021-6-2    ///////////
//can put rectangle in text frame now

///////////    Version 3    ///////////    2021-8-10    ///////////
//added makeHeader(), setImageTextFramePosition(), getPDFs(), setPrefix()
//redesign dialog.OKbutton and dialog.cancelButton

///////////    Version 4    ///////////    2021-8-11    ///////////
//added deleteEmptyRectangles(), deleteExtraRetruns(), deleteEmptyPages()

///////////    Version 5    ///////////    2021-8-13    ///////////
//added getSectionListAndPushToSections() to get sections contents from server

///////////    Version 6    ///////////    2021-8-16    ///////////
//added getSectionNumberListAndPushToSectionNumbers() to get sectionNumbers contents from server

///////////    Version 7    ///////////    2021-8-26    ///////////
//changed find what is  "公司註冊處處長" to "\\t公司註冊處處長" for apply "Item-Tail"


///////////    Version 8    ///////////    2021-9-9    ///////////
//added app.pdfPlacePreferences.pdfCrop = PDFCrop.cropMedia to insertPDFs() to ensure import pdfs with media box

/////////////////////////////////////////////////////////////////////////////////////////////////////    SCRIPT    /////////////////////////////////////////////////////////////////////////////////////////////////////

const activeDocument = app.activeDocument
const issue = activeDocument.name.toString()[5] + activeDocument.name.toString()[6]
const version = activeDocument.name.toString()[4]
const gnNumber = activeDocument.name.toString().slice(0, 3)
const path = activeDocument.fullName.toString().slice(0, 49)
var headerTextFrame
var imageTextFrame
const header = "第xxx號公告\r公司註冊處\r公司條例\(第622章\)\r"
const sectionNumbers = [];
const sections = []

getSectionNumberListAndPushToSectionNumbers()
getSectionListAndPushToSections()
makeDialog();
dialog.center(); 
if(dialog.show() === 1){
  type = dialog.kindType.selection.index;
}else{
  exit();
}
makeHeader()
setPrefix()
insertPDFs()
deleteEmptyRectangles()
deleteExtraRetruns()
deleteEmptyPages()

/////////////////////////////////////////////////////////////////////////////////////////////////////    FUNCTION    /////////////////////////////////////////////////////////////////////////////////////////////////////

function makeDialog() {
    dialog = new Window('dialog', "", "x:0, y:0, width:300, height:80");
    dialog.panel = dialog.add('panel', [15, 10, 195, 70], "");
    dialog.panel.add('statictext', [10, 20, 75, 45], "條例編號：");
    dialog.kindType = dialog.panel.add('dropdownlist', [80, 15, 165, 40]);
    dialog.kindType.onChange = function showSectionNumbers() {
        kind(dialog.kindType, placementINFO.pgCount, "條例編號");
    }
    for(var x = 0; x < sectionNumbers.length; x ++){
        dialog.kindType.add('item', sectionNumbers[x]);
    }
    dialog.OKbutton = dialog.add('button', [210, 10, 285, 35], "OK");
    dialog.OKbutton.onClick = function onOKclicked() {
        dialog.close(1);
    }
    dialog.cancelButton = dialog.add('button', [210, 45, 285, 70], "Cancel");
    dialog.cancelButton.onClick = function onCANclicked() {
        dialog.close(0);
    }
  return dialog;
}


function getSectionNumberListAndPushToSectionNumbers() {
    var path = File(app.activeScript).path
    var  sectionListNumber = File("//EXZIP18/Gazette_N/_Operator's/InDesign Utilities/zzz__setting-indesign/CS6_Printer & Adobe/Scripts/Scripts Panel/FindChangeSupport/Company Section Numbers.txt")
    sectionListNumber = File(sectionListNumber);
    sectionListNumber.open("r");
    do {
        line = sectionListNumber.readln();
        sectionNumber = line.split(',');
        sectionNumbers.push(sectionNumber)
    } while(sectionListNumber.eof == false) {
        sectionListNumber.close();
    }
}


function getSectionListAndPushToSections() {
    var path = File(app.activeScript).path
    var  sectionList = File("//EXZIP18/Gazette_N/_Operator's/InDesign Utilities/zzz__setting-indesign/CS6_Printer & Adobe/Scripts/Scripts Panel/FindChangeSupport/Company - C.txt")
    sectionList = File(sectionList);
    sectionList.open("r");
    do {
        line = sectionList.readln();
        section = line.split(',');
        sections.push(section)
    } while(sectionList.eof == false) {
        sectionList.close();
    }
}


function makeHeader() {
    headerTextFrame = activeDocument.pages[0].textFrames[0]
    headerTextFrame.geometricBounds = ["0", "28", "503", "0",];    //["上", "左", "下", "右",]
    app.selection = headerTextFrame.rectangles[0]
    app.cut()
    headerTextFrame.contents = header + sections[type] + "\r"

    try {
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;
        app.findChangeGrepOptions.properties = ({includeFootnotes: false, includeMasterPages: false, includeHiddenLayers: false, wholeWord: false});
        app.findGrepPreferences.properties = ({findWhat:"xxx"});
        app.changeGrepPreferences.properties = ({changeTo: gnNumber});
        app.changeGrepPreferences.appliedParagraphStyle =  "ab  BM-Item--Head_no 5pt"
        activeDocument.changeGrep();
    } catch (e) {alert(e)}

    try {
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;
        app.findGrepPreferences.properties = ({findWhat:"公司註冊處"});
        app.changeGrepPreferences.properties = ({changeTo:""});
        app.changeGrepPreferences.appliedParagraphStyle =  "H-Roman-Center_0.001pt"
        activeDocument.changeGrep();
    } catch (e) {alert(e)}

    try {
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;
        app.findGrepPreferences.properties = ({findWhat:"公司條例\\(第622章\\)"});
        app.changeGrepPreferences.properties = ({changeTo:""});
        app.changeGrepPreferences.appliedParagraphStyle =  "H-Bold-Center"
        activeDocument.changeGrep();
    } catch (e) {alert(e)}

    try {
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;
        app.findGrepPreferences.properties = ({findWhat:"：\\r"});
        app.changeGrepPreferences.properties = ({changeTo:"：\\r\\r"});
        app.changeGrepPreferences.appliedParagraphStyle =  "Text"
        activeDocument.changeGrep();
    } catch (e) {alert(e)}

    try {
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;
        app.findGrepPreferences.findWhat = "(?<=：\\r)\\r";
        app.changeGrepPreferences.properties = ({changeTo:"\\r\\r"});
        app.changeGrepPreferences.appliedParagraphStyle =  "Picture"
        activeDocument.changeGrep();
    } catch (e) {alert(e)}

    try {
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;
        app.findGrepPreferences.findWhat = "(?<=\\r)\\r(?=\\r)";
        app.selection = app.activeDocument.findGrep();
        app.paste()
    } catch (e) {alert(e)}

    try {
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;
        app.findGrepPreferences.findWhat = "\\t公司註冊處處長";
        app.changeGrepPreferences.appliedParagraphStyle =  "Item-Tail"
        activeDocument.changeGrep();
    } catch (e) {alert(e)}
}


function setPrefix() {
    activeDocument.sections[0].sectionPrefix = gnNumber + "-"
}


function getPDFs(artworkFolder) {
	try { 
        return artworkFolder.getFiles("*"+ ".pdf"); 
	} catch (e) {alert(e)}
}


function insertPDFs() {
    var PDFs = getPDFs(Folder(path + '_Artwork/' + gnNumber))
    app.pdfPlacePreferences.pdfCrop = PDFCrop.cropMedia
    for(var x = 0; x < activeDocument.pages.length; x ++) {
        try {
            activeDocument.pages[x].textFrames[0].rectangles[0].place(File(PDFs[x].fullName)); 
        } catch(e) {
            x = activeDocument.pages.length
        }
    }
}


function deleteEmptyRectangles() {
    for(var x = activeDocument.pages.length - 1; x > 0; x --) {
        if(activeDocument.pages[x].allGraphics[0] === undefined) {
            try {
                activeDocument.pages[x].textFrames[0].rectangles[0].remove()
            } catch(e) {}
        } else {
            activeDocument.pages[x].textFrames[0].geometricBounds = ["0", "28", "528", "0",]
            x = 0
        }
    }
}


function deleteExtraRetruns() {
    try {
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;
        app.findGrepPreferences.findWhat = "\\r+(\\d)";
        app.changeGrepPreferences.properties = ({changeTo:"\\r$1"});
        activeDocument.changeGrep();
    } catch (e) {alert(e)}
}


function deleteEmptyPages() {
    for(var x = activeDocument.pages.length - 1; x > 0; x --) {
        if(activeDocument.pages[x].textFrames[0].contents.length === 0){
            activeDocument.pages[x].remove()
        }
    }
}






/*
const sectionNumbers = ["291(3)", "291(6)", "291AA(7)", "291AA(9)", "744(3)", "745(2)(b)", "746(2)", "747(2)", "751(1)", "751(3)", "796(3)", "797(2)(b)", "798(2)"];

const sections = [
  "在《公司條例》(第622章) (下稱「該條例」)附表9的生效日期前不時有效的《前身公司條例》，其第291(3)條已被廢除，但在被廢除後，根據該條例附表11第129條，具有持續效力。現依據《前身公司條例》第291(3)條公布，在本公告刊登當日起計三個月期滿時，除非有相反因由提出，否則下列公司的名稱將從登記冊上剔除，而公司亦予以解散：", 
  "在《公司條例》(第622章) (下稱「該條例」)附表9的生效日期前不時有效的《前身公司條例》，其第291(6)條已被廢除，但在被廢除後，根據該條例附表11第129條，具有持續效力。現依據《前身公司條例》第291(6)條公布，下列公司的名稱已從公司登記冊上剔除，並由本公告刊登當日起予以解散：", 
  "在《公司條例》(第622章) (下稱「該條例」)附表9的生效日期前不時有效的《前身公司條例》，其第291AA(7)條已被廢除，但在被廢除後，根據該條例附表11第129條，具有持續效力。現依據《前身公司條例》第291AA(7)條公布，除非處長在本公告刊登日期後三個月內收到對下列公司的撤銷註冊而提出的反對，否則處長可將下列公司的註冊撤銷和解散下列公司：", 
  "在《公司條例》(第622章) (下稱「該條例」)附表9的生效日期前不時有效的《前身公司條例》，其第291AA(9)條已被廢除，但在被廢除後，根據該條例附表11第129條，具有持續效力。現依據《前身公司條例》第291AA(9)條公布，下列公司的註冊在本公告刊登當日撤銷，而下列公司亦在註冊撤銷時解散：", 
  "現根據《公司條例》第744(3)條公布，除非有反對因由提出，否則在本公告的日期後的3個月終結時，下列公司的名稱將會從公司登記冊剔除，而下列公司將會解散：", 
  "現根據《公司條例》第745(2)(b)條公布，除非有反對因由提出，否則在本公告的日期後的3個月終結時，下列公司的名稱將會從公司登記冊剔除，而下列公司將會解散：", 
  "現根據《公司條例》第746(2)條公布，下列公司的名稱已從公司登記冊剔除，而下列公司由本公告刊登當日起即告解散：", 
  "現根據《公司條例》第747(2)條公布，除非有反對因由提出，否則在本公告的日期後的3個月終結時，下列公司的名稱將會從公司登記冊剔除，而下列公司將會解散：", 
  "現依據《公司條例》第751(1)條公布，除非處長在本公告刊登的日期後3個月內收到對下列公司的撤銷註冊的反對，否則處長可撤銷下列公司的註冊：", 
  "公司註冊處處長現依據《公司條例》第751(3)條宣布，下列公司的註冊在本公告刊登當日撤銷，而依據《公司條例》第751(6)條，下列公司亦在註冊撤銷時即告解散：", 
  "現根據《公司條例》第796(3)條公布，除非有反對因由提出，否則在本公告的日期後的3個月終結時，下列公司的名稱將會從公司登記冊剔除，而下列公司將不再是註冊非香港公司：", 
  "現根據《公司條例》第797(2)(b)條公布，除非有反對因由提出，否則在本公告的日期後的3個月終結時，下列公司的名稱將會從公司登記冊剔除，而下列公司將不再是註冊非香港公司：", 
  "現根據《公司條例》第798(2)條公布，下列公司的名稱已從公司登記冊剔除，而下列公司即不再是註冊非香港公司：", 
]
*/