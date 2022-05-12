﻿///////////    Version 1    ///////////

/////////////////////////////////////////////////////////////////////////////////////////////////////    LOG    /////////////////////////////////////////////////////////////////////////////////////////////////////

///////////    Version 1    ///////////    2021-8-11    ///////////
//added save(), getHeader(), applyEnglishMasterPage(), changeEnglishItemTail(), saveAs()

/////////////////////////////////////////////////////////////////////////////////////////////////////    SCRIPT    /////////////////////////////////////////////////////////////////////////////////////////////////////

const activeDocument = app.activeDocument
const issue = activeDocument.name.toString()[5] + activeDocument.name.toString()[6]
const version = activeDocument.name.toString()[4]
const gnNumber = activeDocument.name.toString().slice(0, 3)
const path = activeDocument.fullName.toString().slice(0, 37)
var headerTextFrame
var imageTextFrame
const header = "G.N. xxx\rCompanies Registry\rCompanies Ordinance \(Chapter 622\)\r"
const sectionNumbers = ["291\\(3\\)", "291\\(6\\)", "291AA\\(7\\)", "291AA\\(9\\)", "744\\(3\\)", "745\\(2\\)\\(b\\)", "746\\(2\\)", "747\\(2\\)", "751\\(1\\)", "751\\(3\\)", "796\\(3\\)", "797\\(2\\)\\(b\\)", "798\\(2\\)"];
const sections = [
  "Pursuant to section 291(3) of the predecessor Companies Ordinance (as in force from time to time before the commencement date of Schedule 9 to the Companies Ordinance (Chapter 622) (the ‘Ordinance’)) which is repealed but has a continuing effect under section 129 of Schedule 11 to the Ordinance, notice is hereby given that at the expiration of three months from the date hereof, the names of the undermentioned companies will, unless cause is shown to the contrary, be struck off the Companies Register and the companies will be dissolved:—", 
  "Pursuant to section 291(6) of the predecessor Companies Ordinance (as in force from time to time before the commencement date of Schedule 9 to the Companies Ordinance (Chapter 622) (the ‘Ordinance’)) which is repealed but has a continuing effect under section 129 of Schedule 11 to the Ordinance, notice is hereby given that the names of the undermentioned companies have been struck off the Companies Register. Such companies are accordingly dissolved as from the date of the publication of this notice:—", 
  "Pursuant to section 291AA(7) of the predecessor Companies Ordinance (as in force from time to time before the commencement date of Schedule 9 to the Companies Ordinance (Chapter 622) (the ‘Ordinance’)) which is repealed but has a continuing effect under section 129 of Schedule 11 to the Ordinance, notice is hereby given that unless an objection to the deregistration of the undermentioned companies is received within 3 months after the date of the publication of this Gazette Notice, the Registrar may deregister the undermentioned companies and dissolve them:—", 
  "Pursuant to section 291AA(9) of the predecessor Companies Ordinance (as in force from time to time before the commencement date of Schedule 9 to the Companies Ordinance (Chapter 622) (the ‘Ordinance’)) which is repealed but has a continuing effect under section 129 of Schedule 11 to the Ordinance, the undermentioned companies are deregistered upon the date of the publication of this Gazette Notice. The undermentioned companies are accordingly dissolved on deregistration:—", 
  "Pursuant to section 744(3) of the Companies Ordinance, notice is hereby given that unless cause is shown to the contrary, the names of the undermentioned companies will be struck off the Companies Register, and the companies dissolved, at the end of 3 months after the date hereof:—", 
  "Pursuant to section 745(2)(b) of the Companies Ordinance, notice is hereby given that unless cause is shown to the contrary, the names of the undermentioned companies will be struck off the Companies Register, and the companies dissolved, at the end of 3 months after the date hereof:—", 
  "Pursuant to section 746(2) of the Companies Ordinance, notice is hereby given that the names of the undermentioned companies have been struck off the Companies Register, and the companies dissolved as from the date of the publication of this notice:—", 
  "Pursuant to section 747(2) of the Companies Ordinance, notice is hereby given that unless cause is shown to the contrary, the names of the undermentioned companies will be struck off the Companies Register, and the companies dissolved, at the end of 3 months after the date hereof:—", 
  "Pursuant to section 751(1) of the Companies Ordinance, notice is hereby given that unless an objection to the deregistration of the undermentioned companies is received by the Registrar within 3 months after the date of publication of this Gazette Notice, the Registrar may deregister the undermentioned companies:—", 
  "Pursuant to section 751(3) of the Companies Ordinance, the Registrar hereby declares that the undermentioned companies are deregistered on the date of publication of this Gazette Notice, and pursuant to section 751(6) of the Companies Ordinance, the undermentioned companies are dissolved on deregistration:—", 
  "Pursuant to section 796(3) of the Companies Ordinance, notice is hereby given that unless cause is shown to the contrary, the names of the undermentioned companies will be struck off the Companies Register, and the companies will no longer be registered non-Hong Kong companies, at the end of 3 months after the date hereof:—", 
  "Pursuant to section 797(2)(b) of the Companies Ordinance, notice is hereby given that unless cause is shown to the contrary, the names of the undermentioned companies will be struck off the Companies Register, and the companies will no longer be registered non-Hong Kong companies, at the end of 3 months after the date hereof:—", 
  "Pursuant to section 798(2) of the Companies Ordinance, notice is hereby given that the names of the undermentioned companies have been struck off the Companies Register, and the companies are no longer registered non-Hong Kong companies:—", 
]

save()
getHeader()
applyEnglishMasterPage()
changeEnglishItemTail()
saveAs()


/////////////////////////////////////////////////////////////////////////////////////////////////////    FUNCTION    /////////////////////////////////////////////////////////////////////////////////////////////////////

function save() {
    activeDocument.save()
}


function getHeader() {
    for(var x = 0; x < sectionNumbers.length; x ++) {
        app.findGrepPreferences.findWhat = sectionNumbers[x]
        result = app.activeDocument.findGrep();
        if(result.length > 0) {
            makeHeader(sections[x])
            x = sectionNumbers.length
        }
    }
}


function makeHeader(sections) {
    headerTextFrame = activeDocument.pages[0].textFrames[0]
    headerTextFrame.geometricBounds = ["0", "28", "503", "0",];    //["上", "左", "下", "右",]
    app.selection = headerTextFrame.rectangles[0]
    app.cut()
    headerTextFrame.contents = header + sections + "\r"

    try {
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;
        app.findChangeGrepOptions.properties = ({includeFootnotes: false, includeMasterPages: false, includeHiddenLayers: false, wholeWord: false});
        app.findGrepPreferences.properties = ({findWhat:"xxx"});
        app.changeGrepPreferences.properties = ({changeTo: gnNumber});
        app.changeGrepPreferences.appliedParagraphStyle =  "ab  BM-Item--Head-E no 4pt"
        activeDocument.changeGrep();
    } catch (e) {alert(e)}

    try {
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;
        app.findGrepPreferences.properties = ({findWhat:"Companies Registry"});
        app.changeGrepPreferences.appliedParagraphStyle =  "Department_Right"
        activeDocument.changeGrep();
    } catch (e) {alert(e)}

    try {
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;
        app.findGrepPreferences.properties = ({findWhat:"Companies Ordinance \\(Chapter 622\\)"});
        app.changeGrepPreferences.appliedParagraphStyle =  "H--Caps-E"
        activeDocument.changeGrep();
    } catch (e) {alert(e)}

    try {
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;
        app.findGrepPreferences.properties = ({findWhat:":—"});
        app.changeGrepPreferences.properties = ({changeTo:":—\\r\\r"});
        app.changeGrepPreferences.appliedParagraphStyle =  "Text-E"
        activeDocument.changeGrep();
    } catch (e) {alert(e)}

    try {
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;
        app.findGrepPreferences.findWhat = "(?<=:—\\r)\\r\\r";
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
}


function applyEnglishMasterPage() {
    for(var x = 0; x < activeDocument.pages.length; x ++) {
        activeDocument.pages[x].appliedMaster = activeDocument.masterSpreads.item('E-主版')
    }
}


function changeEnglishItemTail() {
    var itemTail = activeDocument.masterSpreads.item('E-主版').textFrames[2].contents
    try {
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;
        app.findGrepPreferences.findWhat = "(?<=\\r)\\d+年.+公司註冊處處長\\n\\t.+代行\\)";
        app.changeGrepPreferences.properties = ({changeTo: itemTail});
        app.changeGrepPreferences.appliedParagraphStyle = "Item-Tail-E"
        activeDocument.changeGrep();
    } catch (e) {alert(e)}
    
    try {
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;
        app.findGrepPreferences.findWhat = "(?<=\\r).+\\t";
        app.changeGrepPreferences.appliedCharacterStyle = "Times_Italic"
        activeDocument.changeGrep();
    } catch (e) {alert(e)}

    try {
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;
        app.findGrepPreferences.findWhat = "(?<= for )Registrar of Companies";
        app.changeGrepPreferences.appliedCharacterStyle = "Times_Italic"
        activeDocument.changeGrep();
    } catch (e) {alert(e)}
}


function saveAs() {
    activeDocument.save(File(path + 'English/E' + issue + '\\' + gnNumber + '-E' + issue + '.indd'));
    activeDocument.close();
}