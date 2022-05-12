///////////    Version 7   ///////////

/////////////////////////////////////////////////////////////////////////////////////////////////////    LOG    /////////////////////////////////////////////////////////////////////////////////////////////////////

///////////    Version 1    ///////////    2021-8-11    ///////////
//added save(), getHeader(), applyEnglishMasterPage(), changeEnglishItemTail(), saveAs()

///////////    Version 2    ///////////    2021-8-12    ///////////
//redesign getHeader() and reordered the excute stack

///////////    Version 3    ///////////
//added fitFrameToContent()

///////////    Version 4    ///////////    2021-8-13    ///////////
//added getSectionListAndPushToSections() to get section contents from server

///////////    Version 5    ///////////    2021-8-13    ///////////
//change [[:graph:]] to ~a

///////////    Version 6    ///////////    2021-8-16    ///////////
//added getSectionListAndPushToSections() to get sections contents from server
//added getSectionNumberListAndPushToSectionNumbers() to get sectionNumbers contents from server

///////////    Version 7    ///////////    2021-8-16    ///////////
//added new logic to getSectionNumberListAndPushToSectionNumbers to regextify string before push to sectionNumbers[]

/////////////////////////////////////////////////////////////////////////////////////////////////////    SCRIPT    /////////////////////////////////////////////////////////////////////////////////////////////////////

const activeDocument = app.activeDocument
const issue = activeDocument.name.toString()[5] + activeDocument.name.toString()[6]
const version = activeDocument.name.toString()[4]
const gnNumber = activeDocument.name.toString().slice(0, 3)
const path = activeDocument.fullName.toString().slice(0, 37)
var headerTextFrame
var imageTextFrame
const header = "G.N. xxx\rCompanies Registry\rCompanies Ordinance \(Chapter 622\)\r"
const sectionNumbers = [];
const sections = []


save()
saveAs()
getSectionNumberListAndPushToSectionNumbers()
getSectionListAndPushToSections()
getHeader()
applyEnglishMasterPage()
changeEnglishItemTail()
fitFrameToContent()
save()


/////////////////////////////////////////////////////////////////////////////////////////////////////    FUNCTION    /////////////////////////////////////////////////////////////////////////////////////////////////////

function save() {
    activeDocument.save()
}


function saveAs() {
    activeDocument.save(File(path + 'English/E' + issue + '\\' + gnNumber + '-E' + issue + '.indd'));
}


function getSectionNumberListAndPushToSectionNumbers() {
    var path = File(app.activeScript).path
    var  sectionListNumber = File("//EXZIP18/Gazette_N/_Operator's/InDesign Utilities/zzz__setting-indesign/CS6_Printer & Adobe/Scripts/Scripts Panel/FindChangeSupport/Company Section Numbers.txt")
    sectionListNumber = File(sectionListNumber);
    sectionListNumber.open("r");
    do {
        line = sectionListNumber.readln();
        sectionNumber = line.split(',');
        
        var regexSectionNumber = sectionNumber.toString().slice(0, sectionNumber.toString().indexOf("("))
        regexSectionNumber += "\\("
        if(sectionNumber.toString().indexOf(")(") !== -1) {
            regexSectionNumber += sectionNumber.toString().slice(sectionNumber.toString().indexOf("(") + 1, sectionNumber.toString().indexOf(")("))
            regexSectionNumber += "\\)\\("
        }
        regexSectionNumber += sectionNumber.toString().slice(sectionNumber.toString().lastIndexOf("(") + 1, sectionNumber.toString().lastIndexOf(")"))
        regexSectionNumber += "\\)"        
       
        sectionNumbers.push(regexSectionNumber)
    } while(sectionListNumber.eof == false) {
        sectionListNumber.close();
    }
}


function getSectionListAndPushToSections() {
    var path = File(app.activeScript).path
    var  sectionList = File("//EXZIP18/Gazette_N/_Operator's/InDesign Utilities/zzz__setting-indesign/CS6_Printer & Adobe/Scripts/Scripts Panel/FindChangeSupport/Company - E.txt")
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
    try {
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;
        app.findChangeGrepOptions.properties = ({includeFootnotes: false, includeMasterPages: false, includeHiddenLayers: false, wholeWord: false});
        app.findGrepPreferences.properties = ({findWhat:"：\\r"});
        app.changeGrepPreferences.properties = ({changeTo: "：~P"});
        activeDocument.changeGrep();
    } catch (e) {alert(e)}

    activeDocument.pages[0].textFrames[0].contents = header + sections + "\r"

    try {
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;
        app.findGrepPreferences.findWhat = "~a";
        app.changeGrepPreferences.appliedParagraphStyle =  "Picture"
        activeDocument.changeGrep();
    } catch (e) {alert(e)}

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
        app.changeGrepPreferences.appliedParagraphStyle =  "Text-E"
        activeDocument.changeGrep();
    } catch (e) {alert(e)}
}


function applyEnglishMasterPage() {
    for(var x = 0; x < activeDocument.pages.length; x ++) {
        activeDocument.pages[x].appliedMaster = activeDocument.masterSpreads.item('E-主版')
    }
}


function changeEnglishItemTail() {
    var itemTail = activeDocument.masterSpreads.item('E-主版').textFrames[0].contents

    try {
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;
        app.findGrepPreferences.findWhat = "(?<=\\r)\\d+年.+公司註冊處處長\\n\\t.+代行\\)";
        app.changeGrepPreferences.appliedParagraphStyle = "Item-Tail-E"
        activeDocument.changeGrep();
    } catch (e) {alert(e)}

    try {
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;
        app.findGrepPreferences.findWhat = "(?<=\\r)\\d+年.+公司註冊處處長\\n\\t.+代行\\)";
        app.changeGrepPreferences.properties = ({changeTo: itemTail});
        activeDocument.changeGrep();
    } catch (e) {alert(e)}

    try {
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;
        app.findGrepPreferences.findWhat = "(?<=\\t).+(?=Registrar of Companies)";
        app.changeGrepPreferences.appliedCharacterStyle = "Times_Regular"
        activeDocument.changeGrep();
    } catch (e) {alert(e)}

}


function fitFrameToContent() {
    activeDocument.pages[activeDocument.pages.length - 1].textFrames[0].fit(FitOptions.frameToContent);
}








/*
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
*/