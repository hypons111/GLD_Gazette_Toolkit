///////////    Version 16    ///////////

/////////////////////////////////////////////////////////////////////////////////////////////////////    LOG    /////////////////////////////////////////////////////////////////////////////////////////////////////

///////////    Version 11    ///////////    2021-5-21    ///////////
//Fixed 造字 paragraph style
//changed function isMultiTextFrame() from "switch" to "if " to fixed unknown problam at function deleteOversetedTextFrame(). REF file on 17-???-???

///////////    Version 12    ///////////    2021-5-27    ///////////
//change finish alert to "已完成\請檢查後執行 \"大報 - 上網 - 2\"" 

///////////    Version 13    ///////////    2021-5-29    ///////////
//added one more condition on the isMultiTextFrame() & groupDoubleTextFrame() for the useless structure tag 

///////////    Version 14    ///////////    2021-5-31    ///////////
//removed groupCanceledGN() 
//set preflight function on and put finish alert in it
//added deleteEmptyTextFrame()
//deleted groupDoubleTextFrame()

///////////    Version 15    ///////////    2021-6-3    ///////////
//added sendItemToBack() for the GN text frame coverd the contents text frame in multi text frames cases
//initialized zero point to contents area

///////////    Version 16    ///////////    2021-6-4    ///////////
//added spreadOversetedGN()
//removed spreadChapter380()
//added checkOversetedGNNumberTextFrame() & splitOversetedGNNumberTextFrame()

/////////////////////////////////////////////////////////////////////////////////////////////////////    LOG    /////////////////////////////////////////////////////////////////////////////////////////////////////

const activeDocument = app.activeDocument
var oversetedGNNumberTextFrame = false

setDocumentPreference()
deleteMasterPageItems()

checkOversetedGNNumberTextFrame()

if(oversetedGNNumberTextFrame === false) {
    for(var x = 0; x < activeDocument.pages.length; x ++) {
        ungroup(activeDocument.pages[x])
        deleteHeader(activeDocument.pages[x])
        for(var y = 0; y < activeDocument.pages[x].textFrames.length; y ++) {
            deleteEmptyTextFrame(activeDocument.pages[x].textFrames[y])
            if(isMultiTextFrame(activeDocument.pages[x].textFrames[y]) === false) {
                deleteOversetedTextFrame(activeDocument.pages[x].textFrames[y])
            }
        }
        groupChapter622(activeDocument.pages[x])
        makeGNPerPage(activeDocument.pages[x])
    }

    spreadOversetedGN();

    for(var x = 0; x < activeDocument.pages.length; x ++) {
        for(var y = 0; y < activeDocument.pages[x].groups.length; y ++) {
            relocateGroups(activeDocument.pages[x].groups[y])
        }
        for(var y = 0; y < activeDocument.pages[x].textFrames.length; y ++) {
            relocateTextFrames(activeDocument.pages[x].textFrames[y])
        }
        for(var y = 0; y < activeDocument.pages[x].textFrames.length; y ++) {
            spreadGNs(activeDocument.pages[x].textFrames[y])
        }
    }

    for(var x = activeDocument.pages.length - 1; x >= 0; x --) {
        deleteBlankPage(activeDocument.pages[x])
    }

    for(var x = activeDocument.pages.length - 1; x >= 0; x --) {
        for(var y = 0; y < activeDocument.pages[x].allPageItems.length; y ++) {
            sendItemToBack(activeDocument.pages[x].textFrames[y])
            applyPictureParagraphStyle(activeDocument.pages[x].allPageItems[y])
        }
    }
} else {
    splitOversetedGNNumberTextFrame()
}

preflight()


//////////////////////////////////////////////////////////////////////////////////////    FUNCTION    /////////////////////////////////////////////////////////////////////////////////////

function setDocumentPreference() {
    activeDocument.documentPreferences.allowPageShuffle = true
    activeDocument.spreads.everyItem().allowPageShuffle = true
    activeDocument.sections[0].continueNumbering = false
    activeDocument.sections[0].pageNumberStart = 1
    activeDocument.documentPreferences.facingPages = false
    activeDocument.zeroPoint = [0, 0]
    activeDocument.zeroPoint = ["2p2.172", "4p7"]
}


function deleteMasterPageItems() {
    for(var x = 0; x < activeDocument.masterSpreads.length; x ++){
        for(var i = activeDocument.masterSpreads[x].allPageItems.length - 1; i >= 0; i --) {
            activeDocument.masterSpreads[x].allPageItems[i].remove()
        }
    }
}


function checkOversetedGNNumberTextFrame() {
    app.findGrepPreferences.appliedParagraphStyle = "";
    app.findGrepPreferences.findWhat = "(?<=\\r)(第\\d+號公告\\r第)";
    myFind = activeDocument.findGrep();
    if(myFind.length > 0) {
        oversetedGNNumberTextFrame = true
        splitOversetedGNNumberTextFrame(myFind[0])
    } else {
        app.findGrepPreferences.findWhat = "(?<=\\r)(G.N. \\d+\\rG.N.)"; 
        myFind = activeDocument.findGrep();
        if(myFind.length > 0) {
            oversetedGNNumberTextFrame = true
            splitOversetedGNNumberTextFrame(myFind[0])
        }
    }
}


function ungroup(page) {
    for(var y = 0; y < page.groups.length; y ++) {
        page.groups[y].ungroup()
    }
}    


function deleteEmptyTextFrame(textFrame) {
    if(textFrame.contents === ""){
        textFrame.remove();
    }
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
            case "APPOINTMENTS" :
            case "NOTICES" :
            case "TENDERS" :
            case "NOTICES REPEATED" :
            case "End of Main Gazette of this issue." :
            case "\tIndex Numbers\rSelected materials" :
            case "\tIndex Numbers\rLabour" :
                page.textFrames[y].remove()
        }
    }
}


function isMultiTextFrame(textFrames) {
    try {
        if(textFrames.startTextFrame.contents.toString().slice(0, 5) === "G.N. "
        || textFrames.startTextFrame.contents.toString().slice(1, 6) === "G.N. "
        || textFrames.startTextFrame.contents.toString().slice(2, 7) === "G.N. ") {
            return false
        } else if(textFrames.startTextFrame.contents.toString().slice(0, 1) === "第"
        || textFrames.startTextFrame.contents.toString().slice(1, 2) === "第"
        || textFrames.startTextFrame.contents.toString().slice(2, 3) === "第") {
            return false
        }
        return true
    } catch(e) {}
}


function deleteOversetedTextFrame(textFrame) {
    if(textFrame.parentStory.textContainers.length > 1){
        for(var z = textFrame.parentStory.textContainers.length - 1; z > 0; z --){
            textFrame.parentStory.textContainers[z].remove();
        }
    } 
}


function groupChapter622(page) {
    var items =page.textFrames
    for(var y = 0; y < items.length; y ++) { 
        if(items[y].contents.toString().slice(7, 79) == "Companies Registry\rCompanies Ordinance \(Chapter 622\)\rPursuant to section"
            || items[y].contents.toString().slice(8, 80) == "Companies Registry\rCompanies Ordinance \(Chapter 622\)\rPursuant to section"
            || items[y].contents.toString().slice(9, 81) == "Companies Registry\rCompanies Ordinance \(Chapter 622\)\rPursuant to section"
            || items[y].contents.toString().slice(10, 82) == "Companies Registry\rCompanies Ordinance \(Chapter 622\)\rPursuant to section"
            || items[y].contents.toString().slice(11, 83) == "Companies Registry\rCompanies Ordinance \(Chapter 622\)\rPursuant to section"
            || items[y].contents.toString().slice(2, 24) == "號公告\r公司註冊處\r公司條例\(第622章\)\r"
            || items[y].contents.toString().slice(3, 25) == "號公告\r公司註冊處\r公司條例\(第622章\)\r"
            || items[y].contents.toString().slice(4, 26) == "號公告\r公司註冊處\r公司條例\(第622章\)\r"
            || items[y].contents.toString().slice(5, 27) == "號公告\r公司註冊處\r公司條例\(第622章\)\r"
            || items[y].contents.toString().slice(6, 28) == "號公告\r公司註冊處\r公司條例\(第622章\)\r") {
                if(items[y].startTextFrame) {
                    try {
                        if(items[y].geometricBounds[0] < items[y + 1].geometricBounds[0]) {
                            if(items[y + 1].contents.toString().slice(0, 4) !== "G.N." || items[y].contents.toString().slice(0, 2) !== "第\\d") {
                                activeDocument.groups.add([items[y], items[y + 1]])
                            }
                        }
                    } catch(e) {}
                }
        }
    }
}


function makeGNPerPage(page) {
    while(page.textFrames.length > 1) {
        activeDocument.pages.add(LocationOptions.AFTER, page)
        page.textFrames[0].move(activeDocument.pages[x + 1])
    }
    if(page.textFrames.length > 0){
        for(var y = 0; y <= page.groups.length; y ++) {
            activeDocument.pages.add(LocationOptions.AFTER, page)
            try{
                page.groups[0].move(activeDocument.pages[x + 1])
            } catch(e) {}
        }
    }
}


function spreadGNs(textFrame) {
    if(textFrame.overflows) {
        var nextThreadingTextFrame = activeDocument.pages[x].textFrames.add()
        if(textFrame.contents.toString().slice(0, 1) !== "\\w" && activeDocument.pages[x].textFrames.length > 1) {
            y += 1
        }   
        nextThreadingTextFrame.geometricBounds = ["0", "28p0", "504", "0"];  
        textFrame.nextTextFrame = nextThreadingTextFrame
        activeDocument.pages.add(LocationOptions.AFTER, activeDocument.pages[x])
        nextThreadingTextFrame.move(activeDocument.pages[x + 1])
    }
}


function relocateTextFrames(textFrame) {
    textFrame.geometricBounds = ["0", "28p0", "504", "0",];    //["上", "左", "下", "右",]
}


function relocateGroups(group) {
    group.geometricBounds = ["0", "28p0", activeDocument.pages[x].groups[y].geometricBounds[2] - activeDocument.pages[x].groups[y].geometricBounds[0], "0",];
}


function deleteBlankPage(page) {
    if(page.allGraphics.length === 0 && page.textFrames.length === 0 && page.groups.length === 0){
        page.remove()
    } else if(page.textFrames.length === 1 && page.textFrames[0].startTextFrame.contents === "") {
        page.remove()
    }
}

function sendItemToBack(item) {
    try {
        if(item.contents.toString().slice(0, 5) === "G.N. "
        || item.contents.toString().slice(1, 6) === "G.N. "
        || item.contents.toString().slice(2, 7) === "G.N. ") {
            item.sendToBack()
        } else if(item.contents.toString().slice(0, 1) === "第"
        || item.contents.toString().slice(1, 2) === "第"
        || item.contents.toString().slice(2, 3) === "第") {
            item.sendToBack()
        }
    } catch(e) {}
}


function spreadOversetedGN() {
	var scriptVersion = app.scriptPreferences.version;
	app.scriptPreferences.version = 8.0;
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
    app.findChangeGrepOptions.properties = ({includeHiddenLayers:true, includeMasterPages:true, includeFootnotes:true, kanaSensitive:true, widthSensitive:true});
    app.findGrepPreferences.properties = ({});
	try {
		app.findGrepPreferences.appliedParagraphStyle = getParagraphStyle(activeDocument, 'ab  BM-Item--Head', 'paragraphStyles');
		app.changeGrepPreferences.properties = ({startParagraph:1885500011});
		activeDocument.changeGrep();
	} catch (e) {}
	try {
		app.findGrepPreferences.appliedParagraphStyle = getParagraphStyle(activeDocument, 'ab  BM-Item--Head_no 5pt', 'paragraphStyles');
		app.changeGrepPreferences.properties = ({startParagraph:1885500011});
		activeDocument.changeGrep();
	} catch (e) {}
	try {
		app.findGrepPreferences.appliedParagraphStyle = getParagraphStyle(activeDocument, 'ab  BM-Item--Head-E', 'paragraphStyles');
		app.changeGrepPreferences.properties = ({startParagraph:1885500011});
		activeDocument.changeGrep();
	} catch (e) {}
	try {
		app.findGrepPreferences.appliedParagraphStyle = getParagraphStyle(activeDocument, 'ab  BM-Item--Head-E no 4pt', 'paragraphStyles');
		app.changeGrepPreferences.properties = ({startParagraph:1885500011});
		activeDocument.changeGrep();
	} catch (e) {}
	app.findGrepPreferences = NothingEnum.NOTHING;
	app.changeGrepPreferences = NothingEnum.NOTHING;
	app.scriptPreferences.version = scriptVersion;
};


function getParagraphStyle(document, paragraphStyleName, property) {
	stringResult = paragraphStyleName.match (/^(.*?[^\\]):(.*)$/);
	var styleName = (stringResult) ? stringResult[1] : paragraphStyleName;
	styleName = styleName.replace (/\\:/g, ':');
	remainingString = (stringResult) ? stringResult[2] : '';
	var newProperty = (stringResult) ? property.replace(/s$/, '') + 'Groups' : property;
	var styleOrGroup = document[newProperty].itemByName(styleName);
	if (remainingString.length > 0 && styleOrGroup.isValid) styleOrGroup = getParagraphStyle (styleOrGroup, remainingString, property);
	return styleOrGroup;
};


function applyPictureParagraphStyle(item) {
    try {
        if(activeDocument.pages[x].allPageItems[y].geometricBounds[0] < 0) {      
            item.parent.appliedParagraphStyle = "Picture";         
        }
    } catch(e) {}
}       
       
       
function splitOversetedGNNumberTextFrame(result) {
    try{
        result.select()
        app.doScript('C:/Program Files (x86)/Adobe/Adobe InDesign CS6/Scripts/Scripts Panel/Samples/JavaScript/SplitStory.jsx', ScriptLanguage.JAVASCRIPT)
    } catch(e) {}
}
       
       
function preflight() {
    var preflightProfile = app.preflightProfiles[1]
    var preflightProcess = app.preflightProcesses.add(activeDocument, preflightProfile)
    preflightProcess.waitForProcess()
    preflightResult = preflightProcess.processResults
    if(preflightResult.indexOf('None') === 0) {
        alert("已完成\r請檢查後執行 \"大報 - 上網 - 2\"")
    } else {
        alert("檔案有問題")
    }
}

/////////////////////////////////////////////////////////////////////////////////////////////////////    TRASH    /////////////////////////////////////////////////////////////////////////////////////////////////////

function groupDoubleTextFrame(page) {
    if(page.allGraphics.length > 0) {
        var gns = []
        for(var y = 0; y <= page.textFrames.length; y ++) {
            try{
                if(page.textFrames[y].contents.toString().slice(0, 5) === "G.N. "
                || page.textFrames[y].contents.toString().slice(1, 6) === "G.N. "
                || page.textFrames[y].contents.toString().slice(2, 7) === "G.N. "
                || page.textFrames[y].contents.toString().slice(0, 1) === "第"
                || page.textFrames[y].contents.toString().slice(1, 2) === "第"
                || page.textFrames[y].contents.toString().slice(2, 3) === "第") {
                    gns.push(page.textFrames[y])
                }
            } catch(e) {}
        } 
        for(var z = 0; z <= page.allGraphics.length; z ++) {
            if(gns.length === 1 && page.groups.length === 0) {
                page.groups.add(page.pageItems)
            } else if(gns.length > 1) {
                alert(page.allGraphics[z].parent)
                alert("此檔案需要手動")
            }
        }
    }
}




