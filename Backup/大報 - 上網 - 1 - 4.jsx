///////////    Version 4    ///////////
///////////    2021-5-16    ///////////

const activeDocument = app.activeDocument

activeDocument.documentPreferences.allowPageShuffle = true
activeDocument.sections[0].continueNumbering = false
activeDocument.sections[0].pageNumberStart = 1
activeDocument.documentPreferences.facingPages = false

deleteMasterPageItems()

seperateThreadingGNs()

for(var x = 0; x < activeDocument.pages.length; x ++) {
    for(var y = 0; y < activeDocument.pages[x].textFrames.length; y ++) {
        app.activeDocument.pages[x].groups.everyItem().ungroup()
        deleteHeader(activeDocument.pages[x].textFrames[y])
        deleteOversetedTextFrame(activeDocument.pages[x].textFrames[y])
    }
}

groupCompany()

deleteBlankPage()

for(var i = 0; i < activeDocument.pages.length; i ++) {
    seperateGNs()
    positionTextFrames()
    spreadThreadingText()
//    positionTextFrames()
}

preflight()


function deleteMasterPageItems() {
    while(activeDocument.masterSpreads.item("A-主版").textFrames.length > 0) {
        activeDocument.masterSpreads.item("A-主版").textFrames[0].remove()
    } 
}


function deleteHeader(header) {
    if(header.contents === "委任令" 
      || header.contents === "公  告" 
      || header.contents === "招  標"
      || header.contents === "再登的公告"
      || header.contents === "本期憲報全文完"
      || header.contents === "APPOINTMENTS"
      || header.contents === "NOTICES" 
      || header.contents === "TENDERS"
      || header.contents === "NOTICES REPEATED"
      || header.contents === "End of Main Gazette of this issue.") {
        header.remove()
    }
}


function seperateThreadingGNs() {
    app.findGrepPreferences.appliedParagraphStyle = "";
    app.findGrepPreferences.findWhat = "(?<=\\r)(第\\d+號公告\\r)";
    myFind = activeDocument.findGrep();
    app.changeGrepPreferences.changeTo = "~P$1"
    activeDocument.changeGrep()
    if(myFind.length === 0) {
        app.findGrepPreferences.findWhat = "(?<=\\r)(G.N. \\d+\\r)"; 
        myFind = activeDocument.findGrep();
        app.changeGrepPreferences.changeTo = "~P$1"
        activeDocument.changeGrep()
    }
}


function deleteOversetedTextFrame(textFrame) {
    var selection = textFrame;
	try {
        if(selection.parentStory.textContainers.length > 1){
            for(var i = selection.parentStory.textContainers.length - 1; i > 0; i --){
                selection.parentStory.textContainers[i].remove();
            }
        } 
    } catch (e) {}
}


function deleteBlankPage() {
    for(var x = activeDocument.pages.length - 1; x >= 0; x --) {
        if(activeDocument.pages[x].allGraphics.length === 0 && activeDocument.pages[x].textFrames.length === 0){
            activeDocument.pages[x].remove()
        }
    }
}


function positionTextFrames() {
	try {    
//        activeDocument.pages[i].groups[0].geometricBounds = ["0", "0", "494.999", "28p0",];    //["上", "左", "下", "右",]
    } catch (e) {}
    try {    
        activeDocument.pages[i].textFrames[0].geometricBounds = ["0", "0", "494.999", "28p0",];    //["上", "左", "下", "右",]
    } catch (e) {}
    try {    
        activeDocument.pages[x].pageItems[y].geometricBounds = ["0", "0", "494.999", "28p0",];    //["上", "左", "下", "右",]
    } catch (e) {}
}


function seperateGNs() {
    if(activeDocument.pages[i].textFrames.length > 1) {
        while(activeDocument.pages[i].textFrames.length > 1) {
            activeDocument.pages.add(LocationOptions.AFTER, activeDocument.pages[i])
            activeDocument.pages[i].textFrames[0].duplicate(activeDocument.pages[i + 1])
            activeDocument.pages[i].textFrames[0].remove()
        }
    }
}


function spreadThreadingText() {
	try {
        if(activeDocument.pages[i].textFrames[0].overflows) {
            var newTextFrame = activeDocument.pages.item(i).textFrames.add().geometricBounds = ["0", "0", "494.999", "28p0",];    //["上", "左", "下", "右"]
            activeDocument.pages[i].textFrames.everyItem().select()
            a = app.selection[1]
            b = app.selection[0]
            a.nextTextFrame = b
            activeDocument.pages.add(LocationOptions.AFTER, activeDocument.pages[i])
            activeDocument.pages[i].textFrames[0].move(activeDocument.pages[i + 1])
        }
    } catch (e) {}
}

function groupCompany() {
    for(var x = 0; x < activeDocument.pages.length; x ++) {
        var items = activeDocument.pages[x].textFrames
        for(var y = 0; y < items.length; y ++) {
           if(items[y].contents.toString().slice(7, 79) == "Companies Registry\rCompanies Ordinance \(Chapter 622\)\rPursuant to section"
           || items[y].contents.toString().slice(8, 80) == "Companies Registry\rCompanies Ordinance \(Chapter 622\)\rPursuant to section"
           || items[y].contents.toString().slice(9, 81) == "Companies Registry\rCompanies Ordinance \(Chapter 622\)\rPursuant to section"
           || items[y].contents.toString().slice(10, 82) == "Companies Registry\rCompanies Ordinance \(Chapter 622\)\rPursuant to section"
           || items[y].contents.toString().slice(11, 83) == "Companies Registry\rCompanies Ordinance \(Chapter 622\)\rPursuant to section") {
                if(items[y].startTextFrame) {
                    try {
                        if(items[y].geometricBounds[0] < items[y + 1].geometricBounds[0]) {
                            if(items[y + 1].contents.toString().slice(0, 4) !== "G.N." || items[y].contents.toString().slice(0, 4) !== "G.N.") {
                                activeDocument.groups.add([items[y], items[y + 1]])
                            }
                        }
                    } catch(e) {}
                }
            }
        }
    }
}

function preflight() {
    var preflightProfile = app.preflightProfiles[1]
    var preflightProcess = app.preflightProcesses.add(activeDocument, preflightProfile)
    preflightProcess.waitForProcess()
    preflightResult = preflightProcess.processResults
    if(preflightResult.indexOf('None') === 0) {
    } else { alert('檔案有問題') }
}



