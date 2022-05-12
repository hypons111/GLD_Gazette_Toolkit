﻿///////////    Version 5    ///////////
///////////    2021-5-17    ///////////

const activeDocument = app.activeDocument



setDocumentPreference()
deleteMasterPageItems()

for(var x = 0; x < activeDocument.pages.length; x ++) {
    ungroup(activeDocument.pages[x])
    deleteHeader(activeDocument.pages[x])
    for(var y = 0; y < activeDocument.pages[x].textFrames.length; y ++) {
        if(isDoubleTextFrame(activeDocument.pages[x].textFrames[y]) === false) {
            deleteOversetedTextFrame(activeDocument.pages[x].textFrames[y])
        }
        groupChapter622()
        makeGNPerPage(activeDocument.pages[x])
    }
}

spreadChapter380()

for(var x = 0; x < activeDocument.pages.length; x ++) {
    for(var y = 0; y < activeDocument.pages[x].textFrames.length; y ++) {
        spreadGNs(activeDocument.pages[x].textFrames[y])
    }
    for(var y = 0; y < activeDocument.pages[x].groups.length; y ++) {
        relocateGroups(activeDocument.pages[x].groups[y])
    }
    for(var y = 0; y < activeDocument.pages[x].textFrames.length; y ++) {
        relocateTextFrames(activeDocument.pages[x].textFrames[y])
        deleteBlankPage(activeDocument.pages[x].textFrames[y])
    }
}



//preflight()


//////////////////////////////////////////////////////////////////////////////////////    FUNCTION    /////////////////////////////////////////////////////////////////////////////////////

function setDocumentPreference() {
    activeDocument.documentPreferences.allowPageShuffle = true
    activeDocument.sections[0].continueNumbering = false
    activeDocument.sections[0].pageNumberStart = 1
    activeDocument.documentPreferences.facingPages = false
}


function deleteMasterPageItems() {
    for(var x = 0; x < activeDocument.masterSpreads.length; x ++){
        for(var i = activeDocument.masterSpreads[x].allPageItems.length - 1; i > 0; i --) {
            activeDocument.masterSpreads[x].allPageItems[i].remove()
        }
    }
}


function ungroup(page) {
    for(var y = 0; y < page.groups.length; y ++) {
        page.groups[y].ungroup()
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
            case "APPOINTMENTS" :
            case "NOTICES" :
            case "TENDERS" :
            case "NOTICES REPEATED" :
            case "End of Main Gazette of this issue." :
                page.textFrames[y].remove()
        }
    }
}


function isDoubleTextFrame(textFrames) {
    switch("G.N. ") {
        case textFrames.startTextFrame.contents.toString().slice(0, 5) :
        case textFrames.startTextFrame.contents.toString().slice(1, 6) :
            return false
    }
    switch("第") {
        case textFrames.startTextFrame.contents.toString().slice(0, 1) :
        case textFrames.startTextFrame.contents.toString().slice(1, 2) :
            return false
    }
}


function deleteOversetedTextFrame(textFrame) {
    if(textFrame.parentStory.textContainers.length > 1){
        for(var z = textFrame.parentStory.textContainers.length - 1; z > 0; z --){
            textFrame.parentStory.textContainers[z].remove();
        }
    } 
}


function groupChapter622() {
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


function makeGNPerPage(page) {
    while(page.textFrames.length > 1) {
        activeDocument.pages.add(LocationOptions.AFTER, page)
        page.textFrames[0].move(activeDocument.pages[x + 1])
    }
    if(page.textFrames.length > 0){
        for(var y = 0; y < page.groups.length; y ++) {
            activeDocument.pages.add(LocationOptions.AFTER, page)
            page.groups[y]
            page.groups[y].move(activeDocument.pages[x + 1])
        }
    }
}


function spreadGNs(textFrame) {
    if(textFrame.overflows) {
        var nextThreadingTextFrame = activeDocument.pages[x].textFrames.add()
        if(textFrame.contents.toString().slice(0, 1) !== "\\w" && activeDocument.pages[x].textFrames.length > 1) {
            y += 1
        }         
        textFrame.nextTextFrame = nextThreadingTextFrame
        nextThreadingTextFrame.geometricBounds = ["0", "28p0", "494.999", "0"];
        activeDocument.pages.add(LocationOptions.AFTER, activeDocument.pages[x])
        nextThreadingTextFrame.move(activeDocument.pages[x + 1])
    }
}


function spreadChapter380() {
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


function relocateTextFrames(textFrame) {
    textFrame.geometricBounds = ["0", "28p0", "494.999", "0",];    //["上", "左", "下", "右",]
}


function relocateGroups(group) {
    group.geometricBounds = ["0", "28p0", activeDocument.pages[x].groups[y].geometricBounds[2] - activeDocument.pages[x].groups[y].geometricBounds[0], "0",];
}


function deleteBlankPage(item) {
    try {
        if(!item){
            activeDocument.pages[x].remove()
        }
    } catch(e) {}
}


function deleteBlankPage() {
    for(var x = activeDocument.pages.length - 1; x >= 0; x --) {
        if(activeDocument.pages[x].allGraphics.length === 0 && activeDocument.pages[x].textFrames.length === 0){
            activeDocument.pages[x].remove()
        }
        try{    
            for(var y = 0; y < activeDocument.pages[x].textFrames.length; y ++) {
                if(activeDocument.pages[x].textFrames[y].contents === ""){
//                    activeDocument.pages[x].remove()
                }
            }
        } catch(e) {}        
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