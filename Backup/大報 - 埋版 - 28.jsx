///////////    Version 28    ///////////

/////////////////////////////////////////////////////////////////////////////////////////////////////    LOG    /////////////////////////////////////////////////////////////////////////////////////////////////////

///////////    Version 3    ///////////    2021-6-7    ///////////
//added openIndd(), preparationBeforeCopy(), copy(), close(), paste(), addNewPage(), moveItem()

///////////    Version 4    ///////////    2021-6-8    ///////////
//redesigned the logic of the process

///////////    Version 5    ///////////    2021-6-9    ///////////
//redesigned the logic of the process

///////////    Version 6    ///////////    2021-6-10    ///////////
//redesigned the logic of the process and basically works

///////////    Version 7    ///////////    2021-6-10    ///////////
//orderList can use hypen 
//deleteLable() includes "詳校" 

///////////    Version 8    ///////////    2021-6-15    ///////////
//can insert gns one by one, but no UI
//changed deleteLable() delete rules
//if the current gn is the first gn at the current page, do not move its geometricBounds[0], and dont not cut its next text frame even its shorter than 28.661

///////////    Version 9    ///////////    2021-6-16    ///////////
//can choose auto or manual 
//functionized manual part 

///////////    Version 10    ///////////    2021-6-21    ///////////
//automatic define auto or manual by searching list.txt on desktop
//added a loop for getting the new GN number

///////////    Version 11    ///////////    2021-6-21    ///////////
//rearrange automatic part scripts
//list text can use all capital

///////////    Version 12    ///////////    2021-6-21    ///////////
//fixed the bug when cancel getNewestTop() after entered a number
//cancaled preserve space for appointment\
//manaul orderList can use hypen 
//about when not threaded text frame

///////////    Version 13    ///////////    2021-6-22    ///////////
//made input gn number slot active

///////////    Version 14    ///////////    2021-6-24    ///////////
//fixed images need or dont need to apply picture paragraph style

///////////    Version 15    ///////////    2021-7-5    ///////////
//fixed the first gn in new page doesnt fit to contents
//added squeezes gn when have enough space
//added manaul mode has to launch everytime
//added addNewPage()

///////////    Version 16    ///////////    2021-7-7    ///////////
//fixed getNewestTop() cant get coordinate from groups

///////////    Version 17    ///////////    2021-7-7    ///////////
//added try catch to handle and alert errors

///////////    Version 18    ///////////    2021-7-12    ///////////
//added insertHeader() for automatic
//added rename list file name after used 
//fixed deleteLable() bug

///////////    Version 19    ///////////    2021-7-13    ///////////
//added ~ for file gn number
//handled those error alerts

///////////    Version 20    ///////////    2021-7-13    ///////////
//added zoom back to the starting page after insert gns

///////////    Version 21    ///////////    2021-7-14    ///////////
//fixed text frame after header to zero spacing after rearrange excution order

///////////    Version 22    ///////////    2021-7-15    ///////////
//improved isThreadTextFrame() performance

///////////    Version 23    ///////////    2021-7-20    ///////////
//added copyAndDeleteText() to fix list.txt close() too slow 
//changed list.txt name format after used

///////////    Version 24    ///////////    2021-7-26    ///////////
//added isThreadingGNNumberTextFrame()

///////////    Version 25    ///////////    2021-7-26    ///////////
//redesign isThreadTextFrame() logic

///////////    Version 26    ///////////    2021-7-28    ///////////
//combined isThreadTextFrame() and isThreadingGNNumberTextFrame() into isSingleTextFrame()
//redesign isSingleTextFrame() logic
//redesign deleteLable() logic

///////////    Version 27    ///////////    2021-8-18    ///////////
//made openFile() accept "_col", "_Col", "_COL", "-col", "-Col", "-COL" in file name
//added new logic to modifyGN()

///////////    Version 28    ///////////    2021-9-14    ///////////
//added return false to openFile() to abort the execution when cant open the correct file.

/////////////////////////////////////////////////////////////////////////////////////////////////////    LOG    /////////////////////////////////////////////////////////////////////////////////////////////////////
const mainDocument = app.activeDocument
const issue = mainDocument.name.toString()[3] + mainDocument.name.toString()[4]
const version = mainDocument.name.toString()[6]
const versionFullName = getVersionFullName(version)
const spacing = 45
var headerSwitch = false
var currentTop = getNewestTop()
const orderList = []
const orderListTextFile = File('~/Desktop/list.txt') || File('~/Desktop/List.txt') || File('~/Desktop/LIST.txt')
const activePage = app.activeWindow.activePage.name


if(isAuto()) {
    if(insertHeader(getHeader(version))) {
        for (var x = 0; x < orderList.length; x ++) {
            try{
                if(openFile() === false) {
                    break
                }
                var activeDocument = openFile()
                ungroup(activeDocument)
                deleteLable(activeDocument)
                if(isSingleTextFrame(activeDocument) && activeDocument.name.toString()[0] + activeDocument.name.toString()[1] != "MA") {
//                    applyPictureParagraphStyle(activeDocument)
                    deletePagesFromePageTwo(activeDocument)
                    app.selection = activeDocument.pages[0].pageItems
                    app.copy()
                    activeDocument.close(SaveOptions.no)
                    documentSettings()
                    addNewPage()
                    mainDocument.pages.item(-1).select()
                    mainDocument.select(NothingEnum.NOTHING)
                    app.paste()
                    modifyGN()
                }
            } catch(e) {
                alert(e)
            }
        }
    copyAndDeleteText()
    }
    mainDocument.documentPreferences.facingPages = true
    app.activeWindow.activePage = app.activeDocument.pages.item(activePage)
    app.activeDocument.layoutWindows[0].zoom(ZoomOptions.FIT_SPREAD)
} else {
    if(getGNNumber(inputGNNumber())) {
        try{
            modifyNumber(orderList)
            for (var x = 0; x < orderList.length; x ++) {
                if(openFile() === false) {
                    break
                }
                var activeDocument = openFile()
                ungroup(activeDocument)
                deleteLable(activeDocument)
                if(isSingleTextFrame(activeDocument) && activeDocument.name.toString()[0] + activeDocument.name.toString()[1] != "MA") {
//                    applyPictureParagraphStyle(activeDocument)
                    deletePagesFromePageTwo(activeDocument)
                    app.selection = activeDocument.pages[0].pageItems
                    app.copy()
                    activeDocument.close(SaveOptions.no)
                    documentSettings()
                    addNewPage()
                    mainDocument.pages.item(-1).select()
                    mainDocument.select(NothingEnum.NOTHING)
                    app.paste()
                    modifyGN()
                }
            }
        } catch(e) {
            alert(e)
        }
    }
    mainDocument.documentPreferences.facingPages = true
    app.activeDocument.layoutWindows[0].zoom(ZoomOptions.FIT_SPREAD)
}

/////////////////////////////////////////////////////////////////////////////////////////////////////    FUNCTIONS    /////////////////////////////////////////////////////////////////////////////////////////////////////

function getVersionFullName(v) {
    if(v === "C") {
        return  "Chinese"
    } else if (v === "E") {
        return "English"
    }
}


function getNewestTop() {
    try {
        if(headerSwitch === true
         || mainDocument.pages[getNewestPage()].pageItems[0].geometricBounds[2] - mainDocument.pages[getNewestPage()].pageItems[0].geometricBounds[0] === 24
         || mainDocument.pages[getNewestPage()].pageItems[0].geometricBounds[2] - mainDocument.pages[getNewestPage()].pageItems[0].geometricBounds[0] === 18) {
             if(mainDocument.pages[getNewestPage()].pageItems[0].contents.toString()[0] === "委"
             || mainDocument.pages[getNewestPage()].pageItems[0].contents.toString()[0] === "公"
             || mainDocument.pages[getNewestPage()].pageItems[0].contents.toString()[0] === "招"
             || mainDocument.pages[getNewestPage()].pageItems[0].contents.toString()[0] === "A"
             || mainDocument.pages[getNewestPage()].pageItems[0].contents.toString()[0] === "N"
             || mainDocument.pages[getNewestPage()].pageItems[0].contents.toString()[0] === "T") {
                return mainDocument.pages[getNewestPage()].pageItems[0].geometricBounds[2]
              }
        } else {
            return mainDocument.pages[getNewestPage()].pageItems[0].geometricBounds[2] + spacing
        }
    } catch(e) {
        return 0
    } 
}


function documentSettings() {
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.neverInteract;
    mainDocument.documentPreferences.allowPageShuffle = true
    mainDocument.spreads.everyItem().allowPageShuffle = true
    mainDocument.documentPreferences.facingPages = false
}


function isAuto() {
    getOrderList()
    modifyNumber(orderList)
    return orderList.length
}


function inputGNNumber() {
    var dialog = new Window('dialog', "", "x:0, y:0, width:290, height:75");
    dialog.panel = dialog.add('panel', [10, 10, 185, 65], "");

    dialog.panel.add('statictext',  [12, 18, 80, 198], "公告編號"); 
    dialog.gN = dialog.panel.add('edittext', [80, 13, 155, 38], "");
    dialog.gN.active = true
    dialog.gN.onChange = function() {
        gnNumber = dialog.gN.text
    }

    dialog.okButton = dialog.add('button', [200, 10, 275, 35], "OK");
    dialog.okButton.onClick = function() { dialog.close(1) }
    dialog.cancelButton = dialog.add('button', [200, 40, 275, 65], "Cancel");
    dialog.cancelButton.onClick = function () {
        dialog.close(0)
        gnNumber = 0
        return false
    }
    dialog.center(); 
    dialog.show()
    try {
        return gnNumber;
    } catch(e) {
        alert("請輸入公告編號")
    }
}


function getOrderList() {
    orderListTextFile.open("r")
    const temp = []
    for (var x = 0; x < orderListTextFile.length; x ++) {
        temp.push(orderListTextFile.readln())
    } 
    for (var y = 0; y < temp.length; y ++) {
        if(temp[y].slice(1, 2) === "~" && Number(temp[y]) !== 0) {
            orderList.push(temp[y].replace("~", "-"))
        } else if(temp[y].slice(2, 3) === "~" && Number(temp[y]) !== 0) {
            orderList.push(temp[y].replace("~", "-"))
        } else if(temp[y].slice(3, 4) === "~" && Number(temp[y]) !== 0) {
            orderList.push(temp[y].replace("~", "-"))
        } else if(temp[y].slice(4, 5) === "~" && Number(temp[y]) !== 0) {
            orderList.push(temp[y].replace("~", "-"))
        } else if(temp[y].slice(1, 2) === "-" && Number(temp[y]) !== 0) {
            for(var z = Number(temp[y].slice(0, 1)); z <= Number(temp[y].slice(2, temp[y].length)); z ++) {
                orderList.push(z)
            }
        } else if(temp[y].slice(2, 3) === "-" && Number(temp[y]) !== 0) {
            for(var z = Number(temp[y].slice(0, 2)); z <= Number(temp[y].slice(3, temp[y].length)); z ++) {
                orderList.push(z)
            }
        } else if(temp[y].slice(3, 4) === "-" && Number(temp[y]) !== 0) {
            for(var z = Number(temp[y].slice(0, 3)); z <= Number(temp[y].slice(4, temp[y].length)); z ++) {
                orderList.push(z)
            }
        } else if(temp[y].slice(4, 5) === "-" && Number(temp[y]) !== 0) {
            for(var z = Number(temp[y].slice(0, 4)); z <= Number(temp[y].slice(5, temp[y].length)); z ++) {
                orderList.push(z)
            }
        } else if(Number(temp[y]) !== 0) {
            orderList.push(temp[y])
        }
    } 
}



function copyAndDeleteText() {
    orderListTextFile.copy('~/Desktop/' + issue + '期 ' + new Date().getHours() + "點 " + new Date().getMinutes() + "分 " + new Date().getMinutes() + '秒.txt')
    orderListTextFile.remove()
}


function getGNNumber(number) {
    orderList.splice(0, orderList.length)
    try{
        if(number.slice(1, 2) === "~" && Number(number) !== 0) {
            orderList.push(number.replace("~", "-"))
        } else if(number.slice(2, 3) === "~" && Number(number) !== 0) {
            orderList.push(number.replace("~", "-"))
        } else if(number.slice(3, 4) === "~" && Number(number) !== 0) {
            orderList.push(number.replace("~", "-"))
        } else if(number.slice(4, 5) === "~" && Number(number) !== 0) {
            orderList.push(number.replace("~", "-"))
        } else if(number.slice(1, 2) === "-" && Number(number) !== 0) {
            for(var z = Number(number.slice(0, 1)); z <= Number(number.slice(2, number.length)); z ++) {
                orderList.push(z)
            }
        } else if(number.slice(2, 3) === "-" && Number(number) !== 0) {
            for(var z = Number(number.slice(0, 2)); z <= Number(number.slice(3, number.length)); z ++) {
                orderList.push(z)
            }
        } else if(number.slice(3, 4) === "-" && Number(number) !== 0) {
            for(var z = Number(number.slice(0, 3)); z <= Number(number.slice(4, number.length)); z ++) {
                orderList.push(z)
            }
        } else if(number.slice(4, 5) === "-" && Number(number) !== 0) {
            for(var z = Number(number.slice(0, 4)); z <= Number(number.slice(5, number.length)); z ++) {
                orderList.push(z)
            }
        } else if(Number(number) !== 0) {
            orderList.push(number)
        }
        return true
    } catch(e) {
        return false
    }
}




function modifyNumber(number) {
    for (var x = 0; x < number.length; x ++) {
        if (number[x] < 10 ) {
            number[x] = "00" + number[x]
        } else if (number[x] < 100 ) {
            number[x] = "0" + number[x]
        }
    }
}


function insertHeader(t) {
    var header
    var height
    try{
        if(t.toString()[0] === "無") {
            return true
        } else if(t.toString()[0] !== "無" && t !== false) {
            switch(t.toString()[0]) {
                case "委" :
                    header = document.masterSpreads.item(2).textFrames[0]
                    height = "2p"
                    break;
                case "公" :
                    if(currentTop !== 0) {
                        currentTop = 0
                        mainDocument.pages.add()
                    }
                    header = document.masterSpreads.item(3).textFrames[0]
                    height = "2p"
                    break;
                case "招" :
                    if(currentTop !== 0) {
                        currentTop = 0
                        mainDocument.pages.add()
                    }
                    header = document.masterSpreads.item(4).textFrames[0]
                    height = "2p"
                    break;
                case "A" :
                    header = document.masterSpreads.item(2).textFrames[0]
                    height = "1p6"
                    break;
                case "N" :
                    if(currentTop !== 0) {
                        currentTop = 0
                        mainDocument.pages.add()
                    }
                    header = document.masterSpreads.item(3).textFrames[0]
                    height = "1p6"
                    break;
                case "T" :
                    if(currentTop !== 0) {
                        currentTop = 0
                        mainDocument.pages.add()
                    }
                    header = document.masterSpreads.item(4).textFrames[0]
                    height = "1p6"
                    break;
            }
            header.move(mainDocument.pages[mainDocument.pages.length - 1])
            moveItem(header, "0", "28p0", height, "0", false)
            headerSwitch = true
            return true
        }
    } catch(e) {
        alert("揀錯標題")
    }
    return false
}



function getHeader(v) {
    var result
    const header = ["無", "委任令", "公告", "招標"];
    if(v === "E") {
        header = ["無", "Appointments", "Notices", "Tenders"];
    }

    var dialog = new Window('dialog', "", "x:0, y:0, width:330, height:75");
    dialog.panel = dialog.add('panel', [10, 10, 225, 65], "");

    dialog.panel.add('statictext',  [12, 18, 80, 198], "請選擇標題"); 
    var dialogPanel = dialog.kindType = dialog.panel.add('dropdownlist', [85, 13, 205, 38], header);
    dialogPanel.selection = 0;    

    dialog.okButton = dialog.add('button', [240, 10, 315, 35], "OK");
    dialog.okButton.onClick = function() {
        dialog.close(1)
        result = dialogPanel.selection;
    }
    dialog.cancelButton = dialog.add('button', [240, 40, 315, 65], "Cancel");
    dialog.cancelButton.onClick = function () {
        dialog.close(0)
        gnNumber = 0
        result = false
    }
    dialog.center(); 
    dialog.show()
    return result
}


function headerValidator(){
  kind(dialog.kindType, placementINFO.pgCount, "條例編號");
}
function onOKclicked(){
  dialog.close(1);
}
function onCANclicked(){
  dialog.close(0);
}



function openFile() {
    try {
        app.open (File('//EXZIP18/Gazette_N/-Main Gazette--' + versionFullName + '/' + version + issue + '/' + orderList[x] + '-' + version + issue + '.indd'));
        return app.activeDocument
    } catch(e) {
        try {
            app.open (File('//EXZIP18/Gazette_N/-Main Gazette--' + versionFullName + '/' + version + issue + '/' + orderList[x] + '-' + version + issue + "_col" + '.indd'));
            return app.activeDocument
        } catch(e) {
            try {
                app.open (File('//EXZIP18/Gazette_N/-Main Gazette--' + versionFullName + '/' + version + issue + '/' + orderList[x] + '-' + version + issue + "_Col" + '.indd'));
                return app.activeDocument
            } catch(e) {
                try {
                    app.open (File('//EXZIP18/Gazette_N/-Main Gazette--' + versionFullName + '/' + version + issue + '/' + orderList[x] + '-' + version + issue + "_COL" + '.indd'));
                    return app.activeDocument
                } catch(e) {
                    try {
                        app.open (File('//EXZIP18/Gazette_N/-Main Gazette--' + versionFullName + '/' + version + issue + '/' + orderList[x] + '-' + version + issue + "-col" +'.indd'));
                        return app.activeDocument
                    } catch(e) {
                        try {
                            app.open (File('//EXZIP18/Gazette_N/-Main Gazette--' + versionFullName + '/' + version + issue + '/' + orderList[x] + '-' + version + issue + "-Col" + '.indd'));
                            return app.activeDocument
                        } catch(e) {
                            try {
                                app.open (File('//EXZIP18/Gazette_N/-Main Gazette--' + versionFullName + '/' + version + issue + '/' + orderList[x] + '-' + version + issue + "-COL" + '.indd'));
                                return app.activeDocument
                            } catch(e) {
                                alert(e)
                                return false
                             }
                         }
                     }
                 }
             }
         }
     }
} 


function isSingleTextFrame(document) {
    for(var y = 0; y < document.pages.length; y ++) {
        if(document.pages[y].pageItems.length > 1) {
            alert(orderList[x] + '-' + version + issue + '.indd' + "沒有使用串連字字框，請手動置入。")
            alert(orderList[x] + '-' + version + issue + '.indd' + "沒有使用串連字字框，請手動置入。")
            alert(orderList[x] + '-' + version + issue + '.indd' + "沒有使用串連字字框，請手動置入。")
            document.close(SaveOptions.no)
            x = orderList.length
            return false
        }
    }    
    for(var y = 0; y < document.pages[0].pageItems.length; y ++) {   
        if(document.pages[0].pageItems[y].parentStory.textContainers.length < document.pages.length
         && document.pages[0].pageItems[y].geometricBounds[1] === 0) {
            alert(orderList[x] + '-' + version + issue + '.indd' + "沒有使用串連字字框，請手動置入。")
            alert(orderList[x] + '-' + version + issue + '.indd' + "沒有使用串連字字框，請手動置入。")
            alert(orderList[x] + '-' + version + issue + '.indd' + "沒有使用串連字字框，請手動置入。")
            document.close(SaveOptions.no)
            x = orderList.length
            return false
        }
    }
    return true
}



function deleteLable(document) {
    for(var x = 0; x < document.pages.length; x ++) {
        for(var y = document.pages[x].pageItems.length - 1; y >= 0; y --) {
            if(document.pages[x].pageItems[y].geometricBounds[0] > "665.717"
             || document.pages[x].pageItems[y].geometricBounds[1] > "28"
             || document.pages[x].pageItems[y].geometricBounds[2] < "0"
             || document.pages[x].pageItems[y].geometricBounds[3] < "0"
            ) {
                document.pages[x].pageItems[y].remove()
            }
        }
    }
}


function ungroup(document) {
    for(var x = 0; x < document.pages.length; x ++) {
        for(var y = 0; y < document.pages[x].groups.length; y ++) {
            document.pages[x].groups[y].ungroup()
        }
    }
}


function deletePagesFromePageTwo(document) {
    if(document.pages.length === document.pages[0].textFrames[0].parentStory.textContainers.length) {
        for(var x = document.pages.length - 1; x > 0; x --){
            document.pages[x].remove();
        }
    }
}


function applyPictureParagraphStyle(document) {
    for(var x = 0; x < document.pages.length; x ++) {
        for(var y = 0; y < document.pages[x].allGraphics.length; y ++) {
            if(document.pages[x].allGraphics[y].parent.parent.appliedParagraphStyle.name !== "Picture"
             && document.pages[x].allGraphics[y].parent.geometricBounds[1] === 0
             && Number(document.pages[x].allGraphics[y].parent.geometricBounds[3]) - Number(document.pages[x].allGraphics[y].parent.geometricBounds[1]) > 26) {
                document.pages[x].allGraphics[y].parent.parent.appliedParagraphStyle = "Picture"
            }
        }
    }
}


function modifyGN() {
    var activeGN = mainDocument.pages[getNewestPage()].textFrames[0]
    if(497 - currentTop <= mainDocument.pages[getNewestPage()].textFrames[0].geometricBounds[2] - mainDocument.pages[getNewestPage()].textFrames[0].geometricBounds[0] + spacing
    && 497 - currentTop >= mainDocument.pages[getNewestPage()].textFrames[0].geometricBounds[2] - mainDocument.pages[getNewestPage()].textFrames[0].geometricBounds[0] + 21) {
        var TempTop = 497 - mainDocument.pages[getNewestPage()].textFrames[0].geometricBounds[2] - mainDocument.pages[getNewestPage()].textFrames[0].geometricBounds[0]
        moveItem(activeGN, TempTop, "28p0", "497", "0", true)
    }
    if(activeGN.geometricBounds[2] - activeGN.geometricBounds[0] >= 495 - currentTop
     || mainDocument.pages[getNewestPage()].textFrames[0].geometricBounds[2] > 497 - spacing) {
        moveItem(activeGN, currentTop, "28p0", "497", "0", false)
    } else {
        moveItem(activeGN, currentTop, "28p0", "497", "0", true)
    }
    while(mainDocument.pages[getNewestPage()].textFrames[0].overflows) {
        var previousThreadingTextFrame = mainDocument.pages[getNewestPage()].textFrames[0]
        mainDocument.pages.add()
        currentTop = "0"
        var nextThreadingTextFrame = mainDocument.pages[getNewestPage()].textFrames.add()
        previousThreadingTextFrame.nextTextFrame = nextThreadingTextFrame
        moveItem(nextThreadingTextFrame, currentTop, "28p0", "497", "0", false)
        currentTop = getNewestTop()
        if(mainDocument.pages[getNewestPage() - 1].textFrames[0].geometricBounds[2] - mainDocument.pages[getNewestPage() - 1].textFrames[0].geometricBounds[0] < 28.661) {
            mainDocument.pages[getNewestPage() - 1].textFrames[0].remove()
            mainDocument.pages[getNewestPage()].textFrames[0].fit(FitOptions.frameToContent);
        } 
    }
    mainDocument.pages[getNewestPage()].textFrames[0].fit(FitOptions.frameToContent);
    if(mainDocument.pages[getNewestPage()].textFrames[0].geometricBounds[2] - mainDocument.pages[getNewestPage()].textFrames[0].geometricBounds[0] < 28.661) {
        if(mainDocument.pages[getNewestPage() - 1].textFrames[0].geometricBounds[0] < 2) {
            var rearrangingGN = mainDocument.pages[getNewestPage()].textFrames[0]
            moveItem(mainDocument.pages[getNewestPage() - 1].textFrames[0], "0", "28p0", "480", "0", true)
            moveItem(rearrangingGN, "0", "28p0", "495", "0", true)
        } else {
            mainDocument.pages[getNewestPage()].remove()
            mainDocument.pages[getNewestPage()].textFrames[0].fit(FitOptions.frameToContent)
            var rearrangingGN = mainDocument.pages[getNewestPage()].textFrames[0]
            moveItem(rearrangingGN, rearrangingGN.geometricBounds[0] - (rearrangingGN.geometricBounds[2] - 495), "28p0", "495", "0", false)
        }
    }
    headerSwitch = false
}    


function addNewPage() {
    currentTop = getNewestTop()
    if(currentTop > 451.5) {
        mainDocument.pages.add()
        currentTop = "0"
    } 
}

function getNewestPage() {
    return mainDocument.pages.length - 1
}


function moveItem(item, top, left, bottom, right, dofit) {
    item.geometricBounds = [top, left, bottom, right];
    if(dofit) {
        item.fit(FitOptions.frameToContent);
    }
}


function deleteBlankPage(page) {
    if(page.allGraphics.length === 0 && page.textFrames.length === 0 && page.groups.length === 0){
        page.remove()
    } else if(page.textFrames.length === 1 && page.textFrames[0].startTextFrame.contents === "") {
        page.remove()
    }
}


/////////////////////////////////////////////////////////////////////////////////////////////////////    KEEP    /////////////////////////////////////////////////////////////////////////////////////////////////////

/*
for(var x = 0; x < mainDocument.pages.length; x ++) {
    for(var y = 0; y < mainDocument.pages[x].textFrames.length; y ++) {
        if(mainDocument.pages[x].textFrames[y].geometricBounds[0] > 470
        || mainDocument.pages[x].textFrames[y].geometricBounds[1] > 28
        || mainDocument.pages[x].textFrames[y].geometricBounds[2] < 0
        || mainDocument.pages[x].textFrames[y].geometricBounds[3] < 0) {
            mainDocument.pages[x].textFrames[y].remove()
        }        
    }
    if(mainDocument.pages.length > 1) {
        deleteBlankPage(mainDocument.pages[x])
    }
}
*/

/*
    app.findGrepPreferences.appliedParagraphStyle = "";
    app.findGrepPreferences.findWhat = "(?<=\\r)(第\\d+號公告\\r第)";
    myFind = activeDocument.findGrep();
    if(myFind.length > 0) {
        alert(orderList[x] + '-' + version + issue + '.indd' + "沒有使用串連字字框，請手動置入。")
        alert(orderList[x] + '-' + version + issue + '.indd' + "沒有使用串連字字框，請手動置入。")
        alert(orderList[x] + '-' + version + issue + '.indd' + "沒有使用串連字字框，請手動置入。")
        document.close(SaveOptions.no)
        x = orderList.length
        return false
    } else {
        app.findGrepPreferences.findWhat = "(?<=\\r)(G.N. \\d+\\rG.N.)"; 
        myFind = activeDocument.findGrep();
        if(myFind.length > 0) {
            alert(orderList[x] + '-' + version + issue + '.indd' + "沒有使用串連字字框，請手動置入。")
            alert(orderList[x] + '-' + version + issue + '.indd' + "沒有使用串連字字框，請手動置入。")
            alert(orderList[x] + '-' + version + issue + '.indd' + "沒有使用串連字字框，請手動置入。")
            document.close(SaveOptions.no)
            x = orderList.length
            return false
        }
    }
*/
/*
    for(var y = 0; y < page.textFrames.length; y ++) {
        if(page.textFrames[y].contents === "詳校") {
            page.textFrames[y].remove()
        }
    }
*/