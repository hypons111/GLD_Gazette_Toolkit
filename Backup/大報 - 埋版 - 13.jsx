///////////    Version 13    ///////////

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

/////////////////////////////////////////////////////////////////////////////////////////////////////    LOG    /////////////////////////////////////////////////////////////////////////////////////////////////////
const mainDocument = app.activeDocument
var activeDocument
const issue = mainDocument.name.toString()[3] + mainDocument.name.toString()[4]
const version = mainDocument.name.toString()[6]
const versionFullName = getVersionFullName(version)
const spacing = 45
var currentTop = getNewestTop()
const orderList = []
const orderListTextFile = File('~/Desktop/list.txt') || File('~/Desktop/List.txt') || File('~/Desktop/LIST.txt')
var abort = false

app.scriptPreferences.userInteractionLevel = UserInteractionLevels.neverInteract;
mainDocument.documentPreferences.allowPageShuffle = true
mainDocument.spreads.everyItem().allowPageShuffle = true
mainDocument.documentPreferences.facingPages = false


if(isAuto()) {
    for (var x = 0; x < orderList.length; x ++) {
        openFile()
        activeDocument = app.activeDocument
        ungroup(activeDocument)
        deleteLable(activeDocument.pages[0])
        if(isThreadTextFrame(activeDocument) && activeDocument.name.toString()[0] + activeDocument.name.toString()[1] != "MA") {
            deletePagesFromePageTwo(activeDocument)
            app.selection = activeDocument.pages[0].pageItems
            app.copy()
            activeDocument.close(SaveOptions.no)
            mainDocument.pages.item(-1).select()
            mainDocument.select(NothingEnum.NOTHING)
            app.paste()
            modifyGN()
        }
    }
} else {
    var isContinue = true
    while(isContinue) {        
        getGNNumber(inputGNNumber())
        try{
            modifyNumber(orderList)
            for (var x = 0; x < orderList.length; x ++) {
                openFile()
                activeDocument = app.activeDocument
                ungroup(activeDocument)
                deleteLable(activeDocument.pages[0])
                if(isThreadTextFrame(activeDocument) && activeDocument.name.toString()[0] + activeDocument.name.toString()[1] != "MA") {
                    deletePagesFromePageTwo(activeDocument)
                    app.selection = activeDocument.pages[0].pageItems
                    app.copy()
                    activeDocument.close(SaveOptions.no)
                    mainDocument.pages.item(-1).select()
                    mainDocument.select(NothingEnum.NOTHING)
                    app.paste()
                    modifyGN()
                }
            }
        } catch(e) {}
    }
}

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

mainDocument.documentPreferences.facingPages = true

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
        return mainDocument.pages[getNewestPage()].textFrames[0].geometricBounds[2] + spacing
    } catch(e) {
            return 0
    } 
}


function isAuto() {
    getOrderList()
    modifyNumber(orderList)
    return orderList.length
}


function inputGNNumber() {
    var gnNumber
    dialog = new Window('dialog', "", "x:0, y:0, width:290, height:75");
    dialog.panel = dialog.add('panel', [10, 10, 185, 65], "");

    dialog.panel.add('statictext',  [10, 16, 80, 196], "Number"); 
    dialog.gN = dialog.panel.add('edittext', [80, 11, 155, 36], "");
    dialog.gN.active = true
    dialog.gN.onChange = function() {
        gnNumber = dialog.gN.text
    }

    dialog.okButton = dialog.add('button', [200, 10, 275, 35], "OK");
    dialog.okButton.onClick = function() { dialog.close(1) }
    dialog.cancelButton = dialog.add('button', [200, 40, 275, 65], "Cancel");
    dialog.cancelButton.onClick = function () {
        isContinue = false
        dialog.close(0)
        gnNumber = 0
    }
    dialog.center(); 
    dialog.show()
    return gnNumber;
}


function getOrderList() {
    orderListTextFile.open("r")
    const temp = []
    for (var x = 0; x < orderListTextFile.length; x ++) {
        temp.push(orderListTextFile.readln())
    } 
    for (var y = 0; y < temp.length; y ++) {
        if(temp[y].slice(1, 2) === "-" && Number(temp[y]) !== 0) {
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


function getGNNumber(number) {
    orderList.splice(0, orderList.length)
    try{
        if(number.slice(1, 2) === "-" && Number(number) !== 0) {
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
    } catch(e) {}
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


function openFile() {
        app.open (File('//EXZIP18/Gazette_N/-Main Gazette--' + versionFullName + '/Testing' + version + issue + '/' + orderList[x] + '-' + version + issue + '.indd'));
} 


function isThreadTextFrame() {
    if(activeDocument.pages.length > activeDocument.pages[0].textFrames[0].parentStory.textContainers.length
    || activeDocument.pages.length < activeDocument.pages[0].textFrames.length) {
        alert(orderList[x] + '-' + version + issue + '.indd' + "沒有使用串連字字框，請手動置入。")
        activeDocument.close(SaveOptions.no)
        isContinue = false
        x = orderList.length
        return false
    }
    return true
}


function deleteLable(page) {
    for(var y = 0; y < page.pageItems.length; y ++) {
        if(page.pageItems[y].geometricBounds[2] < "0"
        || page.pageItems[y].geometricBounds[3] < "0"
        || page.pageItems[y].geometricBounds[0] > "665.717"
        || page.pageItems[y].geometricBounds[1] > "28"
        ) {
            page.pageItems[y].remove()
        }
    }
    for(var y = 0; y < page.textFrames.length; y ++) {
        if(page.textFrames[y].contents === "詳校") {
            page.textFrames[y].remove()
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


function deletePagesFromePageTwo(activeDocument) {
    if(activeDocument.pages.length === activeDocument.pages[0].textFrames[0].parentStory.textContainers.length) {
        for(var x = activeDocument.pages.length - 1; x > 0; x --){
            activeDocument.pages[x].remove();
        }
    }
}


function modifyGN() {
    var activeGN = mainDocument.pages[getNewestPage()].textFrames[0]

    if(activeGN.geometricBounds[2] - activeGN.geometricBounds[0] >= 495 - currentTop) {
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

