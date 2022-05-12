///////////    Version 7    ///////////

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

/////////////////////////////////////////////////////////////////////////////////////////////////////    LOG    /////////////////////////////////////////////////////////////////////////////////////////////////////
const mainDocument = app.activeDocument
var activeDocument
const issue = mainDocument.name.toString()[3] + mainDocument.name.toString()[4]
const version = mainDocument.name.toString()[6]
const orderList = []
const orderListTextFile = File('~/Desktop/list.txt') || File('~/Desktop/List.txt')
var currentTop = "0"
const spacing = 45

mainDocument.documentPreferences.allowPageShuffle = true
mainDocument.spreads.everyItem().allowPageShuffle = true
mainDocument.documentPreferences.facingPages = false

getOrderList()
modifyNumber(orderList)

for (var x = 0; x < orderList.length; x ++) {
/////////// get ///////////
    app.open (File('//192.200.9.12/Gazette_N/-Main Gazette--Chinese/TestingC23/' + orderList[x] + '-' + version + issue + '.indd'));
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.neverInteract;
    activeDocument = app.activeDocument
    ungroup(activeDocument)
    deleteLable(activeDocument.pages[0])
    if(activeDocument.pages.length > activeDocument.pages[0].textFrames[0].parentStory.textContainers.length
    || activeDocument.pages.length < activeDocument.pages[0].textFrames.length) {
        alert(orderList[x] + '-' + version + issue + '.indd' + "沒有使用串連字字框，請手動置入。")
        break
    } else {
        deletePagesFromePageTwo(activeDocument)
        if(activeDocument.name.toString()[0] + activeDocument.name.toString()[1] != "MA") {
            app.selection = activeDocument.pages[0].pageItems
            app.copy()
            activeDocument.close(SaveOptions.no)
        }
/////////// get ///////////
        
/////////// paste ///////////
        app.paste()
/////////// paste ///////////
        
/////////// modify ///////////
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

        if(mainDocument.pages[getNewestPage()].textFrames[0].geometricBounds[2] - mainDocument.pages[getNewestPage()].textFrames[0].geometricBounds[0] < 28.661){
            mainDocument.pages[getNewestPage()].remove()
            mainDocument.pages[getNewestPage()].textFrames[0].fit(FitOptions.frameToContent)
            var rearrangingGN = mainDocument.pages[getNewestPage()].textFrames[0]
            moveItem(rearrangingGN, rearrangingGN.geometricBounds[0] - (rearrangingGN.geometricBounds[2] - 495), "28p0", "495", "0", false)
        }



        currentTop = getNewestTop()

        if(currentTop > 451.5) {
            mainDocument.pages.add()
            currentTop = "0"
        } 
    }
}


for(var x = 0; x < mainDocument.pages.length; x ++) {
    for(var y = 0; y < mainDocument.pages[x].textFrames.length; y ++) {
        if(mainDocument.pages[x].textFrames[y].geometricBounds[0] > 470
        || mainDocument.pages[x].textFrames[y].geometricBounds[1] > 28
        || mainDocument.pages[x].textFrames[y].geometricBounds[2] < 0
        || mainDocument.pages[x].textFrames[y].geometricBounds[3] < 0) {
            alert("here")
            mainDocument.pages[x].textFrames[y].remove()
        }        
        if(mainDocument.pages[x].textFrames[y].contents === "" && !mainDocument.pages[x].textFrames[y].overflows) {
//            mainDocument.pages[x].textFrames[y].remove()
        }
    }
//    deleteBlankPage(mainDocument.pages[x])
}



/*
///////////     將突出嚟既GN推返去497
        if(activeGN.geometricBounds[2] > 497 && activeGN.geometricBounds[2] < 495 + 31.417) {
            moveItem(activeGN, currentTop - (activeGN.geometricBounds[2] - 497), "28p0", activeGN.geometricBounds[2] - (activeGN.geometricBounds[2] - 497), "0", true)
        }
*/


//mainDocument.documentPreferences.facingPages = true

/////////////////////////////////////////////////////////////////////////////////////////////////////    FUNCTIONS    /////////////////////////////////////////////////////////////////////////////////////////////////////

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


function modifyNumber(number) {
    for (var x = 0; x < number.length; x ++) {
        if (number[x] < 10 ) {
            number[x] = "00" + number[x]
        } else if (number[x] < 100 ) {
            number[x] = "0" + number[x]
        }
    }
}


function deleteLable(page) {
    for(var y = 0; y < page.textFrames.length; y ++) {
        if(page.textFrames[y].geometricBounds[0] == "78"
        && page.textFrames[y].geometricBounds[1] == "34.5"
        && page.textFrames[y].geometricBounds[2] == "135"
        && page.textFrames[y].geometricBounds[3] == "40.5") {
            page.textFrames[y].remove()
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


function getNewestPage() {
    return mainDocument.pages.length - 1
}


function getNewestTop() {
    return mainDocument.pages[getNewestPage()].textFrames[0].geometricBounds[2] + spacing
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
function cutOverFlowGN(gn) {
    if(Number(gn.geometricBounds[2]) > 500) {
        gn.geometricBounds = [gn.geometricBounds[0], gn.geometricBounds[1], "500", gn.geometricBounds[3]]
        mainDocument.pages.add()
        currentPage += 1
        currentTop = "0"
    }
}

function spreadOversetedGN(gn) {
    if(gn.overflows) {
        var nextThreadingTextFrame = mainDocument.pages[currentPage - 1].textFrames.add()
        gn.nextTextFrame = nextThreadingTextFrame
        mainDocument.pages.add(LocationOptions.AFTER, mainDocument.pages[currentPage - 1])
        nextThreadingTextFrame.move(mainDocument.pages[currentPage])
        nextThreadingTextFrame.geometricBounds = ["0", "28p0", "495", "0"];
        nextThreadingTextFrame.fit(FitOptions.frameToContent);
        currentTop = Number(nextThreadingTextFrame.geometricBounds[2]) + 60
    }
}

/*   
        if(Number(activeGN.geometricBounds[2]) > 495 || activeGN.overflows) {
            activeGN.geometricBounds = [activeGN.geometricBounds[0], activeGN.geometricBounds[1], "495", activeGN.geometricBounds[3]]
            var nextThreadingTextFrame = mainDocument.pages[getNewestPage()].textFrames.add()
            activeGN.nextTextFrame = nextThreadingTextFrame
            mainDocument.pages.add()
            currentTop = "0"
            nextThreadingTextFrame.move(mainDocument.pages[getNewestPage()])
            moveItem(nextThreadingTextFrame, currentTop, "28p0", "496", "0")
            currentTop = Number(nextThreadingTextFrame.geometricBounds[2]) + spacing
        }
*/