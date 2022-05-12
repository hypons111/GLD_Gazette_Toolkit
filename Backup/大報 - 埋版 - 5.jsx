///////////    Version 5    ///////////

/////////////////////////////////////////////////////////////////////////////////////////////////////    LOG    /////////////////////////////////////////////////////////////////////////////////////////////////////

///////////    Version 3    ///////////    2021-6-7    ///////////
//added openIndd(), preparationBeforeCopy(), copy(), close(), paste(), addNewPage(), moveItem()

///////////    Version 4    ///////////    2021-6-8    ///////////
//redesigned the logic of the process

///////////    Version 5    ///////////    2021-6-9    ///////////
//redesigned the logic of the process

/////////////////////////////////////////////////////////////////////////////////////////////////////    LOG    /////////////////////////////////////////////////////////////////////////////////////////////////////
const mainDocument = app.activeDocument
const issue = mainDocument.name.toString()[3] + mainDocument.name.toString()[4]
const version = mainDocument.name.toString()[6]
const orderList = []
const orderListTextFile = File('~/Desktop/list.txt') || File('~/Desktop/List.txt')
var currentTop = "0"
const spacing = 40

mainDocument.documentPreferences.allowPageShuffle = true
mainDocument.spreads.everyItem().allowPageShuffle = true
mainDocument.documentPreferences.facingPages = false

getOrderList()
modifyNumber(orderList)

for (var x = 0; x < orderList.length; x ++) {
/////////// get ///////////
    app.open (File('//EXZIP18/Gazette_N/-Main Gazette--Chinese/Testing' + version + issue + '/' + orderList[x] + '-' + version + issue + '.indd'));
    deleteLable(app.activeDocument)
    ungroup(app.activeDocument)
    if(app.activeDocument.pages.length > app.activeDocument.pages[0].textFrames[0].parentStory.textContainers.length
    || app.activeDocument.pages.length < app.activeDocument.pages[0].textFrames.length) {
        alert(orderList[x] + '-' + version + issue + '.indd' + "沒有使用串連字字框，請手動置入。")
        break
    } else {
        deletePagesFromePageTwo(app.activeDocument)
        if(app.activeDocument.name.toString()[0] + app.activeDocument.name.toString()[1] != "MA") {
            app.selection = app.activeDocument.pages[0].pageItems
            app.copy()
            app.activeDocument.close(SaveOptions.no)
        }
/////////// get ///////////
        
/////////// paste ///////////
        app.paste()
/////////// paste ///////////
        
/////////// modify ///////////
        var activeItem = mainDocument.pages[newPage()].textFrames[0]
        
        moveItem(activeItem, currentTop, "28p0", "496", "0", true)
        
        if(activeItem.geometricBounds[2] > 495 && activeItem.geometricBounds[2] < 495 + 31.417) {
            moveItem(activeItem, currentTop - (activeItem.geometricBounds[2] - 495), "28p0", activeItem.geometricBounds[2] - (activeItem.geometricBounds[2] - 495), "0", true)
        }
        
        currentTop = Number(activeItem.geometricBounds[2]) + spacing

        while(mainDocument.pages[newPage()].textFrames[0].overflows) {
            var previousThreadingTextFrame = mainDocument.pages[newPage()].textFrames[0]
            mainDocument.pages.add()
            currentTop = "0"
            var nextThreadingTextFrame = mainDocument.pages[newPage()].textFrames.add()
            previousThreadingTextFrame.nextTextFrame = nextThreadingTextFrame
            moveItem(nextThreadingTextFrame, currentTop, "28p0", "496", "0", false)
            currentTop = Number(nextThreadingTextFrame.geometricBounds[2]) + spacing
        }

        mainDocument.pages[newPage()].textFrames[0].fit(FitOptions.frameToContent);

        if(currentTop > 495) {
            mainDocument.pages.add()
            currentTop = "0"
        } 
    }
}


for(var x = 0; x < mainDocument.pages.length; x ++) {
    for(var y = 0; y < mainDocument.pages[x].textFrames.length; y ++) {
        if(mainDocument.pages[x].textFrames[y].contents === "" && !mainDocument.pages[x].textFrames[y].overflows) {
            mainDocument.pages[x].textFrames[y].remove()
        }
    }
    deleteBlankPage(mainDocument.pages[x])
}

//mainDocument.documentPreferences.facingPages = true

/////////////////////////////////////////////////////////////////////////////////////////////////////    FUNCTIONS    /////////////////////////////////////////////////////////////////////////////////////////////////////

function getOrderList() {
    orderListTextFile.open("r")
    const temp = []
    for (var x = 0; x < orderListTextFile.length; x ++) {
        temp.push(Number(orderListTextFile.readln()))
    } 
    for (var y = 0; y < temp.length; y ++) {
        if(Number(temp[y]) !== 0) {
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


function ungroup(document) {
    for(var x = 0; x < document.pages.length; x ++) {
        for(var y = 0; y < document.pages[x].groups.length; y ++) {
            document.pages[x].groups[y].ungroup()
        }
    }
}


function deleteLable(activeDocument) {
    for(var x = 0; x < activeDocument.pages.length; x ++) {
        for(var y = 0; y < activeDocument.pages[x].textFrames.length; y ++) {
            if(activeDocument.pages[x].textFrames[y].geometricBounds[0] == "78"
            && activeDocument.pages[x].textFrames[y].geometricBounds[1] == "34.5"
            && activeDocument.pages[x].textFrames[y].geometricBounds[2] == "135"
            && activeDocument.pages[x].textFrames[y].geometricBounds[3] == "40.5") {
                activeDocument.pages[x].textFrames[y].remove()
            }
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


function newPage() {
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
        if(Number(activeItem.geometricBounds[2]) > 495 || activeItem.overflows) {
            activeItem.geometricBounds = [activeItem.geometricBounds[0], activeItem.geometricBounds[1], "495", activeItem.geometricBounds[3]]
            var nextThreadingTextFrame = mainDocument.pages[newPage()].textFrames.add()
            activeItem.nextTextFrame = nextThreadingTextFrame
            mainDocument.pages.add()
            currentTop = "0"
            nextThreadingTextFrame.move(mainDocument.pages[newPage()])
            moveItem(nextThreadingTextFrame, currentTop, "28p0", "496", "0")
            currentTop = Number(nextThreadingTextFrame.geometricBounds[2]) + spacing
        }
*/