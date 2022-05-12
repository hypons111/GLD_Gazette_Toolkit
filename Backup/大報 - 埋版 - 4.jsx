///////////    Version 4    ///////////

/////////////////////////////////////////////////////////////////////////////////////////////////////    LOG    /////////////////////////////////////////////////////////////////////////////////////////////////////

///////////    Version 3    ///////////    2021-6-7    ///////////
//added openIndd(), preparationBeforeCopy(), copy(), close(), paste(), addNewPage(), moveItem()

///////////    Version 4    ///////////    2021-6-8    ///////////
//redesigned the logic of the process

/////////////////////////////////////////////////////////////////////////////////////////////////////    LOG    /////////////////////////////////////////////////////////////////////////////////////////////////////
const mainDocument = app.activeDocument
const issue = mainDocument.name.toString()[3] + mainDocument.name.toString()[4]
const orderList = []
const orderListTextFile = File('~/Desktop/list.txt') || File('~/Desktop/List.txt')
var currentTop = "0"


mainDocument.documentPreferences.allowPageShuffle = true
mainDocument.spreads.everyItem().allowPageShuffle = true
mainDocument.documentPreferences.facingPages = false

getOrderList()

modifyNumber(orderList)


for (var x = 0; x < orderList.length; x ++) {
    openIndd('//EXZIP18/Gazette_N/-Main Gazette--Chinese/TestingC' + issue + '/' + orderList[x] + '-C' + issue + '.indd')
    deleteLable(app.activeDocument)
    deletePagesFromePageTwo(app.activeDocument)
    if(app.activeDocument.name.toString()[0] + app.activeDocument.name.toString()[1] != "MA") {
    app.selection = app.activeDocument.pages[0].pageItems
    app.copy()
    app.activeDocument.close(SaveOptions.no)
    }
    app.paste()
    var activeItem = mainDocument.pages[getActivePage()].textFrames[0]
    
    moveItem(activeItem, currentTop, "28p0", "496", "0")

    currentTop = Number(activeItem.geometricBounds[2]) + 30
   
    if(Number(activeItem.geometricBounds[2]) > 495) {
        activeItem.geometricBounds = [activeItem.geometricBounds[0], activeItem.geometricBounds[1], "495", activeItem.geometricBounds[3]]
        var nextThreadingTextFrame = mainDocument.pages[getActivePage()].textFrames.add()
        activeItem.nextTextFrame = nextThreadingTextFrame
        mainDocument.pages.add()
        currentTop = "0"
        nextThreadingTextFrame.move(mainDocument.pages[getActivePage()])
        moveItem(nextThreadingTextFrame, currentTop, "28p0", "496", "0")
        currentTop = Number(nextThreadingTextFrame.geometricBounds[2]) + 30
    }


    while(Number(mainDocument.pages[getActivePage()].textFrames[0].geometricBounds[2]) > 495) {
        var previousThreadingTextFrame = mainDocument.pages[getActivePage()].textFrames[0]
        var nextThreadingTextFrame = mainDocument.pages[getActivePage()].textFrames.add()

alert(previousThreadingTextFrame.contents)
       /* 
        previousThreadingTextFrame.nextTextFrame = nextThreadingTextFrame
        nextThreadingTextFrame.geometricBounds = ["0", "28p0", "496", "0"];  
        mainDocument.pages.add()
        nextThreadingTextFrame.move(mainDocument.pages[getActivePage()])
        */
    }


    if(currentTop > 500) {
        mainDocument.pages.add()
        currentTop = "0"
    } 
}


mainDocument.documentPreferences.facingPages = true

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


function openIndd(path) {
    try{
        app.open (File(path));
    } catch(e) {}
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


function getActivePage() {
    return mainDocument.pages.length - 1
}


function moveItem(item, top, left, bottom, right) {
    item.geometricBounds = [top, left, bottom, right];    //["上", "左", "下", "右",]
    if(!item.overflows) {
        try {
            item.fit(FitOptions.frameToContent);
        } catch(e) {}
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
