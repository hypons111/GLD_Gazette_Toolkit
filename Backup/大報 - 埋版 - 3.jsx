///////////    Version 3    ///////////

/////////////////////////////////////////////////////////////////////////////////////////////////////    LOG    /////////////////////////////////////////////////////////////////////////////////////////////////////

///////////    Version 3    ///////////    2021-6-7    ///////////
//added openIndd(), preparationBeforeCopy(), copy(), close(), paste(), addNewPage(), moveItem()

/////////////////////////////////////////////////////////////////////////////////////////////////////    LOG    /////////////////////////////////////////////////////////////////////////////////////////////////////
const mainDocument = app.activeDocument
const issue = mainDocument.name.toString()[3] + mainDocument.name.toString()[4]
const orderList = []
const orderListTextFile = File('~/Desktop/list.txt') || File('~/Desktop/List.txt')
var currentUP = "0"
var currentPage = 0
getOrderList()

mainDocument.documentPreferences.allowPageShuffle = true
mainDocument.spreads.everyItem().allowPageShuffle = true
mainDocument.documentPreferences.facingPages = false

for (var x = 0; x < orderList.length - 1; x ++) {
    openIndd('//EXZIP18/Gazette_N/-Main Gazette--Chinese/TestingC' + issue + '/' + orderList[x] + '-C' + issue + '.indd')
    preparationBeforeCopy(app.activeDocument)
    copy(app.activeDocument.pages[0].pageItems)
    close(app.activeDocument)
    paste(mainDocument.pages[x])
    addNewPage(mainDocument)
    moveItem(mainDocument.pages[x].pageItems[0], "0", "28p0", "504", "0")
}

/////////////////////////////////////////////////////////////////////////////////////////////////////    FUNCTIONS    /////////////////////////////////////////////////////////////////////////////////////////////////////

function getOrderList() {
    orderListTextFile.open("r")
    for (var x = 0; x < orderListTextFile.length; x ++) {
        orderList.push(orderListTextFile.readln())
    } 
}

function openIndd(path) {
    try{
        app.open (File(path));
    } catch(e) {}
}

function preparationBeforeCopy(activeDocument) {
    if(activeDocument.pages.length === activeDocument.pages[0].textFrames[0].parentStory.textContainers.length) {
        for(var x = activeDocument.pages.length - 1; x > 0; x --){
            activeDocument.pages[x].remove();
        }
    }
}

function copy(items) {
    app.selection = items
    app.copy()
}

function close(document) {
    document.close();
}

function paste(page) {
    app.paste(page)
}

function addNewPage(document) {
    document.pages.add()
}

function moveItem(item, up, left, down, right) {
    item.geometricBounds = [up, left, down, right];    //["上", "左", "下", "右",]
    try {
        item.fit(FitOptions.frameToContent);
    } catch(e) {}
}





/////////////////////////////////////////////////////////////////////////////////////////////////////    KEEP    /////////////////////////////////////////////////////////////////////////////////////////////////////
function cutOverFlowGN(gn) {
    if(Number(gn.geometricBounds[2]) > 500) {
        gn.geometricBounds = [gn.geometricBounds[0], gn.geometricBounds[1], "500", gn.geometricBounds[3]]
        mainDocument.pages.add()
        currentPage += 1
        currentUP = "0"
    }
}

function spreadOversetedGN(gn) {
    if(gn.overflows) {
        var nextThreadingTextFrame = mainDocument.pages[currentPage - 1].textFrames.add()
        gn.nextTextFrame = nextThreadingTextFrame
        mainDocument.pages.add(LocationOptions.AFTER, mainDocument.pages[currentPage - 1])
        nextThreadingTextFrame.move(mainDocument.pages[currentPage])
        nextThreadingTextFrame.geometricBounds = ["0", "28p0", "504", "0"];
        nextThreadingTextFrame.fit(FitOptions.frameToContent);
        currentUP = Number(nextThreadingTextFrame.geometricBounds[2]) + 60
    }
}
