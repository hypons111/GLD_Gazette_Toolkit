const activeDocument = app.activeDocument
const activeDocumentTextFrame = activeDocument.pages[0].textFrames[0]
const path = activeDocument.filePath
const issue = activeDocument.name.toString()[3] + activeDocument.name.toString()[4]
const orderList = []
const orderListFile = File('~/Desktop/orderLog.txt')



if(activeDocumentTextFrame instanceof TextFrame) {
    app.selection = activeDocumentTextFrame.texts.item(0).texts
}


getOrderList()

for (var i = 0; i < orderList.length; i ++) {
   app.open (File('~/Desktop/新增資料夾/'+ orderList[i] + '-C' + issue + '.indd'));
   copy()
   paste()
}




/////////////////////////////////////////////  FUNCTIONS /////////////////////////////////////////////

function getOrderList() {
    orderListFile.open("r");
    var tempNumber = 0
    for (var i = 0; i < orderListFile.length; i ++) {
        tempNumber = Number(orderListFile.readln())
        if(tempNumber > 0) {
            orderList.push(modifyFileName(tempNumber))
        }
    } 
}

function modifyFileName(number) {
    if (number > 0) {
        if (number < 10 ) {
            return "00" +number
        } else if (name < 100 ) {
            return "0" + number
        }else {
            return number
        }
    }
}

function copy() {
    while (gnFile.pages.length > 1) {
        gnFile.pages[1].textFrames[0].remove()
        gnFile.pages[1].remove()
    }
    var selectedContents = gnFile.pages[0].textFrames[0].texts.item(0).parentStory.texts[0].select()
    app.copy()
    gnFile.close();
}

function paste() {
    if(activeDocumentTextFrame instanceof TextFrame) {
        app.paste()
        spacing()
        app.paste()
    }
}

function spacing() {
    var spaceFile = app.documents.add();
    space = document.pages.item(0).textFrames.add();    
    space.contents = "\r\r";
    spaceFile.pages[0].textFrames[0].texts.item(0).parentStory.texts[0].select()
    app.copy()    
    spaceFile.close(SaveOptions.NO);
}


