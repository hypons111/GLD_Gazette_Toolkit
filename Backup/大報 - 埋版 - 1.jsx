var oldInteractionPref = app.scriptPreferences.userInteractionLevel;
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
var usePrefs = true;
var orderList = []
var gnFileName
var gnFile



// Look for and read prefs file
gnOrder = File((Folder(app.activeScript)).parent + "/gnOrderList.txt");
if(gnOrder.exists){
    readPrefs();
} else {
    alert("no file")
}


// function to read prefs from a file
function readPrefs(){
    if(usePrefs){
        gnOrder.open("r");
        for (var i = 0; i < gnOrder.length / 5; i ++) {
            var order = [] 
            order[i] = Number(gnOrder.readln());
            orderList.push(order[i])
        } 
        gnOrder.close();
    }
}

alert(orderList.length)

var outPutFile = app.open (File('~/Desktop/Main.indd'));
var outPutFileTextFrame = outPutFile.pages[0].textFrames[0]
if(outPutFileTextFrame instanceof TextFrame) {
    app.selection = outPutFileTextFrame.texts.item(0).texts
}

for (var i = 0; i < orderList.length; i ++) {
    if (orderList[i] > 0) {
        if (orderList[i] < 10 ) {
            gnFileName = "00" + orderList[i]
        } else if (orderList[i] < 100 ) {
            gnFileName = "0" + orderList[i]
        }else {
            gnFileName = orderList[i]
        }
    }
   gnFile = app.open (File('~/Desktop/新增資料夾/'+ gnFileName + '-C11.indd'));
   alert(gnFileName)
    copy()
    pasting()
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

function spacing() {
    var spaceFile = app.documents.add();
    space = document.pages.item(0).textFrames.add();    
    space.contents = "\r\r\r\r\r";
    spaceFile.pages[0].textFrames[0].texts.item(0).parentStory.texts[0].select()
    app.copy()    
    spaceFile.close();
}

function pasting() {
    if(outPutFileTextFrame instanceof TextFrame) {
        app.paste()
        spacing()
        app.paste()
    }
}