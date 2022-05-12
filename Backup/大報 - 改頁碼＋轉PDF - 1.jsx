var orderList = []
var activeFile
var thisIssue = ''
var firstPageNumber = 0

pageList = File('~/Desktop/頁碼表.txt');
if(pageList.exists){
    readPrefs();
    getThisIssue()
    getFirstPageNumber()
    changeNameAndSaveAs()
} else {
    alert("no file")
}


//  讀取頁碼表.txt
function readPrefs(){
    pageList.open("r");
    for (var i = 0; i < pageList.length / 5; i ++) {
        var order = [] 
        order[i] = Number(pageList.readln())
        if(order[i] > 0) {
            orderList.push(order[i])
        }
    } 
    pageList.close();
}

//  回傳今期期數
function getThisIssue() {
    var issueList = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '30', '31', '32', '33', '34', '35', '36', '37', '38', '39', '40', '41', '42', '43', '44', '45', '46', '47', '48', '49', '50', '51', '52', '53']
    var guess = 'MA-' + issueList[0] + '-C01'
    for(var i = 0; i < issueList.length; i ++) { 
        try {
            var fileName = app.open (File('~/Desktop/新增資料夾/MA-' + issueList[i] + '-C01.indd'));
            if(fileName) {
                thisIssue = issueList[i]
                fileName.close();
                break
            }
        } catch (error) {}
    }
}

function getFirstPageNumber() {
    firstPageNumber = orderList[0]
}

//  更改.indd 檔名
function changeNameAndSaveAs() {
    for(var i = 0; i < orderList.length; i ++) {
        try {
            activeFile = app.open (File('~/Desktop/新增資料夾/MA-' + thisIssue + '-C0' + (i + 1) + '.indd'));
        } catch (error) {
            activeFile = app.open (File('~/Desktop/新增資料夾/MA-' + thisIssue + '-C' + (i + 1) + '.indd'));
        }
        if(activeFile) {
            changePageNumber()
            activeFile.save(File('~/Desktop/Here/MA-' + thisIssue + '-' + orderList[i] + '.indd'));
            activeFile.close();
        } 
    }
}

//  更改.indd 頁碼
function changePageNumber() {
    app.activeDocument.documentPreferences.allowPageShuffle = true
    app.activeDocument.sections[0].continueNumbering = false
    app.activeDocument.sections[0].pageNumberStart = firstPageNumber
}


