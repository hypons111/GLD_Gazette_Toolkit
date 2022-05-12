///////////    Version 12    ///////////

/////////////////////////////////////////////////////////////////////////////////////////////////////    LOG    /////////////////////////////////////////////////////////////////////////////////////////////////////

///////////    Version 8    ///////////    2021-6-28    ///////////
//fixed exorted pdf has empty pages problem (testing)
//removed changeNameAndSaveAndDelete()

///////////    Version 9    ///////////    2021-7-15    ///////////
//redesign main() logic
//auto add "_col" to file name when the file is colored

///////////    Version 10    ///////////    2021-7-21    ///////////
//fixed getPageNumberFromDialog() bug which caused by I forgot to put the second parameter in.

///////////    Version 11    ///////////    2021-8-31    ///////////
//added updatePageInfoFile(), getStartCellNumber to generate the page number list automaticly

///////////    Version 12    ///////////    2021-9-1    ///////////
//getDateMonthYear() to generate the launch date (but still have bug when chinese new year holiday)

/////////////////////////////////////////////////////////////////////////////////////////////////////    SCRIPT    /////////////////////////////////////////////////////////////////////////////////////////////////////

const activeDocument = app.activeDocument
var name = activeDocument.fullName.toString().slice(49)
const path =  activeDocument.fullName.toString().slice(0, 49)
const version = path.slice(45, 46)
const issue = path.slice(46, 48)
var startPage
var endPage
const pageLength = activeDocument.pages.length
const logFile = File(activeDocument.fullName.toString().slice(0, 46) + issue + '/log.txt');
var color = ""
var state = false


main()
if(state === true) {
    updatePageInfoFile()
}


/////////////////////////////////////////////////////////////////////////////////////////////////////    FUNCTION    /////////////////////////////////////////////////////////////////////////////////////////////////////

function main() {
    var preflightProfile = app.preflightProfiles[1]
    var preflightProcess = app.preflightProcesses.add(activeDocument, preflightProfile)
    
    preflightProcess.waitForProcess()
    preflightResult = preflightProcess.processResults

    if(preflight() !== false) {
        execute()
    }
}


function preflight() {
    if(preflightResult.indexOf('None') !== 0) {
        for(var x = 0; x < preflightProcess.aggregatedResults[2].length; x ++) {
            if(preflightProcess.aggregatedResults[2][x].toString().slice(2, 4) === "連結"
             || preflightProcess.aggregatedResults[2][x].toString().slice(2, 4) === "文字"
             || preflightProcess.aggregatedResults[2][x].toString().slice(2, 4) === "文件") {
                alert('檔案有問題')
                return false
            }
        }
        for(var x = 0; x < preflightProcess.aggregatedResults[2].length; x ++) {
            if(preflightProcess.aggregatedResults[2][x].toString().slice(2, 4) === "顏色") {
                color = "_col"
            }
        }  
    }
}


function execute() {
    logFile.open("r");
    const pageNumber = Number(logFile.readln())
    if(pageNumber > 0) {
        if(activeDocument.name.match('MA-' + issue + '-C') || activeDocument.name.match('MA-' + issue + '-E')){
            updateLogFile(getPageNumberFromLog(pageNumber))
        } else { alert('開錯File') }
    } else {
        if(activeDocument.name.match('Contents-C' + issue) || activeDocument.name.match('Contents-E' + issue)) {
            updateLogFile(getPageNumberFromDialog())
        } else { alert('開錯File') }
    }
}


function makeDialog() {
    dialog = new Window('dialog', "", "x:0, y:0, width:290, height:75");
    dialog.panel = dialog.add('panel', [10, 10, 185, 65], "");

    dialog.panel.add('statictext',  [10, 16, 80, 196], "第一頁頁碼"); 
    dialog.pageNumber = dialog.panel.add('edittext', [80, 11, 155, 36], "");
    dialog.pageNumber.active = true
    dialog.pageNumber.onChange = function () {
        pageValidator(dialog.pageNumber, placementINFO.pgCount, "");
    }    
    
    dialog.OKButton = dialog.add('button', [200, 10, 275, 35], "OK");
    dialog.OKButton.onClick = function () {
        dialog.close(1)
    }    

    dialog.cancelButton = dialog.add('button', [200, 40, 275, 65], "Cancel");
    dialog.cancelButton.onClick = function () {
        dialog.close(0)
    }    

    return dialog;
}



function getPageNumberFromDialog() {
    makeDialog();
    dialog.center();
    if(dialog.show() === 1) {
        const firstPageNumber = Number(dialog.pageNumber.text)
        if(firstPageNumber > 0) {
            if(firstPageNumber % 2 === 1) {
                changePageNumber(firstPageNumber)
                exportPDF(firstPageNumber)
                rename(activeDocument.fullName, firstPageNumber)
                return pageLength + firstPageNumber
            } else {
                alert('目錄要單數版開始')
            }
        } else { alert('輸入頁碼') }
    } else { exit() }   
}


function getPageNumberFromLog(pageNumber) {
    changePageNumber(pageNumber)
    exportPDF(pageNumber)
    rename(activeDocument.fullName, pageNumber)
    return pageLength + pageNumber
}


function updateLogFile(nextPageNumber) {
    logFile.open("w");
    logFile.write(nextPageNumber)
    logFile.close();
}





function changePageNumber(pageNumber) {
    activeDocument.documentPreferences.allowPageShuffle = true
    activeDocument.spreads.everyItem().allowPageShuffle = true
    activeDocument.sections[0].continueNumbering = false
    activeDocument.sections[0].sectionPrefix = ""
    activeDocument.sections[0].pageNumberStart = pageNumber
    startPage = activeDocument.pages[0].name
    endPage = activeDocument.pages[activeDocument.pages.length - 1].name
}


function exportPDF(pageNumber) {
    const exportPreset = app.pdfExportPresets.item('Output_pageup');
    activeDocument.exportFile(ExportFormat.pdfType, File('//APSRVX/HotFolderRoot/Gaz_PDF_Hires(new)/MA-' + issue + '-' + pageNumber + color + '.pdf'), false, exportPreset)
}

 
function rename(documentPath, pageNumber) {
    activeDocument.close(SaveOptions.yes)
    File(documentPath).rename('MA-' + issue + '-' + pageNumber + color + '.indd')
    state = true
}


function updatePageInfoFile() {
    const pageInfoDocument = app.open(File('//EXZIP18/Gazette_N/-Main Gazette--Chinese/C' + issue + '/MA-' + issue + '-PageNo.indd'))
    const leftTable = pageInfoDocument.textFrames[1].tables[0]
    const rightTable = pageInfoDocument.textFrames[0].tables[0]
    const cellNumber = getStartCellNumber(version)

    if(name.toString().slice(0, 8) === "Contents" && version === "C") {
        leftTable.cells.item(0).contents = "憲報第  " + issue + "  期"
        leftTable.cells.item(1).contents = (getDateMonthYear(new Date().getMonth().toString(), new Date().getDate().toString()))
        rightTable.cells.item(2).contents = Number(startPage - 2).toString()
        rightTable.cells.item(5).contents = Number(startPage - 1).toString()
        rightTable.cells.item(8).contents = startPage + " – " + endPage
    } else if(name.toString().slice(0, 8) === "Contents" && version === "E") {
        rightTable.cells.item(14).contents = startPage + " – " + endPage
    } else {
        while(leftTable.cells.item(cellNumber).contents !== "E") {
            cellNumber += 4
        }
        leftTable.cells.item(cellNumber).contents = name.toString().slice(0, 9) + " " + startPage + " – " + endPage
        leftTable.cells.item(cellNumber + 1).contents = pageLength.toString()
        if(version === "C") {
            rightTable.cells.item(11).contents = leftTable.cells.item(6).contents.slice(leftTable.cells.item(6).contents.indexOf(" ") + 1, leftTable.cells.item(6).contents.indexOf(" – ")) + " – " + endPage
        } else if(version === "E") {
            rightTable.cells.item(17).contents = leftTable.cells.item(8).contents.slice(leftTable.cells.item(8).contents.indexOf(" ") + 1, leftTable.cells.item(8).contents.indexOf(" – ")) + " – " + endPage
        }
    }
    pageInfoDocument.close(SaveOptions.yes)
}


function getStartCellNumber(v) {
    if(v === "C") {
        return 6
    } else if(v === "E") {
        return 8
    }
}

function getDateMonthYear(month, date) {
    switch(month) {
        case "0" :
            if(date === "31") {
                return "2" + " February " + new Date().getFullYear()
            } else if(date === "30") {
                return "1" + " February " + new Date().getFullYear()
            } else {
                return Number(date) + 2 + " January " + new Date().getFullYear()
            }
        case "1" :
            if(date === "28") {
                return "2" + " March " + new Date().getFullYear()
            } else if(date === "27") {
                return "1" + " March " + new Date().getFullYear()
            } else {
                return Number(date) + 2 + " February " + new Date().getFullYear()
            }
        case "2" :
            if(date === "31") {
                return "2" + " April " + new Date().getFullYear()
            } else if(date === "30") {
                return "1" + " April " + new Date().getFullYear()
            } else {
                return Number(date) + 2 + " March " + new Date().getFullYear()
            }
        case "3" :
            if(date === "30") {
                return "2" + " May " + new Date().getFullYear()
            } else if(date === "29") {
                return "1" + " May " + new Date().getFullYear()
            } else {
                return Number(date) + 2 + " April " + new Date().getFullYear()
            }
        case "4" :
            if(date === "31") {
                return "2" + " June " + new Date().getFullYear()
            } else if(date === "30") {
                return "1" + " June " + new Date().getFullYear()
            } else {
                return Number(date) + 2 + " May " + new Date().getFullYear()
            }
        case "5" :
            if(date === "30") {
                return "2" + " July " + new Date().getFullYear()
            } else if(date === "29") {
                return "1" + " July " + new Date().getFullYear()
            } else {
                return Number(date) + 2 + " June " + new Date().getFullYear()
            }
        case "6" :
            if(date === "31") {
                return "2" + " August " + new Date().getFullYear()
            } else if(date === "30") {
                return "1" + " August " + new Date().getFullYear()
            } else {
                return Number(date) + 2 + " July " + new Date().getFullYear()
            }
        case "7" :
            if(date === "31") {
                return "2" + " September " + new Date().getFullYear()
            } else if(date === "30") {
                return "1" + " September " + new Date().getFullYear()
            } else {
                return Number(date) + 2 + " August " + new Date().getFullYear()
            }
        case "8" :
            if(date === "30") {
                return "2" + " October " + new Date().getFullYear()
            } else if(date === "29") {
                return "1" + " October " + new Date().getFullYear()
            } else {
                return Number(date) + 2 + " September " + new Date().getFullYear()
            }
        case "9" :
            if(date === "31") {
                return "2" + " November " + new Date().getFullYear()
            } else if(date === "30") {
                return "1" + " November " + new Date().getFullYear()
            } else {
                return Number(date) + 2 + " October " + new Date().getFullYear()
            }
        case "10" :
            if(date === "30") {
                return "2" + " December " + new Date().getFullYear()
            } else if(date === "29") {
                return "1" + " December " + new Date().getFullYear()
            } else {
                return Number(date) + 2 + " November " + new Date().getFullYear()
            }
        case "11" :
            if(date === "31") {
                return "2" + " January " + Number(new Date().getFullYear() + 1)
            } else if(date === "30") {
                return "1" + " January " + Number(new Date().getFullYear() + 1)
            } else {
                return Number(date) + 2 + " December " + new Date().getFullYear()
            }
    }
}
/////////////////////////////////////////////////////////////////////////////////////////////////////    KEEP    /////////////////////////////////////////////////////////////////////////////////////////////////////
/*
function getMonth(month) {
    switch(month) {
        case "0" :
            return "January"
        case "1" :
            return "February"
        case "2" :
            return "March"
        case "3" :
            return "April"
        case "4" :
            return "May"
        case "5" :
            return "June"
        case "6" :
            return "July"
        case "7" :
            return "August"
        case "8" :
            return "September"
        case "9" :
            return "October"
        case "10" :
            return "November"
        case "11" :
            return "December"
    }
}
*/
