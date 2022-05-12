const logFile = File('~/Desktop/log.txt');

main()

function main() {
    logFile.open("r");
    const pageNumber = Number(logFile.readln())
    const path = getPath()
    const issue = getIssue()
    if(pageNumber > 0) {
        updateLogFile(getPageNumberFromLog(path, issue, pageNumber))
    } else {
        updateLogFile(getPageNumberFromDialog(path, issue))
    }
}

function getPageNumberFromLog(path, issue, pageNumber) {
    numberOfpages = app.activeDocument.pages.length
    changePageNumber(pageNumber)
    exportPDF(issue, pageNumber)
    changeNameAndSaveAndDelete(path, issue, pageNumber)
    return numberOfpages + pageNumber
}

function getPageNumberFromDialog(path, issue) {
    makeDialog();
    dialog.center();
    if(dialog.show() === 1) {
        const firstPageNumber = getFirstPageNumber()
        if(getFirstPageNumber() > 0) {
            numberOfpages = app.activeDocument.pages.length
            changePageNumber(firstPageNumber)
            exportPDF(issue, firstPageNumber)
            changeNameAndSaveAndDelete(path, issue, firstPageNumber)
            return numberOfpages + firstPageNumber
        } else { alert('輸入頁碼') }
    } else { exit() }   
}

function updateLogFile(nextPageNumber) {
    logFile.open("w");
    logFile.write(nextPageNumber)
    logFile.close();
}

function makeDialog() {
    dialog = new Window('dialog', "", "x:0, y:0, width:290, height:75");
    dialog.panel = dialog.add('panel', [10, 10, 185, 65], "");

    dialog.panel.add('statictext',  [10, 16, 80, 196], "第一頁頁碼"); 
    dialog.pageNumber = dialog.panel.add('edittext', [80, 11, 155, 36], "");
    dialog.pageNumber.onChange = pageNumberValidator;

    dialog.OKbut = dialog.add('button', [200, 10, 275, 35], "OK");
    dialog.OKbut.onClick = onOKclicked;
    dialog.CANbut = dialog.add('button', [200, 40, 275, 65], "Cancel");
    dialog.CANbut.onClick = onCANclicked;

    return dialog;
}
function pageNumberValidator() {
    pageValidator(dialog.pageNumber, placementINFO.pgCount, "");
}
function onOKclicked() {
    dialog.close(1);
}
function onCANclicked() {
    dialog.close(0);
}


function getFirstPageNumber() {
    return Number(dialog.pageNumber.text)
}
function getPath() {
    return app.activeDocument.filePath
}
function getIssue() {
    return app.activeDocument.name.toString()[3] + app.activeDocument.name.toString()[4]
}
function getNumberOfPages() {
    return app.activeDocument.pages.length
}


function changePageNumber(firstPageNumber) {
    app.activeDocument.documentPreferences.allowPageShuffle = true
    app.activeDocument.sections[0].continueNumbering = false
    app.activeDocument.sections[0].pageNumberStart = firstPageNumber
}
function exportPDF(issue, firstPageNumber) {
    const exportPreset = app.pdfExportPresets.item('Output_pageup');
//    app.activeDocument.exportFile(ExportFormat.pdfType, File('//APSRVX/HotFolderRoot/Gaz_PDF_Hires(new)/MA-' + issue + '-' + firstPageNumber + '.pdf'), false, exportPreset)
    app.activeDocument.exportFile(ExportFormat.pdfType, File('~/Desktop/MA-' + issue + '-' + firstPageNumber + '.pdf'), false, exportPreset)
}
function changeNameAndSaveAndDelete(path, issue, firstPageNumber) {
    const oldFile = app.activeDocument.fullName
    app.activeDocument.save(File(path + '/MA-' + issue + '-' + firstPageNumber + '.indd'));
    app.activeDocument.close();
//     oldFile.remove()
}


/*
updateLogFile(execute(newPageNumber))

function updateLogFile(nextPageNumber) {
    logFile.open("w");
    logFile.write(nextPageNumber)
    logFile.close();
}

function execute(newPageNumber) {
    const issue = getIssue()
    const path = getPath()
    if(newPageNumber > 0) {
            numberOfpages = app.activeDocument.pages.length
            changePageNumber(newPageNumber)
            exportPDF(issue, newPageNumber)
            changeNameAndSaveAndDelete(path, issue, newPageNumber)
        return numberOfpages + newPageNumber
    } else {
        makeDialog();
        dialog.center();
        
        if(dialog.show() === 1) {
            const firstPageNumber = getFirstPageNumber()
            const numberOfpages = 0
            if(getFirstPageNumber() > 0) {
                numberOfpages = app.activeDocument.pages.length
                changePageNumber(firstPageNumber)
                exportPDF(issue, firstPageNumber)
                changeNameAndSaveAndDelete(path, issue, firstPageNumber)
                return numberOfpages + firstPageNumber
            }
        } else {
            exit()
        }   
    }
}
*/