const activeDocument = app.activeDocument
const path = activeDocument.filePath
const issue = activeDocument.filePath.toString()[46] + activeDocument.filePath.toString()[47]
const numberOfPages = activeDocument.pages.length
const oldDocumentPath = activeDocument.fullName
const logFilePath = oldDocumentPath.toString().slice(0, 46)
const logFile = File(logFilePath + issue + '/log.txt');




main()




function main() {
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

function getPageNumberFromLog(pageNumber) {
    changePageNumber(pageNumber)
    exportPDF(pageNumber)
    changeNameAndSaveAndDelete(pageNumber)
    return numberOfPages + pageNumber
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
                changeNameAndSaveAndDelete(firstPageNumber)
                return numberOfPages + firstPageNumber
            } else {
                alert('目錄要單數版開始')
            }
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


function changePageNumber(pageNumber) {
    activeDocument.documentPreferences.allowPageShuffle = true
    activeDocument.sections[0].continueNumbering = false
    activeDocument.sections[0].pageNumberStart = pageNumber
}
function exportPDF(pageNumber) {
    const exportPreset = app.pdfExportPresets.item('Output_pageup');
    activeDocument.exportFile(ExportFormat.pdfType, File('//APSRVX/HotFolderRoot/Gaz_PDF_Hires(new)/MA-' + issue + '-' + pageNumber + '.pdf'), false, exportPreset)
}
function changeNameAndSaveAndDelete(pageNumber) {
    activeDocument.save(File(path + '/MA-' + issue + '-' + pageNumber + '.indd'));
    activeDocument.close();
    oldDocumentPath.remove()
}