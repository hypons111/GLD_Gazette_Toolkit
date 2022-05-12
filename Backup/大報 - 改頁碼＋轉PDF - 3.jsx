makeDialog();
dialog.center();




if(dialog.show() === 1){
    if(getFirstPageNumber() > 0) {
        changePageNumber()
        exportPDF()
        changeNameAndSaveAndDelete()
    }
}else{
    exit();
}




function makeDialog() {
    dialog=new Window('dialog', "", "x:0, y:0, width:290, height:75");
    dialog.panel=dialog.add('panel', [10, 10, 185, 65], "");

    dialog.panel.add('statictext',  [10, 16, 80, 196], "第一頁頁碼"); 
    dialog.pageNumber=dialog.panel.add('edittext', [80, 11, 155, 36], "");
    dialog.pageNumber.onChange=pageNumberValidator;

    dialog.OKbut=dialog.add('button', [200, 10, 275, 35], "OK");
    dialog.OKbut.onClick=onOKclicked;
    dialog.CANbut=dialog.add('button', [200, 40, 275, 65], "Cancel");
    dialog.CANbut.onClick=onCANclicked;

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
function changePageNumber() {
    app.activeDocument.documentPreferences.allowPageShuffle = true
    app.activeDocument.sections[0].continueNumbering = false
    app.activeDocument.sections[0].pageNumberStart = getFirstPageNumber()
}
function exportPDF(firstPageNumber) {
    var exportPreset = app.pdfExportPresets.item('Output_pageup');
    app.activeDocument.exportFile(ExportFormat.pdfType, File('//APSRVX/HotFolderRoot/Gaz_PDF_Hires(new)/MA-' + getIssue() + '-' + getFirstPageNumber() + '.pdf'), false, exportPreset)
}
function changeNameAndSaveAndDelete(firstPageNumber) {
    var oldFile = app.activeDocument.fullName
    app.activeDocument.save(File(getPath() + '/MA-' + getIssue() + '-' + getFirstPageNumber() + '.indd'));
    app.activeDocument.close();
    oldFile.remove()
}