var thisIssue = ''
var firstPageNumber = 0



getThisIssue()
makeDialog();
dialog.center();

if(dialog.show() === 1){
    firstPageNumber = Number(dialog.pageNumber.text);
}else{
    exit();
}

if(firstPageNumber !== 0) {
    changePageNumber()
    exportPDF()
    changeNameAndSaveAs()
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


function getThisIssue() {
    var issueList = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '30', '31', '32', '33', '34', '35', '36', '37', '38', '39', '40', '41', '42', '43', '44', '45', '46', '47', '48', '49', '50', '51', '52', '53']
    var guess = 'MA-' + issueList[0] + '-C01'
    for(var i = 0; i < issueList.length; i ++) { 

        try {
            var fileName = app.open (File('//EXZIP18/Gazette_N/-Main Gazette--Chinese/C' + issueList[i] + '/000-C' + issueList[i] + '.indt'));
            if(fileName) {
                thisIssue = issueList[i]
                fileName.close();
                break
            }
        } catch (error) {}
    }
}


function changePageNumber() {
    app.activeDocument.documentPreferences.allowPageShuffle = true
    app.activeDocument.sections[0].continueNumbering = false
    app.activeDocument.sections[0].pageNumberStart = firstPageNumber
}


function exportPDF() {
    var exportPreset = app.pdfExportPresets.item('Output_pageup');
    app.activeDocument.exportFile(ExportFormat.pdfType, File('//APSRVX/HotFolderRoot/Gaz_PDF_Hires(new)/MA-' + thisIssue + '-' + firstPageNumber + '.pdf'), false, exportPreset)
}

function changeNameAndSaveAs() {
    app.activeDocument.save(File('~/Desktop/Here/MA-' + thisIssue + '-' + firstPageNumber + '.indd'));
    app.activeDocument.close();
}




