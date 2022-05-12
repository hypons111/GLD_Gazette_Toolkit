///////////    Version 4    ///////////

/////////////////////////////////////////////////////////////////////////////////////////////////////    LOG    /////////////////////////////////////////////////////////////////////////////////////////////////////

///////////    Version 3    ///////////    2021-2-19    ///////////
//made input gn number slot active

///////////    Version 4    ///////////    2021-7-26    ///////////
//improvded search result performace
//changed terms on buttoms 
/////////////////////////////////////////////////////////////////////////////////////////////////////    LOG    /////////////////////////////////////////////////////////////////////////////////////////////////////

var i;
var count;
var myFind;
var execute = 0

makeDialog();
dialog.center(); 
if(dialog.show() === 1){
    gN=Number(dialog.gN.text);
}else{
    exit();
}

app.findGrepPreferences.appliedParagraphStyle = "";

app.findGrepPreferences.findWhat = "(?<!\\r)(?<=^)第\\d{1, 5}號公告(?=\\r)|(?<=^)第\\d\+號公告\\r(?=勞工處\\r破產欠薪保障條例\\(第)";
myFind = app.activeDocument.findGrep();
if(myFind.length === 0) {
    app.findGrepPreferences.findWhat = "(?<!\\r)(?<=^)G.N. \\d{1, 5}(?=\\r)|(?<=^)G.N. \\d\+\\r(?=Labour Department\\rProtection of Wages on Insolvency Ordinance \\(Chapter )"; 
    myFind = app.activeDocument.findGrep();
}

if(myFind.length > 0){
    count = gN-1;
    app.findGrepPreferences.findWhat = "\\d\+";
    var changed = 0
    for(i = 0; i < myFind.length; i ++) {
        zoom()
        makeChangingDialog(myFind[i].contents, pepsi(count + 1));
        if(execute === 0) {
            app.changeGrepPreferences.changeTo = pepsi(++count);
            myFind[i].changeGrep();
            changed ++
        } else if(execute === 1){
            continue
        } else if(execute === 2){
            alert("已停止")
            break
        }
    }
    alert("最後的公告是 " + pepsi(count) + " 號\r已更改 " + changed + " 個公告");
}


// Create dialog box
function makeDialog() {
    dialog=new Window('dialog', "", "x:0, y:0, width:290, height:75");
    dialog.panel=dialog.add('panel', [10, 10, 185, 65], "");

    dialog.panel.add('statictext',  [10, 16, 80, 196], "第一個 G.N. "); 
    dialog.gN=dialog.panel.add('edittext', [80, 11, 155, 36], "");
    dialog.gN.onChange=gNValidator;
    dialog.gN.active = true
    // The buttons
    dialog.OKbut=dialog.add('button', [200, 10, 275, 35], "OK");
    dialog.OKbut.onClick=onOKclicked;
    dialog.CANbut=dialog.add('button', [200, 40, 275, 65], "Cancel");
    dialog.CANbut.onClick=onCANclicked;

    return dialog;
}

// Start G.N.
function gNValidator() {
    pageValidator(dialog.gN, placementINFO.pgCount, "Start Contents");
}

// Take care of OK being clicked
function onOKclicked() {
    dialog.close(1);
}
// Take care of Cancel being clicked
function onCANclicked() {
    dialog.close(0);
}


// Create changing box
function makeChangingDialog(target, number) {
    dialog=new Window('dialog', target + "     改成     " + number, "x:800, y:450, width:260, height:45");

    // buttons
    dialog.Changebut = dialog.add('button', [10, 10, 85, 35], "變更");
    dialog.Changebut.onClick = onChangeclicked;
    dialog.Nextbut = dialog.add('button', [95, 10, 170, 35], "下一個");
    dialog.Nextbut.onClick = onNextclicked;
    dialog.Abortbut = dialog.add('button', [180, 10, 245, 35], "停止");
    dialog.Abortbut.onClick = onAbortclicked;
//    dialog.center(); 
    dialog.show()
    return dialog;
}

function onChangeclicked(){
    dialog.close(0);
    execute = 0
    return
}
function onNextclicked(){
    dialog.close(0);
    execute = 1
    return
}
function onAbortclicked(){
    dialog.close(0);    
    execute = 2
    return
}

function zoom() {
    app.selection = myFind[i]
    app.activeDocument.layoutWindows[0].zoomPercentage = app.activeDocument.layoutWindows[0].zoomPercentage;
}


function pepsi (num) {
    var str = String(num);
    return str;
}
