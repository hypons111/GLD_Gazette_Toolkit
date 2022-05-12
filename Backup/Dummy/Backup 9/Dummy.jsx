var type;
var nop32 = 0, nop16 = 0, nop12 = 0, nop8 = 0, nop4 = 0;
var saddleNop32 = 0, saddleNop16 = 0, saddleNop12 = 0, saddleNop8 = 0, saddleNop4 = 0;
var blank = 0;
var jobName;
var totalPage;
var startContents, endContents;
var startpage, endPage;
var pageCounter;
var sectionCounter = 1;
var cover;
var imprint;
var content;
var text;
var 類型 = ["大報", "憲報 No. 1", "憲報 No. 2", "憲報 No. 3", "憲報 No. 4", "憲報 No. 5", "憲報 No. 6", "特報", "簽字", "No. 1 + Index A", "No. 2 + Index B", "No. 3 + Index C", "Sup 1 + Index A", "Sup 2 + Index B", "Sup 3 + Index C"];

var document = app.documents.add();
document.documentPreferences.facingPages = false;
app.documents.item(0).viewPreferences.horizontalMeasurementUnits = MeasurementUnits.millimeters;
app.documents.item(0).viewPreferences.verticalMeasurementUnits = MeasurementUnits.millimeters;
document.marginPreferences.top = 0;
document.marginPreferences.left = 0;
document.marginPreferences.bottom = 0;
document.marginPreferences.right = 0;
document.documentPreferences.pageWidth = 210;
document.documentPreferences.pageHeight = 297;
blue = document.colors.add({name:"Blue", space:ColorSpace.CMYK, model:ColorModel.process, colorValue:[100, 0, 0, 0]});


//  Create 32 pages section master page
document.masterSpreads.add();
var mP32 = document.masterSpreads.item(1);
mP32.rectangles.add({geometricBounds:[45, 15, 71.516, 52.5], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[71.516, 15, 98.032, 52.5], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[98.032, 15, 124.548, 52.5], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[124.548, 15, 151.064, 52.5], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[45, 52.5 , 71.516, 90], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[71.516, 52.5 , 98.032, 90], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[98.032, 52.5 , 124.548, 90], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[124.548, 52.5 , 151.064, 90], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[45, 90, 71.516, 127.5], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[71.516, 90, 98.032, 127.5], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[98.032, 90, 124.548, 127.5], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[124.548, 90, 151.064, 127.5], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[45, 127.5, 71.516, 165], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[71.516, 127.5, 98.032, 165], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[98.032, 127.5, 124.548, 165], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[124.548, 127.5, 151.064, 165], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[175.936, 15, 202.452, 52.5], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[202.452, 15, 228.968, 52.5], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[228.968, 15, 255.484, 52.5], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[255.484, 15, 282, 52.5], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[175.936, 52.5 , 202.452, 90], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[202.452, 52.5 , 228.968, 90], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[228.968, 52.5 , 255.484, 90], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[255.484, 52.5 , 282, 90], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[175.936, 90, 202.452, 127.5], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[202.452, 90, 228.968, 127.5], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[228.968, 90, 255.484, 127.5], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[255.484, 90, 282, 127.5], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[175.936, 127.5, 202.452, 165], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[202.452, 127.5, 228.968, 165], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[228.968, 127.5, 255.484, 165], strokeWeight:1});
mP32.rectangles.add({geometricBounds:[255.484, 127.5, 282, 165], strokeWeight:1});

//  Create 16 pages section master page
document.masterSpreads.add();
var mP16 = document.masterSpreads.item(2);
mP16.rectangles.add({geometricBounds:[45, 15, 71.516, 52.5], strokeWeight:1});
mP16.rectangles.add({geometricBounds:[71.516, 15, 98.032, 52.5], strokeWeight:1});
mP16.rectangles.add({geometricBounds:[98.032, 15, 124.548, 52.5], strokeWeight:1});
mP16.rectangles.add({geometricBounds:[124.548, 15, 151.064, 52.5], strokeWeight:1});
mP16.rectangles.add({geometricBounds:[45, 52.5 , 71.516, 90], strokeWeight:1});
mP16.rectangles.add({geometricBounds:[71.516, 52.5 , 98.032, 90], strokeWeight:1});
mP16.rectangles.add({geometricBounds:[98.032, 52.5 , 124.548, 90], strokeWeight:1});
mP16.rectangles.add({geometricBounds:[124.548, 52.5 , 151.064, 90], strokeWeight:1});
mP16.rectangles.add({geometricBounds:[45, 90, 71.516, 127.5], strokeWeight:1});
mP16.rectangles.add({geometricBounds:[71.516, 90, 98.032, 127.5], strokeWeight:1});
mP16.rectangles.add({geometricBounds:[98.032, 90, 124.548, 127.5], strokeWeight:1});
mP16.rectangles.add({geometricBounds:[124.548, 90, 151.064, 127.5], strokeWeight:1});
mP16.rectangles.add({geometricBounds:[45, 127.5, 71.516, 165], strokeWeight:1});
mP16.rectangles.add({geometricBounds:[71.516, 127.5, 98.032, 165], strokeWeight:1});
mP16.rectangles.add({geometricBounds:[98.032, 127.5, 124.548, 165], strokeWeight:1});
mP16.rectangles.add({geometricBounds:[124.548, 127.5, 151.064, 165], strokeWeight:1});

//  Create 12 pages section master page
document.masterSpreads.add();
var mP12 = document.masterSpreads.item(3);
mP12.rectangles.add({geometricBounds:[45, 15, 71.516, 52.5], strokeWeight:1});
mP12.rectangles.add({geometricBounds:[71.516, 15, 98.032, 52.5], strokeWeight:1});
mP12.rectangles.add({geometricBounds:[45, 52.5 , 71.516, 90], strokeWeight:1});
mP12.rectangles.add({geometricBounds:[71.516, 52.5 , 98.032, 90], strokeWeight:1});
mP12.rectangles.add({geometricBounds:[45, 90, 71.516, 127.5], strokeWeight:1});
mP12.rectangles.add({geometricBounds:[71.516, 90, 98.032, 127.5], strokeWeight:1});
mP12.rectangles.add({geometricBounds:[175.936, 15, 202.452, 52.5], strokeWeight:1});
mP12.rectangles.add({geometricBounds:[202.452, 15, 228.968, 52.5], strokeWeight:1});
mP12.rectangles.add({geometricBounds:[175.936, 52.5 , 202.452, 90], strokeWeight:1});
mP12.rectangles.add({geometricBounds:[202.452, 52.5 , 228.968, 90], strokeWeight:1});
mP12.rectangles.add({geometricBounds:[175.936, 90, 202.452, 127.5], strokeWeight:1});
mP12.rectangles.add({geometricBounds:[202.452, 90, 228.968, 127.5], strokeWeight:1});

//  Create 8 pages section master page
document.masterSpreads.add();
var mP8 = document.masterSpreads.item(4);
mP8.rectangles.add({geometricBounds:[45, 15, 71.516, 52.5], strokeWeight:1});
mP8.rectangles.add({geometricBounds:[71.516, 15, 98.032, 52.5], strokeWeight:1});
mP8.rectangles.add({geometricBounds:[45, 52.5 , 71.516, 90], strokeWeight:1});
mP8.rectangles.add({geometricBounds:[71.516, 52.5 , 98.032, 90], strokeWeight:1});
mP8.rectangles.add({geometricBounds:[175.936, 15, 202.452, 52.5], strokeWeight:1});
mP8.rectangles.add({geometricBounds:[202.452, 15, 228.968, 52.5], strokeWeight:1});
mP8.rectangles.add({geometricBounds:[175.936, 52.5 , 202.452, 90], strokeWeight:1});
mP8.rectangles.add({geometricBounds:[202.452, 52.5 , 228.968, 90], strokeWeight:1});

//  Create 4 pages section master page
document.masterSpreads.add();
var mP4 = document.masterSpreads.item(5);
mP4.rectangles.add({geometricBounds:[45, 15, 71.516, 52.5], strokeWeight:1});
mP4.rectangles.add({geometricBounds:[71.516, 15, 98.032, 52.5], strokeWeight:1});
mP4.rectangles.add({geometricBounds:[45, 52.5 , 71.516, 90], strokeWeight:1});
mP4.rectangles.add({geometricBounds:[71.516, 52.5 , 98.032, 90], strokeWeight:1});



//  Create number page header
document.masterSpreads.add();
var numberPage = document.masterSpreads.item(6);
var tableFrameHeader = document.masterSpreads.item(6).textFrames.add();
tableFrameHeader.geometricBounds = ["28", "10", "34", "200"];    //["上", "左", "下", "右",]
var numberPageHeader = tableFrameHeader.tables.add({columnCount:11, bodyRowCount:1});
numberPageHeader.rows.everyItem().height = 5.75;
numberPageHeader.columns.everyItem().width = 18.328;
numberPageHeader.columns[0].width = 6.469;


////////////////////////////////////////////////////////////////////////
numberPageHeader.rows[0].bottomEdgeStrokeColor = "Blue";
numberPageHeader.rows[0].columns[0].bottomEdgeStrokeWeight = 10;
////////////////////////////////////////////////////////////////////////


//  Create number page table
var tableFrame = document.masterSpreads.item(6).textFrames.add();
tableFrame.geometricBounds = ["34", "10", "240.25", "200"];    //["上", "左", "下", "右",]
var numberPageTable = tableFrame.tables.add({columnCount:21, bodyRowCount:32});
numberPageTable.rows.everyItem().height = 6.437;
numberPageTable.columns.everyItem().width = 6.469;
for(var i=1; i<20; i=i+2){
    numberPageTable.columns[i].width = 11.859;
}

//  Create number page footer
var tableFrameFooter = document.masterSpreads.item(6).textFrames.add();
tableFrameFooter.geometricBounds = ["240.25 ", "10", "252", "200"];    //["上", "左", "下", "右",]
var numberPageFooter = tableFrameFooter.tables.add({columnCount:11, bodyRowCount:2});
numberPageFooter.rows.everyItem().height = 5.75;
numberPageFooter.columns.everyItem().width = 18.328;
numberPageFooter.columns[0].width = 6.469;


var tag = ["", "A", "B", "C", "D", "E",];
if(app.documents.length == 0){
    var tmp = new Array();
    tmp["width"] = 612;
    tmp["height"] = 792;
    placementINFO["pgSize"]  = tmp;
}   

//  declare dialog data information
makeDialog();
dialog.center(); 
if(dialog.show() == 1){
    jobName = String(dialog.jobName.text);
    cover = Number(dialog.cover.text);
    startContents = Number(dialog.startContents.text);
    endContents = Number(dialog.endContents.text);
    startPage = Number(dialog.startPage.text);
    endPage = Number(dialog.endPage.text);
    type = dialog.kindType.selection.index;
}else{
    exit();
}


//  32 page section information
mP32TF = document.masterSpreads.item(1).textFrames.add();  
mP32TF.geometricBounds = ["10", "15", "30", "165"];    //["上", "左", "下", "右",]
mP32TF.texts.item(0).contents = 類型[type] + " 32 Pages " + jobName;
mP32TF.texts.item(0).appliedFont = "新細明體";
mP32TF.texts.item(0).pointSize = 20;
mP32TF.texts.item(0).justification = Justification.leftAlign;
mP32TF.texts.item(0).fillColor = "Blue";

//  16 page section information
mP16TF = document.masterSpreads.item(2).textFrames.add();  
mP16TF.geometricBounds = ["10", "15", "30", "165"];    //["上", "左", "下", "右",]
mP16TF.texts.item(0).contents = 類型[type] + " 16 Pages";
mP16TF.texts.item(0).appliedFont = "新細明體";
mP16TF.texts.item(0).pointSize = 20;
mP16TF.texts.item(0).justification = Justification.leftAlign;
mP32TF.texts.item(0).fillColor = "Blue";

//  12 page section information
mP12TF = document.masterSpreads.item(3).textFrames.add();  
mP12TF.geometricBounds = ["10", "15", "30", "165"];    //["上", "左", "下", "右",]
mP12TF.texts.item(0).contents = 類型[type] + " 12 Pages";
mP12TF.texts.item(0).appliedFont = "新細明體";
mP12TF.texts.item(0).pointSize = 20;
mP12TF.texts.item(0).justification = Justification.leftAlign;
mP32TF.texts.item(0).fillColor = "Blue";

//  8 page section information
mP8TF = document.masterSpreads.item(4).textFrames.add();  
mP8TF.geometricBounds = ["10", "15", "30", "165"];    //["上", "左", "下", "右",]
mP8TF.texts.item(0).contents = 類型[type] + " 8 Pages";
mP8TF.texts.item(0).appliedFont = "新細明體";
mP8TF.texts.item(0).pointSize = 20;
mP8TF.texts.item(0).justification = Justification.leftAlign;
mP32TF.texts.item(0).fillColor = "Blue";

//  4 page section information
mP4TF = document.masterSpreads.item(5).textFrames.add();  
mP4TF.geometricBounds = ["10", "15", "30", "165"];    //["上", "左", "下", "右",]
mP4TF.texts.item(0).contents = 類型[type] + " 4 Pages";
mP4TF.texts.item(0).appliedFont = "新細明體";
mP4TF.texts.item(0).pointSize = 20;
mP4TF.texts.item(0).justification = Justification.leftAlign;
mP32TF.texts.item(0).fillColor = "Blue";


// Create dialog box
function makeDialog()
{
    dialog = new Window('dialog', "HAHAHA", "x:0, y:0, width:550, height:550");
    dialog.panel = dialog.add('panel', [10, 10, 390, 390], "");


    dialog.panel.add('statictext', [10,15,50,40], "類型：");
    dialog.kindType = dialog.panel.add('dropdownlist', [100,10,215,35]);
    dialog.kindType.onChange = kindValidator;
    for(var i=0; i<類型.length; i++){
        dialog.kindType.add('item', 類型[i]);
    }
    
    dialog.panel.add('statictext',  [10,55,90,80], "Job Name："); 
    dialog.jobName = dialog.panel.add('edittext', [100,50,215,75], "");
    dialog.jobName.onChange = jobNameValidator;
    
    
    dialog.panel.add('statictext',  [10, 95, 50, 120], "封面："); 
    dialog.cover = dialog.panel.add('edittext', [100, 90, 150, 115], "");
    dialog.cover.onChange = coverValidator;

    dialog.panel.add('statictext',  [10, 175, 50, 195], "目錄："); 
    dialog.startContents = dialog.panel.add('edittext', [100, 170, 150, 195], "");
    dialog.startContents.onChange = startContentsValidator;
    dialog.panel.add('statictext',  [155, 175, 160, 200], "-");
    dialog.endContents = dialog.panel.add('edittext', [165, 170, 215, 195], "");
    dialog.cropType = dialog.panel.add('dropdownlist', [220, 170, 260, 195]);
    for(var i=0; i<tag.length; i++){
        dialog.cropType.add('item', tag[i]);
    }

    dialog.panel.add('statictext',  [10, 215, 50, 240], "內文："); 
    dialog.startPage = dialog.panel.add('edittext', [100, 210, 150, 235], "");
    dialog.startPage.onChange = startPageValidator;
    dialog.panel.add('statictext',  [155, 215, 160, 240], "-");
    dialog.endPage = dialog.panel.add('edittext', [165, 210, 215, 235], ""); 
    dialog.endPage.onChange = endPageValidator;      
    dialog.cropType = dialog.panel.add('dropdownlist', [220, 210, 260, 235]);
    for(var i=0; i<tag.length; i++){
        dialog.cropType.add('item', tag[i]);
    }

    dialog.placeOnLayer = dialog.panel.add('checkbox', [10, 250, 220, 285], " \"This\" Page");

    // The buttons
    dialog.OKbut = dialog.add('button', [450, 20, 525, 45], "OK");
    dialog.OKbut.onClick = onOKclicked;
    dialog.CANbut = dialog.add('button', [450, 50, 525, 75], "Cancel");
    dialog.CANbut.onClick = onCANclicked;

    return dialog;
}

//  get total page number
totalPage = endPage - startContents + 1;

//  make pageCounter = total page number to count sections
pageCounter = totalPage;

//  count 32 pages section
for(; pageCounter >= 32;){
    nop32 ++;
    pageCounter = pageCounter - 32;
}
if(pageCounter == 31){
    nop32 ++;
    pageCounter = pageCounter - 32;
    blank = blank + 1;
}else if(pageCounter == 30){
    nop32 ++;
    pageCounter = pageCounter - 32;
    blank = blank + 2;
}else if(pageCounter == 29){
    nop32 ++;
    pageCounter = pageCounter - 32;
    blank = blank + 3;
}

//  count 16 pages section
if(pageCounter >= 16){
    nop16 ++;
    pageCounter = pageCounter - 16;
}else if(pageCounter == 15){
    nop16 ++;
    pageCounter = pageCounter - 16;
    blank = blank + 1;
}else if(pageCounter == 14){
    nop16 ++;
    pageCounter = pageCounter - 16;
    blank = blank + 2;
}else if(pageCounter == 13){
    nop16 ++;
    pageCounter = pageCounter - 16;
    blank = blank + 3;
}

//  count 12 pages section
if(pageCounter >= 12){
    nop12 ++;
    pageCounter = pageCounter - 12;
}else if(pageCounter == 11){
    nop12 ++;
    pageCounter = pageCounter - 12;
    blank = blank + 1;
}else if(pageCounter == 10){
    nop12 ++;
    pageCounter = pageCounter - 12;
    blank = blank + 2;
}else if(pageCounter == 9){
    nop12 ++;
    pageCounter = pageCounter - 12;
    blank = blank + 3;
}

//  count 8 pages section
if(pageCounter >= 8){
    nop8 ++;
    pageCounter = pageCounter - 8;
}else if(pageCounter == 7){
    nop8 ++;
    pageCounter = pageCounter - 8;
    blank = blank + 1;
}else if(pageCounter == 6){
    nop8 ++;
    pageCounter = pageCounter - 8;
    blank = blank + 2;
}else if(pageCounter == 5){
    nop8 ++;
    pageCounter = pageCounter - 8;
    blank = blank + 3;
}

//  count 4 pages section
if(pageCounter >= 4){
    nop4 ++;
    pageCounter = pageCounter - 4;
}else if(pageCounter == 3){
    nop4 ++;
    pageCounter = pageCounter - 4;
    blank = blank + 1;
}else if(pageCounter == 2){
    nop4 ++;
    pageCounter = pageCounter - 4;
    blank = blank + 2;
}else if(pageCounter == 1){
    nop4 ++;
    pageCounter = pageCounter - 4;
    blank = blank + 3;
}

// make information page
makeText(document.pages.item(0), [10, 20, 40, 70], "Blank Page\n" + blank, "Times New Roman", 25, Justification.centerAlign, false, 0);
makeText(document.pages.item(0), [40, 20, 70, 70], "32 Pages\n" + nop32, "Times New Roman", 25, Justification.centerAlign, false, 0);
makeText(document.pages.item(0), [70, 20, 100, 70], "16  Pages\n" + nop16, "Times New Roman", 25, Justification.centerAlign, false, 0);
makeText(document.pages.item(0), [100, 20, 130, 70], "12  Pages\n" + nop12, "Times New Roman", 25, Justification.centerAlign, false, 0);
makeText(document.pages.item(0), [130, 20, 160, 70], "8 Pages\n" + nop8, "Times New Roman", 25, Justification.centerAlign, false, 0);
makeText(document.pages.item(0), [160, 20, 190, 70], "4 Pages\n" + nop4, "Times New Roman", 25, Justification.centerAlign, false, 0);
makeText(document.pages.item(0), [10, 70, 40, 120], "totalPage\n" + totalPage, "Times New Roman", 25, Justification.centerAlign, false, 0);
if (totalPage > 224){
    makeText(document.pages.item(0), [40, 70, 70, 120], " (膠裝) ", "Times New Roman", 25, Justification.centerAlign, false, 0);
}else{
    makeText(document.pages.item(0), [40, 70, 70, 120], " (騎釘) ", "Times New Roman", 25, Justification.centerAlign, false, 0);
}

//  Pefect Bound
if (totalPage > 224){
//  make 32 pages number table
    for(; nop32 > 1;){
    nop32 --;
        document.pages.add().appliedMaster = app.documents[0].masterSpreads.item ('B-主版');

        makeText(document.pages.item(sectionCounter), [66.024, 20.492, 103.524, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [223.476, 20.492, 260.976, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [223.476, 132.992, 260.976, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [66.024, 132.992, 103.524, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [92.54, 132.992, 130.04, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [196.96, 132.992, 234.46, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [196.96, 20.492, 234.46, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [92.54, 20.492, 130.04, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
         makeText(document.pages.item(sectionCounter), [92.54, 57.992, 130.04, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [196.96, 57.992, 234.46, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [196.96, 95.492, 234.46, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [92.54, 95.492, 130.04, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [66.024, 95.492, 103.524, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [223.476, 95.492, 260.976, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [223.476, 57.992, 260.976, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [66.024, 57.992, 103.524, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [39.508, 57.992, 77.008, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [249.992, 57.992, 287.492, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [249.992, 95.492, 287.492, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [39.508, 95.492, 77.008, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [119.056, 95.492, 156.556, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [170.444, 95.492, 207.944, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [170.444, 57.992, 207.944, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [119.056, 57.992, 156.556, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [119.056, 20.492, 156.556, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [170.444, 20.492, 207.944, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [170.444, 132.992, 207.944, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [119.056, 132.992, 156.556, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [39.508, 132.992, 77.008, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [249.992, 132.992, 287.492, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [249.992, 20.492, 287.492, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [39.508, 20.492, 77.008, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        sectionCounter ++;
    }

    //  make 4 pages number table
    for(; nop4 > 0;){
        nop4 --;
        document.pages.add().appliedMaster = app.documents[0].masterSpreads.item ('F-主版');

        makeText(document.pages.item(sectionCounter), [66.024, 20.492, 103.524, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        if(startContents <= endPage){
                makeText(document.pages.item(sectionCounter), [66.024, 57.992, 103.524, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
                startContents ++;
        }else{
                makeBlueText(document.pages.item(sectionCounter), [66.024, 57.992, 103.524, 84.508], "Blank", "Times New Roman", 25, Justification.centerAlign, false, 90);
                blank --;
        }        
        if(startContents <= endPage){
                makeText(document.pages.item(sectionCounter), [39.508, 57.992, 77.008, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
                startContents ++;
        }else{
                makeBlueText(document.pages.item(sectionCounter), [39.508, 57.992, 77.008, 84.508], "Blank", "Times New Roman", 25, Justification.centerAlign, false, 90);
                blank --;
        }
        if(startContents <= endPage){
                makeText(document.pages.item(sectionCounter), [39.508, 20.492, 77.008, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
                startContents ++;
        }else{
                makeBlueText(document.pages.item(sectionCounter), [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, Justification.centerAlign, false, -90);
                blank --;
        }        
        sectionCounter ++;
    }

    //  make 8 pages number table
    for(; nop8 > 0;){
        nop8 --;
        document.pages.add().appliedMaster = app.documents[0].masterSpreads.item ('E-主版');

        makeText(document.pages.item(sectionCounter), [66.024, 20.492, 103.524, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [170.444, 20.492, 207.944, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [170.444, 57.992, 207.944, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [66.024, 57.992, 103.524, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [39.508, 57.992, 77.008, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        if(startContents <= endPage){
                makeText(document.pages.item(sectionCounter), [196.96, 57.992, 234.46, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
                startContents ++;
        }else{
                makeBlueText(document.pages.item(sectionCounter), [196.96, 57.992, 234.46, 84.508], "Blank", "Times New Roman", 25, Justification.centerAlign, false, 90);
                blank --;
        }        
        if(startContents <= endPage){
                makeText(document.pages.item(sectionCounter), [196.96, 20.492, 234.46, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
                startContents ++;
        }else{
                makeBlueText(document.pages.item(sectionCounter), [196.96, 20.492, 234.46, 47.008], "Blank", "Times New Roman", 25, Justification.centerAlign, false, -90);
                blank --;
        }
        if(startContents <= endPage){
                makeText(document.pages.item(sectionCounter), [39.508, 20.492, 77.008, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
                startContents ++;
        }else{
                makeBlueText(document.pages.item(sectionCounter), [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, Justification.centerAlign, false, -90);
                blank --;
        }        
        sectionCounter ++;
    }

    //  make 12 pages number table
    for(; nop12 > 0;){
        nop12 --;
        document.pages.add().appliedMaster = app.documents[0].masterSpreads.item ('D-主版');

        makeText(document.pages.item(sectionCounter), [66.024, 20.492, 103.524, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;

        makeText(document.pages.item(sectionCounter), [170.444, 20.492, 207.944, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [66.024, 95.492, 103.524, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [170.44, 95.492, 207.944, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [170.444, 57.992, 207.944, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [66.024, 57.992, 103.524, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [39.508, 57.992, 77.008, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [196.96, 57.992, 234.46, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [196.96, 95.492, 234.46, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        if(startContents <= endPage){
                makeText(document.pages.item(sectionCounter), [39.508, 95.492, 77.008, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
                startContents ++;
        }else{
                makeBlueText(document.pages.item(sectionCounter), [39.508, 95.492, 77.008, 122.008], "Blank", "Times New Roman", 25, Justification.centerAlign, false, -90);
                blank --;
        }        
        if(startContents <= endPage){
                makeText(document.pages.item(sectionCounter), [196.96, 20.492, 234.46, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
                startContents ++;
        }else{
                makeBlueText(document.pages.item(sectionCounter), [196.96, 20.492, 234.46, 47.008], "Blank", "Times New Roman", 25, Justification.centerAlign, false, -90);
                blank --;
        }
        if(startContents <= endPage){
                makeText(document.pages.item(sectionCounter), [39.508, 20.492, 77.008, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
                startContents ++;
        }else{
                makeBlueText(document.pages.item(sectionCounter), [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, Justification.centerAlign, false, -90);
                blank --;
        }        
        sectionCounter ++;
    }

    //  make 16 pages number table
    for(; nop16 > 0;){
        nop16 --;
        document.pages.add().appliedMaster = app.documents[0].masterSpreads.item ('C-主版');

        makeText(document.pages.item(sectionCounter), [66.024, 20.492, 103.524, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [66.024, 132.992, 103.524, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [92.54, 132.992, 130.04, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [92.54, 20.492, 130.04, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [92.54, 57.992, 130.04, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [92.54, 95.492, 130.04, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [66.024, 95.492, 103.524, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [66.024, 57.992, 103.524, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [39.508, 57.992, 77.008, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [39.508, 95.492, 77.008, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [119.056, 95.492, 156.556, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [119.056, 57.992, 156.556, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [119.056, 20.492, 156.556, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        if(startContents <= endPage){
                makeText(document.pages.item(sectionCounter), [119.056, 132.992, 156.556, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
                startContents ++;
        }else{
                makeBlueText(document.pages.item(sectionCounter), [119.056, 132.992, 156.556, 159.508], "Blank", "Times New Roman", 25, Justification.centerAlign, false, 90);
                blank --;
        }
        if(startContents <= endPage){
                makeText(document.pages.item(sectionCounter), [39.508, 132.992, 77.008, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
                startContents ++;
        }else{
                makeBlueText(document.pages.item(sectionCounter), [39.508, 132.992, 77.008, 159.508], "Blank", "Times New Roman", 25, Justification.centerAlign, false, 90);
                blank --;
        }
        if(startContents <= endPage){
                makeText(document.pages.item(sectionCounter), [39.508, 20.492, 77.008, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
                startContents ++;
        }else{
                makeBlueText(document.pages.item(sectionCounter), [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, Justification.centerAlign, false, -90);
                blank --;
        }
        sectionCounter ++;
    }

    //  make the last 32 pages number table
    for(; nop32 > 0;){
        nop32 --;
        document.pages.add().appliedMaster = app.documents[0].masterSpreads.item ('B-主版');

        if(startContents )
        makeText(document.pages.item(sectionCounter), [66.024, 20.492, 103.524, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [223.476, 20.492, 260.976, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [223.476, 132.992, 260.976, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [66.024, 132.992, 103.524, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [92.54, 132.992, 130.04, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [196.96, 132.992, 234.46, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [196.96, 20.492, 234.46, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [92.54, 20.492, 130.04, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
         makeText(document.pages.item(sectionCounter), [92.54, 57.992, 130.04, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [196.96, 57.992, 234.46, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [196.96, 95.492, 234.46, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [92.54, 95.492, 130.04, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [66.024, 95.492, 103.524, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [223.476, 95.492, 260.976, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [223.476, 57.992, 260.976, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [66.024, 57.992, 103.524, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [39.508, 57.992, 77.008, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [249.992, 57.992, 287.492, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [249.992, 95.492, 287.492, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [39.508, 95.492, 77.008, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [119.056, 95.492, 156.556, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [170.444, 95.492, 207.944, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [170.444, 57.992, 207.944, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [119.056, 57.992, 156.556, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [119.056, 20.492, 156.556, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [170.444, 20.492, 207.944, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [170.444, 132.992, 207.944, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [119.056, 132.992, 156.556, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [39.508, 132.992, 77.008, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        if(startContents <= endPage){
                makeText(document.pages.item(sectionCounter), [249.992, 132.992, 287.492, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
                startContents ++;
        }else{
                makeBlueText(document.pages.item(sectionCounter), [249.992, 132.992, 287.492, 159.508], "Blank", "Times New Roman", 25, Justification.centerAlign, false, 90);
                blank --;
        }
        if(startContents <= endPage){
                makeText(document.pages.item(sectionCounter), [249.992, 20.492, 287.492, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
                startContents ++;
        }else{
                makeBlueText(document.pages.item(sectionCounter), [249.992, 20.492, 287.492, 47.008], "Blank", "Times New Roman", 25, Justification.centerAlign, false, -90);
                blank --;
        }
        if(startContents <= endPage){
                makeText(document.pages.item(sectionCounter), [39.508, 20.492, 77.008, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
                startContents ++;
        }else{
                makeBlueText(document.pages.item(sectionCounter), [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, Justification.centerAlign, false, -90);
                blank --;
        }        
        sectionCounter ++;
    }



//  Saddle Sitiched  //  make 32 pages number table
}else{
        for(; saddleNop32 < (nop32 - 1);){
        document.pages.add().appliedMaster = app.documents[0].masterSpreads.item ('B-主版');

        makeText(document.pages.item(sectionCounter), [66.024, 20.492, 103.524, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [223.476, 20.492, 260.976, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [223.476, 132.992, 260.976, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [66.024, 132.992, 103.524, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [92.54, 132.992, 130.04, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [196.96, 132.992, 234.46, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [196.96, 20.492, 234.46, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [92.54, 20.492, 130.04, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [92.54, 57.992, 130.04, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [196.96, 57.992, 234.46, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [196.96, 95.492, 234.46, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [92.54, 95.492, 130.04, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [66.024, 95.492, 103.524, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [223.476, 95.492, 260.976, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [223.476, 57.992, 260.976, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [66.024, 57.992, 103.524, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        saddleNop32 ++;
        sectionCounter ++;
    }

    //  Saddle Sitiched  //  make 4 pages number table
    if(saddleNop4 < nop4){
        document.pages.add().appliedMaster = app.documents[0].masterSpreads.item ('F-主版');

        makeText(document.pages.item(sectionCounter), [66.024, 20.492, 103.524, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [66.024, 57.992, 103.524, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        saddleNop4 ++;
        if(nop32 > 0){
            sectionCounter ++;
        }
    }

    //  Saddle Sitiched  //  make 8 pages number table
    if(saddleNop8 < nop8){
        document.pages.add().appliedMaster = app.documents[0].masterSpreads.item ('E-主版');

        makeText(document.pages.item(sectionCounter), [66.024, 20.492, 103.524, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [170.444, 20.492, 207.944, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [170.444, 57.992, 207.944, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [66.024, 57.992, 103.524, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        saddleNop8 ++;
        if(nop32 > 0){
            sectionCounter ++;
        }
    }

    //  Saddle Sitiched  //  make 12 pages number table
    if(saddleNop12 < nop12){
        document.pages.add().appliedMaster = app.documents[0].masterSpreads.item ('D-主版');

        makeText(document.pages.item(sectionCounter), [66.024, 20.492, 103.524, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [170.444, 20.492, 207.944, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [66.024, 95.492, 103.524, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [170.44, 95.492, 207.944, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [170.444, 57.992, 207.944, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [66.024, 57.992, 103.524, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        saddleNop12 ++;
        if(nop32 > 0){
            sectionCounter ++;
        }
    }

    //  Saddle Sitiched  //  make 16 pages number table
    if(saddleNop16 < nop16){
        document.pages.add().appliedMaster = app.documents[0].masterSpreads.item ('C-主版');

        makeText(document.pages.item(sectionCounter), [66.024, 20.492, 103.524, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [66.024, 132.992, 103.524, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [92.54, 132.992, 130.04, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [92.54, 20.492, 130.04, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [92.54, 57.992, 130.04, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [92.54, 95.492, 130.04, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [66.024, 95.492, 103.524, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [66.024, 57.992, 103.524, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        saddleNop16 ++;
        if(nop32 > 0){
            sectionCounter ++;
        }
    }

    //  Saddle Sitiched  //  make middle 32 pages number table
    if(saddleNop32 == (nop32 - 1)){
        document.pages.add().appliedMaster = app.documents[0].masterSpreads.item ('B-主版');

        makeText(document.pages.item(sectionCounter), [66.024, 20.492, 103.524, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [223.476, 20.492, 260.976, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [223.476, 132.992, 260.976, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [66.024, 132.992, 103.524, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [92.54, 132.992, 130.04, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [196.96, 132.992, 234.46, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [196.96, 20.492, 234.46, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [92.54, 20.492, 130.04, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
         makeText(document.pages.item(sectionCounter), [92.54, 57.992, 130.04, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [196.96, 57.992, 234.46, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [196.96, 95.492, 234.46, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [92.54, 95.492, 130.04, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [66.024, 95.492, 103.524, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [223.476, 95.492, 260.976, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [223.476, 57.992, 260.976, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [66.024, 57.992, 103.524, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [39.508, 57.992, 77.008, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [249.992, 57.992, 287.492, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [249.992, 95.492, 287.492, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [39.508, 95.492, 77.008, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [119.056, 95.492, 156.556, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [170.444, 95.492, 207.944, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [170.444, 57.992, 207.944, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [119.056, 57.992, 156.556, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [119.056, 20.492, 156.556, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [170.444, 20.492, 207.944, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [170.444, 132.992, 207.944, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [119.056, 132.992, 156.556, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [39.508, 132.992, 77.008, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [249.992, 132.992, 287.492, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [249.992, 20.492, 287.492, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [39.508, 20.492, 77.008, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        saddleNop32 ++
        sectionCounter --;
    }

    //  Saddle Sitiched  //  make another half 16 pages number table 
    if(saddleNop16 > 0){
        makeText(document.pages.item(sectionCounter), [39.508, 57.992, 77.008, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [39.508, 95.492, 77.008, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [119.056, 95.492, 156.556, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [119.056, 57.992, 156.556, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [119.056, 20.492, 156.556, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        if(startContents <= endPage){
                makeText(document.pages.item(sectionCounter), [119.056, 132.992, 156.556, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
                startContents ++;
        }else{
                makeBlueText(document.pages.item(sectionCounter), [119.056, 132.992, 156.556, 159.508], "Blank", "Times New Roman", 25, Justification.centerAlign, false, 90);
                blank --;
        }
        if(startContents <= endPage){
                makeText(document.pages.item(sectionCounter), [39.508, 132.992, 77.008, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
                startContents ++;
        }else{
                makeBlueText(document.pages.item(sectionCounter), [39.508, 132.992, 77.008, 159.508], "Blank", "Times New Roman", 25, Justification.centerAlign, false, 90);
                blank --;
        }
        if(startContents <= endPage){
                makeText(document.pages.item(sectionCounter), [39.508, 20.492, 77.008, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
                startContents ++;
        }else{
                makeBlueText(document.pages.item(sectionCounter), [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, Justification.centerAlign, false, -90);
                blank --;
        }
        sectionCounter --;
    }

    //  Saddle Sitiched  //  make another half 12 pages number table
    if(saddleNop12 > 0){
        makeText(document.pages.item(sectionCounter), [39.508, 57.992, 77.008, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [196.96, 57.992, 234.46, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [196.96, 95.492, 234.46, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        if(startContents <= endPage){
                makeText(document.pages.item(sectionCounter), [39.508, 95.492, 77.008, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
                startContents ++;
        }else{
                makeBlueText(document.pages.item(sectionCounter), [39.508, 95.492, 77.008, 122.008], "Blank", "Times New Roman", 25, Justification.centerAlign, false, -90);
                blank --;
        }        
        if(startContents <= endPage){
                makeText(document.pages.item(sectionCounter), [196.96, 20.492, 234.46, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
                startContents ++;
        }else{
                makeBlueText(document.pages.item(sectionCounter), [196.96, 20.492, 234.46, 47.008], "Blank", "Times New Roman", 25, Justification.centerAlign, false, -90);
                blank --;
        }
        if(startContents <= endPage){
                makeText(document.pages.item(sectionCounter), [39.508, 20.492, 77.008, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
                startContents ++;
        }else{
                makeBlueText(document.pages.item(sectionCounter), [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, Justification.centerAlign, false, -90);
                blank --;
        }        
        sectionCounter --;
    }

    //  Saddle Sitiched  //  make another half 8 pages number table
    if(saddleNop8 > 0){
        makeText(document.pages.item(sectionCounter), [39.508, 57.992, 77.008, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        if(startContents <= endPage){
                makeText(document.pages.item(sectionCounter), [196.96, 57.992, 234.46, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
                startContents ++;
        }else{
                makeBlueText(document.pages.item(sectionCounter), [196.96, 57.992, 234.46, 84.508], "Blank", "Times New Roman", 25, Justification.centerAlign, false, 90);
                blank --;
        }        
        if(startContents <= endPage){
                makeText(document.pages.item(sectionCounter), [196.96, 20.492, 234.46, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
                startContents ++;
        }else{
                makeBlueText(document.pages.item(sectionCounter), [196.96, 20.492, 234.46, 47.008], "Blank", "Times New Roman", 25, Justification.centerAlign, false, -90);
                blank --;
        }
        if(startContents <= endPage){
                makeText(document.pages.item(sectionCounter), [39.508, 20.492, 77.008, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
                startContents ++;
        }else{
                makeBlueText(document.pages.item(sectionCounter), [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, Justification.centerAlign, false, -90);
                blank --;
        }        
        sectionCounter --;
    }

    //  Saddle Sitiched  //  make another half 4 pages number table
    if(saddleNop4 > 0){
        if(startContents <= endPage){
                makeText(document.pages.item(sectionCounter), [39.508, 57.992, 77.008, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
                startContents ++;
        }else{
                makeBlueText(document.pages.item(sectionCounter), [39.508, 57.992, 77.008, 84.508], "Blank", "Times New Roman", 25, Justification.centerAlign, false, 90);
                blank --;
        }
        if(startContents <= endPage){
                makeText(document.pages.item(sectionCounter), [39.508, 20.492, 77.008, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
                startContents ++;
        }else{
                makeBlueText(document.pages.item(sectionCounter), [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, Justification.centerAlign, false, -90);
                blank --;
        }        
        sectionCounter --;
    }

    //  Saddle Sitiched  //  make another half 32 pages number table
    while(saddleNop32 <= nop32 && saddleNop32 > 1){
        makeText(document.pages.item(sectionCounter), [39.508, 57.992, 77.008, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [249.992, 57.992, 287.492, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [249.992, 95.492, 287.492, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [39.508, 95.492, 77.008, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [119.056, 95.492, 156.556, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [170.444, 95.492, 207.944, 122.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [170.444, 57.992, 207.944, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [119.056, 57.992, 156.556, 84.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [119.056, 20.492, 156.556, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [170.444, 20.492, 207.944, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [170.444, 132.992, 207.944, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [119.056, 132.992, 156.556, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        makeText(document.pages.item(sectionCounter), [39.508, 132.992, 77.008, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
        startContents ++;
        if(startContents <= endPage){
                makeText(document.pages.item(sectionCounter), [249.992, 132.992, 287.492, 159.508], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, 90);
                startContents ++;
        }else{
                makeBlueText(document.pages.item(sectionCounter), [249.992, 132.992, 287.492, 159.508], "Blank", "Times New Roman", 25, Justification.centerAlign, false, 90);
                blank --;
        }
        if(startContents <= endPage){
                makeText(document.pages.item(sectionCounter), [249.992, 20.492, 287.492, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
                startContents ++;
        }else{
                makeBlueText(document.pages.item(sectionCounter), [249.992, 20.492, 287.492, 47.008], "Blank", "Times New Roman", 25, Justification.centerAlign, false, -90);
                blank --;
        }
        if(startContents <= endPage){
                makeText(document.pages.item(sectionCounter), [39.508, 20.492, 77.008, 47.008], "↑\n\n" + startContents, "Times New Roman", 25, Justification.centerAlign, false, -90);
                startContents ++;
        }else{
                makeBlueText(document.pages.item(sectionCounter), [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, Justification.centerAlign, false, -90);
                blank --;
        }        
        saddleNop32 --;
        sectionCounter --;
    }
}









// Validate the dialog
function coverValidator(){
    pageValidator(dialog.cover, placementINFO.pgCount, "Cover");
}
function startContentsValidator(){
    pageValidator(dialog.startContents, placementINFO.pgCount, "Start Contents");
}
function endContentsValidator(){
    pageValidator(dialog.endContents, placementINFO.pgCount, "End Contents");
}
function startPageValidator(){
    pageValidator(dialog.startPage, placementINFO.pgCount, "Start Page");
}
function endPageValidator(){
    pageValidator(dialog.endPage, placementINFO.pgCount, "End Page");
}
function kindValidator(){
    kind(dialog.startPage, placementINFO.pgCount, "類型");
}
function jobNameValidator(){
    jobNameValidator(dialog.jobName, placementINFO.pgCount, "Job Name");
}

// Take care of OK being clicked
function onOKclicked(){
    dialog.close(1);
}

// Take care of Cancel being clicked
function onCANclicked(){
    dialog.close(0);
}

// Make text frame
function makeText(myPage, myBounds, myString, myFontName, myPointSize, myJustification, myFitToContent, myTransfrom, sectionCounter){
    var textFrame = myPage.textFrames.add({geometricBounds:myBounds});
    textFrame.transform (CoordinateSpaces.pasteboardCoordinates, AnchorPoint.centerAnchor, app.transformationMatrices.add({counterclockwiseRotationAngle: myTransfrom}));    
    textFrame.texts.item(0).insertionPoints.item(0).contents = myString
    textFrame.texts.item(0).parentStory.appliedFont = app.fonts.item(myFontName);
    textFrame.texts.item(0).parentStory.pointSize = myPointSize;
    textFrame.texts.item(0).justification = myJustification;
    textFrame.textFramePreferences.verticalJustification = VerticalJustification.CENTER_ALIGN;
    sectionCounter = 1;
    if(myFitToContent == true){
        textFrame.fit(FitOptions.frameToContent);
    }
    return textFrame;
}

function makeBlueText(myPage, myBounds, myString, myFontName, myPointSize, myJustification, myFitToContent, myTransfrom, myFontColor, sectionCounter){
    var textFrame = myPage.textFrames.add({geometricBounds:myBounds});
    textFrame.transform (CoordinateSpaces.pasteboardCoordinates, AnchorPoint.centerAnchor, app.transformationMatrices.add({counterclockwiseRotationAngle: myTransfrom}));    
    textFrame.texts.item(0).insertionPoints.item(0).contents = myString
    textFrame.texts.item(0).parentStory.appliedFont = app.fonts.item(myFontName);
    textFrame.texts.item(0).parentStory.pointSize = myPointSize;
    textFrame.texts.item(0).justification = myJustification;
    textFrame.textFramePreferences.verticalJustification = VerticalJustification.CENTER_ALIGN;
    textFrame.texts.item(0).fillColor = "Blue";
    sectionCounter = 1;
    if(myFitToContent == true){
        textFrame.fit(FitOptions.frameToContent);
    }
    return textFrame;
}











//  trash
/*
if(app.documents.length == 0){
    var tmp = new Array();
    tmp["width"] = 612;
    tmp["height"] = 792;
    placementINFO["pgSize"]  = tmp;
}
*/