var type;
var nop32 = 0, nop16 = 0, nop12 = 0, nop8 = 0, nop4 = 0;
var saddleNop32 = 0, saddleNop16 = 0, saddleNop12 = 0, saddleNop8 = 0, saddleNop4 = 0;
var blank = 0;
var jobName;
var totalPage;
var startContents, endContents;
var startpage1, startpage2, startpage3, startpage4, startpage5, startpage6, startpage7, startpage8, startpage9, startpage10, startpage11, startpage12, startpage13, startpage14, startpage15;
var endPage1, endPage2, endPage3, endPage4, endPage5, endPage6, endPage7, endPage8, endPage9, endPage10, endPage11, endPage12, endPage13, endPage14, endPage15;
var blank1, blank2, blank3, blank4, blank5, blank6, blank7, blank8, blank9, blank10, blank11, blank12, blank13, blank14, blank15;
var sectionCounter;
var layoutCounter = 1;
var cover;
var imprint;
var content;
var text;
var 類型 = ["大報", "憲報 No. 1", "憲報 No. 2", "憲報 No. 3", "憲報 No. 4", "憲報 No. 5", "憲報 No. 6", "特報", "簽字", "No. 1 + Index A", "No. 2 + Index B", "No. 3 + Index C", "Sup 1 + Index A", "Sup 2 + Index B", "Sup 3 + Index C"];


app.documents.item(0).viewPreferences.horizontalMeasurementUnits = MeasurementUnits.millimeters;
app.documents.item(0).viewPreferences.verticalMeasurementUnits = MeasurementUnits.millimeters;
document.marginPreferences.top = 0;
document.marginPreferences.left = 0;
document.marginPreferences.bottom = 0;
document.marginPreferences.right = 0;
document.documentPreferences.pageWidth = 210;
document.documentPreferences.pageHeight = 297;



var tag = ["", "A", "B", "C", "D", "E",];
if(app.documents.length === 0){
    var tmp = new Array();
    tmp["width"] = 612;
    tmp["height"] = 792;
    placementINFO["pgSize"]  = tmp;
}   

//  declare dialog data information
makeDialog();
dialog.center(); 
if(dialog.show() === 1){
    jobName = String(dialog.jobName.text);
    cover = Number(dialog.cover.text);
    imprint = Number(dialog.imprint.text);
    startContents = Number(dialog.startContents.text);
    endContents = Number(dialog.endContents.text);
    startPage1 = Number(dialog.startPage1.text);
    endPage1 = Number(dialog.endPage1.text);
    startPage2 = Number(dialog.startPage2.text);
    endPage2 = Number(dialog.endPage2.text);
    startPage3 = Number(dialog.startPage3.text);
    endPage3 = Number(dialog.endPage3.text);
    startPage4 = Number(dialog.startPage4.text);
    endPage4 = Number(dialog.endPage4.text);
    startPage5 = Number(dialog.startPage5.text);
    endPage5 = Number(dialog.endPage5.text);
    startPage6 = Number(dialog.startPage6.text);
    endPage6 = Number(dialog.endPage6.text);
    startPage7 = Number(dialog.startPage7.text);
    endPage7 = Number(dialog.endPage7.text);
    startPage8 = Number(dialog.startPage8.text);
    endPage8 = Number(dialog.endPage8.text);
    startPage9 = Number(dialog.startPage9.text);
    endPage9 = Number(dialog.endPage9.text);
    startPage10 = Number(dialog.startPage10.text);
    endPage10 = Number(dialog.endPage10.text);
/*  
    startPage11 = Number(dialog.startPage11.text);
    endPage11 = Number(dialog.endPage11.text);
    startPage12 = Number(dialog.startPage12.text);
    endPage12 = Number(dialog.endPage12.text);
    startPage13 = Number(dialog.startPage13.text);
    endPage13 = Number(dialog.endPage13.text);
    startPage14 = Number(dialog.startPage14.text);
    endPage14 = Number(dialog.endPage14.text);
    startPage15 = Number(dialog.startPage15.text);
    endPage15 = Number(dialog.endPage15.text);
*/    
    type = dialog.kindType.selection.index;
}else{
    exit();
}


//  32 page section information
makeText(document.masterSpreads.item(1), [10, 15, 30, 165], 類型[type] + " 32 Pages " + jobName, "新細明體", 20, "Blue", Justification.leftAlign, VerticalJustification.TOP_ALIGN, 0, false);
//  16 page section information
makeText(document.masterSpreads.item(2), [10, 15, 30, 165], 類型[type] + " 16 Pages " + jobName, "新細明體", 20, "Blue", Justification.leftAlign, VerticalJustification.TOP_ALIGN, 0, false);
//  12 page section information
makeText(document.masterSpreads.item(3), [10, 15, 30, 165], 類型[type] + " 12 Pages " + jobName, "新細明體", 20, "Blue", Justification.leftAlign, VerticalJustification.TOP_ALIGN, 0, false);
//  8 page section information
makeText(document.masterSpreads.item(4), [10, 15, 30, 165], 類型[type] + " 8 Pages " + jobName, "新細明體", 20, "Blue", Justification.leftAlign, VerticalJustification.TOP_ALIGN, 0, false);
//  4 page section information
makeText(document.masterSpreads.item(5), [10, 15, 30, 165], 類型[type] + " 4 Pages " + jobName, "新細明體", 20, "Blue", Justification.leftAlign, VerticalJustification.TOP_ALIGN, 0, false);



// Create dialog box
function makeDialog()
{
    dialog = new Window('dialog', "HAHAHA", "x:0, y:0, width:550, height:690");
    dialog.panel = dialog.add('panel', [10, 10, 300, 680], "");

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
    dialog.cover = dialog.panel.add('edittext', [100, 90, 155, 115], "");
    dialog.cover.onChange = coverValidator;
    
    dialog.panel.add('statictext',  [10, 135, 60, 150], "Imprint："); 
    dialog.imprint = dialog.panel.add('edittext', [100, 130, 155, 155], "");
    dialog.imprint.onChange = imprintValidator;

    dialog.panel.add('statictext',  [10, 175, 50, 195], "目錄："); 
    dialog.startContents = dialog.panel.add('edittext', [100, 170, 155, 195], "");
    dialog.startContents.onChange = startContentsValidator;
    dialog.panel.add('statictext',  [160, 175, 165, 200], "-");
    dialog.endContents = dialog.panel.add('edittext', [170, 170, 220, 195], "");
    dialog.cropType = dialog.panel.add('dropdownlist', [225, 170, 265, 195]);
    for(var i=0; i<tag.length; i++){
        dialog.cropType.add('item', tag[i]);
    }

    dialog.panel.add('statictext',  [10, 215, 50, 240], "內文："); 
    dialog.startPage1 = dialog.panel.add('edittext', [100, 210, 155, 235], "");
    dialog.startPage1.onChange = startPage1Validator;
    dialog.panel.add('statictext',  [160, 215, 165, 240], "-");
    dialog.endPage1 = dialog.panel.add('edittext', [170, 210, 220, 235], ""); 
    dialog.endPage1.onChange = endPage1Validator;      
    dialog.cropType = dialog.panel.add('dropdownlist', [225, 210, 265, 235]);
    for(var i=0; i<tag.length; i++){
        dialog.cropType.add('item', tag[i]);
    }

    dialog.panel.add('statictext',  [10, 255, 50, 290], "內文："); 
    dialog.startPage2 = dialog.panel.add('edittext', [100, 250, 155, 275], "");
    dialog.startPage2.onChange = startPage2Validator;
    dialog.panel.add('statictext',  [160, 255, 165, 290], "-");
    dialog.endPage2 = dialog.panel.add('edittext', [170, 250, 220, 275], ""); 
    dialog.endPage2.onChange = endPage2Validator;      
    dialog.cropType = dialog.panel.add('dropdownlist', [225, 250, 265, 275]);
    for(var i=0; i<tag.length; i++){
        dialog.cropType.add('item', tag[i]);
    }

    dialog.panel.add('statictext',  [10, 295, 50, 310], "內文："); 
    dialog.startPage3 = dialog.panel.add('edittext', [100, 290, 155, 315], "");
    dialog.startPage3.onChange = startPage3Validator;
    dialog.panel.add('statictext',  [160, 295, 165, 310], "-");
    dialog.endPage3 = dialog.panel.add('edittext', [170, 290, 220, 315], ""); 
    dialog.endPage3.onChange = endPage3Validator;      
    dialog.cropType = dialog.panel.add('dropdownlist', [225, 290, 265, 275]);
    for(var i=0; i<tag.length; i++){
        dialog.cropType.add('item', tag[i]);
    }

    dialog.panel.add('statictext',  [10, 335, 50, 350], "內文："); 
    dialog.startPage4 = dialog.panel.add('edittext', [100, 330, 155, 355], "");
    dialog.startPage4.onChange = startPage4Validator;
    dialog.panel.add('statictext',  [160, 335, 165, 350], "-");
    dialog.endPage4 = dialog.panel.add('edittext', [170, 330, 220, 355], ""); 
    dialog.endPage4.onChange = endPage4Validator;      
    dialog.cropType = dialog.panel.add('dropdownlist', [225, 330, 265, 315]);
    for(var i=0; i<tag.length; i++){
        dialog.cropType.add('item', tag[i]);
    }

    dialog.panel.add('statictext',  [10, 375, 50, 390], "內文："); 
    dialog.startPage5 = dialog.panel.add('edittext', [100, 370, 155, 395], "");
    dialog.startPage5.onChange = startPage5Validator;
    dialog.panel.add('statictext',  [160, 375, 165, 390], "-");
    dialog.endPage5 = dialog.panel.add('edittext', [170, 370, 220, 395], ""); 
    dialog.endPage5.onChange = endPage5Validator;      
    dialog.cropType = dialog.panel.add('dropdownlist', [225, 370, 265, 355]);
    for(var i=0; i<tag.length; i++){
        dialog.cropType.add('item', tag[i]);
    }

    dialog.panel.add('statictext',  [10, 415, 50, 430], "內文："); 
    dialog.startPage6 = dialog.panel.add('edittext', [100, 410, 155, 435], "");
    dialog.startPage6.onChange = startPage6Validator;
    dialog.panel.add('statictext',  [160, 415, 165, 430], "-");
    dialog.endPage6 = dialog.panel.add('edittext', [170, 410, 220, 435], ""); 
    dialog.endPage6.onChange = endPage6Validator;      
    dialog.cropType = dialog.panel.add('dropdownlist', [225, 410, 265, 395]);
    for(var i=0; i<tag.length; i++){
        dialog.cropType.add('item', tag[i]);
    }

    dialog.panel.add('statictext',  [10, 455, 50, 470], "內文："); 
    dialog.startPage7 = dialog.panel.add('edittext', [100, 450, 155, 475], "");
    dialog.startPage7.onChange = startPage7Validator;
    dialog.panel.add('statictext',  [160, 455, 165, 470], "-");
    dialog.endPage7 = dialog.panel.add('edittext', [170, 450, 220, 475], ""); 
    dialog.endPage7.onChange = endPage7Validator;      
    dialog.cropType = dialog.panel.add('dropdownlist', [225, 450, 265, 335]);
    for(var i=0; i<tag.length; i++){
        dialog.cropType.add('item', tag[i]);
    }

    dialog.panel.add('statictext',  [10, 495, 50, 510], "內文："); 
    dialog.startPage8 = dialog.panel.add('edittext', [100, 490, 155, 515], "");
    dialog.startPage8.onChange = startPage8Validator;
    dialog.panel.add('statictext',  [160, 495, 165, 510], "-");
    dialog.endPage8 = dialog.panel.add('edittext', [170, 490, 220, 515], ""); 
    dialog.endPage8.onChange = endPage8Validator;      
    dialog.cropType = dialog.panel.add('dropdownlist', [225, 490, 265, 375]);
    for(var i=0; i<tag.length; i++){
        dialog.cropType.add('item', tag[i]);
    }

    dialog.panel.add('statictext',  [10, 535, 50, 550], "內文："); 
    dialog.startPage9 = dialog.panel.add('edittext', [100, 530, 155, 555], "");
    dialog.startPage9.onChange = startPage9Validator;
    dialog.panel.add('statictext',  [160, 535, 165, 550], "-");
    dialog.endPage9 = dialog.panel.add('edittext', [170, 530, 220, 555], ""); 
    dialog.endPage9.onChange = endPage9Validator;      
    dialog.cropType = dialog.panel.add('dropdownlist', [225, 530, 265, 415]);
    for(var i=0; i<tag.length; i++){
        dialog.cropType.add('item', tag[i]);
    }

    dialog.panel.add('statictext',  [10, 575, 50, 590], "內文："); 
    dialog.startPage10 = dialog.panel.add('edittext', [100, 570, 155, 595], "");
    dialog.startPage10.onChange = startPage10Validator;
    dialog.panel.add('statictext',  [160, 575, 165, 590], "-");
    dialog.endPage10 = dialog.panel.add('edittext', [170, 570, 220, 595], ""); 
    dialog.endPage10.onChange = endPage10Validator;      
    dialog.cropType = dialog.panel.add('dropdownlist', [225, 570, 265, 455]);
    for(var i=0; i<tag.length; i++){
        dialog.cropType.add('item', tag[i]);
    }

    dialog.placeOnLayer = dialog.panel.add('checkbox', [10, 600, 220, 635], "\"This\" Page");
    dialog.placeOnLayer = dialog.panel.add('checkbox', [10, 640, 220, 655], "Holiday");
    
    // The buttons
    dialog.OKbut = dialog.add('button', [450, 20, 525, 45], "OK");
    dialog.OKbut.onClick = onOKclicked;
    dialog.CANbut = dialog.add('button', [450, 50, 525, 75], "Cancel");
    dialog.CANbut.onClick = onCANclicked;

    return dialog;
}

//  get total page number
//contents = endContents - startContents + 1;
//contentsCounter = startContents;


//  make sectionCounter = total page number to count sections
sectionCounter = endPage15 - startContents + 1;

//  count 32 pages section
for(; sectionCounter >= 32;){
    nop32 ++;
    sectionCounter = sectionCounter - 32;
}
if(sectionCounter === 31){
    nop32 ++;
    sectionCounter = sectionCounter - 32;
    blank = blank + 1;
}else if(sectionCounter === 30){
    nop32 ++;
    sectionCounter = sectionCounter - 32;
    blank = blank + 2;
}else if(sectionCounter === 29){
    nop32 ++;
    sectionCounter = sectionCounter - 32;
    blank = blank + 3;
}

//  count 16 pages section
if(sectionCounter >= 16){
    nop16 ++;
    sectionCounter = sectionCounter - 16;
}else if(sectionCounter === 15){
    nop16 ++;
    sectionCounter = sectionCounter - 16;
    blank = blank + 1;
}else if(sectionCounter === 14){
    nop16 ++;
    sectionCounter = sectionCounter - 16;
    blank = blank + 2;
}else if(sectionCounter === 13){
    nop16 ++;
    sectionCounter = sectionCounter - 16;
    blank = blank + 3;
}

//  count 12 pages section
if(sectionCounter >= 12){
    nop12 ++;
    sectionCounter = sectionCounter - 12;
}else if(sectionCounter === 11){
    nop12 ++;
    sectionCounter = sectionCounter - 12;
    blank = blank + 1;
}else if(sectionCounter === 10){
    nop12 ++;
    sectionCounter = sectionCounter - 12;
    blank = blank + 2;
}else if(sectionCounter === 9){
    nop12 ++;
    sectionCounter = sectionCounter - 12;
    blank = blank + 3;
}

//  count 8 pages section
if(sectionCounter >= 8){
    nop8 ++;
    sectionCounter = sectionCounter - 8;
}else if(sectionCounter === 7){
    nop8 ++;
    sectionCounter = sectionCounter - 8;
    blank = blank + 1;
}else if(sectionCounter === 6){
    nop8 ++;
    sectionCounter = sectionCounter - 8;
    blank = blank + 2;
}else if(sectionCounter === 5){
    nop8 ++;
    sectionCounter = sectionCounter - 8;
    blank = blank + 3;
}

//  count 4 pages section
if(sectionCounter >= 4){
    nop4 ++;
    sectionCounter = sectionCounter - 4;
}else if(sectionCounter === 3){
    nop4 ++;
    sectionCounter = sectionCounter - 4;
    blank = blank + 1;
}else if(sectionCounter === 2){
    nop4 ++;
    sectionCounter = sectionCounter - 4;
    blank = blank + 2;
}else if(sectionCounter === 1){
    nop4 ++;
    sectionCounter = sectionCounter - 4;
    blank = blank + 3;
}

//  Make information page


//  Create number page header
document.pages.item(0).appliedMaster = app.documents[0].masterSpreads.item ('G-主版')
makeText(document.pages.item(0), [252.104, 24.7 , 263.736, 29.7], "" + nop32, "Times New Roman", 13, "Red", Justification.centerAlign, VerticalJustification.BOTTOM_ALIGN, 0, false);
makeText(document.pages.item(0), [252.104, 72.2, 263.736, 77.2], "" + nop16, "Times New Roman", 13, "Red",  Justification.centerAlign, VerticalJustification.BOTTOM_ALIGN, 0, false);
makeText(document.pages.item(0), [252.104, 119.7 , 263.736, 124.7], "" + nop12, "Times New Roman", 13, "Red",  Justification.centerAlign, VerticalJustification.BOTTOM_ALIGN, 0, false);
makeText(document.pages.item(0), [252.104, 167.2, 263.736, 172.2], "" + nop8, "Times New Roman", 13, "Red",  Justification.centerAlign, VerticalJustification.BOTTOM_ALIGN, 0, false);
makeText(document.pages.item(0), [263.736, 24.7, 275.411, 29.7], "" + nop4, "Times New Roman", 13, "Red",  Justification.centerAlign, VerticalJustification.BOTTOM_ALIGN, 0, false);
if (totalPage > 135){
    makeText(document.pages.item(0), [275.411 , 24.7, 287.086, 29.7], "1", "Times New Roman", 13, "Red",  Justification.centerAlign, VerticalJustification.BOTTOM_ALIGN, 0, false);
}
if (類型[type] === "特報"  &&  nop4>0 && totalPage<5){
    makeText(document.pages.item(0), [263.736, 43.6, 275.411, 200 ], "→　Flat Work \\\\ gaz sm33_4pp pf Logo single", "Times New Roman", 13, "Red",  Justification.leftAlign, VerticalJustification.BOTTOM_ALIGN, 0, false);
}else if (類型[type]=="Sup 1 + Index A" || 類型[type]=="Sup 2 + Index B" || 類型[type]=="Sup 3 + Index C"  &&  nop4>0 && totalPage<5){
    makeText(document.pages.item(0), [263.736, 43.6, 275.411, 200 ], "→　Flat Work \\\\ gaz sm33_4pp pf W/B single", "Times New Roman", 13, "Red",  Justification.leftAlign, VerticalJustification.BOTTOM_ALIGN, 0, false);
}



//  Pefect Bound
if (totalPage > 135){
//  make 32 pages number table
    for(; nop32 > 1;){
    nop32 --;
        document.pages.add().appliedMaster = app.documents[0].masterSpreads.item ('B-主版');

        makeText(document.pages.item(layoutCounter), [66.024, 20.492, 103.524, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [223.476, 20.492, 260.976, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [223.476, 132.992, 260.976, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [66.024, 132.992, 103.524, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [92.54, 132.992, 130.04, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [196.96, 132.992, 234.46, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [196.96, 20.492, 234.46, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [92.54, 20.492, 130.04, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
         makeText(document.pages.item(layoutCounter), [92.54, 57.992, 130.04, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [196.96, 57.992, 234.46, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [196.96, 95.492, 234.46, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [92.54, 95.492, 130.04, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [66.024, 95.492, 103.524, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [223.476, 95.492, 260.976, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [223.476, 57.992, 260.976, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [66.024, 57.992, 103.524, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [39.508, 57.992, 77.008, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [249.992, 57.992, 287.492, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [249.992, 95.492, 287.492, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [39.508, 95.492, 77.008, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [119.056, 95.492, 156.556, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [170.444, 95.492, 207.944, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [170.444, 57.992, 207.944, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [119.056, 57.992, 156.556, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [119.056, 20.492, 156.556, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [170.444, 20.492, 207.944, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [170.444, 132.992, 207.944, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [119.056, 132.992, 156.556, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [39.508, 132.992, 77.008, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [249.992, 132.992, 287.492, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [249.992, 20.492, 287.492, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        layoutCounter ++;
    }

    //  make 4 pages number table
    for(; nop4 > 0;){
        nop4 --;
        document.pages.add().appliedMaster = app.documents[0].masterSpreads.item ('F-主版');

        makeText(document.pages.item(layoutCounter), [66.024, 20.492, 103.524, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [66.024, 57.992, 103.524, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [66.024, 57.992, 103.524, 84.508], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                blank --;
        }        
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [39.508, 57.992, 77.008, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [39.508, 57.992, 77.008, 84.508], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                blank --;
        }
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                blank --;
        }        
        layoutCounter ++;
    }

    //  make 8 pages number table
    for(; nop8 > 0;){
        nop8 --;
        document.pages.add().appliedMaster = app.documents[0].masterSpreads.item ('E-主版');

        makeText(document.pages.item(layoutCounter), [66.024, 20.492, 103.524, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [170.444, 20.492, 207.944, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [170.444, 57.992, 207.944, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [66.024, 57.992, 103.524, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [39.508, 57.992, 77.008, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [196.96, 57.992, 234.46, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [196.96, 57.992, 234.46, 84.508], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                blank --;
        }        
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [196.96, 20.492, 234.46, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [196.96, 20.492, 234.46, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                blank --;
        }
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                blank --;
        }        
        layoutCounter ++;
    }

    //  make 12 pages number table
    for(; nop12 > 0;){
        nop12 --;
        document.pages.add().appliedMaster = app.documents[0].masterSpreads.item ('D-主版');

        makeText(document.pages.item(layoutCounter), [66.024, 20.492, 103.524, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [170.444, 20.492, 207.944, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [66.024, 95.492, 103.524, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [170.44, 95.492, 207.944, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [170.444, 57.992, 207.944, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [66.024, 57.992, 103.524, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [39.508, 57.992, 77.008, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [196.96, 57.992, 234.46, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [196.96, 95.492, 234.46, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [39.508, 95.492, 77.008, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [39.508, 95.492, 77.008, 122.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                blank --;
        }        
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [196.96, 20.492, 234.46, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [196.96, 20.492, 234.46, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                blank --;
        }
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                blank --;
        }        
        layoutCounter ++;
    }

    //  make 16 pages number table
    for(; nop16 > 0;){
        nop16 --;
        document.pages.add().appliedMaster = app.documents[0].masterSpreads.item ('C-主版');

        makeText(document.pages.item(layoutCounter), [66.024, 20.492, 103.524, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [66.024, 132.992, 103.524, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [92.54, 132.992, 130.04, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [92.54, 20.492, 130.04, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [92.54, 57.992, 130.04, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [92.54, 95.492, 130.04, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [66.024, 95.492, 103.524, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [66.024, 57.992, 103.524, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [39.508, 57.992, 77.008, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [39.508, 95.492, 77.008, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [119.056, 95.492, 156.556, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [119.056, 57.992, 156.556, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [119.056, 20.492, 156.556, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [119.056, 132.992, 156.556, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [119.056, 132.992, 156.556, 159.508], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                blank --;
        }
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [39.508, 132.992, 77.008, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [39.508, 132.992, 77.008, 159.508], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                blank --;
        }
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                blank --;
        }
        layoutCounter ++;
    }

    //  make the last 32 pages number table
    for(; nop32 > 0;){
        nop32 --;
        document.pages.add().appliedMaster = app.documents[0].masterSpreads.item ('B-主版');

        if(startContents )
        makeText(document.pages.item(layoutCounter), [66.024, 20.492, 103.524, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [223.476, 20.492, 260.976, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [223.476, 132.992, 260.976, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [66.024, 132.992, 103.524, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [92.54, 132.992, 130.04, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [196.96, 132.992, 234.46, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [196.96, 20.492, 234.46, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [92.54, 20.492, 130.04, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
         makeText(document.pages.item(layoutCounter), [92.54, 57.992, 130.04, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [196.96, 57.992, 234.46, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [196.96, 95.492, 234.46, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [92.54, 95.492, 130.04, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [66.024, 95.492, 103.524, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [223.476, 95.492, 260.976, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [223.476, 57.992, 260.976, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [66.024, 57.992, 103.524, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [39.508, 57.992, 77.008, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [249.992, 57.992, 287.492, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [249.992, 95.492, 287.492, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [39.508, 95.492, 77.008, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [119.056, 95.492, 156.556, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [170.444, 95.492, 207.944, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [170.444, 57.992, 207.944, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [119.056, 57.992, 156.556, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [119.056, 20.492, 156.556, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [170.444, 20.492, 207.944, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [170.444, 132.992, 207.944, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [119.056, 132.992, 156.556, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [39.508, 132.992, 77.008, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [249.992, 132.992, 287.492, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [249.992, 132.992, 287.492, 159.508], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                blank --;
        }
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [249.992, 20.492, 287.492, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [249.992, 20.492, 287.492, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                blank --;
        }
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                blank --;
        }        
        layoutCounter ++;
    }



//  Saddle Sitiched  //  make 32 pages number table
}else{
        for(; saddleNop32 < (nop32 - 1);){
        document.pages.add().appliedMaster = app.documents[0].masterSpreads.item ('B-主版');

        makeText(document.pages.item(layoutCounter), [66.024, 20.492, 103.524, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [223.476, 20.492, 260.976, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [223.476, 132.992, 260.976, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [66.024, 132.992, 103.524, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [92.54, 132.992, 130.04, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [196.96, 132.992, 234.46, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [196.96, 20.492, 234.46, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [92.54, 20.492, 130.04, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [92.54, 57.992, 130.04, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [196.96, 57.992, 234.46, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [196.96, 95.492, 234.46, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [92.54, 95.492, 130.04, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [66.024, 95.492, 103.524, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [223.476, 95.492, 260.976, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [223.476, 57.992, 260.976, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [66.024, 57.992, 103.524, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        saddleNop32 ++;
        layoutCounter ++;
    }

    //  Saddle Sitiched  //  make 4 pages number table
    if(saddleNop4 < nop4){
        document.pages.add().appliedMaster = app.documents[0].masterSpreads.item ('F-主版');

        makeText(document.pages.item(layoutCounter), [66.024, 20.492, 103.524, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [66.024, 57.992, 103.524, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        saddleNop4 ++;
        if(nop32 > 0){
            layoutCounter ++;
        }
    }

    //  Saddle Sitiched  //  make 8 pages number table
    if(saddleNop8 < nop8){
        document.pages.add().appliedMaster = app.documents[0].masterSpreads.item ('E-主版');

        makeText(document.pages.item(layoutCounter), [66.024, 20.492, 103.524, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [170.444, 20.492, 207.944, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [170.444, 57.992, 207.944, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [66.024, 57.992, 103.524, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        saddleNop8 ++;
        if(nop32 > 0){
            layoutCounter ++;
        }
    }

    //  Saddle Sitiched  //  make 12 pages number table
    if(saddleNop12 < nop12){
        document.pages.add().appliedMaster = app.documents[0].masterSpreads.item ('D-主版');

        makeText(document.pages.item(layoutCounter), [66.024, 20.492, 103.524, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [170.444, 20.492, 207.944, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [66.024, 95.492, 103.524, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [170.44, 95.492, 207.944, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [170.444, 57.992, 207.944, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [66.024, 57.992, 103.524, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        saddleNop12 ++;
        if(nop32 > 0){
            layoutCounter ++;
        }
    }

    //  Saddle Sitiched  //  make 16 pages number table
    if(saddleNop16 < nop16){
        document.pages.add().appliedMaster = app.documents[0].masterSpreads.item ('C-主版');

        makeText(document.pages.item(layoutCounter), [66.024, 20.492, 103.524, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [66.024, 132.992, 103.524, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [92.54, 132.992, 130.04, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [92.54, 20.492, 130.04, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [92.54, 57.992, 130.04, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [92.54, 95.492, 130.04, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [66.024, 95.492, 103.524, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [66.024, 57.992, 103.524, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        saddleNop16 ++;
        if(nop32 > 0){
            layoutCounter ++;
        }
    }

    //  Saddle Sitiched  //  make middle 32 pages number table
    if(saddleNop32 === nop32 - 1){
        document.pages.add().appliedMaster = app.documents[0].masterSpreads.item ('B-主版');

        makeText(document.pages.item(layoutCounter), [66.024, 20.492, 103.524, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [223.476, 20.492, 260.976, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [223.476, 132.992, 260.976, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [66.024, 132.992, 103.524, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [92.54, 132.992, 130.04, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [196.96, 132.992, 234.46, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [196.96, 20.492, 234.46, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [92.54, 20.492, 130.04, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
         makeText(document.pages.item(layoutCounter), [92.54, 57.992, 130.04, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [196.96, 57.992, 234.46, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [196.96, 95.492, 234.46, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [92.54, 95.492, 130.04, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [66.024, 95.492, 103.524, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [223.476, 95.492, 260.976, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [223.476, 57.992, 260.976, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [66.024, 57.992, 103.524, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [39.508, 57.992, 77.008, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [249.992, 57.992, 287.492, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [249.992, 95.492, 287.492, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [39.508, 95.492, 77.008, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [119.056, 95.492, 156.556, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [170.444, 95.492, 207.944, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [170.444, 57.992, 207.944, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [119.056, 57.992, 156.556, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [119.056, 20.492, 156.556, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [170.444, 20.492, 207.944, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [170.444, 132.992, 207.944, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [119.056, 132.992, 156.556, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [39.508, 132.992, 77.008, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;

    if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [249.992, 132.992, 287.492, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [249.992, 132.992, 287.492, 159.508], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                blank --;
        }
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [249.992, 20.492, 287.492, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [249.992, 20.492, 287.492, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                blank --;
        }
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                blank --;
        }                
                
        saddleNop32 ++
        layoutCounter --;
    }

    //  Saddle Sitiched  //  make another half 16 pages number table 
    if(saddleNop16 > 0){
        makeText(document.pages.item(layoutCounter), [39.508, 57.992, 77.008, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [39.508, 95.492, 77.008, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [119.056, 95.492, 156.556, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [119.056, 57.992, 156.556, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [119.056, 20.492, 156.556, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [119.056, 132.992, 156.556, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [119.056, 132.992, 156.556, 159.508], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                blank --;
        }
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [39.508, 132.992, 77.008, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [39.508, 132.992, 77.008, 159.508], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                blank --;
        }
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                blank --;
        }
        layoutCounter --;
    }

    //  Saddle Sitiched  //  make another half 12 pages number table
    if(saddleNop12 > 0){
        makeText(document.pages.item(layoutCounter), [39.508, 57.992, 77.008, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [196.96, 57.992, 234.46, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [196.96, 95.492, 234.46, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [39.508, 95.492, 77.008, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [39.508, 95.492, 77.008, 122.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                blank --;
        }        
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [196.96, 20.492, 234.46, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [196.96, 20.492, 234.46, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                blank --;
        }
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                blank --;
        }        
        layoutCounter --;
    }

    //  Saddle Sitiched  //  make another half 8 pages number table
    if(saddleNop8 > 0){
        makeText(document.pages.item(layoutCounter), [39.508, 57.992, 77.008, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [196.96, 57.992, 234.46, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [196.96, 57.992, 234.46, 84.508], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                blank --;
        }        
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [196.96, 20.492, 234.46, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [196.96, 20.492, 234.46, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                blank --;
        }
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                blank --;
        }        
        layoutCounter --;
    }

    //  Saddle Sitiched  //  make another half 4 pages number table
    if(saddleNop4 > 0){
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [39.508, 57.992, 77.008, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [39.508, 57.992, 77.008, 84.508], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                blank --;
        }
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                blank --;
        }        
        layoutCounter --;
    }

    //  Saddle Sitiched  //  make another half 32 pages number table
    while(saddleNop32 <= nop32 && saddleNop32 > 1){
        makeText(document.pages.item(layoutCounter), [39.508, 57.992, 77.008, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [249.992, 57.992, 287.492, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [249.992, 95.492, 287.492, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [39.508, 95.492, 77.008, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [119.056, 95.492, 156.556, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [170.444, 95.492, 207.944, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [170.444, 57.992, 207.944, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [119.056, 57.992, 156.556, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [119.056, 20.492, 156.556, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [170.444, 20.492, 207.944, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [170.444, 132.992, 207.944, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [119.056, 132.992, 156.556, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(document.pages.item(layoutCounter), [39.508, 132.992, 77.008, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [249.992, 132.992, 287.492, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [249.992, 132.992, 287.492, 159.508], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                blank --;
        }
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [249.992, 20.492, 287.492, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [249.992, 20.492, 287.492, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                blank --;
        }
        if(startContents <= endPage10){
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                blank --;
        }        
        saddleNop32 --;
        layoutCounter --;
    }
}


//    Insert blank page
/*
if(startPage1 !== (endContents+1)){
            blank1 = endContents+1;
}
if(startPage2 !== (endPage1+1)){
            blank2 = endpage1+1;
}
if(startPage3 !== (endPage2+1)){
            blank3 = endpage2+1;
}
if(startPage4 !== (endPage3+1)){
            blank4 = endpage3+1;
}
if(startPage5 !== (endPage4+1)){
            blank5 = endpage4+1;
}
if(startPage6 !== (endPage5+1)){
            blank6 = endpage5+1;
}
if(startPage7 !== (endPage6+1)){
            blank7 = endpage6+1;
}
if(startPage8 !== (endPage7+1)){
            blank8 = endpage7+1;
}
if(startPage9 !== (endPage8+1)){
            blank9 = endpage8+1;
}
if(startPage10 !== (endPage9+1)){
            blank10 = endpage9+1;
}
if(startPage11 !== (endPage10+1)){
            blank11 = endpage10+1;
}
if(startPage12 !== (endPage11+1)){
            blank12 = endpage11+1;
}
if(startPage13 !== (endPage12+1)){
            blank13 = endpage12+1;
}
if(startPage14 !== (endPage13+1)){
            blank14 = endpage13+1;
}
if(startPage15 !== (endPage14+1)){
            blank15 = endpage14+1;
}
*/




// Validate the dialog
function coverValidator(){
    pageValidator(dialog.cover, placementINFO.pgCount, "Cover");
}
function imprintValidator(){
    pageValidator(dialog.imprint, placementINFO.pgCount, "Imprint");
}
function startContentsValidator(){
    pageValidator(dialog.startContents, placementINFO.pgCount, "Start Contents");
}
function endContentsValidator(){
    pageValidator(dialog.endContents, placementINFO.pgCount, "End Contents");
}
function startPage1Validator(){
    pageValidator(dialog.startPage1, placementINFO.pgCount, "Start Page");
}
function endPage1Validator(){
    pageValidator(dialog.endPage1, placementINFO.pgCount, "End Page");
}
function startPage2Validator(){
    pageValidator(dialog.startPage2, placementINFO.pgCount, "Start Page");
}
function endPage2Validator(){
    pageValidator(dialog.endPage2, placementINFO.pgCount, "End Page");
}
function startPage3Validator(){
    pageValidator(dialog.startPage3, placementINFO.pgCount, "Start Page");
}
function endPage3Validator(){
    pageValidator(dialog.endPage3, placementINFO.pgCount, "End Page");
}
function startPage4Validator(){
    pageValidator(dialog.startPage4, placementINFO.pgCount, "Start Page");
}
function endPage4Validator(){
    pageValidator(dialog.endPage4, placementINFO.pgCount, "End Page");
}
function startPage5Validator(){
    pageValidator(dialog.startPage5, placementINFO.pgCount, "Start Page");
}
function endPage5Validator(){
    pageValidator(dialog.endPage5, placementINFO.pgCount, "End Page");
}
function startPage6Validator(){
    pageValidator(dialog.startPage6, placementINFO.pgCount, "Start Page");
}
function endPage6Validator(){
    pageValidator(dialog.endPage6, placementINFO.pgCount, "End Page");
}
function startPage7Validator(){
    pageValidator(dialog.startPage7, placementINFO.pgCount, "Start Page");
}
function endPage7Validator(){
    pageValidator(dialog.endPage7, placementINFO.pgCount, "End Page");
}
function startPage8Validator(){
    pageValidator(dialog.startPage8, placementINFO.pgCount, "Start Page");
}
function endPage8Validator(){
    pageValidator(dialog.endPage8, placementINFO.pgCount, "End Page");
}
function startPage9Validator(){
    pageValidator(dialog.startPage9, placementINFO.pgCount, "Start Page");
}
function endPage9Validator(){
    pageValidator(dialog.endPage9, placementINFO.pgCount, "End Page");
}
function startPage10Validator(){
    pageValidator(dialog.startPage10, placementINFO.pgCount, "Start Page");
}
function endPage10Validator(){
    pageValidator(dialog.endPage10, placementINFO.pgCount, "End Page");
}
function startPage11Validator(){
    pageValidator(dialog.startPage11, placementINFO.pgCount, "Start Page");
}
function endPage11Validator(){
    pageValidator(dialog.endPage11, placementINFO.pgCount, "End Page");
}
function startPage12Validator(){
    pageValidator(dialog.startPage12, placementINFO.pgCount, "Start Page");
}
function endPage12Validator(){
    pageValidator(dialog.endPage12, placementINFO.pgCount, "End Page");
}
function startPage13Validator(){
    pageValidator(dialog.startPage13, placementINFO.pgCount, "Start Page");
}
function endPage13Validator(){
    pageValidator(dialog.endPage13, placementINFO.pgCount, "End Page");
}
function startPage14Validator(){
    pageValidator(dialog.startPage14, placementINFO.pgCount, "Start Page");
}
function endPage14Validator(){
    pageValidator(dialog.endPage14, placementINFO.pgCount, "End Page");
}
function startPage15Validator(){
    pageValidator(dialog.startPage15, placementINFO.pgCount, "Start Page");
}
function endPage15Validator(){
    pageValidator(dialog.endPage15, placementINFO.pgCount, "End Page");
}
function kindValidator(){
    kind(dialog.kindType, placementINFO.pgCount, "類型");
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
function makeText(myPage, myBounds, myString, myFontName, myPointSize, myColor, myJustification, myVerticalJustification, myTransfrom, layoutCounter, myFitToContent){
    var textFrame = myPage.textFrames.add({geometricBounds:myBounds});
    textFrame.texts.item(0).insertionPoints.item(0).contents = myString
    textFrame.texts.item(0).parentStory.appliedFont = app.fonts.item(myFontName);
    textFrame.texts.item(0).parentStory.pointSize = myPointSize;
    textFrame.texts.item(0).fillColor = myColor;
    textFrame.texts.item(0).justification = myJustification;
    textFrame.textFramePreferences.verticalJustification = myVerticalJustification;
    textFrame.transform (CoordinateSpaces.pasteboardCoordinates, AnchorPoint.centerAnchor, app.transformationMatrices.add({counterclockwiseRotationAngle: myTransfrom}));    
    layoutCounter = 1;
    if(myFitToContent === true){
        textFrame.fit(FitOptions.frameToContent);
    }
    return textFrame;
}








//  trash
/*
if(app.documents.length === 0){
    var tmp = new Array();
    tmp["width"] = 612;
    tmp["height"] = 792;
    placementINFO["pgSize"]  = tmp;
}



function makeText(myPage, myBounds, myString, myFontName, myPointSize, myJustification, myFitToContent, myTransfrom, layoutCounter){
    var textFrame = myPage.textFrames.add({geometricBounds:myBounds});
    textFrame.transform (CoordinateSpaces.pasteboardCoordinates, AnchorPoint.centerAnchor, app.transformationMatrices.add({counterclockwiseRotationAngle: myTransfrom}));    
    textFrame.texts.item(0).insertionPoints.item(0).contents = myString
    textFrame.texts.item(0).parentStory.appliedFont = app.fonts.item(myFontName);
    textFrame.texts.item(0).parentStory.pointSize = myPointSize;
    textFrame.texts.item(0).justification = myJustification;
    textFrame.textFramePreferences.verticalJustification = VerticalJustification.CENTER_ALIGN;
    layoutCounter = 1;
    if(myFitToContent === true){
        textFrame.fit(FitOptions.frameToContent);
    }
    return textFrame;
}



function makeBlueText(myPage, myBounds, myString, myFontName, myPointSize, myJustification, myFitToContent, myTransfrom, myFontColor, layoutCounter){
    var textFrame = myPage.textFrames.add({geometricBounds:myBounds});
    textFrame.transform (CoordinateSpaces.pasteboardCoordinates, AnchorPoint.centerAnchor, app.transformationMatrices.add({counterclockwiseRotationAngle: myTransfrom}));    
    textFrame.texts.item(0).insertionPoints.item(0).contents = myString
    textFrame.texts.item(0).parentStory.appliedFont = app.fonts.item(myFontName);
    textFrame.texts.item(0).parentStory.pointSize = myPointSize;
    textFrame.texts.item(0).justification = myJustification;
    textFrame.textFramePreferences.verticalJustification = VerticalJustification.CENTER_ALIGN;
    textFrame.texts.item(0).fillColor = "Blue";
    layoutCounter = 1;
    if(myFitToContent === true){
        textFrame.fit(FitOptions.frameToContent);
    }
    return textFrame;
}



mP12TF = document.masterSpreads.item(3).textFrames.add();  
mP12TF.geometricBounds = ["10", "15", "30", "165"];    //["上", "左", "下", "右",]
mP12TF.texts.item(0).contents = 類型[type] + " 12 Pages";
mP12TF.texts.item(0).appliedFont = "新細明體";
mP12TF.texts.item(0).pointSize = 20;
mP12TF.texts.item(0).justification = Justification.leftAlign;
mP12TF.texts.item(0).fillColor = "Blue";


mP8TF = document.masterSpreads.item(4).textFrames.add();  
mP8TF.geometricBounds = ["10", "15", "30", "165"];    //["上", "左", "下", "右",]
mP8TF.texts.item(0).contents = 類型[type] + " 8 Pages";
mP8TF.texts.item(0).appliedFont = "新細明體";
mP8TF.texts.item(0).pointSize = 20;
mP8TF.texts.item(0).justification = Justification.leftAlign;
mP8TF.texts.item(0).fillColor = "Blue";


mP4TF = document.masterSpreads.item(5).textFrames.add();  
mP4TF.geometricBounds = ["10", "15", "30", "165"];    //["上", "左", "下", "右",]
mP4TF.texts.item(0).contents = 類型[type] + " 4 Pages";
mP4TF.texts.item(0).appliedFont = "新細明體";
mP4TF.texts.item(0).pointSize = 20;
mP4TF.texts.item(0).justification = Justification.leftAlign;
mP4TF.texts.item(0).fillColor = "Blue";


page1 = endPage1 - startPage1 + 1;
page1Counter = startPage1;

page2 = endPage2 - startPage2 + 1;
page2Counter = startPage2;

page3 = endPage3 - startPage3 + 1;
page3Counter = startPage3;

page4 = endPage4 - startPage4 + 1;
page4Counter = startPage4;

page5 = endPage5 - startPage5 + 1;
page5Counter = startPage5;

page6 = endPage6 - startPage6 + 1;
page6Counter = startPage6;

page7 = endPage7 - startPage7 + 1;
page7Counter = startPage7;

page8 = endPage8 - startPage8 + 1;
page8Counter = startPage8;

page9 = endPage9 - startPage9 + 1;
page9Counter = startPage9;

page10 = endPage10 - startPage10 + 1;
page10Counter = startPage10;

page11 = endPage11 - startPage11 + 1;
page11Counter = startPage11;

page12 = endPage12 - startPage12 + 1;
page12Counter = startPage12;

page13 = endPage13 - startPage13 + 1;
page13Counter = startPage13;

page14 = endPage14 - startPage14 + 1;
page14Counter = startPage14;

page15 = endPage15 - startPage15 + 1;
page15Counter = startPage15;
*/