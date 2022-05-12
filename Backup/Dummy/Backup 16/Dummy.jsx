﻿var type;
var jobName;
var cover;
var imprint;
var startContents, endContents;
var startpage1, startpage2, startpage3, startpage4, startpage5, startpage6, startpage7, startpage8, startpage9, startpage10, startpage11, startpage12, startpage13, startpage14, startpage15;
var endPage1, endPage2, endPage3, endPage4, endPage5, endPage6, endPage7, endPage8, endPage9, endPage10, endPage11, endPage12, endPage13, endPage14, endPage15;
var blank1, blank2, blank3, blank4, blank5, blank6, blank7, blank8, blank9, blank10, blank11, blank12, blank13, blank14, blank15;
var 類型=["大報", "憲報 No. 1", "憲報 No. 2", "憲報 No. 3", "憲報 No. 4", "憲報 No. 5", "憲報 No. 6", "特報", "簽字", "No. 1 + Index A", "No. 2 + Index B", "No. 3 + Index C", "Sup 1 + Index A", "Sup 2 + Index B", "Sup 3 + Index C"];

var pageNumber;
var totalPage;
var pageCounter;
var sectionCounter
var layoutCounter=1;
var pb32=0, pb16=0, pb12=0, pb8=0, pb4=0;
var ss32=0, ss16=0, ss12=0, ss8=0, ss4=0;
var layout32F, layout32L, layout16, layout12, layout8, layout4;

var listPageNumber;
var t=33.07;
var l=16.65;
var b=39.5;
var r=28.5;
var i;
var x;
var section32 = new Array(i);
var section16 = new Array(i);
var section12 = new Array(i);
var section8 = new Array(i);
var section4 = new Array(i);
var column = new Array(i); 

app.documents.item(0).viewPreferences.horizontalMeasurementUnits=MeasurementUnits.millimeters;
app.documents.item(0).viewPreferences.verticalMeasurementUnits=MeasurementUnits.millimeters;
document.marginPreferences.top=0;
document.marginPreferences.left=0;
document.marginPreferences.bottom=0;
document.marginPreferences.right=0;
document.documentPreferences.pageWidth=210;
document.documentPreferences.pageHeight=297;


var tag=["", "A", "B", "C", "D", "E",];
if(app.documents.length === 0){
    var tmp=new Array();
    tmp["width"]=612;
    tmp["height"]=792;
    placementINFO["pgSize"] =tmp;
}   

//  declare dialog data information
makeDialog();
dialog.center(); 
if(dialog.show() === 1){
    jobName=String(dialog.jobName.text);
    cover=Number(dialog.cover.text);
    imprint=Number(dialog.imprint.text);
    startContents=Number(dialog.startContents.text);
    endContents=Number(dialog.endContents.text);
    startPage1=Number(dialog.startPage1.text);
    endPage1=Number(dialog.endPage1.text);
    startPage2=Number(dialog.startPage2.text);
    endPage2=Number(dialog.endPage2.text);
    startPage3=Number(dialog.startPage3.text);
    endPage3=Number(dialog.endPage3.text);
    startPage4=Number(dialog.startPage4.text);
    endPage4=Number(dialog.endPage4.text);
    startPage5=Number(dialog.startPage5.text);
    endPage5=Number(dialog.endPage5.text);
    startPage6=Number(dialog.startPage6.text);
    endPage6=Number(dialog.endPage6.text);
    startPage7=Number(dialog.startPage7.text);
    endPage7=Number(dialog.endPage7.text);
    startPage8=Number(dialog.startPage8.text);
    endPage8=Number(dialog.endPage8.text);
    startPage9=Number(dialog.startPage9.text);
    endPage9=Number(dialog.endPage9.text);
    startPage10=Number(dialog.startPage10.text);
    endPage10=Number(dialog.endPage10.text);
    startPage11=Number(dialog.startPage11.text);
    endPage11=Number(dialog.endPage11.text);
    startPage12=Number(dialog.startPage12.text);
    endPage12=Number(dialog.endPage12.text);
    startPage13=Number(dialog.startPage13.text);
    endPage13=Number(dialog.endPage13.text);
    startPage14=Number(dialog.startPage14.text);
    endPage14=Number(dialog.endPage14.text);
    startPage15=Number(dialog.startPage15.text);
    endPage15=Number(dialog.endPage15.text);
    type=dialog.kindType.selection.index;
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
    dialog=new Window('dialog', "HAHAHA", "x:0, y:0, width:550, height:1000");
    dialog.panel=dialog.add('panel', [10, 10, 300, 990], "");

    dialog.panel.add('statictext', [10,15,50,40], "Type：");
    dialog.kindType=dialog.panel.add('dropdownlist', [100,10,215,35]);
    dialog.kindType.onChange=kindValidator;
    for(var i=0; i<類型.length; i++){
        dialog.kindType.add('item', 類型[i]);
    }
    
    dialog.panel.add('statictext',  [10,45,90,70], "Job Name："); 
    dialog.jobName=dialog.panel.add('edittext', [100,40,215,65], "");
    dialog.jobName.onChange=jobNameValidator;
    
    dialog.panel.add('statictext',  [10, 75, 50, 110], "Cover："); 
    dialog.cover=dialog.panel.add('edittext', [100, 70, 155, 95], "");
    dialog.cover.onChange=coverValidator;
    
    dialog.panel.add('statictext',  [10, 120, 60, 135], "Imprint："); 
    dialog.imprint=dialog.panel.add('edittext', [100, 115, 155, 140], "");
    dialog.imprint.onChange=imprintValidator;

    dialog.panel.add('statictext',  [10, 175, 50, 195], "目錄："); 
    dialog.startContents=dialog.panel.add('edittext', [100, 170, 155, 195], "");
    dialog.startContents.onChange=startContentsValidator;
    dialog.panel.add('statictext',  [160, 175, 165, 200], "-");
    dialog.endContents=dialog.panel.add('edittext', [170, 170, 220, 195], "");
    dialog.cropType=dialog.panel.add('dropdownlist', [225, 170, 265, 195]);
    for(var i=0; i<tag.length; i++){
        dialog.cropType.add('item', tag[i]);
    }

    dialog.panel.add('statictext',  [10, 215, 50, 240], "內文："); 
    dialog.startPage1=dialog.panel.add('edittext', [100, 210, 155, 235], "");
    dialog.startPage1.onChange=startPage1Validator;
    dialog.panel.add('statictext',  [160, 215, 165, 240], "-");
    dialog.endPage1=dialog.panel.add('edittext', [170, 210, 220, 235], ""); 
    dialog.endPage1.onChange=endPage1Validator;      
    dialog.cropType=dialog.panel.add('dropdownlist', [225, 210, 265, 235]);
    for(var i=0; i<tag.length; i++){
        dialog.cropType.add('item', tag[i]);
    }

    dialog.panel.add('statictext',  [10, 255, 50, 290], "內文："); 
    dialog.startPage2=dialog.panel.add('edittext', [100, 250, 155, 275], "");
    dialog.startPage2.onChange=startPage2Validator;
    dialog.panel.add('statictext',  [160, 255, 165, 290], "-");
    dialog.endPage2=dialog.panel.add('edittext', [170, 250, 220, 275], ""); 
    dialog.endPage2.onChange=endPage2Validator;      
    dialog.cropType=dialog.panel.add('dropdownlist', [225, 250, 265, 275]);
    for(var i=0; i<tag.length; i++){
        dialog.cropType.add('item', tag[i]);
    }

    dialog.panel.add('statictext',  [10, 295, 50, 310], "內文："); 
    dialog.startPage3=dialog.panel.add('edittext', [100, 290, 155, 315], "");
    dialog.startPage3.onChange=startPage3Validator;
    dialog.panel.add('statictext',  [160, 295, 165, 310], "-");
    dialog.endPage3=dialog.panel.add('edittext', [170, 290, 220, 315], ""); 
    dialog.endPage3.onChange=endPage3Validator;      
    dialog.cropType=dialog.panel.add('dropdownlist', [225, 290, 265, 315]);
    for(var i=0; i<tag.length; i++){
        dialog.cropType.add('item', tag[i]);
    }

    dialog.panel.add('statictext',  [10, 335, 50, 350], "內文："); 
    dialog.startPage4=dialog.panel.add('edittext', [100, 330, 155, 355], "");
    dialog.startPage4.onChange=startPage4Validator;
    dialog.panel.add('statictext',  [160, 335, 165, 350], "-");
    dialog.endPage4=dialog.panel.add('edittext', [170, 330, 220, 355], ""); 
    dialog.endPage4.onChange=endPage4Validator;      
    dialog.cropType=dialog.panel.add('dropdownlist', [225, 330, 265, 355]);
    for(var i=0; i<tag.length; i++){
        dialog.cropType.add('item', tag[i]);
    }

    dialog.panel.add('statictext',  [10, 375, 50, 390], "內文："); 
    dialog.startPage5=dialog.panel.add('edittext', [100, 370, 155, 395], "");
    dialog.startPage5.onChange=startPage5Validator;
    dialog.panel.add('statictext',  [160, 375, 165, 390], "-");
    dialog.endPage5=dialog.panel.add('edittext', [170, 370, 220, 395], ""); 
    dialog.endPage5.onChange=endPage5Validator;      
    dialog.cropType=dialog.panel.add('dropdownlist', [225, 370, 265, 395]);
    for(var i=0; i<tag.length; i++){
        dialog.cropType.add('item', tag[i]);
    }

    dialog.panel.add('statictext',  [10, 415, 50, 430], "內文："); 
    dialog.startPage6=dialog.panel.add('edittext', [100, 410, 155, 435], "");
    dialog.startPage6.onChange=startPage6Validator;
    dialog.panel.add('statictext',  [160, 415, 165, 430], "-");
    dialog.endPage6=dialog.panel.add('edittext', [170, 410, 220, 435], ""); 
    dialog.endPage6.onChange=endPage6Validator;      
    dialog.cropType=dialog.panel.add('dropdownlist', [225, 410, 265, 435]);
    for(var i=0; i<tag.length; i++){
        dialog.cropType.add('item', tag[i]);
    }

    dialog.panel.add('statictext',  [10, 455, 50, 470], "內文："); 
    dialog.startPage7=dialog.panel.add('edittext', [100, 450, 155, 475], "");
    dialog.startPage7.onChange=startPage7Validator;
    dialog.panel.add('statictext',  [160, 455, 165, 470], "-");
    dialog.endPage7=dialog.panel.add('edittext', [170, 450, 220, 475], ""); 
    dialog.endPage7.onChange=endPage7Validator;      
    dialog.cropType=dialog.panel.add('dropdownlist', [225, 450, 265, 475]);
    for(var i=0; i<tag.length; i++){
        dialog.cropType.add('item', tag[i]);
    }

    dialog.panel.add('statictext',  [10, 495, 50, 510], "內文："); 
    dialog.startPage8=dialog.panel.add('edittext', [100, 490, 155, 515], "");
    dialog.startPage8.onChange=startPage8Validator;
    dialog.panel.add('statictext',  [160, 495, 165, 510], "-");
    dialog.endPage8=dialog.panel.add('edittext', [170, 490, 220, 515], ""); 
    dialog.endPage8.onChange=endPage8Validator;      
    dialog.cropType=dialog.panel.add('dropdownlist', [225, 490, 265, 515]);
    for(var i=0; i<tag.length; i++){
        dialog.cropType.add('item', tag[i]);
    }

    dialog.panel.add('statictext',  [10, 535, 50, 550], "內文："); 
    dialog.startPage9=dialog.panel.add('edittext', [100, 530, 155, 555], "");
    dialog.startPage9.onChange=startPage9Validator;
    dialog.panel.add('statictext',  [160, 535, 165, 550], "-");
    dialog.endPage9=dialog.panel.add('edittext', [170, 530, 220, 555], ""); 
    dialog.endPage9.onChange=endPage9Validator;      
    dialog.cropType=dialog.panel.add('dropdownlist', [225, 530, 265, 555]);
    for(var i=0; i<tag.length; i++){
        dialog.cropType.add('item', tag[i]);
    }

    dialog.panel.add('statictext',  [10, 575, 50, 590], "內文："); 
    dialog.startPage10=dialog.panel.add('edittext', [100, 570, 155, 595], "");
    dialog.startPage10.onChange=startPage10Validator;
    dialog.panel.add('statictext',  [160, 575, 165, 590], "-");
    dialog.endPage10=dialog.panel.add('edittext', [170, 570, 220, 595], ""); 
    dialog.endPage10.onChange=endPage10Validator;      
    dialog.cropType=dialog.panel.add('dropdownlist', [225, 570, 265, 595]);
    for(var i=0; i<tag.length; i++){
        dialog.cropType.add('item', tag[i]);
    }

    dialog.panel.add('statictext',  [10, 615, 50, 630], "內文："); 
    dialog.startPage11=dialog.panel.add('edittext', [100, 610, 155, 635], "");
    dialog.startPage11.onChange=startPage11Validator;
    dialog.panel.add('statictext',  [160, 615, 165, 630], "-");
    dialog.endPage11=dialog.panel.add('edittext', [170, 610, 220, 635], ""); 
    dialog.endPage11.onChange=endPage11Validator;      
    dialog.cropType=dialog.panel.add('dropdownlist', [225, 610, 265, 635]);
    for(var i=0; i<tag.length; i++){
        dialog.cropType.add('item', tag[i]);
    }

    dialog.panel.add('statictext',  [10, 655, 50, 670], "內文："); 
    dialog.startPage12=dialog.panel.add('edittext', [100, 650, 155, 675], "");
    dialog.startPage12.onChange=startPage12Validator;
    dialog.panel.add('statictext',  [160, 655, 165, 670], "-");
    dialog.endPage12=dialog.panel.add('edittext', [170, 650, 220, 675], ""); 
    dialog.endPage12.onChange=endPage12Validator;      
    dialog.cropType=dialog.panel.add('dropdownlist', [225, 650, 265, 675]);
    for(var i=0; i<tag.length; i++){
        dialog.cropType.add('item', tag[i]);
    }

    dialog.panel.add('statictext',  [10, 695, 50, 710], "內文："); 
    dialog.startPage13=dialog.panel.add('edittext', [100, 690, 155, 715], "");
    dialog.startPage13.onChange=startPage13Validator;
    dialog.panel.add('statictext',  [160, 695, 165, 710], "-");
    dialog.endPage13=dialog.panel.add('edittext', [170, 690, 220, 715], ""); 
    dialog.endPage13.onChange=endPage13Validator;      
    dialog.cropType=dialog.panel.add('dropdownlist', [225, 690, 265, 715]);
    for(var i=0; i<tag.length; i++){
        dialog.cropType.add('item', tag[i]);
    }

    dialog.panel.add('statictext',  [10, 735, 50, 750], "內文："); 
    dialog.startPage14=dialog.panel.add('edittext', [100, 730, 155, 755], "");
    dialog.startPage14.onChange=startPage14Validator;
    dialog.panel.add('statictext',  [160, 735, 165, 750], "-");
    dialog.endPage14=dialog.panel.add('edittext', [170, 730, 220, 755], ""); 
    dialog.endPage14.onChange=endPage14Validator;      
    dialog.cropType=dialog.panel.add('dropdownlist', [225, 730, 265, 755]);
    for(var i=0; i<tag.length; i++){
        dialog.cropType.add('item', tag[i]);
    }

    dialog.panel.add('statictext',  [10, 695, 50, 710], "內文："); 
    dialog.startPage15=dialog.panel.add('edittext', [100, 770, 155, 795], "");
    dialog.startPage15.onChange=startPage15Validator;
    dialog.panel.add('statictext',  [160, 695, 165, 710], "-");
    dialog.endPage15=dialog.panel.add('edittext', [170, 770, 220, 795], ""); 
    dialog.endPage15.onChange=endPage15Validator;      
    dialog.cropType=dialog.panel.add('dropdownlist', [225, 770, 265, 795]);
    for(var i=0; i<tag.length; i++){
        dialog.cropType.add('item', tag[i]);
    }



//    dialog.placeOnLayer=dialog.panel.add('checkbox', [10, 600, 220, 635], "\"This\" Page");
    dialog.placeOnLayer=dialog.panel.add('checkbox', [10, 775, 220, 790], "Holiday");
    
    // The buttons
    dialog.OKbut=dialog.add('button', [450, 20, 525, 45], "OK");
    dialog.OKbut.onClick=onOKclicked;
    dialog.CANbut=dialog.add('button', [450, 50, 525, 75], "Cancel");
    dialog.CANbut.onClick=onCANclicked;

    return dialog;
}

//  get total page number
if(endPage15){
    totalPage=endPage15 - startContents + 1;
    pageCounter=endPage15
}else if(endPage14){
    totalPage=endPage14 - startContents + 1;
    pageCounter=endPage14
}else if(endPage13){
    totalPage=endPage13 - startContents + 1;
    pageCounter=endPage13
}else if(endPage12){
    totalPage=endPage12 - startContents + 1;
    pageCounter=endPage12
}else if(endPage11){
    totalPage=endPage11 - startContents + 1;
    pageCounter=endPage11
}else if(endPage10){
    totalPage=endPage10 - startContents + 1;
    pageCounter=endPage10
}else if(endPage9){
    totalPage=endPage9 - startContents + 1;
    pageCounter=endPage9
}else if(endPage8){
    totalPage=endPage8 - startContents + 1;
    pageCounter=endPage8
}else if(endPage7){
    totalPage=endPage7 - startContents + 1;
    pageCounter=endPage7
}else if(endPage6){
    totalPage=endPage6 - startContents + 1;
    pageCounter=endPage6
}else if(endPage5){
    totalPage=endPage5 - startContents + 1;
    pageCounter=endPage5
}else if(endPage4){
    totalPage=endPage4 - startContents + 1;
    pageCounter=endPage4
}else if(endPage3){
    totalPage=endPage3 - startContents + 1;
    pageCounter=endPage3
}else if(endPage2){
    totalPage=endPage2 - startContents + 1;
    pageCounter=endPage2
}else if(endPage1){
    totalPage=endPage1 - startContents + 1;
    pageCounter=endPage1;
}


//  make pageCounter=total page number to count sections
sectionCounter=totalPage;

//  count 32 pages section
for(; sectionCounter >= 32;){
    pb32 ++;
    sectionCounter=sectionCounter - 32;
    section32.push(32);
}
if(sectionCounter === 31){
    pb32 ++;
    sectionCounter=sectionCounter - 32;
    section32.push(32);
}else if(sectionCounter === 30){
    pb32 ++;
    sectionCounter=sectionCounter - 32;
    section32.push(32);
}else if(sectionCounter === 29){
    pb32 ++;
    sectionCounter=sectionCounter - 32;
    section32.push(32);
}

//  count 16 pages section
if(sectionCounter >= 16){
    pb16 ++;
    sectionCounter=sectionCounter - 16;
    section16.push(16);
}else if(sectionCounter === 15){
    pb16 ++;
    sectionCounter=sectionCounter - 16;
    section16.push(16);
}else if(sectionCounter === 14){
    pb16 ++;
    sectionCounter=sectionCounter - 16;
    section16.push(16);
}else if(sectionCounter === 13){
    pb16 ++;
    sectionCounter=sectionCounter - 16;
    section16.push(16);
}

//  count 12 pages section
if(sectionCounter >= 12){
    pb12 ++;
    sectionCounter=sectionCounter - 12;
    section12.push(12);
}else if(sectionCounter === 11){
    pb12 ++;
    sectionCounter=sectionCounter - 12;
    section12.push(12);
}else if(sectionCounter === 10){
    pb12 ++;
    sectionCounter=sectionCounter - 12;
    section12.push(12);
}else if(sectionCounter === 9){
    pb12 ++;
    sectionCounter=sectionCounter - 12;
    section12.push(12);
}

//  count 8 pages section
if(sectionCounter >= 8){
    pb8 ++;
    sectionCounter=sectionCounter - 8;
    section8.push(8);
}else if(sectionCounter === 7){
    pb8 ++;
    sectionCounter=sectionCounter - 8;
    section8.push(8);
}else if(sectionCounter === 6){
    pb8 ++;
    sectionCounter=sectionCounter - 8;
    section8.push(8);
}else if(sectionCounter === 5){
    pb8 ++;
    sectionCounter=sectionCounter - 8;
    section8.push(8);
}

//  count 4 pages section
if(sectionCounter >= 4){
    pb4 ++;
    sectionCounter=sectionCounter - 4;
    section4.push(4);
}else if(sectionCounter === 3){
    pb4 ++;
    sectionCounter=sectionCounter - 4;
    section4.push(4);
}else if(sectionCounter === 2){
    pb4 ++;
    sectionCounter=sectionCounter - 4;
    section4.push(4);
}else if(sectionCounter === 1){
    pb4 ++;
    sectionCounter=sectionCounter - 4;
    section4.push(4);
}


//  Create number page header
document.pages.item(0).appliedMaster=app.documents[0].masterSpreads.item ('G-主版')
makeText(document.pages.item(0), [252.104, 24.7 , 263.736, 29.7], "" + pb32, "Times New Roman", 13, "Red", Justification.centerAlign, VerticalJustification.BOTTOM_ALIGN, 0, false);
makeText(document.pages.item(0), [252.104, 72.2, 263.736, 77.2], "" + pb16, "Times New Roman", 13, "Red",  Justification.centerAlign, VerticalJustification.BOTTOM_ALIGN, 0, false);
makeText(document.pages.item(0), [252.104, 119.7 , 263.736, 124.7], "" + pb12, "Times New Roman", 13, "Red",  Justification.centerAlign, VerticalJustification.BOTTOM_ALIGN, 0, false);
makeText(document.pages.item(0), [252.104, 167.2, 263.736, 172.2], "" + pb8, "Times New Roman", 13, "Red",  Justification.centerAlign, VerticalJustification.BOTTOM_ALIGN, 0, false);
makeText(document.pages.item(0), [263.736, 24.7, 275.411, 29.7], "" + pb4, "Times New Roman", 13, "Red",  Justification.centerAlign, VerticalJustification.BOTTOM_ALIGN, 0, false);
if (sectionCounter > 136){
    makeText(document.pages.item(0), [275.411 , 24.7, 287.086, 29.7], "1", "Times New Roman", 13, "Red",  Justification.centerAlign, VerticalJustification.BOTTOM_ALIGN, 0, false);
}
if (類型[type] === "特報"  &&  pb4>0 && totalPage<5){
    makeText(document.pages.item(0), [263.736, 43.6, 275.411, 200 ], "→　Flat Work \\\\ gaz sm33_4pp pf Logo single", "Times New Roman", 13, "Red",  Justification.leftAlign, VerticalJustification.BOTTOM_ALIGN, 0, false);
}else if (類型[type]=="Sup 1 + Index A" || 類型[type]=="Sup 2 + Index B" || 類型[type]=="Sup 3 + Index C"  &&  pb4>0 && totalPage<5){
    makeText(document.pages.item(0), [263.736, 43.6, 275.411, 200 ], "→　Flat Work \\\\ gaz sm33_4pp pf W/B single", "Times New Roman", 13, "Red",  Justification.leftAlign, VerticalJustification.BOTTOM_ALIGN, 0, false);
}

listPageNumber=startContents;

if (totalPage > 136){
    for (x=0; x<(section32.length-1); x=x+1){
        for (i=0; i<section32[x]; i=i+1){
            column[i] = document.pages.item(0).textFrames.add();
            column[i].geometricBounds = [t, l, b, r];    //["上", "左", "下", "右",]
            column[i].texts.item(0).contents = listPageNumber + "";
            column[i].texts.item(0).pointSize = 12;
            column[i].texts.item(0).justification = Justification.centerAlign;
            column[i].textFramePreferences.verticalJustification = VerticalJustification.CENTER_ALIGN
            listPageNumber=listPageNumber+1;
            t=t+6.437;
            b=b+6.437;
        }
        t=33.07;
        b=39.5;
        l=l+18.328;
        r=r+18.328;
    }
    for (x=0; x<section16.length; x=x+1){
        for (i=0; i<section16[x]; i=i+1){
            column[i] = document.pages.item(0).textFrames.add();
            column[i].geometricBounds = [t, l, b, r];    //["上", "左", "下", "右",]
            column[i].texts.item(0).contents = listPageNumber + "";
            column[i].texts.item(0).pointSize = 12;
            column[i].texts.item(0).justification = Justification.centerAlign;
            column[i].textFramePreferences.verticalJustification = VerticalJustification.CENTER_ALIGN
            listPageNumber=listPageNumber+1;
            t=t+6.437;
            b=b+6.437;
        }
        t=33.07;
        b=39.5;
        l=l+18.328;
        r=r+18.328;
    }
    for (x=0; x<section12.length; x=x+1){
        for (i=0; i<section12[x]; i=i+1){
            column[i] = document.pages.item(0).textFrames.add();
            column[i].geometricBounds = [t, l, b, r];    //["上", "左", "下", "右",]
            column[i].texts.item(0).contents = listPageNumber + "";
            column[i].texts.item(0).pointSize = 12;
            column[i].texts.item(0).justification = Justification.centerAlign;
            column[i].textFramePreferences.verticalJustification = VerticalJustification.CENTER_ALIGN
            listPageNumber=listPageNumber+1;
            t=t+6.437;
            b=b+6.437;
        }
        t=33.07;
        b=39.5;
        l=l+18.328;
        r=r+18.328;
    }
    for (x=0; x<section8.length; x=x+1){
        for (i=0; i<section8[x]; i=i+1){
            column[i] = document.pages.item(0).textFrames.add();
            column[i].geometricBounds = [t, l, b, r];    //["上", "左", "下", "右",]
            column[i].texts.item(0).contents = listPageNumber + "";
            column[i].texts.item(0).pointSize = 12;
            column[i].texts.item(0).justification = Justification.centerAlign;
            column[i].textFramePreferences.verticalJustification = VerticalJustification.CENTER_ALIGN
            listPageNumber=listPageNumber+1;
            t=t+6.437;
            b=b+6.437;
        }
        t=33.07;
        b=39.5;
        l=l+18.328;
        r=r+18.328;
    }
    for (x=0; x<section4.length; x=x+1){
        for (i=0; i<section4[x]; i=i+1){
            column[i] = document.pages.item(0).textFrames.add();
            column[i].geometricBounds = [t, l, b, r];    //["上", "左", "下", "右",]
            column[i].texts.item(0).contents = listPageNumber + "";
            column[i].texts.item(0).pointSize = 12;
            column[i].texts.item(0).justification = Justification.centerAlign;
            column[i].textFramePreferences.verticalJustification = VerticalJustification.CENTER_ALIGN
            listPageNumber=listPageNumber+1;
            t=t+6.437;
            b=b+6.437;
        }
        t=33.07;
        b=39.5;
        l=l+18.328;
        r=r+18.328;
    }
        for (i=0; i<section32[x]; i=i+1){
            column[i] = document.pages.item(0).textFrames.add();
            column[i].geometricBounds = [t, l, b, r];    //["上", "左", "下", "右",]
            column[i].texts.item(0).contents = listPageNumber + "";
            column[i].texts.item(0).pointSize = 12;
            column[i].texts.item(0).justification = Justification.centerAlign;
            column[i].textFramePreferences.verticalJustification = VerticalJustification.CENTER_ALIGN
            listPageNumber=listPageNumber+1;
            t=t+6.437;
            b=b+6.437;
        }
        t=33.07;
        b=39.5;
        l=l+18.328;
        r=r+18.328;
}else{
    
}

alert(totalPage);


//  Pefect Bound
if (totalPage > 136){alert("膠裝 599");
//  make 32 pages number table
    for(; pb32 > 1;){
    pb32 --;
        document.pages.add().appliedMaster=app.documents[0].masterSpreads.item ('B-主版');

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
    for(; pb4 > 0;){
        pb4 --;
        document.pages.add().appliedMaster=app.documents[0].masterSpreads.item ('F-主版');

        makeText(document.pages.item(layoutCounter), [66.024, 20.492, 103.524, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        if(startContents <= pageCounter){
                makeText(document.pages.item(layoutCounter), [66.024, 57.992, 103.524, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [66.024, 57.992, 103.524, 84.508], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                app.findGrepPreferences.findWhat="(?<!\\r)^"+startContents+"$";
                app.changeGrepPreferences.changeTo="Blank"; 
                app.activeDocument.changeGrep();
                app.changeGrepPreferences = NothingEnum.nothing;
                startContents ++;
         }        
        if(startContents <= pageCounter){
                makeText(document.pages.item(layoutCounter), [39.508, 57.992, 77.008, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [39.508, 57.992, 77.008, 84.508], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                app.findGrepPreferences.findWhat="(?<!\\r)^"+startContents+"$";
                app.changeGrepPreferences.changeTo="Blank";
                app.activeDocument.changeGrep();
                app.changeGrepPreferences = NothingEnum.nothing;
                startContents ++;
        }
        if(startContents <= pageCounter){
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                app.findGrepPreferences.findWhat="(?<!\\r)^"+startContents+"$";
                app.changeGrepPreferences.changeTo="Blank";
                app.activeDocument.changeGrep();
                app.changeGrepPreferences = NothingEnum.nothing;
                startContents ++;
        }        
        layoutCounter ++;
    }

    //  make 8 pages number table
    for(; pb8 > 0;){
        pb8 --;
        document.pages.add().appliedMaster=app.documents[0].masterSpreads.item ('E-主版');

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
        if(startContents <= pageCounter){
                makeText(document.pages.item(layoutCounter), [196.96, 57.992, 234.46, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [196.96, 57.992, 234.46, 84.508], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                app.findGrepPreferences.findWhat="(?<!\\r)^"+startContents+"$";
                app.changeGrepPreferences.changeTo="Blank";
                app.activeDocument.changeGrep();
                app.changeGrepPreferences = NothingEnum.nothing;
                startContents ++;
        }        
        if(startContents <= pageCounter){
                makeText(document.pages.item(layoutCounter), [196.96, 20.492, 234.46, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [196.96, 20.492, 234.46, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                app.findGrepPreferences.findWhat="(?<!\\r)^"+startContents+"$";
                app.changeGrepPreferences.changeTo="Blank";
                app.activeDocument.changeGrep();
                app.changeGrepPreferences = NothingEnum.nothing;
                startContents ++;
        }
        if(startContents <= pageCounter){
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                app.findGrepPreferences.findWhat="(?<!\\r)^"+startContents+"$";
                app.changeGrepPreferences.changeTo="Blank";
                app.activeDocument.changeGrep();
                app.changeGrepPreferences = NothingEnum.nothing;
                startContents ++;
        }        
        layoutCounter ++;
    }

    //  make 12 pages number table
    for(; pb12 > 0;){
        pb12 --;
        document.pages.add().appliedMaster=app.documents[0].masterSpreads.item ('D-主版');

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
        if(startContents <= pageCounter){
                makeText(document.pages.item(layoutCounter), [39.508, 95.492, 77.008, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [39.508, 95.492, 77.008, 122.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                app.findGrepPreferences.findWhat="(?<!\\r)^"+startContents+"$";
                app.changeGrepPreferences.changeTo="Blank";
                app.activeDocument.changeGrep();
                app.changeGrepPreferences = NothingEnum.nothing;
                startContents ++;
        }        
        if(startContents <= pageCounter){
                makeText(document.pages.item(layoutCounter), [196.96, 20.492, 234.46, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [196.96, 20.492, 234.46, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                app.findGrepPreferences.findWhat="(?<!\\r)^"+startContents+"$";
                app.changeGrepPreferences.changeTo="Blank";
                app.activeDocument.changeGrep();
                app.changeGrepPreferences = NothingEnum.nothing;
                startContents ++;
        }
        if(startContents <= pageCounter){
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                app.findGrepPreferences.findWhat="(?<!\\r)^"+startContents+"$";
                app.changeGrepPreferences.changeTo="Blank";
                app.activeDocument.changeGrep();
                app.changeGrepPreferences = NothingEnum.nothing;
                startContents ++;
        }        
        layoutCounter ++;
    }

    //  make 16 pages number table
    for(; pb16 > 0;){
        pb16 --;
        document.pages.add().appliedMaster=app.documents[0].masterSpreads.item ('C-主版');

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
        if(startContents <= pageCounter){
                makeText(document.pages.item(layoutCounter), [119.056, 132.992, 156.556, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [119.056, 132.992, 156.556, 159.508], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                app.findGrepPreferences.findWhat="(?<!\\r)^"+startContents+"$";
                app.changeGrepPreferences.changeTo="Blank";
                app.activeDocument.changeGrep();
                app.changeGrepPreferences = NothingEnum.nothing;
                startContents ++;        
        }
        if(startContents <= pageCounter){
                makeText(document.pages.item(layoutCounter), [39.508, 132.992, 77.008, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [39.508, 132.992, 77.008, 159.508], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                app.findGrepPreferences.findWhat="(?<!\\r)^"+startContents+"$";
                app.changeGrepPreferences.changeTo="Blank";
                app.activeDocument.changeGrep();
                app.changeGrepPreferences = NothingEnum.nothing;
                startContents ++;
        }
        if(startContents <= pageCounter){
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                app.findGrepPreferences.findWhat="(?<!\\r)^"+startContents+"$";
                app.changeGrepPreferences.changeTo="Blank";
                app.activeDocument.changeGrep();
                app.changeGrepPreferences = NothingEnum.nothing;
                startContents ++;        
        }
        layoutCounter ++;
    }

    //  make the last 32 pages number table
    for(; pb32 > 0;){
        pb32 --;
        document.pages.add().appliedMaster=app.documents[0].masterSpreads.item ('B-主版');

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
        if(startContents <= pageCounter){
                makeText(document.pages.item(layoutCounter), [249.992, 132.992, 287.492, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [249.992, 132.992, 287.492, 159.508], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                app.findGrepPreferences.findWhat="(?<!\\r)^"+startContents+"$";
                app.changeGrepPreferences.changeTo="Blank";
                app.activeDocument.changeGrep();
                app.changeGrepPreferences = NothingEnum.nothing;
                startContents ++;
        }
        if(startContents <= pageCounter){
                makeText(document.pages.item(layoutCounter), [249.992, 20.492, 287.492, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [249.992, 20.492, 287.492, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                app.findGrepPreferences.findWhat="(?<!\\r)^"+startContents+"$";
                app.changeGrepPreferences.changeTo="Blank";
                app.activeDocument.changeGrep();
                app.changeGrepPreferences = NothingEnum.nothing;
                startContents ++;
        }
        if(startContents <= pageCounter){
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(document.pages.item(layoutCounter), [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                app.findGrepPreferences.findWhat="(?<!\\r)^"+startContents+"$";
                app.changeGrepPreferences.changeTo="Blank";
                app.activeDocument.changeGrep();
                app.changeGrepPreferences = NothingEnum.nothing;
                startContents ++;
        }        
        layoutCounter ++;
    }



//  Saddle Sitiched  //  make 32 pages number table
}else{        alert("騎釘  995");
       if(pb32>1){
        layout32F=document.pages.add();
        makeText(layout32F, [66.024, 20.492, 103.524, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout32F, [223.476, 20.492, 260.976, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout32F, [223.476, 132.992, 260.976, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout32F, [66.024, 132.992, 103.524, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout32F, [92.54, 132.992, 130.04, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout32F, [196.96, 132.992, 234.46, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout32F, [196.96, 20.492, 234.46, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout32F, [92.54, 20.492, 130.04, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout32F, [92.54, 57.992, 130.04, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout32F, [196.96, 57.992, 234.46, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout32F, [196.96, 95.492, 234.46, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout32F, [92.54, 95.492, 130.04, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout32F, [66.024, 95.492, 103.524, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout32F, [223.476, 95.492, 260.976, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout32F, [223.476, 57.992, 260.976, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout32F, [66.024, 57.992, 103.524, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        layout32F.appliedMaster=app.documents[0].masterSpreads.item ('B-主版');
        ss32 ++;
        layoutCounter ++;
    }

    //  Saddle Sitiched  //  make 4 pages number table
    if(ss4 < pb4){
        layout4=document.pages.add();
        makeText(layout4, [66.024, 20.492, 103.524, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        if(startContents <= pageCounter){
                makeText(layout4, [66.024, 57.992, 103.524, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                startContents ++;
        }else{
                makeText(layout4, [66.024, 57.992, 103.524, 84.508], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                app.findGrepPreferences.findWhat="(?<!\\r)^"+startContents+"$";
                app.changeGrepPreferences.changeTo="Blank";
                app.activeDocument.changeGrep();
                app.changeGrepPreferences = NothingEnum.nothing;        
                startContents ++;        
        }
        layout4.appliedMaster=app.documents[0].masterSpreads.item ('F-主版');
        ss4 ++; 
        layoutCounter ++;
    }

    //  Saddle Sitiched  //  make 8 pages number table
    if(ss8 < pb8){
        layout8=document.pages.add();
        makeText(layout8, [66.024, 20.492, 103.524, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout8, [170.444, 20.492, 207.944, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout8, [170.444, 57.992, 207.944, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout8, [66.024, 57.992, 103.524, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        layout8.appliedMaster=app.documents[0].masterSpreads.item ('E-主版');
        ss8 ++;
        layoutCounter ++;
    }

    //  Saddle Sitiched  //  make 12 pages number table
    if(ss12 < pb12){
        layout12=document.pages.add();
        makeText(layout12, [66.024, 20.492, 103.524, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout12, [170.444, 20.492, 207.944, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout12, [66.024, 95.492, 103.524, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout12, [170.44, 95.492, 207.944, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout12, [170.444, 57.992, 207.944, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout12, [66.024, 57.992, 103.524, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        layout12.appliedMaster=app.documents[0].masterSpreads.item ('D-主版');
        ss12 ++;
        layoutCounter ++;
    }

    //  Saddle Sitiched  //  make 16 pages number table
    if(ss16 < pb16){
        layout16=document.pages.add();
        makeText(layout16, [66.024, 20.492, 103.524, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout16, [66.024, 132.992, 103.524, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout16, [92.54, 132.992, 130.04, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout16, [92.54, 20.492, 130.04, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout16, [92.54, 57.992, 130.04, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout16, [92.54, 95.492, 130.04, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout16, [66.024, 95.492, 103.524, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout16, [66.024, 57.992, 103.524, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        layout16.appliedMaster=app.documents[0].masterSpreads.item ('C-主版');
        ss16 ++;
        layoutCounter ++;
    }

    //  Saddle Sitiched  //  make middle 32 pages number table
    for(var i=pb32-2; i>0; i=i-1){
        document.pages.add().appliedMaster=app.documents[0].masterSpreads.item ('B-主版');
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
        ss32 ++;
        layoutCounter ++;
    }

    //  Saddle Sitiched  //  make final 32 pages number table
    if(pb32>0){
        layout32L=document.pages.add();
        makeText(layout32L, [66.024, 20.492, 103.524, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout32L, [223.476, 20.492, 260.976, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout32L, [223.476, 132.992, 260.976, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout32L, [66.024, 132.992, 103.524, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout32L, [92.54, 132.992, 130.04, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout32L, [196.96, 132.992, 234.46, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout32L, [196.96, 20.492, 234.46, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout32L, [92.54, 20.492, 130.04, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout32L, [92.54, 57.992, 130.04, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout32L, [196.96, 57.992, 234.46, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout32L, [196.96, 95.492, 234.46, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout32L, [92.54, 95.492, 130.04, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout32L, [66.024, 95.492, 103.524, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout32L, [223.476, 95.492, 260.976, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout32L, [223.476, 57.992, 260.976, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout32L, [66.024, 57.992, 103.524, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout32L, [39.508, 57.992, 77.008, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout32L, [249.992, 57.992, 287.492, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout32L, [249.992, 95.492, 287.492, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout32L, [39.508, 95.492, 77.008, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout32L, [119.056, 95.492, 156.556, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout32L, [170.444, 95.492, 207.944, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout32L, [170.444, 57.992, 207.944, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout32L, [119.056, 57.992, 156.556, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout32L, [119.056, 20.492, 156.556, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout32L, [170.444, 20.492, 207.944, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout32L, [170.444, 132.992, 207.944, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout32L, [119.056, 132.992, 156.556, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout32L, [39.508, 132.992, 77.008, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout32L, [249.992, 132.992, 287.492, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout32L, [249.992, 20.492, 287.492, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout32L, [39.508, 20.492, 77.008, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        layout32L.appliedMaster=app.documents[0].masterSpreads.item ('B-主版');
        ss32 ++;
        layoutCounter --;
    }

    //  Saddle Sitiched  //  make another half  middle 32 pages number table
    while(ss32 <= pb32 && ss32 > 2 && pb32 > 2 ){
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
        ss32 --;
        layoutCounter --;
    }


    //  Saddle Sitiched  //  make another half 16 pages number table 
    if(ss16 > 0){
        makeText(layout16, [39.508, 57.992, 77.008, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout16, [39.508, 95.492, 77.008, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout16, [119.056, 95.492, 156.556, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout16, [119.056, 57.992, 156.556, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout16, [119.056, 20.492, 156.556, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        if(startContents <= pageCounter){
                makeText(layout16, [119.056, 132.992, 156.556, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                startContents ++;
        }else{
                makeText(layout16, [119.056, 132.992, 156.556, 159.508], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                app.findGrepPreferences.findWhat="(?<!\\r)^"+startContents+"$";
                app.changeGrepPreferences.changeTo="Blank";
                app.activeDocument.changeGrep();
                app.changeGrepPreferences = NothingEnum.nothing;
                startContents ++;
        }
        if(startContents <= pageCounter){
                makeText(layout16, [39.508, 132.992, 77.008, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                startContents ++;
        }else{
                makeText(layout16, [39.508, 132.992, 77.008, 159.508], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                app.findGrepPreferences.findWhat="(?<!\\r)^"+startContents+"$";
                app.changeGrepPreferences.changeTo="Blank";
                app.activeDocument.changeGrep();
                app.changeGrepPreferences = NothingEnum.nothing;        
                startContents ++;        
        }
        if(startContents <= pageCounter){
                makeText(layout16, [39.508, 20.492, 77.008, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(layout16, [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                app.findGrepPreferences.findWhat="(?<!\\r)^"+startContents+"$";
                app.changeGrepPreferences.changeTo="Blank";
                app.activeDocument.changeGrep();
                app.changeGrepPreferences = NothingEnum.nothing;        
                startContents ++;        
        }
        layoutCounter --;
    }

    //  Saddle Sitiched  //  make another half 12 pages number table
    if(ss12 > 0){
        makeText(layout12, [39.508, 57.992, 77.008, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout12, [196.96, 57.992, 234.46, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout12, [196.96, 95.492, 234.46, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        if(startContents <= pageCounter){
                makeText(layout12, [39.508, 95.492, 77.008, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(layout12, [39.508, 95.492, 77.008, 122.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                app.findGrepPreferences.findWhat="(?<!\\r)^"+startContents+"$";
                app.changeGrepPreferences.changeTo="Blank";
                app.activeDocument.changeGrep();
                app.changeGrepPreferences = NothingEnum.nothing;        
                startContents ++;        
        }        
        if(startContents <= pageCounter){
                makeText(layout12, [196.96, 20.492, 234.46, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(layout12, [196.96, 20.492, 234.46, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                app.findGrepPreferences.findWhat="(?<!\\r)^"+startContents+"$";
                app.changeGrepPreferences.changeTo="Blank";
                app.activeDocument.changeGrep();
                app.changeGrepPreferences = NothingEnum.nothing;        
                startContents ++;        
        }
        if(startContents <= pageCounter){
                makeText(layout12, [39.508, 20.492, 77.008, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(layout12, [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                app.findGrepPreferences.findWhat="(?<!\\r)^"+startContents+"$";
                app.changeGrepPreferences.changeTo="Blank";
                app.activeDocument.changeGrep();
                app.changeGrepPreferences = NothingEnum.nothing;        
                startContents ++;        
        }        
        layoutCounter --;
    }

    //  Saddle Sitiched  //  make another half 8 pages number table
    if(ss8 > 0){
        makeText(layout8, [39.508, 57.992, 77.008, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        if(startContents <= pageCounter){
                makeText(layout8, [196.96, 57.992, 234.46, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                startContents ++;
        }else{
                makeText(layout8, [196.96, 57.992, 234.46, 84.508], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                app.findGrepPreferences.findWhat="(?<!\\r)^"+startContents+"$";
                app.changeGrepPreferences.changeTo="Blank";
                app.activeDocument.changeGrep();
                app.changeGrepPreferences = NothingEnum.nothing;        
                startContents ++;        
        }        
        if(startContents <= pageCounter){
                makeText(layout8, [196.96, 20.492, 234.46, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(layout8, [196.96, 20.492, 234.46, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                app.findGrepPreferences.findWhat="(?<!\\r)^"+startContents+"$";
                app.changeGrepPreferences.changeTo="Blank";
                app.activeDocument.changeGrep();
                app.changeGrepPreferences = NothingEnum.nothing;
                startContents ++;        
        }
        if(startContents <= pageCounter){
                makeText(layout8, [39.508, 20.492, 77.008, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(layout8, [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                app.findGrepPreferences.findWhat="(?<!\\r)^"+startContents+"$";
                app.changeGrepPreferences.changeTo="Blank";
                app.activeDocument.changeGrep();
                app.changeGrepPreferences = NothingEnum.nothing;        
                startContents ++;        
        }        
        layoutCounter --;
    }

    //  Saddle Sitiched  //  make another half 4 pages number table
    if(ss4 > 0){
        if(startContents <= pageCounter){
                makeText(layout4, [39.508, 57.992, 77.008, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                startContents ++;
        }else{
                makeText(layout4, [39.508, 57.992, 77.008, 84.508], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                app.findGrepPreferences.findWhat="(?<!\\r)^"+startContents+"$";
                app.changeGrepPreferences.changeTo="Blank";
                app.activeDocument.changeGrep();
                app.changeGrepPreferences = NothingEnum.nothing;        
                startContents ++;        
        }
        if(startContents <= pageCounter){
                makeText(layout4, [39.508, 20.492, 77.008, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(layout4, [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                app.findGrepPreferences.findWhat="(?<!\\r)^"+startContents+"$";
                app.changeGrepPreferences.changeTo="Blank";
                app.activeDocument.changeGrep();
                app.changeGrepPreferences = NothingEnum.nothing;
                startContents ++;        
        }        
        layoutCounter --;
    }

    //  Saddle Sitiched  //  make another half first 32 pages number table
    while(ss32 > 1 && pb32 > 1){
        makeText(layout32F, [39.508, 57.992, 77.008, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout32F, [249.992, 57.992, 287.492, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout32F, [249.992, 95.492, 287.492, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout32F, [39.508, 95.492, 77.008, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout32F, [119.056, 95.492, 156.556, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout32F, [170.444, 95.492, 207.944, 122.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout32F, [170.444, 57.992, 207.944, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout32F, [119.056, 57.992, 156.556, 84.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout32F, [119.056, 20.492, 156.556, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout32F, [170.444, 20.492, 207.944, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
        startContents ++;
        makeText(layout32F, [170.444, 132.992, 207.944, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout32F, [119.056, 132.992, 156.556, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        makeText(layout32F, [39.508, 132.992, 77.008, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
        startContents ++;
        if(startContents <= pageCounter){
                makeText(layout32F, [249.992, 132.992, 287.492, 159.508], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                startContents ++;
        }else{
                makeText(layout32F, [249.992, 132.992, 287.492, 159.508], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, 90, false);
                app.findGrepPreferences.findWhat="(?<!\\r)^"+startContents+"$";
                app.changeGrepPreferences.changeTo="Blank";
                app.activeDocument.changeGrep();
                app.changeGrepPreferences = NothingEnum.nothing;        
                startContents ++;        
        }
        if(startContents <= pageCounter){
                makeText(layout32F, [249.992, 20.492, 287.492, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(layout32F, [249.992, 20.492, 287.492, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                app.findGrepPreferences.findWhat="(?<!\\r)^"+startContents+"$";
                app.changeGrepPreferences.changeTo="Blank";
                app.activeDocument.changeGrep();
                app.changeGrepPreferences = NothingEnum.nothing;        
                startContents ++;        
        }
        if(startContents <= pageCounter){
                makeText(layout32F, [39.508, 20.492, 77.008, 47.008], "↑\r\r" + startContents, "Times New Roman", 25, "Black", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                startContents ++;
        }else{
                makeText(layout32F, [39.508, 20.492, 77.008, 47.008], "Blank", "Times New Roman", 25, "Blue", Justification.centerAlign, VerticalJustification.CENTER_ALIGN, -90, false);
                app.findGrepPreferences.findWhat="(?<!\\r)^"+startContents+"$";
                app.changeGrepPreferences.changeTo="Blank";
                app.activeDocument.changeGrep();
                app.changeGrepPreferences = NothingEnum.nothing;        
                startContents ++;        
        }        
        ss32 --;
        layoutCounter --;
    }
}


if(startPage1-1 !== endContents && startPage1>0){
    Blank1=startPage1-1;
    app.findGrepPreferences.findWhat="\r"+Blank1+"$|(?<!\\r)"+Blank1+"$";
    app.changeGrepPreferences.changeTo="Blank";
    app.activeDocument.changeGrep();
    app.changeGrepPreferences = NothingEnum.nothing;
}
if(startPage2-1 !== endPage1 && startPage2){
    Blank2=startPage2-1;
    app.findGrepPreferences.findWhat="\r"+Blank2+"$|(?<!\\r)"+Blank2+"$";
    app.changeGrepPreferences.changeTo="Blank";
    app.activeDocument.changeGrep();
    app.changeGrepPreferences = NothingEnum.nothing;
}
if(startPage3-1 !== endPage2 && startPage3){
    Blank3=startPage3-1;
    app.findGrepPreferences.findWhat="\r"+Blank3+"$|(?<!\\r)"+Blank3+"$";
    app.changeGrepPreferences.changeTo="Blank";
    app.activeDocument.changeGrep();
    app.changeGrepPreferences = NothingEnum.nothing;
}
if(startPage4-1 !== endPage3 && startPage4){
    Blank4=startPage4-1;
    app.findGrepPreferences.findWhat="\r"+Blank4+"$|(?<!\\r)"+Blank4+"$";
    app.changeGrepPreferences.changeTo="Blank";
    app.activeDocument.changeGrep();
    app.changeGrepPreferences = NothingEnum.nothing;
}
if(startPage5-1 !== endPage4 && startPage5){
    Blank5=startPage5-1;
    app.findGrepPreferences.findWhat="\r"+Blank5+"$|(?<!\\r)"+Blank5+"$";
    app.changeGrepPreferences.changeTo="Blank";
    app.activeDocument.changeGrep();
    app.changeGrepPreferences = NothingEnum.nothing;
}
if(startPage6-1 !== endPage5 && startPage6){
    Blank6=startPage6-1;
    app.findGrepPreferences.findWhat="\r"+Blank6+"$|(?<!\\r)"+Blank6+"$";
    app.changeGrepPreferences.changeTo="Blank";
    app.activeDocument.changeGrep();
    app.changeGrepPreferences = NothingEnum.nothing;
}
if(startPage7-1 !== endPage6 && startPage7){
    Blank7=startPage7-1;
    app.findGrepPreferences.findWhat="\r"+Blank7+"$|(?<!\\r)"+Blank7+"$";
    app.changeGrepPreferences.changeTo="Blank";
    app.activeDocument.changeGrep();
    app.changeGrepPreferences = NothingEnum.nothing;
}
if(startPage8-1 !== endPage7 && startPage8){
    Blank8=startPage8-1;
    app.findGrepPreferences.findWhat="\r"+Blank8+"$|(?<!\\r)"+Blank8+"$";
    app.changeGrepPreferences.changeTo="Blank";
    app.activeDocument.changeGrep();
    app.changeGrepPreferences = NothingEnum.nothing;
}
if(startPage9-1 !== endPage8 && startPage9){
    Blank9=startPage9-1;
    app.findGrepPreferences.findWhat="\r"+Blank9+"$|(?<!\\r)"+Blank9+"$";
    app.changeGrepPreferences.changeTo="Blank";
    app.activeDocument.changeGrep();
    app.changeGrepPreferences = NothingEnum.nothing;
}
if(startPage10-1 !== endPage9 && startPage10){
    Blank10=startPage10-1;
    app.findGrepPreferences.findWhat="\r"+Blank10+"$|(?<!\\r)"+Blank10+"$";
    app.changeGrepPreferences.changeTo="Blank";
    app.activeDocument.changeGrep();
    app.changeGrepPreferences = NothingEnum.nothing;
}
if(startPage11-1 !== endPage10 && startPage11){
    Blank11=startPage11-1;
    app.findGrepPreferences.findWhat="\r"+Blank11+"$|(?<!\\r)"+Blank11+"$";
    app.changeGrepPreferences.changeTo="Blank";
    app.activeDocument.changeGrep();
    app.changeGrepPreferences = NothingEnum.nothing;
}
if(startPage12-1 !== endPage11 && startPage12){
    Blank12=startPage12-1;
    app.findGrepPreferences.findWhat="\r"+Blank12+"$|(?<!\\r)"+Blank12+"$";
    app.changeGrepPreferences.changeTo="Blank";
    app.activeDocument.changeGrep();
    app.changeGrepPreferences = NothingEnum.nothing;
}
if(startPage13-1 !== endPage12 && startPage13){
    Blank13=startPage13-1;
    app.findGrepPreferences.findWhat="\r"+Blank13+"$|(?<!\\r)"+Blank13+"$";
    app.changeGrepPreferences.changeTo="Blank";
    app.activeDocument.changeGrep();
    app.changeGrepPreferences = NothingEnum.nothing;
}
if(startPage14-1 !== endPage13 && startPage14){
    Blank14=startPage14-1;
    app.findGrepPreferences.findWhat="\r"+Blank14+"$|(?<!\\r)"+Blank14+"$";
    app.changeGrepPreferences.changeTo="Blank";
    app.activeDocument.changeGrep();
    app.changeGrepPreferences = NothingEnum.nothing;
}
if(startPage15-1 !== endPage14 && startPage15){
    Blank15=startPage15-1;
    app.findGrepPreferences.findWhat="\r"+Blank15+"$|(?<!\\r)"+Blank15+"$";
    app.changeGrepPreferences.changeTo="Blank";
    app.activeDocument.changeGrep();
    app.changeGrepPreferences = NothingEnum.nothing;
}


//  Change Blank to blue color
app.findGrepPreferences.findWhat="Blank";
app.changeGrepPreferences.fillColor = "Blue"; 
app.activeDocument.changeGrep();



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
    var textFrame=myPage.textFrames.add({geometricBounds:myBounds});
    textFrame.texts.item(0).insertionPoints.item(0).contents=myString
    textFrame.texts.item(0).parentStory.appliedFont=app.fonts.item(myFontName);
    textFrame.texts.item(0).parentStory.pointSize=myPointSize;
    textFrame.texts.item(0).fillColor=myColor;
    textFrame.texts.item(0).justification=myJustification;
    textFrame.textFramePreferences.verticalJustification=myVerticalJustification;
    textFrame.transform (CoordinateSpaces.pasteboardCoordinates, AnchorPoint.centerAnchor, app.transformationMatrices.add({counterclockwiseRotationAngle: myTransfrom}));    
    layoutCounter=1;
    if(myFitToContent === true){
        textFrame.fit(FitOptions.frameToContent);
    }
    return textFrame;
}





//  trash
/*
if(app.documents.length === 0){
    var tmp=new Array();
    tmp["width"]=612;
    tmp["height"]=792;
    placementINFO["pgSize"] =tmp;
}



function makeText(myPage, myBounds, myString, myFontName, myPointSize, myJustification, myFitToContent, myTransfrom, sectionCounter){
    var textFrame=myPage.textFrames.add({geometricBounds:myBounds});
    textFrame.transform (CoordinateSpaces.pasteboardCoordinates, AnchorPoint.centerAnchor, app.transformationMatrices.add({counterclockwiseRotationAngle: myTransfrom}));    
    textFrame.texts.item(0).insertionPoints.item(0).contents=myString
    textFrame.texts.item(0).parentStory.appliedFont=app.fonts.item(myFontName);
    textFrame.texts.item(0).parentStory.pointSize=myPointSize;
    textFrame.texts.item(0).justification=myJustification;
    textFrame.textFramePreferences.verticalJustification=VerticalJustification.CENTER_ALIGN;
    sectionCounter=1;
    if(myFitToContent === true){
        textFrame.fit(FitOptions.frameToContent);
    }
    return textFrame;
}



function makeBlueText(myPage, myBounds, myString, myFontName, myPointSize, myJustification, myFitToContent, myTransfrom, myFontColor, sectionCounter){
    var textFrame=myPage.textFrames.add({geometricBounds:myBounds});
    textFrame.transform (CoordinateSpaces.pasteboardCoordinates, AnchorPoint.centerAnchor, app.transformationMatrices.add({counterclockwiseRotationAngle: myTransfrom}));    
    textFrame.texts.item(0).insertionPoints.item(0).contents=myString
    textFrame.texts.item(0).parentStory.appliedFont=app.fonts.item(myFontName);
    textFrame.texts.item(0).parentStory.pointSize=myPointSize;
    textFrame.texts.item(0).justification=myJustification;
    textFrame.textFramePreferences.verticalJustification=VerticalJustification.CENTER_ALIGN;
    textFrame.texts.item(0).fillColor="Blue";
    sectionCounter=1;
    if(myFitToContent === true){
        textFrame.fit(FitOptions.frameToContent);
    }
    return textFrame;
}



*/