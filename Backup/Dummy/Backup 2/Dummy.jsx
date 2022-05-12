//sample    rectangle = document.rectangles.add({geometricBounds:[39.508, 20.492, 77.008, 47.008], strokeWeight:1, strokeColor:black});  //sample
//sample    dialog.pan4 = dialog.add('panel', [210,200,438,350], "類型：");
//sample    mapPages = Number(dialog.mapPages.value);
//sample    reverseOrder = Number(dialog.reverseOrder.value);
//sample    docStartPG = Number(dialog.docStartPG.text);

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


var master = document.masterSpreads.item(0);

document.masterSpreads.add();
var mP32 = document.masterSpreads.item(1);
mP32TF = document.masterSpreads.item(1).textFrames.add();  
mP32TF.geometricBounds = ["20", "20", "277", "190",];    //["上", "左", "下", "右",]
mP32TF.texts.item(0).contents = "32 Pages";
mP32TF.texts.item(0).pointSize = 50;
mP32TF.texts.item(0).justification = Justification.centerAlign;
mP32TF.textFramePreferences.verticalJustification = VerticalJustification.CENTER_ALIGN;


document.masterSpreads.add();
var mP16 = document.masterSpreads.item(2);
mP16TF = document.masterSpreads.item(2).textFrames.add();  
mP16TF.geometricBounds = ["20", "20", "277", "190",];    //["上", "左", "下", "右",]
mP16TF.texts.item(0).contents = "16 Pages";
mP16TF.texts.item(0).pointSize = 50;
mP16TF.texts.item(0).justification = Justification.centerAlign;
mP16TF.textFramePreferences.verticalJustification = VerticalJustification.CENTER_ALIGN;

document.masterSpreads.add();
var mP12 = document.masterSpreads.item(3);
mP12TF = document.masterSpreads.item(3).textFrames.add();  
mP12TF.geometricBounds = ["20", "20", "277", "190",];    //["上", "左", "下", "右",]
mP12TF.texts.item(0).contents = "12 Pages";
mP12TF.texts.item(0).pointSize = 50;
mP12TF.texts.item(0).justification = Justification.centerAlign;
mP12TF.textFramePreferences.verticalJustification = VerticalJustification.CENTER_ALIGN;

document.masterSpreads.add();
var mP8 = document.masterSpreads.item(4);
mP8TF = document.masterSpreads.item(4).textFrames.add();  
mP8TF.geometricBounds = ["20", "20", "277", "190",];    //["上", "左", "下", "右",]
mP8TF.texts.item(0).contents = "8 Pages";
mP8TF.texts.item(0).pointSize = 50;
mP8TF.texts.item(0).justification = Justification.centerAlign;
mP8TF.textFramePreferences.verticalJustification = VerticalJustification.CENTER_ALIGN;

document.masterSpreads.add();
var mP4 = document.masterSpreads.item(5);
mP4TF = document.masterSpreads.item(5).textFrames.add();  
mP4TF.geometricBounds = ["20", "20", "277", "190",];    //["上", "左", "下", "右",]
mP4TF.texts.item(0).contents = "4 Pages";
mP4TF.texts.item(0).pointSize = 50;
mP4TF.texts.item(0).justification = Justification.centerAlign;
mP4TF.textFramePreferences.verticalJustification = VerticalJustification.CENTER_ALIGN;


/*
master32.rectangles.add({geometricBounds:[45, 15, 71.516, 52.5], strokeWeight:1});
master32.rectangles.add({geometricBounds:[45, 15, 71.516, 52.5], strokeWeight:1});
master32.rectangles.add({geometricBounds:[71.516, 15, 98.032, 52.5], strokeWeight:1});
master32.rectangles.add({geometricBounds:[98.032, 15, 124.548, 52.5], strokeWeight:1});
master32.rectangles.add({geometricBounds:[124.548, 15, 151.064, 52.5], strokeWeight:1});

master32.rectangles.add({geometricBounds:[45, 52.5 , 71.516, 90], strokeWeight:1});
master32.rectangles.add({geometricBounds:[71.516, 52.5 , 98.032, 90], strokeWeight:1});
master32.rectangles.add({geometricBounds:[98.032, 52.5 , 124.548, 90], strokeWeight:1});
master32.rectangles.add({geometricBounds:[124.548, 52.5 , 151.064, 90], strokeWeight:1});

master32.rectangles.add({geometricBounds:[45, 90, 71.516, 127.5], strokeWeight:1});
master32.rectangles.add({geometricBounds:[71.516, 90, 98.032, 127.5], strokeWeight:1});
master32.rectangles.add({geometricBounds:[98.032, 90, 124.548, 127.5], strokeWeight:1});
master32.rectangles.add({geometricBounds:[124.548, 90, 151.064, 127.5], strokeWeight:1});

master32.rectangles.add({geometricBounds:[45, 127.5, 71.516, 165], strokeWeight:1});
master32.rectangles.add({geometricBounds:[71.516, 127.5, 98.032, 165], strokeWeight:1});
master32.rectangles.add({geometricBounds:[98.032, 127.5, 124.548, 165], strokeWeight:1});
master32.rectangles.add({geometricBounds:[124.548, 127.5, 151.064, 165], strokeWeight:1});

master32.rectangles.add({geometricBounds:[175.936, 15, 202.452, 52.5], strokeWeight:1});
master32.rectangles.add({geometricBounds:[202.452, 15, 228.968, 52.5], strokeWeight:1});
master32.rectangles.add({geometricBounds:[228.968, 15, 255.484, 52.5], strokeWeight:1});
master32.rectangles.add({geometricBounds:[255.484, 15, 282, 52.5], strokeWeight:1});

master32.rectangles.add({geometricBounds:[175.936, 52.5 , 202.452, 90], strokeWeight:1});
master32.rectangles.add({geometricBounds:[202.452, 52.5 , 228.968, 90], strokeWeight:1});
master32.rectangles.add({geometricBounds:[228.968, 52.5 , 255.484, 90], strokeWeight:1});
master32.rectangles.add({geometricBounds:[255.484, 52.5 , 282, 90], strokeWeight:1});

master32.rectangles.add({geometricBounds:[175.936, 90, 202.452, 127.5], strokeWeight:1});
master32.rectangles.add({geometricBounds:[202.452, 90, 228.968, 127.5], strokeWeight:1});
master32.rectangles.add({geometricBounds:[228.968, 90, 255.484, 127.5], strokeWeight:1});
master32.rectangles.add({geometricBounds:[255.484, 90, 282, 127.5], strokeWeight:1});

master32.rectangles.add({geometricBounds:[175.936, 127.5, 202.452, 165], strokeWeight:1});
master32.rectangles.add({geometricBounds:[202.452, 127.5, 228.968, 165], strokeWeight:1});
master32.rectangles.add({geometricBounds:[228.968, 127.5, 255.484, 165], strokeWeight:1});
master32.rectangles.add({geometricBounds:[255.484, 127.5, 282, 165], strokeWeight:1});
*/




var nop32 = 0, nop16 = 0, nop12 = 0, nop8 = 0, nop4 = 0;
var blankPage = 0;

var 類型 = ["大報", "憲報 No. 1", "憲報 No. 2", "憲報 No. 3", "憲報 No. 4", "憲報 No. 5", "憲報 No. 6", "特報", "簽字", "No. 1 + Index A", "No. 2 + Index B", "No. 3 + Index C", "Sup 1 + Index A", "Sup 2 + Index B", "Sup 3 + Index C", ];
    if(app.documents.length == 0){
        var tmp = new Array();
        tmp["width"] = 612;
        tmp["height"] = 792;
        placementINFO["pgSize"]  = tmp;
}

var tag = ["", "A", "B", "C", "D", "E",];
    if(app.documents.length == 0){
        var tmp = new Array();
        tmp["width"] = 612;
        tmp["height"] = 792;
        placementINFO["pgSize"]  = tmp;
}

var 
// Get and display the dialog
dialog = makeDialog();
dialog.center(); // Center dialog in screen
if(dialog.show() == 1){
    cover = Number(dialog.cover.text);
    startContents = Number(dialog.startContents.text);
    endContents = Number(dialog.endContents.text);
    startPage = Number(dialog.startPage.text);
    endPage = Number(dialog.endPage.text);
}else{
    exit();
}

// Create dialog box
function makeDialog()
{
	dialog = new Window('dialog', "HAHAHA", "x:0, y:0, width:500, height:500");
    dialog.panel = dialog.add('panel', [10, 10, 300, 390], "");

	dialog.panel.add('statictext', [10,15,50,40], "類型：");
	dialog.cropType = dialog.panel.add('dropdownlist', [50,10,165,35]);
	for(i=0;i<類型.length;i++){
		dialog.cropType.add('item', 類型[i]);
	}
    
    dialog.panel.add('statictext',  [10, 55, 50, 80], "封面："); 
    dialog.cover = dialog.panel.add('edittext', [50, 50, 110, 75], "");
    dialog.cover.onChange = coverValidator;

    dialog.panel.add('statictext',  [10, 95, 50, 120], "目錄："); 
    dialog.startContents = dialog.panel.add('edittext', [50, 90, 110, 115], "");
    dialog.startContents.onChange = startContentsValidator;
    dialog.panel.add('statictext',  [115, 95, 125, 120], "-");
    dialog.endContents = dialog.panel.add('edittext', [125, 90, 185, 115], "");
    dialog.endContents.onChange = endContentsValidator;
    dialog.cropType = dialog.panel.add('dropdownlist', [200, 90, 235, 115]);
    for(i=0;i<tag.length;i++){
        dialog.cropType.add('item', tag[i]);
    }
    
    dialog.panel.add('statictext',  [10, 135, 50, 160], "內文："); 
    dialog.startPage = dialog.panel.add('edittext', [50, 130, 110, 155], "");
    dialog.startPage.onChange = startPageValidator;   
    dialog.panel.add('statictext',  [115, 135, 125, 160], "-");
    dialog.endPage = dialog.panel.add('edittext', [125, 130, 185, 155], ""); 
    dialog.endPage.onChange = endPageValidator;      
    dialog.cropType = dialog.panel.add('dropdownlist', [200, 130, 235, 155]);
    for(i=0;i<tag.length;i++){
        dialog.cropType.add('item', tag[i]);
    }

    dialog.placeOnLayer = dialog.panel.add('checkbox', [10, 170, 220, 195], " \"This\" Page");


	// The buttons
	dialog.OKbut = dialog.add('button', [400, 20, 480, 45], "OK");
	dialog.OKbut.onClick = onOKclicked;
	dialog.CANbut = dialog.add('button', [400, 50, 480, 75], "Cancel");
	dialog.CANbut.onClick = onCANclicked;
	return dialog;
}

var totalPage = endPage - startContents + 1;
var blankPage = 0;

makeText(document.pages.item(0), [10, 70, 40, 130], "totalPage\n" + totalPage, "Times New Roman", 25, Justification.centerAlign, false, 0);

for(; totalPage >= 32;){
    nop32 ++;
    totalPage = totalPage - 32;
}
if(totalPage == 31){
    nop32 ++;
    totalPage = totalPage - 32;
    blankPage = blankPage + 1;
}else if(totalPage == 30){
    nop32 ++;
    totalPage = totalPage - 32;
    blankPage = blankPage + 2;
}else if(totalPage == 29){
    nop32 ++;
    totalPage = totalPage - 32;
    blankPage = blankPage + 3;
}

if(totalPage >= 16){
    nop16 ++;
    totalPage = totalPage - 16;
}else if(totalPage == 15){
    nop16 ++;
    totalPage = totalPage - 16;
    blankPage = blankPage + 1;
}else if(totalPage == 14){
    nop16 ++;
    totalPage = totalPage - 16;
    blankPage = blankPage + 2;
}else if(totalPage == 13){
    nop16 ++;
    totalPage = totalPage - 16;
    blankPage = blankPage + 3;
}

if(totalPage >= 12){
    nop12 ++;
    totalPage = totalPage - 12;
}else if(totalPage == 11){
    nop12 ++;
    totalPage = totalPage - 12;
    blankPage = blankPage + 1;
}else if(totalPage == 10){
    nop12 ++;
    totalPage = totalPage - 12;
    blankPage = blankPage + 2;
}else if(totalPage == 9){
    nop12 ++;
    totalPage = totalPage - 12;
    blankPage = blankPage + 3;
}

if(totalPage >= 8){
    nop8 ++;
    totalPage = totalPage - 8;
}else if(totalPage == 7){
    nop8 ++;
    totalPage = totalPage - 8;
    blankPage = blankPage + 1;
}else if(totalPage == 6){
    nop8 ++;
    totalPage = totalPage - 8;
    blankPage = blankPage + 2;
}else if(totalPage == 5){
    nop8 ++;
    totalPage = totalPage - 8;
    blankPage = blankPage + 3;
}

if(totalPage >= 4){
    nop4 ++;
    totalPage = totalPage - 4;
}else if(totalPage == 3){
    nop4 ++;
    totalPage = totalPage - 4;
    blankPage = blankPage + 1;
}else if(totalPage == 2){
    nop4 ++;
    totalPage = totalPage - 4;
    blankPage = blankPage + 2;
}else if(totalPage == 1){
    nop4 ++;
    totalPage = totalPage - 4;
    blankPage = blankPage + 3;
}

makeText(document.pages.item(0), [10, 20, 40, 70], "Blank Page\n" + blankPage, "Times New Roman", 25, Justification.centerAlign, false, 0);
makeText(document.pages.item(0), [40, 20, 70, 70], "32 Pages\n" + nop32, "Times New Roman", 25, Justification.centerAlign, false, 0);
makeText(document.pages.item(0), [70, 20, 100, 70], "16  Pages\n" + nop16, "Times New Roman", 25, Justification.centerAlign, false, 0);
makeText(document.pages.item(0), [100, 20, 130, 70], "12  Pages\n" + nop12, "Times New Roman", 25, Justification.centerAlign, false, 0);
makeText(document.pages.item(0), [130, 20, 160, 70], "8 Pages\n" + nop8, "Times New Roman", 25, Justification.centerAlign, false, 0);
makeText(document.pages.item(0), [160, 20, 190, 70], "4 Pages\n" + nop4, "Times New Roman", 25, Justification.centerAlign, false, 0);

for(; nop32 > 1;){
    nop32 --;
    page = document.pages.add();
    page.appliedMaster = app.documents[0].masterSpreads.item ('B-主版');
}
for(; nop4 > 0;){
    nop4 --;
    page = document.pages.add();
    page.appliedMaster = app.documents[0].masterSpreads.item ('F-主版');
}
for(; nop8 > 0;){
    nop8 --;
    page = document.pages.add();
    page.appliedMaster = app.documents[0].masterSpreads.item ('E-主版');
}
for(; nop12 > 0;){
    nop12 --;
    page = document.pages.add();
    page.appliedMaster = app.documents[0].masterSpreads.item ('D-主版');
}
for(; nop16 > 0;){
    nop16 --;
    page = document.pages.add();
    page.appliedMaster = app.documents[0].masterSpreads.item ('C-主版');
}
for(; nop32 > 0;){
    nop32 --;
    page = document.pages.add();
    page.appliedMaster = app.documents[0].masterSpreads.item ('B-主版');
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

// Take care of OK being clicked
function onOKclicked(){
	dialog.close(1);
}

// Take care of Cancel being clicked
function onCANclicked(){
	dialog.close(0);
}

// Make text frame
function makeText(myPage, myBounds, myString, myFontName, myPointSize, myJustification, myFitToContent, myTransfrom){
    var textFrame = myPage.textFrames.add({geometricBounds:myBounds});
    textFrame.transform (CoordinateSpaces.pasteboardCoordinates, AnchorPoint.centerAnchor, app.transformationMatrices.add({counterclockwiseRotationAngle: myTransfrom}));	
    textFrame.texts.item(0).insertionPoints.item(0).contents = myString
    textFrame.texts.item(0).parentStory.appliedFont = app.fonts.item(myFontName);
    textFrame.texts.item(0).parentStory.pointSize = myPointSize;
    textFrame.texts.item(0).justification = myJustification;
    textFrame.textFramePreferences.verticalJustification = VerticalJustification.CENTER_ALIGN;
    if(myFitToContent == true){
        textFrame.fit(FitOptions.frameToContent);
    }
	return textFrame;
}





/*dont delete
makeText(document.pages.item(0), [39.508, 20.492, 77.008, 47.008], "↑\n\n" + blankPage, "Times New Roman", 25, Justification.centerAlign, false, -90);
makeText(document.pages.item(0), [66.024, 20.492, 103.524, 47.008], "↑\n\n" + nop32, "Times New Roman", 25, Justification.centerAlign, false, -90);
makeText(document.pages.item(0), [92.54, 20.492, 130.04, 47.008], "↑\n\n" + nop16, "Times New Roman", 25, Justification.centerAlign, false, -90);
makeText(document.pages.item(0), [119.056, 20.492, 156.556, 47.008], "↑\n\n" + nop12, "Times New Roman", 25, Justification.centerAlign, false, -90);
makeText(document.pages.item(0), [172.008, 20.492, 156.556, 47.008], "↑\n\n" + nop8, "Times New Roman", 25, Justification.centerAlign, false, -90);
makeText(document.pages.item(0), [199.56, 20.492, 242.564, 47.008], "↑\n\n" + nop4, "Times New Roman", 25, Justification.centerAlign, false, -90);
    */
    
 /*dont delete   
for(totalPage; totalPage >= 32;){
    
    if(totalPage >= 63 && totalPage <= 33){
        totalPage = totalPage - 32;
        if(totalPage >= 16){
            totalPage = totalPage - 16;
            page = document.pages.add();
            page.appliedMaster = app.documents[0].masterSpreads.item ('C-主版');
        if(totalPage < 4){
            page = document.pages.add();
            page.appliedMaster = app.documents[0].masterSpreads.item ('A-主版');
            }
        }
        if(totalPage >= 12){
            totalPage = totalPage - 12;
            page = document.pages.add();
            page.appliedMaster = app.documents[0].masterSpreads.item ('D-主版');
            if(totalPage < 4){
                page = document.pages.add();
                page.appliedMaster = app.documents[0].masterSpreads.item ('A-主版');
            }
        }
        if(totalPage >= 8){
            totalPage = totalPage - 8;
            page = document.pages.add();
            page.appliedMaster = app.documents[0].masterSpreads.item ('E-主版');
            if(totalPage < 4){
                page = document.pages.add();
                page.appliedMaster = app.documents[0].masterSpreads.item ('A-主版');
            }
        }
        if(totalPage >= 4){
            totalPage = totalPage - 4;
            page = document.pages.add();
            page.appliedMaster = app.documents[0].masterSpreads.item ('F-主版');
            if(totalPage < 4){
                page = document.pages.add();
                page.appliedMaster = app.documents[0].masterSpreads.item ('A-主版');
            }
        }
        totalPage = totalPage + 32;
    }else{
        totalPage = totalPage - 32;
        page = document.pages.add();
        page.appliedMaster = app.documents[0].masterSpreads.item ('B-主版');
    }
}
for(totalPage; totalPage >= 16;){
//   nop16 ++;
    totalPage = totalPage - 16;
    page = document.pages.add();
    page.appliedMaster = app.documents[0].masterSpreads.item ('C-主版');
}
for(totalPage; totalPage >= 12;){
//    nop12 ++;
    totalPage = totalPage - 12;
    page = document.pages.add();
    page.appliedMaster = app.documents[0].masterSpreads.item ('D-主版');
}
for(totalPage; totalPage >= 8;){
//    nop8 ++;
    totalPage = totalPage - 8;
    page = document.pages.add();
    page.appliedMaster = app.documents[0].masterSpreads.item ('E-主版');
}
for(totalPage; totalPage >= 4;){
//    nop4 ++;
    totalPage = totalPage - 4;
    page = document.pages.add();
    page.appliedMaster = app.documents[0].masterSpreads.item ('F-主版');
}
*/

