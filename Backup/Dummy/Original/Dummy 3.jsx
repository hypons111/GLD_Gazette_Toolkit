var page = 0;
var startPage;
var endPage;

var document = app.documents.add();
app.documents.item(0).viewPreferences.horizontalMeasurementUnits = MeasurementUnits.millimeters;
app.documents.item(0).viewPreferences.verticalMeasurementUnits = MeasurementUnits.millimeters;
document.documentPreferences.pageWidth = 210;
document.documentPreferences.pageHeight = 297;
var blue = document.colors.add({name:"Blue", space:ColorSpace.CMYK, model:ColorModel.process, colorValue:[100, 0, 0, 0]});
var red = document.colors.add({name:"Red", space:ColorSpace.CMYK, model:ColorModel.process, colorValue:[0, 100, 0, 0]});

var myRangeButtons, myBasedOnButtons;
var myLabelWidth = 100;
var myDialog = app.dialogs.add({name:"Testing"});
with(myDialog){
    with(dialogColumns.add()){
        with(borderPanels.add()){
            with(dialogRows.add()){
                with(dialogColumns.add()){
                    staticTexts.add({staticLabel:"Kind:", minWidth:myLabelWidth});
                }
                with(dialogColumns.add()){
                    var coverPageNumber = measurementEditboxes.add();
                }
            }
            with(dialogRows.add()){
                with(dialogColumns.add()){
                    staticTexts.add({staticLabel:"Issue:", minWidth:myLabelWidth});
                }
                with(dialogColumns.add()){
                    var coverPageNumber = measurementEditboxes.add();
                }
            }
            with(dialogRows.add()){
                with(dialogColumns.add()){
                    staticTexts.add({staticLabel:"Cover:", minWidth:myLabelWidth});
                }
                with(dialogColumns.add()){
                    var coverPageNumber = measurementEditboxes.add();
                }
            }
            with(dialogRows.add()){
                with(dialogColumns.add()){
                    staticTexts.add({staticLabel:"Contents:", minWidth:myLabelWidth});
                }
                with(dialogColumns.add()){
                    var contentsPageNumber = measurementEditboxes.add();
                }
            }
            with(dialogRows.add()){
                with(dialogColumns.add()){
                    staticTexts.add({staticLabel:"Start Page:", minWidth:myLabelWidth});
                }
                with(dialogColumns.add()){
                    var startPage = measurementEditboxes.add();
                }
            }
             with(dialogRows.add()){
                with(dialogColumns.add()){
                    staticTexts.add({staticLabel:"End Page:", minWidth:myLabelWidth});
                }
                with(dialogColumns.add()){
                    var endPage = measurementEditboxes.add();
                }
            }       
        }
    }
}
myReturn = myDialog.show();
if (myReturn == true){
    var cover = coverPageNumber.editValue;
    var contents = contentsPageNumber.editValue;
    var text = startPage.editValue;
    var totalPage = endPage.editValue - startPage.editValue;
    myDialog.destroy();
}


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




document.rectangles.add({geometricBounds:[45, 15, 71.516, 52.5], strokeWeight:1});
document.rectangles.add({geometricBounds:[71.516, 15, 98.032, 52.5], strokeWeight:1});
document.rectangles.add({geometricBounds:[98.032, 15, 124.548, 52.5], strokeWeight:1});
document.rectangles.add({geometricBounds:[124.548, 15, 151.064, 52.5], strokeWeight:1});

document.rectangles.add({geometricBounds:[45, 52.5 , 71.516, 90], strokeWeight:1});
document.rectangles.add({geometricBounds:[71.516, 52.5 , 98.032, 90], strokeWeight:1});
document.rectangles.add({geometricBounds:[98.032, 52.5 , 124.548, 90], strokeWeight:1});
document.rectangles.add({geometricBounds:[124.548, 52.5 , 151.064, 90], strokeWeight:1});

document.rectangles.add({geometricBounds:[45, 90, 71.516, 127.5], strokeWeight:1});
document.rectangles.add({geometricBounds:[71.516, 90, 98.032, 127.5], strokeWeight:1});
document.rectangles.add({geometricBounds:[98.032, 90, 124.548, 127.5], strokeWeight:1});
document.rectangles.add({geometricBounds:[124.548, 90, 151.064, 127.5], strokeWeight:1});

document.rectangles.add({geometricBounds:[45, 127.5, 71.516, 165], strokeWeight:1});
document.rectangles.add({geometricBounds:[71.516, 127.5, 98.032, 165], strokeWeight:1});
document.rectangles.add({geometricBounds:[98.032, 127.5, 124.548, 165], strokeWeight:1});
document.rectangles.add({geometricBounds:[124.548, 127.5, 151.064, 165], strokeWeight:1});



document.rectangles.add({geometricBounds:[175.936, 15, 202.452, 52.5], strokeWeight:1});
document.rectangles.add({geometricBounds:[202.452, 15, 228.968, 52.5], strokeWeight:1});
document.rectangles.add({geometricBounds:[228.968, 15, 255.484, 52.5], strokeWeight:1});
document.rectangles.add({geometricBounds:[255.484, 15, 282, 52.5], strokeWeight:1});

document.rectangles.add({geometricBounds:[175.936, 52.5 , 202.452, 90], strokeWeight:1});
document.rectangles.add({geometricBounds:[202.452, 52.5 , 228.968, 90], strokeWeight:1});
document.rectangles.add({geometricBounds:[228.968, 52.5 , 255.484, 90], strokeWeight:1});
document.rectangles.add({geometricBounds:[255.484, 52.5 , 282, 90], strokeWeight:1});

document.rectangles.add({geometricBounds:[175.936, 90, 202.452, 127.5], strokeWeight:1});
document.rectangles.add({geometricBounds:[202.452, 90, 228.968, 127.5], strokeWeight:1});
document.rectangles.add({geometricBounds:[228.968, 90, 255.484, 127.5], strokeWeight:1});
document.rectangles.add({geometricBounds:[255.484, 90, 282, 127.5], strokeWeight:1});

document.rectangles.add({geometricBounds:[175.936, 127.5, 202.452, 165], strokeWeight:1});
document.rectangles.add({geometricBounds:[202.452, 127.5, 228.968, 165], strokeWeight:1});
document.rectangles.add({geometricBounds:[228.968, 127.5, 255.484, 165], strokeWeight:1});
document.rectangles.add({geometricBounds:[255.484, 127.5, 282, 165], strokeWeight:1});

// 1
makeText(document.pages.item(0), [39.508, 20.492, 77.008, 47.008], "↑\n\n" + (31 + text), "Times New Roman", 25, Justification.centerAlign, false, -90);
//2
makeText(document.pages.item(0), [66.024, 20.492, 103.524, 47.008], "↑\n\n" + (0 + text), "Times New Roman", 25, Justification.centerAlign, false, -90);
//3
makeText(document.pages.item(0), [92.54, 20.492, 130.04, 47.008], "↑\n\n" + (7 + text), "Times New Roman", 25, Justification.centerAlign, false, -90);
//4
makeText(document.pages.item(0), [119.056, 20.492, 156.556, 47.008], "↑\n\n" + (24 + text), "Times New Roman", 25, Justification.centerAlign, false, -90);
//5
makeText(document.pages.item(0), [39.508, 57.992, 77.008, 84.508], "↑\n\n" + (16 + text), "Times New Roman", 25, Justification.centerAlign, false, 90);
//6
makeText(document.pages.item(0), [66.024, 57.992, 103.524, 84.508], "↑\n\n" + (15 + text), "Times New Roman", 25, Justification.centerAlign, false, 90);
//7
makeText(document.pages.item(0), [92.54, 57.992, 130.04, 84.508], "↑\n\n" + (8 +  text), "Times New Roman", 25, Justification.centerAlign, false, 90);
//8
makeText(document.pages.item(0), [119.056, 57.992, 156.556, 84.508], "↑\n\n" + (23 + text), "Times New Roman", 25, Justification.centerAlign, false, 90);
//9
makeText(document.pages.item(0), [39.508, 95.492, 77.008, 122.008], "↑\n\n" + (19 + text), "Times New Roman", 25, Justification.centerAlign, false, -90);
//10
makeText(document.pages.item(0), [66.024, 95.492, 103.524, 122.008], "↑\n\n" + (12 + text), "Times New Roman", 25, Justification.centerAlign, false, -90);
//11
makeText(document.pages.item(0), [92.54, 95.492, 130.04, 122.008], "↑\n\n" + (11 + text), "Times New Roman", 25, Justification.centerAlign, false, -90);
//12
makeText(document.pages.item(0), [119.056, 95.492, 156.556, 122.008], "↑\n\n" + (20 + text), "Times New Roman", 25, Justification.centerAlign, false, -90);
//13
makeText(document.pages.item(0), [39.508, 132.992, 77.008, 159.508], "↑\n\n" + (28 + text), "Times New Roman", 25, Justification.centerAlign, false, 90);
//14
makeText(document.pages.item(0), [66.024, 132.992, 103.524, 159.508], "↑\n\n" + (3 + text), "Times New Roman", 25, Justification.centerAlign, false, 90);
//15
makeText(document.pages.item(0), [92.54, 132.992, 130.04, 159.508], "↑\n\n" + (4 + text), "Times New Roman", 25, Justification.centerAlign, false, 90);
//16
makeText(document.pages.item(0), [119.056, 132.992, 156.556, 159.508], "↑\n\n" + (27 + text), "Times New Roman", 25, Justification.centerAlign, false, 90);
//17
makeText(document.pages.item(0), [170.444, 20.492, 207.944, 47.008], "↑\n\n" + (25 + text), "Times New Roman", 25, Justification.centerAlign, false, -90);
//18
makeText(document.pages.item(0), [196.96, 20.492, 234.46, 47.008], "↑\n\n" + (6 + text), "Times New Roman", 25, Justification.centerAlign, false, -90);
//19
makeText(document.pages.item(0), [223.476, 20.492, 260.976, 47.008], "↑\n\n" + (1 + text), "Times New Roman", 25, Justification.centerAlign, false, -90);
//20
makeText(document.pages.item(0), [249.992, 20.492, 287.492, 47.008], "↑\n\n" + (30 + text), "Times New Roman", 25, Justification.centerAlign, false, -90);
//21
makeText(document.pages.item(0), [170.444, 57.992, 207.944, 84.508], "↑\n\n" + (22 + text), "Times New Roman", 25, Justification.centerAlign, false, 90);
//22
makeText(document.pages.item(0), [196.96, 57.992, 234.46, 84.508], "↑\n\n" + (9 + text), "Times New Roman", 25, Justification.centerAlign, false, 90);
//23
makeText(document.pages.item(0), [223.476, 57.992, 260.976, 84.508], "↑\n\n" + (14 + text), "Times New Roman", 25, Justification.centerAlign, false, 90);
//24
makeText(document.pages.item(0), [249.992, 57.992, 287.492, 84.508], "↑\n\n" + (17 + text), "Times New Roman", 25, Justification.centerAlign, false, 90);
//25
makeText(document.pages.item(0), [170.444, 95.492, 207.944, 122.008], "↑\n\n" + (21 + text), "Times New Roman", 25, Justification.centerAlign, false, -90);
//26
makeText(document.pages.item(0), [196.96, 95.492, 234.46, 122.008], "↑\n\n" + (10 + text), "Times New Roman", 25, Justification.centerAlign, false, -90);
//27
makeText(document.pages.item(0), [223.476, 95.492, 260.976, 122.008], "↑\n\n" + (13 + text), "Times New Roman", 25, Justification.centerAlign, false, -90);
//28
makeText(document.pages.item(0), [249.992, 95.492, 287.492, 122.008], "↑\n\n" + (18 + text), "Times New Roman", 25, Justification.centerAlign, false, -90);
//29
makeText(document.pages.item(0), [170.444, 132.992, 207.944, 159.508], "↑\n\n" + (26 + text), "Times New Roman", 25, Justification.centerAlign, false, 90);
//30
makeText(document.pages.item(0), [196.96, 132.992, 234.46, 159.508], "↑\n\n" + (5 + text), "Times New Roman", 25, Justification.centerAlign, false, 90);
//31
makeText(document.pages.item(0), [223.476, 132.992, 260.976, 159.508], "↑\n\n" + (2 + text), "Times New Roman", 25, Justification.centerAlign, false, 90);
//32
makeText(document.pages.item(0), [249.992, 132.992, 287.492, 159.508], "↑\n\n" + (29 + text), "Times New Roman", 25, Justification.centerAlign, false, 90);







