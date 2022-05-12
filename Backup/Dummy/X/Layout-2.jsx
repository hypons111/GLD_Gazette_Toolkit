//AnimationEncyclopedia.jsx
//An InDesign CS6 JavaScript
//
//Author:  Brenda Burden
//Creates a 6-page sample doc demonstrating the new InDesign CS6 Animation feature

// Search Reference
// PAGE ONE - Sample of Animation Properties 
// PAGE TWO - Sample of all Animation Events
// PAGE THREE - Sample of Additional Animation Properties and Settings, including Duration, Play count, Loop, Speed and Origin
// PAGE FOUR - Sample of Animate To, Animate From, To Current location
// PAGE FIVE - Sample of creating Timing Groups, setting Delays, and looping a Timing Group
// PAGE SIX - Complex A-B-C-D animation that can be created only through scripting (cannot be created in the UI)




main()
function main(){
	mySetSWFExportPreferences();
	var myDocument = app.documents.add();
	mySetUnits(MeasurementUnits.millimeters);
	myDocument.transparencyPreferences.blendingSpace = BlendingSpace.CMYK;
	myDocument.documentPreferences.facingPages = false;
	myDocument.documentPreferences.pageWidth = 210;
	myDocument.documentPreferences.pageHeight = 297;
	var myWhite = myMakeColor("White", ColorSpace.CMYK, ColorModel.process, [0, 0, 0, 0]);
    var myGray = myMakeColor("Gray", ColorSpace.CMYK, ColorModel.process, [0, 0, 0, 70]);	
    var myBlack = myMakeColor("Black", ColorSpace.CMYK, ColorModel.process, [0, 0, 0, 100]);
    

	mySetUnits(MeasurementUnits.millimeters);
//    myMakeRectangle(myDocument.pages.item(0), [20, 20, 138.5, 190], "Background Rectangle", myWhite, myBlack, 2); 
//    myMakeRectangle(myDocument.pages.item(0), [277, 20, 157, 190], "Background Rectangle", myWhite, myBlack, 2); 
//    myMakeTextFrame(myDocument.pages.item(0), [0, 0, 10, 20], "Testing", "Myriad Pro", 6, Justification.centerAlign, true);
	
    var pagenumber = 0;
    var myTextFrame = myDocument.pages.item(0).textFrames.add();
//    myTextFrame.geometricBounds = [49.36, 20, 19.647, 62.5];
//    myTextFrame.contents = "1";
//    myTextFrame.geometricBounds = [79.073 , 20, 49.36, 62.5];
//    myTextFrame.contents = "2";
    
    myMakeTextFrame(myDocument.pages.item(0), [49.36, 20, 19.647, 62.5], "1", "Arial", 18, Justification.centerAlign, false);
    myMakeTextFrame(myDocument.pages.item(0), [79.073 , 20, 49.36 , 62.5], "2", "Arial", 18, Justification.centerAlign, false);
}


//Utility functions.
function mySetUnits(myUnits){
	app.documents.item(0).viewPreferences.horizontalMeasurementUnits = myUnits;
	app.documents.item(0).viewPreferences.verticalMeasurementUnits = myUnits;
}
function myMakeTextFrame(myPage, myBounds, myString, myFontName, myPointSize, myJustification, myFitToContent){
	var myTextFrame = myPage.textFrames.add({geometricBounds:myBounds});
	myTextFrame.texts.item(0).insertionPoints.item(0).contents = myString
	myTextFrame.texts.item(0).parentStory.appliedFont = app.fonts.item(myFontName);
	myTextFrame.texts.item(0).parentStory.pointSize = myPointSize;
	myTextFrame.texts.item(0).parentStory.fillColor = app.documents.item(0).swatches.item("Black");
    myTextFrame.texts.item(0).justification = myJustification;
    if(myFitToContent == true){
		myTextFrame.fit(FitOptions.frameToContent);
	}
    return myTextFrame;	
}

function myMakeRectangle(myPage, myBounds, myString, myFillColor,  myStrokeColor, myStrokeWeight){
	var myRectangle = myPage.rectangles.add({geometricBounds:myBounds, fillColor:myFillColor, strokeWeight:myStrokeWeight, strokeColor:myStrokeColor, name:myString});
	return myRectangle;
}

function myMakeColor(myColorName, myColorSpace, myColorModel, myColorValue){
	var myColor;
	var myDocument = app.documents.item(0);
	myColor = myDocument.colors.item(myColorName);
	if(myColor.isValid == false){
		myColor = myDocument.colors.add({name:myColorName});
	}
	myColor.properties = {space:myColorSpace, model:myColorModel, colorValue:myColorValue};
	return myColor;
}

function mySetSWFExportPreferences(){
	app.swfExportPreferences.rasterizePages= false;
	app.swfExportPreferences.flattenTransparency = false;
	app.swfExportPreferences.dynamicMediaHandling = DynamicMediaHandlingOptions.includeAllMedia;
	app.swfExportPreferences.frameRate = 24;	
}
