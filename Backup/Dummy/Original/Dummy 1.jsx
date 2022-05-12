var myDocument = app.documents.add();
myDocument.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.points;    //millimeters
myDocument.viewPreferences.verticalMeasurementUnits = MeasurementUnits.points;
myDocument.viewPreferences.rulerOrigin = RulerOrigin.pageOrigin;


var master32 = myDocument.masterSpreads.item(0);
master32.rectangles.add({geometricBounds:[45, 15, 71.516, 52.5], strokeWeight:1});


myDocument.masterSpreads.add();
var b = myDocument.masterSpreads.item(1);
var myRightPage = b.pages.item(0);
var myRightFooter = myRightPage.rectangles.add();
myRightFooter.geometricBounds = [728, 70, 742, 528];