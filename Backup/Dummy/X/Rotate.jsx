main()
function main(){
    var myDocument = app.documents.add();
    var myText = myDocument.textFrames.add();
    myText.geometricBounds = [20, 100, 100, 20];
    myText.contents = "Testing";
    var rotate90 = app.transformationMatrices.add({counterclockwiseRotationAngle: 90});
    var rotate180 = app.transformationMatrices.add({counterclockwiseRotationAngle: 180});
    var rotate270 = app.transformationMatrices.add({counterclockwiseRotationAngle: 270});
    myText.transform (CoordinateSpaces.pasteboardCoordinates, AnchorPoint.centerAnchor, rotate180);
}
