main();


function main(){
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
	var myFilteredFiles;
	var myFolder = Folder.selectDialog("", ""); 
	if(myFolder != null){ 
        myFilteredFiles = openFile(myFolder);
    }
    end();
}


function openFile(myFolder){
	var myFiles = myFolder.getFiles("*.indd"); 
    if(myFiles.length != 0){ 
        for(var i = 0; i < myFiles.length; i++){ 
            app.open(File(myFolder+ "/" + myFiles[i].name))
            for(var x = 0; x <  app.activeDocument.pages.length; x ++) {
                app.activeDocument.pages[x].groups.everyItem().ungroup()
                for(var y = 0; y <  app.activeDocument.pages[x].textFrames.length; y ++) {
                    deleteOversetedTextFrame(app.activeDocument.pages[x].textFrames[y])
                }
            }
            deleteBlankPage()
            app.activeDocument.save();
            app.activeDocument.close();
        } 
    } 
}


function deleteOversetedTextFrame(textFrame) {
    var selection = textFrame;
	try {
        if(selection.parentStory.textContainers.length > 1){
            for(var i = selection.parentStory.textContainers.length - 1; i > 0; i --){
                selection.parentStory.textContainers[i].remove();
            }
        } 
    } catch (e) {}
}


function deleteBlankPage() {
    for(var x = app.activeDocument.pages.length - 1; x >= 0; x --) {
        if(app.activeDocument.pages[x].textFrames.length === 0){
            app.activeDocument.pages[x].remove()
        }
    }
}


function end() {
    alert("已完成")
}