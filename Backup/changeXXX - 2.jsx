///////////    Version 2   ///////////

/////////////////////////////////////////////////////////////////////////////////////////////////////    LOG    /////////////////////////////////////////////////////////////////////////////////////////////////////

///////////    Version 1    ///////////    2021-8-20    ///////////
//added changeXXX()

///////////    Version 2    ///////////    2021-10-5    ///////////
//cleared find change preferences before excute the event handler

/////////////////////////////////////////////////////////////////////////////////////////////////////    LOG    /////////////////////////////////////////////////////////////////////////////////////////////////////
#targetengine "session"

if(Number(app.activeDocument.name.slice(0, 3)) > 0) {
    changeXXX()
} else {
    app.activeDocument.eventListeners.everyItem().remove()
    app.activeDocument.eventListeners.add("afterSave", changeXXX)
    app.activeDocument.eventListeners.add("afterSaveAs", changeXXX)
}

function changeXXX() {
    app.findGrepPreferences = NothingEnum.NOTHING;
    app.changeGrepPreferences = NothingEnum.NOTHING;    
    
    app.findChangeGrepOptions.properties = ({includeMasterPages: true, kanaSensitive:true, widthSensitive:true});
    app.findGrepPreferences.findWhat = "(?<=第)xxx";
    app.changeGrepPreferences.changeTo = app.activeDocument.name.slice(0, 3).toString()
    app.activeDocument.changeGrep();
    
    app.findGrepPreferences.findWhat = "(?<=G.N. )xxx";
    app.changeGrepPreferences.changeTo = app.activeDocument.name.slice(0, 3).toString()
    app.activeDocument.changeGrep();
    
    app.findGrepPreferences.findWhat = "xxx(?=\n-)";
    app.changeGrepPreferences.changeTo = app.activeDocument.name.slice(0, 3).toString()
    app.activeDocument.changeGrep();

    app.activeDocument.eventListeners.everyItem().remove()
}
