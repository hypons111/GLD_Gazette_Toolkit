///////////    Version 9    ///////////

/////////////////////////////////////////////////////////////////////////////////////////////////////    LOG    /////////////////////////////////////////////////////////////////////////////////////////////////////

///////////    Version 1    ///////////    2021-5-25    ///////////
//Strart

///////////    Version 2    ///////////    2021-5-27    ///////////
//added a function make the year & volume plus 1 when 31th of December

///////////    Version 3    ///////////    2021-5-27    ///////////
//make script looks better

///////////    Version 4    ///////////    2021-5-27    ///////////
//fix double struture which i dont know what its for 

///////////    Version 5    ///////////    2021-5-28    ///////////
//show total exported pdf number at the finish alert

///////////    Version 6    ///////////    2021-6-17    ///////////
//fixed chinese version gn number with 5 digit

///////////    Version 7    ///////////    2021-6-17    ///////////
//fixed struture tag after gn number, takes me 5hrs to fix this useless piece of shit

///////////    Version 8    ///////////    2021-6-23    ///////////
//fixed when has 第xxx號 at the top of the contents

///////////    Version 9    ///////////    2021-7-27    ///////////
//added getExportPreset() to automate add bookmark and remove tag on acorbat in indesign

/////////////////////////////////////////////////////////////////////////////////////////////////////    SCRIPT    /////////////////////////////////////////////////////////////////////////////////////////////////////

const activeDocument = app.activeDocument
var GNs = []
var version = getVersion()
const year = getYear(new Date().getMonth() + 1, new Date().getDate())
const volume = Number(year) - 1996
const issue = activeDocument.name.toString().slice(3, 5)
var gnNumber = []
const presetList = []


for(var x = 0; x < activeDocument.pages.length; x ++) {
    ungroup(activeDocument.pages[x])
}


for(var x = 0; x < activeDocument.pages.length; x ++) {
    for(var y = 0; y < activeDocument.pages[x].textFrames.length; y ++) {
        if(activeDocument.pages[x].textFrames[y].contents.toString().slice(0, 5) === "G.N. ") {
            getGnNumber(activeDocument.pages[x].textFrames[y].contents.toString().slice(5, 10))
         } else if(activeDocument.pages[x].textFrames[y].contents.toString().slice(1, 6) === "G.N. ") {
            getGnNumber(activeDocument.pages[x].textFrames[y].contents.toString().slice(6, 11))
         } else if(activeDocument.pages[x].textFrames[y].contents.toString().slice(2, 7) === "G.N. ") {
            getGnNumber(activeDocument.pages[x].textFrames[y].contents.toString().slice(7, 12))
         } else if(activeDocument.pages[x].textFrames[y].contents.toString().slice(0, 1) === "第") {
            getGnNumber(activeDocument.pages[x].textFrames[y].contents.toString().slice(1, 7))
        } else if(activeDocument.pages[x].textFrames[y].contents.toString().slice(1, 2) === "第") {
            getGnNumber(activeDocument.pages[x].textFrames[y].contents.toString().slice(2, 8))
        }
    } 
}

function getExportPreset() {
    for(var x = 0; x < GNs.length; x ++) {
        
        var deleteSwitch = true
        
        for(var y = Number(GNs[x] - 1); y < Number(GNs[x + 1] - 1); y ++) {
            for(var z = 0; z < activeDocument.pages[y].allGraphics.length; z ++) {
                if(activeDocument.pages[y].allGraphics[z].imageTypeName === "TIFF"){
                    deleteSwitch = false
                    break
                }
            }
        }

        if(deleteSwitch === true) {
            presetList.push(app.pdfExportPresets.item('[大報]'))
        } else{
            presetList.push(app.pdfExportPresets.item('[大報有標籤]'))
        }
    }
}

//const exportPreset = app.pdfExportPresets.item('eGaz_PDF');
//const exportPresetRemoveTag = app.pdfExportPresets.item('eGaz_PDF delete 標籤');

getExportPreset()


for(var x = 0; x < GNs.length; x ++) {
    if(x === GNs.length - 1) {
        app.pdfExportPreferences.pageRange = GNs[x].toString() + "-"
    } else if(GNs[x] === (GNs[x + 1] - 1)) {
        app.pdfExportPreferences.pageRange = GNs[x].toString()
    } else {
        app.pdfExportPreferences.pageRange = GNs[x].toString() + "-" + (GNs[x + 1] - 1).toString()
    }
    addBookmarks(gnNumber[x])
    setDocumentTitle(gnNumber[x])
    activeDocument.exportFile(ExportFormat.pdfType, File('~/Desktop/Gaz_PDF/' + version + year + volume + issue + gnNumber[x] + '.pdf'), false, presetList[x])
    activeDocument.bookmarks[activeDocument.bookmarks.length - 1].remove()
}

alert(gnNumber.length + "個PDF已儲存在桌面 Gaz_PDF ")


/////////////////////////////////////////////////////////////////////////////////////////////////////    FUNCTION    /////////////////////////////////////////////////////////////////////////////////////////////////////

function getVersion() {
    app.findGrepPreferences.findWhat = "(?<!\\r)(?<=^)第\\d\+號公告$|(?<=^)第\\d\+號公告\\r勞工處\\r破產欠薪保障條例\\(第";
    if(app.activeDocument.findGrep().length > 0) {
        return "cgn"
    } else {
        app.findGrepPreferences.findWhat = "(?<!\\r)(?<=^)G.N. \\d\+$|(?<=^)G.N. \\d\+\\rLabour Department\\rProtection of Wages on Insolvency Ordinance \\(Chapter ";
        if(app.activeDocument.findGrep().length > 0)
        return "egn"
    }
}


function getYear(month, date) {
    if(month === 12 && date === 31) {
        return new Date().getFullYear() + 1
    }
    return new Date().getFullYear()
}

function ungroup(page) {
    for(var y = 0; y < page.groups.length; y ++) {
        page.groups[y].ungroup()
    }
}    


function getGnNumber(string) {
    if(version === "egn") {
        switch("\r") {
            case string.slice(1, 2) :
                if(Number(string.slice(0, 1)) >= 0) {
                    gnNumber.push(string.slice(0, 1))
                } else {
                    gnNumber.push(string.slice(0, 0))
                }
                break;
            case string.slice(2, 3) :
                if(Number(string.slice(1, 2)) >= 0) {
                    gnNumber.push(string.slice(0, 2))
                } else {
                    gnNumber.push(string.slice(0, 1))
                }
                break;
            case string.slice(3, 4) :
                if(Number(string.slice(2, 3)) >= 0) {
                    gnNumber.push(string.slice(0, 3))
                } else {
                    gnNumber.push(string.slice(0, 2))
                }
                break;
            case string.slice(4, 5) :
                if(Number(string.slice(3, 4)) >= 0) {
                    gnNumber.push(string.slice(0, 4))
                } else {
                    gnNumber.push(string.slice(0, 3))
                }
                break;

            default:
                if(Number(string.slice(4, 5)) >= 0) {
                    gnNumber.push(string.slice(0, 5))
                } else {
                    gnNumber.push(string.slice(0, 4))
                }
        }
        GNs.push(x + 1)
    } else if(version === "cgn") {
        switch(string.slice(4, string.length - 1)) {
            case "號" :
                if(string.slice(5, string.length) === "公") {
                    gnNumber.push(string.slice(0, string.length - 2))
                    GNs.push(x + 1)
                }
                break;
            case "公" :
                gnNumber.push(string.slice(0, string.length - 3))
                GNs.push(x + 1)
                break;
            case "告" :
                gnNumber.push(string.slice(0, string.length - 4))
                GNs.push(x + 1)
                break;
            case "\r" :
                gnNumber.push(string.slice(0, string.length - 5))
                GNs.push(x + 1)
                break;
            default:
                gnNumber.push(string)
        }
    }
}


function addBookmarks(gnNumber) {
    var bookmarkName = "第" + gnNumber + "號公告"
    if(version === "egn") {
        bookmarkName = "G.N. " + gnNumber
    }
    activeDocument.bookmarks.add(activeDocument.pages[GNs[x] - 1])
    activeDocument.bookmarks[0].name = bookmarkName
}


function setDocumentTitle(gnNumber) {
    activeDocument.metadataPreferences.documentTitle = version + year + volume + issue + gnNumber
}