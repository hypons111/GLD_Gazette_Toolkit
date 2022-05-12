///////////    Version 2    ///////////

/////////////////////////////////////////////////////////////////////////////////////////////////////    LOG    /////////////////////////////////////////////////////////////////////////////////////////////////////

///////////    Version 1    ///////////    2021-5-25    ///////////
//Strart

///////////    Version 2    ///////////    2021-5-27    ///////////
//added a function make the year & volume plus 1 when 31th of December


/////////////////////////////////////////////////////////////////////////////////////////////////////    SCRIPT    /////////////////////////////////////////////////////////////////////////////////////////////////////

const activeDocument = app.activeDocument
const exportPreset = app.pdfExportPresets.item('eGaz_PDF');
var GNs = []
var version = ""
var year = adjustVolumeAndYear(new Date().getMonth() + 1, new Date().getDate())
var volume = Number(year) - 1996
const issue = activeDocument.name.toString().slice(3, 5)
var gnNumber = []



for(var x = 0; x < activeDocument.pages.length; x ++) {
    ungroup(activeDocument.pages[x])
}


for(var x = 0; x < activeDocument.pages.length; x ++) {
    for(var y = 0; y < activeDocument.pages[x].textFrames.length; y ++) {
        if(activeDocument.pages[x].textFrames[y].contents.toString().slice(0, 5) === "G.N. ") {
            version = "egn"
            getGnNumber(activeDocument.pages[x].textFrames[y].contents.toString().slice(5, 10))
            GNs.push(x + 1)
         } else if(activeDocument.pages[x].textFrames[y].contents.toString().slice(1, 6) === "G.N. ") {
            version = "egn"
            getGnNumber(activeDocument.pages[x].textFrames[y].contents.toString().slice(6, 11))
            GNs.push(x + 1)
         } else if(activeDocument.pages[x].textFrames[y].contents.toString().slice(0, 1) === "第") {
            version = "cgn"
            getGnNumber(activeDocument.pages[x].textFrames[y].contents.toString().slice(1, 6))
            GNs.push(x + 1)
        } else if(activeDocument.pages[x].textFrames[y].contents.toString().slice(1, 2) === "第") {
            version = "cgn"
            getGnNumber(activeDocument.pages[x].textFrames[y].contents.toString().slice(2, 7))
            GNs.push(x + 1)
        }
    } 
}

for(var x = 0; x < GNs.length; x ++) {
    if(x === GNs.length - 1) {
        app.pdfExportPreferences.pageRange = GNs[x].toString() + "-"
    } else if(GNs[x] === (GNs[x + 1] - 1)) {
        app.pdfExportPreferences.pageRange = GNs[x].toString()
    } else {
        app.pdfExportPreferences.pageRange = GNs[x].toString() + "-" + (GNs[x + 1] - 1).toString()
    }
    activeDocument.exportFile(ExportFormat.pdfType, File('~/Desktop/Gaz_PDF/' + version + year + volume + issue + gnNumber[x] + '.pdf'), false, exportPreset)
}

alert("PDF已儲存在桌面 Gaz_PDF ")


/////////////////////////////////////////////////////////////////////////////////////////////////////    FUNCTION    /////////////////////////////////////////////////////////////////////////////////////////////////////

function ungroup(page) {
    for(var y = 0; y < page.groups.length; y ++) {
        page.groups[y].ungroup()
    }
}    

function getGnNumber(string) {
    if(version === "egn") {
        switch("\r") {
            case string.slice(4, 5) :
                gnNumber.push(string.slice(0, string.length - 1))
                break;
            case string.slice(3, 4) :
                gnNumber.push(string.slice(0, string.length - 2))
                break;
            case string.slice(2, 3) :
                gnNumber.push(string.slice(0, string.length - 3))
                break;
            case string.slice(1, 2) :
                gnNumber.push(string.slice(0, string.length - 4))
                break;
            default:
                gnNumber.push(string)
        }
    } else if(version === "cgn") {
        switch(string.slice(4, string.length)) {
            case "號" :
                gnNumber.push(string.slice(0, string.length - 1))
                break;
            case "公" :
                gnNumber.push(string.slice(0, string.length - 2))
                break;
            case "告" :
                gnNumber.push(string.slice(0, string.length - 3))
                break;
            case "\r" :
                gnNumber.push(string.slice(0, string.length - 4))
                break;
        }
    }
}

function adjustVolumeAndYear(month, date) {
    if(month === 12 && date === 31) {
        return new Date().getFullYear() + 1
    }
    return new Date().getFullYear()
}
