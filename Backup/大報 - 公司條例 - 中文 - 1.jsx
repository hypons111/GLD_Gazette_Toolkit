﻿///////////    Version 1    ///////////

/////////////////////////////////////////////////////////////////////////////////////////////////////    LOG    /////////////////////////////////////////////////////////////////////////////////////////////////////

///////////    Version 1    ///////////    2021-5-31    ///////////
//can get artwork path now

/////////////////////////////////////////////////////////////////////////////////////////////////////    LOG    /////////////////////////////////////////////////////////////////////////////////////////////////////

const activeDocument = app.activeDocument
const header = "第xxx號公告\r公司註冊處\r公司條例\(第622章\)\r"
const sections = [
  "在《公司條例》(第622章) (下稱「該條例」)附表9的生效日期前不時有效的《前身公司條例》，其第291(3)條已被廢除，但在被廢除後，根據該條例附表11第129條，具有持續效力。現依據《前身公司條例》第291(3)條公布，在本公告刊登當日起計三個月期滿時，除非有相反因由提出，否則下列公司的名稱將從登記冊上剔除，而公司亦予以解散：", 
  "在《公司條例》(第622章) (下稱「該條例」)附表9的生效日期前不時有效的《前身公司條例》，其第291(6)條已被廢除，但在被廢除後，根據該條例附表11第129條，具有持續效力。現依據《前身公司條例》第291(6)條公布，下列公司的名稱已從公司登記冊上剔除，並由本公告刊登當日起予以解散：", 
  "在《公司條例》(第622章) (下稱「該條例」)附表9的生效日期前不時有效的《前身公司條例》，其第291AA(7)條已被廢除，但在被廢除後，根據該條例附表11第129條，具有持續效力。現依據《前身公司條例》第291AA(7)條公布，除非處長在本公告刊登日期後三個月內收到對下列公司的撤銷註冊而提出的反對，否則處長可將下列公司的註冊撤銷和解散下列公司：", 
  "在《公司條例》(第622章) (下稱「該條例」)附表9的生效日期前不時有效的《前身公司條例》，其第291AA(9)條已被廢除，但在被廢除後，根據該條例附表11第129條，具有持續效力。現依據《前身公司條例》第291AA(9)條公布，下列公司的註冊在本公告刊登當日撤銷，而下列公司亦在註冊撤銷時解散：", 
  "現根據《公司條例》第744(3)條公布，除非有反對因由提出，否則在本公告的日期後的3個月終結時，下列公司的名稱將會從公司登記冊剔除，而下列公司將會解散：", 
  "現根據《公司條例》第745(2)(b)條公布，除非有反對因由提出，否則在本公告的日期後的3個月終結時，下列公司的名稱將會從公司登記冊剔除，而下列公司將會解散：", 
  "現根據《公司條例》第746(2)條公布，下列公司的名稱已從公司登記冊剔除，而下列公司由本公告刊登當日起即告解散：", 
  "現根據《公司條例》第747(2)條公布，除非有反對因由提出，否則在本公告的日期後的3個月終結時，下列公司的名稱將會從公司登記冊剔除，而下列公司將會解散：", 
  "現依據《公司條例》第751(1)條公布，除非處長在本公告刊登的日期後3個月內收到對下列公司的撤銷註冊的反對，否則處長可撤銷下列公司的註冊：", 
  "公司註冊處處長現依據《公司條例》第751(3)條宣布，下列公司的註冊在本公告刊登當日撤銷，而依據《公司條例》第751(6)條，下列公司亦在註冊撤銷時即告解散：", 
  "現根據《公司條例》第796(3)條公布，除非有反對因由提出，否則在本公告的日期後的3個月終結時，下列公司的名稱將會從公司登記冊剔除，而下列公司將不再是註冊非香港公司：", 
  "現根據《公司條例》第797(2)(b)條公布，除非有反對因由提出，否則在本公告的日期後的3個月終結時，下列公司的名稱將會從公司登記冊剔除，而下列公司將不再是註冊非香港公司：", 
  "現根據《公司條例》第798(2)條公布，下列公司的名稱已從公司登記冊剔除，而下列公司即不再是註冊非香港公司：", 
]
const sectionsNumber = ["291(3)", "291(6)", "291AA(7)", "291AA(9)", "744(3)", "745(2)(b)", "746(2)", "747(2)", "751(1)", "751(3)", "796(3)", "797(2)(b)", "798(2)"];


makeDialog();
dialog.center(); 
if(dialog.show() === 1){
  type = dialog.kindType.selection.index;
}else{
  exit();
}
activeDocument.pages[0].textFrames[0].contents = header + sections[type] + "\r"





var cPDFs = getPDFs(Folder(activeDocument.fullName.toString().slice(0, 49) + '_Artwork/011'))

function getPDFs(artworkFolder) {
	if(artworkFolder !== null){ 
        return artworkFolder.getFiles("*"+ ".pdf"); 
	}
}

for(var x = 0; x < cPDFs.length - 1; x ++) {
    activeDocument.pages[0].textFrames[0].place(File(cPDFs[x])); 
}




function makeDialog()
{
  dialog = new Window('dialog', "請選擇條例編號", "x:0, y:0, width:310, height:80");
  dialog.panel = dialog.add('panel', [10, 10, 220, 70], "");

  dialog.panel.add('statictext', [10, 20, 110, 45], "請選擇條例編號：");
  dialog.kindType = dialog.panel.add('dropdownlist', [115, 15, 195, 40]);
  dialog.kindType.onChange = sectionsNumberValidator;
  for(var x = 0; x < sectionsNumber.length; x ++){
    dialog.kindType.add('item', sectionsNumber[x]);
  }
 

    dialog.OKbut=dialog.add('button', [225, 10, 300, 35], "OK");
    dialog.OKbut.onClick=onOKclicked;
    dialog.CANbut=dialog.add('button', [225, 45, 300, 70], "Cancel");
    dialog.CANbut.onClick=onCANclicked;
  return dialog;
}

function sectionsNumberValidator(){
  kind(dialog.kindType, placementINFO.pgCount, "條例編號");
}
function onOKclicked(){
  dialog.close(1);
}
function onCANclicked(){
  dialog.close(0);
}