var pptxgenjs = require("pptxgenjs");
const { app, BrowserWindow, ipcMain, dialog } = require("electron");

app.on("ready", createWindow);

function createWindow() {
  // Create the browser window.
  win = new BrowserWindow({
    width: 800,
    height: 600
  });

  // and load the index.html of the app.
  win.loadFile("index.html");
  win.on("ready-to-show", () => {
    win.show();
  });
}

//catch item add
ipcMain.on("formdata", function(e, data) {
  e.preventDefault();
  console.log(data);
  createPPT(data);
});

/* data = {
  project: "Title",
  answer0: "this is answer 1",
  answer1: "this is answer 2",
  answer2: "this is answer 3",
  answer3: "this is answer 4",
  answer4: "this is answer 5"
};
createPPT(data); */

createPPT = data => {
  //function createPPT(data) {
  console.log("ppt");
  var questions = [
    "Lorem ipsum dolor sit amet, iudico doming cum an, nemore posidonium constituam cu vis. Mea ei hinc nemore?",
    "Lorem ipsum dolor sit amet, iudico doming cum an, nemore posidonium constituam cu vis. Mea ei hinc nemore?",
    "Lorem ipsum dolor sit amet, iudico doming cum an, nemore posidonium constituam cu vis. Mea ei hinc nemore?",
    "Lorem ipsum dolor sit amet, iudico doming cum an, nemore posidonium constituam cu vis. Mea ei hinc nemore?",
    "Lorem ipsum dolor sit amet, iudico doming cum an, nemore posidonium constituam cu vis. Mea ei hinc nemore?"
  ];
  var pptx = new pptxgenjs();
  //---------------------------------------------------------------------------------------------------
  //Template design
  pptx.defineSlideMaster({
    title: "MASTER_SLIDE",
    bkgd: "dcdce8",
    objects: [
      {
        line: {
          x: 5,
          y: 0.8,
          w: 0,
          h: 3.5,
          line: "0303bf",
          lineSize: 1
        }
      },
      {
        rect: {
          x: 0.1,
          y: 0.1,
          w: 9.8,
          h: "96%",
          line: "0303bf",
          lineSize: 1
        }
      }
      //    { image: { x: 6, y: 0, w: 1.67, h: 0.75, path: "images/logo.png" } }
    ]
  });

  var slide = pptx.addNewSlide("MASTER_SLIDE");
  //template design ends
  //---------------------------------------------------------------------------------------------------
  //var slide = pptx.addNewSlide();
  //Title
  slide.addText(data.project, {
    x: 0.2,
    y: 0.1,
    w: "98%",
    valign: "top",
    fontSize: 14,
    bold: "true",
    color: "0303bf"
  });
  //Question and Answers
  slide.addText(questions[0], {
    x: 0.2,
    y: 0.5,
    w: "48%",
    valign: "top",
    fontSize: 8,
    autoFit: "true",
    color: "0303bf"
  });
  slide.addText(data.answer0, {
    x: 0.2,
    y: 0.8,
    w: "48%",
    valign: "top",
    fontSize: 8,
    autoFit: "true",
    color: "000000"
  });
  slide.addText(questions[1], {
    x: 0.2,
    y: 2.0,
    w: "48%",
    valign: "top",
    fontSize: 8,
    autoFit: "true",
    color: "0303bf"
  });
  slide.addText(data.answer1, {
    x: 0.2,
    y: 2.3,
    w: "48%",
    valign: "top",
    fontSize: 8,
    autoFit: "true",
    color: "000000"
  });
  slide.addText(questions[2], {
    x: 0.2,
    y: 3.3,
    w: "48%",
    valign: "top",
    fontSize: 8,
    autoFit: "true",
    color: "0303bf"
  });
  slide.addText(data.answer2, {
    x: 0.2,
    y: 3.6,
    w: "48%",
    valign: "top",
    fontSize: 8,
    autoFit: "true",
    color: "000000"
  });
  slide.addText(questions[3], {
    x: 5.0,
    y: 0.5,
    w: "48%",
    valign: "top",
    fontSize: 8,
    autoFit: "true",
    color: "0303bf"
  });
  slide.addText(data.answer3, {
    x: 5.0,
    y: 0.8,
    w: "48%",
    valign: "top",
    fontSize: 8,
    autoFit: "true",
    color: "000000"
  });
  slide.addText(questions[4], {
    x: 5.0,
    y: 3.0,
    w: "48%",
    valign: "top",
    fontSize: 8,
    autoFit: "true",
    color: "0303bf"
  });
  slide.addText(data.answer4, {
    x: 5.0,
    y: 3.3,
    w: "48%",
    valign: "top",
    fontSize: 8,
    autoFit: "true",
    color: "000000"
  });
  pptx.save("CaseStudy-" + data.project, function(filename) {
    console.log("Created: " + filename);
    dialog.showMessageBox(win, {
      title: "Message",
      buttons: [],
      message: "PPT created with filename " + filename
    });
    return;
  });
};
