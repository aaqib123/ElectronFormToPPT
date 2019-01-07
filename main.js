var pptxgenjs = require("pptxgenjs");
const { app, BrowserWindow, ipcMain } = require("electron");

app.on("ready", createWindow);

function createWindow() {
  // Create the browser window.
  win = new BrowserWindow({ width: 800, height: 600 });

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

createPPT = data => {
  console.log("ppt");
  var questions = [
    "Lorem ipsum dolor sit amet, iudico doming cum an, nemore posidonium constituam cu vis. Mea ei hinc nemore?",
    "Lorem ipsum dolor sit amet, iudico doming cum an, nemore posidonium constituam cu vis. Mea ei hinc nemore?",
    "Lorem ipsum dolor sit amet, iudico doming cum an, nemore posidonium constituam cu vis. Mea ei hinc nemore?",
    "Lorem ipsum dolor sit amet, iudico doming cum an, nemore posidonium constituam cu vis. Mea ei hinc nemore?",
    "Lorem ipsum dolor sit amet, iudico doming cum an, nemore posidonium constituam cu vis. Mea ei hinc nemore?"
  ];
  var pptx = new pptxgenjs();

  var slide = pptx.addNewSlide();
  //Title
  slide.addText(data.project, {
    x: 0.2,
    y: 0.0,
    w: "98%",
    valign: "top",
    fontSize: 12,
    bold: "true",
    color: "000000"
  });
  //Question and Answers
  slide.addText(questions[0], {
    x: 0.2,
    y: 0.5,
    w: "48%",
    valign: "top",
    fontSize: 8,
    autoFit: "true",
    color: "000000"
  });
  slide.addText(data.answer1, {
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
    color: "000000"
  });
  slide.addText(data.answer2, {
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
    y: 3.0,
    w: "48%",
    valign: "top",
    fontSize: 8,
    autoFit: "true",
    color: "000000"
  });
  slide.addText(data.answer3, {
    x: 0.2,
    y: 3.3,
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
    color: "000000"
  });
  slide.addText(data.answer4, {
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
    color: "000000"
  });
  slide.addText(data.answer5, {
    x: 5.0,
    y: 3.3,
    w: "48%",
    valign: "top",
    fontSize: 8,
    autoFit: "true",
    color: "000000"
  });
  pptx.save("CaseStudy", function(filename) {
    console.log("Created: " + filename);
  });
};
