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
  var pptx = new pptxgenjs();
  pptx.defineSlideMaster({
    title: "MASTER_SLIDE",
    bkgd: "FFFFFF",
    objects: [
      { line: { x: 3.5, y: 1.0, w: 6.0, line: "0088CC", lineSize: 5 } },
      { rect: { x: 0.0, y: 5.3, w: "100%", h: 0.75, fill: "F1F1F1" } },
      {
        text: {
          text: "Status Report",
          options: { x: 3.0, y: 5.3, w: 5.5, h: 0.75 }
        }
      },
      { image: { x: 11.3, y: 6.4, w: 1.67, h: 0.75, path: "images/logo.png" } }
    ],
    slideNumber: { x: 0.3, y: "90%" }
  });
  var slide = pptx.addNewSlide();

  var opts = { x: 1.0, y: 1.0, fontSize: 42, color: "00FF00" };

  slide.addText(data.answer1, opts);
  pptx.save("CaseStudy", function(filename) {
    console.log("Created: " + filename);
  });
};
