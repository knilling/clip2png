/* 
 * png2ppt.js
 *
 * Copyright (c) 2018 Christopher Crawford
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 *
 */

function FS(){
    return new ActiveXObject("Scripting.FileSystemObject");
}

function readFile(f) {
    return FS().OpenTextFile(f,1).ReadAll();
}

eval(readFile("lib\\includelib.js"));
include("lib\\json3.min.js");
include("lib\\underscore-min.js");
include("lib\\msgbox.js");

var MSCONST = (function () {
    var stuff = {
        ppLayoutBlank     : 12,
        ppLayoutTitle     : 1,
        ppLayoutTitleOnly : 11,
        msoLineSingle     : 1,
        msoFalse          : 0,
        msoTrue           : -1
    };

    return {
        get: function(name) { return stuff[name]; }
    };

})();

function addBorder(pic){
    pic.Line.Weight = 0.25;
    pic.Line.Style = MSCONST.get('msoLineSingle');
}

function setMargin(pictureUnits,availableUnits,marginEdge,slideUnits){
    var margin = 0;
    if(pictureUnits > availableUnits) {
        margin = marginEdge;
    }
    else {
        var a  = (slideUnits - marginEdge) / 2;
        var b  = slideUnits - a;
        var c  = pictureUnits / 2;
        margin = b - c;
    }
    return margin;
}

function center(pic,slide,sHeight,sWidth){
    var pHeight     = pic.Height;
    var pWidth      = pic.Width;    
    var title       = slide.Shapes.title;
    var tHeight     = title.Height;
    var tTop        = title.Top;
    var tBottom     = tHeight + tTop;
    var tLeft       = title.Left;
    var tWidth      = title.Width;
    //var tRight      = tLeft + tWidth;

    // Available Height
    var aHeight     = sHeight - tBottom - 1;
    var aHeight_Top = sHeight - aHeight;

    pic.Top  = setMargin(pHeight,aHeight,aHeight_Top,sHeight);
    pic.Left = setMargin(pWidth,sWidth,0,sWidth);
}

function addSlide(ppt){
    var newSlidePosition = ppt.Slides.Count + 1;
    var slide = ppt.Slides.Add(newSlidePosition,MSCONST.get('ppLayoutTitleOnly'));
    return slide;
}

function insertPic(pic,slide) {
    var s                = slide.Shapes;
    var linkToFile       = MSCONST.get('msoFalse');
    var saveWithDocument = MSCONST.get('msoTrue');
    var left             = 0;
    var top              = 0;
    var picOnSlide       = s.AddPicture( pic, linkToFile, saveWithDocument, left, top );
    return picOnSlide;
}

function initConfig(){
    var s = readFile("settings.json");
    var config = JSON.parse(s);

    var fixPath = function(s){
    var re = /\|/gi;
    return s.replace(re,'\\');
    }
    
    config.fullPath = fixPath(config.fullPath);
    config.projectPath = config.fullPath + "\\" + config.projectName;
    config.tmp = config.projectPath + "\\tmp"
    config.tmpDoc = config.tmp + "\\tmp.docx";
    config.tmpPicName = "image1.png";
    config.tmpPicPath = config.tmp + "\\" + config.tmpPicName;
    config.pixPath = config.projectPath + "\\screenshots";
    config.report = config.projectPath + "\\report.json";
    config.cwd = new ActiveXObject("WScript.Shell").CurrentDirectory;
    config.bin = config.cwd + "\\bin";
    return config;
}

function insertTitle(text,slide) {
    slide.Shapes.Title.TextFrame.TextRange.Text = text;
}

function png2ppt(){
    var powerpnt     = new ActiveXObject("PowerPoint.Application")
    powerpnt.Visible = true;
    var ppt          = powerpnt.Presentations.Add();
    var sHeight      = ppt.PageSetup.SlideHeight;
    var sWidth       = ppt.PageSetup.SlideWidth;
    var config       = initConfig();
    var s            = readFile(config.report);
    var report       = JSON.parse(s);

    var i = new Enumerator(report);
    while (!i.atEnd()) {
        var note = i.item();
        var slide = addSlide(ppt);
        var fileName = config.pixPath + "\\" + note.screenshot;
        var title = note.caption;
        insertTitle(note.caption,slide);
        var pic = insertPic(fileName,slide)
        addBorder(pic);
        center(pic,slide,sHeight,sWidth);
        i.moveNext();
    }

}

function main() {
    png2ppt();
}

main();
