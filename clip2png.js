/* 
 * clip2png.js
 *
 * Copyright (c) 2016 Christopher Crawford
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
(function(ws) {
    var thisIsa32BitSystem = function() {
  	var wmi = GetObject("winmgmts:root\\cimv2");
  	var osSettingsCollection = wmi.ExecQuery("select * from Win32_OperatingSystem");
  	var i = new Enumerator(osSettingsCollection);
  	var config = i.item();
  	if( /32-bit/.test(config.OSArchitecture) ){
  	    return true;
  	}
  	else{
  	    return false;
	}
    };
    
    if(!thisIsa32BitSystem()){
        if(!/SysWOW64/.test(ws.Path)){
            var cmd = 'C:\\Windows\\SysWOW64\\cscript.exe //nologo "' + ws.scriptFullName + '"';
            var args = ws.arguments;
            for (var i = 0, len = args.length; i < len; i++) {
                var arg = args(i);
                cmd += ' ' + (~arg.indexOf(' ') ? '"' + arg + '"' : arg);
            }
            new ActiveXObject('WScript.Shell').run(cmd,0);
            ws.quit();	  
        }
    }
    
})(WScript);

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

function print(s) {
    WScript.Echo(s);
}

function printObj(obj) {
    WScript.Echo(JSON.stringify(obj));
}

function extract(o){
    var SzPath = "bin\\7z1602-extra";
    var SzExe = "7za.exe";
    var command = "e";
    var archivePath = o.to;
    var switches = "-o" + archivePath;
    var archiveName = o.from;
    var filePath = "word\\media";
    var fileName = o.file;
    var cmd = [ "cmd /c",
                [SzPath,SzExe].join("\\"),
                command,
                switches,
                archiveName,
                [filePath,fileName].join("\\")
              ].join(" ");

    var exec = function(cmd){
	var shell = new ActiveXObject("WScript.Shell");
	var results = shell.Run(cmd,0,true);
    };

    exec(cmd);
}

var WORD = (function () {
    var instance;

    function createInstance() {
        var object = new ActiveXObject("Word.Application");
        return object;
    }
    
    return {
        getInstance: function () {
            try{
                instance.Version;
            }
            catch(e){
                instance = createInstance();
                return instance;
            }
            return instance;
        },
        Quit: function(){
            try{
                instance.Version;
            }
            catch(e){
                instance = null
                return;
            }
            var saveChanges = false;
            instance.Quit(saveChanges);
            instance = null;
            return;
        }
    };
})();

function initFolder(f){
    var fs = FS();
    if(fs.FolderExists(f)){
        return;
    }
    try {
        fs.CreateFolder(f);
    }
    catch(e){
        initFolder(f.split("\\").reverse().slice(1).reverse().join("\\"));
        fs.CreateFolder(f);
    }
}

function delFolder(f){
    var fs = FS();
    if(!fs.FolderExists(f)){
        return;
    }
    fs.DeleteFolder(f,true);
}

function getNewFilename(screenshots) {
    var fs = FS();
    var folder;
    try {
        folder = fs.GetFolder(screenshots);
    }
    catch(e) {
        initFolder(screenshots);
        folder = fs.GetFolder(screenshots);
    }
    var files = new Enumerator(folder.Files);
    var l = [];
    while (!files.atEnd()) {
	var f = files.item();
	if (f.Name.match(/\d{3}\.png$/gi)) {
	    l.push(f.Name.split(".")[0]);
	}
	files.moveNext();
    }

    if(l.length === 0) {
	l.push("000");
    }

    var a = l.sort( function(a,b) { return a-b; } );
    var b = a.pop();
    var c = parseInt(b,10) + 1;

    var pad = function(num, size) {
	var s = "00" + num;
	s = s.split('').reverse().join('');
	s = s.substr(0,size)
	return s.split('').reverse().join('');
    };
    
    var d = pad(c.toString(),3) + ".png";
    return fs.BuildPath(folder.Path, d);
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

function getCaption(){
    var s = VB.InputBox("Write a complete sentence about this screenshot.");
    if(s.length===0){
        VB.MsgBox("You cannot leave this empty.",16,"Not so fast...");
        return getCaption();
    }
    else{
        return s;
    }
}

function getTimeEstimate(){
    var n = VB.InputBox("How many minutes has it been since the last step?");
    if(isNaN(n) || n.length===0){
        VB.MsgBox("You must provide a number.",16,"Not so fast...");
        return getTimeEstimate();
    }
    else{
        return n;
    }
}

function REPORT(path){
    this.path = path;
    var fs = FS();
    var report;

    if(!fs.FileExists(this.path)){
        this.report = [];
    }
    else {
        var s = readFile(this.path);
        this.report = JSON.parse(s);
    }

    this.writeToFile = function(){
        var f = fs.OpenTextFile(this.path, 2, true);
        f.Write(JSON.stringify(this.report));
        f.Close();
    };

    this.add = function(entry){
        this.report.push(entry);
    };
}

function DOC(){
    this.doc = WORD.getInstance().Documents.Add();
    this.IMG_MAX_HEIGHT = 3.71; // inches
    this.IMG_MAX_WIDTH = 6; // inches
    this.POINTS_PER_INCH = 72;

    this.paste = function() {
        this.doc.Range().Paste();
    };

    this.save = function(path){
        this.doc.SaveAs2(path);
    };

    this.close = function(){
        var saveChanges = false;
        this.doc.Close(saveChanges);
        //if no more docs, quit Word
        if(! WORD.getInstance().Documents.Count > 0){
            WORD.Quit();
        }
    };

    this.hasAnImage = function(){
        if(this.doc.InlineShapes.Count > 0){
            return true;
        }
	else {
	    return false;
	}
    };

    this.imgWidth = function(){
        if(this.hasAnImage()){
            return this.doc.InlineShapes.Item(1).Width / this.POINTS_PER_INCH;
        }
        else {
            return 0;
        }
    };

    this.imgHeight = function(){
        if(this.hasAnImage()){
            return this.doc.InlineShapes.Item(1).Height / this.POINTS_PER_INCH;
        }
        else {
            return 0;
        }
    };

    this.checkDimensions = function(){
        if(this.imgWidth() > this.IMG_MAX_WIDTH){
            throw "ImageTooWide";
        }
        if(this.imgHeight() > this.IMG_MAX_HEIGHT){
            throw "ImageTooTall";
        }
    }

}

function clip2png(){
    var config = initConfig();
    var doc = new DOC();
    doc.paste();
    if(doc.hasAnImage()){
	try {
	    doc.checkDimensions();
	}
	catch(e){
	    var s1 = "Could not save your screenshot.";
	    var imgIsToo = function(adjective,n,nMax){
		return [
		    "The picture is too " + adjective + ".  ",
		    "(" + n + " inches " + adjective + ".  ",
		    "Needs to be " + nMax + " inches or below.)"
		].join("");
	    };
	    if(e==="ImageTooTall"){
		var s2 = imgIsToo("tall", doc.imgHeight(), doc.IMG_MAX_HEIGHT);
		VB.MsgBox(s2,16,s1);
	    }
	    if(e==="ImageTooWide"){
		var s2 = imgIsToo("wide", doc.imgWidth(), doc.IMG_MAX_WIDTH);
		VB.MsgBox(s2,16,s1);
	    }
	    doc.close();
	    WScript.Quit();
	}
	initFolder(config.tmp);
	doc.save(config.tmpDoc);
	doc.close();
	
	extract({'file': config.tmpPicName,
		 'from': config.tmpDoc,
		 'to'  : config.tmp});
        
	var newFilePath = "";
	if(FS().FileExists(config.tmpPicPath)){
	    newFilePath = getNewFilename(config.pixPath);
	}
	else{
	    VB.MsgBox("There was no picture on your clipboard!",16,"Could not save your screenshot.");
	    delFolder(config.tmp);
	    WScript.Quit();
	}
	FS().MoveFile(config.tmpPicPath, newFilePath);
	delFolder(config.tmp);
	var report = new REPORT(config.report);
	var newEntry = {};
	newEntry.screenshot = newFilePath.split("\\").pop();
	newEntry.caption = getCaption();
	newEntry.minutes = getTimeEstimate();
	report.add(newEntry);
	report.writeToFile();
    }
    else {
        VB.MsgBox("There was no picture on your clipboard!",16,"Could not save your screenshot.");
        doc.close();
        WScript.Quit();
    }
}

function main() {
    clip2png();
}

main();
