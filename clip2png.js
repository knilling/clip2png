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
  }
    
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

function readFile(f) {
    return new ActiveXObject("Scripting.FileSystemObject").
        OpenTextFile(f,1).ReadAll();
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

function formatStep(step) {
    return [ step.step,
             step.caption,
             "Estimated Time to Complete: " + step.minutes,
             "Screenshot: " + step.screenshot
           ].join("\n");
}

function genSteps(n) {
    var l = []
    for(i = 1; i <= n; i++){
        var s = "Step " + i;
        l.push(s);
    }
    return l;
}

function addSteps(report){
    var steps = genSteps(report.length);
    var z = _.zip(steps,report);
    return _.map(z,function(x){x[1]['step'] = x[0]; return x[1]; });    
}

function report() {
    var r = JSON.parse(readFile("report.json"));
    return addSteps(r);
}

function exec(cmd){
    var shell = WScript.CreateObject("WScript.Shell");
    var results = shell.Run(cmd,0,true);
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
    exec(cmd);
}

function extractionSuccessful(f) {
    var fs = new ActiveXObject("Scripting.FileSystemObject");
    return fs.FileExists(f);
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

function newWordDoc(){
   return WORD.getInstance().Documents.Add();
}

function paste(doc) {
    doc.Range().Paste();
}

function save(doc,path){
    doc.SaveAs2(path);
}

function close(doc){
    var saveChanges = false;
    doc.Close(saveChanges);
    //if no more docs, quit Word
    if(! WORD.getInstance().Documents.Count > 0){
        WORD.Quit();
    }
}

function fixPath(s){
    var re = /\|/gi;
    return s.replace(re,'\\');
}

function getCurrentDirectory() {
    var s = WScript.ScriptFullName.split('\\');
    s.pop();
    return s.join('\\');
}

function getCurrentScriptName(){
    var s = WScript.ScriptFullName.split('\\');
    return s.pop();
}

function initFolder(f){
    var fs = new ActiveXObject("Scripting.FileSystemObject");
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
    //exec("cmd /c rmdir /s /q " + f);
    var fs = new ActiveXObject("Scripting.FileSystemObject");
    if(!fs.FolderExists(f)){
        return;
    }
    fs.DeleteFolder(f,true);
}

function pad(num, size) {
    var s = "00" + num;
	s = s.split('').reverse().join('');
	s = s.substr(0,size)
    return s.split('').reverse().join('');
}

function getNewFilename(screenshots) {
    var fs = new ActiveXObject("Scripting.FileSystemObject");
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
    var d = pad(c.toString(),3) + ".png";
    return fs.BuildPath(folder.Path, d);
}

function initConfig(){
    var s = readFile("settings.json");
    var config = JSON.parse(s);
    config.fullPath = fixPath(config.fullPath);
    config.projectPath = config.fullPath + "\\" + config.projectName;
    config.tmp = config.projectPath + "\\tmp"
    config.tmpDoc = config.tmp + "\\tmp.docx";
    config.tmpPicName = "image1.png";
    config.tmpPicPath = config.tmp + "\\" + config.tmpPicName;
    config.pixPath = config.projectPath + "\\screenshots";
    config.report = config.projectPath + "\\report.json";
    config.cwd = WScript.CreateObject ("WScript.Shell").CurrentDirectory;
    config.bin = config.cwd + "\\bin";
    return config;
}

function move(o) {
    fs = new ActiveXObject("Scripting.FileSystemObject");
    fs.MoveFile(o.file, o.to);
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

function readReport(r){
    var fs = new ActiveXObject("Scripting.FileSystemObject");
    if(!fs.FileExists(r)){
        return [];
    }
    var s = readFile(r);
    return JSON.parse(s);
}

function writeReport(report_obj, report_path){
    var fs = new ActiveXObject("Scripting.FileSystemObject");
    var f = fs.OpenTextFile(report_path, 2, true);
    f.Write(JSON.stringify(report_obj));
    f.Close();
}

function hasAnImage(doc){
    if(doc.InlineShapes.Count > 0){
        return true;
    }
}

function imageDimensions(doc){
    return {"width":  doc.InlineShapes.Item(1).Width / 72,
            "height": doc.InlineShapes.Item(1).Height / 72}
}

function imageTooWide(doc){
    var img = imageDimensions(doc);
    if(img.width >= 5.9){
        return true;
    }
    else {
        return false;
    }
}

function imageTooTall(doc){
    var img = imageDimensions(doc);
    if(img.height >= 6){
        return true;
    }
    else {
        return false;
    }
}

function clip2png(){
    var config = initConfig();
    var doc = newWordDoc();
    paste(doc);
    if(hasAnImage(doc)){
        if(!imageTooWide(doc)){
            if(!imageTooTall(doc)){
                initFolder(config.tmp);
                save(doc,config.tmpDoc);
                close(doc);
                extract({'file': config.tmpPicName,
                         'from': config.tmpDoc,
                         'to'  : config.tmp});
                
                var newFilePath = "";
                if(extractionSuccessful(config.tmpPicPath)){
                    newFilePath = getNewFilename(config.pixPath);
                }
                else{
                    VB.MsgBox("There was no picture on your clipboard!",16,"Could not save your screenshot.");
                    delFolder(config.tmp);
                    WScript.Quit();
                }
                move({"file": config.tmpPicPath, "to": newFilePath});
                delFolder(config.tmp);
                var report = readReport(config.report);
                var newEntry = {};
                newEntry.screenshot = newFilePath.split("\\").pop();
                newEntry.caption = getCaption();
                newEntry.minutes = getTimeEstimate();
                report.push(newEntry);
                writeReport(report,config.report);
            }
            else {
                var img = imageDimensions(doc);
                VB.MsgBox("The picture is too tall. (" + img.height + " inches)",16,"Could not save your screenshot.");
                close(doc);
                WScript.Quit();
            }
        }
        else{
            var img = imageDimensions(doc);
            VB.MsgBox("The picture is too wide. (" + img.width + " inches)",16,"Could not save your screenshot.");
            close(doc);
            WScript.Quit();
        }
    }
    else {
        VB.MsgBox("There was no picture on your clipboard!",16,"Could not save your screenshot.");
        close(doc);
        WScript.Quit();
    }
}

//function main() {
//    print(_.map(report(),formatStep).join("\n\n"));
//}

function main() {
    clip2png();
}

main();
