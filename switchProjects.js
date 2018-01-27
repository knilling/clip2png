/* 
 * switchProjects.js
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

function writeToFile(file,data) {
    var forReading = 1;
    var forWriting = 2;
    var forAppending = 3;
    var create = true;
    var f = FS().OpenTextFile(file, forWriting, create);
    f.Write(data);
    f.Close();
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

function initConfig(){
    var s = readFile("settings.json");
    var config = JSON.parse(s);

    var fixPath = function(s){
        var re = /\|/gi;
        return s.replace(re,'\\');
    }
    
    return config;
}

function getProjectName(){
    var s = VB.InputBox("What is the name of the project you want?");
    if(s.length===0){
        VB.MsgBox("You cannot leave this empty.",16,"Not so fast...");
        return getProjectName();
    }
    else{
        return s;
    }
}

function writeConfig(settings){
        writeToFile("settings.json",JSON.stringify(settings));
};

function switchProjects(){
    var settings = initConfig();
    var newProjectName = getProjectName();
    settings.projectName = newProjectName;
    writeConfig(settings);
}

function main() {
    switchProjects();
}

main();
