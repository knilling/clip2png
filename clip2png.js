/*
 *Copyright (c) 2016 Christopher Crawford
 *
 *Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
 *
 *The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
 *
 *THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 *
 */

var powerpnt = new ActiveXObject("PowerPoint.Application");
var ppt = powerpnt.Presentations.Add(0);
var slides = ppt.Slides;
var slide = slides.Add(1, 12);
var my_slide = slide.Shapes;
try {
	var shapes_collection = my_slide.Paste();
	var screen_shot = shapes_collection(1);
	ppt.PageSetup.SlideHeight = screen_shot.Height;
	ppt.PageSetup.SlideWidth = screen_shot.Width;
	screen_shot.Left = 0;
	screen_shot.Top = 0;
	screen_shot.ScaleHeight(1,-1);
	screen_shot.ScaleWidth(1,-1);
	var name = getFileName();
	ppt.Slides(1).Export(name,"PNG");
}
catch(e) {
}
powerpnt.Quit();


WScript.Quit();

function getFileName() {
	var script_path = WScript.ScriptFullName;
	var fs = new ActiveXObject("Scripting.FileSystemObject");
	var script = fs.GetFile(script_path);
	var folder = script.ParentFolder;
	var files = new Enumerator(folder.Files);
	var i = 1;
	var file_array = [];
	while (!files.atEnd()) {
		var file = files.item();
		if (file.Name.match(/\d{3}\.png$/gi)) {
			//WScript.Echo(file.Path);
			file_array.push(file.Name.split(".")[0]);
		}
		files.moveNext();
	}
	if(file_array.length === 0) {
		file_array.push("0");
	}
	var new_file_array = file_array.sort(function(a,b){return a-b});
	var name = new_file_array.pop();
	var new_name = parseInt(name,10) + 1;
	new_name = new_name + "";
	var new_new_name = pad(new_name,3) + ".png";
	var path = fs.BuildPath(folder.Path, new_new_name)
	return path;
}

function pad(num, size) {
    var s = "00" + num;
	s = s.split('').reverse().join('');
	s = s.substr(0,size)
    return s.split('').reverse().join('');
}