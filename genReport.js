/* 
 * genReport.js
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

function app(){
    return WORD.getInstance();
}

function DOC(){
    this.doc = app().Documents.Add();

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
}

function right(n){
    for(i=0; i<n; i++){
        app().Selection.MoveRight();
    }
}

function left(n){
    for(i=0; i<n; i++){
        app().Selection.MoveLeft();
    }
}

function up(n){
    for(i=0; i<n; i++){
        app().Selection.MoveUp();
    }
}

function down(n){
    for(i=0; i<n; i++){
        app().Selection.MoveDown();
    }
}

function newParagraph(){
    app().Selection.Paragraphs.Add();
}

function nextLine(){
    newParagraph();
    down(1);
}

function newPage(){
    var my_wdPageBreak = 7;
    app().Selection.InsertBreak(my_wdPageBreak);
}

function text(s){
    app().Selection.text = s;
}

function stylized(txt, style) {
    text(txt);
    app().Selection.style = app().ActiveDocument.Styles(style)
    nextLine();
}

function bulletedList(a){
    app().Selection.Range.ListFormat.ApplyBulletDefault();
    for(var i=0; i < a.length; i++){
        text(a[i]);
        nextLine();
    }
    app().Selection.Range.ListFormat.RemoveNumbers();
}

function newTable(rows, cols, border){
    var t = app().ActiveDocument.Tables.Add(app().Selection.Range, rows, cols);
    t.TopPadding = 0;
    t.RightPadding = 0;
    t.LeftPadding = 0;
    t.BottomPadding = 0;
    t.Select();
    app().Selection.style = app().ActiveDocument.Styles("No Spacing");
    left(1);
    if(border){
        t.Borders.Enable = true;
    }
    else {
        t.Borders.Enable = false;
    }
    return t;
}

function h1(s){
    stylized(s, "Heading 1");
}

function h2(s){
    stylized(s, "Heading 2");
}

function pic(path, border){
    var p = app().ActiveDocument.Shapes.AddPicture(path, false, true);
    //wdWrapInline = 7
    var my_wdWrapInline = 7
    p.WrapFormat.Type = my_wdWrapInline;
    if (border){
        p.Line.Weight = 1
        //RGB(0,0,0) = 0
        p.Line.ForeColor.RGB = 0
    }
    return p;
}

function execSumHeader(logo){
    //Add heading table
    var t = newTable(2, 2, false);
    var c = t.Cell(2, 1);
    t.Cell(1, 1).Merge(c);
    pic(logo, false);
    
    // Populate heading table
    right(2);
    text("<<Report Title>>");
    app().Selection.style = app().ActiveDocument.Styles("Title");
    right(1);
    newParagraph();
    down(1);
    text("UBNETDEF Field Guide");
    app().Selection.style = app().ActiveDocument.Styles("Subtitle");
    down(1);
    text("<<Author Name>>");
    right(1);
    newParagraph();
    down(1);
    text("<<YYYY-MM-DD>>");
    down(1);
}

function execSumContent(){
    // Populate Excutive Summary Page
    h1("Executive Summary");
    h2("Objective");
    text("After completing this guide, the reader will be able to <<finish this statement>>.");
    nextLine();
    
    h2("Requirements");
    text("In order to complete this guide, the reader will need the following:");
    nextLine();
    
    bulletedList(["<<Stuff>>", "<<Things>>", "<<More Things>>"]);
    
    h2("Time Estimate");
    text("The reader can expect the following procedure to take <<X>> minutes.");
    
    nextLine();
    newPage();
}

function executiveSummary(logo){
    execSumHeader(logo);
    execSumContent();
}

function table_of_contents(){
    // Add Table of Contents
    h1("Table of Contents");
    app().ActiveDocument.TablesOfContents.Add(app().Selection.Range);
    down(1);
    newPage();
}

function addRow(table){
    table.rows.Add();
    down(1);
}

function addTableHeaders(t){
    text("Step");
    right(2);
    text("Time (minutes)");
}

function addTableData(t, steps){
    t.Cell(1, 1).Select();
    for(var i = 0; i < steps.length; i++){
        addRow(t);
        text(steps[i]);
    }
}

function addTotalRow(t){
    addRow(t);
    text("Tota Time");
}

function setColumnWidths(t){
    // wdAdjustNone = 0
    var my_wdAdjustNone = 0;
    t.Columns(1).SetWidth(404, my_wdAdjustNone);
    t.Columns(2).SetWidth(72, my_wdAdjustNone);
}

function centerTable(t){
    // wdAlignRowCenter = 1
    my_wdAlignRowCenter = 1
    t.rows.Alignment = my_wdAlignRowCenter
}

function formatHeaderRow(t){
    var rng = t.rows(1).Range;
    rng.Font.Bold = true;
    // wdAlignParagraphCenter = 1
    my_wdAlignParagraphCenter = 1
    rng.ParagraphFormat.Alignment = my_wdAlignParagraphCenter;
}

function setTableFonts(t){
    rng = t.rows(2).Range;
    rng.End = t.rows(t.rows.Count - 1).Range.End
    rng.Font.Name = "Courier New"
    t.Cell(t.rows.Count, 2).Range.Font.Name = "Courier New"
    t.Range.Font.Size = 8
}

function setAlignmentForTimeData(t){
    var rng = t.Cell(2, 2).Range;
    rng.End = t.Cell(t.rows.Count, 2).Range.End;
    rng.Select();
    // wdAlignParagraphRight = 2
    var my_wdAlignParagraphRight = 2
    app().Selection.ParagraphFormat.Alignment = my_wdAlignParagraphRight;
}

function italicizeSteps(t){
    var rng = t.Cell(2, 1).Range;
    rng.End = t.Cell(t.rows.Count - 1, 1).Range.End;
    rng.Select();
    app().Selection.Font.Italic = true
}

function setTableBorders(t){
    // Format table borders
    // wdLineWidth075pt = 6
    my_wdLineWidth075pt = 6
    t.Borders.InsideLineWidth = my_wdLineWidth075pt
    var rng = t.Cell(2, 1).Range;
    rng.End = t.Cell(t.rows.Count - 1, 2).Range.End
    rng.Select();
    // wdLineWidth150pt = 12
    my_wdLineWidth150pt = 12
    app().Selection.Borders.OutsideLineWidth = my_wdLineWidth150pt;
        
    var rng = t.Cell(1, 1).Range;
    rng.End = t.Cell(t.rows.Count, 1).Range.End
    rng.Select();
    // wdLineStyleSingle = 1
    my_wdLineStyleSingle = 1;
    app().Selection.Borders.OutsideLineStyle = my_wdLineStyleSingle;
    app().Selection.Borders.OutsideLineWidth = my_wdLineWidth150pt;
    
    t.Borders.OutsideLineStyle = my_wdLineStyleSingle;
    // wdLineWidth225pt = 18
    my_wdLineWidth225pt = 18;
    t.Borders.OutsideLineWidth = my_wdLineWidth225pt;
}

function setTablePadding(t){
    t.LeftPadding = 5;
    t.RightPadding = 15;
}

function boldTotalsRow(t){
    var rng = t.rows(t.rows.Count).Range
    rng.Select();
    app().Selection.Font.Bold = true;
}

function removeItalicsFromTimeData(t){
    // Make sure time data is not italicized
    t.Columns(2).Select();
    app().Selection.Font.Italic = false;
}

function shadeBandedRows(t){
    // wdColorGray20 = 13421772
    my_wdColorGray20 = 13421772;
    for(var i = 1; i < t.rows.Count; i++){
        if(i % 2 === 0){
            t.rows(i).Shading.BackgroundPatternColor = my_wdColorGray20;
        }
    }
}

function time_estimate(steps){
    // Add Time Estimate Section
    h1("Time Estimate Table");
    nextLine();
    
    // Add Time Estimate Table
    var t = newTable(1, 2, true);
        
    addTableHeaders(t);
    addTableData(t, steps);
    addTotalRow(t);
    centerTable(t);
    formatHeaderRow(t);
    setTableFonts(t);
    setAlignmentForTimeData(t);
    italicizeSteps(t);
    setTableBorders(t);
    setTablePadding(t);
    boldTotalsRow(t);
    removeItalicsFromTimeData(t);
    setColumnWidths(t);
    shadeBandedRows(t);
    down(1);
    newPage();
}

function procedure_step(i){
    var t = newTable(6, 1, false);
    h2(i);
    app().Selection.TypeBackspace();
    down(2);
    text("Estimated Time Required: " + "<<X>>" + " minutes");
    down(2);
    var p = pic("C:\\Users\\Chris\\Desktop\\ubnetdef.png", true);
    // wdAlignParagraphCenter = 1
    my_wdAlignParagraphCenter = 1;
    app().Selection.ParagraphFormat.Alignment = my_wdAlignParagraphCenter
    down(2);
    newPage();
}

function procedure(steps){
    // Add Procedure Section
    stylized("Procedure", "Heading 1");
    for(var i = 0; i < steps.length; i++){
        procedure_step(steps[i]);
    }
    app().Selection.TypeBackspace();
    app().Selection.TypeBackspace();
    app().Selection.TypeBackspace();
}

function update_toc(){
    app().ActiveDocument.TablesOfContents(1).Update();
}

function add_page_numbers(){
    // wdHeaderFooterPrimary = 1
    // wdAlignPageNumberCenter = 1
    my_wdHeaderFooterPrimary = 1;
    my_wdAlignPageNumberCenter = 1;
    app().ActiveDocument.Sections(1).Footers(my_wdHeaderFooterPrimary).PageNumbers.Add(my_wdAlignPageNumberCenter, false);
}

function genReport(){
    add_page_numbers();
    logo = "C:\\Users\\Chris\\Desktop\\ubnetdef.png"
    executiveSummary(logo);
    table_of_contents();
    var steps = ["One", "Two", "Three"];
    time_estimate(steps);
    procedure(steps);
    update_toc();
}
