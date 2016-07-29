/*
 * msgbox.js
 * Copyright (c) 2016 Christopher Crawford
 *
 * Using ideas from:
 * [+] http://with-love-from-siberia.blogspot.com/2009/12/msgbox-inputbox-in-jscript.html
 * [+] http://eloquentjavascript.net/10_modules.html
 *
 * Have to min the js, in order for it to load as a library.
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
(function(exports){var vb={};vb.Function=function(func){return function(){return vb.Function.eval.call(this,func,arguments);};};vb.Function.eval=function(func){var args=Array.prototype.slice.call(arguments[1]);for(var i=0;i<args.length;i++){if(typeof args[i]!='string'){continue;} args[i]='"'+args[i].replace(/"/g,'" + Chr(34) + "')+'"';} var vbe;vbe=new ActiveXObject('ScriptControl');vbe.Language='VBScript';return vbe.eval(func+'('+args.join(', ')+')');};exports.InputBox=vb.Function('InputBox');exports.MsgBox=vb.Function('MsgBox');})(this.VB={});