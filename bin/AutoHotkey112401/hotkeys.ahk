#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Recommended for catching common errors.
#NoTrayIcon
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

^+a::Run, "C:\Windows\Sysnative\SnippingTool.exe"
^+s::Run, mspaint
^+x::Run, cscript.exe //nologo clip2png.js, ..\.. , Hide
^+p::Run, cscript.exe //nologo png2ppt.js, ..\.. , Hide
^+=::Run, cscript.exe //nologo switchProjects.js, ..\.. , Hide
^+1::
FormatTime, mytime, , yyyy-MM-dd
Send %mytime%
Return
^+2::
FormatTime, mytime, , yyyy-MM-dd_HHmm
Send %mytime%
Return
^+z::
WinGetTitle, Title, A
WinKill, %Title%
Return
