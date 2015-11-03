;
; AutoHotkey Version: 1.x
; Language:       English
; Platform:       Win9x/NT
; Author:         MPagel       
; Website:        http://www.github.com/MPagel
;
; Script Function:

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

SysGet, _ , MonitorWorkArea ; right and bottom seem to be 1s based, whereas left and top are 0s based. Right - Left is giving 1px more than the working area of the screen

titleBarSize := 26
fontSize := 12
vertSpacer := 16
box1VStart := titleBarSize + vertSpacer - fontSize // 2 ; or 36
lineHeight := 21
windowBorder := 3
waw := _Right - _Left ; set waw smaller or get it from the initial window size if you don't want to take up the whole width of the screen
wah := _Bottom - _Top - titleBarSize ; minus windowBorder?
horizSpacer := 10
buttonHeight := 26
minButtonPadding := 4 ; default padding around top and bottom of each button. This is a minimum (minus object border).

importCancel := buttonHeight + minButtonPadding * 2
box2Height := 6 * lineHeight ; approximate size of "When importing VRL data set that already exists" dialogue
box1Height := wah - box2Height - box1VStart - importCancel - windowBorder
filelistVStart := titleBarSize + vertSpacer * 2 ; or 60 ; or box1VStart + 24
filelistHeightMax := box1Height + box1Vstart - filelistVStart - buttonHeight - minButtonPadding * 2
filelistLines := filelistHeightMax // lineHeight ; // 2 if you want to force pairing of .vrl with matching _edited.vrl
filelistHeight := filelistLines * lineHeight + 2 ; * 2 on left side if paired
buttonPadding2 := box1Height + box1VStart - filelistVStart - filelistHeight - buttonHeight
buttonPadding := buttonPadding2 // 2
rightAlnDetails := waw - 97 ; could look up right boundary of TAdvStringGrid1 (Xg+Wg=Rg) after its control move as well as size (Wb) of TButton1. then rightAlnDetails := Rg-Wb
box3VStart := box1VStart + box1Height + vertSpacer + box2Height + 1
box3VEnd := wah + titleBarSize - windowBorder
box3Pad2 := box3VEnd - box3VStart - buttonHeight
box3Pad := box3Pad2 // 2

; MsgBox, %filelistHeightMax% %filelistHeight% %buttonPadding2% %box1Height% %box1VStart% %filelistVStart%
DetectHiddenWindows, On
SetWinDelay, 10
WinActivate, Edit and Import VRL Files
WinMove, A, , _Left, _Top, waw, wah + titleBarSize
SetControlDelay, 0
SetWinDelay, 1
ControlMove, TGroupBox1, windowBorder + horizSpacer, box1VStart, , box1Height, A ; box1 is 975 in top-bottom 
ControlMove, TAdvStringGrid1, horizSpacer * 2 + windowBorder, filelistVStart, , filelistHeight, A
ControlMove, TButton4, horizSpacer * 2 + windowBorder, filelistHeight + filelistVStart + buttonPadding , , , A ; Details
ControlMove, TButton3, , filelistHeight + filelistVStart + buttonPadding , , , A ; Auto Correct
ControlMove, TButton2, , filelistHeight + filelistVStart + buttonPadding , , , A ; Reset
ControlMove, TButton1, rightAlnDetails, filelistHeight + filelistVStart + buttonPadding, , , A ; Help
ControlMove, TGroupBox2, windowBorder + horizSpacer, box1VStart + box1Height + vertSpacer, , box2Height, A ; when importing a VRL data set that already exists in the database
ControlMove, TPanel1, , box1VStart + box1Height + vertSpacer + box2Height + 1 + box3Pad,  , importCancel , A 
; import and cancel buttons. seems to want to stick bottom of button about 17 px above the bottom of the screen. See below Exit for more details.
ControlMove, TPanel2, , , , lineHeight * 4 + lineHeight // 2, A ; 4.5 * LH... may be 5 * fontSize instead
WinSet, Redraw
SetControlDelay, 0 ; default is 20, but using that seems to cause a lag in button response?
SetWinDelay, 0 ; default is 100

ExitApp

; Previous version code had the following
; ControlMove, TPanel1, , box1VStart + box1Height + vertSpacer + box2Height + minButtonPadding + 1,  , importCancel , A 
; 36 + 975 + 16 + 126 + 4 + 1 = 1158 of 1199 max

; ControlGetPos, x1, y1, w1, h1, TPanel1, A
; ControlGetPos, x2, y2, w2, h2, TPanel2, A
; MsgBox, Width: %waw%`tHeight: %wah%`tBox1 Height: %box1Height%`tFile List Height: %filelistHeight%`nPanel 1: %x1% %y1% %w1% %h1%`nPanel 2: %x2% %y2% %w2% %h2%`n

; WinSet, Style, +0x70000, A ; WS_MAXIMIZEBOX 0x10000 + WS_MINIMIZEBOX 0x20000 + WS_SIZEBOX 0x40000 ; would have to add functions for these buttons, though as behavior is not built-in.

; =================================================================================
; Function: AutoXYWH
;   Move and resize control automatically when GUI resizes.
; Parameters:
;   DimSize - Can be one or more of x/y/w/h  optional followed by a fraction
;             add a '*' to DimSize to 'MoveDraw' the controls rather then just 'Move', this is recommended for Groupboxes
;   cList   - variadic list of ControlIDs
;             ControlID can be a control HWND, associated variable name, ClassNN or displayed text.
;             The later (displayed text) is possible but not recommend since not very reliable 
; Examples:
;   AutoXYWH("xy", "Btn1", "Btn2")
;   AutoXYWH("w0.5 h 0.75", hEdit, "displayed text", "vLabel", "Button1")
;   AutoXYWH("*w0.5 h 0.75", hGroupbox1, "GrbChoices")
; ---------------------------------------------------------------------------------
; Version: 2015-5-29 / Added 'reset' option (by tmplinshi)
;          2014-7-03 / toralf
;          2014-1-2  / tmplinshi
; requires AHK version : 1.1.13.01+
; =================================================================================
;AutoXYWH(DimSize, cList*){       ; http://ahkscript.org/boards/viewtopic.php?t=1079
;  static cInfo := {}
; 
;  If (DimSize = "reset")
;    Return cInfo := {}
; 
;  For i, ctrl in cList {
;    ctrlID := A_Gui ":" ctrl
;    If ( cInfo[ctrlID].x = "" ){
;        GuiControlGet, i, %A_Gui%:Pos, %ctrl%
;        MMD := InStr(DimSize, "*") ? "MoveDraw" : "Move"
;        fx := fy := fw := fh := 0
;        For i, dim in (a := StrSplit(RegExReplace(DimSize, "i)[^xywh]")))
;            If !RegExMatch(DimSize, "i)" dim "\s*\K[\d.-]+", f%dim%)
;              f%dim% := 1
;        cInfo[ctrlID] := { x:ix, fx:fx, y:iy, fy:fy, w:iw, fw:fw, h:ih, fh:fh, gw:A_GuiWidth, gh:A_GuiHeight, a:a , m:MMD}
;    }Else If ( cInfo[ctrlID].a.1) {
;        dgx := dgw := A_GuiWidth  - cInfo[ctrlID].gw  , dgy := dgh := A_GuiHeight - cInfo[ctrlID].gh
;        For i, dim in cInfo[ctrlID]["a"]
;            Options .= dim (dg%dim% * cInfo[ctrlID]["f" dim] + cInfo[ctrlID][dim]) A_Space
;        GuiControl, % A_Gui ":" cInfo[ctrlID].m , % ctrl, % Options
;} } }
