#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#2::
FullPath := "D:\Book1.xlsx"	; please adjust full path to your Workbook...
Xl := "", XlSheet := ""	; release references
Xl := ComObjGet(FullPath)		; get reference to WorkBook
Xl.Application.Windows(Xl.Name).Visible := 1	; just do it - too long to explain why...
CellN19 := Xl.WorkSheets(1).Range("N19")
CellO19 := Xl.WorkSheets(1).Range("O19")
CellN20 := Xl.WorkSheets(1).Range("N20")
CellO20 := Xl.WorkSheets(1).Range("O20")
Cell_One := CellN19.Text ;Shipment cal start Cell
Cell_Two := CellO19.Text ;Shipment cal start Cell
Cell_Three := CellN20.Text ;Open Order cal start Cell
Cell_Four := CellO20.Text ;Open Order cal end Cell
Sleep 500
Send, %Cell_One%
;MSGBox,%Cell_One%
;Xl.Close(1)	; save changes and close Workbook
WinActivate,Book1 - Excel
Xl.WorkSheets(1).Range("A1").Activate()
Xl := "", XlSheet := ""	; release references
ExitApp
Esc::ExitApp

#4::
WinActivate,Internet Explorer_Server1
ControlSend, Internet Explorer_Server1, {Tab}, ahk_exe Iqwin32.exe