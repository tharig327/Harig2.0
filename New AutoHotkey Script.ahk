#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
Xl := "", XlSheet := "" ;release references
FullPath := "https://d.docs.live.net/26dcd5873a316737/Documents/Book%201.xlsx"	; please adjust full path to your Workbook...
Xl := ComObjGet(FullPath)		; get reference to WorkBook
Xl.Application.Windows(Xl.Name).Visible := 1	; just do it - too long to explain why...
Xl.WorkSheets(1).Calculate
CellN19 := Xl.WorkSheets(1).Range("N19")
CellO19 := Xl.WorkSheets(1).Range("O19")
CellN20 := Xl.WorkSheets(1).Range("N20")
CellO20 := Xl.WorkSheets(1).Range("O20")
Cell_One := CellN19.Text ;Shipment cal start Cell
Cell_Two := CellO19.Text ;Shipment cal start Cell
Cell_Three := CellN20.Text ;Open Order cal start Cell
Cell_Four := CellO20.Text ;Open Order cal end Cell
msgbox,%Cell_Four%