;^F12::
SetDefaultMouseSpeed, 0
WinMinimizeAll
;Traytip , Report Updater, Now Running Reports. Please Do not Move Or Use Mouse or Keyboard. Win+2 will kill process, 95
Xl := "", XlSheet := "" ;release references
FullPath := "\\Hlserver\Company Data\Wirecutting\Master Scheduler (V2).xlsx"	; please adjust full path to your Workbook...
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
;MSGBox,%Cell_One%
Sleep 2000
Xl.Close(1) ; save changes and close Workbook
;WinActivate,Master Scheduler (V2) - Excel
;Xl.WorkSheets(1).Range("A1").Activate()
Xl := "", XlSheet := ""	; release references
;Main startup Starts IQMS and Logs in to User;

MSGBox 4, IQMS Report Updater, Would you like to Update Reports. Reports will Run in this Order, Shipment items, Commponments, Open Orders,20

IFMSGBox ,Timeout
{
Xl := "", XlSheet := "" ;release references
ExitApp
}
IFMSGBOX ,No
{
Xl := "", XlSheet := "" ;release references
MSGBOX , You Have Killed The App
ExitApp
}
IFMSGBOX ,Yes
{
Ifexist C:\Program Files\IQMS\IQWin32\Iqwin32.exe
Run C:\Program Files\IQMS\IQWin32\Iqwin32.exe
Ifexist C:\Program Files (x86)\IQMS\IQWin32\Iqwin32.exe
Run C:\Program Files (x86)\IQMS\IQWin32\Iqwin32.exe
}
;Run C:\Program Files (x86)\IQMS\IQWin32\Iqwin32.exe
;Run Notepad
WinWait,ahk_class TLogin , ,10
if ErrorLevel
{
    MSGBox, NO Login window found.
ExitApp
}
else
;Sleep 10000
if WinActive("ahk_class TLogin")
;IfWinActive, Login Please , ,
{
ControlSend, TEdit1, wire, ahk_class TLogin
ControlSend, TEdit1, {enter}, ahk_class TLogin
;Send, {enter}
Sleep 5000
}
;Shipping Reports; ; Runs the Shipping Reports;

;MouseMove, 190, 63, 0, ;Remaped for Sales and Distribution tab
;MouseClick, , , , , ;Clicks left button once.
;MouseMove, 177, 98, 0, ;Remaped for Shipping and Pack Slips
;MouseClick, , , , , ;Clicks left button once.
;ControlSend, TPageControl1, {Alt}, ahk_exe Iqwin32.exe
;ControlSend, TPageControl1, F, ahk_exe Iqwin32.exe
if WinActive ahk_class TIQLauncher
sleep 1000
Send, {Alt}
Send,F
Send,D
Send,A
Sleep 1000
Send, {enter}
Sleep 500
Send, {enter}
Sleep 1000
Send, {Alt}
Send,R
Send,P
Sleep 2000
;WinWaitActive,Registered Reports (ID: FrmPsMaint), 10
WinWait,Registered Reports (ID: FrmPsMaint)
ControlSend, TwwIncrementalSearch1, Shipments by Item Number, ahk_exe Iqwin32.exe
;Send,Shipments by Item Number
Sleep 2000
;ControlSend, TwwIncrementalSearch1, {Tab}, ahk_exe Iqwin32.exe
;Send, {Tab}
Sleep 500
ControlSend, TwwDBGrid1, {End}, ahk_exe Iqwin32.exe
ControlSend, TwwDBGrid1, {Up 8}, ahk_exe Iqwin32.exe
;Send, {Down 8}
Sleep 2000
ControlSend, TwwDBGrid1, {Tab}, ahk_exe Iqwin32.exe
;Send {Tab}
ControlSend, TPageControl1, {End}, ahk_exe Iqwin32.exe
;Send {End}
SLEEP 3000
; Retrieve the ID/HWND of the active window
id := WinExist("A")
WinMaximize, A
Sleep 500
MouseMove, 270, 728
;MouseMove, 343, 728

;MouseMove, 270, 666
Click, 3
Send {Tab 4}
Send {Enter}
Click, 3
MouseMove, 68, 122, 0, ;Click Calander cell
;Click,3
sleep 3000
Id := WinExist("A")
;Id := WinWaitActive("A")
Send, {Tab}
ControlSend, A, {Tab 2}
ControlSend, Internet Explorer_Server1, %Cell_One%, ahk_exe Iqwin32.exe
ControlSend, Internet Explorer_Server1, {Tab 4}, ahk_exe Iqwin32.exe
;MouseMove, 287, 124, 0, ;Click Calander cell
;MouseClick, , , , , ;Clicks left button once.;
sleep 500
MouseMove, 405, 122, 0, ;Click Calander cell
MouseClick, 405, 122, 0,, , , , ;Clicks left button once.;
ControlSend, Internet Explorer_Server1, %Cell_Two%, ahk_exe Iqwin32.exe
ControlSend, Internet Explorer_Server1, {Tab 4}, ahk_exe Iqwin32.exe
Sleep 500
ControlSend, Internet Explorer_Server1, {Enter}, ahk_exe Iqwin32.exe
;MouseMove, 596, 895, 0, ;Click Calander cell
;Send, %Cell_Two%
Msgbox, 1, , Component Reports Starting Now,10

IfMsgBox, Cancel
{
ExitApp
}
Else

;IfMsgBox, Timeout;

{
WinActivate,ahk_exe Iqwin32.exe
Send {Esc}
Send {Alt}
Send F
Send C
Sleep 2000
;Component Reports;
if WinActive("ahk_exe Iqwin32.exe")
Sleep 500
Send, {Alt}
sleep 500
Send,F
Send,M
Send,I
Sleep 1000
Send, {enter}
Sleep 500
Send, {enter}
Sleep 1000
Send, {Alt}
Send,R
Send,P
;Sleep 5000

WinWait, ahk_class TFrmRepDef
ControlSend, TwwIncrementalSearch1, Inventory Listing by Class, ahk_exe Iqwin32.exe
;Send,Shipments by Item Number
Sleep 2000
;ControlSend, TwwIncrementalSearch1, {Tab}, ahk_exe Iqwin32.exe
;Send, {Tab}
Sleep 500
;ControlSend, TwwDBGrid1, {End}, ahk_exe Iqwin32.exe
ControlSend, TwwDBGrid1, {Down 3}, ahk_exe Iqwin32.exe
;Send, {Down 8}
Sleep 2000
ControlSend, TwwDBGrid1, {Tab}, ahk_exe Iqwin32.exe ;Selection Criteria;
;Send {Tab}
ControlSend, TPageControl1, {End}, ahk_exe Iqwin32.exe ;Destination;
;Send {End}
SLEEP 3000
; Retrieve the ID/HWND of the active window
;id := WinExist("A")
;WinMaximize, A
Sleep 500
MouseMove, 270, 728
;MouseMove, 343, 728
;MouseMove, 270, 666
Click, 3
;Send {Tab 4}
;Send {Enter}
;Click, 3
ControlSend, Internet Explorer_Server1, {Altdown}P{Altup}, ahk_class WindowsForms10.Window.8.app.0.378734a_r61_ad1
;MouseMove, 68, 122, 0, ;Click Calander cell
;Click,3
sleep 3000
id := WinExist("A")
;Send, {Tab}
;ControlSend, A, {Tab 2}
Sleep 1000
;Send {Tab}
Sleep 500
ControlSend, Internet Explorer_Server1, {Tab}, ahk_class WindowsForms10.Window.8.app.0.378734a_r61_ad1
Sleep 500
ControlSend, Internet Explorer_Server1, {DownArrow}, ahk_class WindowsForms10.Window.8.app.0.378734a_r61_ad1
;Send {DownArrow}
;ControlSend, Internet Explorer_Server1, {Shiftdown}A{Shiftup} , ahk_exe Iqwin32.exe
Send {Tab}
Sleep 500
Send {Enter}
Sleep 4000
;Component Reports;
}

Msgbox, 1, , Open Orders Reports Starting Now,10

IfMsgBox, Cancel
{
ExitApp
}
Else

Sleep 3000
WinActivate,ahk_class TFrmRepDef
Send {Esc}
Send {Alt}
Send F
Send X

Sleep 2000

;Open Orders Reports;

if WinActive("ahk_exe Iqwin32.exe")
Sleep 1000
Send, {Alt}
sleep 500
Send,F
Send,D
Send,O
Sleep 1000
Send, {enter}
Sleep 500
Send, {enter}
Sleep 1000
Send, {Alt}
Send,R
Send,P
Sleep 5000
WinWait,Registered Reports (ID: FrmMainOrder)
ControlSend, TwwIncrementalSearch1, Open Order Report, ahk_exe Iqwin32.exe
;Send,Shipments by Item Number
Sleep 2000
;ControlSend, TwwIncrementalSearch1, {Tab}, ahk_exe Iqwin32.exe
;Send, {Tab}
Sleep 500
;ControlSend, TwwDBGrid1, {End}, ahk_exe Iqwin32.exe
;ControlSend, TwwDBGrid1, {Down 3}, ahk_exe Iqwin32.exe
;Send, {Down 8}
Sleep 2000
ControlSend, TwwDBGrid1, {Tab}, ahk_exe Iqwin32.exe ;Selection Criteria;
;Send {Tab}
ControlSend, TPageControl1, {End}, ahk_exe Iqwin32.exe ;Destination;
;Send {End}
SLEEP 3000
; Retrieve the ID/HWND of the active window
;id := WinExist("A")
;WinMaximize, A
Sleep 500
MouseMove, 270, 728
;MouseMove, 343, 728
;MouseMove, 270, 666
Click, 3
Send {Tab 4}
Send {Enter}
Click, 3
MouseMove, 68, 122, 0, ;Click Calander cell
;Click,3
sleep 3000
id := WinExist("A")
Send, {Tab}
ControlSend, A, {Tab 2}
ControlSend, Internet Explorer_Server1, %Cell_Three%, ahk_exe Iqwin32.exe
ControlSend, Internet Explorer_Server1, {Tab 4}, ahk_exe Iqwin32.exe
;MouseMove, 287, 124, 0, ;Click Calander cell
;MouseClick, , , , , ;Clicks left button once.;
sleep 500
MouseMove, 405, 122, 0, ;Click Calander cell
MouseClick, 405, 122, 0,, , , , ;Clicks left button once.;
ControlSend, Internet Explorer_Server1, %Cell_Four%, ahk_exe Iqwin32.exe
ControlSend, Internet Explorer_Server1, {Tab 4}, ahk_exe Iqwin32.exe
Sleep 500
ControlSend, Internet Explorer_Server1, {Enter}, ahk_exe Iqwin32.exe

MSGBox, Reports are complete
WinActivate,ahk_class TFrmRepDef
Send {Esc}
Sleep 500
Send {Alt}
Send F
Send X
Xl.WorkSheets(1).Range("A1").Activate()
Xl := "", XlSheet := ""	;release references
ExitApp

#2::
{

MSGBox, You have chosen to kill the app
ExitApp
}


;Open Orders Reports;

ExitApp

