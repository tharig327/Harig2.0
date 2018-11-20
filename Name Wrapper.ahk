#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
::9000743::19916 ;20972743 SRC HOSE
::9000742::19915 ;20972742 SRC HOSE
::9000741::19914 ;20972741 SRC HOSE
::9000740::19913 ;20972740 SRC HOSE

::WIRE CUTTING::WIRE_CUTTING

5::
vSun += 1-A_WDay, Days
FormatTime, vSun1, %vSun%, ShortDate
vFri += 6-A_WDay, Days
FormatTime, vFri1, %vFri%, ShortDate
vSun365 += -365-A_WDay, Days
FormatTime, vSun3651, %vSun365%, ShortDate
vFri120 += 120-A_WDay, Days
FormatTime, vFri1201, %vFri120%, ShortDate
MsgBox, , ,%vSun1%
MsgBox, , ,%vFri1%
MsgBox, , ,%vSun3651%
MsgBox, , ,%vFri1201%

ExitApp
return

