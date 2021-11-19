; Script Name - Hourly Counter - Working hours counter
/**
This script provides a personalised and privacy focused way to pitch in 
your working hours by creating a small widget to appear on your screen
with simple three button interface.


*/
global csv_file := "Working_Hours_" . A_MMM . "_" . A_Year . " .csv"
global timestamp =
global is_running := false ; Checks the current state

; Set gui positioning to bottom right corner
x_gui:= A_ScreenWidth - 350
y_gui := A_ScreenHeight - 150

Gui, Color, White, Black
Gui, Font, s15, Arial
Gui, Add, Button, vstart gstart x10 y10 w30 h30 , % Chr(127939)
Gui, Add, Button, vstop gstop xp+35 yp w30 h30 , % Chr(129486)
Gui, Add, Button, vdesc gdesc xp+35 yp w30 h30 , % Chr(128172)
Gui, Add, Button, vfinish gfinish xp+35 yp w30 h30 ,  % Chr(127937)
Gui, Add, Text, vtimer xp+35 yp w90 h30 +Center, 00:00:00
Gui, Add, Button, xp+100 yp gopen w30 h30 , % chr(128194)
Gui, Add, Button, xp+35 yp gclose w30 h30 , % chr(128683)

Gui, Show, w330 h50 x%x_gui% y%y_gui%, Hourly Counter
Gui, +LastFound +AlwaysOnTop +ToolWindow

GuiControl , Disable, stop
GuiControl , Disable, desc
GuiControl , Disable, finish

Gui, task: Add, Text, x12 y10 w110 h20 , Add Task Description (What you are working on)
Gui, task: Font, s15, Arial
Gui, task: Add, Button, Disabled x262 y9 w60 h30,% chr(128247)
Gui, task: Font, , 
Gui, task: Add, Edit, x12 y49 w310 h120 vdescription,
Gui, task: Add, Button, x122 y179 w100 h30 , Done

; Insert app entry on start
FileAppend , % "`r`nInitiated, " . A_Hour . ":" . A_Min . ":" . A_Sec . ", App Started " . A_DD . "/" . A_MMM . "/" . A_YYYY . ",`r`n", % csv_file

Menu, Tray, NoStandard
Menu, Tray, Icon, Shell32.dll, 24
Menu, Tray, Add, Show, show
Menu, Tray, Add, Exit, exit

return

show:
	WinShow , Hourly Counter
	return
	
close:
GuiClose:
	WinHide , Hourly Counter
	return

exit:
	ex_action("exit")
ExitApp

start:
	GuiControl , Disable, start
	GuiControl , Enable, stop
	GuiControl , Enable, desc
	GuiControl , Enable, finish
	ex_action("start")
	SetTimer , stopwatch, 1
return

stop:
	GuiControl , Enable, start
	GuiControl , Disable, stop
	GuiControl , Disable, desc
	ex_action("stop")
	SetTimer , stopwatch, Off
	return
desc:
	Gui, task: Show, w336 h219, Add Task Description
	return
finish:
	GuiControl, Enable, start
	GuiControl, Disable, stop
	GuiControl , Disable, finish
	ex_action("finish")
	SetTimer , stopwatch, Off
return

taskButtonDone:
	Gui, task: submit, Nohide
	ex_action("update", description)
	WinHide , Add Task Description
return

open:
	Run, % A_ScriptDir
return

stopwatch:
	GuiControl , , timer, % FormatTimeStamp(A_TickCount - timestamp)
return


ex_action(command:="start", desc := ""){
	global timestamp
	global is_running
	
	switch (command){
		case "start":
			timestamp := A_TickCount
			FileAppend , % "Started, " . A_Hour . ":" . A_Min . ":" . A_Sec . ", Session started, `r`n", % csv_file
			WinSetTitle, Hourly Counter | Running...
			is_running := true
		return
		case "stop":
			FileAppend,  % "Stopped|Paused, " . A_Hour . ":" . A_Min . ":" . A_Sec . ", Session stopped/paused, " . FormatTimeStamp(A_TickCount - timestamp ) . " `r`n", % csv_file
			WinSetTitle, Hourly Counter | Stopped/Paused
			is_running := false
		return
		case "resume": ; Not currently in use
			timestamp := A_TickCount
			FileAppend , % "Resumed, " . A_Hour . ":" . A_Min . ":" . A_Sec . ", Session Resumed, `r`n", % csv_file
			WinSetTitle, Hourly Counter | Running (Resumed)...
			is_running := true
		return
		case "finish":
			FileAppend , % "Finished, " . A_Hour . ":" . A_Min . ":" . A_Sec . ", Session Finished, " . (is_running ? FormatTimeStamp(A_TickCount - timestamp ): "") . " `r`n", % csv_file
			WinSetTitle, Hourly Counter | Finished
			is_running := false
		return
		case "update":
			FileAppend , % "Description, " . A_Hour . ":" . A_Min . ":" . A_Sec . ", " . desc . " , `r`n", % csv_file
			return
		case "exit":
			FileAppend , % "Exited, " . A_Hour . ":" . A_Min . ":" . A_Sec . ", App Exited " . ( A_DD . "/" . A_MMM . "/" . A_YYYY ) . ", " . (is_running ? FormatTimeStamp(A_TickCount - timestamp ): "") . " `r`n`r`n", % csv_file
		return	

	}

}

FormatTimeStamp(delta)  ; Convert the specified number of milliseconds to hh:mm:ss format.
{
	delta := delta/1000
	hours := Format("{:02}", floor(delta/3600))
	mins := Format("{:02}", floor((delta/60) - (hours*60)))
	sec :=  Format("{:02}", floor(delta - hours*3600 - mins*60))
	return % hours . ":" . mins . ":" . sec
}