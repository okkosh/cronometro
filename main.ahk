;@Ahk2Exe-SetMainIcon icon.ico
;@Ahk2Exe-ExeName cronometro
; Script Name - Cronometro - Working hours counter
/**
This script provides a personalised and privacy focused way to pitch in
your working hours by creating a small widget on your screen's top left
corner with simple three button interface.

TODO: Configurable widget location
*/

IniRead , csv_file, config.ini, file, location

; Use a default csv_file in case of non-existent location
if not csv_file {
	csv_file := A_ScriptDir . "\" . get_filename()
} else {
	csv_file := csv_file . "\" . get_filename()
}

; Checks for the autostart
if FileExist(A_Startup . "\oc.lnk")
	is_auto_start:= "checked"


global timestamp ; Timestamp to compare the time with
 ; Checks the current state running/paused
global is_running := false
global is_paused := false
; Time difference/delta to keep track of between each pause/start
global time_delta := 0

; Set gui positioning to bottom right corner
x_gui:= (A_ScreenWidth - 350*A_ScreenDPI/96)
y_gui := (A_ScreenHeight - 150*A_ScreenDPI/96)
global gui_name := "Cronometro"

; Main Gui
Gui, Color, White, Black
Gui, Font, s15, Arial
Gui, Add, Button, vstart gstart x10 y10 w30 h30 , % Chr(9210)
Gui, Add, Button, vstop gpause xp+35 yp w30 h30 , % Chr(9208)
Gui, Add, Button, vfinish gfinish xp+35 yp w30 h30 , % Chr(9209)
Gui, Add, Button, vdesc gdesc xp+35 yp w30 h30 , % Chr(128221)
Gui, Add, Text, vtimer xp+35 yp w90 h30 +Center, 00:00:00
Gui, Add, Button, xp+100 yp gopen w30 h30 , % chr(128194)
Gui, Add, Button, xp+35 yp gconfigure w30 h30 , % chr(128736)
; Show the main gui
Gui, Show, w330 h50 x%x_gui% y%y_gui%, % gui_name
Gui, +LastFound +AlwaysOnTop +ToolWindow +Owner
; Disable irrelevant controls (Since this is our first run)
GuiControl , Disable, stop
GuiControl , Disable, desc
GuiControl , Disable, finish

; Task window to add task descriptions
Gui, task: Add, Text, x12 y10 w210 , Add Task Description`r`n(Describe complete details of your Task)
Gui, task: Font, s15, Arial
Gui, task: Add, Button, Disabled x262 y9 w60 h30,% chr(128247)
Gui, task: Font, ,
Gui, task: Add, Edit, x12 y49 w310 h120 vdescription,
Gui, task: Add, Button, x122 y179 w100 h30 , Done

; Retrieve folder from csv_file to show inside the configuration window
SplitPath , csv_file, , csv_folder
split_checked := get_filename(true)
; labels are based according to the config
; Do not mess this up unless you are absolutely sure about what you are doing
radio_labels := ["Current day", "Current month", "App run", "Current year"]

; Create a configuration window
Gui, config: Add, CheckBox, von_start gautoStart  %is_auto_start% x32 y9 w210 h20 , Start with windows
Gui, config: Add, Text, x32 y39 w210 h20 , Split worksheet files based on
Gui, config: Add, Radio, vworsheets_split xp yp w0 h0, ; Fake label to store radio info
; This will make sure that your last configuration will be selected on the radio control
for key, label in radio_labels {
	radio_checked :=
	if (key = split_checked)
		radio_checked := "Checked"
	Gui, config: Add, Radio, %radio_checked% xp yp+20 h20 , % label
}
Gui, config: Add, Edit, vcsv_folder_browser ReadOnly x22 y169 w200 h20 , %csv_folder%
Gui, config: Add, Button, x122 y199 w100 h20 , Browse
Gui, config: Add, Button, x62 y239 w100 h30 , Save
Gui, config: Add, Text, x22 y149 w200 h20 , Save worksheets to

; Make sure to log the App run
if not FileExist(csv_file)
	FileAppend , % "ACTION, DATE, TIME, TOTAL`r`n", % csv_file

; Customise the tray menu
Menu, Tray, NoStandard
Menu, Tray, Add, Show, show
Menu, Tray, Add, Configure, configure
Menu, Tray, Add, Open Worksheet Folder, open
Menu, Tray, Add, Exit, exit

; Handles System Shutdown and Logoff
OnMessage(0x11, "WM_QUERYENDSESSION")
; Handles Tray clicks
OnMessage(0x404, "AHK_NOTIFYICON")
return

; Show the main gui
show:
	WinShow , % gui_name
	return

; Pressing cross button won't let you exit the app
; TODO: Make it configurable
GuiClose:
	WinHide , % gui_name
	return

; Actual app exit
; TODO : Prevent shutdown action to destroy app data
; while the app is tracking time in the background
exit:
	ex_action("exit")
ExitApp

; Start the time tracking
start:
	; Enable and disable controls according to this action
	; TODO: optimise disabling by removing redundant code
	GuiControl , Disable, start
	GuiControl , Enable, stop
	GuiControl , Enable, desc
	GuiControl , Enable, finish

	if is_paused{
		ex_action("resume")
	} else {
		ex_action("start")
	}
	; start a timer watch
	; for aesthetics purpose only
	SetTimer , stopwatch, 1
return

; Pauses the time tracking
pause:
	GuiControl , Enable, start
	GuiControl , Disable, stop
	GuiControl , Disable, desc
	ex_action("pause")
	SetTimer , stopwatch, Off
	return

; Add task desctiption (Will only work when you are running a task)
desc:
	Gui, task: Show, w336 h219, Add Task Description
	return

; finish a task
finish:
	GuiControl, Enable, start
	GuiControl, Disable, stop
	GuiControl , Disable, finish
	ex_action("finish")
	SetTimer , stopwatch, Off
return

; Description of a task is updated.
taskButtonDone:
	Gui, task: submit, Nohide
	ex_action("update", description)
	WinHide , Add Task Description
return

; Configure Chronometro
configure:
	Gui, config: Show, w240 h281 , Configure
return

; Open the folder containing worksheets
open:
	Run, % csv_folder
return

; A time controlled subroutine to show timer on gui
stopwatch:
	GuiControl , , timer, % FormatTimeStamp(A_TickCount - timestamp + time_delta)
return

; A g-label to make sure that this script link
; will be added into the startup folder
autoStart:
	Gui, config:submit, nohide
	if on_start
		FileCreateShortcut, %A_ScriptFullPath%, %A_Startup%\oc.lnk, %A_ScriptDir%
	else
		FileDelete, %A_Startup%\oc.lnk
	return

; Save configurations and make them effective
configButtonSave:
	Gui, config:submit, nohide
	; worksheets_split -1 because, we are using a fake radio button
	; to store our variable
	set_filename(worsheets_split-1)
	; Make sure that we are using the same file based on
	; worksheet settings
	IniWrite , %csv_folder% , config.ini, file, location
	csv_file := csv_folder . "/" . get_filename()
	; Hide the main configuration window
	WinHide , Configure
	return

; Browse Button for selecting the worksheet folder
configButtonBrowse:
	Gui, config:submit, nohide
	FileSelectFolder , csv_folder, csv_folder, , Select a folder to save your worksheets
	if csv_folder {
		GuiControl , config: , csv_folder_browser, %csv_folder%
	}
	return

/**
Method name - ex_action - execute action
params:-
	command - action name based on button pressed. start, stop/pause, resume, finish, update, exit
	desc - Task description (optional)
returns :-
	null
*/
ex_action(command:="start", desc := ""){
	; Use the golbal namespace variables
	global timestamp, is_running, csv_file, is_paused

	; There are two main parts of a particular action
	; 1. Appending the CSV file based on the action selected
	; 2. Updating the is_running/is_paused status

	switch (command){
		case "start":
			timestamp := A_TickCount
			FileAppend , % "`r`nSession Started, " . ( A_DD . "/" . A_MMM . "/" . A_YYYY ) . ", " . A_Hour . ":" . A_Min . ":" . A_Sec . ", `r`n", % csv_file
			WinSetTitle, % gui_name . "| Running..."
			is_running := true
		return
		case "stop":
		case "pause":
			FileAppend,  % "Paused (Total Session - " .  FormatTimeStamp(A_TickCount - timestamp + time_delta) . ") , " . ( A_DD . "/" . A_MMM . "/" . A_YYYY ) . ", " . A_Hour . ":" . A_Min . ":" . A_Sec . ", `r`n", % csv_file
			WinSetTitle, % gui_name . "| Paused"
			is_running := false
			is_paused := true
			time_delta := A_TickCount - timestamp + time_delta
		return
		case "resume":
			timestamp := A_TickCount
			FileAppend , % "Session Resumed, " . ( A_DD . "/" . A_MMM . "/" . A_YYYY ) . ", " . A_Hour . ":" . A_Min . ":" . A_Sec . ", `r`n", % csv_file
			WinSetTitle, % gui_name . "| Running (Resumed)..."
			is_paused := false
			is_running := true
		return
		case "finish":
			FileAppend , % "Finished, " . ( A_DD . "/" . A_MMM . "/" . A_YYYY ) . ", " . A_Hour . ":" . A_Min . ":" . A_Sec . ", " . (is_running ? FormatTimeStamp(A_TickCount - timestamp + time_delta): FormatTimeStamp(time_delta)) . "`r`n", % csv_file
			WinSetTitle,  % gui_name . "| Finished"
			is_running := false
			is_paused := false
			time_delta := 0
			GuiControl , , timer,  % FormatTimeStamp(0)
		return
		case "update":
			FileAppend , % " " """" . desc . " "" [Info Added], "  . ( A_DD . "/" . A_MMM . "/" . A_YYYY ) . ", " . A_Hour . ":" . A_Min . ":" . A_Sec . ",`r`n", % csv_file
			return
		case "exit":
			; Only log the elapsed time when our app is actively running/paused as a failsafe against aburptly exiting the app
			FileAppend , % "Exited, " . ( A_DD . "/" . A_MMM . "/" . A_YYYY ) . ", " . A_Hour . ":" . A_Min . ":" . A_Sec . ", "
			. (is_running ? FormatTimeStamp(A_TickCount - timestamp + time_delta):  ( is_paused ?  FormatTimeStamp(time_delta) : "" )) . " `r`n`r`n", % csv_file
		return
	}
}

/**
Method name - get_filename
params:-
	raw - whether to return the actual index of radio control (i.e. raw index) or not
returns :-
	filename based on the selection of our worksheet split type in configuration (default month based)
*/
get_filename(raw:=false){
	IniRead , csv_split, config.ini, file, split
	switch (csv_split){
		case "day":
			return  raw ? 1 : "working_hours_raw_" . A_DD . "_" . A_MMM . "_" . A_YYYY . ".csv"
		case "month":
			return raw ? 2 : "working_hours_raw_" . A_MMM . "_" . A_YYYY . ".csv"
		case "app":
			return raw ? 3 : "working_hours_app_run_" . A_Now . A_MMM . "_" . A_YYYY . ".csv"
		case "year":
			return raw ? 4 : "working_hours_raw_" . A_YYYY . ".csv"
		case Default:
			return raw ? 2: "working_hours_raw_" . A_MMM . "_" . A_YYYY . ".csv"
	}
}

/**
Method name - set_filename - used internally for the configuration purpose only
params:-
	Index - Index of the radiobutton
returns :-
	null
*/
set_filename(index){
	switch (index){
		case 1: IniWrite , % "day" , config.ini, file, split
		case 2: IniWrite , % "month" , config.ini, file, split
		case 3: IniWrite , % "app" , config.ini, file, split
		case 4: IniWrite , % "year" , config.ini, file, split
	}
}

/**
Method name - FormatTimeStamp - Convert the specified number of milliseconds to hh:mm:ss format.
params:-
	delta - Total time elapsed in ms
returns :-
	Time string -> Number of hours:mins:secs elapsed
*/
FormatTimeStamp(delta)
{
	delta := delta/1000
	hours := Format("{:02}", floor(delta/3600))
	mins := Format("{:02}", floor((delta/60) - (hours*60)))
	sec :=  Format("{:02}", floor(delta - hours*3600 - mins*60))
	return % hours . ":" . mins . ":" . sec
}

/**
The following snippets has been taken from the Autohotkey Documentation
*/

;Clicking the icon will lead to Main window
AHK_NOTIFYICON(wParam, lParam, uMsg, hWnd)
{
	global gui_name
	if (lParam = 0x0201)
		WinShow , % gui_name
}

WM_QUERYENDSESSION(wParam, lParam)
{
	global is_running, is_paused

    ENDSESSION_LOGOFF := 0x80000000
    if (lParam & ENDSESSION_LOGOFF)  ; User is logging off.
        EventType := "Logoff"
    else  ; System is either shutting down or restarting.
        EventType := "Shutdown"
    try
    {
		if (is_running or is_paused) {
			; Set a prompt for the OS shutdown UI to display.  We do not display
			; our own confirmation prompt because we have only 5 seconds before
			; the OS displays the shutdown UI anyway.  Also, a program without
			; a visible window cannot block shutdown without providing a reason.
			BlockShutdown("Chronometro attempting to prevent " EventType ".")
			return false
		} else {
			ex_action("exit")
			return true
		}
    }
    catch
    {
        ; ShutdownBlockReasonCreate is not available, so this is probably
        ; Windows XP, 2003 or 2000, where we can actually prevent shutdown.
        MsgBox, 4,, %EventType% in progress.  Allow it?
        IfMsgBox Yes
		{
			ex_action("exit")
			return true  ; Tell the OS to allow the shutdown/logoff to continue.
        } else {
            return false  ; Tell the OS to abort the shutdown/logoff.
		}
    }
}

BlockShutdown(Reason)
{
    ; If your script has a visible GUI, use it instead of A_ScriptHwnd.
    DllCall("ShutdownBlockReasonCreate", "ptr", A_ScriptHwnd, "wstr", Reason)
    OnExit("StopBlockingShutdown")
}

StopBlockingShutdown()
{
	ex_action("exit")
    OnExit(A_ThisFunc, 0)
    DllCall("ShutdownBlockReasonDestroy", "ptr", A_ScriptHwnd)
}
