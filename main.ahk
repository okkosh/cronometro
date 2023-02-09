;@Ahk2Exe-SetMainIcon icon.ico
;@Ahk2Exe-ExeName cronometro
; Script Name - Cronometro - Working hours counter
/**
This script provides a personalised and privacy focused way to pitch in
your working hours by creating a small widget on your screen's bottom right
corner with simple button interface.

TODO: Configurable widget location
*/

IniRead , csv_file, config.ini, file, location
IniRead , username_, config.ini, email, user
IniRead , server_, config.ini, email, server
IniRead , port_, config.ini, email, port
IniRead , from_, config.ini, email, from
IniRead , to_, config.ini, email, to
Cred_suffix := "AHK_Creds"

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
x_gui:= (A_ScreenWidth - 390*A_ScreenDPI/96)
y_gui := (A_ScreenHeight - 150*A_ScreenDPI/96)
global gui_name := "Cronometro"

; Enabled icon set
play_icon := "res\icons8-play-64.png"
pause_icon := "res\icons8-pause-64.png"
stop_icon := "res\icons8-stop-64.png"
edit_file_icon := "res\icons8-edit-file-64.png"
opened_folder_icon := "res\icons8-opened-folder-64.png"
settings_icon := "res\icons8-settings-64.png"
share_icon := "res\icons8-share-rounded-64.png"


; Disabled icon set
play_icon_dis := "res\icons8-play-disabled-64.png"
pause_icon_dis := "res\icons8-pause-disabled-64.png"
stop_icon_dis := "res\icons8-stop-disabled-64.png"
edit_file_icon_dis := "res\icons8-edit-file-disabled-64.png"
opened_folder_icon_dis := "res\icons8-opened-folder-disabled-64.png"
settings_icon_dis := "res\icons8-settings-disabled-64.png"

; Minimise to tray icon
mini_tray := "res\icons8-close-32.png"


; Main Gui
Gui, Color, 2f3337, 000000
Gui, Add, Pic, vstart cWhite gstart x10 y10 w25 h25 , % play_icon
start_TT := "Start/resume a task/session"
Gui, Add, Pic, vstop cWhite  gpause xp+35 yp w25 h25 , % pause_icon_dis
stop_TT := "Pause an already running session"
Gui, Add, Pic, vfinish cWhite  gfinish xp+35 yp w25 h25 , % stop_icon_dis
finish_TT := "Finish a task/session"
Gui, Font, s15, Arial
Gui, Add, Text, vtimer cWhite  xp+35 yp w90 h30 +Center, 00:00:00
Gui, Add, Pic, vdesc cWhite gdesc xp+100 yp w25 h25 , % edit_file_icon_dis
desc_TT := "Add task's description"
Gui, Add, Pic, xp+35 yp cWhite  vopen_folder gopen w25 h25 , % opened_folder_icon
open_folder_TT := "Open worksheet directory"
Gui, Add, Pic, xp+35 yp cWhite  vconfigure_crono gconfigure w25 h25 , % settings_icon
configure_crono_TT := "Cronometro settings"
Gui, Add, Pic, xp+35 yp cWhite  vshare_hours gshare w25 h25 , % share_icon
configure_crono_TT := "Share working hours"

Gui, Add, Pic, xp+27 y5 gGuiClose w12 h12 , % mini_tray
; Show the main gui
Gui, Show, w355 h15 x%x_gui% y%y_gui%, % gui_name
Gui, +LastFound +AlwaysOnTop -Caption  +Owner
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
Gui, config: Add, CheckBox, von_start gautoStart  %is_auto_start% x32 y20 w210 h20 , Start with windows
Gui, config: Add, Text, x32 yp+30 w210 h20 , Split worksheet files based on
Gui, config: Add, Radio, vworsheets_split xp yp w0 h0, ; Fake label to store radio info
; This will make sure that your last configuration will be selected on the radio control
for key, label in radio_labels {
	radio_checked :=
	if (key = split_checked)
		radio_checked := "Checked"
	Gui, config: Add, Radio, %radio_checked% xp yp+20 h20 , % label
}
Gui, config: Add, Text, x22 yp+20 w200 h20 , Save worksheets to
Gui, config: Add, Edit, vcsv_folder_browser ReadOnly x22 yp+20 w200 h20 , %csv_folder%
Gui, config: Add, Button, x122 yp+20 w100 h20 , Browse

Gui, config: Add, Text, x240 y10 h20 , Email Settings:
Gui, config: Add, Text, x240 yp+20 h20 , Server (SMTP/IMAP)
Gui, config: Add, Edit, vserver xp yp+20 w200 h20, %server_%
Gui, config: Add, Text, xp yp+30 w200 h20 , Port (587/25)
Gui, config: Add, Edit, vport xp yp+20 w200 h20, %port_%
Gui, config: Add, Text, xp yp+30 w200 h20 , Username
Gui, config: Add, Edit, vusername xp yp+20 w200 h20, %username_%
Gui, config: Add, Text, xp yp+30 w200 h20 , Password
Gui, config: Add, Edit, vpassword password xp yp+20 w200 h20, %password_%
Gui, config: Add, Text, xp yp+30 w200 h20 , From Email
Gui, config: Add, Edit, vfrom xp yp+20 w200 h20, %from_%
Gui, config: Add, Text, xp yp+30 w200 h20 , To Email
Gui, config: Add, Edit, vto xp yp+20 w200 h20, %to_%

Gui, config: Add, Button, x110 yp+30 w200 h30 , Save

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
; Handles Mouse hovers for tooltips
OnMessage(0x200, "WM_MOUSEMOVE")
; Make window a bit transparent
WinSet, Transparent, 240, % gui_name
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
	; Set Icons accordingly
	GuiControl,, start, % play_icon_dis
	GuiControl,, stop, % pause_icon
	GuiControl,, finish, % stop_icon
	GuiControl,, desc, % edit_file_icon

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
	; Set Icons accordingly
	GuiControl,, start, % play_icon
	GuiControl,, stop, % pause_icon_dis
	GuiControl,, desc, % edit_file_icon_dis

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
	; Set Icons accordingly
	GuiControl,, start, % play_icon
	GuiControl,, stop, % pause_icon_dis
	GuiControl,, finish, % stop_icon_dis
	ex_action("finish")
	SetTimer , stopwatch, Off
return

share:
	if server_ =
	{
		MsgBox, 64, Email not configured, Please configure your email settings first.
		return
	}
	TrayTip, Cronometro, Sharing your yesterday's report, 2, 17
	creds := CredRead(Cred_suffix)
	password := creds.password

	L_day := A_DD - 1
	Year := A_YYYY
	if (L_day <= 0){
		if (A_MM = 01){
			Month := "Dec"
			Year := A_YYYY-1
		}
		L_day := last_day(Month)
	}
	total_time := "00:00:00"
	mail_body := ""

	date_ := L_day . "/" . A_MMM . "/" . A_YYYY

	Loop, Read, % csv_file
	{
		If InStr(A_LoopReadLine, date_ ) {
			mail_body := mail_body . "<tr>"
			Loop, Parse, A_LoopReadLine, CSV
			{
				if (A_index == 4) { ; Last field

					if (A_LoopField <> " ") {
						total_time := add_time(total_time, A_LoopField)
					}
					mail_body :=  mail_body . "<td><b>" .  A_LoopField . "</b></td>"
				} else {
					mail_body :=  mail_body . "<td>" .  A_LoopField . "</td>"
				}

			}
			mail_body := mail_body . "</tr>"
		}
	}

	raw_html =
		(<!DOCTYPE html>
		<html><style>table, th, td `{border:1px solid black;`}</style><body><h2>Working hours on %date_%</h2><table style='width:100`%; border: 1px solid black'>
		<tr><th>Action</th><th>Date</th><th>Time</th><th>Total</th></tr>%mail_body%</table><br/><h4>Total Time Worked: %total_time%</h4><p>Auto generated by Cronometro</p></body></html>
		)
	Runwait, %A_WorkingDir%\bin\SwithMail.exe /s /from %from_% /pass %password% /Server %server_% /Port %port_% /to %to_% /TLS /HTML "true" /sub "Working hours %date_%" /b "%raw_html%", , hide
	MsgBox, 64, Shared, Your previous day report has been shared with %to_%
return

; Description of a task is updated.
taskButtonDone:
	Gui, task: submit, Nohide
	ex_action("update", description)
	WinHide , Add Task Description
return

; Configure Chronometro
configure:
	Gui, config: Show, , Configure
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
	IniWrite , %username% , config.ini, email, user
	IniWrite , %server% , config.ini, email, server
	IniWrite , %port% , config.ini, email, port
	IniWrite , %from% , config.ini, email, from
	IniWrite , %to% , config.ini, email, to
	CredWrite(Cred_suffix, username, password)
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
			if (is_running or is_paused)
				FileAppend , % "Exited/Finished, " . ( A_DD . "/" . A_MMM . "/" . A_YYYY ) . ", " . A_Hour . ":" . A_Min . ":" . A_Sec . ", "
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

; Method to get last day of any month
last_day(month){
	switch (month){
		case 1, 3, 5, 7,8,10,12:
			return 31
		case 4,6,9,11:
			return 30
		case 2:
			if ((Mod(A_YYYY,4) == 0) && (Mod(A_YYYY,4) || Mod(A_YYYY,100)!= 0)){
				return 28
			}
			return 27

	}
}

; Add two times/hours
add_time(work_1, work_2){
	time_1 := StrSplit(work_1, ":")
	time_2 := StrSplit(work_2, ":")

	final_time_s :=  time_1[3] + time_2[3]
	final_time_m := (time_1[2] + time_2[2]) + floor(final_time_s/60)
	final_time_h := (time_1[1] + time_2[1]) + floor(final_time_m/60)

	return  Format("{:02}",final_time_h) . ":" .  Format("{:02}", Mod(final_time_m, 60)) . ":" .  Format("{:02}",Mod(final_time_s, 60))
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

; This snippet is taken from AHK forums for tooltip on hover
WM_MOUSEMOVE(wparam, lParam, msg, hwnd)
{
    static CurrControl, PrevControl, _TT  ; _TT is kept blank for use by the ToolTip command below.
    CurrControl := A_GuiControl
    If (CurrControl <> PrevControl and not InStr(CurrControl, " "))
    {
        ToolTip  ; Turn off any previous tooltip.
        SetTimer, DisplayToolTip, 1000
        PrevControl := CurrControl
    } else {
		if wparam = 1 ; LButton
			PostMessage, 0xA1, 2,,, A ; WM_NCLBUTTONDOWN
	}

    return


	DisplayToolTip:
	try
		ToolTip % %CurrControl%_TT
	catch
		ToolTip
	SetTimer, RemoveToolTip, -2000
	return

	RemoveToolTip:
	ToolTip
	return
}

/**
 Snippet taken from ahk forums

*/
CredWrite(name, username, password)
{
	VarSetCapacity(cred, 24 + A_PtrSize * 7, 0)
	cbPassword := StrLen(password)*2
	NumPut(1         , cred,  4+A_PtrSize*0, "UInt") ; Type = CRED_TYPE_GENERIC
	NumPut(&name     , cred,  8+A_PtrSize*0, "Ptr")  ; TargetName = name
	NumPut(cbPassword, cred, 16+A_PtrSize*2, "UInt") ; CredentialBlobSize
	NumPut(&password , cred, 16+A_PtrSize*3, "UInt") ; CredentialBlob
	NumPut(3         , cred, 16+A_PtrSize*4, "UInt") ; Persist = CRED_PERSIST_ENTERPRISE (roam across domain)
	NumPut(&username , cred, 24+A_PtrSize*6, "Ptr")  ; UserName
	return DllCall("Advapi32.dll\CredWriteW"
	, "Ptr", &cred ; [in] PCREDENTIALW Credential
	, "UInt", 0    ; [in] DWORD        Flags
	, "UInt") ; BOOL
}

CredDelete(name)
{
	return DllCall("Advapi32.dll\CredDeleteW"
	, "WStr", name ; [in] LPCWSTR TargetName
	, "UInt", 1    ; [in] DWORD   Type,
	, "UInt", 0    ; [in] DWORD   Flags
	, "UInt") ; BOOL
}

CredRead(name)
{
	DllCall("Advapi32.dll\CredReadW"
	, "Str", name   ; [in]  LPCWSTR      TargetName
	, "UInt", 1     ; [in]  DWORD        Type = CRED_TYPE_GENERIC (https://learn.microsoft.com/en-us/windows/win32/api/wincred/ns-wincred-credentiala)
	, "UInt", 0     ; [in]  DWORD        Flags
	, "Ptr*", pCred ; [out] PCREDENTIALW *Credential
	, "UInt") ; BOOL
	if !pCred
		return
	name := StrGet(NumGet(pCred + 8 + A_PtrSize * 0, "UPtr"), 256, "UTF-16")
	username := StrGet(NumGet(pCred + 24 + A_PtrSize * 6, "UPtr"), 256, "UTF-16")
	len := NumGet(pCred + 16 + A_PtrSize * 2, "UInt")
	password := StrGet(NumGet(pCred + 16 + A_PtrSize * 3, "UPtr"), len/2, "UTF-16")
	DllCall("Advapi32.dll\CredFree", "Ptr", pCred)
	return {"name": name, "username": username, "password": password}
}
