; How to change default editor: change in registry regedit: HKEY_CLASSES_ROOT\AutoHotkeyScript\Shell\Edit\Command -> notepad++ instead of notepad

; #-Win... !-Alt..... ^-Ctrl..... + shift

; ^+a:: send report_switching_activity [get_nets 
; ^+s:: send select_object [get_nets 


!r:: ; Reload this script after changes
	reload, %A_ScriptFullPath% 
	return 
													
;***************************Key Control Shortcuts*************************

;  Rename XPP to PDM
;  ^+p::
;  send {CTRLDOWN}a{CTRLUP}
;  send {CTRLDOWN}c{CTRLUP}
;  StringReplace, clipboard, clipboard,  XPP, PDM, All
;  StringReplace, clipboard, clipboard,  Xilinx Power Planner, Power Design Manager, All
;  send {CTRLDOWN}v{CTRLUP}
;  sleep 500
;  send {TAB}
;  sleep 1000
;  send {CTRLDOWN}f{CTRLUP}
;  sleep 1000
;  send XPP
;  send {TAB}
;  sleep 1000
;  send PDM
;  send {TAB}
;  sleep 100
;  send {TAB}
;  sleep 100
;  send {TAB}
;  sleep 100
;  send {SPACE}
;  sleep 2000
;  send {CTRLDOWN}f{CTRLUP}
;  sleep 1000
;  send {BACKSPACE}{BACKSPACE}{BACKSPACE}
;  send Xilinx Power Planner
;  send {TAB}
;  sleep 1000
;  send Power Design Manager
;  send {TAB}
;  sleep 300
;  send {TAB}
;  sleep 300
;  send {TAB}
;  sleep 300
;  send {TAB}
;  sleep 300
;  send {SPACE}
;  sleep 1000
;  return



$*Shift:: ; Pressing shift will almost constrain mouse to a horizontal movement
Send {Shift Down}
MouseGetPos, ox, oy
SetTimer, WatchTheMouse, 1
KeyWait Shift
return

$*Shift Up::
Send {Shift Up}
SetTimer, WatchTheMouse, off
return

WatchTheMouse:
MouseGetPos, nx, ny
MouseMove nx, oy, 0
return


^+v:: ; Copy Paste content. Very useful for VNC screen when copy is not working fine
sendMode Input
StringReplace, clipboard, clipboard, `r`n, , All
send %ClipBoard%
sendMode Event
return

;**************** Skype shortcuts - My org doesn't use Skype anymore *************
;   ChangeSkypeMode(mode)
;   {
;      IfWinNotActive, Skype for Business, , WinActivate, Skype for Business, 
;      WinWaitActive, Skype for Business, 
;      Send, {ALTDOWN}f{ALTUP}
;      sleep 100
;      Send m
;      Sleep 500
;      Send %mode%
;      Sleep 100
;      return
;   }
;   
;   !a::  ; Skype Reset the status - Available
;      ChangeSkypeMode("t")
;      return
;   
;   !b::  ; Skype DND - Busy
;      ChangeSkypeMode("d")
;      return
;   
;   !c::  ;  Skype offline - Cut off
;      ChangeSkypeMode("f")
;      return
;   return

; ************ Slack Short-cuts ***************


;****************************Application shortcuts**************************

GetWeblink()
{
	sleep, 100
	send, {F6}
	sleep, 500
   Send, {CTRLDOWN}c{CTRLUP}
   sleep, 500
   Send, {ENTER}
   sleep, 1000
	return, %clipboard%
	
}
GetDateTime()
{
	FormatTime, var_d , , dd
	FormatTime, var_M , , MM
	FormatTime, var_y , , yyyy
	FormatTime, var_min , , mm
	FormatTime, var_s , , ss
	FormatTime, var_H , , HH
	DATETIME := var_y . "/" . var_M . "/" . var_d . " - " . var_H . ":" . var_min . ":" . var_s
	return DATETIME
}
GetDate()
{
	FormatTime, var_d , , dd
	FormatTime, var_MM1 , , MMM
	FormatTime, var_y , , yyyy

	StringUpper, var_m, var_MM1  

	date:= var_d . var_m . var_y
   return, date
}

!d:: ; Curent DATE in DDMMYYYY Format
   date := GetDate()
	;msgbox % date
	SendMode Input
	Send % date
	SendMode, Event
return

#d::  ;clipboard change from linux to windows path (Windows+D)
	; Sandbox Dir
	StringReplace, clipboard, clipboard,  /proj/cogsparusr/sandboxes/darshakk, U:, All
	; XCO Staff
	StringReplace, clipboard, clipboard,  /proj/xcohdstaff2/darshakk, W:, All
	; XSJ Staff
	StringReplace, clipboard, clipboard,  /proj/xsjhdstaff1/darshakk, Y:, All
	;msgbox %clipboard%
	;XCO Staff nobkup
	StringReplace, clipboard, clipboard, /wrk/xcohdnobkup2/darshakk, T:, All
	;XSJ Staff nobkup
	StringReplace, clipboard, clipboard, /wrk/xsjhdnobkup1/darshakk, Q:, All
	; XCO Dese Dir
	StringReplace, clipboard, clipboard, /wrk/dese1/darshakk, S:, All
	StringReplace, clipboard, clipboard, /proj/dese/users/darshakk, S:, All

	StringReplace, clipboard, clipboard, /, \, All
return

!t::	;; time and date stamp

	DATETIME:= getDateTime()
	SendMode Input
	send %DATETIME%
	SendMode Event
	
return	

FileNameFromTitle(winTitle)
{
   fileName = %winTitle%
   fileName := RegExReplace(fileName, "[+]", "p")
   fileName := RegExReplace(fileName, "[^a-zA-Z0-9]", " ")
   fileName := RegExReplace(fileName, "\s+", "_")
   fileName := RegExReplace(fileName, "_+", "_")
   return, fileName
}

!s:: ;snip of entire screen

   folderPath = C:\Users\%A_UserName%\MySnippets
   IfNotExist, %folderPath%
   {
      FileCreateDir, %folderPath%
   }
   date := GetDate()
  
   ; Move mouse to center of active window
   sleep, 100
   WinGetTitle, winTitle, A
   CoordMode, Mouse, Relative
   WinGetActiveStats, AWTitle, AWWidth, AWHeight, AWX, AWY
   MPosX := (AWWidth//2)
   MPosY := (AWHeight//2)
   MouseMove, %MPosX%, %MPosY%
 
   fileName := FileNameFromTitle(winTitle)
   ; Create file name by looking at file count in folder
   fileCount := ComObjCreate("Scripting.FileSystemObject").GetFolder(folderPath).Files.Count
   fileCount := SubStr("00000" . fileCount,-4)
   if !WinExist("Snipping Tool")
   {
      RunCmd("snippingtool.exe")
   }
   WinWait, Snipping Tool, 
   IfWinNotActive, Snipping Tool, , WinActivate, Snipping Tool, 
   WinWaitActive, Snipping Tool, 
   Send, {ALTDOWN}m{ALTUP}
   sleep 300
   Send, W
   sleep 500
   MouseClick, left
   sleep 1000
   WinWait, Snipping Tool, 
   IfWinNotActive, Snipping Tool, , WinActivate, Snipping Tool, 
   WinWaitActive, Snipping Tool, 
   sleep, 300
   Send, {CTRLDOWN}c{CTRLUP}
   sleep, 300
   Send, {ALTDOWN}f{ALTUP}
   sleep 300
   Send, x
   sleep 300
   Send {ENTER}
   WinWait, Save As, 
   IfWinNotActive, Save As, , WinActivate, Save As,
   WinWaitActive, Save As,
   sendMode Event
   SendInput %folderPath%\%fileCount%_%fileName%_%date%
   sleep 1000
   Send {TAB}
   sleep 300
   send, j
   sleep 300
   send, {ENTER}
   sleep 100
   return

!p:: ; Full page PDF in Chrome

FullPageScreenShot:
	link := GetWeblink()
	ToolTip, Downloading %link%, A_ScreenWidth //3, A_ScreenHeight //2

	datetime:= GetDateTime()
   folderPath = C:\Users\%A_UserName%\MyWebPages
   IfNotExist, %folderPath%
   {
      FileCreateDir, %folderPath%
   }
   failedList = %folderPath%\failedDownloads.txt
   IfNotExist, %failedList%
   {
		FileAppend, Websites that failed to download:`n, %failedList%
   }
   succededList = %folderPath%\succededDownloads.txt
   IfNotExist, %succededList%
   {
		FileAppend, Websites that successeded to download:`n, %succededList%
   }
  
   ; Check if we are running chrome
   WinGetTitle, winTitle, A
   fileName := StrReplace(winTitle, " - Google Chrome","")
   ; MsgBox, %fileName%
   if(fileName == winTitle) ;not a chrome window
      return
   
   fileName := FileNameFromTitle(fileName)
   ; Create file name by looking at file count in folder
   fileCount := ComObjCreate("Scripting.FileSystemObject").GetFolder(folderPath).Files.Count
   fileCount := SubStr("00000" . fileCount,-4)
   
   ; Commands: https://www.autohotkey.com/docs/commands/Send.htm
   ; Full screen capture using Chrome extension 
   CaptureScreen:
   Loop, 20
   {
   SendInput {PgDn}
   Sleep 100
   }
   Sendinput {Alt Down}
   Sleep 10
   Sendinput {Shift Down}
   Sleep 10
   Sendinput p
   Sleep 10
   Sendinput {Alt Up}
   Sleep 10
   Sendinput {Shift Up}
   Sleep 100
   WinWaitActive, Screen Capture Result, ,30,
   Sleep 5000
	IfWinNotActive, Screen Capture Result,
	{
		; Goto, CaptureScreen
		FileAppend, %datetime%, %failedList%
		FileAppend, %A_Tab%, %failedList%
		FileAppend, %link%`n, %failedList%
		Tooltip
		return
	}
   Sleep 300
   ClickDownloadPdf:
   Sleep 300
   Sendinput {Ctrl Down}
   Sleep 300
   SendInput f
   Sleep 300
   SendInput {Ctrl Up}
   Sleep 300
   SendInput Download PDF
   Sleep 300
   SendInput {Enter}
   Sleep 1000
   SendInput {Esc}
   Sleep 1000
   Sendinput {Enter}
   Sleep 1000
   WinWaitActive, Save As, ,5
   IfWinNotActive, Save As,
   {
      Goto, ClickDownloadPdf
   }
   Sleep 1000
   sendMode Event
   SendInput %folderPath%\%fileCount%_%fileName%.pdf
   Sleep 300
   SendInput {Enter}
   Sleep 1000
   WinWaitActive, Screen Capture Result, ,10,
   SendInput {Ctrl Down}
   Sleep 300
   SendInput w
   Sleep 300
   SendInput {Ctrl Up}
   Sleep 1000
	FileAppend, %datetime%, %succededList%
	FileAppend, %A_Tab%, %succededList%
	FileAppend, %link%`n, %succededList%
	Clipboard = %folderPath%\%fileCount%_%fileName%.pdf
	Tooltip
return

RunCmd(cmd)		;executes command in Run window
{
   send {LWin Down}r{LWin Up}
   sleep 500
   IfWinNotActive, Run, , WinActivate, Run, 
   WinWaitActive, Run, 
   sendMode Input
   sleep 100
   send %cmd%
   sendMode Event
   sleep 200
   send {Enter}
   return
}


RunJavaScript(inlineScript)		;executes javascript in url bar
{
	Send {F6}			;default key to focus the url in chrome 
	Sleep 100
	SendInput javascript:
	Sleep 100
	SendInput %inlineScript%
	Sleep 100
	SendInput {Enter}
	Sleep 1000
   return
}


!i:: ; Read links.txt and PDF print all the pages
	FileSelectFile, LinksFile, 3, , Open file contains links, Text Documents (*.txt)
	if (LinksFile = "")
	{
		MsgBox, No valid file selected. 
		return
	}
	else
	{
		FileRead, links, %LinksFile%
		Loop, Parse, links, `n, `r ; line by line
		{
			link:= A_LoopField
			SetTitleMatchMode, 2
			WinActivate, ahk_exe chrome.exe
			Send ^t
			WinWaitActive, New Tab
			sendMode Event
			send %link%
			sleep 100
			send {enter}
			sleep 100
			while (A_Cursor = "AppStarting")
				continue
			sleep 100
         GoSub, FullPageScreenShot
			sleep 100
			Send ^w
		}
	}
	return



; !i::  ; Educative.io Download A Course by iterating over all lessons
;    Loop, 500
;    {
;       WinGetTitle, prevWinTitle, A
;       ;MsgBox, %prevWinTitle%
;       GoSub, FullPageScreenShot
;       ;MsgBox, %prevWinTitle%
;       Sleep 1000
;       inlineScript = document.getElementById("next_lesson").focus()
;       RunJavaScript(inlineScript) ; focus button
;       Send {Enter} ; Click button
;       Sleep 15000
;       WinGetTitle, winTitle, A
;       if(prevWinTitle == winTitle) ; Button not found or didn't do anything. Exit
;       {
;    	 MsgBox, Download completed
;    	 return     
;       }
;    }
;    return


!l:: ;keep screen unlocked by dancing the mouse
CoordMode, Mouse, Screen
MouseGetPos, CurrentX, CurrentY
Loop {
    Sleep, 100
    LastX := CurrentX
    LastY := CurrentY
    MouseGetPos, CurrentX, CurrentY
    If (CurrentX = LastX and CurrentY = LastY) {
        MouseMove, 5, 5, , R
        Sleep, 100
        MouseMove, -5, -5, , R
    }
}
return
