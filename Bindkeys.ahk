; How to change default editor: change in registry regedit: HKEY_CLASSES_ROOT\AutoHotkeyScript\Shell\Edit\Command -> notepad++ instead of notepad

; #-Win... !-Alt..... ^-Ctrl..... + shift

; ^+a:: send report_switching_activity [get_nets 
; ^+s:: send select_object [get_nets 


!r:: ; Reload this script after changes
	reload, %A_ScriptFullPath% 
	return 
													
;***************************Key Control Shortcuts*************************

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

;***************************Skype shortcuts******************************
ChangeSkypeMode(mode)
{
   IfWinNotActive, Skype for Business, , WinActivate, Skype for Business, 
   WinWaitActive, Skype for Business, 
   Send, {ALTDOWN}f{ALTUP}
   sleep 100
   Send m
   Sleep 500
   Send %mode%
   Sleep 100
   return
}

!a::  ; Skype Reset the status - Available
   ChangeSkypeMode("t")
   return

!b::  ; Skype DND - Busy
   ChangeSkypeMode("d")
   return

!c::  ;  Skype offline - Cut off
   ChangeSkypeMode("f")
   return
return

;****************************Application shortcuts**************************

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

	FormatTime, var_d , , dd
	FormatTime, var_M , , MM
	FormatTime, var_y , , yy
	FormatTime, var_min , , mm
	FormatTime, var_s , , ss
	FormatTime, var_H , , HH
	DATETIME := var_y . var_M . var_d . var_H . var_min . var_s
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
   folderPath = C:\Users\%A_UserName%\MyWebPages
   IfNotExist, %folderPath%
   {
      FileCreateDir, %folderPath%
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
   Sleep 3000
   IfWinNotActive, Screen Capture Result,
   {
	  Send {F5}			;default key to reload in chrome
      Sleep 30000
	  IfWinNotActive, Screen Capture Result,
	  {
		Goto, CaptureScreen
	  }
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
   Clipboard = %folderPath%\%fileCount%_%fileName%.pdf
return

RunCmd(cmd)		;executes javascript in url bar
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


!i::  ; Educative.io Download A Course by iterating over all lessons
   Loop, 500
   {
      WinGetTitle, prevWinTitle, A
      ;MsgBox, %prevWinTitle%
      GoSub, FullPageScreenShot
      ;MsgBox, %prevWinTitle%
      Sleep 1000
      inlineScript = document.getElementById("next_lesson").focus()
      RunJavaScript(inlineScript) ; focus button
      Send {Enter} ; Click button
      Sleep 15000
      WinGetTitle, winTitle, A
      if(prevWinTitle == winTitle) ; Button not found or didn't do anything. Exit
      {
   	 MsgBox, Download completed
   	 return     
      }
   }
   return


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
