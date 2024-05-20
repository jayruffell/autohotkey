;GLOBAL COMMANDS NEEDED TO LOAD FIRST::

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode 2   ;sets window title matching (for #IfWinActive) to partial matches

; add file explorer to a group called 'Explorer', as it has multiple ahk classes
GroupAdd, Explorer, ahk_class CabinetWClass
GroupAdd, Explorer, ahk_class ExploreWClass
Return
;=====================================================================================================
;=====================================================================================================

;-------------------------
; loading / editing AHK script
;-------------------------

;-------------------------
; F KEYS
;-------------------------

; f8:: Send {LWin} ;now debugger shortcut in VSCode - hash out unless I"m on my keyboard that has broken winkey.

; temp hotstrings for current project - "open current folder" and "open current code" in dragon
^+f8:: Run "C:\dev\r\packages\ausmodels\dev"
^+f10:: Run C:\dev\r\packages\ausmodels\dev\
^+f11:: Run C:\dev\r\packages\ausmodels\dev\compare-model-outputs
^+!f12:: Run C:\dev\r\packages\ausmodels\dev\model-outputs

; rest
!f1:: Run C:\dev\r\jon\aussie_models\utility_wind_solar_curt_combined
!f2:: Run C:\dev\r\packages\ausmodels\dev
;!f3:: Run C:\Users\Owner\OneDrive - Haast Energy Trading\Documents\
!f4:: Run C:\dev\HomeHealthCare\
!f5:: Run C:\Users\Owner\Downloads\
;!f12:: Run C:\Program Files (x86)\Common Files\Nuance\NaturallySpeaking15\dragonbar.exe
; !f12:: Run C:\Program Files (x86)\Common Files\Nuance\NaturallySpeaking15\dragonbar_exe_renamedToAvoidWinkeyProblem.exe

f1:: Run C:\Users\Owner\Documents
f2:: Run C:\Users\Owner\Downloads
;f2:: Send {LWinDown} {tab} {LWinUp} ; Show windows including virtual desktops
f3::
 Send {LWinDown} {r} {LWinUp} ; paste clipboard into Run.
 Sleep 500
 Send ^v
 Sleep 500
 Send {enter}
f5:: Send !{f4}
F7:: Run C:\Program Files\WindowsApps\MSTeams_24102.2223.2870.9480_x64__8wekyb3d8bbwe\ms-teams.exe
f9:: #+s ; snipping tool

; CHROME: switch to open program if there is one, otherwise open
f12:: 
if WinExist("Google Chrome")
  WinActivate 
else
  Run C:\Program Files\Google\Chrome\Application\chrome.exe
Return

^F8:: 
if WinExist("to do today")
  WinActivate 
else
  openFileWithBlackNotepad("C:\Users\Owner\Desktop\to do today.txt")
Return

; OUTLOOK: switch to open program if there is one, otherwise open
^F9:: 
if WinExist("Outlook")
  WinActivate 
else
  Run C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE
Return

; OUTLOOK: as above, but for new message window
^+F9:: 
if WinExist("Message")
  WinActivate 
else
  {
        MsgBox, "No message window to switch to"
    }
Return


^+F12:: Run C:\Program Files\Google\Chrome\Application\chrome.exe --profile-directory="Profile 1"
+F12:: Run C:\Program Files\Google\Chrome\Application\chrome.exe --profile-directory="Default"


; VS Code: specific workspaces (more below tho)
^f1:: 
 Send {CtrlUp} {LWin}  
 Sleep 1000 
 Send {Down} 
 Sleep 1000 
 Send {Right}
Return
 
; windows key - NB can just do Send {LWin} if I map straight to an F-key
; ^f2:: Send {LWinDown} {LWinUp} ; not working

; Reference docs
^f3:: Run C:\Users\Owner\OneDrive - Haast Energy Trading\Documents\reference docs

; FILE EXPLORER: switch to open program if there is one, otherwise open
^f4:: 
if WinExist("ahk_group Explorer") ;group is defined at top of script
  WinActivate 
else
  Send #{e}
Return

; WORD: switch to open program if there is one, otherwise open
^f5:: 
if WinExist("Word")
  WinActivate 
else
  Run "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
Return

; NOTEPAD: switch to open program if there is one, otherwise open
^f6:: 
  if WinExist("Black Notepad")
; if WinExist("Notepad")
; if WinExist("Jarte")
  WinActivate 
else {
  ; Run C:\Program Files\WindowsApps\Microsoft.WindowsNotepad_11.2401.26.0_x64__8wekyb3d8bbwe\Notepad\Notepad.exe
  ;Run C:\Windows\system32\notepad.exe
  openBlackNotepadCentred()
  ;  Run C:\Program Files (x86)\Jarte\Jarte.exe
}
Return

; VS Code: switch to open program if there is one, otherwise open
^f7:: 
if WinExist("Visual Studio Code")
  WinActivate
else
  Run C:\Users\Owner\AppData\Local\Programs\Microsoft VS Code\Code.exe
Return

; ; Define a hotkey to activate Power BI Desktop using its window title
; ^F8:: ; This sets the hotkey to Ctrl + Alt + P. You can change it to any other combination.
;     IfWinExist, Power BI Desktop
;     {
;         WinActivate
;     }
;     else
;     {
;         MsgBox, "haven't set up functionality to open PBI... just got to find the executable."
;     }
; return

; Define a hotkey to activate Outlook reminders using window title
^F11:: ; This sets the hotkey to Ctrl + Alt + P. You can change it to any other combination.
    IfWinExist, Reminder
    {
        WinActivate
    }
    else
    {
        MsgBox, "no reminders window based on autohotkey criteria "
    }
return


#+f8:: Run "C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE"
; open AHK script to edit with VS Code
!^f8:: Run C:\Users\Owner\AppData\Local\Programs\Microsoft VS Code\Code.exe "C:\Users\Owner\Documents\autohotkey\" 
; reload AHK script
!^f9:: Run "C:\Users\Owner\Documents\autohotkey\motherscript.ahk" 
^f10:: Run C:\Users\Owner\AppData\Local\Programs\Microsoft VS Code\Code.exe "C:\dev\r\jay\VS-Code-scratchpad\"
;stacks report/power notes stuff
^F12:: Run C:\Program Files\Google\Chrome\Application\chrome.exe --new-window https://haastenergy.atlassian.net/wiki/spaces/ASR/overview
#+F11:: Run C:\Users\Owner\OneDrive - Haast Energy Trading\Documents\reference docs\DUID info.xlsx

; scripts starting "shift + alt + x"
+!1:: Run C:\Program Files\Google\Chrome\Application\chrome.exe https://www.asxenergy.com.au/futures_au
+!2:: Run C:\Program Files\Google\Chrome\Application\chrome.exe https://app.powerbi.com/groups/df569d6a-0f19-4d8e-b25f-9e784c1b0d84/reports/f49573c7-5e50-4429-aec3-8e5b72ab6567/ReportSection ;asx prices
+!3:: Run C:\Program Files\Google\Chrome\Application\chrome.exe https://app.powerbi.com/groups/df569d6a-0f19-4d8e-b25f-9e784c1b0d84/list ;all reports
+!4:: Run C:\Program Files\Google\Chrome\Application\chrome.exe https://app.powerbi.com/groups/df569d6a-0f19-4d8e-b25f-9e784c1b0d84/reports/891a5a05-eb67-44d5-8ebf-a35a20b883e8/ReportSection ;aupm

; Some hacks for feeding into Dragon commands (which won't let me do Windows Key + x commands)
#+!d:: Send {LWinDown} {d} {LWinUp}
#+!r:: Send {LWinDown} {r} {LWinUp}

openBlackNotepad()
{
  Run C:\Program Files\WindowsApps\55809savaged.BlackNotepad_1.0.9.0_neutral__p2aarvyam9pfc\BlackNotepad\BlackNotepad.exe
}
openBlackNotepadCentred()
{
  openBlackNotepad()
  Sleep 1000
  CenterActiveWindow()
}
openFileWithBlackNotepad(file_path)
{
 openBlackNotepadCentred()
send ^o
sleep 1000
send %file_path%
sleep 500
send {enter}
}

;-------------------------
; autoomating "snip it quad" commands
;-------------------------

; HELPER FUNS ----
SnippingTool() { ; hack, was having with calling this another way
  Send #+s
}
HoldDownMouseAndMove(Xfrom, Yfrom, Xto, Yto) ;X Y are coords passed to native Mouseclick fn.
{
  MouseClick, left, Xfrom, Yfrom, ,  , D
  Sleep 1000
  MouseClick, left, Xto, Yto, ,  , U
}
SnipItQuad(Xfrom, Yfrom, Xto, Yto) {
  SnippingTool()
  Sleep 1000
  HoldDownMouseAndMove(Xfrom, Yfrom, Xto, Yto)
}

;; for testing each of the below:
; ^+F1:: SnipItQuad_5minPrices() ; to test
SnipItQuad_PriceLast3d() {
  ; NB use window spy 'screen' values here!
  SnipItQuad(2840, 1200, 3430, 1645)
}
SnipItQuad_7dSummary() {
  ; NB use window spy 'screen' values here!
  SnipItQuad(1569, 600, 3730, 2090)
}
SnipItQuad_5minPrices() {
  ; NB use window spy 'screen' values here!
  SnipItQuad(1580, 1502, 3730, 2090)
}
SnipItQuad_PortfolioOffers() {
  ; NB use window spy 'screen' values here!
  SnipItQuad(987, 736, 3639, 2090)
}

SearchInConfluence(string) { ;to get to correct place to paste in. Also hits down arrow, so it actually shd select the line immediately below the string
  Send ^f
  Sleep 500
  Send, %string%
  Sleep 500
  Send {Enter}
  Send {Escape}
  Send {Down}
}
; ^+F2:: SearchInConfluence("3. Portfolio Offers") ; to test
CopyFromStacksAndPasteInConfluence(figureNumber) {
  ; figure number is numeric, 0-3 at mo
  
  WinActivate, bid stacks
  Sleep 1000
  
  ; click relevant tab on stacks report
  If (figureNumber = 0) {
    MouseClick, Left, 667, 603 ; daily price tab
  }
  If (figureNumber = 1) {
    MouseClick, Left, 667, 690 ; bids tab
  }
  If (figureNumber = 2) {
    MouseClick, Left, 667, 690 ; bids tab
  }
  If (figureNumber = 3) {
    MouseClick, Left, 667, 790 ; portfolio offers tab
  }
  Sleep 1000

; snip
If (figureNumber = 0) {
    SnipItQuad_PriceLast3d()
  }
  If (figureNumber = 1) {
    SnipItQuad_7dSummary()
  }
  If (figureNumber = 2) {
    HoldDownMouseAndMove(1077, 647, 1560, 630) ; convert to yday-only prices
    Sleep 1000
    MouseClick, Left, 1077, 1250
    Sleep 1000
    SnipItQuad_5minPrices()
  }
  If (figureNumber = 3) {
    SnipItQuad_PortfolioOffers()
  }
  Sleep 1000

  MouseClick, Left, 3500, 1800 ; click on snip when it comes up in notificationsd - annoying hack cos wasn't pasting properly
  Sleep 2000

  ; paste in confluence
  WinActivate, Edit - 
  Sleep 1000
  If (figureNumber = 0) {
    Send ^{Home}
  }
  If (figureNumber = 1) {
    SearchInConfluence("1. Daily view")
  }
  If (figureNumber = 2) {
    SearchInConfluence("2. 5min prices")
  }
  If (figureNumber = 3) {
    SearchInConfluence("3. Portfolio Offers")
  }
  Sleep 1000
  Send ^v

  WinActivate, Snipping Tool ; sometimes need to do "copy switch and paste" from this window when above paste doesnt work

  SoundBeep
}

; AND WITHOUT FURTHER ADO...  "populate visuals" in dragon
;  *** should have conf page open in Edit mode (win title starting 'Edit - ' ***
;  *** and bid stacks dash also open ***
^+F5:: CopyFromStacksAndPasteInConfluence(0)
^+F1:: CopyFromStacksAndPasteInConfluence(1)
^+F2:: CopyFromStacksAndPasteInConfluence(2)
^+F3:: CopyFromStacksAndPasteInConfluence(3)

; another convenience functiuon - 'top of stacks report' in DNS
^+F4:: 
  SearchInConfluence("Summary")
Return

;-------------------------
; clicking windows and taskbar – "win one", "Win two" etc.
;-------------------------

; Function to click win 1, win 2 etc. calculates the different window coords by multiplying the window number by the panel height (height in taskbar - can calc using window spy) so easy to update for different screen sizes.
ClickTaskbarWin_vert(win_num, xcoord, ycoord_win1, panel_height) ;tb on side of screen
{
  ycoord := ycoord_win1 + panel_height*(win_num-1)
  CoordMode, Mouse, Screen
  MouseClick, left, xcoord, ycoord
  CoordMode, Mouse, Window 
}
; Function to click win 1, win 2 etc. calculates the different window coords by multiplying the window number by the panel height (height in taskbar - can calc using window spy) so easy to update for different screen sizes.
ClickTaskbarWin_hor(win_num, ycoord, xcoord_win1, panel_width) ;at top/botty
{
  xcoord := xcoord_win1 + panel_width*(win_num-1)
  CoordMode, Mouse, Screen
  MouseClick, left, xcoord, ycoord
  CoordMode, Mouse, Window 
}
^!F1::ClickTaskbarWin_hor(1, 50, 450, 410) ;_hor or _vert
^!F2::ClickTaskbarWin_hor(2, 50, 450, 410)
^!F3::ClickTaskbarWin_hor(3, 50, 450, 410)
^!F4::ClickTaskbarWin_hor(4, 50, 450, 410)
^!F5::ClickTaskbarWin_hor(5, 50, 450, 410)
^!F6::ClickTaskbarWin_hor(6, 50, 450, 410)
^!F7::ClickTaskbarWin_hor(7, 50, 450, 410)
#!F8::ClickTaskbarWin_hor(8, 50, 450, 410)
;#!F9::ClickTaskbarWin_hor(9)
;+!F9::ClickTaskbarWin_vert(9)


;-------------------------
; moving windows around monitors - replaces shitty win11 commands
;-------------------------

; Commands for current screen:
#Up::WinMaximize, A  ; maximise active - "win maximise" in dragon
#Down::WinRestore, A  ; restore active  - "win restore" in dragon
#Right::
  WinMaximize, A  ; RH of screen -  "win right" in dragon
  Sleep 500
  Send #{right}
  Sleep 500
  Send {Esc}
Return
#Left::
  WinMaximize, A  ; LH of screen -  "win left" in dragon
  Sleep 500
  Send #{left}
  Sleep 500
  Send {Esc}
Return
; centre window -  "win centre" in dragon
; from https://superuser.com/questions/403187/need-autohotkey-to-center-active-window
#!Up::
  WinRestore, A ;jay addition, doesn't work when max'd
  CenterActiveWindow()
Return
CenterActiveWindow()
{
    windowWidth := A_ScreenWidth * 0.5 ; desired width
    windowHeight := A_ScreenHeight * 0.95 ; desired height
    WinGetTitle, windowName, A
    WinMove, %windowName%, , A_ScreenWidth/2-(windowWidth/2), A_ScreenHeight/2-(windowHeight/2), windowWidth, windowHeight
}

; Same commands as above, but move to other screen first.
; dragon commands same but suffix with "other"
; **BASED ON 2 MONITORS - NEED TO UPDATE IF >2 MONITORS"
; note if I just want to move to the other screen without maximising/right half/left path then just do the native Windows command, 
+#Up::  
   WinMaximize, A
   Send {Shift down}{LWin down}{Left}{Shift up}{LWin up} ; will move to left OR right monitor IF ONLY 2 MONITORS
  WinMaximize, A  ; maximise active
Return
+#Right::
  WinMaximize, A  
   Send {Shift down}{LWin down}{Left}{Shift up}{LWin up} ; will move to left OR right monitor IF ONLY 2 MONITORS
  WinMaximize, A  ; RH of screen
  Sleep 1000
  Send #{right}
  Sleep 500
  Send {Esc}
Return
+#Left::  
  WinMaximize, A
   Send {Shift down}{LWin down}{Left}{Shift up}{LWin up} ; will move to left OR right monitor IF ONLY 2 MONITORS
  WinMaximize, A  ; LH of screen
  Sleep 1000
  Send #{left}
  Sleep 500
  Send {Esc}
Return

; maximise across both screens.  -  "win max both" in dragon
; also put in function just for sake of it - copying CenterActiveWindow() above.
; https://stackoverflow.com/questions/9828808/how-can-i-maximize-a-window-across-multiple-monitors
^!#Up::
  MaxBothScreens()
return
MaxBothScreens()
{
  pixelOffsetFromTop := 90 ; for if I have taskbar and dragonbar on top.
  pixelStretchDownFromBottom := 5 ; if I wanna stretch win down slightly below bottom of normal screen.
  WinGetActiveTitle, Title
  WinRestore, %Title%
  SysGet, X1, 76 ; see SysGet documentation https://www.autohotkey.com/docs/commands/SysGet.htm for what numbers mean.
  SysGet, Y1, 77
  SysGet, Width, 78
  SysGet, Height, 1 ; SysGet 1 is equivalent to A_ScreenHeight as used above; see SysGet documentation.
  WinMove, %Title%,, X1, Y1 + pixelOffsetFromTop, Width*1, Height*1 - pixelOffsetFromTop + pixelStretchDownFromBottom ; Width*x, Height*x - in case I want to have not maximised.
}

;-------------------------
; HOTSTRINGS
;-------------------------

; note - the if not statement isn't working! poss cos deprecated... perhaps add all possible 
; programs taht I do want instead with an 'or', instead of those I don't... per example here
; (https://www.autohotkey.com/docs/commands/WinActive.htm)
:*:uu::
if not WinActive("ahk_group Explorer")
    Send {_}
    Send {left 2}
    Send {del}
    Send {right}
Return

:*:`;tj::{enter}{enter}Thanks,{enter}{enter}Jay

:*:`;cf:: ;confluence. "general" and "r" spaces only are relevant to me
 Run C:\Program Files\Google\Chrome\Application\chrome.exe https://haastenergy.atlassian.net/wiki/spaces/R/overview
Return

:*:`;tm:: ;time management via jira
 Run C:\Program Files\Google\Chrome\Application\chrome.exe https://haastenergy.atlassian.net/jira/software/c/projects/GEN/boards/5?assignee=60c674cfa01e11006a39ab36
Return

:*:`;gm:: 
 Run C:\Program Files\Google\Chrome\Application\chrome.exe https://gmail.com
Return

:*:`;gc:: 
 Run C:\Program Files\Google\Chrome\Application\chrome.exe https://calendar.google.com/calendar/u/0/r/week
Return

:*:`;do:: 
 Run C:\Users\Owner\Documents
Return

:*:`;dl:: 
 Run C:\Users\Owner\Downloads
Return

:*:`;ri:: 
 Run C:\Program Files\Google\Chrome\Application\chrome.exe https://ap-southeast-2.console.aws.amazon.com/ec2/v2/home?region=ap-southeast-2#Instances:sort=instanceState ;ec2 running instances
Return

:*:`;s3:: 
 Run C:\Program Files\Google\Chrome\Application\chrome.exe https://s3.console.aws.amazon.com/s3/buckets/nz-omg-ann-aipl-staging/digital_pathway_reporting/?region=ap-southeast-2&tab=overview ;S3
Return

; Opening via hotstrings - other

:*:`;bjm::
  Send B3ingjohnmalkovich
  Send +1
  Send {Enter}
Return

:*:`;jh::jay@haastenergy.com
:*:`;rj::ruffell.jay@gmail.com
:*:`;fp::Spring@21{!}
:*:`;sp::altheg81
:*:`;021::0210528720

/*
;-------------------------
; slack: opening specific channel from anywhere
;-------------------------

; EDIT: not working, cos ctrl + s stops save!
; CTRL + S + {first letter of slack thread name} - need #If 
; method below to use 2 keys with a modifier
#If, GetKeyState("Ctrl") ;start of context sensitive hotkeys
s & j::
  Run C:\Users\Owner\AppData\Local\Microsoft\WindowsApps\Slack.exe
  Sleep 1500
  Send ^k
  Sleep 1500
  Send j     ; for jon
  Sleep 1500
  Send {enter}
return
#If ;end of context sensitive hotkeys

#If, GetKeyState("Ctrl") ;start of context sensitive hotkeys
s & e::
  Run C:\Users\Owner\AppData\Local\Microsoft\WindowsApps\Slack.exe
  Sleep 1500
  Send ^k
  Sleep 1500
  Send e   ; for elbert 
  Sleep 1500
  Send {enter}
return
#If ;end of context sensitive hotkeys

#If, GetKeyState("Ctrl") ;start of context sensitive hotkeys
s & r::
  Run C:\Users\Owner\AppData\Local\Microsoft\WindowsApps\Slack.exe
  Sleep 1500
  Send ^k
  Sleep 1500
  Send r    ;for rob
  Sleep 1500
  Send {enter}
return
#If ;end of context sensitive hotkeys

#If, GetKeyState("Ctrl") ;start of context sensitive hotkeys
s & c::
  Run C:\Users\Owner\AppData\Local\Microsoft\WindowsApps\Slack.exe
  Sleep 1500
  Send ^k
  Sleep 1500
  Send charles jonathan ;charls and jon    
  Sleep 1500
  Send {enter}
return
#If ;end of context sensitive hotkeys

#If, GetKeyState("Ctrl") ;start of context sensitive hotkeys
s & g::
  Run C:\Users\Owner\AppData\Local\Microsoft\WindowsApps\Slack.exe
  Sleep 1500
  Send ^k
  Sleep 1500
  Send g ;general
  Sleep 1500
  Send {enter}
return
#If ;end of context sensitive hotkeys

#If, GetKeyState("Ctrl") ;start of context sensitive hotkeys
s & d::
  Run C:\Users\Owner\AppData\Local\Microsoft\WindowsApps\Slack.exe
  Sleep 1500
  Send ^k
  Sleep 1500
  Send d ;dev
  Sleep 1500
  Send {enter}
return
#If ;end of context sensitive hotkeys

;^sj:: 
;  Run C:\Users\Owner\AppData\Local\Microsoft\WindowsApps\Slack.exe
;  Sleep 200
;  Send ^k
;  Sleep 200
;  Send j
;Return
*/

;#left::
;WinGet, mm, MinMax, A
;WinRestore, A
;WinGetPos, X, Y,,,A
;WinMove, A,, X-A_ScreenWidth, Y
;if(mm = 1){
;    WinMaximize, A
;}
;return
;					;moves window from right to left monitor
;#right::
;WinGet, mm, MinMax, A
;WinRestore, A
;WinGetPos, X, Y,,,A
;WinMove, A,, X+A_ScreenWidth, Y
;if(mm = 1){
;    WinMaximize, A
;}
;return
					
;-------------------------
; MODIFIER KEYS
;-------------------------

; Because Dragon is a pain with the Windows key, add a quick letter and backspace to stop Dragon opening. Not working - causes win modifier key to fail
;LWin::
; Send {LWin}
; Sleep 500
; Send a
; Sleep 500
; Send {backspace}
;Return

;; remap winkey + xxx keys to right control (keyboard winkey not working)  
;>^up::  
; Send #{up}
;Return

;>^down::
; Send #{down}
;Return

;>^right::
; Send #{right}
;Return

;>^left::
; Send #{left}
;Return					

; #IfWinActive
					
;=====================================================================================================
;=====================================================================================================

;WORD

#IfWinActive ahk_class OpusApp ; applies only to Word from now on

^!p::
Send ^a
Send ^x
WinActivate,RStudio
Send ^v
Return				; cut and paste to TinnR/RStudio 



;#### temporary bookmarking:

:*:`;mh::				; "Mark Here". "*" means no ending character required to activate
 Send {home}				
 Send `===== TEMPORARY BOOKMARK =====	
Return					

:*:`;gom::				; "Go Mark". "*" means no ending character required to activate
 Send ^f				
 Sleep 200
 Send `===== TEMPORARY BOOKMARK =====	
 Send {enter}
 Sleep 200
 Send {escape}
Return		

#IfWinActive

;=====================================================================================================
;===================================================================================================

;WINDOWS EXPLORER

#IfWinActive ahk_class ExploreWClass 

F6:: Send !d

#IfWinActive

#IfWinActive ahk_class CabinetWClass               ;2 possible window names

F6:: Send !d

#IfWinActive


;=====================================================================================================
;=====================================================================================================

;GIT BASH

#IfWinActive ahk_class mintty 

:*:`;gs::git status {enter}
:*:`;gda::git diff {enter} ;'git diff all'
:*:`;gds::git diff{space} ;'git diff specific files'
:*:`;ga::git add{space}
:*:`;gp::git push origin master
:*:`;gic::
  Send git commit -m "" 
  Send {left}
Return

#IfWinActive


;=====================================================================================================
;=====================================================================================================
;===================================================================================================

;OUTLOOK

#IfWinActive Outlook 

#^!F1::ClickTaskbarWin_vert(1, 500, 740, 210) ;_hor or _vert
#^!F2::ClickTaskbarWin_vert(2, 500, 740, 210)
#^!F3::ClickTaskbarWin_vert(3, 500, 740, 210)
#^!F4::ClickTaskbarWin_vert(4, 500, 740, 210) 
#^!F5::ClickTaskbarWin_vert(5, 500, 740, 210)
#^!F6::ClickTaskbarWin_vert(6, 500, 740, 210)
#^!F7::ClickTaskbarWin_vert(7, 500, 740, 210)
#^!F8::ClickTaskbarWin_vert(8, 500, 740, 210)

:*:`;hg::
 Send Hey guys,
 Send {enter}{enter}			
Return

:*:`;hc::
 Send Hey Charles,
 Send {enter}{enter}			
Return

:*:`;tj::
 Send {enter}{enter}
 Send Thanks,
 Send {enter}{enter}
 Send Jay				
Return

#IfWinActive


;=====================================================================================================;=====================================================================================================
;===================================================================================================

;OUTLOOK - new message window

#IfWinActive Message 

:*:`;hg::
 Send Hey guys,
 Send {enter}{enter}			
Return

:*:`;hs::
 Send Hey Sophia,
 Send {enter}{enter}			
Return

:*:`;tj::
 Send {enter}{enter}
 Send Thanks,
 Send {enter}{enter}
 Send Jay				
Return

#IfWinActive

;=====================================================================================================
;=====================================================================================================
					
;Spyder
#IfWinActive ahk_exe pythonw.exe
::hh::               ;hashkey followed by space.
 Send +3
 Send {space}
Return

#IfWinActive

;=====================================================================================================
;=====================================================================================================

#IfWinActive ahk_exe ms-teams.exe
#^!F1::ClickTaskbarWin_vert(1, 500, 740, 110) ;_hor or _vert
#^!F2::ClickTaskbarWin_vert(2, 500, 740, 110)
#^!F3::ClickTaskbarWin_vert(3, 500, 740, 110)
#^!F4::ClickTaskbarWin_vert(4, 500, 740, 110) 
#^!F5::ClickTaskbarWin_vert(5, 500, 740, 110)
#^!F6::ClickTaskbarWin_vert(6, 500, 740, 110)
#^!F7::ClickTaskbarWin_vert(7, 500, 740, 110)
#^!F8::ClickTaskbarWin_vert(8, 500, 740, 110)
#IfWinActive

;=====================================================================================================
;=====================================================================================================

#IfWinActive Visual Studio Code			; commands largely just copied from Rstudio code

; toggle through open tabs - not working!
F1::
  Send {Ctrl down}{Tab}
  ; Send {Tab down}{Tab up}
  Sleep 5000
  Send {Ctrl up}
  ; MsgBox "Ctrl key now up", T1 ; T should time out, but not working.
  
  ; toggle cursor to indicate that control is now up:
  Send {Down}
  Sleep, 100
  Send {Up}
  Sleep, 100
  Send {Down}
  Sleep, 100
  Send {Up}
Return


;----------------
;main hotstrings - more below
;----------------

:*:`;si::sessionInfo()

:*:`;sh::stop("here")

:*:ee::               ;"end enter" - go to end of line then press enter, to start a new line. need space first, which then gets deleted
 Send {backspace} 
 Send {end}
 Send {enter}
Return

:*:ll::                ;"line end" - go to end of line. need space first, which then gets deleted. Just easier than pressing {end} key
 Send {backspace} 
 Send {end}
Return

;:*:ss::               ;"home key" - go to start of line. need space first, ;which then gets deleted. Just easier than pressing {home} key
; Send {backspace}
; Send {home}
;Return

:*:`;tb:: 
 Send traceback() 
 Send {enter}
Return

:*:`;b::               
 Send browser()
Return

:*:`;pr::           
 Send print()
 Send {left}
Return

:*:`;rh::          
 Send print(head())
 Send {left}
 Send {left}
Return

::`;s::               
 Send str()
 Send {left}
Return

:*:`;h::               
 Send head()
 Send {left}
Return

:*:`;t::               
 Send tail()
 Send {left}
Return

::hh::               ;hashkey followed by space.
 Send +3
 Send {space}
Return

:*:dd::               ;dollarsign. need space first, which then gets deleted
 Send {backspace}
 Send +4
Return

:*:`;fu::             ;'full  underline'
 Send +3
 Send ____________________________________
 Send {enter}
 Send {enter}
Return

:*:`;el::             ;'equal line'
 Send +3
 Send {space}{*}{*}{*}{*}{*}{*}{*}{*}{*}{*}{*}{*}{*}{*}{*}{*}{*}{*}{*}{*}
 Send {enter}
Return

:*:,,::`<- `		;hit comma twice to assign variable

;#### temporary bookmarking:

:*:`;mh::				; "Mark Here". "*" means no ending character required to activate
 Send {home}				
 Send `===== TEMPORARY BOOKMARK =====	
 Send {enter}
Return					

:*:`;gom::				; "Go Mark". "*" means no ending character required to activate
 Send ^f				
 Sleep 200
 Send `===== TEMPORARY BOOKMARK =====	
 Send {enter}
 Sleep 200
 Send {escape}
Return		

;-------------------------
; navigation
;-------------------------

; console focus - hit space first, which then gets deleted
:*:`;cc:: 
 Send {backspace}
 Send ^k
 Sleep 100
 Send {2}
Return		

; editor focus - hit space first, which then gets deleted
:*:`;ef:: 
 Send {backspace}
 Send ^k
 Sleep 100
 Send {1}
Return		

; kill terminal - hit space first, which then gets deleted
:*:`;kk:: 
 Send {backspace}
 Send ^k
 Sleep 100
 Send {2}
 Sleep 100
 Send ^!k
Return		

;-------------------------
; SQL
;-------------------------

:*:`;sa:: 
 Send SELECT * FROM
 Send {space}
Return
:*:`;stt:: 
 Send SELECT TOP 10 * FROM
 Send {space}
Return

;-------------------------
; Dplyr commands
;-------------------------

::nl::             ; 'next line' - go to end of line, create '%>%' and press enter. Need to hit space first
 Send {backspace}
 Send {end}
 Send {space}
 Send +5
 Send +.
 Send +5
 Send {enter}
Return

::sel::               
 Send select()
 Send {left}
Return
::gb::
 Send group_by()
 Send {left}
Return
::mut::               
 Send mutate()
 Send {left}
Return
::tra::               
 Send transmute()
 Send {left}
Return
::sz::
 Send summarise()
 Send {left}
Return
::fil::               
 Send filter()
 Send {left}
Return
::lj::               
 Send left_join()
 Send {left}
Return
::tl::               
 Send tally()
 Send {left}
Return

::oo::               ; pipeline operator
 Send +5
 Send +.
 Send +5
 Send {enter}
Return

::ooo::               ; as above but no enter at end - useful when inserting a line within an existing pipe.
 Send +5
 Send +.
 Send +5
 Send {down}
 Send {home}
Return

:*:`;dd:: ; dd %>%
 Send dd{space}
 Send +5
 Send +.
 Send +5
 Send {enter}
Return

;-------------------------
; ggplot commands
;-------------------------

::gps::             ; 'plus sign' - go to end of line, create '+' and press enter. Need to hit space first
 Send {backspace}
 Send {end}
 Send {space}
 Send {+}
 Send {enter}
Return

::gpsne::             ; 'plus sign no enter' - go to end of line, create '+' but no pressing enter. Need to hit space first
 Send {backspace}
 Send {end}
 Send {space}
 Send {+}
 Send {down}
 Send {home}
Return

::gg::               ; cursor finishing in brackets
 Send ggplot(aes())
 Send {left 2}
Return
::gp::               
 Send geom_point()
 Send {left}
Return
::gl::               
 Send geom_line()
 Send {left}
Return
::gs::               
 Send geom_smooth()
 Send {left}
Return
::fw::               
 Send facet_wrap(~, scales = 'free')
 Send {left 18}
Return
::ggs::               
 Send ggsave('.png')
 Send {left 6}
Return

;-------------------------
; Other Rstudio commands
;-------------------------

RAlt:: 					; send lines with right alt key
  Send ^{enter}
Return

^e::				;cancel operation (switch to code window, press escape, and switch back to source)
 Send ^2
 Send {escape}
 Send ^1
Return

#IfWinActive

;=====================================================================================================
;=====================================================================================================

;=====================================================================================================
;=====================================================================================================
;=====================================================================================================
;=====================================================================================================
