; GLOBAL COMMANDS NEEDED TO LOAD FIRST::

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode 2   ;sets window title matching (for #IfWinActive) to partial matches

;=====================================================================================================
;=====================================================================================================

;TEMP while windows key aint working.
F8:: Send {LWin}
F9:: Send #d

;OPENING DOCUMENTS AND FOLDERS

; F1:: Run \\100.97.137.78\nz_akl_omd\OMD\Annalect\Jay general methods & code
; F2:: Run \\100.97.137.78\nz_akl_omd\OMD\Annalect\Jay other documents
F3:: Run \\100.97.137.78\nz_akl_omd\OMD\Annalect\Clients\OMD

; open outlook if its not open already, otherwise just switch to open window
F7:: 
if WinExist("Outlook")
    WinActivate ; use the window found above
else
    Run C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Outlook.lnk
Return

; Repeat for teams
!F7:: 
if WinExist("Teams")
    WinActivate ; use the window found above
else
  Run C:\Users\jay.ruffell\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Teams.lnk
Return

; Opening via hotstrings - URLs

:*:`;ct:: ; Close Teams - quikly switch to and then back again to current app if annoying popup shows up
if WinExist("Teams")
  WinActivate
  Sleep, 2000
  Send !{tab}
Return

;------------------------------------------------------------------------
; DICTATION: if word or outlook, use the in-app dictation shortcuts. otherwise open up a word doc especially for dication to copy across
F1:: Send !e{2}{d}{d} ; outlook: message not popped out
F2:: Send !h{d}{d} ; outlook: message popped out, or word
:*:`;dic:: 
    Run C:\Users\jay.ruffell\OneDrive - OneWorkplace\Desktop\dictation.docx  
    WinWaitActive, dictation
    Send !h{d}{d}
Return ; dictation scratchpad
 
; original code, not working now for unkonwn reasons. 'if' statements aren't running for outlook windows
; :*:`;dic:: 
;  if WinActive("Outlook")
;    Send !e{2}{d}{d}
;  if WinActive("Message")
;    Send !h{d}{d} 
;  if WinActive("Word")
;    Send !h{d}{d}
;  else {
;    Run C:\Users\jay.ruffell\OneDrive - OneWorkplace\Desktop\dictation.docx  
;    WinWaitActive, dictation
;    Send !h{d}{d}
;  }
; Return
; for using scratchpad, select dictated text and send back to orig app
:*:`;sb:: 
  Send ^a
  Sleep 200
  Send ^x
  Sleep 200
  Send !{tab}
  Sleep 1000
  Send ^v
Return
;------------------------------------------------------------------------
:*:`;ftc:: ; forticlient VPN login - opens forticlient then logs in
Run C:\Program Files\Fortinet\FortiClient\FortiClient.exe
    Sleep 6000 
    Send {tab}{tab}
    Sleep 1000
    Send jay.ruffell
    Send {tab}
    Sleep 1000
    Send Spring{@}21{!}
    Send {enter}
    Sleep 2000
    Send !{F4}
Return

:*:`;adh:: ; ads data hub
 Run C:\Program Files\Google\Chrome\Application\chrome.exe https://adsdatahub.google.com/#/queries
Return

:*:`;bq:: ; ads data hub
 Run C:\Program Files\Google\Chrome\Application\chrome.exe https://console.cloud.google.com/bigquery?_ga=2.62333803.698029576.1586206027-1272128723.1544390697&_gac=1.175792662.1586221277.EAIaIQobChMIk76e1Y7V6AIVlouPCh11xwZ5EAAYASAAEgLTDfD_BwE&project=omg-newzealand-adsdatahub&supportedpurview=project&p=omg-newzealand-adsdatahub&d=twg_nz_the_warehouse_adh&page=dataset
Return


:*:`;sbd:: ; When running query in ADH, in table name field after naming table and then putting a space after, Specify sandbox data with '_sandbox' suffix of table name, then set start and end date as full range available, then run. 
 Send {backspace}
 Send _sandbox
 Sleep 500
 Send {tab}
 Sleep 500
 Send Aug 17, 2018
 Sleep 500
 Send {tab}
 Sleep 500
 Send {tab}
 Sleep 500
 Send Sep 18, 2018
 Sleep 500
 Send {tab}
 Sleep 500
 Send {tab}
 Sleep 500
 Send {tab}
 Sleep 500
 Send {tab}
 Sleep 500
 Send {enter}
Return

:*:`;m2:: ; When running query in ADH, in table name field after naming table and then putting a space after, Specify daternage with '_may20' suffix of table name, then set start and end date as this range then run. EDIT: no longer adding suffix.
 ;Send {backspace}
 ;Send _may20
 Sleep 500
 Send {tab}
 Sleep 500
 Send May 1, 2020
 Sleep 500
 Send {tab}
 Sleep 500
 Send {tab}
 Sleep 500
 Send May 31, 2020
 Sleep 500
 Send {tab}
 Sleep 500
 Send {tab}
 Sleep 500
 Send {tab}
 Sleep 500
 Send {tab}
 Sleep 500
 Send {enter}
Return

:*:`;tm:: ; task manager 
 Run C:\Program Files\Google\Chrome\Application\chrome.exe https://app.asana.com/0/409477042493367/list
Return

:*:`;uec:: 
 Run C:\Program Files\Google\Chrome\Application\chrome.exe https://170415891119.signin.aws.amazon.com/console ;AWS console
Return

:*:`;ri:: 
 Run C:\Program Files\Google\Chrome\Application\chrome.exe https://ap-southeast-2.console.aws.amazon.com/ec2/v2/home?region=ap-southeast-2#Instances:sort=instanceState ;ec2 running instances
Return

:*:`;s3:: 
 Run C:\Program Files\Google\Chrome\Application\chrome.exe https://s3.console.aws.amazon.com/s3/buckets/nz-omg-ann-aipl-staging/digital_pathway_reporting/?region=ap-southeast-2&tab=overview ;S3
Return

:*:`;tt:: 
 Run C:\Users\jay.ruffell\OneDrive - OneWorkplace\Desktop\TODAYS TASKS.docx
Return

; Opening via hotstrings - other

:*:`;cmi:: 
 Run \\100.97.137.78\nz_akl_omd\OMD\Annalect\Jay general methods & code\annalect-data-science-cmi-dashboard
Return

:*:`;op:: 
  Run \\100.97.137.78\nz_akl_omd\OMD\Annalect\Clients\OMD\Open Polytechnic\Jay analyses
Return

:*:`;sky:: 
 ;Run O:\OMD\Annalect\Clients\OMD\SKY\Jay analyses\ 
 Run \\100.97.137.78\nz_akl_omd\OMD\Annalect\Clients\OMD\SKY\Jay analyses\ 
Return

:*:`;aai:: 
 ;Run O:\OMD\Annalect\Clients\OMD\AA Insurance\Digi Attribution
 Run \\100.97.137.78\nz_akl_omd\OMD\Annalect\Clients\OMD\AA Insurance\Digi Attribution
Return

:*:`;twg:: 
 ;Run O:\OMD\Annalect\Clients\OMD\TWG\Jay\Digi pathway
 Run \\100.97.137.78\nz_akl_omd\OMD\Annalect\Clients\OMD\TWG\Jay\Digi pathway
Return

:*:`;gem:: 
 ;Run \\100.97.137.78\nz_akl_omd\OMD\Annalect\Clients\Total Media\GE Money
 Run \\100.97.137.78\nz_akl_omd\OMD\Annalect\Clients\Total Media\GE Money\Digital attribution 2020
Return

:*:`;sae:: ;save as excel - for csv files
 Send {alt} 
 Sleep 1000
 Send {f}
 Sleep 1000
 Send {alt} 
 Sleep 1000
 Send {a}
 Sleep 1000
 Send {y}{3}
 Sleep 1000
 Send {home}
 Sleep 1000
 Send {tab}
 Send {enter}
Return

:*:`;sky:: 
 Run \\100.97.137.78\nz_akl_omd\OMD\Annalect\Clients\OMD\SKY\Jay analyses\   ; sky folder
Return

:*:`;iam::170415891119 ; AWS IAM number
:*:`;ep::5tatsMan{!}          ; ec2 pw - cloud R machine o(9otvN!)g
:*:`;sp::5gGeB{;}C7nhFn)8o7RPpDGtsHlcYwolq51          ; ec2 pw  'subnet old' machine

F12:: Run C:\Program Files\Google\Chrome\Application\chrome.exe


;=====================================================================================================
;=====================================================================================================

;HOTKEYS TRIGGERED BY LONGPRESS OF INDIVIDUAL KEYS

;see http://www.autohotkey.com/board/topic/80697-long-keypress-hotkeys-wo-modifiers/page-1 for full discussion, and some good e.g.s

;1. Copy all by longpress of 'a':

;$a::
;SetKeyDelay, -1
;KeyWait, a
;return
; $a Up::
; If (A_PriorHotKey = "$a" AND A_TimeSincePriorHotkey < 750)
;   Send a
; else
;   SendInput, {Ctrldown}a{Ctrlup} ; select all
; return

;2. {end} by longpress of ''':

;$'::
;KeyWait, '
;return
; $' Up::
; If (A_PriorHotKey = "$'" AND A_TimeSincePriorHotkey < 750)
;   Send `'
; else
;   SendInput, {End} 
; return

;3. {home} by longpress of ';':

;$`;::
;KeyWait, `;
;return
; $`; Up::
; If (A_PriorHotKey = "$`;" AND A_TimeSincePriorHotkey < 750):
;   Send `;
; else
;   SendInput, {Home} 
; return

;=====================================================================================================
;=====================================================================================================

;MISC

:*:`;cc:: 
MouseMove 900,500
Sleep, 500
Click
Return					;click in centre of screen

:*:`;sl:: 
 Send {home}
 Send +{end} 
Return	          ;select line

:*:`;dl:: 
 Send {home}
 Send +{end} 
 Send {delete}
 Send {backspace}
Return	          ;delete line

:*:`;cl:: 
 Send {home}
 Send +{end} 
 Send ^c
Return	          ;copy line

:*:`;xl:: 
 Send {home}
 Send +{end} 
 Send ^x
Return	          ;select line

:*:`;ah:: 
MouseMove 3050,2300
Sleep, 500
Click right
Return					;bring up autohotkey menu from taskbar

:*:`;st:: 
 Run C:\Program Files\FreeCountdownTimer\FreeCountdownTimer.exe
 Sleep 200

RAlt & down::
 Send {Ctrl Down}
 MouseClick,WheelDown,,,3,0,D,R   ;zoom in and out as per ctrl+mousescrolling (used in MS Word, for example)
 Send {Ctrl Up}
Return

RAlt & up::
 Send {Ctrl Down}
 MouseClick,WheelUp,,,3,0,D,R
 Send {Ctrl Up}
Return

RAlt & right:: 
 Send ^#{tab} 	;opens Aero 'flip 3d' to flip through open windows
Return

;###some keys for making keyboard naviagtion easier:----------------------------------------------

;Capslock::
;   Send {LCtrl Down}
;   KeyWait, Capslock ; Maps caps to control. keywait waits until the Capslock button is released. BUT DOESNT WORK WHEN CAPSLOCK ARROW NAVIGATION (BELOW) IS ON.
;   Send, {LCtrl Up}
;Return

;##Capslock plus middle row of LH keys: move up down left right by letter or word:
;Capslock & k:: Send {up}	
;Capslock & l:: Send {down}
;Capslock & j:: Send {left}
;Capslock & `;:: Send {right}
;Capslock & h:: Send ^{left}
;Capslock & ':: Send ^{right} 

;##Capslock plus bottom row of LH keys: highlight left right up down by letter or word:
;Capslock & n:: Send +{home}		
;Capslock & RShift:: Send +{end}								
;Capslock & m:: Send ^+{left}
;Capslock & /:: Send ^+{right}
;Capslock & ,:: Send +{up}
;Capslock & .:: Send +{down}

;##Capslock plus top row of LH keys: delete/backspace by letter or word:
;Capslock & i:: Send {backspace}	
;Capslock & o:: Send {delete}
;Capslock & u:: Send ^{backspace}
;Capslock & p:: Send ^{delete}
;Capslock & y:: 
;  Send +{home} 
;  Send {delete}
; Return
;Capslock & [::
;  Send +{end} 
;  Send {delete}
; Return

;##Capslock plus middle row of LH keys: move up down left right by letter or word:
Capslock & k:: Send {down}	
Capslock & l:: Send {right}
Capslock & j:: Send {left}
Capslock & `;:: Send ^{right}
Capslock & h:: Send ^{left}
Capslock & ':: Send {end}  ;key for 'up' is in top row of keys below

;##Capslock plus bottom row of LH keys: highlight left right up down by letter or word:
Capslock & n:: Send +^{left}		
Capslock & RShift:: Send +{end}								
Capslock & m:: Send +{left}
Capslock & /:: Send ^+{right}
;Capslock & ,:: Send +{up}
Capslock & .:: Send +{right}

;##Capslock plus top row of LH keys: delete/backspace by letter or word:
Capslock & i:: Send {up}	
Capslock & o:: Send {delete}
Capslock & u:: Send {backspace}
Capslock & p:: Send ^{delete}
Capslock & y:: Send ^{backspace}
Capslock & [::
  Send +{end} 
  Send {delete}
 Return

;-----------------------------------------------------

;`:: Send !{tab}		;switch windows

;:*:kl:: 
;    Send {PgDn}		;using hotkeys for page down and page up. the asterisk means no ending (space or enter, for example) is required.
;Return

;:*:lk:: 
;    Send {PgUp}			
;Return


;###hotstrings: 

:*:`;rj::ruffell.jay@gmail.com
:*:`;fp::Spring@21{!}
:*:`;sp::altheg81
:*:`;454::4548601499399117
:*:`;5j::5 Janet Place
:*:`;021::0210528720

;:*?:  ::. `				; makes a full stop space after pressing space twice. BUT causing "PU" command to cause problems in IE11 for some reason.

^+!f::
WinActivate, Internet Explorer
Return 					;make chrome active window

;F2:: 
;Send {LWin}
;Sleep, 300
;Send {up}
;Sleep, 300
;Send {up}
;Sleep, 300
;Send {enter}          	;put comp to sleep 
;Sleep, 300
;Send {down}
;Sleep, 300
;Send {enter}
;Return 

#o:: Send E:\ArcGIS\Default1.gdb\	;output location for arcgis tools. Cant go inside ArcGIS code cos active 							;..window is tool name, not arcGIS name

^`::
Send {NumpadMult}			;switch to normal mode in dragon
Sleep, 100
Send m
Sleep, 100
Send n
Return

>!c:: 
WinActivate, View			; switches to free clipboard and clears clipboard
Click right 70,220
Send l
Send !{tab}
MouseMove 70,10
Return

^+!#m:: 
WinActivate, View  
MouseMove 70,70
Sleep, 300
Click
MouseMove 70,10
Return

^+!#n:: 
WinActivate, View
MouseMove 70,85
Sleep, 300
Click
MouseMove 70,10
Return

^+!#b:: 
WinActivate, View
MouseMove 70,100			; inserts free clipboard item 1-6
Sleep, 300
Click
MouseMove 70,10
Return

^+!#v:: 
WinActivate, View
MouseMove 70,115
Sleep, 300
Click
MouseMove 70,10
Return

^+!#c:: 
WinActivate, View	
MouseMove 70,130
Sleep, 300
Click
MouseMove 70,10
Return

^+!#x:: 
WinActivate, View
MouseMove 70,145
Sleep, 300
Click
MouseMove 70,10
Return


;right click at caret (cursor) plus optional offsets - using for navigating from contents page in Microsoft Word (see Vocola code):
!^+h::
xoff=0
yoff=0
MouseGetPos,X,Y
MouseMove,(A_CaretX-xoff),(A_CaretY-yoff)
click right
;MouseMove,(X),(Y)   returns to orig pos... Don't need
Return

f5::
Send !{f4}				; close
Return

::5 j:: 5 Janet Place, Laingholm

#p::
Send {Lwin}
Sleep, 200 
Send {up} 
Send 	{enter}
Return					; program files

#r::					; my recent documents
Send #r
Sleep, 200
Send shell:recent
Sleep, 200
Send {enter}
Return

#left::
WinGet, mm, MinMax, A
WinRestore, A
WinGetPos, X, Y,,,A
WinMove, A,, X-A_ScreenWidth, Y
if(mm = 1){
    WinMaximize, A
}
return
					;moves window from right to left monitor

#right::
WinGet, mm, MinMax, A
WinRestore, A
WinGetPos, X, Y,,,A
WinMove, A,, X+A_ScreenWidth, Y
if(mm = 1){
    WinMaximize, A
}
return
					

^+!e:: Send ruffell.jay@gmail.com             ;email address

;=====================================================================================================
;=====================================================================================================


;FOXIT

#IfWinActive ahk_class classFoxitReader					; Makes fullscreen and set to highlight for FOXIT ONLY. all commands bewteen the two IfWinActive commands only apply to foxit.
f11::
Send {f11}
Send !ch
Send !vz {enter}200{enter}
Return
#IfWinActive							

;=====================================================================================================
;=====================================================================================================


;PDF-XCHANGE

#IfWinActive XChange

F10:: 
 Send {F11} 		;toggle menu bar
Return

F11:: 
 Send {F12}		;toggle full screen
Return

#IfWinActive							

;=====================================================================================================
;=====================================================================================================


;TINN R/R

^l::
WinActivate,R Console
WinActivate,Tinn
Return 					;make Tinn R active window and R 'next active' 


#IfWinActive Tinn			; applies only to TinnR from now on

^y::
Send ^+z
Return

^p:: WinActivate,Word			; make Microsoft Word the active window
^b:: WinActivate, R Graphics		; make Graphics the active window

^m:: 
Send {home}+{end}
Send ^c
WinActivate,R Console
Send ^v
Send {enter}  
Sleep, 500
WinActivate,Tinn
Send {home}{down}
Return					;send line to R:

^k:: 
Send ^c
WinActivate,R Console
Send ^v
Send {enter}  
Sleep, 500
WinActivate,Tinn
Send {end}{home}
Return					;send selection to R:

^#m:: 
Send {home}+{end}
Send ^c
WinActivate,R Console
Send ^v
Send {enter}  
Sleep, 500
WinActivate,Tinn
Send {home}{down}
WinActivate,R Graphics
Return					;send line and make graph window active:

^#k:: 
Send ^c
WinActivate,R Console
Send ^v
Send {enter}  
Sleep, 500
WinActivate,Tinn
Send {end}{home}
WinActivate,R Graphics
Return					;send seln and make graph window active:


^n:: Send {Home}+{End}			;highlight line:
					

^!c:: Send !r {enter}{down}{enter}       ;connect to R

^!m::
WinMaximize, A			
Send !{tab}
WinGet, mm, MinMax, A
WinRestore, A
WinGetPos, X, Y,,,A
WinMove, A,, X-A_ScreenWidth, Y
if(mm = 1){
    WinMaximize, A
}
WinMaximize, A
Send !{tab}
Return   				; Open Tinn and R on diff monitors and maximise

::b1..::Bird10$
::b3..::Bird30$
::ne..::NAT.EX
::nb..::NAT.BROAD
::neb..::NAT.EX.BROAD
::nbt..::NAT.BROAD.TEA
::nebt..::NAT.EX.BROAD.TEA
::nf..::NAT.FOREST
::sn..::SP.NAT
::st..::SP.TOT
::si..::SP.INT				;variable names


#IfWinActive
					
;=====================================================================================================
;=====================================================================================================

;RStudio

#IfWinActive RStudio			; applies only to RStudio from now on


^w::
 Send ^{w}				;close tab
Return

RAlt:: 					; send lines with right alt key
  Send ^{enter}
Return

^tab:: 				;switch tabs
  Send ^!{down}
Return

^e::				;cancel operation (switch to code window, press escape, and switch back to source)
 Send ^2
 Send {escape}
 Send ^1
Return

^b:: WinActivate, R Graphics		; make Graphics the active window

^g::								; use 'windows()' function to call graphics window (after closing existing graph windows), then run code, then view the plot in this window
  Send ^2
  Send windows()
  Send {enter}
  WinWaitActive, R Graphics
  ;Send !{tab}
  WinActivate, RStudio
  WinWaitActive, RStudio
  Send ^1
  Send ^{enter}
  Sleep 200
;  Send !{tab}
  WinActivate, R Graphics
Return

;^p:: WinActivate,Word			; make Microsoft Word the active window

^n:: Send {Home}+{End}			;highlight line:

; ~Enter::
; If (A_PriorHotKey = A_ThisHotKey and A_TimeSincePriorHotkey < 500)
; {
;  Send ^{enter} ; 							double enter sends code, BUT also sends an enter...  so not useful at present.
; }
; Return

;Mouseclick scripts:

 ~LButton::
 If (A_PriorHotKey = A_ThisHotKey and A_TimeSincePriorHotkey < 500)
 {
  Send ^{left} 									;double click selects word
  Send ^+{right}
 }
 Return

 ~RButton::
 If (A_PriorHotKey = A_ThisHotKey and A_TimeSincePriorHotkey < 500)
 {
  Sleep 50 ; wait for right-click menu, fine tune for your PC		;double right click does 'paste click'
  Send {Esc} ; close it
  Send ^v ; or your double-right-click action here
 }
 Return


;HOTSTRINGS:

::`;sd:: 
 Send spark_disconnect_all()
Return

::`;t:: 
 Send traceback() 
 Send {enter}
Return

::`;js::               ;'jay source' - sources doc to current line by inserting a stop() at the line, then sourcing.
 Send {end}
 Send {enter}
 Send stop('Sourcing to this line, then stopping') 
 Sleep 500
 Send ^+s
 Send +{home}
Return

::`;rp::            
 Run Notepad
 WinActivate, Notepad
 Sleep 200
 Send ^v
 Sleep 200
 Send !e
 Sleep 200
 Send r
 Sleep 200
 Send \
 Sleep 200
 Send {tab}
 Sleep 200
 Send /
 Sleep 200
 Send !a
 Sleep 200
 Send {escape}
 Sleep 200
 Send ^a
 Send ^c
 Sleep 200
 Send !{F4}
 Sleep 200
 Send n
 Sleep 200
 WinActivate, RStudio
 Sleep 200
 Send ^v
Return	                        ;'replace path' - replaces backslashes with forward slashes in file path. *Should have path in clipboard already*


:*:`;hs::
 Send ^+{home}
 Sleep 200 
 Send ^{enter}
 Sleep 500
 Send {down}
Return					;'highlight start' [but actually 'buy from start']

:*:`;mw::
 MouseMove 3820,175
 Sleep, 500
 Click
Return					;'maximise window' of script

::`;s::               ;'str()
 Send str()
 Send {left}
Return

::`;h::               ;'head('
 Send head()
 Send {left}
Return

::hh::               ;hashkey followed by space.
 Send +3
 Send {space}
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

:*:`;fu::             ;'full  underline'
 Send +3
 Send __________________________________________________________
 Send {enter}
 Send {enter}
Return

:*:`;el::             ;'equal line'
 Send +3
 Send {+}{+}{+}{+}{+}{+}{+}{+}{+}{+}{+}{+}{+}{+}{+}{+}{+}{+}{+}
 Send {enter}
Return

::`;t::
 Send tempDF                ;note you have to push enter or space for the tempDFs... cost otherwise you can't do 'tempDF$' without first activating 'tempDF'
Return

::`;td::
 Send tempDF$
Return
 
::`;ts::
 Send tempDF_subset
Return

:*:,,::`<- `		;hit comma twice to assign variable

;#### temporary bookmarking:

:*:`;mh::				; "Mark Here". "*" means no ending character required to activate
 Send {home}				
 Send `===== TEMPORARY BOOKMARK =====	
 Send {enter}
Return					

:*:`;gm::				; "Go Mark". "*" means no ending character required to activate
 Send ^f				
 Sleep 200
 Send `===== TEMPORARY BOOKMARK =====	
 Send {enter}
 Sleep 200
 Send {escape}
Return		

::b1..::Bird10$
::b3..::Bird30$
::ne..::NAT.EX
::nb..::NAT.BROAD
::neb..::NAT.EX.BROAD
::nbt..::NAT.BROAD.TEA
::nebt..::NAT.EX.BROAD.TEA
::nf..::NAT.FOREST
::sn..::SP.NAT
::st..::SP.TOT
::si..::SP.INT				;variable names

;# BOOKMARK NAVIGATION=====================================

::`;b1::				;press comma, b, m, <number> to navigate to bookmark
 Send ^f
 Sleep 200
 Send ^a {del}			
 Send +3 			;hash
 Sleep 200 
 Send {home} 
 Send {delete} 
 Send {end}	;deletes space automatically insterted before hash
 Send 1.	
 Sleep 200
 Send {escape}
Return		


::`;b2::				;press comma, b, m, <number> to navigate to bookmark
 Send ^f
 Sleep 200
 Send ^a {del}			
 Send +3 			;hash
 Sleep 200 
 Send {home} 
 Send {delete} 
 Send {end}	;deletes space automatically insterted before hash
 Send 2.
 Sleep 200
 Send {escape}
Return		


::`;b3::				;press comma, b, m, <number> to navigate to bookmark
 Send ^f
 Sleep 200
 Send ^a {del}			
 Send +3 			;hash
 Sleep 200 
 Send {home} 
 Send {delete} 
 Send {end}	;deletes space automatically insterted before hash
 Send 3.	
 Sleep 200
 Send {escape}
Return		

::`;b4::				;press comma, b, m, <number> to navigate to bookmark
 Send ^f
 Sleep 200
 Send ^a {del}			
 Send +3 			;hash
 Sleep 200 
 Send {home} 
 Send {delete} 
 Send {end}	;deletes space automatically insterted before hash
 Send 4.
 Sleep 200
 Send {escape}
Return		

::`;b5::				;press comma, b, m, <number> to navigate to bookmark
 Send ^f
 Sleep 200
 Send ^a {del}			
 Send +3 			;hash
 Sleep 200 
 Send {home} 
 Send {delete} 
 Send {end}	;deletes space automatically insterted before hash
 Send 5.
 Sleep 200
 Send {escape}
Return		

::`;b6::				;press comma, b, m, <number> to navigate to bookmark
 Send ^f
 Sleep 200
 Send ^a {del}			
 Send +3 			;hash
 Sleep 200 
 Send {home} 
 Send {delete} 
 Send {end}	;deletes space automatically insterted before hash
 Send 6.
 Sleep 200
 Send {escape}
Return		

::`;b7::				;press comma, b, m, <number> to navigate to bookmark
 Send ^f
 Sleep 200
 Send ^a {del}			
 Send +3 			;hash
 Sleep 200 
 Send {home} 
 Send {delete} 
 Send {end}	;deletes space automatically insterted before hash
 Send 7.	
 Sleep 200
 Send {escape}
Return		


::`;b8::				;press comma, b, m, <number> to navigate to bookmark
 Send ^f
 Sleep 200
 Send ^a {del}			
 Send +3 			;hash
 Sleep 200 
 Send {home} 
 Send {delete} 
 Send {end}	;deletes space automatically insterted before hash
 Send 8.
 Sleep 200
 Send {escape}
Return		


::`;b9::				;press comma, b, m, <number> to navigate to bookmark
 Send ^f
 Sleep 200
 Send ^a {del}			
 Send +3 			;hash
 Sleep 200 
 Send {home} 
 Send {delete} 
 Send {end}	;deletes space automatically insterted before hash
 Send 9.	
 Sleep 200
 Send {escape}
Return		

::`;b10::				;press comma, b, m, <number> to navigate to bookmark
 Send ^f
 Sleep 200
 Send ^a {del}			
 Send +3 			;hash
 Sleep 200 
 Send {home} 
 Send {delete} 
 Send {end}	;deletes space automatically insterted before hash
 Send 10.
 Sleep 200
 Send {escape}
Return		

::`;b11::				;press comma, b, m, <number> to navigate to bookmark
 Send ^f
 Sleep 200
 Send ^a {del}			
 Send +3 			;hash
 Sleep 200 
 Send {home} 
 Send {delete} 
 Send {end}	;deletes space automatically insterted before hash
 Send 11.
 Sleep 200
 Send {escape}
Return		

::`;b12::				;press comma, b, m, <number> to navigate to bookmark
 Send ^f
 Sleep 200
 Send ^a {del}			
 Send +3 			;hash
 Sleep 200 
 Send {home} 
 Send {delete} 
 Send {end}	;deletes space automatically insterted before hash
 Send 12.
 Sleep 200
 Send {escape}
Return		

::`;b13::				;press comma, b, m, <number> to navigate to bookmark
 Send ^f
 Sleep 200
 Send ^a {del}			
 Send +3 			;hash
 Sleep 200 
 Send {home} 
 Send {delete} 
 Send {end}	;deletes space automatically insterted before hash
 Send 13.
 Sleep 200
 Send {escape}
Return		

::`;b14::				;press comma, b, m, <number> to navigate to bookmark
 Send ^f
 Sleep 200
 Send ^a {del}			
 Send +3 			;hash
 Sleep 200 
 Send {home} 
 Send {delete} 
 Send {end}	;deletes space automatically insterted before hash
 Send 14.
 Sleep 200
 Send {escape}
Return		

::`;b15::				;press comma, b, m, <number> to navigate to bookmark
 Send ^f
 Sleep 200
 Send ^a {del}			
 Send +3 			;hash
 Sleep 200 
 Send {home} 
 Send {delete} 
 Send {end}	;deletes space automatically insterted before hash
 Send 15.
 Sleep 200
 Send {escape}
Return		


#IfWinActive
					
;=====================================================================================================
;=====================================================================================================

;FIREFOX

#IfWinActive aahk_class MozillaWindowClass					;all commands bewteen the two IfWinActive commands only apply to FF.

#1::
Click 450,384
Return				

#2::
Click 450, 431
Return				
				;gmail open first,2nd, 3rd mail etc
#3::
Click 450, 478
Return				

#4::
Click 450, 531
Return				

^!p::
Click 11,63			; save pdf
Return

#IfWinActive



;=====================================================================================================
;=====================================================================================================

;ARCGIS

#IfWinActive ArcMap					;all commands bewteen the two IfWinActive commands only apply to GIS.

^p::
MouseGetPos, xpos, ypos 
Click 1380, 60
Click %xpos%, %ypos%, 0		;pan
Return	
			
^z::
MouseGetPos, xpos, ypos 
Click 1340, 60
Click %xpos%, %ypos%, 0		;zoom
Return	
			
+^s::
MouseGetPos, xpos, ypos 
Click 1535, 60
Click %xpos%, %ypos%, 0		;select by rect
Return	
			
^l::
MouseGetPos, xpos, ypos 
Click 1565, 60
Click %xpos%, %ypos%, 0
Return				;clear selectn

^d::
MouseGetPos, xpos, ypos 
Click 1620, 60
Click %xpos%, %ypos%, 0		;identify tool
Return				

^g::
MouseGetPos, xpos, ypos 
Click 1410, 60
Click %xpos%, %ypos%, 0		;glob extent
Return				

^e::
MouseGetPos, xpos, ypos 
Click 1490, 60
Click %xpos%, %ypos%, 0		;prev extent
Return

#c::
Send {esc}
Send !wt
Sleep, 500
Click 1700,500
Return				;cata window

#1::
Send {esc}
Send !wt
Sleep, 500
Click 1700,500
Send {left}{left}{left}{left}{left}{left}{left}{left}{left}{home}f{right}{down}{down}{right}dd{right}
Return											;expand default1

#f::
Send {esc}
Send !wt
Sleep, 500
Click 1700,500
Send {left}{left}{left}{left}{left}{left}{left}{left}{left}{home}f{right}{down}{down}{right}d{right}
Return											;expand default

#s::
Send {esc}
Send !wt
Sleep, 500
Click 1700,500
Send {left}{left}{left}{left}{left}{left}{left}{left}{left}{home}{right}
Return											;expand default1 shapefiles

f3::
Click 170, 80
Return				; table of contents

#IfWinActive							


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

:*:`;gm::				; "Go Mark". "*" means no ending character required to activate
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
;ADOBE READER

#IfWinActive Adobe

F11::
Send !v
Sleep, 200
Send f
Return				; full screen

#IfWinActive


#IfWinActive Internet Explorer

RAlt & down::
 Send ^`-
Return

RAlt & up::                  ;not working
 Send ^`+
Return

#IfWinActive
					
;=====================================================================================================
;===================================================================================================

;WINDOWS EXPLORER

#IfWinActive ahk_class ExploreWClass 
#w::               
 Send !d
Return                                             ; highlights address bar. Note hotstrings don't work here.

#IfWinActive

#IfWinActive ahk_class CabinetWClass               ;2 possible window names

#w::               
 Send !d
Return

#IfWinActive


;=====================================================================================================
;=====================================================================================================
;===================================================================================================

;OUTLOOK

#IfWinActive Outlook 

:*:`;hg::
 Send Hey guys,
 Send {enter}{enter}			
Return

:*:`;hr::
 Send Hey Rach,
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

:*:`;hr::
 Send Hey Rach,
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

;=====================================================================================================
;=====================================================================================================
;=====================================================================================================
;=====================================================================================================




;MOUSE POSITION INFO FROM http://www.autohotkey.com/docs/Tutorial.htm:
;Mouse Clicks: To send a mouse click to a window it is first necessary to determine the X and Y coordinates where the click should occur. This can be done with either AutoScriptWriter or Window ;Spy, which are included with AutoHotkey. The following steps apply to the Window Spy method:

;    Launch Window Spy from a script's tray-icon menu or the Start Menu.
 ;   Activate the window of interest either by clicking its title bar, alt-tabbing, or other means (Window Spy will stay "always on top" by design).
  ;  Move the mouse cursor to the desired position in the target window and write down the mouse coordinates displayed by Window Spy (or on Windows XP and earlier, press Shift-Alt-Tab to activate ;Window Spy so that the "frozen" coordinates can be copied and pasted).
 ;   Apply the coordinates discovered above to the Click35 command. The following example clicks the left mouse button:
  ;  Click 112, 223

;To move the mouse without clicking, use MouseMove36. To drag the mouse, use MouseClickDrag37. 