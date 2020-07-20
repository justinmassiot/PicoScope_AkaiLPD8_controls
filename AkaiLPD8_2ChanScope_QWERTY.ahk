
; parameters are to be set here
DeviceID := 0
CALLBACK_WINDOW := 0x10000


#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#Persistent


; ======================================
; Start the MIDI device through the WinMM DLL

Gui, +LastFound
hWnd := WinExist()
OpenCloseMidiAPI()
OnExit, Sub_Exit

hMidiIn = ; create an empty object
VarSetCapacity(hMidiIn, 4, 0)
result := DllCall("winmm.dll\midiInOpen", UInt, &hMidiIn, UInt, DeviceID, UInt, hWnd, UInt, 0, UInt, CALLBACK_WINDOW, "UInt")
;result := DllCall("winmm.dll\midiInOpen", UInt, &hMidiIn, UInt, DeviceID)

If result
  GoSub, sub_exit ; silent close if no MIDI device is found

hMidiIn := NumGet(hMidiIn) ; because midiInOpen writes the value in 32 bit binary number, AHK stores it as a string
result := DllCall("winmm.dll\midiInStart", UInt, hMidiIn)
If result
{
  MsgBox, 48, MIDI failure, Cant start the MIDI communication`n(midiInStart returned %result%)
  ;       48 = Warning
  GoSub, sub_exit
}
Else{
  MsgBox, 64, MIDI device found, MIDI device found at ID=%DeviceID%`nhMidiIn = %hMidiIn%
  ;       64 = Info
}


; ======================================
; Bind MIDI messages to the event handler

;  #define MM_MIM_OPEN         0x3C1           /* MIDI input */
;  #define MM_MIM_CLOSE        0x3C2
;  #define MM_MIM_DATA         0x3C3
;  #define MM_MIM_LONGDATA     0x3C4
;  #define MM_MIM_ERROR        0x3C5
;  #define MM_MIM_LONGERROR    0x3C6
OnMessage(0x3C1, "midiInHandler")
OnMessage(0x3C2, "midiInHandler")
OnMessage(0x3C3, "midiInHandler")
OnMessage(0x3C4, "midiInHandler")
OnMessage(0x3C5, "midiInHandler")
OnMessage(0x3C6, "midiInHandler")

return


; ======================================
; Exit function

Sub_Exit:
If (hMidiIn)
  DllCall("winmm.dll\midiInClose", UInt, hMidiIn)
OpenCloseMidiAPI()
ExitApp


; ======================================
; Open/Close function

OpenCloseMidiAPI()
{
  Static hModule
  If hModule
    DllCall("FreeLibrary", UInt, hModule), hModule := ""
  If (0 = hModule := DllCall("LoadLibrary", Str, "winmm.dll")) {
    MsgBox Cannot load library winmm.dll
    ExitApp
  } 
}


; ======================================
; Event handler

midiInHandler(hInput, midiMsg, wMsg) 
{
  Static Rotary1 := 0
  Static Rotary2 := 0
  Static Rotary3 := 0
  Static Rotary4 := 0
  Static Rotary5 := 0
  Static Rotary6 := 0
  Static Rotary7 := 0
  Static Rotary8 := 0

  ; https://www.midi.org/specifications-old/item/table-1-summary-of-midi-message
  ; midiMsg = h xxJJKKLL (Little Endian)
  ;   LL <=> b 1MMM'NNNN (status) <=> 1=status, MMM=type, NNNN=channel
  ;   KK <=> b 0XXX'XXXX (data 1)
  ;   JJ <=> b 0XXX'XXXX (data 2)
  msgType := (midiMsg >> 4) & 0x07 ; 0 = Note On
                                   ; 1 = Note Off
                                   ; 2 = Polyphonic aftertouch
                                   ; 3 = Control change
                                   ; 4 = Program change
                                   ; 5 = Channel aftertouch
                                   ; 6 = Pitch wheel
                                   ; 7 = System message
  channel := midiMsg & 0x0F
  byteData1 := (midiMsg >> 8) & 0x7F
  byteData2 := (midiMsg >> 16) & 0x7F
  
  ; https://www.autohotkey.com/docs/commands/Send.htm
  If (msgType = 1) { ; touch
    ; get the focus on the PicoScope application before sending any keystroke
    if WinExist("ahk_exe PicoScope.exe") and !WinActive("ahk_exe PicoScope.exe") ; doesn't work with `WinActive(PicoScope)`
      WinActivate, ahk_exe PicoScope.exe
    
    switch byteData1 {
      case 0: Send a ; Channel.#0.Coupling.Next  /  Upper Left touchpad
      case 1: Send s ; Channel.#1.Coupling.Next
      case 2: Send x ; Trigger.TriggerSource.Next
      ;case 3  /  Upper Right touchpad
      case 4: Send +z ; Channel.#0.Enabled  /  Lower Left touchpad
      case 5: Send +x ; Channel.#1.Enabled
      case 6: Send z ; Trigger.TriggerMode.Next
      case 7: Send {Space} ; Run  /  Lower Right touchpad
    }
  }
  
  Else If (msgType = 3) { ; rotary
    ; get the focus on the PicoScope application before sending any keystroke
    if WinExist("ahk_exe PicoScope.exe") and !WinActive("ahk_exe PicoScope.exe") ; doesn't work with `WinActive(PicoScope)`
      WinActivate, ahk_exe PicoScope.exe
    
    switch byteData1 {
      case 1: ; rotary 1  ------------------------------  Upper Left rotary knob
        If (byteData2 = 64) {
          Send +a ; Channel.#0.Offset.Reset
        }
        Else If (byteData2//12 > Rotary1) {
          Send +1+1+1+1+1+1+1+1+1+1 ; Channel.#0.Offset.Decrement
        }
        Else If (byteData2//12 < Rotary1) {
          Send +q+q+q+q+q+q+q+q+q+q ; Channel.#0.Offset.Increment
        }
        Rotary1 := byteData2//12 ; 11 steps
      case 2: ; rotary 2
        If (byteData2 = 64) {
          Send +s ; Channel.#1.Offset.Reset
        }
        Else If (byteData2//12 > Rotary2) {
          Send +2+2+2+2+2+2+2+2+2+2 ; Channel.#1.Offset.Decrement
        }
        Else If (byteData2//12 < Rotary2) {
          Send +w+w+w+w+w+w+w+w+w+w ; Channel.#1.Offset.Increment
        }
        Rotary2 := byteData2//12 ; 11 steps
      case 3: ; rotary 3
        If (byteData2//3 > Rotary3) {
          Send v ; Trigger.Threshold.Increment
        }
        Else If (byteData2//3 < Rotary3) {
          Send b ; Trigger.Threshold.Decrement
        }
        Rotary3 := byteData2//3 ; 42 steps
      case 4: ; rotary 4  -----------------------------  Upper Right rotary knob
        If (byteData2//12 > Rotary4) {
          Send m ; Trigger.PreTrigger.Decrement
        }
        Else If (byteData2//12 < Rotary4) {
          Send n ; Trigger.PreTrigger.Increment
        }
        Rotary4 := byteData2//12 ; 11 steps
      case 5: ; rotary 5  ------------------------------  Lower Left rotary knob
        If (byteData2//12 > Rotary5) {
          Send 1 ; Channel.#0.Range.Previous
        }
        Else If (byteData2//12 < Rotary5) {
          Send q ; Channel.#0.Range.Next
        }
        Rotary5 := byteData2//12 ; 11 steps
      case 6: ; rotary 6
        If (byteData2//12 > Rotary6) {
          Send 2 ; Channel.#1.Range.Previous
        }
        Else If (byteData2//12 < Rotary6) {
          Send w ; Channel.#1.Range.Next
        }
        Rotary6 := byteData2//12 ; 11 steps
      ;case 7
      case 8: ; rotary 8  -----------------------------  Lower Right rotary knob
        If (byteData2//3 > Rotary8) {
          Send {Up} ; CollectionTime.Previous
        }
        Else If (byteData2//3 < Rotary8) {
          Send {Down} ; CollectionTime.Next
        }
        Rotary8 := byteData2//3 ; 42 steps
    }
  }
}


;Esc::GoSub, sub_exit
