
; "#defines"
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
;MsgBox, hWnd = %hWnd%`nPress OK to open winmm.dll library
OpenCloseMidiAPI()
OnExit, Sub_Exit
;MsgBox, winmm.dll loaded.`nPress OK to open midi device`nDevice ID = %DeviceID%`nhWnd = %hWnd%`ndwFlags = CALLBACK_WINDOW

hMidiIn = ; create an empty object
VarSetCapacity(hMidiIn, 4, 0)
result := DllCall("winmm.dll\midiInOpen", UInt, &hMidiIn, UInt, DeviceID, UInt, hWnd, UInt, 0, UInt, CALLBACK_WINDOW, "UInt")
;result := DllCall("winmm.dll\midiInOpen", UInt, &hMidiIn, UInt, DeviceID)

If result
{
	MsgBox, 48, MIDI failure, Failed to open MIDI device at ID=%DeviceID%`n(midiInOpen returned %result%)
  ;       48 = Warning
	GoSub, sub_exit
}

hMidiIn := NumGet(hMidiIn) ; because midiInOpen writes the value in 32 bit binary number, AHK stores it as a string
MsgBox, 64, MIDI device found, MIDI device found at ID=%DeviceID%`nhMidiIn = %hMidiIn%
;       64 = Info

result := DllCall("winmm.dll\midiInStart", UInt, hMidiIn)
If result
{
	MsgBox, 48, Cant start the MIDI communication`n(midiInStart returned %result%)
  ;       48 = Warning
	GoSub, sub_exit
}


; ======================================
; Bind MIDI messages to the event handler

;	#define MM_MIM_OPEN         0x3C1           /* MIDI input */
;	#define MM_MIM_CLOSE        0x3C2
;	#define MM_MIM_DATA         0x3C3
;	#define MM_MIM_LONGDATA     0x3C4
;	#define MM_MIM_ERROR        0x3C5
;	#define MM_MIM_LONGERROR    0x3C6
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

	ToolTip,
	(
    MIDI message received: %wMsg%
    wParam = %hInput%
    lParam = %midiMsg%
      msgType = %msgType%
      channel = %channel%
      data1 = %byteData1%
      data2 = %byteData2%
	)
}


Esc::GoSub, sub_exit
