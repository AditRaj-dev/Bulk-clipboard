CoordMode, Mouse, Screen
CoordMode, Pixel, Screen

^!l::  ; Ctrl + Alt + L starts automation
Loop, 41
{
    ; Click answer/input area
    MouseMove, 2059, 621, 0
    Click
    Sleep, 100

    ; Select all
    Send, ^a
    Sleep, 100

    ; Paste
    Send, ^v
    Sleep, 150

    ; Click Submit button
    MouseMove, 2694, 1711, 0
    Click
    Sleep, 3000   ; wait 3 seconds

    ; Click Next Question button
    MouseMove, 1658, 1227, 0
    Click
    Sleep, 300    ; small buffer before next loop
}
return

; Emergency stop
Esc::ExitApp
