CoordMode, Mouse, Screen
CoordMode, Pixel, Screen

^!c::  ; Press F6 to start
Loop, 41
{
    ; Move to first position and click
    MouseMove, 2059, 621, 0
    Click
    Sleep, 100

    ; Ctrl + A
    Send, ^a
    Sleep, 100

    ; Ctrl + C
    Send, ^c
    Sleep, 250

    ; Move to second position and click
    MouseMove, 502, 1709, 0
    Click
    Sleep, 4000
}
return

; Emergency stop
Esc::ExitApp
