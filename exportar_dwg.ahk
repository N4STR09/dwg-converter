#SingleInstance Force
SetKeyDelay, 50
SetMouseDelay, 50
SetWinDelay, 100

; ============================
; SETUP SIEMPRE
; ============================

MsgBox, 64, Setup, Coloca el raton sobre ALMACENAR y pulsa F12.
KeyWait, F12, D
MouseGetPos, ALMACENAR_X, ALMACENAR_Y
MsgBox, 64, OK, ALMACENAR:`nX=%ALMACENAR_X%  Y=%ALMACENAR_Y%

MsgBox, 64, Setup, Coloca el raton sobre DWG y pulsa F12.
KeyWait, F12, D
MouseGetPos, DWG_X, DWG_Y
MsgBox, 64, OK, DWG:`nX=%DWG_X%  Y=%DWG_Y%

MsgBox, 64, Setup, Coloca el raton sobre la línea de comandos y pulsa F12.
KeyWait, F12, D
MouseGetPos, CMD_X, CMD_Y
MsgBox, 64, OK, Línea de comandos:`nX=%CMD_X%  Y=%CMD_Y%

; ============================
; PROGRAMA
; ============================

; Activar la ventana de OneSpace
; WinActivate, OneSpace
; WinWaitActive, OneSpace
Sleep, 5000

; ================================
;  SCRIPT DE EXPORTACIÓN MASIVA DWG
;  OneSpace Designer Drafting 2002+
;  Windows NT
; ================================

#SingleInstance Force
SetTitleMatchMode, 2
SetKeyDelay, 50
SetMouseDelay, 50
SetWinDelay, 100W

; Lista de archivos SIN extensión
FileList := "
(
3.11.00-1
3.11.00
3
03.26.03-001U
03.26.03-001a
03.26.03-001
)"

; -------------------------------
; BUCLE PRINCIPAL
; -------------------------------

Loop, Parse, FileList, `n, `r
{
    Nombre := A_LoopField
    if (Nombre = "")
        continue

    ; Clic en ALMACENAR
    Click, %ALMACENAR_X%, %ALMACENAR_Y%
    Sleep, 300

    ; Clic en DWG
    Click, %DWG_X%, %DWG_Y%
    Sleep, 300

    ; Clic en la línea de comandos (CAMBIA ESTAS COORDENADAS)
    Click, %CMD_X%, %CMD_Y%
    Sleep, 200

    ; Escribir nombre
    Send, '%Nombre%
    Sleep, 200

    Send, {Enter}
    Sleep, 1200

}

MsgBox, Proceso terminado.
ExitApp
