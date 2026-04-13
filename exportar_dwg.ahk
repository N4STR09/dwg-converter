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

; ============================
; SETUP
; ============================

MsgBox, 64, Setup, Vamos a configurar las coordenadas. Pulsa OK para continuar.

; --- ALMACENAR ---
MsgBox, 64, Setup, Coloca el ratón sobre ALMACENAR y pulsa F12.
Hotkey, F12, SetAlmacenar, On
KeyWait, F12, D
Hotkey, F12, Off

; --- DWG ---
MsgBox, 64, Setup, Coloca el ratón sobre DWG y pulsa F12.
Hotkey, F12, SetDWG, On
KeyWait, F12, D
Hotkey, F12, Off

; --- LÍNEA DE COMANDOS ---
MsgBox, 64, Setup, Coloca el ratón sobre la línea de comandos y pulsa F12.
Hotkey, F12, SetCMD, On
KeyWait, F12, D
Hotkey, F12, Off

MsgBox, 64, Setup, Configuración completada. Ya puedes ejecutar el script normalmente.
ExitApp

; ============================
; LEER CONFIGURACIÓN
; ============================

IniRead, ALMACENAR_X, %ConfigFile%, COORDS, ALMACENAR_X
IniRead, ALMACENAR_Y, %ConfigFile%, COORDS, ALMACENAR_Y
IniRead, DWG_X, %ConfigFile%, COORDS, DWG_X
IniRead, DWG_Y, %ConfigFile%, COORDS, DWG_Y
IniRead, CMD_X, %ConfigFile%, COORDS, CMD_X
IniRead, CMD_Y, %ConfigFile%, COORDS, CMD_Y

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
