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
; SELECCIONAR CARPETA
; ============================

FileSelectFolder, Carpeta, , 3, Selecciona la carpeta con los archivos
if Carpeta =
{
    MsgBox, No seleccionaste carpeta. Saliendo.
    ExitApp
}

; Construir lista de archivos automáticamente
FileList := ""

Loop, %Carpeta%\*.*, 0
{
    Nombre := A_LoopFileName
    StringSplit, Partes, Nombre, .
    FileList := FileList . Partes1 . "`n"
}

MsgBox, 64, OK, Se han cargado los nombres desde:`n%Carpeta%

; ============================
; PROGRAMA
; ============================

Sleep, 5000

SetTitleMatchMode, 2
SetKeyDelay, 50
SetMouseDelay, 50
SetWinDelay, 100W

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

    ; Clic en la línea de comandos
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