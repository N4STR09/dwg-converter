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

MsgBox, 64, Setup, Coloca el raton sobre DWG y pulsa F12.
KeyWait, F12, D
MouseGetPos, DWG_X, DWG_Y

MsgBox, 64, Setup, Coloca el raton sobre la línea de comandos y pulsa F12.
KeyWait, F12, D
MouseGetPos, CMD_X, CMD_Y

; ============================
; SELECCIONAR CARPETA
; ============================

FileSelectFolder, Carpeta, , 3, Selecciona la carpeta con los archivos
if Carpeta =
{
    MsgBox, No seleccionaste carpeta. Saliendo.
    ExitApp
}

; ============================
; GENERAR LISTA FILTRADA
; ============================

FileList := ""

Loop, %Carpeta%\*.*, 0
{
    NombreCompleto := A_LoopFileName

    ; ============================
    ; CORRECCIÓN DEFINITIVA
    ; NO cortar por puntos
    ; NO interpretar extensiones
    ; Usar el nombre EXACTO
    ; ============================

    Base := NombreCompleto   ; nombre tal cual

    ; Si ya existe el DWG correspondiente → saltar
    DWGPath := Carpeta . "\" . Base . ".dwg"
    if FileExist(DWGPath)
        continue

    ; Si el archivo ES un DWG → ignorarlo
    if (SubStr(Base, -3) = ".dwg" or SubStr(Base, -3) = ".DWG") or SubStr(NombreCompleto, -3) = ".bak"
    or SubStr(NombreCompleto, -3) = ".tmp" or SubStr(NombreCompleto, -3) = ".log"
        continue

    ; Añadir a la lista
    FileList := FileList . Base . "`n"
}

MsgBox, 64, OK, Archivos pendientes de exportar:`n%FileList%

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