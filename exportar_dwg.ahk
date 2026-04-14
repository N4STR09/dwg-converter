#SingleInstance Force
SetKeyDelay, 50
SetMouseDelay, 50
SetWinDelay, 100

; ============================
; CONFIGURACIÓN DEL LOG
; ============================

LogFile := A_ScriptDir . "\export_log.txt"
CSVFile := A_ScriptDir . "\export_resultados.csv"

; Cabecera del log TXT
FileAppend, `n`n==============================`n, %LogFile%
FileAppend, Inicio: %A_Now%`n, %LogFile%
FileAppend, ==============================`n, %LogFile%

; Cabecera del CSV (solo si no existe)
if !FileExist(CSVFile)
    FileAppend, Nombre;Estado;Ruta;FechaHora`n, %CSVFile%

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

FileAppend, Carpeta seleccionada: %Carpeta%`n, %LogFile%

; ============================
; GENERAR LISTA FILTRADA
; ============================

FileList := ""

Loop, %Carpeta%\*.*, 0
{
    NombreCompleto := A_LoopFileName

    ; ============================
    ; EXCEPCIONES DE EXTENSIONES
    ; ============================

    if (SubStr(NombreCompleto, -3) = ".bak"
     or SubStr(NombreCompleto, -3) = ".tmp"
     or SubStr(NombreCompleto, -3) = ".log")
    {
        FileAppend, Ignorado por extensión: %NombreCompleto%`n, %LogFile%
        FileAppend, %NombreCompleto%;Ignorado extensión;%Carpeta%;%A_Now%`n, %CSVFile%
        continue
    }

    ; ============================
    ; USAR NOMBRE COMPLETO TAL CUAL
    ; ============================

    Base := NombreCompleto

    ; Si ya existe el DWG correspondiente → saltar
    DWGPath := Carpeta . "\" . Base . ".dwg"
    if FileExist(DWGPath)
    {
        FileAppend, Saltado (ya existe DWG): %Base%`n, %LogFile%
        FileAppend, %Base%;Saltado (ya existe DWG);%Carpeta%;%A_Now%`n, %CSVFile%
        continue
    }

    ; Si el archivo ES un DWG → ignorarlo
    if (SubStr(Base, -3) = ".dwg" or SubStr(Base, -3) = ".DWG")
    {
        FileAppend, Ignorado (es DWG): %Base%`n, %LogFile%
        FileAppend, %Base%;Ignorado (es DWG);%Carpeta%;%A_Now%`n, %CSVFile%
        continue
    }

    ; Añadir a la lista
    FileList := FileList . Base . "`n"
    FileAppend, Añadido a la cola: %Base%`n, %LogFile%
    FileAppend, %Base%;Añadido a cola;%Carpeta%;%A_Now%`n, %CSVFile%
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

    FileAppend, Procesando: %Nombre%`n, %LogFile%
    FileAppend, %Nombre%;Procesando;%Carpeta%;%A_Now%`n, %CSVFile%

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

    ; Registrar resultado
    if ErrorLevel
    {
        FileAppend, ERROR procesando: %Nombre%`n, %LogFile%
        FileAppend, %Nombre%;ERROR;%Carpeta%;%A_Now%`n, %CSVFile%
    }
    else
    {
        FileAppend, OK: %Nombre%`n, %LogFile%
        FileAppend, %Nombre%;OK;%Carpeta%;%A_Now%`n, %CSVFile%
    }
}

MsgBox, Proceso terminado.

FileAppend, Fin: %A_Now%`n, %LogFile%
FileAppend, ==============================`n, %LogFile%

FileAppend, FIN;Script completado;%Carpeta%;%A_Now%`n, %CSVFile%

ExitApp
