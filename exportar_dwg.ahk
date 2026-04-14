#SingleInstance Force
SetKeyDelay, 50
SetMouseDelay, 50
SetWinDelay, 100

; ============================
; CONFIGURACIÓN DEL LOG
; ============================

LogFile := A_ScriptDir . "\export_log.txt"
CSVFile := A_ScriptDir . "\export_resultados.csv"

FileAppend, `n`n==============================`n, %LogFile%
FileAppend, Inicio: %A_Now%`n, %LogFile%
FileAppend, ==============================`n, %LogFile%

if !FileExist(CSVFile)
    FileAppend, Nombre;Estado`n, %CSVFile%

; ============================
; CONTADORES
; ============================

TotalEncontrados := 0
TotalIgnorados := 0
TotalSaltados := 0
TotalCola := 0
TotalProcesados := 0
TotalErrores := 0

; ============================
; SETUP
; ============================

MsgBox, 64, Setup, Coloca el raton sobre ALMACENAR y pulsa F12.
KeyWait, F12, D
MouseGetPos, ALMACENAR_X, ALMACENAR_Y

MsgBox, 64, Setup, Coloca el raton sobre DWG y pulsa F12.
KeyWait, F12, D
MouseGetPos, DWG_X, DWG_Y

MsgBox, 64, Setup, Coloca el raton sobre la linea de comandos y pulsa F12.
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
    TotalEncontrados++

    if (SubStr(NombreCompleto, -3) = ".bak"
     or SubStr(NombreCompleto, -3) = ".tmp"
     or SubStr(NombreCompleto, -3) = ".log")
    {
        FileAppend, Ignorado por extension: %NombreCompleto%`n, %LogFile%
        FileAppend, %NombreCompleto%;Ignorado`n, %CSVFile%
        TotalIgnorados++
        continue
    }

    Base := NombreCompleto

    DWGPath := Carpeta . "\" . Base . ".dwg"
    if FileExist(DWGPath)
    {
        FileAppend, Saltado (ya existe DWG): %Base%`n, %LogFile%
        FileAppend, %Base%;Saltado`n, %CSVFile%
        TotalSaltados++
        continue
    }

    if (SubStr(Base, -3) = ".dwg" or SubStr(Base, -3) = ".DWG")
    {
        FileAppend, Ignorado (es DWG): %Base%`n, %LogFile%
        FileAppend, %Base%;Ignorado`n, %CSVFile%
        TotalIgnorados++
        continue
    }

    FileList .= Base . "`n"
    FileAppend, Agregado a la cola: %Base%`n, %LogFile%
    TotalCola++
}

; ============================
; PROGRAMA
; ============================

Sleep, 3000
SetTitleMatchMode, 2

; ============================
; RUTA DE ONESPACE
; ============================

OneSpaceCMD := "C:\Archivos de programa\CoCreate\OSD_Drafting_11.65\old_ui\ME10F"

; -------------------------------
; BUCLE PRINCIPAL
; -------------------------------

Loop, Parse, FileList, `n, `r
{
    Nombre := A_LoopField
    if (Nombre = "")
        continue

    FileAppend, Procesando: %Nombre%`n, %LogFile%

    Inicio := A_TickCount
    Timeout := 20000

    ; ============================
    ; EXPORTACIÓN
    ; ============================

    Click, %ALMACENAR_X%, %ALMACENAR_Y%
    Sleep, 300

    Click, %DWG_X%, %DWG_Y%
    Sleep, 300

    Click, %CMD_X%, %CMD_Y%
    Sleep, 200

    Send, '%Nombre%
    Sleep, 200
    Send, {Enter}

    ErrorLevel := 0

    ; ============================
    ; ESPERA INTELIGENTE
    ; ============================

    Loop
    {
        Sleep, 200

        ; 1. Timeout
        if (A_TickCount - Inicio > Timeout)
        {
            ErrorLevel := 1
            break
        }

        ; 2. Proceso cerrado → detener todo
        Process, Exist, ME10F.exe
        if (ErrorLevel = 0)
        {
            ; Registrar error
            FileAppend, ERROR CRITICO: OneSpace se ha cerrado procesando %Nombre%`n, %LogFile%
            FileAppend, %Nombre%;Error`n, %CSVFile%
            TotalErrores++

            ; Mostrar mensaje
            MsgBox, 16, ERROR CRITICO, OneSpace se ha cerrado inesperadamente.`n`nEl proceso se detendrá.

            ; Reabrir OneSpace
            Run, %OneSpaceCMD%
            Sleep, 3000

            ; Mostrar resumen final
            Resumen =
            (
Resumen final:

Total encontrados: %TotalEncontrados%
Ignorados: %TotalIgnorados%
Saltados: %TotalSaltados%
En cola: %TotalCola%
Procesados OK: %TotalProcesados%
Errores: %TotalErrores%
            )

            MsgBox, 48, Resumen, %Resumen%

            ExitApp
        }

        ; 3. DWG generado → OK
        if FileExist(Carpeta . "\" . Nombre . ".dwg")
        {
            ErrorLevel := 0
            break
        }
    }

    ; ============================
    ; RESULTADO FINAL DEL ARCHIVO
    ; ============================

    if (ErrorLevel = 1)
    {
        FileAppend, ERROR procesando: %Nombre%`n, %LogFile%
        FileAppend, %Nombre%;Error`n, %CSVFile%
        TotalErrores++
    }
    else
    {
        FileAppend, OK: %Nombre%`n, %LogFile%
        FileAppend, %Nombre%;Procesado`n, %CSVFile%
        TotalProcesados++
    }
}
