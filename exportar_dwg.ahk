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
; FUNCIÓN RESUMEN FINAL
; ============================

ResumenFinal() {
    global TotalEncontrados, TotalIgnorados, TotalSaltados, TotalCola
    global TotalProcesados, TotalErrores

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
}

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

IllegalChars := "<>:|?*"

Loop, %Carpeta%\*.*, 0
{
    NombreCompleto := A_LoopFileName
    RutaCompleta := Carpeta . "\" . NombreCompleto
    TotalEncontrados++

    ; EXTENSIONES PROHIBIDAS
    if (SubStr(NombreCompleto, -3) = ".bak"
     or SubStr(NombreCompleto, -3) = ".tmp"
     or SubStr(NombreCompleto, -3) = ".log")
    {
        FileAppend, Ignorado por extension: %NombreCompleto%`n, %LogFile%
        FileAppend, %NombreCompleto%;Ignorado extension`n, %CSVFile%
        TotalIgnorados++
        continue
    }

    Base := NombreCompleto

    ; ARCHIVO VACÍO / CORRUPTO
    FileGetSize, Tamano, %RutaCompleta%

    if (Tamano = 0)
    {
        FileAppend, Ignorado (archivo vacío): %Base%`n, %LogFile%
        FileAppend, %Base%;Archivo vacio`n, %CSVFile%
        TotalIgnorados++
        continue
    }

    if (Tamano < 2000)
    {
        FileAppend, Ignorado (posible corrupto - muy pequeño): %Base% (%Tamano% bytes)`n, %LogFile%
        FileAppend, %Base%;Posible corrupto (pequeño)`n, %CSVFile%
        TotalIgnorados++
        continue
    }

    if (Tamano > 20000000)
    {
        FileAppend, Ignorado (demasiado grande): %Base% (%Tamano% bytes)`n, %LogFile%
        FileAppend, %Base%;Demasiado grande`n, %CSVFile%
        TotalIgnorados++
        continue
    }

    ; NOMBRES ILEGALES (SIN REGEX)
    Loop, Parse, IllegalChars
    {
        if InStr(Base, A_LoopField)
        {
            FileAppend, Ignorado (caracter ilegal: %A_LoopField%): %Base%`n, %LogFile%
            FileAppend, %Base%;Nombre ilegal`n, %CSVFile%
            TotalIgnorados++
            continue, 2
        }
    }

    ; PERMISOS DE LECTURA
    FileRead, TestLectura, %RutaCompleta%
    if (ErrorLevel)
    {
        FileAppend, Ignorado (sin permisos de lectura): %Base%`n, %LogFile%
        FileAppend, %Base%;Sin permisos`n, %CSVFile%
        TotalIgnorados++
        continue
    }

    ; ARCHIVO BLOQUEADO (MÉTODO COMPATIBLE)
    TempName := RutaCompleta . ".locktest"
    FileMove, %RutaCompleta%, %TempName%, 1
    if ErrorLevel
    {
        FileAppend, Ignorado (archivo bloqueado): %Base%`n, %LogFile%
        FileAppend, %Base%;Bloqueado`n, %CSVFile%
        TotalIgnorados++
        continue
    }
    FileMove, %TempName%, %RutaCompleta%, 1

    ; SI YA EXISTE EL DWG
    DWGPath := Carpeta . "\" . Base . ".dwg"
    if FileExist(DWGPath)
    {
        FileAppend, Saltado (ya existe DWG): %Base%`n, %LogFile%
        FileAppend, %Base%;Saltado`n, %CSVFile%
        TotalSaltados++
        continue
    }

    ; IGNORAR SI YA ES DWG
    if (SubStr(Base, -3) = ".dwg" or SubStr(Base, -3) = ".DWG")
    {
        FileAppend, Ignorado (es DWG): %Base%`n, %LogFile%
        FileAppend, %Base%;Ignorado (DWG)`n, %CSVFile%
        TotalIgnorados++
        continue
    }

    ; AÑADIR A LA COLA
    FileList .= Base . "`n"
    FileAppend, Agregado a la cola: %Base%`n, %LogFile%
    TotalCola++
}

; ============================
; PROCESO PRINCIPAL
; ============================

Sleep, 3000
SetTitleMatchMode, 2

OneSpaceCMD := "C:\Archivos de programa\CoCreate\OSD_Drafting_11.65\old_ui\ME10F"

Loop, Parse, FileList, `n, `r
{
    Nombre := A_LoopField
    if (Nombre = "")
        continue

    FileAppend, Procesando: %Nombre%`n, %LogFile%

    Inicio := A_TickCount
    Timeout := 20000

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

    Loop
    {
        Sleep, 200

        if (A_TickCount - Inicio > Timeout)
        {
            ErrorLevel := 1
            break
        }

        Process, Exist, ME10F.exe
        if (ErrorLevel = 0)
        {
            FileAppend, ERROR CRITICO: OneSpace se ha cerrado procesando %Nombre%`n, %LogFile%
            FileAppend, %Nombre%;Error`n, %CSVFile%
            TotalErrores++

            MsgBox, 16, ERROR CRITICO, OneSpace se ha cerrado inesperadamente.`n`nEl proceso se detendra.

            Run, %OneSpaceCMD%
            Sleep, 3000

            ResumenFinal()
            ExitApp
        }

        if FileExist(Carpeta . "\" . Nombre . ".dwg")
        {
            ErrorLevel := 0
            break
        }
    }

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

ResumenFinal()
ExitApp