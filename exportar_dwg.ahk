#SingleInstance Force          ; Evita que se ejecuten varias copias del script
SetKeyDelay, 50                ; Delay entre teclas enviadas
SetMouseDelay, 50              ; Delay entre clics de ratón
SetWinDelay, 100               ; Delay entre acciones de ventana

; ============================
; CONFIGURACIÓN DEL LOG
; ============================

LogFile := A_ScriptDir . "\export_log.txt"   ; Archivo donde se guardan lo que ha ocurrido en cada proceso

; Cabecera del log
FileAppend, `n`n==============================`n, %LogFile%
FileAppend, Inicio: %A_Now%`n, %LogFile%
FileAppend, ==============================`n, %LogFile%

; ============================
; CONTADORES
; ============================

TotalEncontrados := 0      
TotalIgnorados := 0        
TotalSaltados := 0         
TotalCola := 0             
TotalProcesados := 0       
TotalErrores := 0          

; Para estadísticas avanzadas
TiempoInicioGlobal := A_TickCount
TotalTiempoProcesado := 0

; ============================
; CONFIG.INI
; ============================

ConfigFile := A_ScriptDir . "\config.ini"   ; El archivo externo que controla la configuración 

; Si no existe, se crea con valores por defecto
if !FileExist(ConfigFile){
    IniWrite, C:\Archivos de programa\CoCreate\OSD_Drafting_11.65\old_ui\ME10F, %ConfigFile%, General, RutaOneSpace
    IniWrite, 20000, %ConfigFile%, General, Timeout
    IniWrite, 2000, %ConfigFile%, General, TamanoMinimo
    IniWrite, 20000000, %ConfigFile%, General, TamanoMaximo
}

; Se leen los valores del INI
IniRead, RutaOneSpace, %ConfigFile%, General, RutaOneSpace
IniRead, TimeoutGlobal, %ConfigFile%, General, Timeout
IniRead, TamanoMinimo, %ConfigFile%, General, TamanoMinimo
IniRead, TamanoMaximo, %ConfigFile%, General, TamanoMaximo

; Valores por defecto si el INI está corrupto
if (RutaOneSpace = "ERROR" or RutaOneSpace = "")
    RutaOneSpace := "C:\Archivos de programa\CoCreate\OSD_Drafting_11.65\old_ui\ME10F"

if (TimeoutGlobal = "ERROR" or TimeoutGlobal = "")
    TimeoutGlobal := 20000

if (TamanoMinimo = "ERROR" or TamanoMinimo = "")
    TamanoMinimo := 2000

if (TamanoMaximo = "ERROR" or TamanoMaximo = "")
    TamanoMaximo := 20000000

; ============================
; CONFIGURACIÓN DEL CSV
; ============================

FormatoFecha := A_YYYY "-" A_MM "-" A_DD "_" A_Hour "-" A_Min "-" A_Sec

; CSV único por ejecución
CSVFile := A_ScriptDir . "\export_" . FormatoFecha . ".csv"   

; Cabecera del CSV
FileAppend, Nombre;Estado;Motivo;Fecha;Hora;Color`n, %CSVFile%

; ============================
; FUNCIONES AUXILIARES
; ============================

Log(Msg) {
    global LogFile
    FileAppend, %Msg%`n, %LogFile%   ; Añade texto al log
}

CSV(Nombre, Estado, Motivo, Color) {
    global CSVFile
    if (Color = "")
        Color := "Gris"              ; Color por defecto
    FileAppend, %Nombre%;%Estado%;%Motivo%;%A_YYYY%-%A_MM%-%A_DD%;%A_Hour%:%A_Min%;%Color%`n, %CSVFile%
}

RegistrarIgnorado(Base, Motivo, Color) {
    global TotalIgnorados
    if (Color = "")
        Color := "Gris"

    ; Registro en log
    Log("Ignorado (" . Motivo . "): " . Base)  

    ; Registro en CSV
    CSV(Base, "Ignorado", Motivo, Color)        
    TotalIgnorados++
}

ValidarTamano(Base, Tamano) {
    global TamanoMinimo, TamanoMaximo
    ; Devuelve un texto indicando el motivo del descarte, o vacío si es válido dependiendo de su tamaño
    if (Tamano = 0)
        return "Archivo vacio"
    if (Tamano < TamanoMinimo)
        return "Muy pequeño"
    if (Tamano > TamanoMaximo)
        return "Demasiado grande"
    return ""
}

TieneCaracterIlegal(Base) {
    IllegalChars := "<>:|?*"   ; Caracteres no permitidos
    Loop, Parse, IllegalChars
        if InStr(Base, A_LoopField)
            return A_LoopField
    return ""
}

ArchivoBloqueado(RutaCompleta) {
    ; Intentamos renombrar el archivo
    TempName := RutaCompleta . ".locktest"
    FileMove, %RutaCompleta%, %TempName%, 1
    if ErrorLevel
        return true   
    FileMove, %TempName%, %RutaCompleta%, 1
    return false
}

ProcesarArchivo(Nombre) {
    global ALMACENAR_X, ALMACENAR_Y, DWG_X, DWG_Y, CMD_X, CMD_Y
    global TimeoutGlobal, Carpeta, RutaOneSpace
    global TotalErrores, TotalProcesados, TotalTiempoProcesado

    Log("Procesando: " . Nombre)

    Inicio := A_TickCount   ; Marca de tiempo para el timeout

    ; Secuencia de clics para exportar en OneSpace
    Click, %ALMACENAR_X%, %ALMACENAR_Y%
    Sleep, 300

    Click, %DWG_X%, %DWG_Y%
    Sleep, 300

    Click, %CMD_X%, %CMD_Y%
    Sleep, 200

    ; Enviar el nombre del archivo a OneSpace
    Send, '%Nombre%
    Sleep, 200
    Send, {Enter}

    Loop
    {
        Sleep, 200

        ; Comprobación de timeout
        if (A_TickCount - Inicio > TimeoutGlobal)
        {
            Log("ERROR procesando (timeout): " . Nombre)
            CSV(Nombre, "Error", "Timeout", "Rojo")
            TotalErrores++
            return
        }

        ; Comprobar si OneSpace se ha cerrado
        Process, Exist, ME10F.exe
        if (ErrorLevel = 0)
        {
            Log("ERROR CRITICO: OneSpace se ha cerrado procesando " . Nombre)
            CSV(Nombre, "Error", "OneSpace cerrado", "Rojo")
            TotalErrores++

            MsgBox, 16, ERROR CRITICO, OneSpace se ha cerrado inesperadamente.`n`nEl proceso se detendra.

            Run, %RutaOneSpace%   ; Reinicia OneSpace
            Sleep, 3000

            ResumenFinal()
            ExitApp
        }

        ; Si ya existe el DWG, el archivo se procesó correctamente
        if FileExist(Carpeta . "\" . Nombre . ".dwg")
        {
            Log("OK: " . Nombre)
            CSV(Nombre, "Procesado", "OK", "Verde")
            TotalProcesados++
            TotalTiempoProcesado += (A_TickCount - Inicio)
            return
        }
    }
}

; ============================
; FUNCIÓN RESUMEN FINAL
; ============================

ResumenFinal() {
    global TotalEncontrados, TotalIgnorados, TotalSaltados, TotalCola
    global TotalProcesados, TotalErrores

    ; Resumen que se muestra al final de la ejecución
    Resumen := "Resumen final:`n"
    Resumen .= "Total encontrados: " TotalEncontrados "`n"
    Resumen .= "Ignorados: " TotalIgnorados "`n"
    Resumen .= "Saltados: " TotalSaltados "`n"
    Resumen .= "En cola: " TotalCola "`n"
    Resumen .= "Procesados OK: " TotalProcesados "`n"
    Resumen .= "Errores: " TotalErrores "`n"

    MsgBox, 48, Resumen, %Resumen%
}

; ============================
; ESTADÍSTICAS AVANZADAS
; ============================

EstadisticasAvanzadas() {
    global TotalProcesados, TotalIgnorados, TotalSaltados, TotalEncontrados
    global TiempoInicioGlobal, TotalTiempoProcesado

    ; Cálculo de tiempos y porcentajes
    TiempoTotal := (A_TickCount - TiempoInicioGlobal) / 1000
    TiempoMedio := (TotalProcesados > 0) ? (TotalTiempoProcesado / TotalProcesados / 1000) : 0
    PorMinuto := (TiempoTotal > 0) ? (TotalProcesados / (TiempoTotal / 60)) : 0
    PorcentajeOK := (TotalEncontrados > 0) ? (TotalProcesados / TotalEncontrados * 100) : 0
    PorcentajeIgnorados := (TotalEncontrados > 0) ? (TotalIgnorados / TotalEncontrados * 100) : 0

    ; Texto mostrado en el MsgBox
    Texto := "Estadísticas avanzadas:`n"
    Texto .= "- Tiempo total: " TiempoTotal " s`n"
    Texto .= "- Tiempo medio por archivo: " TiempoMedio " s`n"
    Texto .= "- Archivos por minuto: " PorMinuto "`n"
    Texto .= "- Éxito: " PorcentajeOK " %`n"
    Texto .= "- Ignorados: " PorcentajeIgnorados " %`n"

    MsgBox, 64, Estadísticas, %Texto%
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

Log("Carpeta seleccionada: " . Carpeta)

; ============================
; GENERAR LISTA FILTRADA
; ============================

FileList := ""

Loop, %Carpeta%\*.*, 0
{
    NombreCompleto := A_LoopFileName
    RutaCompleta := Carpeta . "\" . NombreCompleto
    TotalEncontrados++

    ; Extensiones no válidas
    if (SubStr(NombreCompleto, -3) = ".bak"
     or SubStr(NombreCompleto, -3) = ".tmp"
     or SubStr(NombreCompleto, -3) = ".log")
    {
        RegistrarIgnorado(NombreCompleto, "Extension", "")
        continue
    }

    Base := NombreCompleto

    ; Validación de tamaño
    FileGetSize, Tamano, %RutaCompleta%
    motivoTam := ValidarTamano(Base, Tamano)
    if (motivoTam != "")
    {
        RegistrarIgnorado(Base, motivoTam, "")
        continue
    }

    ; Caracteres ilegales
    ilegal := TieneCaracterIlegal(Base)
    if (ilegal != "")
    {
        RegistrarIgnorado(Base, "Caracter ilegal: " . ilegal, "")
        continue
    }

    ; Permisos de lectura
    FileRead, TestLectura, %RutaCompleta%
    if (ErrorLevel)
    {
        RegistrarIgnorado(Base, "Sin permisos", "")
        continue
    }

    ; Archivo bloqueado por otro proceso
    if ArchivoBloqueado(RutaCompleta)
    {
        RegistrarIgnorado(Base, "Bloqueado", "")
        continue
    }

    ; Si ya existe el DWG, no se procesa
    DWGPath := Carpeta . "\" . Base . ".dwg"
    if FileExist(DWGPath)
    {
        Log("Saltado (DWG existe): " . Base)
        CSV(Base, "Saltado", "DWG existente", "Amarillo")
        TotalSaltados++
        continue
    }

    ; Si ya es un DWG, se ignora
    if (SubStr(Base, -3) = ".dwg" or SubStr(Base, -3) = ".DWG")
    {
        RegistrarIgnorado(Base, "Es DWG", "")
        continue
    }

    ; Archivo válido, se añade a la cola
    FileList .= Base . "`n"
    Log("Agregado a la cola: " . Base)
    TotalCola++
}

; ============================
; PROCESO PRINCIPAL
; ============================

Sleep, 3000
SetTitleMatchMode, 2

Loop, Parse, FileList, `n, `r
{
    Nombre := A_LoopField
    if (Nombre = "")
        continue

    ProcesarArchivo(Nombre)
}

ResumenFinal()
EstadisticasAvanzadas()
ExitApp
