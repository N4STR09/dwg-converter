; Activar la ventana de OneSpace
Run, "C:\Archivos de programa\CoCreate\OSD_Drafting_11.65\bin\Start.exe"
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

; -------------------------------
; CONFIGURACIÓN
; -------------------------------

; Coordenadas (CAMBIA ESTAS)
ALMACENAR_X := 1480
ALMACENAR_Y := 150

DWG_X := 1576
DWG_Y := 190

CMD_X := 30
CMD_Y := 840

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
