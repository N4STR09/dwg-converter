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

DWG_X := 1717
DWG_Y := 190

CMD_X := 1188
CMD_Y := 375

; Lista de archivos SIN extensión
FileList := "
(
pieza_001
pieza_002
pieza_003
; ... añade aquí los 20.000 nombres
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
    Sleep, 800

}

MsgBox, Proceso terminado.
ExitApp
