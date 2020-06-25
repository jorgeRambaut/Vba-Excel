Attribute VB_Name = "Grupos"

Private Sub seleccionar_ultima_Fila_Vacia()

Do Until IsEmpty(ActiveCell) And IsEmpty(ActiveCell.Offset(1, 0))
ActiveCell.Offset(1, 0).Select
Loop

End Sub

Private Sub BORRAR()

Range("F4:F21").ClearContents



End Sub

Private Sub Copiar_y_Pegar()

'selecciono
Range("descripcion").Select
'copio
Selection.Copy
'me voy a hoja recoleta
Sheets("Recoleta").Select
'selecciono
Range("a1").Select
Call seleccionar_ultima_Fila_Vacia

'pego
Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
   'borro contenido
   Sheets("Formulario de Carga").Select
    Range("descripcion").ClearContents
    '------------------------------
    'TODO ver codigo para ordenar filas
    '-------------------------------
   MsgBox "Grupo Confirmado", vbExclamation


End Sub

'-----------------------------------Ver como usar funcion call ---------------------

Sub Valido_Datos()
If Range("Hotel") = "Loi Suites Recoleta" Then
'MsgBox "Recoleta"
'----------------------------------------------------------------
'TODO valido Datos ver como llamar a este subproceso para todos los if
'----------------------------------------------------------------

If Range("Status") = "" Then
MsgBox "Completar Status"
Exit Sub
End If
If Range("Nombre_de_Grupo") = "" Then
MsgBox "Completar Nombre"
Exit Sub
End If
If Range("CLIENTE") = "" Then
MsgBox "Completar Cliente"
Exit Sub
End If
If Range("Fecha_in") = "" Then
MsgBox "Completar Fecha"
Exit Sub
End If
If Range("Fecha_in").Value > Range("Fecha_out").Value Then
MsgBox "Fecha in no puede ser mayor a out", vbCritical
Exit Sub
End If
If Range("Fecha_out").Value < Range("Fecha_in").Value Then
MsgBox "Fecha out no puede ser Menor a in", vbCritical
Exit Sub
End If
If Range("Hab") = "" Then
MsgBox "Completar Cantidad de Hab"
Exit Sub
End If
If Range("Categoria_Hab") = "" Then
MsgBox "Completar Hab"
Exit Sub
End If
If Range("Tarifa") = "" Then
MsgBox "Completar Tarifa"
Exit Sub
End If
If Range("Comision") = "" Then
MsgBox "Completar Neta o Comisionable"
Exit Sub
End If
If Range("FOC") = "" Then
MsgBox "¿Hay Hab Free?", vbExclamation
Exit Sub
End If
If Range("Forma_de_pago") = "" Then
MsgBox "Completar Forma de pago"
Exit Sub
End If
If Range("Dead_line") = "" Then
MsgBox "Completar Dead_line"
Exit Sub
End If
If Range("Observaciones") = "" Then
MsgBox "El Grupo,¿No tiene Ningun Requerimiento?", vbYesNo

End If
If Range("Ejecutivo") = "" Then
MsgBox "Completar Ejecutivo"
Exit Sub
End If

'selecciono
Range("descripcion").Select
'copio
Selection.Copy
'me voy a hoja recoleta
Sheets("Recoleta").Select
'selecciono
Range("a1").Select
Call seleccionar_ultima_Fila_Vacia

'pego
Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
   'borro contenido
   Sheets("Formulario de Carga").Select
    Range("descripcion").ClearContents
    'TODO------------------------------
    'ver codigo para ordenar filas
    '-------------------------------
   MsgBox "Grupo Confirmado", vbExclamation


ElseIf Range("Hotel") = "Loi Suites Esmeralda" Then

  If Range("Status") = "" Then
  MsgBox "Completar Status"
Exit Sub
End If
If Range("Nombre_de_Grupo") = "" Then
MsgBox "Completar Nombre"
Exit Sub
End If
If Range("CLIENTE") = "" Then
MsgBox "Completar Cliente"
Exit Sub
End If
If Range("Fecha_in").Value > Range("Fecha_out").Value Then
MsgBox "Fecha in no puede ser mayor a out", vbCritical
Exit Sub
End If
If Range("Fecha_out").Value < Range("Fecha_in").Value Then
MsgBox "Fecha out no puede ser Menor a in", vbCritical
Exit Sub
End If
If Range("Hab") = "" Then
MsgBox "Completar Cantidad de Hab"
Exit Sub
End If
If Range("Categoria_Hab") = "" Then
MsgBox "Completar Hab"
Exit Sub
End If
If Range("Tarifa") = "" Then
MsgBox "Completar Tarifa"
Exit Sub
End If
If Range("Comision") = "" Then
MsgBox "Completar Neta o Comisionable"
Exit Sub
End If
If Range("FOC") = "" Then
MsgBox "¿Hay Hab Free?", vbExclamation
Exit Sub
End If
If Range("Forma_de_pago") = "" Then
MsgBox "Completar Forma de pago"
Exit Sub
End If
If Range("Dead_line") = "" Then
MsgBox "Completar Dead_line"
Exit Sub
End If
If Range("Observaciones") = "" Then
MsgBox "El Grupo,¿No tiene Ningun Requerimiento?", vbYesNo

End If
If Range("Ejecutivo") = "" Then
MsgBox "Completar Ejecutivo"
Exit Sub
End If
'selecciono
Range("descripcion").Select
'copio
Selection.Copy
'me voy a hoja recoleta
Sheets("Esmeralda").Select
'selecciono
Range("a1").Select
Call seleccionar_ultima_Fila_Vacia

'pego
Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
   'borro contenido
   Sheets("Formulario de Carga").Select
    Range("descripcion").ClearContents
    '------------------------------
    'ver codigo para ordenar filas
    '-------------------------------
   MsgBox "Grupo Confirmado", vbExclamation

ElseIf Range("Hotel") = "Loi Suites Chapelco" Then
  
    If Range("Status") = "" Then
  MsgBox "Completar Status"
Exit Sub
End If
If Range("Nombre_de_Grupo") = "" Then
MsgBox "Completar Nombre"
Exit Sub
End If
If Range("CLIENTE") = "" Then
MsgBox "Completar Cliente"
Exit Sub
End If
If Range("Fecha_in").Value > Range("Fecha_out").Value Then
MsgBox "Fecha in no puede ser mayor a out", vbCritical
Exit Sub
End If
If Range("Fecha_out").Value < Range("Fecha_in").Value Then
MsgBox "Fecha out no puede ser Menor a in", vbCritical
Exit Sub
End If
If Range("Hab") = "" Then
MsgBox "Completar Cantidad de Hab"
Exit Sub
End If
If Range("Categoria_Hab") = "" Then
MsgBox "Completar Hab"
Exit Sub
End If
If Range("Tarifa") = "" Then
MsgBox "Completar Tarifa"
Exit Sub
End If
If Range("Comision") = "" Then
MsgBox "Completar Neta o Comisionable"
Exit Sub
End If
If Range("FOC") = "" Then
MsgBox "¿Hay Hab Free?", vbExclamation
Exit Sub
End If
If Range("Forma_de_pago") = "" Then
MsgBox "Completar Forma de pago"
Exit Sub
End If
If Range("Dead_line") = "" Then
MsgBox "Completar Dead_line"
Exit Sub
End If
If Range("Observaciones") = "" Then
MsgBox "El Grupo,¿No tiene Ningun Requerimiento?", vbYesNo

End If
If Range("Ejecutivo") = "" Then
MsgBox "Completar Ejecutivo"
Exit Sub
End If
'selecciono
Range("descripcion").Select
'copio
Selection.Copy
'me voy a hoja recoleta
Sheets("Chapelco").Select
'selecciono
Range("a1").Select
Call seleccionar_ultima_Fila_Vacia

'pego
Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
   'borro contenido
   Sheets("Formulario de Carga").Select
    Range("descripcion").ClearContents
    '------------------------------
    'ver codigo para ordenar filas
    '-------------------------------
   MsgBox "Grupo Confirmado", vbExclamation

 ElseIf Range("Hotel") = "Loi Suites Iguazu" Then
      If Range("Status") = "" Then
  MsgBox "Completar Status"
Exit Sub
End If
If Range("Nombre_de_Grupo") = "" Then
MsgBox "Completar Nombre"
Exit Sub
End If
If Range("CLIENTE") = "" Then
MsgBox "Completar Cliente"
Exit Sub
End If
If Range("Fecha_in").Value > Range("Fecha_out").Value Then
MsgBox "Fecha in no puede ser mayor a out", vbCritical
Exit Sub
End If
If Range("Fecha_out").Value < Range("Fecha_in").Value Then
MsgBox "Fecha out no puede ser Menor a in", vbCritical
Exit Sub
End If
If Range("Hab") = "" Then
MsgBox "Completar Cantidad de Hab"
Exit Sub
End If
If Range("Categoria_Hab") = "" Then
MsgBox "Completar Hab"
Exit Sub
End If
If Range("Tarifa") = "" Then
MsgBox "Completar Tarifa"
Exit Sub
End If
If Range("Comision") = "" Then
MsgBox "Completar Neta o Comisionable"
Exit Sub
End If
If Range("FOC") = "" Then
MsgBox "¿Hay Hab Free?", vbExclamation
Exit Sub
End If
If Range("Forma_de_pago") = "" Then
MsgBox "Completar Forma de pago"
Exit Sub
End If
If Range("Dead_line") = "" Then
MsgBox "Completar Dead_line"
Exit Sub
End If
If Range("Observaciones") = "" Then
MsgBox "El Grupo,¿No tiene Ningun Requerimiento?", vbYesNo
End If
If Range("Ejecutivo") = "" Then
MsgBox "Completar Ejecutivo"
Exit Sub
End If
'selecciono
Range("descripcion").Select
'copio
Selection.Copy
'me voy a hoja recoleta
Sheets("Iguazu").Select
'selecciono
Range("a1").Select
Call seleccionar_ultima_Fila_Vacia

'pego
Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
   'borro contenido
   Sheets("Formulario de Carga").Select
    Range("descripcion").ClearContents
    '------------------------------
    'ver codigo para ordenar filas
    '-------------------------------
   MsgBox "Grupo Confirmado", vbExclamation

Else

MsgBox "Por Favor,Ingrese Hotel", vbRetryCancel

End If

End Sub

'Codigos Varios
'ultimafila = ActiveSheet.Columns("A").Find("*", _
        searchorder:=xlByRows, searchdirection:=xlPrevious).Row

