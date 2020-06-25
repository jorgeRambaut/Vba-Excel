Attribute VB_Name = "CopiarCalendario"

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
'Range("descripcion").Select
'copio
Selection.Copy
'me voy a hoja 1
Sheets("CALENDARIO").Select
'selecciono
Range("a1").Select
Call seleccionar_ultima_Fila_Vacia

'pego
Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
' Vuelvo a Hoja calendario
Sheets("Hoja1").Select

End Sub


Sub Calendario()
Attribute Calendario.VB_ProcData.VB_Invoke_Func = "l\n14"

'ver como se nombran las hojas y el archivo Hoja1 y CALENDARIO

Dim celda As Range
    For Each celda In Selection
    
            If celda Like "Hotel solicitado" Then
            celda.Offset(0, 1).Select
            Range(Selection, Selection.End(xlDown)).Select
            Call Copiar_y_Pegar
            Call seleccionar_ultima_Fila_Vacia
        End If
    Next
 
End Sub

'Codigos Varios
'ultimafila = ActiveSheet.Columns("A").Find("*", _
        searchorder:=xlByRows, searchdirection:=xlPrevious).Row



