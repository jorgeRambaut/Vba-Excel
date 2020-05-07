Attribute VB_Name = "ConversionFechas"
Sub ConversionFechasConHHMMSS()
Dim celda As Range
Range(Selection, Selection.End(xlDown)).Select
Selection.SpecialCells(xlCellTypeVisible).Select
    For Each celda In Selection
    celda = FormatDateTime(celda, vbShortDate)
    Next celda
 
End Sub


Sub ConversionFechasFormatoYYYYMMDD()
Dim celda As Range
Range(Selection, Selection.End(xlDown)).Select
Selection.SpecialCells(xlCellTypeVisible).Select
    For Each celda In Selection
    celda = CDate(celda)
    Next celda
 
End Sub


Sub ConversionPrecio()
Dim celda As Range
Range(Selection, Selection.End(xlDown)).Select
Selection.SpecialCells(xlCellTypeVisible).Select
For Each celda In Selection
celda = Str(celda)
Next celda

End Sub


Sub PasarAdolar()
Dim celda As Range
Dim TipoDecambio As Double
TipoDecambio = InputBox("Ingressar Tipo de Cambio")
Range(Selection, Selection.End(xlDown)).Select
Selection.SpecialCells(xlCellTypeVisible).Select

    For Each celda In Selection
        celda = celda / TipoDecambio
    Next celda
End Sub
