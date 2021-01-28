Attribute VB_Name = "Roomiing_List"
'prog principal

    '-------------------------------------------------------------------------------
    ' RESTO DE CODIGO VER SI SIRVEN
    
    ' OTRA MANERA DE SELECCIONAR RANGO
    ' variable = X
'rango = ("B2" & ":E" & variable)
'range(rango).Select 'selecciona el rango B2:E hasta la fila indicada en la variable
    
    ' LA QUE USAMOS range("A5:O" & ultimaFila & "").Select
    
    'range(Selection, Selection.End(xlToRight)).Select
    'range(Selection, Selection.End(xlDown)).Select
    'ActiveSheet.ListObjects.Add(xlSrcRange, range("$A$5:$O$27"), , xlYes).Name = _
        '"Tabla1"
    'rango = ("A5" & ":O" & ultimaFila)
    ' range("A5:O27").Select
    '-----------------------------------------------------------------------------------

Sub Grupo_Rooming_List()
Attribute Grupo_Rooming_List.VB_ProcData.VB_Invoke_Func = "r\n14"

Call Achico_Matriz

Call Formato_Tabla

Call calculo_totales_hab_Y_paxs

MsgBox "FORMATO APLICADO", vbExclamation

End Sub

Private Sub calculo_totales_hab_Y_paxs()

Dim celda As Range
Dim fecha_in, fecha_out As Date
Dim noches, HAB, noches_totales, fila, PAXS As Integer
Dim rango As Object

'automatico
ActiveSheet.Range("A5").Select
ActiveCell.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Select
PAXS = Application.WorksheetFunction.CountA(Selection)


For Each celda In Selection
    If celda = celda.Offset(1, 0).Value Then
      'nothing
    Else
    fecha_in = ActiveCell.Offset(fila, 10).Value
    fecha_out = ActiveCell.Offset(fila, 11).Value
    noches = DateDiff("D", fecha_in, fecha_out)
    HAB = HAB + 1
    noches_totales = noches_totales + noches
    fila = fila + 1
    End If
Next
      '------------------------------------------
      'DETALLE
      '-------------------------------------------
       Range("S1").Value = Range("B3").Value
       Range("S2").Value = HAB & " " _
       & "Habitaciones" & " " _
       & "Por" & " " & noches_totales & " " _
       & "Noches Totales"
       Range("S3").Value = PAXS & " " & "Paxs"
       Range("S4").Value = "Loisuites Hoteles"
      '--------------------------------------------
     End Sub


Private Sub Achico_Matriz()
     
        Range("A5").Select
    Selection.ColumnWidth = 10
               
      Range("B5:E5").Select
    Selection.ColumnWidth = 0
    
          Range("G5").Select
    Selection.ColumnWidth = 20
    
    Range("H5:J5").Select
     Selection.ColumnWidth = 0
       
    Range("K5:L5").Select
   Selection.ColumnWidth = 10
    
      Range("M5:N5").Select
    Selection.ColumnWidth = 0
    
   Range("P5:Q5").Select
    Selection.ColumnWidth = 0
    
    Range("O5").Select
   Selection.ColumnWidth = 10
      
       ActiveCell.FormulaR1C1 = "iva incl"
    With ActiveCell.Characters(Start:=1, Length:=11).Font
     .Name = "Arial"
     .FontStyle = "Negrita"
      .Size = 11
       .Strikethrough = False
       .Superscript = False
        .Subscript = False
       .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .Color = -1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
   End With
  
End Sub


Private Sub Formato_Tabla()
'doy formato de tabla

'DECLARO FILA COMO LONG
Dim ultimaFila As Long
      
      'BUSCO ULTIMA FILA CON DATOS
        ultimaFila = ActiveSheet.Columns("A").Find("*", _
        SearchOrder:=xlByRows, searchdirection:=xlPrevious).Row
                       
        'SELECCIONO RANGO AL QUE QUIERO DAR FORMATO
        Range("A5:O" & ultimaFila & "").Select
                   
         'LE DOY NOMBRE A LA TABLA
         ActiveSheet.ListObjects.Add(xlSrcRange, Range("A5:O" & ultimaFila & ""), xlYes).Name = _
        "Tabla1"
        
   ' LE DOY FORMATO A LA TABLA
    ActiveSheet.ListObjects("Tabla1").TableStyle = "TableStyleLight16"
    
    

    
    End Sub


