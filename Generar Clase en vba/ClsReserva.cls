VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsReserva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iComportamientos


Sub iComportamientos_Ajustaratexto(columna As String)
Range(columna).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub


Sub iComportamientos_seleccionarCeldasDescendente()

  Range(Selection, Selection.End(xlDown)).Select

End Sub


Function iComportamientos_obtenerParte(texto As String, cantidad As Integer) As String
    
        iComportamientos_obtenerParte = Left(texto, cantidad)
   
End Function

Sub iComportamientos_BuscarTitulo(DatoAbuscar As String, NombreDeSolapaAbuscar As String, RangoInicio As String, RangoFin As String, Optional libro As String)
    Dim tituloEncontrado As Range
      Selection.SpecialCells(xlCellTypeVisible).Select
        
            With Worksheets(NombreDeSolapaAbuscar).Range(RangoInicio & ":" & RangoFin)
            
                    Set tituloEncontrado = .Find(DatoAbuscar, LookIn:=xlValues)
                    
                        If Not tituloEncontrado Is Nothing Then
                            tituloEncontrado.Offset(1, 0).Select
'                            Range.Offset(0, 1).Select
'                            tituloEncontrado(Selection, Selection.End(xlDown)).Select
'                            Range(Selection, Selection.End(xlDown)).Select
                        End If
            End With
    
End Sub
Sub iComportamientos_moverCelda(columna As Integer, fila As Integer)
    Dim mover As Range
'      Selection.SpecialCells(xlCellTypeVisible).Select
        Set mover = mover.Offset(fila, columna).Select
            With mover
                    
                    
                        If Not mover Is Nothing Then
                            mover.Offset(fila, columna).Select
'                            Range.Offset(0, 1).Select
'                            tituloEncontrado(Selection, Selection.End(xlDown)).Select
'                            Range(Selection, Selection.End(xlDown)).Select
                        End If
            End With
     
End Sub

'Function iComportamientos_observacion(ota As String, beneficios As String, condicion_reserva As String, forma_de_pago As String)
'
'    iComportamientos_observacion = ota & " " & condicion_reserva & " " & forma_de_pago & " " & beneficios
'
'End Function

Function iComportamientos_pasarADolar(tarifa As Double, tipoDeCambio As Double) As Double

    iComportamientos_pasarADolar = tarifa / tipoDeCambio

End Function

Function iComportamientos_pasarAPesos(tipoDeCambio As Double, tarifa As Double) As Double
    
  iComportamientos_pasarAPesos = (tipoDeCambio * tarifa)

End Function

Sub iComportamientos_ConvertirValores(valor As String)
Dim celda As Range
'Range(Selection, Selection.End(xlDown)).Select
'Selection.SpecialCells(xlCellTypeVisible).Select
Select Case valor
Case "Letra"
    For Each celda In Selection
    celda = CStr(celda)
    Next celda
Case "Numero"
    For Each celda In Selection
    celda = CInt(celda)
    Next celda
    
Case "Fecha"
    For Each celda In Selection
    celda = CDate(celda)
    Next celda
    
End Select
End Sub

Sub iComportamientos_CorregirGds(gds As String, cantidad_a_quitar As Byte)

Dim celda As Range

'Range(Selection, Selection.End(xlDown)).Select
    
    For Each celda In Selection

        If celda.Value Like "*" & gds & "*" Then
        
           celda = Right(celda, cantidad_a_quitar)
            
        End If
        
    Next celda
    
    End Sub
    
  Sub pasarApesosBookassistARS(tipoDeCambio As Double)
  Dim reserva As New ClsReserva
  Dim celda As Range
  Dim arg As String
  Dim Precio As Double
  Dim precioEnPesos As Double
  Dim iva As Double
  iva = 1.21
    For Each celda In Selection
     arg = reserva.iComportamientos_obtenerParte(CStr(celda), 3)
 
        If arg = "ARS" Then
        
          Precio = celda.Offset(0, 17)
        
          precioEnPesos = reserva.iComportamientos_pasarAPesos(tipoDeCambio, Precio)
          
          celda.Offset(0, 17) = precioEnPesos * iva
        
        End If
 
   
Next celda





  End Sub
  
  Sub pasarApesosDespegar(tipoDeCambio As Double)
  Dim reserva As New ClsReserva
  Dim celda As Range
  Dim Ota As String
  Dim Precio As Double
  Dim precioEnPesos As Double
  
    For Each celda In Selection
 
        If celda = "Despegar" Then
        
          Precio = celda.Offset(0, 16)
        
          precioEnPesos = reserva.iComportamientos_pasarAPesos(tipoDeCambio, Precio)
          
          celda.Offset(0, 16) = precioEnPesos
        
        End If
 
   
Next celda





  End Sub
  
    
    
    
    
    
    
Sub iComportamientos_BuscarDuplicadosYResaltar()
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlDuplicate
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    ActiveWindow.SmallScroll Down:=-114
End Sub

Sub iComportamientos_BuscarSiHayMenoresEnLaReservaYResaltarlos()
   
'Dim celda As range
'ActiveCell.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Select
'Selection.SpecialCells(xlCellTypeVisible).Selec
    ActiveWindow.SmallScroll Down:=-12
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub


Function iComportamientos_buscarReembolsable(reembolsable As String) As Boolean

Dim condicion As String
Dim esReembolsable As Boolean

condicion = "Reembolsable"
    
        
       If condicion = reembolsable Then
           
       esReembolsable = True
       
       Else
       
       esReembolsable = False
       
       End If
   
   iComportamientos_buscarReembolsable = esReembolsable
    
End Function

'Sub iComportamientos_Reemplazar(datoAReemplazar As String, valorReemplazo As String)
'
'    Cells.Replace What:=datoAReemplazar, Replacement:=valorReemplazo, LookAt:=xlPart, SearchOrder _
'        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'End Sub


Sub iComportamientos_Reemplazar(datoAReemplazar As String, valorReemplazo As String)
  Dim celda As Range
  Dim valorCorregido As Double
    For Each celda In Selection
        valorCorregido = Replace(CStr(celda), datoAReemplazar, valorReemplazo)
    celda.Offset(0, 0) = valorCorregido
    Next celda
   
End Sub

Sub iComportamientos_Orden()
' Orden Macro
'ActiveWorkbook.Save
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    
End Sub



Sub iComportamientos_Ocultarcolumnasenblanco(columna As String)

Range(columna).Select
    Selection.ColumnWidth = 0
    
End Sub


