VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iComportamientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Sub seleccionarCeldasDescendente()

  Range(Selection, Selection.End(xlDown)).Select

End Sub

Function obtenerParte(texto As String, cantidad As Integer) As String
    Dim textoobtenido As String
    textoobtenido = Left(texto, cantidad)
End Function


Sub BuscarTitulo(DatoAbuscar As String, NombreDeSolapaAbuscar As String, RangoInicio As String, RangoFin As String, Optional libro As String)
    Dim tituloEncontrado As Range
      Selection.SpecialCells(xlCellTypeVisible).Select
        
            With Worksheets(NombreDeSolapaAbuscar).Range(RangoInicio & ":" & RangoFin)
            
                    Set tituloEncontrado = .Find(DatoAbuscar, LookIn:=xlValues)
                    
                        If Not tituloEncontrado Is Nothing Then
                            tituloEncontrado.Offset(1, 0).Select
                            tituloEncontrado(Selection, Selection.End(xlDown)).Select
                            'Range(Selection, Selection.End(xlDown)).Select
                        End If
            End With
    
End Sub
Sub movercelda(columna As Integer, fila As Integer)
     ActiveSheet.Offset(fila, columna).Select
'    Dim movercelda As Range
'    iComportamientos_moverCelda = movercelda.Offset(Columna, fila)
    
End Sub

'Function observacion(ota As String, beneficios As String, condicion_reserva As String, forma_de_pago As String)
'
'    observacion = ota & " " & condicion_reserva & " " & forma_de_pago & " " & beneficios
'
'End Function

Function pasarADolar(tarifa As Double, tipoDeCambio As Double) As Double

    pasarADolar = tarifa / tipoDeCambio

End Function

Function pasarAPesos(tipoDeCambio As Double, tarifa As Double) As Double

  pasarAPesos = tipoDeCambio * tarifa

End Function

Sub ConvertirValores(valor As String)
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

Sub CorregirGds(gds As String, cantidad_a_quitar As Byte)

Dim celda As Range

'Range(Selection, Selection.End(xlDown)).Select
    
    For Each celda In Selection

        If celda.Value Like "*" & gds & "*" Then
        
           celda = Right(celda, cantidad_a_quitar)
            
        End If
        
    Next celda
    
    End Sub
    
Sub BuscarDuplicadosYResaltar()
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

Sub BuscarSiHayMenoresEnLaReservaYResaltarlos()
   
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


Function buscarReembolsable(reembolsable As String) As Boolean

Dim condicion As String
Dim esReembolsable As Boolean

condicion = "Reembolsable"
    
        
       If condicion = reembolsable Then
           
       esReembolsable = True
       
       Else
       
       esReembolsable = False
       
       End If
   
   buscarReembolsable = esReembolsable
    
End Function

'Sub Reemplazar(datoAReemplazar As String, valorReemplazo As String)
'  Dim celda As Range
'  For Each celda In Selection
'    Cells.Replace What:=datoAReemplazar, Replacement:=valorReemplazo, LookAt:=xlPart, SearchOrder _
'        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'  Next celda
'End Sub

Sub Reemplazar(datoAReemplazar As String, valorReemplazo As String)
      Dim celda As Range
      
    For Each celda In Selection
        Reemplazar = Replace(CStr(celda), datoAReemplazar, valorReemplazo)
    Next celda
    
End Sub

Sub Orden()
    
End Sub
Private Sub Ocultarcolumnasenblanco(columna As String)

    
   
End Sub

Sub Ajustaratexto(columna As String)

    
End Sub
