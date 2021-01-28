Attribute VB_Name = "Refactorizado"

Sub Programa_Rate_Tiger()
Attribute Programa_Rate_Tiger.VB_ProcData.VB_Invoke_Func = "i\n14"
Dim genius As String
Dim expediavip As String
Dim webloi As String
Dim solapa As String

solapa = InputBox("Nombre de Hoja Excel")
webloi = "WEBLOI: Free Upgrade (subject to availability) Free Late Check out (2 hours late)"
expediavip = "Vip - Premium VIP beneficios: 1 Free beverage per person for 2 (once per stay) ECI sujeto a disponibilidad LCO confirmado hasta las 14hs Upgrade sujeto a disponibilidad"
genius = "Genius: Free upgrade ( IMPORTANTE : Sujeto A Disponibilidad) "

ActiveSheet.Range("B6").Select
ActiveSheet.Range("E6").Value = "Extranet"
ActiveSheet.Range("V6").Value = "Observaciones"

Call Orden ' refactorizar
Call columnasenblanco 'refactorizar
Call BuscarTitulo("Channel ID", solapa, "B6", "W6")
Call ConvertirValores("Letra")
Call CorregirGds("ARG", 7)
Call CorregirGds("249-", 6)
Call BuscarSiHayMasDeUnaReservaConMismoGds
Call BuscarTitulo("Children", solapa, "B6", "W6")
Call ConvertirValores("Numero")
Call BuscarSiHayMenoresEnLaReserva
Call BuscarTitulo("Extranet", solapa, "B6", "W6")

'---Cargo observaciones y Politicas -----------------------------------------------

Call buscar_reembolsable_o_standard_Paga_Pax("Booking", "Non Refundable", genius)
Call buscar_reembolsable_o_standard_Paga_Pax("Bookassist", "Non Refundable", webloi)
Call ObservacionExpedia("Expedia", "Non Refundable", expediavip)
Call ObservacionDespegar("Despegar.com", "PROMOS")
Call ObservacionDespegar("Despegar", "PROMOS")
Call ObservacionNtincoming("NTIncoming")
Call ObservacionCuentaCorriente("almundo.com", "Reembolsable-No Reembolsable")
Call ObservacionCuentaCorriente("Best Day", "Reembolsable-No Reembolsable")
'Call ObservacionCuentaCorriente("welcomebeds.com", "Reembolsable-No Reembolsable")
Call Observacion_Welcomebeds("welcomebeds.com", "Reembolsable-No Reembolsable")
Call ObservacionCuentaCorriente("Hotelbeds", "Reembolsable-No Reembolsable")

'------------------------------------------------------------------------------------

Call BuscarTitulo("Booked On", solapa, "B6", "W6")
Call ConvertirValores("Letra")
Call BuscarTitulo("Check-in", solapa, "B6", "W6")
Call ConvertirValores("Letra")
Call BuscarTitulo("Checkout", solapa, "B6", "W6")
Call ConvertirValores("Letra")
Call BuscarTitulo("Rooms", solapa, "B6", "W6")
Call ConvertirValores("Numero")
Call BuscarTitulo("Adults", solapa, "B6", "W6")
Call ConvertirValores("Numero")
Call BuscarTitulo("iva incl", solapa, "B6", "W6")
Call Reemplazar(".", ",")
Call SumaIva
Call BuscarTitulo("Special Request", solapa, "B6", "W6")
Call Ajustar_a_texto
Call BuscarTitulo("Room Type", solapa, "B6", "W6")
Call Descuentos_Hotelbeds("17073", 20, 10)
Call Descuentos_Hotelbeds("17074", 20, 10)
Call Descuentos_Hotelbeds("17173", 20, 10)
Call Descuentos_Hotelbeds("17177", descuento10porciento:=10, descuento5porciento:=5)
Call Descuentos_Hotelbeds("10812", descuento10porciento:=10)


ActiveSheet.Range("B3").Select

'FIN


MsgBox ("Procesos Finalizados")


End Sub

Sub BuscarTitulo(DatoAbuscar As String, NombreDeSolapaAbuscar As String, RangoInicio As String, RangoFin As String, Optional libro As String)
    Dim tituloEncontrado As Range
    Selection.SpecialCells(xlCellTypeVisible).Select
    
        
            With Worksheets(NombreDeSolapaAbuscar).Range(RangoInicio & ":" & RangoFin)
                
                    Set tituloEncontrado = .Find(DatoAbuscar, LookIn:=xlValues)
                    
                        If Not tituloEncontrado Is Nothing Then
                            tituloEncontrado.Offset(1, 0).Select
                            Range(Selection, Selection.End(xlDown)).Select
                        End If
            End With
     
       
        
    
End Sub


Sub prueba()

Call BuscarTitulo("Status", "Detail-Booked", "B6", "W6")



End Sub
Private Sub Ajustar_a_texto()

    Range("C6").Select
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
Private Sub Orden()
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

Private Sub columnasenblanco()

    Range("B6").Select
    Selection.ColumnWidth = 0
           
    Range("H6,J6:L6").Select
    Selection.ColumnWidth = 0
    
    Range("P6").Select
    Selection.ColumnWidth = 0
    
    Range("Q6,R6:S6").Select
    Selection.ColumnWidth = 3
    
    Range("T6").Select
    Selection.ColumnWidth = 0
                
    Range("U6").Select
    Selection.ColumnWidth = 8
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
    
     Range("W6").Select
    Selection.ColumnWidth = 0
   
End Sub

Sub ConvertirValores(valor As String)
Dim celda As Range
Range(Selection, Selection.End(xlDown)).Select
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
End Select
End Sub

Private Sub CorregirGds(gds As String, cantidad_a_quitar As Byte)

Dim celda As Range

Range(Selection, Selection.End(xlDown)).Select

    
    For Each celda In Selection

        If celda.Value Like "*" & gds & "*" Then
        
           celda = Right(celda, cantidad_a_quitar)
            
        End If
        
    Next celda
    
    End Sub
    
'    Private Sub PodaHotelbeds()
'
'Dim celda As Range
'
'Dim palabra As String
'
'Range(Selection, Selection.End(xlDown)).Select
'
'    palabra = "249-"
'
'        palabra = "*" & palabra & "*"
'
'    For Each celda In Selection
'
'        If celda.Value Like palabra Then
'
'           celda = Right(celda, 6)
'
'        End If
'
'    Next celda
'End Sub

Private Sub BuscarSiHayMasDeUnaReservaConMismoGds()

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

Private Sub BuscarSiHayMenoresEnLaReserva()
   
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

Function Observacion_Paga_Pax(channel As String, condicion As String, menores As Integer, beneficios As String)
Observacion_Paga_Pax = channel & " Alojamiento y Extras Paga Pax" _
& vbNewLine & "MAT o TWIN NO ACLARA" _
& vbNewLine & "Condición de la reserva (" & condicion & ")" _
& vbNewLine & "Solicitudes especiales: " _
& vbNewLine & "Menores = " & menores & " NO ACLARA Edad de los menores" _
& vbNewLine & beneficios
End Function

Function Observacion_Cuenta_Corriente(channel As String, condicion As String, menores As Integer)

Observacion_Cuenta_Corriente = " Alojamiento Cta Cte " & channel _
& vbNewLine & "MAT o TWIN NO ACLARA" _
& vbNewLine & "Condición de la reserva (" & condicion & ")" _
& vbNewLine & "Solicitudes especiales: " _
& vbNewLine & "Menores = " & menores _
& vbNewLine & "NO ACLARA Edad de los menores" _

End Function

Function Observacion_Welcomebeds_Tc(channel As String, condicion As String, menores As Integer)

Observacion_Cobro_Tc = " Alojamiento Cobrar de La Tc " & channel _
& vbNewLine & "MAT o TWIN NO ACLARA" _
& vbNewLine & "Condición de la reserva (" & condicion & ")" _
& vbNewLine & "Solicitudes especiales: " _
& vbNewLine & "Menores = " & menores _
& vbNewLine & "NO ACLARA Edad de los menores" _
& vbNewLine & "TC: se activará el día del check in y tendrán hasta 15 días después del check out para cobrarla."


End Function

Function Observacion_Expedia(channel As String, condicion As String, menores As Integer, beneficios As String)

Observacion_Expedia = "A CARGO DEL PAX (Hotel Collects Payment)" _
& vbNewLine & "A CARGO DE " & channel & "(" & channel & " Collects Payment)" _
& vbNewLine & "Elegir el que corresponde" _
& vbNewLine & "MAT o TWIN NO ACLARA" _
& vbNewLine & "Condición de la reserva (" & condicion & ")" _
& vbNewLine & "Solicitudes especiales: " _
& vbNewLine & "Menores = " & menores _
& vbNewLine & "NO ACLARA Edad de los menores" _
& vbNewLine & beneficios

End Function

Function Observacion_Despegar(channel As String, condicion As String, menores As Integer)

Observacion_Despegar = "A CARGO DEL PAX (Hotel Collects Payment)" _
& vbNewLine & "A CARGO DE " & channel & "(" & channel & " Collects Payment)" _
& vbNewLine & "Elegir el que corresponde" _
& vbNewLine & "MAT o TWIN NO ACLARA" _
& vbNewLine & "Condición de la reserva (" & condicion & ")" _
& vbNewLine & "Solicitudes especiales: " _
& vbNewLine & "Menores = " & menores _
& vbNewLine & "NO ACLARA Edad de los menores"

End Function
Function Observacion_Ntincoming(channel As String, menores As Integer)

Observacion_Ntincoming = "Alojamiento TC virtual W2M (" & channel & ") - Extras Paga Pax" _
 & vbNewLine & "MAT o TWIN NO ACLARA" _
 & vbNewLine & "Solicitudes especiales:" _
 & vbNewLine & "Menores = " & menores _
 & vbNewLine & "NO ACLARA Edad de los menores"

End Function


Private Sub buscar_reembolsable_o_standard_Paga_Pax(channel As String, condicion As String, beneficio As String)

Dim extranet As Range
Dim menores As Integer
Dim roomtype As String
Dim observacion As String
Dim reembolsable As String
Dim gds As String
Dim filaInferior As String
Dim filaSuperior As String

reembolsable = "Reembolsable"

    
    
    Range(Selection, Selection.End(xlDown)).Select
    
    For Each extranet In Selection
        menores = extranet.Offset(0, 14).Value
        roomtype = extranet.Offset(0, 10).Value
        gds = extranet.Offset(0, -1).Value
        filaInferior = extranet.Offset(1, -1).Value
        filaSuperior = extranet.Offset(-1, -1).Value
        
       If extranet = channel And _
         roomtype Like "*" & condicion & "*" Then
            extranet.Offset(0, 17).Value = Observacion_Paga_Pax(channel, condicion, menores, beneficio)
            
       ElseIf extranet = channel And roomtype <> "*" & condicion & "*" Then
       extranet.Offset(0, 17).Value = Observacion_Paga_Pax(channel, reembolsable, menores, beneficio)
       End If
       
       If extranet = channel And (gds = filaInferior Or gds = filaSuperior) _
       And roomtype Like "*" & condicion & "*" Then
       extranet.Offset(0, 17).Value = Observacion_Paga_Pax(channel, condicion, menores, beneficio) + " Junto Con Gds " & gds
       
       ElseIf extranet = channel And (gds = filaInferior Or gds = filaSuperior) _
       And roomtype <> "*" & condicion & "*" Then
       extranet.Offset(0, 17).Value = Observacion_Paga_Pax(channel, reembolsable, menores, beneficio) + " Junto Con Gds " & gds
            
       
       End If
       
       
    Next extranet
    
End Sub

Private Sub ObservacionExpedia(channel As String, condicion As String, beneficio As String)

Dim extranet As Range
Dim menores As Integer
Dim roomtype As String
Dim observacion As String
Dim reembolsable As String
Dim gds As String
Dim filaInferior As String
Dim filaSuperior As String

reembolsable = "Reembolsable"

    
    
    Range(Selection, Selection.End(xlDown)).Select
    
    For Each extranet In Selection
        menores = extranet.Offset(0, 14).Value
        roomtype = extranet.Offset(0, 10).Value
        gds = extranet.Offset(0, -1).Value
        filaInferior = extranet.Offset(1, -1).Value
        filaSuperior = extranet.Offset(-1, -1).Value
        
       If extranet = channel And _
         roomtype Like "*" & condicion & "*" Then
            extranet.Offset(0, 17).Value = Observacion_Expedia(channel, condicion, menores, beneficio)
              
       ElseIf extranet = channel And _
       roomtype <> "*" & condicion & "*" Then
       extranet.Offset(0, 17).Value = Observacion_Expedia(channel, reembolsable, menores, beneficio)
       End If
       
       If extranet = channel And (gds = filaInferior Or gds = filaSuperior) _
       And roomtype Like "*" & condicion & "*" Then
       extranet.Offset(0, 17).Value = Observacion_Expedia(channel, condicion, menores, beneficio) + " Junto Con Gds " & gds
             
       ElseIf extranet = channel And (gds = filaInferior Or gds = filaSuperior) _
       And roomtype <> "*" & condicion & "*" Then
       extranet.Offset(0, 17).Value = Observacion_Expedia(channel, reembolsable, menores, beneficio) + " Junto Con Gds " & gds
       End If
        
    Next extranet
    
End Sub

Private Sub ObservacionDespegar(channel As String, condicion As String)

Dim extranet As Range
Dim menores As Integer
Dim roomtype As String
Dim observacion As String
Dim reembolsable As String
Dim gds As String
Dim filaInferior As String
Dim filaSuperior As String

reembolsable = "Reembolsable"

    
    
    Range(Selection, Selection.End(xlDown)).Select
    
    For Each extranet In Selection
        menores = extranet.Offset(0, 14).Value
        roomtype = extranet.Offset(0, 10).Value
        gds = extranet.Offset(0, -1).Value
        filaInferior = extranet.Offset(1, -1).Value
        filaSuperior = extranet.Offset(-1, -1).Value
        
       If extranet = channel And _
         roomtype Like "*" & condicion & "*" Then
            extranet.Offset(0, 17).Value = Observacion_Despegar(channel, condicion, menores)
              
       ElseIf extranet = channel And _
       roomtype <> "*" & condicion & "*" Then
       extranet.Offset(0, 17).Value = Observacion_Despegar(channel, reembolsable, menores)
       End If
       
       If extranet = channel And (gds = filaInferior Or gds = filaSuperior) _
       And roomtype Like "*" & condicion & "*" Then
       extranet.Offset(0, 17).Value = Observacion_Despegar(channel, condicion, menores) + " Junto Con Gds " & gds
             
       ElseIf (extranet = channel) And (gds = filaInferior Or gds = filaSuperior) _
       And roomtype <> "*" & condicion & "*" Then
       extranet.Offset(0, 17).Value = Observacion_Despegar(channel, reembolsable, menores) + " Junto Con Gds " & gds
       End If
        
    Next extranet
    
End Sub

Private Sub ObservacionNtincoming(channel As String)

Dim extranet As Range
Dim menores As Integer
Dim gds As String
Dim filaInferior As String
Dim filaSuperior As String
    
    
    Range(Selection, Selection.End(xlDown)).Select
    
    For Each extranet In Selection
        menores = extranet.Offset(0, 14).Value
        gds = extranet.Offset(0, -1).Value
        filaInferior = extranet.Offset(1, -1).Value
        filaSuperior = extranet.Offset(-1, -1).Value
        
        If extranet = channel Then
            extranet.Offset(0, 17).Value = Observacion_Ntincoming(channel, menores)
        End If
        
        If extranet = channel And (gds = filaInferior Or gds = filaSuperior) Then
        extranet.Offset(0, 17).Value = Observacion_Ntincoming(channel, menores) + " Junto Con Gds " & gds
        End If
        
    Next extranet
    
End Sub


Private Sub ObservacionCuentaCorriente(channel As String, condicion As String)

Dim extranet As Range
Dim menores As Integer
Dim gds As String
Dim filaInferior As String
Dim filaSuperior As String
    
    
    Range(Selection, Selection.End(xlDown)).Select
    
    For Each extranet In Selection
        menores = extranet.Offset(0, 14).Value
        gds = extranet.Offset(0, -1).Value
        filaInferior = extranet.Offset(1, -1).Value
        filaSuperior = extranet.Offset(-1, -1).Value
        
        If extranet = channel Then
            extranet.Offset(0, 17).Value = Observacion_Cuenta_Corriente(channel, condicion, menores)
        End If
        
        If extranet = channel And (gds = filaInferior Or gds = filaSuperior) Then
        extranet.Offset(0, 17).Value = Observacion_Cuenta_Corriente(channel, condicion, menores) + " Junto Con Gds " & gds
        End If
        
    Next extranet
    
End Sub

Private Sub Observacion_Welcomebeds(channel As String, condicion As String)

Dim extranet As Range
Dim menores As Integer
Dim gds As String
Dim filaInferior As String
Dim filaSuperior As String
    
    
    Range(Selection, Selection.End(xlDown)).Select
    
    For Each extranet In Selection
        menores = extranet.Offset(0, 14).Value
        gds = extranet.Offset(0, -1).Value
        filaInferior = extranet.Offset(1, -1).Value
        filaSuperior = extranet.Offset(-1, -1).Value
        
        If extranet = channel Then
            extranet.Offset(0, 17).Value = Observacion_Welcomebeds_Tc(channel, condicion, menores)
        End If
        
        If extranet = channel And (gds = filaInferior Or gds = filaSuperior) Then
        extranet.Offset(0, 17).Value = Observacion_Welcomebeds_Tc(channel, condicion, menores) + " Junto Con Gds " & gds
        End If
        
    Next extranet
    
End Sub



Private Sub SumaIva()
Dim celda As Range
'ActiveCell.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.SpecialCells(xlCellTypeVisible).Select

For Each celda In Selection
      celda.Value = celda.Value * 1.21
Next celda

End Sub

Private Sub Descuentos_Hotelbeds(ratePlan As String, Optional descuento20porciento _
As Double, Optional descuento10porciento As Double, Optional descuento5porciento As Double)

Dim Calculo_descuentos As Double
Dim rate_plan As Range

'descuento20porciento = descuento20porciento / 100
'descuento10porciento = descuento10porciento / 100
'descuento5porciento = descuento5porciento / 100

' -----------------CALCULA DESCUENTOS HOTELBDES----------------------------------
    For Each rate_plan In Selection
    
           If rate_plan.Value Like "*" & ratePlan & "*" Then
                
                descuento20porciento = rate_plan.Offset(0, 6).Value * (descuento20porciento / 100)
                descuento10porciento = ((rate_plan.Offset(0, 6).Value - descuento20porciento) * (descuento10porciento / 100))
                descuento5porciento = (((rate_plan.Offset(0, 6).Value - descuento20porciento) - descuento10porciento) * (descuento5porciento / 100))
                Calculo_descuentos = ((rate_plan.Offset(0, 6).Value - descuento20porciento) - descuento10porciento) - descuento5porciento
                rate_plan.Offset(0, 6).Value = Calculo_descuentos
           
           
                        '-------Resetear % --------------
                        If descuento20porciento <> 0 Then
                        
                            descuento20porciento = 20#
                            
                        End If
                        
                        If descuento10porciento <> 0 Then
                        
                            descuento10porciento = 10#
                            
                        End If
                        
                        If descuento5porciento <> 0 Then
                        
                            descuento5porciento = 5#
                            
                        End If
                    '-------Fin % ----------------------
                       
          End If
                        
    Next

End Sub

Sub Reemplazar(dato_a_reemplazar As String, valor_reemplazo As String)

    Cells.Replace What:=dato_a_reemplazar, Replacement:=valor_reemplazo, LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
End Sub


Function obtenercolor1(celda As Range) As String
Dim sColor As String
sColor = Right("000000" & Hex(celda.Interior.Color), 6)
obtenercolor1 = Right(sColor, 2) & Mid(sColor, 3, 2) & Left(sColor, 2)
End Function
