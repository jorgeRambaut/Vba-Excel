Attribute VB_Name = "prog_igu"
Sub Programa_Rate_Tiger()
Attribute Programa_Rate_Tiger.VB_ProcData.VB_Invoke_Func = "o\n14"

ActiveSheet.Range("B6").Select
Call Orden
Call columnasenblanco

ActiveSheet.Range("D6").Select
ActiveCell.Offset(1, 0).Select
Call DE_LETRA_A_NUMERO
Call PodaBookassist
Call PodaHotelbeds
Call Channel_id_duplicados

ActiveSheet.Range("S6").Select
ActiveCell.Offset(1, 0).Select
Call DE_LETRA_A_NUMERO
Call Niños

ActiveSheet.Range("E6").Select
ActiveCell.Offset(1, 0).Select
buscar_reembolsabe_o_standard

ActiveSheet.Range("G6").Select
ActiveCell.Offset(1, 0).Select
Call DE_LETRA_A_NUMERO

ActiveSheet.Range("M6").Select
ActiveCell.Offset(1, 0).Select
Call DE_LETRA_A_NUMERO

ActiveSheet.Range("N6").Select
ActiveCell.Offset(1, 0).Select
Call DE_LETRA_A_NUMERO

ActiveSheet.Range("Q6:R6").Select
Call DE_LETRA_A_NUMERO


ActiveSheet.Range("U6").Select
ActiveCell.Offset(1, 0).Select
Call DE_LETRA_A_NUMERO
Call IVA


ActiveSheet.Range("V6").Select
'ActiveCell.Offset(1, 0).Select
'Call buscar_reembolsabe_o_standard
Call Ajustar_a_texto

Call Descuentos_Hotelbeds
ActiveSheet.Range("B3").Select
'FIN
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

Private Sub DE_LETRA_A_NUMERO()
Dim celda As Range
Range(Selection, Selection.End(xlDown)).Select
'Selection.SpecialCells(xlCellTypeVisible).Select
For Each celda In Selection
celda = CStr(celda)
Next celda
 
End Sub

Private Sub PodaBookassist()

Dim celda As Range

Dim palabra As String

Range(Selection, Selection.End(xlDown)).Select

    palabra = "ARG"
    
        palabra = "*" & palabra & "*"
    
    For Each celda In Selection

        If celda.Value Like palabra Then
        
           celda = Right(celda, 7)
            
        End If
        
    Next celda
    
    End Sub
    
    Private Sub PodaHotelbeds()

Dim celda As Range

Dim palabra As String

Range(Selection, Selection.End(xlDown)).Select

    palabra = "249-"
    
        palabra = "*" & palabra & "*"
    
    For Each celda In Selection

        If celda.Value Like palabra Then
        
           celda = Right(celda, 6)
            
        End If
        
    Next celda
End Sub

Private Sub Channel_id_duplicados()

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

Private Sub Niños()
   
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


Private Sub buscar_reembolsabe_o_standard()

Dim celda As Range
Dim palabra As String
Dim Observaciion_Bookassist_standard As String
Dim Observacion_Bookassist_nrf As String
Dim Observacion_despegar_nrf As String
Dim Observacion_despegar_standard As String
Dim Observacion_expedia_nrf As String
Dim Observacion_expedia_standard As String
Dim observacion_DOTW As String
Dim Observacion_Globalia_nrf As String
Dim Observacion_Globalia_standard As String
Dim Observacion_Hotelbeds_standard As String
Dim Observacion_Hotelbeds_nrf As String
Dim observacion_NTincoming As String
Dim observacion_welcomebeds_nrf As String
Dim observacion_welcomebeds_standard As String
Dim observaciones_Bestday As String
Dim observaciones_almundo As String

observaciones_almundo = "Alojamiento Cta Cte Al Mundo - Extras Paga Pax" _
& vbNewLine & "MAT o TWIN NO ACLARA" _
& vbNewLine & "Condición de la reserva (Reembolsable - No reembolsable)" _
& vbNewLine & "Solicitudes especiales: Cama extra / Vista especial / Piso alto / etc." _
& vbNewLine & "Edad de los menores X/Y NO ACLARA  junto con gds"

observaciones_Bestday = "Alojamiento Cta Cte Hoteldo - Extras Paga Pax" _
& vbNewLine & "MAT o TWIN NO ACLARA" _
& vbNewLine & "Nacionalidad" _
& vbNewLine & "Solicitudes especiales: Cama extra / Vista especial / Piso alto / etc." _
& vbNewLine & "Edad de los menores X/Y NO ACLARA junto con gds"

observacion_welcomebeds_nrf = "Alojamiento Cta Cte Welcome beds - Extras Paga Pax" _
& vbNewLine & "MAT o TWIN NO ACLARA" _
& vbNewLine & "Condición de la reserva (No reembolsable)" _
& vbNewLine & "Nacionalidad" _
& vbNewLine & "Código de referencia / Localizador" _
& vbNewLine & "Solicitudes especiales: Cama extra / Vista especial / Piso alto / etc." _
& vbNewLine & "Edad de los menores X/Y NO ACLARA junto con gds"

observacion_welcomebeds_standard = "Alojamiento Cta Cte Welcome beds - Extras Paga Pax" _
& vbNewLine & "MAT o TWIN NO ACLARA" _
& vbNewLine & "Condición de la reserva (Reembolsable)" _
& vbNewLine & "Nacionalidad" _
& vbNewLine & "Código de referencia / Localizador" _
& vbNewLine & "Solicitudes especiales: Cama extra / Vista especial / Piso alto / etc." _
& vbNewLine & "Edad de los menores X/Y NO ACLARA junto con gds"

observacion_NTincoming = "Alojamiento TC virtual W2M - Extras Paga Pax" _
 & vbNewLine & "MAT o TWIN NO ACLARA" _
 & vbNewLine & "Solicitudes especiales: Cama extra / Vista especial / Piso alto / etc" _
& vbNewLine & "Edad de los menores X/Y NO ACLARA junto con gds"

Observacion_Hotelbeds_standard = "Alojamiento Cta Cte Hotelbeds - Extras Paga Pax" _
& vbNewLine & "Condición de la reserva (reembolsable)" _
& vbNewLine & "MAT o TWIN NO ACLARA" _
& vbNewLine & "Nacionalidad" _
& vbNewLine & "Solicitudes especiales: Cama extra / Vista especial / Piso alto / etc." _
& vbNewLine & "Edad de los menores X/Y NO ACLARA junto con gds Descuentos Aplicados"

Observacion_Hotelbeds_nrf = "Alojamiento Cta Cte Hotelbeds - Extras Paga Pax" _
& vbNewLine & "Condición de la reserva (No - reembolsable)" _
& vbNewLine & "MAT o TWIN NO ACLARA" _
& vbNewLine & "Nacionalidad" _
& vbNewLine & "Solicitudes especiales: Cama extra / Vista especial / Piso alto / etc." _
& vbNewLine & "Edad de los menores X/Y NO ACLARA junto con gds Descuentos Aplicados"

observacion_DOTW = " Alojamiento NO DEFINIDO - Extras Paga Pax - " _
& vbNewLine & "MAT o TWIN NO ACLARA" _
& vbNewLine & "Condición de la reserva (Reembolsable - No reembolsable)" _
& vbNewLine & "Solicitudes especiales: Cama extra / Vista especial / Piso alto / etc." _
& vbNewLine & "Edad de los menores X/Y NO ACLARA junto con gds"

Observaciion_Bookassist_standard = "Alojamiento y Extras Paga Pax" _
& vbNewLine & "MAT o TWIN NO ACLARA" _
& vbNewLine & "Condición de la reserva (Reembolsable)" _
& vbNewLine & "Solicitudes especiales: WEBLOI Free Upgrade (subject to availability) Free Late Check out (2 hours late)" _
& vbNewLine & "Edad de los menores X/Y NO ACLARA junto con gds"


Observacion_Bookassist_nrf = "Alojamiento y Extras Paga Pax" _
& vbNewLine & "MAT o TWIN NO ACLARA" _
& vbNewLine & "Condición de la reserva (No Reembolsable)" _
& vbNewLine & "Solicitudes especiales: WEBLOI Free Upgrade (subject to availability) Free Late Check out (2 hours late)" _
& vbNewLine & "Edad de los menores X/Y NO ACLARA junto con gds"


Observaciion_Booking_standard = "Alojamiento y Extras Paga Pax" _
& vbNewLine & "MAT o TWIN NO ACLARA" _
& vbNewLine & "Condición de la reserva (Reembolsable)" _
& vbNewLine & "Solicitudes especiales: Genius: Free upgrade ( IMPORTANTE : Sujeto A Disponibilidad) " _
& vbNewLine & "Edad de los menores X/Y NO ACLARA junto con gds"


Observacion_Booking_nrf = "Alojamiento y Extras Paga Pax" _
& vbNewLine & "MAT o TWIN NO ACLARA" _
& vbNewLine & "Condición de la reserva (No Reembolsable)" _
& vbNewLine & "Solicitudes especiales: Genius: Free upgrade ( IMPORTANTE : Sujeto A Disponibilidad) " _
& vbNewLine & "Edad de los menores X/Y NO ACLARA junto con gds"


Observacion_despegar_nrf = "Alojamiento Cta Cte Despegar - Extras Paga Pax" _
& vbNewLine & "Alojamiento y Extras Paga Pax ELEGIR LA OPCION QUE CORRESPONDE" _
& vbNewLine & "MAT o TWIN NO ACLARA" _
& vbNewLine & "Condición de la reserva (No Reembolsable)" _
& vbNewLine & "Edad de los menores X/Y NO ACLARA junto con gds"


Observacion_despegar_standard = "Alojamiento Cta Cte Despegar - Extras Paga Pax" _
& vbNewLine & "Alojamiento y Extras Paga Pax ELEGIR LA OPCION QUE CORRESPONDE" _
& vbNewLine & "MAT o TWIN NO ACLARA" _
& "Condición de la reserva ( Reembolsable)" _
& vbNewLine & "Edad de los menores X/Y NO ACLARA junto con gds"



Observacion_expedia_standard = "A CARGO DE EXPEDIA (Expedia Collects Payment)" _
& vbNewLine & "Alojamiento TC virtual Expedia - Extras Paga Pax" _
& vbNewLine & "MAT o TWIN NO ACLARA" _
& vbNewLine & "Condición de la reserva (Reembolsable)" _
& vbNewLine & "Solicitudes especiales: Vip - Premium VIP beneficios: 1 Free beverage per person for 2 (once per stay)" _
& vbNewLine & "Edad de los menores X/Y NO ACLARA junto con gds" _
& vbNewLine & "--------------------------------------------------" _
& vbNewLine & "A CARGO DEL PAX (Hotel Collects Payment)" _
& vbNewLine & "Alojamiento y Extras Paga Pax" _
& vbNewLine & "MAT o TWIN NO ACLARA" _
& vbNewLine & "Condición de la reserva (Reembolsable)" _
& vbNewLine & "Solicitudes especiales: Vip - Premium VIP beneficios: 1 Free beverage per person for 2 (once per stay)" _
& vbNewLine & "Edad de los menores X/Y NO ACLARA junto con gds"

Observacion_expedia_nrf = "A CARGO DE EXPEDIA (Expedia Collects Payment)" _
& vbNewLine & "Alojamiento TC virtual Expedia - Extras Paga Pax" _
& vbNewLine & "MAT o TWIN NO ACLARA" _
& vbNewLine & "Condición de la reserva (No Reembolsable)" _
& vbNewLine & "Solicitudes especiales: Vip - Premium VIP beneficios: 1 Free beverage per person for 2 (once per stay)" _
& vbNewLine & "Edad de los menores X/Y NO ACLARA junto con gds" _
& vbNewLine & "--------------------------------------------------" _
& vbNewLine & "A CARGO DEL PAX (Hotel Collects Payment)" _
& vbNewLine & "Alojamiento y Extras Paga Pax" _
& vbNewLine & "MAT o TWIN NO ACLARA" _
& vbNewLine & "Condición de la reserva (No Reembolsable)" _
& vbNewLine & "Solicitudes especiales: Vip - Premium VIP beneficios: 1 Free beverage per person for 2 (once per stay)" _
& vbNewLine & "Edad de los menores X/Y NO ACLARA junto con gds"

    Range(Selection, Selection.End(xlDown)).Select
    
    For Each celda In Selection

       If celda = "Bookassist" And _
       celda.Offset(0, 10).Value Like "*Non Refundable*" Then
       celda.Offset(0, 17).Value = Observacion_Bookassist_nrf
       ElseIf celda = "Bookassist" And _
       celda.Offset(0, 10).Value <> "*Non Refundable*" Then
       celda.Offset(0, 17).Value = Observaciion_Bookassist_standard
              
       ElseIf celda = "Booking" And _
       celda.Offset(0, 10).Value Like "*Non Refundable*" Then
       celda.Offset(0, 17).Value = Observacion_Booking_nrf
       ElseIf celda = "Booking" And _
       celda.Offset(0, 10).Value <> "*Non Refundable*" Then
       celda.Offset(0, 17).Value = Observaciion_Booking_standard
              
       ElseIf celda = "Despegar" And _
       celda.Offset(0, 10).Value Like "*MAYORISTA*" Then
       celda.Offset(0, 17).Value = Observacion_despegar_standard
       ElseIf celda = "Despegar" And _
       celda.Offset(0, 10).Value <> "*MAYORISTA*" Then
       celda.Offset(0, 17).Value = Observacion_despegar_nrf
       
       ElseIf celda = "Despegar.com" And _
       celda.Offset(0, 10).Value Like "*MAYORISTA*" Then
       celda.Offset(0, 17).Value = Observacion_despegar_standard
       ElseIf celda = "Despegar.com" And _
       celda.Offset(0, 10).Value <> "*MAYORISTA*" Then
       celda.Offset(0, 17).Value = Observacion_despegar_nrf
            
       ElseIf celda = "DOTW" Then
       celda.Offset(0, 17).Value = observacion_DOTW
 
       ElseIf celda = "Expedia" And _
       celda.Offset(0, 10).Value Like "*Non Refundable*" Then
       celda.Offset(0, 17).Value = Observacion_expedia_nrf
       ElseIf celda = "Expedia" And _
       celda.Offset(0, 10).Value <> "*Non Refundable*" Then
       celda.Offset(0, 17).Value = Observacion_expedia_standard
              
       
       ElseIf celda = "Hotelbeds" And _
       celda.Offset(0, 10).Value Like "*NRF*" Then
       celda.Offset(0, 17).Value = Observacion_Hotelbeds_nrf
       ElseIf celda = "Hotelbeds" And _
       celda.Offset(0, 10).Value <> "*NRF*" Then
       celda.Offset(0, 17).Value = Observacion_Hotelbeds_standard
          
       ElseIf celda = "NTIncoming" Then
       celda.Offset(0, 17).Value = observacion_NTincoming
       
       ElseIf celda = "welcomebeds.com" And _
       celda.Offset(0, 10).Value Like "*BAR*" Then
       celda.Offset(0, 17).Value = observacion_welcomebeds_nrf
       ElseIf celda = "welcomebeds.com" And _
       celda.Offset(0, 10).Value <> "*BAR*" Then
       celda.Offset(0, 17).Value = observacion_welcomebeds_standard
       
       ElseIf celda = "Best Day" Then
       celda.Offset(0, 17).Value = observaciones_Bestday
              
       ElseIf celda = "almundo.com" Then
       celda.Offset(0, 17).Value = observaciones_almundo
       
       
     End If
        
    Next celda
    
End Sub

Private Sub IVA()
Dim celda As Range
'ActiveCell.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.SpecialCells(xlCellTypeVisible).Select

For Each celda In Selection
      celda.Value = celda.Value * 1.21
Next celda

End Sub

Private Sub Descuentos_Hotelbeds()

Dim descuento_del_20 As Double
Dim descuento_del_5 As Double
Dim descuento_del_10 As Double
Dim Calculo_descuentos As Double
Dim rate_plan As Range

'automatico
ActiveSheet.Range("P6").Select
ActiveCell.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Select
'Selection.SpecialCells(xlCellTypeVisible).Select
' -----------------CALCULA DESCUENTOS HOTELBDES----------------------------------
For Each rate_plan In Selection

   If rate_plan.Value Like "*17073*" Then
        descuento_del_20 = 20
        descuento_del_10 = 10
        descuento_del_20 = (rate_plan.Offset(0, 5).Value * descuento_del_20) / 100
        descuento_del_10 = ((rate_plan.Offset(0, 5).Value - descuento_del_20) * descuento_del_10) / 100
        Calculo_descuentos = ((rate_plan.Offset(0, 5).Value - descuento_del_20) - descuento_del_10)
        rate_plan.Offset(0, 5).Value = Calculo_descuentos
               
        ElseIf rate_plan.Value Like "*17074*" Then
        descuento_del_20 = 20
        descuento_del_10 = 10
        descuento_del_20 = (rate_plan.Offset(0, 5).Value * descuento_del_20) / 100
        descuento_del_10 = ((rate_plan.Offset(0, 5).Value - descuento_del_20) * descuento_del_10) / 100
        Calculo_descuentos = ((rate_plan.Offset(0, 5).Value - descuento_del_20) - descuento_del_10)
        rate_plan.Offset(0, 5).Value = Calculo_descuentos
        
            ElseIf rate_plan.Value Like "*17173*" Then
            descuento_del_20 = 20
            descuento_del_10 = 10
            descuento_del_20 = (rate_plan.Offset(0, 5).Value * descuento_del_20) / 100
            descuento_del_10 = ((rate_plan.Offset(0, 5).Value - descuento_del_20) * descuento_del_10) / 100
            Calculo_descuentos = ((rate_plan.Offset(0, 5).Value - descuento_del_20) - descuento_del_10)
            rate_plan.Offset(0, 5).Value = Calculo_descuentos
        
                ElseIf rate_plan.Value Like "*17177*" Then
                descuento_del_5 = 5
                descuento_del_10 = 10
                descuento_del_5 = (rate_plan.Offset(0, 5).Value * descuento_del_5) / 100
                descuento_del_10 = ((rate_plan.Offset(0, 5).Value - descuento_del_5) * descuento_del_10) / 100
                Calculo_descuentos = ((rate_plan.Offset(0, 5).Value - descuento_del_5) - descuento_del_10)
                rate_plan.Offset(0, 5).Value = Calculo_descuentos
            
                    ElseIf rate_plan.Value Like "*10812*" Then
                    descuento_del_10 = 10
                    descuento_del_10 = (rate_plan.Offset(0, 5).Value * descuento_del_10) / 100
                    Calculo_descuentos = (rate_plan.Offset(0, 5).Value - descuento_del_10)
                    rate_plan.Offset(0, 5).Value = Calculo_descuentos
                    
                       
        End If
    
    Next

End Sub
