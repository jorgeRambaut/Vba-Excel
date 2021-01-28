VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsObservacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements iObservacion



Function iObservacion_observacion(tipoDealojamiento As String, Ota As String, condicion_reserva As String) As String
   
    iObservacion_observacion = tipoDealojamiento & " " & Ota _
    & vbNewLine & "MAT o TWIN NO ACLARA " & "Condición de la reserva :" & condicion_reserva _
    & vbNewLine & "Solicitudes Especiales : "

End Function

Sub iObservacion_crearobservacion(TipoDeReembolso As String)
Dim roomtype As Range
Dim nombreota As String
Dim observacionReserva As String
Dim observacion As New ClsObservacion

    For Each roomtype In Selection
        nombreota = roomtype.Offset(0, -10)

        Select Case nombreota

        Case "Expedia"
            If roomtype Like "*" & TipoDeReembolso & "*" Then
             observacionReserva = observacion.iObservacion_observacion("A CARGO DEL PAX (Hotel Collects Payment) / A CARGO DE Expedia(Expedia Collects Payment) ", nombreota, " No Reembolsable")
             roomtype.Offset(0, 7) = observacionReserva + vbNewLine & "Vip - Premium VIP beneficios: 1 Free beverage per person for 2 (once per stay) ECI sujeto a disponibilidad LCO confirmado hasta las 14hs Upgrade sujeto a disponibilidad"
            ElseIf roomtype <> "*" & TipoDeReembolso & "*" Then
            observacionReserva = observacion.iObservacion_observacion("A CARGO DEL PAX (Hotel Collects Payment) / A CARGO DE Expedia(Expedia Collects Payment) ", nombreota, " Reembolsable")
            roomtype.Offset(0, 7) = observacionReserva + vbNewLine & "Vip - Premium VIP beneficios: 1 Free beverage per person for 2 (once per stay) ECI sujeto a disponibilidad LCO confirmado hasta las 14hs Upgrade sujeto a disponibilidad"
            End If
            
        Case "Booking"
            If roomtype Like "*" & TipoDeReembolso & "*" Then
             observacionReserva = observacion.iObservacion_observacion("Alojamiento y Extras Paga Pax ", nombreota, " No Reembolsable")
             roomtype.Offset(0, 7) = observacionReserva + vbNewLine & "Genius: Free upgrade ( IMPORTANTE : Sujeto A Disponibilidad)"
            ElseIf roomtype <> "*" & TipoDeReembolso & "*" Then
             observacionReserva = observacion.iObservacion_observacion("Alojamiento y Extras Paga Pax ", nombreota, " Reembolsable")
             roomtype.Offset(0, 7) = observacionReserva + vbNewLine & "Genius: Free upgrade ( IMPORTANTE : Sujeto A Disponibilidad)"
            End If
            
        Case "Bookassist"
            If roomtype Like "*" & TipoDeReembolso & "*" Then
             observacionReserva = observacion.iObservacion_observacion("Alojamiento y Extras Paga Pax ", nombreota, " No Reembolsable")
             roomtype.Offset(0, 7) = observacionReserva
            ElseIf roomtype <> "*" & TipoDeReembolso & "*" Then
             observacionReserva = observacion.iObservacion_observacion("Alojamiento y Extras Paga Pax ", nombreota, " Reembolsable")
             roomtype.Offset(0, 7) = observacionReserva
            End If
        '***ver de no harcodear los tipos de promo ver otra alternativa ver lista o coleccion****
        Case "Despegar", "Despegar,com", "Despegar.com"
            If roomtype Like "*PROMOS*" Or roomtype Like "*CUPO*" Or roomtype Like "*PAQUETE*" Or roomtype Like "*NR Promociones ARG*" Or roomtype Like "*OPACO*" Then
             observacionReserva = observacion.iObservacion_observacion("A CARGO DEL PAX (Hotel Collects Payment) A CARGO DE Despegar(Despegar Collects Payment) ", nombreota, " No Reembolsable")
             roomtype.Offset(0, 7) = observacionReserva
            ElseIf roomtype <> "*PROMOS*" Or roomtype <> "*CUPO*" Or roomtype <> "*PAQUETE*" Or roomtype <> "*NR Promociones ARG*" Or roomtype <> "*OPACO*" Then
             observacionReserva = observacion.iObservacion_observacion("A CARGO DEL PAX (Hotel Collects Payment) A CARGO DE Despegar(Despegar Collects Payment) ", nombreota, "  Reembolsable")
             roomtype.Offset(0, 7) = observacionReserva
            End If
        Case "almundo,com", "almundo.com"
         observacionReserva = observacion.iObservacion_observacion("Alojamiento Cta Cte ", nombreota, "(Reembolsable)")
         roomtype.Offset(0, 7) = observacionReserva
        Case "Best Day"
         observacionReserva = observacion.iObservacion_observacion("Alojamiento Cta Cte ", nombreota, "(Reembolsable)")
         roomtype.Offset(0, 7) = observacionReserva
        Case "Hotelbeds"
         observacionReserva = observacion.iObservacion_observacion("Alojamiento Cta Cte ", nombreota, "(Reembolsable)")
         roomtype.Offset(0, 7) = observacionReserva
        Case "NTIncoming"
         observacionReserva = observacion.iObservacion_observacion("Alojamiento TC virtual W2M  ", nombreota, "(Reembolsable)")
         roomtype.Offset(0, 7) = observacionReserva
         Case "welcomebeds,com", "welcomebeds.com"
         observacionReserva = observacion.iObservacion_observacion("Alojamiento Cta Cte ", nombreota, "(Reembolsable)")
         roomtype.Offset(0, 7) = observacionReserva
         Case "Prestigia"
         observacionReserva = observacion.iObservacion_observacion("Alojamiento y Extras Paga Pax  ", nombreota, "(Reembolsable)")
         roomtype.Offset(0, 7) = observacionReserva
    End Select
    
  Next roomtype
End Sub

Sub iObservacion_agregarMenoresEnObservacion()

Dim menores As Range
Dim observacion As String

    For Each menores In Selection
    
            If menores > 0 Then
            
              observacion = menores.Offset(0, 3) + vbNewLine & "Menores : " & menores
              menores.Offset(0, 3) = observacion
            End If
    
    Next menores
    
End Sub


Sub iObservacion_agregarJuntoConEnObservacion()
 Dim gds As Range
 Dim observacion As String
 Dim filaInferior As String
 Dim filaSuperior As String
 
 
 
        For Each gds In Selection
        filaSuperior = gds.Offset(-1, 0)
        filaInferior = gds.Offset(1, 0)
                  
          If (gds = filaSuperior) Or (gds = filaInferior) Then
          
            observacion = gds.Offset(0, 18) + vbNewLine & "Junto Con Gds: " & gds
            gds.Offset(0, 18) = observacion
            
          End If
        
        Next gds

End Sub

Sub iObservacion_agregarPromoWebLoi()
 Dim roomtype As Range
  Dim observacion As String
 
 
        For Each roomtype In Selection
       
          
          If roomtype Like "*WEBLOI*" Then
          
            observacion = roomtype.Offset(0, 7) + vbNewLine & "WEBLOI: Free Upgrade (subject to availability) Free Late Check out (2 hours late)"
            roomtype.Offset(0, 7) = observacion
            
          End If
        
        Next roomtype

End Sub

'Sub iObservacion_crearobservacion(TipoDeReembolso As String)
'Dim roomtype As Range
'Dim nombreota As String
'Dim observacionReserva As String
'Dim observacion As New ClsObservacion
'
'For Each roomtype In Selection
'nombreota = roomtype.Offset(0, -10)
'If nombreota = "Expedia" Then
'    If roomtype Like "*" & TipoDeReembolso & "*" Then
'     observacionReserva = observacion.iObservacion_observacion(nombreota, " Vip - Premium VIP beneficios: 1 Free beverage per person for 2 (once per stay) ECI sujeto a disponibilidad LCO confirmado hasta las 14hs Upgrade sujeto a disponibilidad", " No Reembolsable", "A CARGO DEL PAX (Hotel Collects Payment) / A CARGO DE Expedia(Expedia Collects Payment) /Elegir el que corresponde")
'     roomtype.Offset(0, 7) = observacionReserva
'    ElseIf roomtype <> "*" & TipoDeReembolso & "*" Then
'    observacionReserva = observacion.iObservacion_observacion(nombreota, "Vip - Premium VIP beneficios: 1 Free beverage per person for 2 (once per stay) ECI sujeto a disponibilidad LCO confirmado hasta las 14hs Upgrade sujeto a disponibilidad", "Reembolsable", "A CARGO DEL PAX (Hotel Collects Payment) / A CARGO DE Expedia(Expedia Collects Payment) /Elegir el que corresponde")
'    roomtype.Offset(0, 7) = observacionReserva
'    End If
'
'ElseIf nombreota = "Booking" Then
'    If roomtype Like "*" & TipoDeReembolso & "*" Then
'     observacionReserva = observacion.iObservacion_observacion(nombreota, "Genius: Free upgrade ( IMPORTANTE : Sujeto A Disponibilidad)", "No Reembolsable", "Booking Alojamiento y Extras Paga Pax")
'     roomtype.Offset(0, 7) = observacionReserva
'    ElseIf roomtype <> "*" & TipoDeReembolso & "*" Then
'     observacionReserva = observacion.iObservacion_observacion(nombreota, "Genius: Free upgrade ( IMPORTANTE : Sujeto A Disponibilidad)", "Reembolsable", "Alojamiento y Extras Paga Pax")
'     roomtype.Offset(0, 7) = observacionReserva
'    End If
'
'ElseIf nombreota = "Bookassist" Then
'    If roomtype Like "*" & TipoDeReembolso & "*" Then
'     observacionReserva = observacion.iObservacion_observacion(nombreota, "WEBLOI: Free Upgrade (subject to availability) Free Late Check out (2 hours late)", "No Reembolsable", "Alojamiento y Extras Paga Pax")
'     roomtype.Offset(0, 7) = observacionReserva
'    ElseIf roomtype <> "*" & TipoDeReembolso & "*" Then
'     observacionReserva = observacion.iObservacion_observacion(nombreota, "WEBLOI: Free Upgrade (subject to availability) Free Late Check out (2 hours late)", "No Reembolsable", "Alojamiento y Extras Paga Pax")
'     roomtype.Offset(0, 7) = observacionReserva
'    End If
'
'ElseIf nombreota = "almundo,com" Then
'     observacionReserva = observacion.iObservacion_observacion(nombreota, "Solicitudes especiales: ", "Reembolsable/No Reembolsable", "Alojamiento Cta Cte ")
'     roomtype.Offset(0, 7) = observacionReserva
'
'ElseIf nombreota = "Best Day" Then
'     observacionReserva = observacion.iObservacion_observacion(nombreota, "Solicitudes especiales: ", "Reembolsable/No Reembolsable", "Alojamiento Cta Cte ")
'     roomtype.Offset(0, 7) = observacionReserva
'End If
'Next roomtype
'End Sub

