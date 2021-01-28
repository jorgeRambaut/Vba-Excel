Attribute VB_Name = "Refactorizado_Rate"
Public Type condicion

     reservas As ClsReserva
     reembolsableOnoReembolsable As Range
     ObtengoTipoDeCondicion As String
End Type


Sub RateTiger()
Attribute RateTiger.VB_Description = "r"
Attribute RateTiger.VB_ProcData.VB_Invoke_Func = "q\n14"

'ver como usar interface porque estamos programando para una clase y no para una interfaz
Dim reserva As New ClsReserva
Dim observacion As New ClsObservacion
Dim solapa As String
Dim Precio As Double
Dim celda As Range
Dim tipoDeCambioBookassist As Double
Dim tipoDeCambioDespegar As Double

'Pido la hoja que voy a trabajar
solapa = InputBox("Nombre de Solapa")

tipoDeCambioBookassist = InputBox("Ingresar Tipo de Cambio Bookassist")
MsgBox (tipoDeCambioBookassist)

tipoDeCambioDespegar = InputBox("Ingresar Tipo de Cambio Despegar")
MsgBox (tipoDeCambioDespegar)


'creamos observacion
Call reserva.iComportamientos_BuscarTitulo("Room Type", solapa, "B6", "W6")
Call reserva.iComportamientos_seleccionarCeldasDescendente
Call observacion.iObservacion_crearobservacion("Non Refundable")
'agregamos menores si corresponde
Call reserva.iComportamientos_BuscarTitulo("Children", solapa, "B6", "W6")
Call reserva.iComportamientos_seleccionarCeldasDescendente
Call reserva.iComportamientos_ConvertirValores("Numero")
Call reserva.iComportamientos_BuscarSiHayMenoresEnLaReservaYResaltarlos
Call reserva.iComportamientos_seleccionarCeldasDescendente
Call observacion.iObservacion_agregarMenoresEnObservacion
'vemos si hay mas de una reserva con mismo gds
Call reserva.iComportamientos_BuscarTitulo("Channel ID", solapa, "B6", "W6")
'Call reserva.iComportamientos_seleccionarCeldasDescendente
Call reserva.iComportamientos_BuscarDuplicadosYResaltar
Call observacion.iObservacion_agregarJuntoConEnObservacion
'vemos si hay reservas con promo web loi y lo agregamos
Call reserva.iComportamientos_BuscarTitulo("Room Type", solapa, "B6", "W6")
Call reserva.iComportamientos_seleccionarCeldasDescendente
Call observacion.iObservacion_agregarPromoWebLoi

'convertimos a fecha para poder ordenar
Call reserva.iComportamientos_BuscarTitulo("Booked On", solapa, "B6", "W6")
Call reserva.iComportamientos_seleccionarCeldasDescendente
Call reserva.iComportamientos_ConvertirValores("Fecha")

Call reserva.iComportamientos_BuscarTitulo("Check-in", solapa, "B6", "W6")
Call reserva.iComportamientos_seleccionarCeldasDescendente
Call reserva.iComportamientos_ConvertirValores("Fecha")

Call reserva.iComportamientos_BuscarTitulo("Checkout", solapa, "B6", "W6")
Call reserva.iComportamientos_seleccionarCeldasDescendente
Call reserva.iComportamientos_ConvertirValores("Fecha")

Call reserva.iComportamientos_BuscarTitulo("Avg. Daily Rate", solapa, "B6", "W6")
Call reserva.iComportamientos_seleccionarCeldasDescendente
'Call reserva.iComportamientos_ConvertirValores("Numero")
Call reserva.iComportamientos_Reemplazar(".", ",")



'convertimos a pesos GDS ARS
Call reserva.iComportamientos_BuscarTitulo("Channel ID", solapa, "B6", "W6")
Call reserva.iComportamientos_seleccionarCeldasDescendente
Call reserva.pasarApesosBookassistARS(tipoDeCambioBookassist)

'convertimos a pesos Despegar
Call reserva.iComportamientos_BuscarTitulo("Channel", solapa, "E6", "W6")
Call reserva.iComportamientos_seleccionarCeldasDescendente
Call reserva.pasarApesosDespegar(tipoDeCambioDespegar)

'ajustamos excel
Call reserva.iComportamientos_Ajustaratexto("B6")
Call reserva.iComportamientos_Ocultarcolumnasenblanco("B6")
Call reserva.iComportamientos_Ocultarcolumnasenblanco("H6")
Call reserva.iComportamientos_Ocultarcolumnasenblanco("J6")
Call reserva.iComportamientos_Ocultarcolumnasenblanco("K6")
Call reserva.iComportamientos_Ocultarcolumnasenblanco("L6")
Call reserva.iComportamientos_Ocultarcolumnasenblanco("O6")
Call reserva.iComportamientos_Ocultarcolumnasenblanco("P6")
Call reserva.iComportamientos_Ocultarcolumnasenblanco("Q6")
Call reserva.iComportamientos_Ocultarcolumnasenblanco("R6")
Call reserva.iComportamientos_Ocultarcolumnasenblanco("T6")
Call reserva.iComportamientos_Ocultarcolumnasenblanco("W6")



MsgBox ("Procesos Finalizados")

End Sub

'Sub PRUEBA2()
'Dim reserva As New ClsReserva
'Dim arg As String
'Dim celda As Range
'Dim Precio As Double
'Dim precioEnPesos As Double
''Call reserva.iComportamientos_seleccionarCeldasDescendente
''Call reserva.iComportamientos_Reemplazar(".", ",")
'
'Call reserva.iComportamientos_BuscarTitulo("Channel ID", "Detail-Arrival", "B6", "W6")
'Call reserva.iComportamientos_seleccionarCeldasDescendente
'Call reserva.pasarApesosBookassistARS(90.75)
'
''For Each celda In Selection
'' arg = reserva.iComportamientos_obtenerParte(CStr(celda), 3)
''
'' If arg = "ARS" Then
''
''   precio = celda.Offset(0, 17)
''
''   precioEnPesos = reserva.iComportamientos_pasarAPesos(90, precio)
''
''   celda.Offset(0, 17) = precioEnPesos
''
'' End If
''
''
''Next celda


End Sub
