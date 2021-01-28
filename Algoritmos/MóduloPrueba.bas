Attribute VB_Name = "MóduloPrueba"
Sub cancelarduplicados()
Attribute cancelarduplicados.VB_ProcData.VB_Invoke_Func = "h\n14"

Dim reserva As Range
'i = 1
For Each reserva In Selection

If reserva = reserva.Offset(1, 0) Then

reserva.Offset(1, 0).EntireRow.Delete
    
    i = i + 1

End If

Next

End Sub

Sub Cancelaciones()
Attribute Cancelaciones.VB_ProcData.VB_Invoke_Func = "u\n14"

Dim nombrelibro As String
Dim hojalibro As String
Dim columna As String
Dim m_wbBook As Workbook
Dim m_wsSheet As Worksheet
Dim m_rnCheck As Range
Dim TarifaCm As Double
Dim EstadoReservaCm As String
Dim apellido As String
Dim NombreCompleto As String
Dim DireccionCeldaApellidoEncontrado As String
Dim StatusReserva As String
Dim Huesped As String
Dim PrecioRt As Double
Dim PrecioTotalrt As Double
Dim tipoDeCambio As Double
Dim FechaLlegada As Date
Dim FechaSalida As Date
Dim FechaInCm As Date
Dim FechaOutCm As Date
Dim FormatearPrecioCm As Variant
Dim FormatearPrecioRt As Variant
Dim ApellidoEncontrado As Range
Dim ApellidoReporteCm As Range
Dim reserva As clasereserva
Dim NumRes As String



Range("C3").Select
ActiveCell.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.SpecialCells(xlCellTypeVisible).Select

nombrelibro = InputBox("Nombre Como se Guardo ")
hojalibro = InputBox("Nombre Hoja que desea buscar")
columna = InputBox("Columna que desea buscar Por Defecto elegir la que tiene el Numero de Reserva")
'TipoDecambio = InputBox("Colocar Tipo de Cambio")



Set m_wbBook = Workbooks(nombrelibro)
Set m_wsSheet = m_wbBook.Sheets(hojalibro)
Set m_rnCheck = m_wsSheet.Range(columna & ":" & columna)
Set reserva = New clasereserva

            For Each ApellidoReporteCm In Selection
                With reserva
                .NumeroDeconfirmacion = ApellidoReporteCm.Offset(0, 0)
                NumRes = reserva.NumeroDeconfirmacion
'                    FormatearPrecioCm = Format(TarifaCm, "##,##0.00")
                End With

                    With m_rnCheck

                    Set ApellidoEncontrado = .Find(What:=NumRes, LookAt:=xlPart)

                        If Not ApellidoEncontrado Is Nothing Then

'                        DireccionCeldaApellidoEncontrado = ApellidoEncontrado.Address

                        Precio = ApellidoEncontrado.Offset(0, -4)
                                                
                        ApellidoReporteCm.Offset(0, 2) = Precio

                          
                      End If

            End With
        Next
End Sub



'Sub Prueba()
'
'Dim nombrelibro As String
'Dim hojalibro As String
'Dim columna As String
'Dim m_wbBook As Workbook
'Dim m_wsSheet As Worksheet
'Dim m_rnCheck As Range
'Dim TarifaCm As Double
'Dim EstadoReservaCm As String
'Dim apellido As String
'Dim NombreCompleto As String
'Dim DireccionCeldaApellidoEncontrado As String
'Dim StatusReserva As String
'Dim Huesped As String
'Dim PrecioRt As Double
'Dim PrecioTotalrt As Double
'Dim TipoDecambio As Double
'Dim FechaLlegada As Date
'Dim FechaSalida As Date
'Dim FechaInCm As Date
'Dim FechaOutCm As Date
'Dim FormatearPrecioCm As Variant
'Dim FormatearPrecioRt As Variant
'Dim ApellidoEncontrado As Range
'Dim ApellidoReporteCm As Range
'Dim reserva As clasereserva
'
'
'Range("F1").Select
'ActiveCell.Offset(1, 0).Select
'Range(Selection, Selection.End(xlDown)).Select
'Selection.SpecialCells(xlCellTypeVisible).Select
'
'nombrelibro = InputBox("Nombre Como se Guardo ")
'hojalibro = InputBox("Nombre Hoja que desea buscar")
'columna = InputBox("Columna que desea buscar Por defecto elegir F")
'TipoDecambio = InputBox("Colocar Tipo de Cambio")
'
'
'
'Set m_wbBook = Workbooks(nombrelibro)
'Set m_wsSheet = m_wbBook.Sheets(hojalibro)
'Set m_rnCheck = m_wsSheet.Range(columna & ":" & columna)
'Set reserva = New clasereserva
'
'            For Each ApellidoReporteCm In Selection
'                With reserva
'                .nombre = ApellidoReporteCm.Offset(0, 0)
'                .fechain = ApellidoReporteCm.Offset(0, 3)
'                .fechaout = ApellidoReporteCm.Offset(0, 4)
'                .Status = ApellidoReporteCm.Offset(0, -5)
'                .Precio = ApellidoReporteCm.Offset(0, -3)
'
'                NombreCompleto = reserva.nombre
'                reserva.ObtenerPosicion (NombreCompleto)
'                apellido = reserva.nombre
'                FechaInCm = reserva.fechain
'                FechaOutCm = reserva.fechaout
'                EstadoReservaCm = reserva.Status
'                TarifaCm = reserva.Precio
'                FormatearPrecioCm = Format(TarifaCm, "##,##0.00")
'
'                End With
'
'                    With m_rnCheck
'
'                    Set ApellidoEncontrado = .Find(What:=apellido, LookAt:=xlPart)
'
'                        If Not ApellidoEncontrado Is Nothing Then
'
'                        DireccionCeldaApellidoEncontrado = ApellidoEncontrado.Address
'
'                        Huesped = ApellidoEncontrado.Offset(0, 0)
'                        FechaLlegada = ApellidoEncontrado.Offset(0, 1)
'                        FechaSalida = ApellidoEncontrado.Offset(0, 2)
'                        StatusReserva = ApellidoEncontrado.Offset(0, 7)
'                        Precio = ApellidoEncontrado.Offset(0, 5)
'                        PrecioTotalrt = Precio * TipoDecambio
'                        ApellidoEncontrado.Offset(0, 5) = PrecioTotalrt
'                        FormatearPrecioRt = Format(PrecioTotalrt, "##,##0.00")
'
'                            If apellido Like "*" & Huesped & "*" Then
'
'                                ApellidoReporteCm.Offset(0, 0).Interior.ColorIndex = 4
'                                ApellidoEncontrado.Offset(0, 0).Interior.ColorIndex = 4
'
'                            Else
'
'                                MensajeGuestName = "Ver Nombre  " & apellido & " Celda " & ApellidoReporteCm.Address(RowAbsolute:=False, ColumnAbsolute:=False) _
'                                & " / " & Huesped & " celda " & ApellidoEncontrado.Address(RowAbsolute:=False, ColumnAbsolute:=False)
'
'                                ApellidoReporteCm.Offset(0, 0).Interior.ColorIndex = 3
'                                ApellidoEncontrado.Offset(0, 0).Interior.ColorIndex = 3
'                            End If
'
'
'                            If FechaInCm = FechaLlegada Then
'
'                                ApellidoReporteCm.Offset(0, 3).Interior.ColorIndex = 4
'                                ApellidoEncontrado.Offset(0, 1).Interior.ColorIndex = 4
'
'                            Else
'
'                                MensajeCheckin = "Ver Fecha Llegada Cm " & FechaInCm & " Celda " & ApellidoReporteCm.Address(RowAbsolute:=False, ColumnAbsolute:=False) _
'                                & " /" & " Fecha Rate Ingreso " & FechaLlegada & " celda " & ApellidoEncontrado.Address(RowAbsolute:=False, ColumnAbsolute:=False)
'
'                                ApellidoReporteCm.Offset(0, 3).Interior.ColorIndex = 3
'                                ApellidoEncontrado.Offset(0, 1).Interior.ColorIndex = 3
'                            End If
'
'
'                            If FechaOutCm = FechaSalida Then
'
'                                ApellidoReporteCm.Offset(0, 4).Interior.ColorIndex = 4
'                                ApellidoEncontrado.Offset(0, 2).Interior.ColorIndex = 4
'
'                            Else
'
'                                MensajeCheckout = "Ver Fecha Salida Cm " & FechaOutCm & " Celda " & ApellidoReporteCm.Address(RowAbsolute:=False, ColumnAbsolute:=False) _
'                                & " /" & " Fecha Rate Salida" & FechaSalida & " celda " & ApellidoEncontrado.Address(RowAbsolute:=False, ColumnAbsolute:=False)
'
'                                ApellidoReporteCm.Offset(0, 4).Interior.ColorIndex = 3
'                                ApellidoEncontrado.Offset(0, 2).Interior.ColorIndex = 3
'                            End If
'
'                            If EstadoReservaCm Like "*" & StatusReserva & "*" Then
'
'                                ApellidoReporteCm.Offset(0, -5).Interior.ColorIndex = 4
'                                ApellidoEncontrado.Offset(0, 7).Interior.ColorIndex = 4
'
'                            Else
'
'                                MensajeStatus = "Ver Status Cm  " & EstadoReservaCm & " Celda " & ApellidoReporteCm.Address(RowAbsolute:=False, ColumnAbsolute:=False) _
'                                & " / " & "Status Rate " & StatusReserva & " celda " & ApellidoEncontrado.Address(RowAbsolute:=False, ColumnAbsolute:=False)
'
'                                ApellidoReporteCm.Offset(0, -5).Interior.ColorIndex = 3
'                                ApellidoEncontrado.Offset(0, 7).Interior.ColorIndex = 3
'                            End If
'
'                            If FormatearPrecioCm Like "*" & FormatearPrecioRt & "*" Then
'
'                                ApellidoReporteCm.Offset(0, -3).Interior.ColorIndex = 4
'                                ApellidoEncontrado.Offset(0, 5).Interior.ColorIndex = 4
'
'                            Else
'
'                                MensajeTarifa = "Ver Tarifa Cm  " & TarifaCm & " Celda " & ApellidoReporteCm.Address(RowAbsolute:=False, ColumnAbsolute:=False) _
'                                & " / " & "Precio Rate " & Precio & " celda " & ApellidoEncontrado.Address(RowAbsolute:=False, ColumnAbsolute:=False)
'
'                                ApellidoReporteCm.Offset(0, -3).Interior.ColorIndex = 3
'                                ApellidoEncontrado.Offset(0, 5).Interior.ColorIndex = 3
'
'
'
'                            ApellidoReporteCm.Offset(0, 5).Value = "Discrepancias :" & MensajeGuestName & MensajeCheckin _
'                            & " " & MensajeCheckout & MensajeStatus & " " & Mensajeprecio
'
'
'                            MensajeGuestName = ""
'                            MensajeCheckin = ""
'                            MensajeCheckout = ""
'                            MensajeStatus = ""
'                            Mensajeprecio = ""
'
'                            End If
'                      End If
'
'            End With
'        Next
'End Sub
