Attribute VB_Name = "ControlReservas"
Sub ControlReservasOnline()
Attribute ControlReservasOnline.VB_ProcData.VB_Invoke_Func = "k\n14"
Dim nombrelibro As String
Dim hojalibro As String
Dim columna As String
Dim apellido As String
Dim NombreCompleto As String
Dim reserva As clasereserva
Dim m_rnFind As Range
Dim m_stAddress As String
Dim StatusReserva As String
Dim FechaReserva As Date
Dim Cliente As String
Dim TipoHab As String
Dim Huesped As String
Dim NumeroReserva As String
Dim TipoDeHuesped As String
Dim FechaLlegada As Date
Dim FechaSalida As Date
Dim celda As Range
Dim m_wbBook As Workbook
Dim m_wsSheet As Worksheet
Dim m_rnCheck As Range
Dim FechaInCm As Date
Dim FechaOutCm As Date
Dim EstadoReservaCm As String
Dim ClienteCm As String
Dim ChannelRt As String

Range("e1").Select
ActiveCell.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.SpecialCells(xlCellTypeVisible).Select

nombrelibro = InputBox("Nombre Como se Guardo Excel Rate Tiger")
hojalibro = InputBox("Nombre Hoja que desea buscar")
columna = InputBox("Columna que desea buscar Por defecto elegir F")

Set m_wbBook = Workbooks(nombrelibro)
Set m_wsSheet = m_wbBook.Sheets(hojalibro)
Set m_rnCheck = m_wsSheet.Range(columna & ":" & columna)
Set reserva = New clasereserva

For Each celda In Selection
        With reserva
        .nombre = celda.Offset(0, 0)
        .NumeroDeconfirmacion = celda.Offset(0, 1)
        .fechain = celda.Offset(0, 3)
        .fechaout = celda.Offset(0, 4)
        .Status = celda.Offset(0, -4)
        .CanalReserva = celda.Offset(0, -2)
         
        NombreCompleto = reserva.nombre
        reserva.ObtenerPosicion (NombreCompleto)
        apellido = reserva.nombre
        FechaInCm = reserva.fechain
        FechaOutCm = reserva.fechaout
        EstadoReservaCm = reserva.Status
        ClienteCm = reserva.CanalReserva
        
        
        With m_rnCheck
        
        Set m_rnFind = .Find(What:=apellido, LookAt:=xlPart)
        
        If Not m_rnFind Is Nothing Then
        
        
        m_stAddress = m_rnFind.Address
        StatusReserva = m_rnFind.Offset(0, -2)
        FechaReserva = m_rnFind.Offset(0, -1)
        TipoHab = m_rnFind.Offset(0, 3)
        Huesped = m_rnFind.Offset(0, 0)
        NumeroReserva = m_rnFind.Offset(0, -4)
        FechaLlegada = m_rnFind.Offset(0, 1)
        FechaSalida = m_rnFind.Offset(0, 2)
        ChannelRt = m_rnFind.Offset(0, -3)
        
                            
        If ClienteCm Like "*" & ChannelRt & "*" Then
            celda.Offset(0, -2).Interior.ColorIndex = 4
            m_rnFind.Offset(0, -3).Interior.ColorIndex = 4
        Else
            MensajeChannel = "reporte rate tiger " & ChannelRt & _
            " Celda " & m_rnFind.Address(RowAbsolute:=False, ColumnAbsolute:=False) _
            & " / Cm  " & ClienteCm & " celda " & celda.Address(RowAbsolute:=False, ColumnAbsolute:=False)
            celda.Offset(0, -2).Interior.ColorIndex = 3
            m_rnFind.Offset(0, -3).Interior.ColorIndex = 3
        End If

        If EstadoReservaCm Like "*" & StatusReserva & "*" Then
            celda.Offset(0, -4).Interior.ColorIndex = 4
            m_rnFind.Offset(0, -2).Interior.ColorIndex = 4
        Else
            MensajeStatus = "reporte rate tiger " & StatusReserva & _
            " Celda " & m_rnFind.Address(RowAbsolute:=False, ColumnAbsolute:=False) _
            & " / Cm  " & EstadoReservaCm & " celda " & celda.Address(RowAbsolute:=False, ColumnAbsolute:=False)
            celda.Offset(0, -4).Interior.ColorIndex = 3
            m_rnFind.Offset(0, -2).Interior.ColorIndex = 3
        End If

        If FechaInCm = FechaLlegada Then
          m_rnFind.Offset(0, 1).Interior.ColorIndex = 4
          celda.Offset(0, 3).Interior.ColorIndex = 4
        Else
            MensajeCheckin = " Fecha in Cm " & FechaInCm & _
            " Celda " & celda.Address(RowAbsolute:=False, ColumnAbsolute:=False) _
            & " / Rate Tiger " & FechaLlegada & _
            " celda " & m_rnFind.Address(RowAbsolute:=False, ColumnAbsolute:=False)
            m_rnFind.Offset(0, 1).Interior.ColorIndex = 3
            celda.Offset(0, 3).Interior.ColorIndex = 3
        End If

        If FechaOutCm = FechaSalida Then
         m_rnFind.Offset(0, 2).Interior.ColorIndex = 4
         celda.Offset(0, 4).Interior.ColorIndex = 4
        Else
            MensajeCheckout = "Fecha Out reporte Cm " & FechaOutCm & _
            " Celda " & celda.Address(RowAbsolute:=False, ColumnAbsolute:=False) _
            & " / Rate Tiger " & FechaSalida & _
            " celda " & m_rnFind.Address(RowAbsolute:=False, ColumnAbsolute:=False)
            m_rnFind.Offset(0, 2).Interior.ColorIndex = 3
            celda.Offset(0, 4).Interior.ColorIndex = 3
        End If
        
        

         celda.Offset(0, 5).Value = "Discrepancias :" & MensajeStatus & _
         "/ " & Huesped & " /" & apellido & " " & MensajeCheckin _
         & " " & MensajeCheckout & " " & MensajeChannel
         
        MensajeStatus = ""
        MensajeBookedOn = ""
        MensajeGuestName = ""
        MensajeCheckin = ""
        MensajeCheckout = ""
        MensajeChildren = ""
        MensajeChannel = ""
        
End If
End With
End With
Next
MsgBox "Finalizo Control Reservas, Ver Comentarios en Reporte Cm", vbInformation, Title:="Control Reservas"
End Sub


Sub Conciliaciones()
Dim nombrelibro As String
Dim hojalibro As String
Dim columna As String
Dim apellido As String
Dim NombreCompleto As String
Dim reserva As clasereserva
Dim m_rnFind As Range
Dim m_stAddress As String
Dim StatusReserva As String
Dim FechaReserva As Date
Dim Cliente As String
Dim TipoHab As String
Dim Huesped As String
Dim NumeroReserva As String
Dim TipoDeHuesped As String
Dim Precio As String
Dim FechaLlegada As Date
Dim FechaSalida As Date
Dim celda As Range
Dim m_wbBook As Workbook
Dim m_wsSheet As Worksheet
Dim m_rnCheck As Range
Dim ingreso As Date
Dim Salida As Date
Dim CantidadNoches As Integer
Dim total As Double
Dim Importe As Double



Range("e1").Select
ActiveCell.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.SpecialCells(xlCellTypeVisible).Select

nombrelibro = InputBox("Nombre Como se Guardo Expedia")
hojalibro = InputBox("Nombre Hoja que desea buscar")
columna = InputBox("Columna que desea buscar Por defecto elegir D")

Set m_wbBook = Workbooks(nombrelibro)
Set m_wsSheet = m_wbBook.Sheets(hojalibro)
Set m_rnCheck = m_wsSheet.Range(columna & ":" & columna)
Set reserva = New clasereserva



For Each celda In Selection
    With reserva
    .nombre = celda.Offset(0, 0)
    .fechain = celda.Offset(0, 3)
    .fechaout = celda.Offset(0, 4)
    .Status = celda.Offset(0, -4)
    .Precio = celda.Offset(0, -3)
    NombreCompleto = reserva.nombre
    reserva.ObtenerPosicion (NombreCompleto)
    
    apellido = reserva.nombre
    ingreso = reserva.fechain
    Salida = reserva.fechaout
    CantidadNoches = reserva.CalculaNoches(ingreso, Salida)
    Importe = reserva.Precio
    total = CantidadNoches * Importe
          
        With m_rnCheck
        Set m_rnFind = .Find(What:=apellido, LookAt:=xlPart)
        If Not m_rnFind Is Nothing Then
        m_stAddress = m_rnFind.Address
        'StatusReserva = m_rnFind.Offset(0, -2)
        'FechaReserva = m_rnFind.Offset(0, -1)
        'Cliente = m_rnFind.Offset(0, -3)
        'TipoHab = m_rnFind.Offset(0, 3)
        Huesped = m_rnFind.Offset(0, 0)
       ' NumeroReserva = m_rnFind.Offset(0, -4)
        FechaLlegada = m_rnFind.Offset(0, -2)
        FechaSalida = m_rnFind.Offset(0, -1)
        Precio = m_rnFind.Offset(0, 2)
        
'        If reserva.Status Like "*" & StatusReserva & "*" Then
'            CELDA.Offset(0, -4).Interior.ColorIndex = 4
'            m_rnFind.Offset(0, -4).Interior.ColorIndex = 4
'        Else
'            MensajeStatus = "reporte rate tiger " & StatusReserva & _
'            " Celda " & m_rnFind.Address(RowAbsolute:=False, ColumnAbsolute:=False) _
'            & " / Cm  " & reserva.Status & " celda " & CELDA.Address(RowAbsolute:=False, ColumnAbsolute:=False)
'            CELDA.Offset(0, -4).Interior.ColorIndex = 3
'            m_rnFind.Offset(0, -4).Interior.ColorIndex = 3
'        End If

        If ingreso = FechaLlegada Then
          m_rnFind.Offset(0, -2).Interior.ColorIndex = 4
          celda.Offset(0, 3).Interior.ColorIndex = 4
        Else
            MensajeCheckin = " Fecha in Cm " & ingreso & _
            " Celda " & celda.Address(RowAbsolute:=False, ColumnAbsolute:=False) _
            & " / Rate Tiger " & FechaLlegada & _
            " celda " & m_rnFind.Address(RowAbsolute:=False, ColumnAbsolute:=False)
            m_rnFind.Offset(0, 1).Interior.ColorIndex = 3
            celda.Offset(0, 3).Interior.ColorIndex = 3
        End If

        If Salida = FechaSalida Then
         m_rnFind.Offset(0, -1).Interior.ColorIndex = 4
         celda.Offset(0, 4).Interior.ColorIndex = 4
        Else
            MensajeCheckout = "Fecha Out reporte Cm " & Salida & _
            " Celda " & celda.Address(RowAbsolute:=False, ColumnAbsolute:=False) _
            & " / Rate Tiger " & FechaSalida & _
            " celda " & m_rnFind.Address(RowAbsolute:=False, ColumnAbsolute:=False)
            m_rnFind.Offset(0, 2).Interior.ColorIndex = 3
            celda.Offset(0, 4).Interior.ColorIndex = 3
        End If
        
        If total = Precio Then
         m_rnFind.Offset(0, 2).Interior.ColorIndex = 4
         celda.Offset(0, 4).Interior.ColorIndex = 4
        Else
            Mensajeprecio = "Ver Precio reporte Cm " & total & _
            " Celda " & celda.Address(RowAbsolute:=False, ColumnAbsolute:=False) _
            & " / Rate Tiger " & Precio & _
            " celda " & m_rnFind.Address(RowAbsolute:=False, ColumnAbsolute:=False)
            m_rnFind.Offset(0, 2).Interior.ColorIndex = 3
            celda.Offset(0, -3).Interior.ColorIndex = 3
        End If


         celda.Offset(0, 5).Value = "Discrepancias :" & MensajeCheckin _
         & " " & MensajeCheckout & " " & Mensajeprecio
         
        MensajeStatus = ""
        MensajeBookedOn = ""
        MensajeGuestName = ""
        MensajeCheckin = ""
        MensajeCheckout = ""
        MensajeChildren = ""
        
End If
End With
End With
Next
MsgBox "Finalizo Control Reservas, Ver Comentarios en Reporte Cm", vbInformation, Title:="Control Reservas"
End Sub


