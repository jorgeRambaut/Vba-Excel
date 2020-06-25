Attribute VB_Name = "ControlComisiones"
Sub ControlComisionesBooking()

Dim nombrelibro As String
Dim hojalibro As String
Dim columna As String
Dim m_wbBook As Workbook
Dim m_wsSheet As Worksheet
Dim m_rnCheck As Range


Dim EstadoReservaRpt1 As String
Dim EstadoReservaRpt2 As String

Dim apellido As String
Dim apellido2 As String

Dim NombreCompleto As String
Dim DireccionCeldaApellidoEncontrado As String




Dim FechaLlegada As Date
Dim FechaSalida As Date
Dim FechaLlegada2 As Date
Dim FechaSalida2 As Date


Dim ApellidoEncontrado As Range
Dim ApellidoReporteCm As Range

Dim reserva As clasereserva


Range("h1").Select
ActiveCell.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.SpecialCells(xlCellTypeVisible).Select

nombrelibro = InputBox("Nombre Como se Guardo ")
hojalibro = InputBox("Nombre Hoja que desea buscar")
columna = InputBox("Columna que desea buscar")




Set m_wbBook = Workbooks(nombrelibro)
Set m_wsSheet = m_wbBook.Sheets(hojalibro)
Set m_rnCheck = m_wsSheet.Range(columna & ":" & columna)

Set reserva = New clasereserva

            For Each ApellidoReporteCm In Selection
            
                With reserva
                    .nombre = ApellidoReporteCm.Offset(0, 0)
                    .fechain = ApellidoReporteCm.Offset(0, 3)
                    .fechaout = ApellidoReporteCm.Offset(0, 4)
                    .Status = ApellidoReporteCm.Offset(0, -6)
                    .FechaReserva = ApellidoReporteCm.Offset(0, -5)
                    
    
                    NombreCompleto = reserva.nombre
                    
                    reserva.ObtenerPosicion (NombreCompleto)
                    apellido = reserva.nombre
                    FechaLlegada = reserva.fechain
                    FechaSalida = reserva.fechaout
                    EstadoReservaRpt1 = reserva.Status
                
                End With

                    With m_rnCheck
                   

                        Set ApellidoEncontrado = .Find(What:=apellido, LookAt:=xlPart)

                        If Not ApellidoEncontrado Is Nothing Then

                            DireccionCeldaApellidoEncontrado = ApellidoEncontrado.Address
    
                            apellido2 = ApellidoEncontrado.Offset(0, 0)
                            FechaLlegada2 = ApellidoEncontrado.Offset(0, 1)
                            FechaSalida2 = ApellidoEncontrado.Offset(0, 2)
                            EstadoReservaRpt2 = ApellidoEncontrado.Offset(0, 4)
                            
                            Do
                            
                            If apellido2 Like "*" & apellido & "*" _
                              And FechaLlegada = FechaLlegada2 _
                              And FechaSalida = FechaSalida2 And _
                              EstadoReservaRpt1 Like "*" & EstadoReservaRpt2 & "*" Then
                                                           
                              
                                ApellidoReporteCm.Offset(0, 0).Interior.ColorIndex = 4
                                ApellidoEncontrado.Offset(0, 0).Interior.ColorIndex = 4
                                ApellidoReporteCm.Offset(0, 3).Interior.ColorIndex = 4
                                ApellidoEncontrado.Offset(0, 1).Interior.ColorIndex = 4
                                ApellidoReporteCm.Offset(0, 4).Interior.ColorIndex = 4
                                ApellidoEncontrado.Offset(0, 2).Interior.ColorIndex = 4
                                ApellidoReporteCm.Offset(0, -6).Interior.ColorIndex = 4
                                ApellidoEncontrado.Offset(0, 4).Interior.ColorIndex = 4
                                
                                
                            Else
                                ApellidoReporteCm.Offset(0, 0).Interior.ColorIndex = 3
                                ApellidoEncontrado.Offset(0, 0).Interior.ColorIndex = 3
                                ApellidoReporteCm.Offset(0, 3).Interior.ColorIndex = 3
                                ApellidoEncontrado.Offset(0, 1).Interior.ColorIndex = 3
                                ApellidoReporteCm.Offset(0, 4).Interior.ColorIndex = 3
                                ApellidoEncontrado.Offset(0, 2).Interior.ColorIndex = 3
                                ApellidoReporteCm.Offset(0, -6).Interior.ColorIndex = 3
                                ApellidoEncontrado.Offset(0, 4).Interior.ColorIndex = 3
                                
                             End If
                          Set ApellidoEncontrado = .FindNext(ApellidoEncontrado)
                          
                          Loop While Not ApellidoEncontrado Is Nothing And ApellidoEncontrado.Address <> DireccionCeldaApellidoEncontrado
                          
                          End If
                
                End With
        Next
        
    ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveSheet.AutoFilter.Sort.SortFields.Add(Selection _
    , xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color _
        = RGB(255, 255, 255)
        
    ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveSheet.AutoFilter.Sort.SortFields.Add(Selection _
    , xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color _
        = RGB(255, 0, 0)
    

    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub
                              
                              

'                                If apellido2 Like "*" & apellido & "*" Then
'
'                                    ApellidoReporteCm.Offset(0, 0).Interior.ColorIndex = 4
'                                    ApellidoEncontrado.Offset(0, 0).Interior.ColorIndex = 4
'
'                                Else
'                                    ApellidoReporteCm.Offset(0, 0).Interior.ColorIndex = 3
'                                    ApellidoEncontrado.Offset(0, 0).Interior.ColorIndex = 3
'                                End If
'
'
'                                If FechaLlegada = FechaLlegada2 Then
'
'                                    ApellidoReporteCm.Offset(0, 3).Interior.ColorIndex = 4
'                                    ApellidoEncontrado.Offset(0, 1).Interior.ColorIndex = 4
'
'                                Else
'                                    ApellidoReporteCm.Offset(0, 3).Interior.ColorIndex = 3
'                                    ApellidoEncontrado.Offset(0, 1).Interior.ColorIndex = 3
'                                End If
'
'
'                                If FechaSalida = FechaSalida2 Then
'
'                                    ApellidoReporteCm.Offset(0, 4).Interior.ColorIndex = 4
'                                    ApellidoEncontrado.Offset(0, 2).Interior.ColorIndex = 4
'
'                                Else
'
'                                    ApellidoReporteCm.Offset(0, 4).Interior.ColorIndex = 3
'                                    ApellidoEncontrado.Offset(0, 2).Interior.ColorIndex = 3
'                                End If
'
'                                If EstadoReservaRpt1 Like "*" & EstadoReservaRpt2 & "*" Then
'
'                                    ApellidoReporteCm.Offset(0, -6).Interior.ColorIndex = 4
'                                    ApellidoEncontrado.Offset(0, 4).Interior.ColorIndex = 4
'
'                                Else
'
'                                    ApellidoReporteCm.Offset(0, -6).Interior.ColorIndex = 3
'                                    ApellidoEncontrado.Offset(0, 4).Interior.ColorIndex = 3
'                                End If
                                


Sub ControlComisionesExpedia()

Dim nombrelibro As String
Dim hojalibro As String
Dim columna As String
Dim m_wbBook As Workbook
Dim m_wsSheet As Worksheet
Dim m_rnCheck As Range


Dim EstadoReservaRpt1 As String
Dim EstadoReservaRpt2 As String

Dim apellido As String
Dim apellido2 As String

Dim NombreCompleto As String
Dim DireccionCeldaApellidoEncontrado As String




Dim FechaLlegada As Date
Dim FechaSalida As Date
Dim FechaLlegada2 As Date
Dim FechaSalida2 As Date


Dim ApellidoEncontrado As Range
Dim ApellidoReporteCm As Range

Dim reserva As clasereserva


Range("h1").Select
ActiveCell.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.SpecialCells(xlCellTypeVisible).Select

nombrelibro = InputBox("Nombre Como se Guardo ")
hojalibro = InputBox("Nombre Hoja que desea buscar")
columna = InputBox("Columna que desea buscar")




Set m_wbBook = Workbooks(nombrelibro)
Set m_wsSheet = m_wbBook.Sheets(hojalibro)
Set m_rnCheck = m_wsSheet.Range(columna & ":" & columna)

Set reserva = New clasereserva

            For Each ApellidoReporteCm In Selection
            
                With reserva
                    .nombre = ApellidoReporteCm.Offset(0, 0)
                    .fechain = ApellidoReporteCm.Offset(0, 3)
                    .fechaout = ApellidoReporteCm.Offset(0, 4)
                    .Status = ApellidoReporteCm.Offset(0, -6)
                    .FechaReserva = ApellidoReporteCm.Offset(0, -5)
    
                    NombreCompleto = reserva.nombre
                    reserva.ObtenerPosicion (NombreCompleto)
                    apellido = reserva.nombre
                    FechaLlegada = reserva.fechain
                    FechaSalida = reserva.fechaout
                    EstadoReservaRpt1 = reserva.Status
                
                End With

                    With m_rnCheck

                        Set ApellidoEncontrado = .Find(What:=apellido, LookAt:=xlPart)

                        If Not ApellidoEncontrado Is Nothing Then

                            DireccionCeldaApellidoEncontrado = ApellidoEncontrado.Address
    
                            apellido2 = ApellidoEncontrado.Offset(0, 0)
                            FechaLlegada2 = ApellidoEncontrado.Offset(0, 1)
                            FechaSalida2 = ApellidoEncontrado.Offset(0, 2)
                            EstadoReservaRpt2 = ApellidoEncontrado.Offset(0, 7)
                            
                            Do
                            
                              If apellido2 Like "*" & apellido & "*" _
                              And FechaLlegada = FechaLlegada2 _
                              And FechaSalida = FechaSalida2 And _
                              EstadoReservaRpt1 Like "*" & EstadoReservaRpt2 & "*" Then
                              
                              ApellidoReporteCm.Offset(0, 0).Interior.ColorIndex = 4
                              ApellidoEncontrado.Offset(0, 0).Interior.ColorIndex = 4
                              ApellidoReporteCm.Offset(0, 3).Interior.ColorIndex = 4
                              ApellidoEncontrado.Offset(0, 1).Interior.ColorIndex = 4
                              ApellidoReporteCm.Offset(0, 4).Interior.ColorIndex = 4
                              ApellidoEncontrado.Offset(0, 2).Interior.ColorIndex = 4
                              ApellidoReporteCm.Offset(0, 4).Interior.ColorIndex = 4
                              ApellidoEncontrado.Offset(0, 2).Interior.ColorIndex = 4
                              ApellidoReporteCm.Offset(0, -6).Interior.ColorIndex = 4
                              ApellidoEncontrado.Offset(0, 7).Interior.ColorIndex = 4
                              
                              Else
                              ApellidoReporteCm.Offset(0, 0).Interior.ColorIndex = 3
                              ApellidoEncontrado.Offset(0, 0).Interior.ColorIndex = 3
                              ApellidoReporteCm.Offset(0, 3).Interior.ColorIndex = 3
                              ApellidoEncontrado.Offset(0, 1).Interior.ColorIndex = 3
                              ApellidoReporteCm.Offset(0, 4).Interior.ColorIndex = 3
                              ApellidoEncontrado.Offset(0, 2).Interior.ColorIndex = 3
                              ApellidoReporteCm.Offset(0, 4).Interior.ColorIndex = 3
                              ApellidoEncontrado.Offset(0, 2).Interior.ColorIndex = 3
                              ApellidoReporteCm.Offset(0, -6).Interior.ColorIndex = 3
                              ApellidoEncontrado.Offset(0, 7).Interior.ColorIndex = 3
                              
                              
                              
                              End If
                              
                              
                              
                            Set ApellidoEncontrado = .FindNext(ApellidoEncontrado)
                          
                          Loop While Not ApellidoEncontrado Is Nothing And ApellidoEncontrado.Address <> DireccionCeldaApellidoEncontrado
                          
                          End If
                
                End With
        Next
        
        ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveSheet.AutoFilter.Sort.SortFields.Add(Selection _
    , xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color _
        = RGB(255, 255, 255)
        
    ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveSheet.AutoFilter.Sort.SortFields.Add(Selection _
    , xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color _
        = RGB(255, 0, 0)
    

    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub

'                                If apellido2 Like "*" & apellido & "*" Then
'
'                                    ApellidoReporteCm.Offset(0, 0).Interior.ColorIndex = 4
'                                    ApellidoEncontrado.Offset(0, 0).Interior.ColorIndex = 4
'
'                                Else
'                                    ApellidoReporteCm.Offset(0, 0).Interior.ColorIndex = 3
'                                    ApellidoEncontrado.Offset(0, 0).Interior.ColorIndex = 3
'                                End If
'
'
'                                If FechaLlegada = FechaLlegada2 Then
'
'                                    ApellidoReporteCm.Offset(0, 3).Interior.ColorIndex = 4
'                                    ApellidoEncontrado.Offset(0, 1).Interior.ColorIndex = 4
'
'                                Else
'                                    ApellidoReporteCm.Offset(0, 3).Interior.ColorIndex = 3
'                                    ApellidoEncontrado.Offset(0, 1).Interior.ColorIndex = 3
'                                End If
'
'
'                                If FechaSalida = FechaSalida2 Then
'
'                                    ApellidoReporteCm.Offset(0, 4).Interior.ColorIndex = 4
'                                    ApellidoEncontrado.Offset(0, 2).Interior.ColorIndex = 4
'
'                                Else
'
'                                    ApellidoReporteCm.Offset(0, 4).Interior.ColorIndex = 3
'                                    ApellidoEncontrado.Offset(0, 2).Interior.ColorIndex = 3
'                                End If
'
'                                If EstadoReservaRpt1 Like "*" & EstadoReservaRpt2 & "*" Then
'
'                                    ApellidoReporteCm.Offset(0, -6).Interior.ColorIndex = 4
'                                    ApellidoEncontrado.Offset(0, 7).Interior.ColorIndex = 4
'
'                                Else
'
'                                    ApellidoReporteCm.Offset(0, -6).Interior.ColorIndex = 3
'                                    ApellidoEncontrado.Offset(0, 7).Interior.ColorIndex = 3
'                                End If
                                
                      


Sub ControlComisionesDespegarBookassist()

Dim nombrelibro As String
Dim hojalibro As String
Dim columna As String
Dim m_wbBook As Workbook
Dim m_wsSheet As Worksheet
Dim m_rnCheck As Range


Dim EstadoReservaRpt1 As String
Dim EstadoReservaRpt2 As String

Dim apellido As String
Dim apellido2 As String

Dim NombreCompleto As String
Dim DireccionCeldaApellidoEncontrado As String




Dim FechaLlegada As Date
Dim FechaSalida As Date
Dim FechaLlegada2 As Date
Dim FechaSalida2 As Date


Dim ApellidoEncontrado As Range
Dim ApellidoReporteCm As Range

Dim reserva As clasereserva


Range("h1").Select
ActiveCell.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.SpecialCells(xlCellTypeVisible).Select

nombrelibro = InputBox("Nombre Como se Guardo ")
hojalibro = InputBox("Nombre Hoja que desea buscar")
columna = InputBox("Columna que desea buscar")




Set m_wbBook = Workbooks(nombrelibro)
Set m_wsSheet = m_wbBook.Sheets(hojalibro)
Set m_rnCheck = m_wsSheet.Range(columna & ":" & columna)

Set reserva = New clasereserva

            For Each ApellidoReporteCm In Selection
            
                With reserva
                    .nombre = ApellidoReporteCm.Offset(0, 0)
                    .fechain = ApellidoReporteCm.Offset(0, 3)
                    .fechaout = ApellidoReporteCm.Offset(0, 4)
                    .Status = ApellidoReporteCm.Offset(0, -6)
                    .FechaReserva = ApellidoReporteCm.Offset(0, -5)
    
                    NombreCompleto = reserva.nombre
                    reserva.ObtenerPosicion (NombreCompleto)
                    apellido = reserva.nombre
                    FechaLlegada = reserva.fechain
                    FechaSalida = reserva.fechaout
                    EstadoReservaRpt1 = reserva.Status
                
                End With

                    With m_rnCheck

                        Set ApellidoEncontrado = .Find(What:=apellido, LookAt:=xlPart)

                        If Not ApellidoEncontrado Is Nothing Then

                            DireccionCeldaApellidoEncontrado = ApellidoEncontrado.Address
    
                            apellido2 = ApellidoEncontrado.Offset(0, 0)
                            FechaLlegada2 = ApellidoEncontrado.Offset(0, 1)
                            FechaSalida2 = ApellidoEncontrado.Offset(0, 2)
                            EstadoReservaRpt2 = ApellidoEncontrado.Offset(0, -2)
                            
                            Do
                              If apellido2 Like "*" & apellido & "*" _
                              And FechaLlegada = FechaLlegada2 _
                              And FechaSalida = FechaSalida2 And _
                              EstadoReservaRpt1 Like "*" & EstadoReservaRpt2 & "*" Then
                              
                              ApellidoReporteCm.Offset(0, 0).Interior.ColorIndex = 4
                              ApellidoEncontrado.Offset(0, 0).Interior.ColorIndex = 4
                              ApellidoReporteCm.Offset(0, 3).Interior.ColorIndex = 4
                              ApellidoEncontrado.Offset(0, 1).Interior.ColorIndex = 4
                              ApellidoReporteCm.Offset(0, 4).Interior.ColorIndex = 4
                              ApellidoEncontrado.Offset(0, 2).Interior.ColorIndex = 4
                              ApellidoReporteCm.Offset(0, -6).Interior.ColorIndex = 4
                              ApellidoEncontrado.Offset(0, -2).Interior.ColorIndex = 4
                              
                              Else
                              
'                              ApellidoReporteCm.Offset(0, 0).Interior.ColorIndex = 3
'                              ApellidoEncontrado.Offset(0, 0).Interior.ColorIndex = 3
                              ApellidoReporteCm.Offset(0, 3).Interior.ColorIndex = 3
                              ApellidoEncontrado.Offset(0, 1).Interior.ColorIndex = 3
                              ApellidoReporteCm.Offset(0, 4).Interior.ColorIndex = 3
                              ApellidoEncontrado.Offset(0, 2).Interior.ColorIndex = 3
                              ApellidoReporteCm.Offset(0, -6).Interior.ColorIndex = 3
                              ApellidoEncontrado.Offset(0, -2).Interior.ColorIndex = 3
                              End If
                              
                              Set ApellidoEncontrado = .FindNext(ApellidoEncontrado)
                          
                          Loop While Not ApellidoEncontrado Is Nothing And ApellidoEncontrado.Address <> DireccionCeldaApellidoEncontrado
                          
                          End If
                
                End With
        Next
        ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveSheet.AutoFilter.Sort.SortFields.Add(Selection _
    , xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color _
        = RGB(255, 255, 255)
        
    ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveSheet.AutoFilter.Sort.SortFields.Add(Selection _
    , xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color _
        = RGB(255, 0, 0)
    

    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub
                              
                              
                              
'                                If apellido2 Like "*" & apellido & "*" Then
'
'                                    ApellidoReporteCm.Offset(0, 0).Interior.ColorIndex = 4
'                                    ApellidoEncontrado.Offset(0, 0).Interior.ColorIndex = 4
'
'                                Else
'                                    ApellidoReporteCm.Offset(0, 0).Interior.ColorIndex = 3
'                                    ApellidoEncontrado.Offset(0, 0).Interior.ColorIndex = 3
'                                End If
'
'
'                                If FechaLlegada = FechaLlegada2 Then
'
'                                    ApellidoReporteCm.Offset(0, 3).Interior.ColorIndex = 4
'                                    ApellidoEncontrado.Offset(0, 1).Interior.ColorIndex = 4
'
'                                Else
'                                    ApellidoReporteCm.Offset(0, 3).Interior.ColorIndex = 3
'                                    ApellidoEncontrado.Offset(0, 1).Interior.ColorIndex = 3
'                                End If
'
'
'                                If FechaSalida = FechaSalida2 Then
'
'                                    ApellidoReporteCm.Offset(0, 4).Interior.ColorIndex = 4
'                                    ApellidoEncontrado.Offset(0, 2).Interior.ColorIndex = 4
'
'                                Else
'
'                                    ApellidoReporteCm.Offset(0, 4).Interior.ColorIndex = 3
'                                    ApellidoEncontrado.Offset(0, 2).Interior.ColorIndex = 3
'                                End If
'
'                                If EstadoReservaRpt1 Like "*" & EstadoReservaRpt2 & "*" Then
'
'                                    ApellidoReporteCm.Offset(0, -6).Interior.ColorIndex = 4
'                                    ApellidoEncontrado.Offset(0, -2).Interior.ColorIndex = 4
'
'                                Else
'
'                                    ApellidoReporteCm.Offset(0, -6).Interior.ColorIndex = 3
'                                    ApellidoEncontrado.Offset(0, -2).Interior.ColorIndex = 3
'                                End If
                                
                          
Sub PasarAMayusculas()
Dim celda As Range

Range(Selection, Selection.End(xlDown)).Select
Selection.SpecialCells(xlCellTypeVisible).Select

        For Each celda In Selection
        celda = UCase(celda)
        Next

End Sub
