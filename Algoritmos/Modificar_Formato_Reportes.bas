Attribute VB_Name = "Modificar_Formato_Reportes"
Option Explicit

Private Sub SepararCeLdasCombinadas(columna1 As String, columna2 As String)

    Columns(columna1 & ":" & columna2).Select
    ActiveSheet.Shapes.Range(Array("Line 1")).Select
    Selection.UnMerge


End Sub
'***Reporte Resumido de Facturas *******
'***************************************
'Sub pruebareporteresumidodefacturas()
'Sheets("Relatório").Select
'Call separarCeldasCombinadasReporteResumidoDeFacturas
'Call seleccionarRangoYCopiarlo2("A", 11)
'Call pegarEnColumnaSeleccionada("A", 7)
'Call seleccionarRangoYCopiarlo2("C", 11)
'Call pegarEnColumnaSeleccionada("B", 7)
'Call seleccionarRangoYCopiarlo2("F", 11)
'Call pegarEnColumnaSeleccionada("D", 7)
'Call seleccionarRangoYCopiarlo2("G", 11)
'Call pegarEnColumnaSeleccionada("E", 7)
'Call seleccionarRangoYCopiarlo2("I", 11)
'Call pegarEnColumnaSeleccionada("F", 7)
'Call seleccionarRangoYCopiarlo2("K", 11)
'Call pegarEnColumnaSeleccionada("G", 7)
''ORDENAR DATOS QUE ESTAN DESCOMBINADOS
'Call OrdenarDatosDescombinados("L")
'End Sub

'**************************PARA Reporte Solo Datos*******************

Sub resumidoDeFacturasSoloDatos()
Sheets("Relatório").Select
Call crearHoja

Call seleccionarRangoYCopiarlo("A", 1)
Call pegarEnColumnaSeleccionada("A", 7)

Call seleccionarRangoYCopiarlo("C", 1)
Call pegarEnColumnaSeleccionada("B", 7)

Call seleccionarRangoYCopiarlo("D", 1)
Call pegarEnColumnaSeleccionada("D", 7)

Call seleccionarRangoYCopiarlo("E", 1)
Call pegarEnColumnaSeleccionada("E", 7)

Call seleccionarRangoYCopiarlo("F", 1)
Call pegarEnColumnaSeleccionada("F", 7)

Call seleccionarRangoYCopiarlo("G", 1)
Call pegarEnColumnaSeleccionada("G", 7)

Call seleccionarRangoYCopiarlo("H", 1)
Call pegarEnColumnaSeleccionada("H", 7)

Call seleccionarRangoYCopiarlo("I", 1)
Call pegarEnColumnaSeleccionada("I", 7)

Call seleccionarRangoYCopiarlo("J", 1)
Call pegarEnColumnaSeleccionada("J", 7)

Call seleccionarRangoYCopiarlo("K", 1)
Call pegarEnColumnaSeleccionada("L", 7)

Call seleccionarRangoYCopiarlo("L", 1)
Call pegarEnColumnaSeleccionada("M", 7)

Call seleccionarRangoYCopiarlo("M", 1)
Call pegarEnColumnaSeleccionada("N", 7)

Call seleccionarRangoYCopiarlo("O", 1)
Call pegarEnColumnaSeleccionada("Q", 7)
'CREO ENCABEZADO **(TODO VER DE USAR BUCLE O METERLO EN UNA CLASE)**
Call CrearEncabezado("A6", "Nº Factura")
Call CrearEncabezado("B6", "Nombre Huésped")
Call CrearEncabezado("D6", "reserva")
Call CrearEncabezado("E6", "Diarias")
Call CrearEncabezado("F6", "Servicios")
Call CrearEncabezado("G6", "Consulmo")
Call CrearEncabezado("H6", "Diversos")
Call CrearEncabezado("I6", "Ts.Servicio")
Call CrearEncabezado("J6", "Cobros")
Call CrearEncabezado("L6", "Saldo")
Call CrearEncabezado("M6", "Impuesto")
Call CrearEncabezado("N6", "Usuario")
Call CrearEncabezado("O5", "Motivo Anulación")
Call CrearEncabezado("Q6", "Factura Final")
Call CrearEncabezado("G2", "Resumido de Facturas de Hospedaje Emitidas")

End Sub
'VENDIS SOLO DATOS
Sub VendisSoloDatos()
Sheets("Relatório").Select
Call crearHoja

Call seleccionarRangoYCopiarlo("A", 1)
Call pegarEnColumnaSeleccionada("A", 6)

Call seleccionarRangoYCopiarlo("B", 1)
Call pegarEnColumnaSeleccionada("B", 6)

Call seleccionarRangoYCopiarlo("C", 1)
Call pegarEnColumnaSeleccionada("C", 6)

Call seleccionarRangoYCopiarlo("D", 1)
Call pegarEnColumnaSeleccionada("D", 6)

Call seleccionarRangoYCopiarlo("E", 1)
Call pegarEnColumnaSeleccionada("E", 6)

Call seleccionarRangoYCopiarlo("F", 1)
Call pegarEnColumnaSeleccionada("F", 6)

Call seleccionarRangoYCopiarlo("G", 1)
Call pegarEnColumnaSeleccionada("H", 6)

Call seleccionarRangoYCopiarlo("H", 1)
Call pegarEnColumnaSeleccionada("I", 6)

Call seleccionarRangoYCopiarlo("I", 1)
Call pegarEnColumnaSeleccionada("J", 6)

Call seleccionarRangoYCopiarlo("J", 1)
Call pegarEnColumnaSeleccionada("K", 6)

Call seleccionarRangoYCopiarlo("K", 1)
Call pegarEnColumnaSeleccionada("L", 6)

Call seleccionarRangoYCopiarlo("L", 1)
Call pegarEnColumnaSeleccionada("M", 6)

Call seleccionarRangoYCopiarlo("M", 1)
Call pegarEnColumnaSeleccionada("N", 6)

'CREO ENCABEZADO **(TODO VER DE USAR BUCLE O METERLO EN UNA CLASE)**
Call CrearEncabezado("A5", "hab")
Call CrearEncabezado("B5", "Tipo Hab")
Call CrearEncabezado("C5", "Reserva")
Call CrearEncabezado("D5", "Cuenta")
Call CrearEncabezado("E5", "Cod. Deb.")
Call CrearEncabezado("F5", "Descripción")
Call CrearEncabezado("H5", "Factura")
Call CrearEncabezado("I5", "Monto")
Call CrearEncabezado("J5", "Documento - RUT")
Call CrearEncabezado("K5", "Fecha")
Call CrearEncabezado("L5", "Hora")
Call CrearEncabezado("M5", "Usuario")
Call CrearEncabezado("N5", "Designación")
Call CrearEncabezado("H2", "Vendís de Débitos y Créditos")

End Sub

'DEMOSTRATIVO OCUPACION SOLO DATOS



Sub DemostrativoOcupacionSoloDatos()
Sheets("Relatório").Select
Call crearHoja

Call seleccionarRangoYCopiarlo("A", 1)
Call pegarEnColumnaSeleccionada("A", 6)

Call seleccionarRangoYCopiarlo("B", 1)
Call pegarEnColumnaSeleccionada("B", 6)

Call seleccionarRangoYCopiarlo("C", 1)
Call pegarEnColumnaSeleccionada("C", 6)

Call seleccionarRangoYCopiarlo("D", 1)
Call pegarEnColumnaSeleccionada("D", 6)

Call seleccionarRangoYCopiarlo("E", 1)
Call pegarEnColumnaSeleccionada("E", 6)

Call seleccionarRangoYCopiarlo("F", 1)
Call pegarEnColumnaSeleccionada("F", 6)

Call seleccionarRangoYCopiarlo("G", 1)
Call pegarEnColumnaSeleccionada("G", 6)

Call seleccionarRangoYCopiarlo("H", 1)
Call pegarEnColumnaSeleccionada("H", 6)

Call seleccionarRangoYCopiarlo("I", 1)
Call pegarEnColumnaSeleccionada("I", 6)

Call seleccionarRangoYCopiarlo("J", 1)
Call pegarEnColumnaSeleccionada("J", 6)

Call seleccionarRangoYCopiarlo("K", 1)
Call pegarEnColumnaSeleccionada("K", 6)

Call seleccionarRangoYCopiarlo("L", 1)
Call pegarEnColumnaSeleccionada("L", 6)

Call seleccionarRangoYCopiarlo("M", 1)
Call pegarEnColumnaSeleccionada("M", 6)




'CREO ENCABEZADO **(TODO VER DE USAR BUCLE O METERLO EN UNA CLASE)**
Call CrearEncabezado("A5", "Número")
Call CrearEncabezado("B5", "Status")
Call CrearEncabezado("C5", "Fch. Conf.")
Call CrearEncabezado("D5", "Gar.")
Call CrearEncabezado("E5", "Nombre Huésped")
Call CrearEncabezado("F5", "Cliente")
Call CrearEncabezado("G5", "Ad/Ni1/Ni2")
Call CrearEncabezado("H5", "HAB")
Call CrearEncabezado("I4", "Tipo de HAB")
Call CrearEncabezado("J5", "Diaria")
Call CrearEncabezado("K5", "Llegada")
Call CrearEncabezado("L5", "Partida")
Call CrearEncabezado("M5", "")
Call CrearEncabezado("E2", "Demostrativo de ocupación")

End Sub

'DATOS DE LOS HUESPEDES SOLO DATOS



Sub DatosDeLosHuespedesSoloDatos()
Sheets("Relatório").Select
Call crearHoja

Call seleccionarRangoYCopiarlo("A", 1)
Call pegarEnColumnaSeleccionada("A", 4)

Call seleccionarRangoYCopiarlo("B", 1)
Call pegarEnColumnaSeleccionada("B", 4)


Call seleccionarRangoYCopiarlo("C", 1)
Call pegarEnColumnaSeleccionada("C", 4)

Call seleccionarRangoYCopiarlo("D", 1)
Call pegarEnColumnaSeleccionada("D", 4)

Call seleccionarRangoYCopiarlo("E", 1)
Call pegarEnColumnaSeleccionada("E", 4)

Call seleccionarRangoYCopiarlo("F", 1)
Call pegarEnColumnaSeleccionada("F", 4)

Call seleccionarRangoYCopiarlo("G", 1)
Call pegarEnColumnaSeleccionada("G", 4)

Call seleccionarRangoYCopiarlo("H", 1)
Call pegarEnColumnaSeleccionada("H", 4)

Call seleccionarRangoYCopiarlo("I", 1)
Call pegarEnColumnaSeleccionada("I", 4)

Call seleccionarRangoYCopiarlo("J", 1)
Call pegarEnColumnaSeleccionada("J", 4)

Call seleccionarRangoYCopiarlo("K", 1)
Call pegarEnColumnaSeleccionada("L", 4)

Call seleccionarRangoYCopiarlo("M", 1)
Call pegarEnColumnaSeleccionada("N", 4)

Call seleccionarRangoYCopiarlo("N", 1)
Call pegarEnColumnaSeleccionada("O", 4)

Call seleccionarRangoYCopiarlo("O", 1)
Call pegarEnColumnaSeleccionada("P", 4)

Call seleccionarRangoYCopiarlo("P", 1)
Call pegarEnColumnaSeleccionada("Q", 4)

Call seleccionarRangoYCopiarlo("Q", 1)
Call pegarEnColumnaSeleccionada("R", 4)





'CREO ENCABEZADO **(TODO VER DE USAR BUCLE O METERLO EN UNA CLASE)**
Call CrearEncabezado("A3", "Check-in")
Call CrearEncabezado("B3", "Hora")
Call CrearEncabezado("C3", "Check-Out")
Call CrearEncabezado("D3", "HAB")
Call CrearEncabezado("E3", "Tipo")
Call CrearEncabezado("F3", "Nombre")
Call CrearEncabezado("G3", "Nasc.")
Call CrearEncabezado("H3", "Docum.")
Call CrearEncabezado("I3", "Orgão")
Call CrearEncabezado("J3", "Email")
Call CrearEncabezado("K3", "Dirección")
Call CrearEncabezado("L3", "Tel")
Call CrearEncabezado("N3", "Cel")
Call CrearEncabezado("O3", "Sexo")
Call CrearEncabezado("P3", "Nacion")
Call CrearEncabezado("Q3", "Profis.")
Call CrearEncabezado("R3", "Est Civ.")


Call CrearEncabezado("G2", "Dados dos Hóspedes que efetuaram check-in")

End Sub

Private Sub crearHoja()
Worksheets.Add.Name = "Hoja1"
End Sub

Private Sub CrearEncabezado(celda As String, titulo As String)
'ENCABEZADO
'A6 Nº Factura
'B6 Nombre Huésped
'D6 reserva
'E6 Diarias
'F6 Servicios
'G6 Consulmo
'H6 Diversos
'I6 Ts.Servicio
'J6 Cobros
'L6 Saldo
'M6 Impuesto
'N6 Usuario
'O5 Motivo Anulación
'Q6 Factura Final
'G2 Resumido de Facturas de Hospedaje Emitidas
Sheets("Hoja1").Select
Range(celda).Value = titulo



End Sub





Private Sub OrdenarDatosDescombinados(columna As String)
Dim ultimafila As Long
Dim celda As Range
Sheets("Relatório").Select
ultimafila = ActiveSheet.Columns(columna).Find("*", SearchOrder:=xlByRows, searchdirection:=xlPrevious).Row
Range(columna & "11:" & columna & ultimafila).Select

For Each celda In Selection

    If (celda.Value <> Empty) And (celda.Offset(0, 1).Value = Empty) Then
    
       celda.Offset(0, 1) = celda
       celda = Empty
    

    End If


Next celda

End Sub


Private Sub separarCeldasCombinadasReporteResumidoDeFacturas()
'esto funciona
    Cells.Select
    Selection.UnMerge
End Sub

Private Sub seleccionarRangoYCopiarlo(columna As String, fila As Integer)
Dim ultimafila As Long
    Sheets("Relatório").Select
'BUSCO ULTIMA FILA CON DATOS
On Error Resume Next
    ultimafila = ActiveSheet.Columns(columna).Find("*", SearchOrder:=xlByRows, searchdirection:=xlPrevious).Row
    Range(columna & fila & ":" & columna & ultimafila).Copy
End Sub



'*******Reporte Vendis******
'***************************

Sub CopiarVendisDeFormatoNuevoAFormatoViejo()
'separarCeldasCombinadasReporteVendis
'Copia de reporte nuevo a reporte viejo
' la informacion para que coincida con la macro
Call seleccionarRangoYCopiarlo("B")
Call pegarEnColumnaSeleccionada("A", 6)
Call seleccionarRangoYCopiarlo("C")
Call pegarEnColumnaSeleccionada("B", 6)
Call seleccionarRangoYCopiarlo("D")
Call pegarEnColumnaSeleccionada("C", 6)
Call seleccionarRangoYCopiarlo("E")
Call pegarEnColumnaSeleccionada("D", 6)
Call seleccionarRangoYCopiarlo("F")
Call pegarEnColumnaSeleccionada("E", 6)
Call seleccionarRangoYCopiarlo("G")
Call pegarEnColumnaSeleccionada("F", 6)
Call seleccionarRangoYCopiarlo("I")
Call pegarEnColumnaSeleccionada("H", 6)
Call seleccionarRangoYCopiarlo("J")
Call pegarEnColumnaSeleccionada("I", 6)
Call seleccionarRangoYCopiarlo("L")
Call pegarEnColumnaSeleccionada("J", 6)
Call seleccionarRangoYCopiarlo("M")
Call pegarEnColumnaSeleccionada("K", 6)
Call seleccionarRangoYCopiarlo("N")
Call pegarEnColumnaSeleccionada("L", 6)
Call seleccionarRangoYCopiarlo("P")
Call pegarEnColumnaSeleccionada("M", 6)
Call seleccionarRangoYCopiarlo("Q")
Call pegarEnColumnaSeleccionada("N", 6)
    
    
End Sub

Private Sub separarCeldasCombinadasReporteVendis()
    Columns("Q:S").Select
    ActiveSheet.Shapes.Range(Array("Line 1")).Select
    Selection.UnMerge
End Sub




'Private Sub seleccionarRangoYCopiarlo(columna As String)
'Dim ultimafila As Long
'    Sheets("Relatório").Select
''BUSCO ULTIMA FILA CON DATOS
'    ultimafila = ActiveSheet.Columns(columna).Find("*", SearchOrder:=xlByRows, searchdirection:=xlPrevious).Row
'    Range(columna & "4:" & columna & ultimafila).Copy
'End Sub

Private Sub pegarEnColumnaSeleccionada(columna As String, fila As Integer)
    Sheets("Hoja1").Select
    Range(columna & fila).Select
    On Error Resume Next
    ActiveSheet.Paste
    Application.CutCopyMode = False
End Sub


