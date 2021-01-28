Attribute VB_Name = "Modulo_Control_No_Shows"
Sub Control_NoShows_NoReembosables()

'lo que hay que seleccionar en el reporte

'    Status Reserva
'    Fecha Reserva
'    Valor de la Diaria
'    Cliente
'    Categoría de Tarifa
'    Tipo HAB
'    Huésped
'    Número Reserva
'    Tipo de Huésped
'    Fecha Llegada
'    Fecha Salida
'    Ctde Atend n Efect
'    Ctde Reservas
'    Total Receta
'    Ctde.Pernoctes

'TODO VER FUNCIONES CON ARGUMENTOS

Call Filtro("No-Show", 2)
Call NoReembolsable

MsgBox ("Aplicar Filtro en columna tipo de huesped y Seleccionar No Reembolsable" _
 & vbNewLine & "Ver Columna Q")

End Sub

Sub Filtro(criterio As String, field As Integer)

'DECLARO FILA COMO LONG
Dim ultimaFila As Long
      
      'BUSCO ULTIMA FILA CON DATOS
        ultimaFila = ActiveSheet.Columns("A").Find("*", _
        SearchOrder:=xlByRows, searchdirection:=xlPrevious).Row
                       
        'SELECCIONO RANGO PARA APLICAR FILTRO
         Range("A1:P" & ultimaFila & "").Select
         
         'FILTRO POR CRITERIO Y FIELD PASADO COMO PARAMETROS
            Selection.AutoFilter
            ActiveSheet.Range("$A$1:P" & ultimaFila).AutoFilter field:=field, Criteria1:=criterio
End Sub

Sub NoReembolsable()
'Declaro variables
Dim NRF As String
Dim tipoHuesped As String
Dim NoReembolsable As Range
Dim fila As Integer

'defino Variables
 tipoHuesped = "J1" 'ubicacion de columna tipo de huesped
 NRF = "NO REEMBOLSABLE"
 
'-----------------automatico -------------------------------------
Range("Q1").Value = "Cantidad de noches a cobrar" 'DONDE COLOCO LOS DATOS
ActiveSheet.Range(tipoHuesped).Select 'SELECCIONO COLUMNA TIPO HUESPED
ActiveCell.Offset(1, 0).Select 'BAJO UNA CELDA
Range(Selection, Selection.End(xlDown)).Select 'SELECCIONO ULTIMA FILA
Selection.SpecialCells(xlCellTypeVisible).Select ' POR SI HAY FILTROS APLICADOS

'---------------fin automatico-----------------------------------


'---------------Recorro y cargo cuantas noches se tiene que cobrar---------------

    For Each NoReembolsable In Selection
    
        If (NoReembolsable.Offset(0, 0).Value = NRF) Then
          NoReembolsable.Offset(0, 7).Value = NoReembolsable.Offset(0, 6).Value
        End If
     
    Next

'--------------------------------------------------------------------------------
 
End Sub

