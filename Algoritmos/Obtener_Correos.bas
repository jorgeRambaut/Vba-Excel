Attribute VB_Name = "Obtener_Correos"
Sub prueba()

Call obtenerMails(Worksheets("clientes").Range("A2:a41339"))


End Sub

Sub obtenerMails(rango As Range)

Dim celda As Range
Dim nombre As String
Dim mail As String
Dim i As Integer

i = 2

For Each celda In rango
        
        If celda = "Nome" Then
            nombre = celda.Offset(0, 1).Value
            Worksheets("Hoja1").Select
            Range("a" & i).Value = nombre
        End If
        
        If celda = "E-mail" Then
            mail = celda.Offset(0, 1).Value
            Worksheets("Hoja1").Select
            Range("b" & i).Value = mail
            i = i + 1
            
        End If
        
  

Next



End Sub

Sub BuscarTitulo(DatoAbuscar As String, NombreDeSolapaAbuscar _
As String, RangoInicio As String, RangoFin As String, Optional libro As String)
    Dim tituloEncontrado As Range
    Selection.SpecialCells(xlCellTypeVisible).Select
    
        
            With Worksheets(NombreDeSolapaAbuscar).Range(RangoInicio & ":" & RangoFin)
                
                    Set tituloEncontrado = .Find(DatoAbuscar, LookIn:=xlValues)
                    
                        If Not tituloEncontrado Is Nothing Then
                            tituloEncontrado.Offset(0, 0).Select
                            'Range(Selection, Selection.End(xlDown)).Select
                        End If
            End With
     
       
        
    
End Sub
