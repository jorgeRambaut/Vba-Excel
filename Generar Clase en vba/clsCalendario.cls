VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pcelda As String
Private pfecha As Date
Private pprimerdiadelmes As Date
Private pultimodiadelmes As Date
Private pdia As Integer
Private pmes As Integer
Private pdiadelasemana As Integer
Private psemanadelmes As Long
Private pdiasasumar As Integer




Public Property Get celda() As String
    celda = pcelda
End Property

Public Property Let celda(valor As String)
    
While valor <> ""
    If WorksheetFunction.IsText(valor) Then

        pcelda = valor
    
    Else
        MsgBox ("Ingresar Rango de CELDAS")
        valor = InputBox("INGRESAR RANGO DE CELDAS")
        
    End If
    
 Loop
    
End Property


Public Property Get fecha() As Date
    fecha = pfecha
End Property

Public Property Let fecha(valor As Date)
    
While valor <> ""
    If WorksheetFunction.IsDate(valor) Then

        pfecha = valor
    
    Else
        MsgBox ("Ingresar Fecha dd/mm/aaaa")
        valor = InputBox("INGRESAR Fecha dd/mm/aaaa")
        
    End If
    
 Loop
    
End Property

Public Property Get primerdiadelmes() As Date
    primerdiadelmes = pprimerdiadelmes
End Property

Public Property Let primerdiadelmes(valor As Date)
    
While valor <> ""
    If WorksheetFunction.IsDate(valor) Then

        pprimerdiadelmes = valor
    
    Else
        MsgBox ("Ingresar Fecha dd/mm/aaaa")
        valor = InputBox("INGRESAR Fecha dd/mm/aaaa")
        
    End If
    
 Loop
    
End Property

Function ultimodiadelmes(fecha As Date)

ultimodiadelmes = Application.WorksheetFunction.EoMonth(CDate(fecha), 0)

End Function

Public Property Get dia() As Date
    dia = pdia
End Property

Public Property Let dia(valor As Date)
    
While valor <> ""
    If WorksheetFunction.IsDate(valor) Then

        pdia = valor
    
    Else
        MsgBox ("Ingresar Fecha dd/mm/aaaa")
        valor = InputBox("INGRESAR Fecha dd/mm/aaaa")
        
    End If
    
 Loop
    
End Property

Public Property Get mes() As Date
    mes = pmes
End Property

Public Property Let mes(valor As Date)
    
While valor <> ""
    If WorksheetFunction.IsDate(valor) Then

        pmes = valor
    
    Else
        MsgBox ("Ingresar Fecha dd/mm/aaaa")
        valor = InputBox("INGRESAR Fecha dd/mm/aaaa")
        
    End If
    
 Loop
    
End Property


Public Property Get diasasumar() As Integer
    diasasumar = pdiasasumar
End Property

Public Property Let dias(valor As Integer)
    
While valor <> ""
    If WorksheetFunction.IsNumber(valor) Then

        pdias = valor
    
    Else
        MsgBox ("Ingresar VALOR DIA ... 1  ....2")
        valor = InputBox("INGRESAR VALOR")
        
    End If
    
 Loop
    
End Property

Function diadelasemana() As Integer

diadelasemana = DatePart("w", fecha)

End Function


Function semanadelmes() As Integer

semanadelmes = DatePart("ww", fecha) - DatePart("ww", primerdiadelmes) + 1

End Function

Sub CompletarCalendario(fecha As Date, ultimodiadelmes As Date)

Dim diasasumar As Integer
Dim diadelasemana As Integer
Dim semanadelmes As Long

While CDate(fecha) < CDate(ultimodiadelmes)
    fecha = DateAdd("d", diasasumar, fecha)
    'MsgBox (fecha)
    diadelasemana = DatePart("w", fecha)
    semanadelmes = DatePart("ww", fecha) - DatePart("ww", primerdiadelmes) + 1
    Range(celda).Offset(semanadelmes, diadelasemana - 1).Select
    Range(celda).Offset(semanadelmes, diadelasemana - 1).Value = fecha
    diasasumar = 1
Wend

End Sub

Sub moverse(celda As String)

Range(celda).Offset(0, 0).Select


End Sub

Sub BuscarTitulo(DatoAbuscar As String, NombreDeSolapaAbuscar As String, RangoInicio As String, RangoFin As String, Optional libro As String)
    Dim tituloEncontrado As Range
      Selection.SpecialCells(xlCellTypeVisible).Select
        
            With Worksheets(NombreDeSolapaAbuscar).Range(RangoInicio & ":" & RangoFin)
            
                    Set tituloEncontrado = .Find(DatoAbuscar, LookIn:=xlValues)
                    
                        If Not tituloEncontrado Is Nothing Then
                            tituloEncontrado.Offset(1, 0).Select
                            'Range(Selection, Selection.End(xlDown)).Select
                        End If
            End With
    
End Sub

'Private Sub btnConfirmar_Click()
'''enero b5
'''febrero j5
'''marzo r5
'''abril b14
'''mayo j14
'''junio r14
'''julio b25
'''agosto j25
'''septiembre r25
'''octubre b34
'''noviembre j34
'''diciembre r34
'Dim calenadrio As Range
'Dim celda As String
'Dim fecha As Date
'Dim primerdiadelmes As Date
'Dim ultimodiadelmes As Date
'Dim dia As Integer
'Dim mes As Integer
'Dim diadelasemana As Integer
'Dim semanadelmes As Long
'Dim diassumar As Integer
'
'celda = InputBox("CELDA")
'fecha = FrmFechas.txtFecha.Value
'primerdiadelmes = FrmFechas.txtFecha.Value
''Range("b2:x2").Value = fecha
''Range("b22:x22").Value = fecha
'calendario = Range(celda).Offset(0, 0).Select
'ultimodiadelmes = Application.WorksheetFunction.EoMonth(CDate(fecha), 0)
'
'
'While CDate(fecha) < CDate(ultimodiadelmes)
'fecha = DateAdd("d", diasasumar, fecha)
''MsgBox (fecha)
'diadelasemana = DatePart("w", fecha)
'semanadelmes = DatePart("ww", fecha) - DatePart("ww", primerdiadelmes) + 1
'' esto funciona
''WeekNumberFromDate = Int(((dt - StartDate) + 6) / 7) + Abs(Weekday(dt) = Weekday(StartDate))
''MsgBox (semanadelmes)
''MsgBox (semanadelmes)
'Range(celda).Offset(semanadelmes, diadelasemana - 1).Select
'Range(celda).Offset(semanadelmes, diadelasemana - 1).Value = fecha
'diasasumar = 1
'Wend
'End Sub
