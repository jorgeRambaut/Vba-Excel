Attribute VB_Name = "modulopruebacalendario"
Sub btnConfirmar()
'MES        CELDA
'enero      b5
'febrero    j5
'marzo      r5
'abril      b14
'mayo       j14
'junio      r14
'julio      b25
'agosto     j25
'septiembre r25
'octubre    b34
'noviembre  j34
'diciembre  r34

Dim calenadrio As Range
Dim celda As String
Dim fecha As Date
Dim primerdiadelmes As Date
Dim ultimodiadelmes As Date
Dim dia As Integer
Dim mes As Integer
Dim diadelasemana As Integer
Dim semanadelmes As Long
Dim diassumar As Integer

celda = InputBox("CELDA")
fecha = FrmFechas.txtFecha.Value
primerdiadelmes = FrmFechas.txtFecha.Value
Range("b2:x2").Value = fecha
Range("b22:x22").Value = fecha
calendario = Range(celda).Offset(0, 0).Select
ultimodiadelmes = Application.WorksheetFunction.EoMonth(CDate(fecha), 0)


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

Sub pego()
    Dim fecha As Date
    fecha = InputBox("ing fecha")
    Sheets("Diario").Select
    Range("A2").Select
    Range("A2").Value = fecha
    Range(Selection, Selection.End(xlDown)).Select
   
    Range("A2").Select
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=Range("A2:A367")
    Range("A2:A367").Select
    Selection.NumberFormat = "[$-F800]dddd, mmmm dd, yyyy"





End Sub
