Attribute VB_Name = "palabras"
Private Sub ContarPalabras()


Dim tamaño As Integer
Dim i As Integer
Dim f As Integer
Dim a As Long


Dim celda As Range
Dim texto As String


        For Each celda In Selection

             tamaño = Len(celda)
             texto = celda

             
             f = Asc(celda)
             a = InStr(1, celda, f, 0)
             
             
                If f < 97 Then
                MsgBox ("algo")
                End If
            


        Next




End Sub
