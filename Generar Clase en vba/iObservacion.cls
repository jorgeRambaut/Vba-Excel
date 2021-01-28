VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iObservacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function observacion(tipoDealojamiento As String, Ota As String, condicion_reserva As String) As String

    observacion = Ota & " " & condicion_reserva & " " & forma_de_pago & " " & beneficios

End Function


Sub crearobservacion(TipoDeReembolso As String)


End Sub

Sub agregarMenoresEnObservacion()



End Sub

Sub agregarJuntoConEnObservacion()




End Sub

Sub agregarPromoWebLoi()
 

End Sub
