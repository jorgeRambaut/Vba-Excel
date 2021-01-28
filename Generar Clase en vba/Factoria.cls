VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Factoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function Factoria(ByRef className As String) As Object
    Select Case className
    Case "Reserva": Set Factoria = New ClsReserva
    Case Else: MsgBox ("No se reconoce el nombre de la clase.")
    End Select
End Function
