VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formulario 
   Caption         =   "Formulario"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9900
   OleObjectBlob   =   "formulario.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "formulario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub cmdAgregar_Click()
'creo objeto de tipo clsbd
Dim datoscontacto As New clsbd
datoscontacto.seleccionar_ultima_Fila_Vacia
datoscontacto.cargardatosenBD
End Sub



