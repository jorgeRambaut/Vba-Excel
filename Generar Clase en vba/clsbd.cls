VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsbd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private ptipoDeContacto As String
Private pnombre As String
Private papellido As String
Private pdni As String
Private pemail As String
Private ptelefono As String
Private pcumpleaños As String
Private pcalle As String
Private pnumero As String
Private pentreCalles As String
Private pcp As String
Private pbarrio As String
Private plocalidad As String
Private ppartido As String
Private pprovincia As String
Private pfacebook As String
Private pinstagram As String
Private pwhatsapp As String
Private ptelegram As String
Private pcompro As String


'metodos get let set

Public Property Get tipoDeContacto() As String
    tipoDeContacto = ptipoDeContacto
End Property

Public Property Let tipoDeContacto(valor As String)
    While valor = Empty
        MsgBox ("Tipo de Contacto Campo Obligatorio")
        valor = InputBox("Ingresar Tipo de Contacto")
        formulario.txttipodecontacto.Value = valor
    Wend
        ptipoDeContacto = valor
   
End Property

Public Property Get nombre() As String
    nombre = pnombre
End Property

Public Property Let nombre(valor As String)
    While valor = Empty
        MsgBox ("Nombre Campo Obligatorio")
        valor = InputBox("Ingresar Nombre")
    Wend
    pnombre = valor
End Property

Public Property Get apellido() As String
    apellido = papellido
End Property

Public Property Let apellido(valor As String)
    While valor = Empty
        MsgBox ("Apellido Campo Obligatorio")
        valor = InputBox("Ingresar Apellido")
    Wend
        papellido = valor
    
End Property

Public Property Get dni() As String
    dni = pdni
End Property

Public Property Let dni(valor As String)
    While valor = Empty
        MsgBox ("Dni Campo Obligatorio")
        valor = InputBox("Ingresar Dni")
    Wend
        pdni = valor
    
End Property

Public Property Get email() As String
    email = pemail
End Property

Public Property Let email(valor As String)
    While valor = Empty
        MsgBox ("Email Campo Obligatorio")
        valor = InputBox("Ingresar Email")
    Wend
        pemail = valor
    
End Property

Public Property Get telefono() As String
    telefono = ptelefono
End Property

Public Property Let telefono(valor As String)
    While valor = Empty
        MsgBox ("Telefono Campo Obligatorio")
        valor = InputBox("Ingresar Telefono")
    Wend
        ptelefono = valor
    
End Property

Public Property Get cumpleaños() As String
    cumpleaños = pcumpleaños
End Property

Public Property Let cumpleaños(valor As String)
    While valor = Empty
        MsgBox ("Cumpleaños Campo Obligatorio")
        valor = InputBox("Ingresar Cumpleaños dd/mm/aaaa")
    Wend
    pcumpleaños = valor
    
    
End Property


'no los valido porque me hinche las bolas y me chupa un huevo la direccion

Public Property Get calle() As String
    calle = pcalle
End Property

Public Property Let calle(valor As String)
    pcalle = valor
End Property

Public Property Get numero() As String
    numero = pnumero
End Property

Public Property Let numero(valor As String)
    pnumero = valor
End Property

Public Property Get entreCalles() As String
    entreCalles = pentreCalles
End Property

Public Property Let entreCalles(valor As String)
    pentreCalles = valor
End Property

Public Property Get cp() As String
    cp = pcp
End Property

Public Property Let cp(valor As String)
    pcp = valor
End Property

Public Property Get barrio() As String
    barrio = pbarrio
End Property

Public Property Let barrio(valor As String)
    pbarrio = valor
End Property


Public Property Get localidad() As String
    barrio = plocalidad
End Property

Public Property Let localidad(valor As String)
    plocalidad = valor
End Property

Public Property Get partido() As String
    partido = ppartido
End Property

Public Property Let partido(valor As String)
    ppartido = valor
End Property

Public Property Get provincia() As String
    provincia = pprovincia
End Property

Public Property Let provincia(valor As String)
    pprovincia = valor
End Property

Public Property Get facebook() As String
    facebook = pfacebook
End Property

Public Property Let facebook(valor As String)
    While valor = Empty
        MsgBox ("Facebook Campo Obligatorio")
        valor = InputBox("Ingresar Facebook")
    Wend
        pfacebook = valor
    
End Property

Public Property Get instagram() As String
    instagram = pinstagram
End Property

Public Property Let instagram(valor As String)
    While valor = Empty
        MsgBox ("Instagram Campo Obligatorio")
        valor = InputBox("Ingresar Instagram")
    Wend
        pinstagram = valor
 
End Property
Public Property Get whatsapp() As String
    whatsapp = pwhatsapp
End Property

Public Property Let whatsapp(valor As String)
   While valor = Empty
        MsgBox ("Whatsapp Campo Obligatorio")
        valor = InputBox("Ingresar Whatsapp")
    Wend
        pwhatsapp = valor
    
End Property
Public Property Get telegram() As String
    telegram = ptelegram
End Property

Public Property Let telegram(valor As String)
    While valor = Empty
        MsgBox ("Telegram Campo Obligatorio")
        valor = InputBox("Ingresar Telegram")
    Wend
        ptelegram = valor
    
End Property

Public Property Get compro() As String
    compro = pcompro
End Property

Public Property Let compro(valor As String)
   While valor = Empty
        MsgBox ("Compro Campo Obligatorio")
        valor = InputBox("Ingresar Si compro si o no")
    Wend
        pcompro = valor
    
End Property


Sub bd_Ir_BD()
'
' Ir_BD Macro
'
    Sheets("BD").Select
End Sub


Sub detenerPantalla()

Application.ScreenUpdating = False


End Sub

Sub seleccionarbd()

Sheets("BD").Select
Range("A3").EntireRow.Insert
Sheets("Registro").Select

End Sub

Sub seleccionar_ultima_Fila_Vacia()
Range("a1").Select
Do Until IsEmpty(ActiveCell) And IsEmpty(ActiveCell.Offset(1, 0))
ActiveCell.Offset(1, 0).Select
Loop

End Sub


Sub cargardatosenBD() 'tengo que ver de separar estas dos cosas
Dim Datos As New clsbd
    With Datos
    .nombre = formulario.txtNombre.Value
    .apellido = formulario.txtApellido.Value
    .calle = formulario.txtDireccion.Value
    .tipoDeContacto = formulario.txttipodecontacto.Value
    .telefono = formulario.txttelefono.Value
    .cumpleaños = formulario.txtcumpleaños.Value
    .facebook = formulario.txtfacebook.Value
    .instagram = formulario.txtinstagram.Value
    .whatsapp = formulario.txtwhatsapp.Value
    .telegram = formulario.txttelegram.Value
    .compro = formulario.txtcompro.Value
    End With
    
    ActiveCell.Offset(0, 0) = Datos.nombre
    ActiveCell.Offset(0, 1) = Datos.apellido
    ActiveCell.Offset(0, 2) = Datos.calle
    ActiveCell.Offset(0, 3) = Datos.tipoDeContacto
    ActiveCell.Offset(0, 4) = Datos.telefono
    ActiveCell.Offset(0, 5) = Datos.cumpleaños
    ActiveCell.Offset(0, 6) = Datos.facebook
    ActiveCell.Offset(0, 7) = Datos.instagram
    ActiveCell.Offset(0, 8) = Datos.whatsapp
    ActiveCell.Offset(0, 9) = Datos.telegram
    ActiveCell.Offset(0, 10) = Datos.compro
    
End Sub


