VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClaseReserva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Pnombre As String
Private PTipoHabtiacion As String
Private PCanalReserva As String
Private PStatus As String
Private Pdireccion As String
Private Pmail As String
Private Ptelefono As String
Private PRatePlan As String
Private Pobservaciones As String
Private PRequerimientos As String
Private PComentarios As String
Private Pfechain As Date
Private Pfechaout As Date
Private PFechaReserva As Date
Private PNumeroDeconfirmacion As String
Private Padultos As Integer
Private PNiños As Integer
Private PCantidadDeHabitaciones As Integer
Private Pprecio As Double
Private PTotal As Double
Private Pubicacion As Range

Public Property Get ubicacion() As Range
    ubicacion = Pubicacion
End Property

Public Property Let ubicacion(valor As Range)
    Pubicacion = valor
End Property


Public Property Get nombre() As String
    nombre = Pnombre
End Property

Public Property Let nombre(valor As String)
    Pnombre = valor
End Property

Public Property Get TipoHabtiacion() As String
    TipoHabtiacion = PTipoHabtiacion
End Property

Public Property Let TipoHabtiacion(valor As String)
    PTipoHabtiacion = valor
End Property

Public Property Get CanalReserva() As String
    CanalReserva = PCanalReserva
End Property

Public Property Let CanalReserva(valor As String)
    PCanalReserva = valor
End Property

Public Property Get Status() As String
    Status = PStatus
End Property

Public Property Let Status(valor As String)
    PStatus = valor
End Property

Public Property Get direccion() As String
    direccion = Pdireccion
End Property

Public Property Let direccion(valor As String)
    Pdireccion = valor
End Property

Public Property Get mail() As String
    mail = Pmail
End Property

Public Property Let mail(valor As String)
    Pmail = valor
End Property

Public Property Get telefono() As String
    telefono = Ptelefono
End Property

Public Property Let telefono(valor As String)
    Ptelefono = valor
End Property

Public Property Get RatePlan() As String
    RatePlan = PRatePlan
End Property

Public Property Let RatePlan(valor As String)
    PRatePlan = valor
End Property

Public Property Get observaciones() As String
    observaciones = Pobservaciones
End Property

Public Property Let observaciones(valor As String)
    Pobservaciones = valor
End Property

Public Property Get Requerimientos() As String
    Requerimientos = PRequerimientos
End Property

Public Property Let Requerimientos(valor As String)
    PRequerimientos = valor
End Property

Public Property Get Comentarios() As String
    Comentarios = PComentarios
End Property

Public Property Let Comentarios(valor As String)
    PComentarios = valor
End Property

Public Property Get fechain() As Date
    fechain = Pfechain
End Property

Public Property Let fechain(valor As Date)
    Pfechain = valor
End Property

Public Property Get fechaout() As Date
    fechaout = Pfechaout
End Property

Public Property Let fechaout(valor As Date)
    Pfechaout = valor
End Property

Public Property Get FechaReserva() As Date
    FechaReserva = PFechaReserva
End Property

Public Property Let FechaReserva(valor As Date)
    PFechaReserva = valor
End Property

Public Property Get NumeroDeconfirmacion() As String
    NumeroDeconfirmacion = PNumeroDeconfirmacion
End Property

Public Property Let NumeroDeconfirmacion(valor As String)
    PNumeroDeconfirmacion = valor
End Property

Public Property Get adultos() As Integer
    adultos = Padultos
End Property

Public Property Let adultos(valor As Integer)
    Padultos = valor
End Property

Public Property Get Niños() As Integer
    Niños = PNiños
End Property

Public Property Let Niños(valor As Integer)
    PNiños = valor
End Property
Public Property Get CantidadDeHabitaciones() As Integer
    CantidadDeHabitaciones = PCantidadDeHabitaciones
End Property
Public Property Let CantidadDeHabitaciones(valor As Integer)
    PCantidadDeHabitaciones = valor
End Property
Public Property Get Precio() As Double
    Precio = Pprecio
End Property
Public Property Let Precio(valor As Double)
    Pprecio = valor
End Property

Public Function CalculaTotal(CantidadDeHabitaciones As Integer, Precio As Double, CantidadDeNoches As Integer) As Double

CalculaTotal = CantidadDeHabitaciones * Precio * CantidadDeNoches
End Function

Public Function CalculaNoches(fechain As Date, fechaout As Date)
Dim CantidadDeNoches As Integer
CalculaNoches = fechaout - fechain
CalculaNoches = CInt(CantidadDeNoches)
End Function

Public Function ObtenerPosicion(nombre)
Dim Posicion As Integer
Dim Coma As String
Coma = ","
        Selection.SpecialCells(xlCellTypeVisible).Select
        apellido = nombre
        Posicion = InStr(1, apellido, Coma, 1)
        If Posicion = 0 Then
        Exit Function
        Else
        apellido = Left(apellido, Posicion - 1)
        Pnombre = apellido
        End If
End Function

Public Function hojalibro(nombrelibro As String, hojalibro As String, columna As String)
Dim m_wbBook As Workbook
Dim m_wsSheet As Worksheet
Dim m_rnCheck As Range
    Set m_wbBook = Workbooks(nombrelibro)
    Set m_wsSheet = m_wbBook.Sheets(hojalibro)
    Set m_rnCheck = m_wsSheet.Range(columna & ":" & columna)
    
End Function

Public Function BuscarApellido(apellido As String, m_rnCheck As Range)

Dim m_stAddress As String

Set rangoBusqueda = reserva.Find(What:=apellido, LookAt:=xlPart).Select

        If Not rangoBusqueda Is Nothing Then
        
                    Do
                    
                    m_stAddress = m_rnFind.Address
                    
                    Pnombre = apellido
                                     
                    Set rangoBusqueda = rangoBusqueda.FindNext(rangoBusqueda)
        
                    Loop While Not rangoBusqueda Is Nothing And rangoBusqueda.Address <> m_stAddress
        End If
    



End Function




































