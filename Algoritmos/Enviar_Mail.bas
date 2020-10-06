Attribute VB_Name = "Enviar_Mail"
Sub enviarcorreo()

Dim i As Integer
Dim pagina1 As Worksheet
Set pagina1 = ActiveWorkbook.Worksheets("Hoja1")
Dim OutApp As Object
Dim Correo As Object
Dim ContenidoMail As String
Dim pie_de_firma As String
Dim mi_Nombre As String
Dim nombre As String

i = 2

nombre = pagina1.Range("A" & i).Value

mi_Nombre = InputBox("Ingresar Nombre")

pie_de_firma = "<h4>" & mi_Nombre & "</h4>" _
& "<p>Ejecutivo de Reservas</p>" _
& "<p>Loi Suites Hoteles</p>" _
& "<p>+54-11-5777-8950 int 3400 </p>" _
& "<a  href=" & "mailto:reservas4@loisuites.com.ar >" & "reservas4@loisuites.com.ar</a>" & " | " & "<a  href=" & "https://www.loisuites.com.ar >" & "www.loisuites.com.ar</a>"


With Application
.EnableEvents = False
.ScreenUpdating = False
End With
'Comprobar si Outlook esta abierto y en caso de no estarlo abrirlo
On Error Resume Next
Set OutApp = GetObject("", "Outlook.Application")
Err.Clear
If OutApp Is Nothing Then Set OutApp = CreateObject("Outlook.Application")
OutApp.Visible = True
'Set Correo = OutApp.CreateItem(0)


For Each celda In Selection
Set Correo = OutApp.CreateItem(0)
nombre = pagina1.Range("A" & i).Value

ContenidoMail = "<p>Estimado/a" & nombre & "</p>" _
& " <p>Mi nombre es " & mi_Nombre & " le escribo del equipo comercial de Loi Suites Hoteles, esperando que usted y sus seres queridos se encuentren muy bien en estos momentos. </p>" _
& " <p>Nos encantaría poder acercarles las propuestas de alojamiento que estamos ofreciendo para cuando requieran realizar viajes ya sea en la ciudad de Buenos Aires, Iguazú o San Martin de los Andes.</p>" _
& "<p>Nos complace poder  asesorarlos con propuestas que se adapten a sus necesidades, brindando una experiencia única.</p>" _
& "<p>Hemos incorporado en nuestros hoteles los protocolos de seguridad e higiene establecidos por el Gobierno Nacional y la OMS para garantizarles una estancia segura.<p>" _
& "<p>Por favor en caso de estar interesado en recibir nuestras propuestas o tengan algún requerimiento en particular, me encuentro a disposición para poder asesorarlos.<p>" _
& "<p>Se adjunta protocolo de seguridad y el <i><a  href=" & "https://www.loisuites.com.ar >link</a></i> de nuestra web para conocer en detalles nuestros hoteles.</p> " _
& "<p> <b>Es importante saber que todas las reservas en forma directa cuentan con descuentos exclusivos.</b></p>" _
& "<p>Un cordial saludo </p>"

'Crear el correo y mostrarlo
    With Correo
        .To = pagina1.Range("B" & i).Value
        '.CC = pagina1.Range("C9").Value
        .Subject = "Contacto Prueba"
        .HTMLBody = "<HTML><BODY> " & ContenidoMail & " <FOOTER>" & pie_de_firma & "</FOOTER></BODY></HTML>"
        .Display
        '.Send
    End With
    With Application
    .EnableEvents = True
    .ScreenUpdating = True
    End With
 
 i = i + 1
Next

End Sub
