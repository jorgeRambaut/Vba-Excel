Attribute VB_Name = "Enviar_Mail"
Sub Correo_Pax_Directo()
'Variables ------------------------
Dim i As Integer
Dim celda As Range
Dim pagina1 As Worksheet
Dim OutApp As Object
Dim Correo As Object
Dim ContenidoMail As String
Dim pie_de_firma As String
Dim mi_Nombre As String
Dim contacto As String
Dim interno As String
Dim mail As String
Dim solapa_Excel As String
Dim myAttachments As Outlook.Attachments
'----------------------------------------
i = 1



'pido datos para llenar el pie de firma y el cuerpo del correo y lo convierto a tipo oracion.
mi_Nombre = StrConv(InputBox("Ingresar Nombre y Apellido Del Ejecutivo"), vbProperCase)
interno = InputBox("Interno Del Ejecutivo")
mail = InputBox("Ingresar Mail Del Ejecutivo")
solapa_Excel = StrConv(InputBox("Elegir Solapa Eze/Bren/George/Mati"), vbProperCase)

'asigno el nombre de la pagina que voy a usar
Set pagina1 = ActiveWorkbook.Worksheets(solapa_Excel)

'pie de firma ---------------------------------------
pie_de_firma = "<h4>" & mi_Nombre & "</h4>" _
& "<p>Ejecutivo de Reservas</p>" _
& "<p>Loi Suites Hoteles</p>" _
& "<p>+54-11-5777-8950 int " & interno & "</p>" _
& "<a  href=" & mail & ">" & mail & "</a>" & " | " & "<a  href=" & "https://www.loisuites.com.ar >" & "www.loisuites.com.ar</a>"
'-----------------------------------------------------------

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

'recorro las celdas y envio el correo
For Each celda In Selection

Set Correo = OutApp.CreateItem(0)

'convierto a tipo oracion
contacto = StrConv(pagina1.Range("A" & i).Value, vbProperCase)

'adjunto el archivo
Set myAttachments = Correo.Attachments
 myAttachments.Add "C:\Users\JORGE\Desktop\covid\Loi Suites_covid_flyer.png", _
 olByValue, 1, "Medidas De Higiene Y Seguridad"
 myItem.Display
 
ContenidoMail = "<p>Buenos Dias, " & contacto & "</p>" _
& " <p>Lo saluda " & mi_Nombre & " parte del Equipo Comercial de Loi Suites Hoteles, esperando que usted y sus seres queridos se encuentren bien en estos momentos.</p>" _
& " <p>Me complace poder acercarle las propuestas de alojamiento que ofrecemos en Loi Suites Hoteles, caso requieran realizar viajes en la Ciudad de Buenos Aires, Iguazú o San Martin de los Andes.</p>" _
& "<p>Hemos incorporado los protocolos de seguridad e higiene establecidos por el Gobierno Nacional y la OMS para garantizarle una estadía segura. En el presente correo, podrá observar los protocolos implementados, y el <a  href=" & "https://www.loisuites.com.ar > acceso a nuestra web</a> para conocer en detalle nuestras propiedades.</p>" _
& "<p>Es importante destacar que las reservas gestionadas de forma directa con el hotel cuentan con <b>descuentos exclusivos</b>. </p>" _
& "<p>Desde ya, muchas gracias por su tiempo. Si desea recibir mayor información, me encuentro a disposición para poder asistirlo.</p>" _
& "<p>Un cordial saludo.</p>"

'Crear el correo y mostrarlo
    With Correo
        .To = pagina1.Range("B" & i).Value
        '.CC = pagina1.Range("C9").Value
        .Subject = "Propuesta de Alojamiento " & contacto
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

Sub Correo_Pax_Corporativo()
'Variables ------------------------
Dim i As Integer
Dim celda As Range
Dim pagina1 As Worksheet
Dim OutApp As Object
Dim Correo As Object
Dim ContenidoMail As String
Dim pie_de_firma As String
Dim mi_Nombre As String
Dim contacto As String
Dim interno As String
Dim mail As String
Dim solapa_Excel As String
Dim myAttachments As Outlook.Attachments
'----------------------------------------
i = 1



'pido datos para llenar el pie de firma y el cuerpo del correo y lo convierto a tipo oracion.
mi_Nombre = StrConv(InputBox("Ingresar Nombre y Apellido Del Ejecutivo"), vbProperCase)
interno = InputBox("Interno Del Ejecutivo")
mail = InputBox("Ingresar Mail Del Ejecutivo")
solapa_Excel = StrConv(InputBox("Elegir Solapa Eze/Bren/George/Mati"), vbProperCase)

'asigno el nombre de la pagina que voy a usar
Set pagina1 = ActiveWorkbook.Worksheets(solapa_Excel)

'pie de firma ---------------------------------------
pie_de_firma = "<h4>" & mi_Nombre & "</h4>" _
& "<p>Ejecutivo de Reservas</p>" _
& "<p>Loi Suites Hoteles</p>" _
& "<p>+54-11-5777-8950 int " & interno & "</p>" _
& "<a  href=" & mail & ">" & mail & "</a>" & " | " & "<a  href=" & "https://www.loisuites.com.ar >" & "www.loisuites.com.ar</a>"
'-----------------------------------------------------------

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

'recorro las celdas y envio el correo
For Each celda In Selection

Set Correo = OutApp.CreateItem(0)

'convierto a tipo oracion
contacto = StrConv(pagina1.Range("A" & i).Value, vbProperCase)

'adjunto el archivo
Set myAttachments = Correo.Attachments
 myAttachments.Add "C:\Users\JORGE\Desktop\covid\Loi Suites_covid_flyer.png", _
 olByValue, 1, "Medidas De Higiene Y Seguridad"
 myItem.Display
 
ContenidoMail = "<p>Buenos Dias, " & contacto & "</p>" _
& " <p>Lo saluda " & mi_Nombre & " parte del Equipo Comercial de Loi Suites Hoteles, esperando que usted y sus seres queridos se encuentren bien en estos momentos.</p>" _
& " <p>Quisiera conocer si el personal de su empresa (funcionarios, colaboradores, gerentes, CEOs) suelen utilizar servicios de hotelería, para acercarle una propuesta comercial de alojamiento en nuestros hoteles Loi Suites localizados en la Ciudad de Buenos Aires.</p>" _
& "<p>Contamos con dos hoteles en la ciudad, <b>Loi Suites Recoleta</b> y <b>Loi Suites Esmeralda</b> que han implementado loa protocolos de seguridad e higiene establecidos por el Gobierno Nacional y la OMS para garantizarle una estadía segura. En el presente correo, podrá observar los protocolos implementados, y el <a  href=" & "https://www.loisuites.com.ar >acceso a nuestra web</a> para conocer en detalle nuestras propiedades.</p>" _
& "<p><b>Loi Suites Esmeralda</b></p>" _
& "<p>Se ubica en zona Microcentro (Marcelo T. de Alvear 842. Cuenta con las comodidades de un departamento (habitación + kitchenette) y los servicios de un hotel como son limpieza diaria, seguridad 24hs, acceso WiFi y desayuno.</p>" _
& "<p><a href =" & "https://drive.google.com/drive/folders/1ATePDzJMGb8jChGk06NkU_FCB6YVXymX> Click aquí para conocer el hotel.</a></p>" _
& "<p><b>Loi Suites Recoleta</b></p>" _
& "<p>Emplazado en el corazón de uno de los barrios de mayor prestigio de Buenos Aires, a pasos del histórico Cementerio de Recoleta (Vicente López 1955), es un hotel 5 estrellas que se distingue por su maravilloso jardín de invierno de 400m² con piscina climatizada que ofrece un espacio único de luz y vegetación para disfrutar durante su estadía. Ofrece desayuno, gimnasio y sauna y servicio exclusivo de atención al huésped.</p>" _
& "<p><a href =" & "https://drive.google.com/drive/folders/0B5yn0ieZMlx-Nklvb1ZaLUlqS3M> Click aquí para conocer el hotel.</a></p>" _
& "<p>Desde ya, muchas gracias por su tiempo. Si desea recibir mayor información, me encuentro a disposición para poder asistirlo.</p>" _
& "<p>Quedamos atentos a vuestros comentarios.</p>" _
& "<p>Saludos cordiales.</p>"


'Crear el correo y mostrarlo
    With Correo
        .To = pagina1.Range("B" & i).Value
        '.CC = pagina1.Range("C9").Value
        .Subject = "Propuesta de Alojamiento " & contacto
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

'Sub Adjuntar()
' Dim myAttachments As Outlook.Attachments
'
' Set myAttachments = Correo.Attachments
' myAttachments.Add "C:\Users\JORGE\Desktop\covid\Loi Suites_covid_flyer.png", _
' olByValue, 1, "Medidas De Higiene Y Seguridad"
' myItem.Display
'
'End Sub

