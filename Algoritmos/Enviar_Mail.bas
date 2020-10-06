Attribute VB_Name = "Enviar_Mail"
Sub enviar_correo_todos_los_hoteles()

Dim i As Integer
Dim pagina1 As Worksheet
Set pagina1 = ActiveWorkbook.Worksheets("Hoja1")
Dim OutApp As Object
Dim Correo As Object
Dim ContenidoMail As String
Dim pie_de_firma As String
Dim mi_Nombre As String
Dim nombre As String
Dim interno As String
Dim mail As String
Dim myAttachments As Outlook.Attachments

i = 2

nombre = pagina1.Range("A" & i).Value

mi_Nombre = InputBox("Ingresar Nombre y Apellido Del Ejecutivo")
interno = InputBox("Interno Del Ejecutivo")
mail = InputBox("Ingresar Mail Del Ejecutivo")

pie_de_firma = "<h4>" & mi_Nombre & "</h4>" _
& "<p>Ejecutivo de Reservas</p>" _
& "<p>Loi Suites Hoteles</p>" _
& "<p>+54-11-5777-8950 int " & interno & "</p>" _
& "<a  href=" & mail & ">" & mail & "</a>" & " | " & "<a  href=" & "https://www.loisuites.com.ar >" & "www.loisuites.com.ar</a>"


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

'adjunto el archivo
Set myAttachments = Correo.Attachments
 myAttachments.Add "C:\Users\JORGE\Desktop\covid\Loi Suites_covid_flyer.png", _
 olByValue, 1, "Medidas De Higiene Y Seguridad"
 myItem.Display
 
ContenidoMail = "<p>Estimado/a" & nombre & "</p>" _
& " <p>Mi nombre es " & mi_Nombre & " le escribo del equipo comercial de Loi Suites Hoteles, esperando que usted y sus seres queridos se encuentren muy bien en estos momentos. </p>" _
& " <p>Nos encantar�a poder acercarles las propuestas de alojamiento que estamos ofreciendo para cuando requieran realizar viajes ya sea en la ciudad de Buenos Aires, Iguaz� o San Martin de los Andes.</p>" _
& "<p>Nos complace poder  asesorarlos con propuestas que se adapten a sus necesidades, brindando una experiencia �nica.</p>" _
& "<p>Hemos incorporado en nuestros hoteles los protocolos de seguridad e higiene establecidos por el Gobierno Nacional y la OMS para garantizarles una estancia segura.<p>" _
& "<p>Por favor en caso de estar interesado en recibir nuestras propuestas o tengan alg�n requerimiento en particular, me encuentro a disposici�n para poder asesorarlos.<p>" _
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
Sub enviar_correo_hoteles_buenos_aires()

Dim i As Integer
Dim pagina1 As Worksheet
Set pagina1 = ActiveWorkbook.Worksheets("Hoja1")
Dim OutApp As Object
Dim Correo As Object
Dim ContenidoMail As String
Dim pie_de_firma As String
Dim mi_Nombre As String
Dim nombre As String
Dim interno As String
Dim mail As String
Dim myAttachments As Outlook.Attachments

i = 2

nombre = pagina1.Range("A" & i).Value

mi_Nombre = InputBox("Ingresar Nombre y Apellido Del Ejecutivo")
interno = InputBox("Interno Del Ejecutivo")
mail = InputBox("Ingresar Mail Del Ejecutivo")

pie_de_firma = "<h4>" & mi_Nombre & "</h4>" _
& "<p>Ejecutivo de Reservas</p>" _
& "<p>Loi Suites Hoteles</p>" _
& "<p>+54-11-5777-8950 int " & interno & "</p>" _
& "<a  href=" & mail & ">" & mail & "</a>" & " | " & "<a  href=" & "https://www.loisuites.com.ar >" & "www.loisuites.com.ar</a>"


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

'adjunto el archivo
Set myAttachments = Correo.Attachments
 myAttachments.Add "C:\Users\JORGE\Desktop\covid\Loi Suites_covid_flyer.png", _
 olByValue, 1, "Medidas De Higiene Y Seguridad"
 myItem.Display
 
ContenidoMail = "<p>Estimado/a" & nombre & "</p>" _
& " <p>Mi nombre es " & mi_Nombre & " le escribo del equipo comercial de Loi Suites Hoteles, esperando que usted y sus seres queridos se encuentren muy bien en estos momentos. </p>" _
& " <p>Nos encantar�a saber si suelen utilizar los servicios de hoteler�a, para poder acercarles una propuesta comercial de alojamiento en caso que suelan utilizar estos servicios en la ciudad de Buenos Aires, en especial para funcionarios, colaboradores, gerentes, CEO de vuestra empresa que precisen viajar a Bs.As. </p>" _
& "<p>Nos complace poder  asesorarlos con propuestas que se adapten a sus necesidades, brindando una experiencia �nica.</p>" _
& "<p>Contamos con 2 hoteles, de distintas categor�as, en donde hemos implementado las medidas establecidas por el Gobierno Nacional y la OMS. Junto a ello, los protocolos establecidos por la ciudad de Buenos Aires. Con el objetivo de garantizar una estad�a segura.</p>" _
& "<p>En adjunto podr� encontrar una breve presentaci�n de los hoteles y nuestros protocolos.</p>" _
& "<p> Loi suites Esmeralda: ubicado en micro centro (M. T. Alvear 842) cuenta con las comodidades de un departamento (habitaci�n + kitchenet) y los servicios de un hotel. (limpieza diaria, seguridad, wifi y desayuno incluidos): <a  href=" & "https://drive.google.com/drive/folders/1ATePDzJMGb8jChGk06NkU_FCB6YVXymX >Acceso a fotos y contenidos</a> (hacer click aqu� para conocer el hotel)</p>" _
& "<p>Loi Suites Recoleta: ubicado en pleno barrio de la Recoleta, (Vicente Lopez 1955), hotel 5 estrellas, con piscina, gym y sauna). Servicio y atenci�n exclusiva. <a  href=" & "https://drive.google.com/drive/folders/0B5yn0ieZMlx-Nklvb1ZaLUlqS3M > Loi Suites Recoleta</a> (hacer click aqu� para conocer el hotel).</p>" _
& "<p>Si todav�a nuestros ejecutivos comerciales no se ha contactado con ustedes para brindarles nuestras nuevas tarifas e informaci�n sobre cuidado al hu�sped, por favor av�senos que lo contactaremos a la brevedad.</p>" _
& "<p> Quedamos atentos a vuestros comentarios</p>" _
& "<p> Muchas gracias</p>" _
& "<p>Saludos</p>"

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

'Sub Adjuntar()
' Dim myAttachments As Outlook.Attachments
'
' Set myAttachments = Correo.Attachments
' myAttachments.Add "C:\Users\JORGE\Desktop\covid\Loi Suites_covid_flyer.png", _
' olByValue, 1, "Medidas De Higiene Y Seguridad"
' myItem.Display
'
'End Sub

