VERSION 5.00
Begin VB.Form frmEMail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enviar E-MAIL"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   Icon            =   "frmEMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkCopia 
      Caption         =   "Copia remitente"
      Height          =   255
      Left            =   960
      TabIndex        =   20
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   19
      Top             =   5010
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar"
      Height          =   375
      Left            =   2940
      TabIndex        =   18
      Top             =   5010
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4920
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   45
      Width           =   5715
      Begin VB.Frame Frame2 
         Caption         =   "Para"
         ForeColor       =   &H00972E0B&
         Height          =   1035
         Left            =   900
         TabIndex        =   24
         Top             =   180
         Width           =   2055
         Begin VB.OptionButton OptPara 
            Caption         =   "Destinos"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   27
            Top             =   675
            Width           =   1335
         End
         Begin VB.OptionButton OptPara 
            Caption         =   "Clientes"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   26
            Top             =   180
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton OptPara 
            Caption         =   "Proveedores"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   25
            Top             =   428
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "E-mail"
         ForeColor       =   &H00972E0B&
         Height          =   1040
         Left            =   3135
         TabIndex        =   21
         Top             =   180
         Width           =   2175
         Begin VB.OptionButton OptMail 
            Caption         =   "Comercial/Compras"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   23
            Top             =   600
            Width           =   1695
         End
         Begin VB.OptionButton OptMail 
            Caption         =   "Administraci�n"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   22
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   960
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   1320
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   960
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   1740
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   960
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   2160
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   2055
         Index           =   3
         Left            =   960
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "frmEMail.frx":000C
         Top             =   2640
         Width           =   4455
      End
      Begin VB.Label Label1 
         Caption         =   "Para"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   1380
         Width           =   330
      End
      Begin VB.Label Label1 
         Caption         =   "E-Mail"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   10
         Top             =   1800
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Asunto"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   9
         Top             =   2220
         Width           =   555
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   600
         Picture         =   "frmEMail.frx":0012
         Top             =   1380
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Mensaje"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   8
         Top             =   2640
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3675
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   60
      Width           =   5715
      Begin VB.TextBox Text3 
         Height          =   1695
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Text            =   "frmEMail.frx":059C
         Top             =   1800
         Width           =   5355
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   3120
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Otro"
         Height          =   255
         Index           =   2
         Left            =   2460
         TabIndex        =   15
         Top             =   1140
         Width           =   675
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Error"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   14
         Top             =   1140
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sugerencia"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   13
         Top             =   1140
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Mensaje"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   17
         Top             =   1500
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Asunto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Enviar e-Mail Ariadna Software"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   180
         TabIndex        =   12
         Top             =   300
         Width           =   4305
      End
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   360
      Picture         =   "frmEMail.frx":05A2
      Top             =   4950
      Width           =   480
   End
End
Attribute VB_Name = "frmEMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
' +-+-    Autor: DAVID      +-+-
' +-+- Alguns canvis: C�SAR +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-

Option Explicit
Public Opcion As Byte
    '0 - Envio del PDF
    '1 - Envio Mail desde menu soporte
Public DatosEnvio As String
    'Nombre para|email para|Asunto|Mensaje|    y para envio tipo3 el mail de otro persona mail|nombre|
    
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private DatosADevolverBusqueda As String

Dim cad As String
Dim HaDevueltoDatos As Boolean
Dim PrimeraVez As Boolean

Private Sub Enviar(ListaArchivos As Collection)
    Dim imageContentID, success
    Dim mailman As ChilkatMailMan
    Dim Valores As String
    Dim J As Integer
    
    On Error GoTo GotException
    Set mailman = New ChilkatMailMan
    
    'Esta cadena es constante de la lincencia comprada a CHILKAT
    mailman.UnlockComponent "1AriadnaMAIL_BOVuuRWYpC9f"
    mailman.LogMailSentFilename = App.path & "\mailSent.log"
    
' 06/02/2007
' modificacion: he descomentado este grupo de instrucciones
    'Servidor smtp
    Valores = ObtenerValoresEnvioMail  'Empipado: smtphost,smtpuser, pass, diremail
    If Valores = "" Then
        MsgBox "Falta configurar en parametros la opcion de envio mail(servidor, usuario, clave)"
        Exit Sub
    End If
    mailman.Smtphost = RecuperaValor(Valores, 1) ' vParam.SmtpHOST
    mailman.SmtpUsername = RecuperaValor(Valores, 2) 'vParam.SmtpUser
    mailman.SmtpPassword = RecuperaValor(Valores, 3) 'vParam.SmtpPass
    
    mailman.SmtpAuthMethod = "LOGIN"

' he comentado este grupo de instrucciones que estaban a pi�on
'    mailman.SmtpHost = "ariadna.myariadna.com"
'    mailman.SmtpUsername = "manolo"
'    mailman.SmtpPassword = "aritel020763"
    
    ' Create the email, add content, address it, and sent it.
    Dim email As ChilkatEmail
    Set email = New ChilkatEmail
    
    'Si es de SOPORTE
    If Opcion = 1 Then
         'Obtenemos la pagina web de los parametros
        '====David
'        Cad = DevuelveDesdeBD("mailsoporte", "parametros", "fechaini", Format(vParam.fechaini, FormatoFecha), "F")
        '====
        cad = DevuelveDesdeBDNew(cAgro, "sparam", "maiempre", "codempre", 1, "N")
        If cad = "" Then
            MsgBox "Falta configurar en parametros el mail de soporte", vbExclamation
            Exit Sub
        End If

        If cad = "" Then GoTo GotException
        email.AddTo "Soporte Contabilidad", cad
        cad = "Soporte Ariagro. "
        If Option1(0).Value Then cad = cad & Option1(0).Caption
        If Option1(1).Value Then cad = cad & Option1(1).Caption
        If Option1(2).Value Then cad = cad & "Otro: " & Text2.Text
        email.Subject = cad

        'Ahora en text1(3).text generaremos nuestro mensaje
        cad = "Fecha: " & Format(Now, "dd/mm/yyyy") & vbCrLf
        cad = cad & "Hora: " & Format(Now, "hh:mm") & vbCrLf
        cad = cad & "Usuario: " & vUsu.Nombre & vbCrLf
        cad = cad & "Nivel USU: " & vUsu.Nivel & vbCrLf
        cad = cad & "Empresa: " & vEmpresa.nomempre & vbCrLf
        cad = cad & "&nbsp;<hr>"
        cad = cad & Text3.Text & vbCrLf & vbCrLf
        Text1(3).Text = cad
    Else
        'Envio de mensajes normal
        email.AddTo Text1(0).Text, Text1(1).Text
        email.Subject = Text1(2).Text
        
        '### A�ade: Laura 11/10/05
        '### Modifica david.     Lo que hare sera para c
        If Opcion < 4 Then
            cad = RecuperaValor(Valores, 4)
            If chkCopia.Value = 1 Then email.AddBcc RecuperaValor(Valores, 2), cad    'vParam.SmtpPass
'            email.AddBcc RecuperaValor(Valores, 2), cad    'vParam.SmtpPass
        Else
            'Para el multienvio de facturacion y renovacion
            cad = RecuperaValor(DatosEnvio, 3)
            If cad = "1" Then
                cad = RecuperaValor(Valores, 4)
                If chkCopia.Value = 1 Then email.AddBcc RecuperaValor(Valores, 2), cad    'vParam.SmtpPass
            End If
        End If
        'Si la opcion es 3   Envio del mail con tooodos los datos en datosenvio
        If Opcion = 3 Then
            CadenaDesdeOtroForm = RecuperaValor(DatosEnvio, 5)
            If CadenaDesdeOtroForm <> "" Then
                If CadenaDesdeOtroForm <> cad Then
                    'El usuario con el que envia el mail NO es el usuario que le indico con el datosenvio
                    'Por lo cual lo a�ado
                    cad = RecuperaValor(DatosEnvio, 6)
                    If chkCopia.Value = 1 Then email.AddBcc "Aviso tomado", CadenaDesdeOtroForm
                End If
            End If
        End If
        
        
        
        
'        If chkCopia.Value = 1 Then email.AddBcc "AriAgro: " & vEmpresa.nomEmpre, RecuperaValor(Valores, 4)
    End If
    
    'El resto lo hacemos comun
    
    
    
' 06/02/2007 todo esto lo he comentado y he a�adido la parte de abajo
'    'La imagen
'    imageContentID = email.AddRelatedContent(App.path & "\logo.jpg")
'
'
'    cad = "<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//DTD HTML 4.0 Transitional//EN" & Chr(34) & ">"
'    cad = cad & "<HTML><HEAD><TITLE>Mensaje</TITLE></HEAD>"
'    cad = cad & "<TABLE BORDER=""0"" CELLSPACING=1 CELLPADDING=0 WIDTH=576>"
'    'Cuerpo del mensaje
'    cad = cad & "<TR><TD VALIGN=""TOP""><P>"
'    FijarTextoMensaje
'    cad = cad & "</P></TD></TR>"
'    cad = cad & "<TR><TD VALIGN=""TOP""><P><HR ALIGN=""LEFT"" SIZE=1></P>"
'    'La imagen
'    cad = cad & "<P ALIGN=""CENTER""><IMG SRC=" & Chr(34) & "cid:" & imageContentID & Chr(34) & "></P>"
'    cad = cad & "<P ALIGN=""CENTER""><FONT SIZE=2>Mensaje creado desde el programa AriGasol de "
'    cad = cad & "<A HREF=""http://www.ariadnasoftware.com/"">Ariadna&nbsp;"
'    cad = cad & "Software S.L.</A></P><P ALIGN=""CENTER""></P>"
'    cad = cad & "<P>Este correo electr�nico y sus documentos adjuntos est�n dirigidos EXCLUSIVAMENTE a "
'    cad = cad & " los destinatarios especificados. La informaci�n contenida puede ser CONFIDENCIAL"
'    cad = cad & " y/o estar LEGALMENTE PROTEGIDA.</P>"
'    cad = cad & "<P>Si usted recibe este mensaje por ERROR, por favor comun�queselo inmediatamente al"
'    cad = cad & " remitente y ELIM�NELO ya que usted NO ESTA AUTORIZADO al uso, revelaci�n, distribuci�n, "
'    cad = cad & " impresi�n o copia de toda o alguna parte de la informaci�n contenida, gracias "
'    cad = cad & ".</FONT></P><P><HR ALIGN=""LEFT"" SIZE=1></TD>"
'    cad = cad & "</TR></TABLE></BODY></HTML>"
    
    
    cad = "<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//DTD HTML 4.0 Transitional//EN" & Chr(34) & ">"
    cad = cad & "<HTML><HEAD><TITLE>Mensaje</TITLE></HEAD>"
    cad = cad & "<TABLE BORDER=""0"" CELLSPACING=1 CELLPADDING=0 WIDTH=576>"
    'Cuerpo del mensaje
    cad = cad & "<TR><TD VALIGN=""TOP""><P>"
    FijarTextoMensaje
    cad = cad & "</P></TD></TR>"
    cad = cad & "<TR><TD VALIGN=""TOP""><P><hr></P>"
    'La imagen
    'cad = cad & "<P ALIGN=""CENTER""><IMG SRC=" & Chr(34) & "cid:" & imageContentID & Chr(34) & "></P>"
'--monica: no tiene que salir
'    cad = cad & "<P ALIGN=""CENTER""><FONT SIZE=2>Mensaje creado desde el programa " & App.EXEName & " de "
'    cad = cad & "<A HREF=""http://www.ariadnasoftware.com/"">Ariadna&nbsp;"
'    cad = cad & "Software S.L.</A></P><P ALIGN=""CENTER""></P>"
    cad = cad & "<P>Este correo electr�nico y sus documentos adjuntos estan dirigidos EXCLUSIVAMENTE a "
    cad = cad & " los destinatarios especificados. La informaci�n contenida puesde ser CONFIDENCIAL"
    cad = cad & " y/o estar LEGALMENTE PROTEGIDA.</P>"
    cad = cad & "<P>Si usted recibe este mensaje por ERROR, por favor comun�queselo inmediatamente al"
    cad = cad & " remitente y ELIMINELO ya que usted NO ESTA AUTORIZADO al uso, revelaci�n, distribuci�n"
    cad = cad & " impresi�n o copia de toda o alguna parte de la informaci�n contenida, Gracias "
    cad = cad & ".</FONT></P><P><HR ALIGN=""LEFT"" SIZE=1></TD>"
    cad = cad & "</TR></TABLE></BODY></HTML>"
     
    
    
    email.SetHtmlBody (cad)
    
    
    
    email.AddPlainTextAlternativeBody "Programa e-mail NO soporta HTML. " & vbCrLf & Text1(3).Text
    email.From = RecuperaValor(Valores, 4) 'vParam.diremail
    'email.From = "manolo@myariadna.com"
    
'    If Opcion = 0 Then
'        'ADjunatmos el PDF
'        email.AddFileAttachment App.path & "\docum.pdf"
'    End If
    If Opcion <> 1 Then   'Solo la opcion 1 NO lleva attachment
        'ADjunatmos el PDF
        If ListaArchivos Is Nothing Then
            email.AddFileAttachment App.path & "\docum.pdf"
        Else
            
            For J = 1 To ListaArchivos.Count
                email.AddFileAttachment ListaArchivos.item(J)
            Next J
        End If
    End If
        
    
    'email.SendEncrypted = 1
    success = mailman.SendEmail(email)
    If (success = 1) Then
        If Opcion <> 2 And Opcion <> 4 Then
            cad = "Mensaje enviado correctamente."
            MsgBox cad, vbInformation
            Command2(0).SetFocus
        End If
    Else
        cad = "Han ocurrido errores durante el envio.Compruebe el archivo log.xml para mas informacion"
        mailman.SaveXmlLog App.path & "\log.xml"
        MsgBox cad, vbExclamation
    End If
    
'    success = mailman.SendEmail(email)
'    If (success = 1) Then
'        cad = "Mensaje enviado correctamente."
'        MsgBox cad, vbInformation
'        Command2(0).SetFocus
'    Else
'        cad = "Han ocurrido errores durante el envio.Compruebe el archivo log.xml para mas informacion"
'        mailman.SaveXmlLog App.path & "\log.xml"
'        MsgBox cad, vbExclamation
'    End If
    
    
GotException:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set email = Nothing
    Set mailman = Nothing

End Sub

Private Sub Command1_Click()
Dim Col As Collection

    If Not DatosOk Then Exit Sub
    Screen.MousePointer = vbHourglass
    Image2.visible = True
    Me.Refresh
    
'[Monica]11/01/2012: envio de por outlook
'                    cambio Enviar Nothing por lo siguiente
'    Enviar Nothing
    'Opcion cero. Confirmacion entrega pedido
    If Opcion = 0 Then
        cad = RecuperaValor(Me.DatosEnvio, 5)
        If cad <> "" Then
            Set Col = New Collection
            Col.Add cad
        End If
    
    End If
                  
    EnvioNuevo Col

    Image2.visible = False
    Screen.MousePointer = vbDefault
    
    Unload Me
End Sub

'Modificacion MUY IMPORTANTE
'Programaita de envio: arigesmail.exe
'Si la opcion es esa hara OOOtras cosas, si no lo dejamos como esta
Private Sub EnvioNuevo(ListaArch As Collection)

    If vParamAplic.ExeEnvioMail <> "" Then
        'Utliza el programa que lanza desde el outlook
        EnvioDesdeExeNuestro ListaArch
        If Opcion = 0 And DatosEnvio <> "" Then Me.DatosEnvio = "OK"
    Else
    
    
        'El que habia
        Enviar2 ListaArch
    End If

End Sub

Private Sub EnvioDesdeExeNuestro(ListaArchivos As Collection)
Dim Lanza As String
Dim J As Integer

    If Not DatosOk Then Exit Sub
        
    'Dire email
    Lanza = Text1(1).Text & "|"
    'Asunto
    Lanza = Lanza & Text1(2).Text & "|"
    
    'Aqui pondremos lo del texto del BODY
    Lanza = Lanza & Text1(3).Text & "|"
    
    
    'Envio o mostrar
    Lanza = Lanza & "1"   '0. Display        1.send
    
    'Campos reservados para el futuro
    Lanza = Lanza & "||||"
    
    'El/los adjuntos
    If Opcion <> 1 Then   'Solo la opcion 1 NO lleva attachment
        'ADjunatmos el PDF
        If ListaArchivos Is Nothing Then
            Lanza = Lanza & App.path & "\docum.pdf" & "|"
        Else
            
            For J = 1 To ListaArchivos.Count
                   Lanza = Lanza & ListaArchivos.item(J) & "|"
            Next J
        End If
    End If
    
    Lanza = App.path & "\" & vParamAplic.ExeEnvioMail & " " & Lanza
    Shell Lanza, vbNormalFocus

End Sub


'Modificacion: 10 Abril 2007
' Enviar siempre envia el documento llamado docum.pdf
' Ahora necesito enviar varios documentos por mail
' Para ello mandare si en la lista hay algo
' seran los path de los archivos, si no sera docum.pdf
Private Sub Enviar2(ListaArchivos As Collection)
    Dim success
    Dim mailman As ChilkatMailMan
    Dim Valores As String
    Dim J As Integer
    
    On Error GoTo GotException
    Set mailman = New ChilkatMailMan
    
    'Esta cadena es constante de la lincencia comprada a CHILKAT
    mailman.UnlockComponent "1AriadnaMAIL_BOVuuRWYpC9f"
'    mailman.LogMailSentFilename = App.Path & "\mailSent.log"
    
    
    'Servidor smtp
    If vParamAplic.EnvioDesdeOutlook Then
        Valores = "||||"
    Else
        Valores = ObtenerValoresEnvioMail  'Empipado: smtphost,smtpuser, pass, diremail
        If Valores = "" Then
            MsgBox "Falta configurar en paremtros la opcion de envio mail(servidor, usuario, clave)"
            Exit Sub
        End If
        mailman.Smtphost = RecuperaValor(Valores, 1) ' vParam.SmtpHOST
        mailman.SmtpUsername = RecuperaValor(Valores, 2) 'vParam.SmtpUser
        mailman.SmtpPassword = RecuperaValor(Valores, 3) 'vParam.SmtpPass
        
        'David 2 Mayo 2007
        mailman.SmtpAuthMethod = "LOGIN"
        
    End If
    
    ' Create the email, add content, address it, and sent it.
    Dim email As ChilkatEmail
    Set email = New ChilkatEmail
    
    'Si es de SOPORTE
    
    
    If Opcion = 1 Then
         'Obtenemos la pagina web de los parametros
        '====David
'        Cad = DevuelveDesdeBD("mailsoporte", "parametros", "fechaini", Format(vParam.fechaini, FormatoFecha), "F")
        '====
        cad = DevuelveDesdeBDNew(cAgro, "sparam", "maiempre", "codempre", 1, "N")
        If cad = "" Then
            MsgBox "Falta configurar en parametros el mail de soporte", vbExclamation
            Exit Sub
        End If
    
        If cad = "" Then GoTo GotException
        email.AddTo "Soporte Gesti�n", cad
        cad = "Soporte Arigasol. "
        If Option1(0).Value Then cad = cad & Option1(0).Caption
        If Option1(1).Value Then cad = cad & Option1(1).Caption
        If Option1(2).Value Then cad = cad & "Otro: " & Text2.Text
        email.Subject = cad
        
        'Ahora en text1(3).text generaremos nuestro mensaje
        cad = "Fecha: " & Format(Now, "dd/mm/yyyy") & vbCrLf
        cad = cad & "Hora: " & Format(Now, "hh:mm") & vbCrLf
        cad = cad & "Usuario: " & vSesion.Nombre & vbCrLf
        cad = cad & "Nivel USU: " & vSesion.Nivel & vbCrLf
        cad = cad & "Empresa: " & vEmpresa.nomempre & vbCrLf
        cad = cad & "&nbsp;<hr>"
        cad = cad & Text3.Text & vbCrLf & vbCrLf
        Text1(3).Text = cad
    Else
        'Opcion=0 or opcion= 3 or envio=4
        'Envio de mensajes normal
        ' ---- [04/11/2009] [LAURA] : concatenar al final del asunto [ARI] para poder crear regla correo
        
        
        If Opcion <> 6 Then
            email.Subject = Text1(2).Text & " [ARI]"
        Else
            email.Subject = Text1(2).Text
        End If
        ' ----
        email.AddTo Text1(0).Text, Text1(1).Text
        
        '### A�ade: Laura 11/10/05
        '### Modifica david.     Lo que hare sera para c
        If Opcion < 4 Then
            cad = RecuperaValor(Valores, 4)
            email.AddBcc RecuperaValor(Valores, 2), cad    'vParam.SmtpPass
            
        Else
            'Para el multienvio de facturacion y renovacion
            cad = RecuperaValor(DatosEnvio, 3)
            If cad = "1" Then
                cad = RecuperaValor(Valores, 4)
                email.AddBcc RecuperaValor(Valores, 2), cad    'vParam.SmtpPass
            End If
        End If
        'Si la opcion es 3   Envio del mail con tooodos los datos en datosenvio
        If Opcion = 3 Then
            CadenaDesdeOtroForm = RecuperaValor(DatosEnvio, 5)
            If CadenaDesdeOtroForm <> "" Then
                If CadenaDesdeOtroForm <> cad Then
                    'El usuario con el que envia el mail NO es el usuario que le indico con el datosenvio
                    'Por lo cual lo a�ado
                    cad = RecuperaValor(DatosEnvio, 6)
                    email.AddBcc "Aviso tomado", CadenaDesdeOtroForm
                End If
            End If
        End If
    End If
    
    'El resto lo hacemos comun
    'La imagen
    'imageContentID = email.AddRelatedContent(App.Path & "\minilogo.bmp")
    
    
    cad = "<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//DTD HTML 4.0 Transitional//EN" & Chr(34) & ">"
    cad = cad & "<HTML><HEAD><TITLE>Mensaje</TITLE></HEAD>"
    cad = cad & "<TABLE BORDER=""0"" CELLSPACING=1 CELLPADDING=0 WIDTH=576>"
    'Cuerpo del mensaje
    cad = cad & "<TR><TD VALIGN=""TOP""><P>"
    FijarTextoMensaje
    cad = cad & "</P></TD></TR>"
    cad = cad & "<TR><TD VALIGN=""TOP""><P><hr></P>"
    cad = cad & "<FONT SIZE=2>"
    cad = cad & "<P><P><P><P align=""justify"">Este correo electr�nico y sus documentos adjuntos estan dirigidos EXCLUSIVAMENTE a "
    cad = cad & " los destinatarios especificados. La informaci�n contenida puesde ser CONFIDENCIAL"
    cad = cad & " y/o estar LEGALMENTE PROTEGIDA.</P>"
    cad = cad & "<P align=""justify"">Si usted recibe este mensaje por ERROR, por favor comun�queselo inmediatamente al"
    
    cad = cad & " remitente y ELIMINELO ya que usted NO ESTA AUTORIZADO al uso, revelaci�n, distribuci�n"
    cad = cad & " impresi�n o copia de toda o alguna parte de la informaci�n contenida, Gracias "
    cad = cad & ".</FONT></P><P><HR ALIGN=""LEFT"" SIZE=1></TD>"
    cad = cad & "</TR></TABLE></BODY></HTML>"
    
    email.SetHtmlBody (cad)
    
    'Texto alternativo
    cad = ""
    cad = cad & "Este correo electronico y sus documentos adjuntos estan dirigidos EXCLUSIVAMENTE a " & vbCrLf
    cad = cad & " los destinatarios especificados. La informacion contenida puesde ser CONFIDENCIAL" & vbCrLf
    cad = cad & " y/o estar LEGALMENTE PROTEGIDA." & vbCrLf & vbCrLf
    cad = cad & "Si usted recibe este mensaje por ERROR, por favor comuniqueselo inmediatamente al" & vbCrLf
    cad = cad & " remitente y ELIMINELO ya que usted NO ESTA AUTORIZADO al uso, revelacion, distribucion" & vbCrLf
    cad = cad & " impresion o copia de toda o alguna parte de la informacion contenida, Gracias " & vbCrLf

    
    'Por si no acepta HTML
    cad = UCase(cad)
    email.AddPlainTextAlternativeBody Text1(3).Text & vbCrLf & vbCrLf & vbCrLf & cad
    email.From = RecuperaValor(Valores, 4) 'vParam.diremail
    
    
    If Opcion <> 1 Then   'Solo la opcion 1 NO lleva attachment
        'ADjunatmos el PDF
        If ListaArchivos Is Nothing Then
            email.AddFileAttachment App.path & "\docum.pdf"
        Else
            
            For J = 1 To ListaArchivos.Count
                   email.AddFileAttachment ListaArchivos.item(J)
            Next J
        End If
    End If
        
    
    'email.SendEncrypted = 1
    
        'sI ENVIA POR OUTLOOK O NO
     If vParamAplic.EnvioDesdeOutlook Then
        'Si envia por outlook
         mailman.SendViaOutlook email
         success = 1
        
    Else
        success = mailman.SendEmail(email)
    End If
    If (success = 1) Then
        If Opcion <> 2 And Opcion <> 4 And Opcion <> 6 Then
            If vParamAplic.EnvioDesdeOutlook Then
                cad = "Enviado al outlook"
            Else
                cad = "Mensaje enviado correctamente."
            End If
            MsgBox cad, vbInformation
            Command2(0).SetFocus
        End If
        
        ' ---- [04/11/2009] [LAURA] : para saber q se ha enviado con exito y actualizar check de enviado
        If Opcion = 0 And DatosEnvio <> "" Then
            Me.DatosEnvio = "OK"
            Command2_Click (0)
        End If
        ' ---
    Else
        cad = "Han ocurrido errores durante el envio.Compruebe el archivo log.xml para mas informacion"
        mailman.SaveXmlLog App.path & "\log.xml"
        MsgBox cad, vbExclamation
    End If
    
    
GotException:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set email = Nothing
    Set mailman = Nothing

End Sub


Private Sub Command2_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Activate()
     If PrimeraVez Then
        PrimeraVez = False
        If Opcion = 2 Or Opcion >= 4 Then
            If Opcion = 2 Then
                HacerMultiEnvio
            Else
                'Opcion 4 y 5
                Me.Command1.visible = False
                Command2(0).visible = False
                DoEvents
                HacerMultiEnvioFacturacion
                
                
                
                Me.Command1.visible = True
                Command2(0).visible = True
                DoEvents
            End If
            Unload Me
        ' ---- [04/11/2009] [LAURA] : A�adir bot�n para enviar informe confirmacion entrega del Pedido
        ' ----                        para ello aqui a�ado opcion=0
        Else
            If Opcion = 3 Or (Opcion = 0) Then
                If DatosEnvio <> "" Then
                    'Fuerzo el envio de mail
        
                    Text1(0).Text = RecuperaValor(DatosEnvio, 1)
                    Text1(1).Text = RecuperaValor(DatosEnvio, 2)
                    Text1(2).Text = RecuperaValor(DatosEnvio, 3)
                    Text1(3).Text = RecuperaValor(DatosEnvio, 4)
                    Me.Refresh
                    DoEvents
                    
                    If Opcion = 3 Then
                        Command1_Click
                        Unload Me
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    Image2.visible = False
    limpiar Me
    Frame1(0).visible = (Opcion = 0) Or (Opcion = 2)
    Frame1(1).visible = (Opcion = 1)
    If Opcion = 1 Then HabilitarText

    '###Descomentar
'    cad = DevuelveDesdeBD("smtpHost", "spara1", "fechaini", Format(vParam.fechaini, FormatoFecha), "F")
'    cad = DevuelveDesdeBDnew(conAri, "spara1", "smtphost", "codigo", "1", "N")
'    Me.Command1.Enabled = (cad <> "")

    cad = ""
    If vParamAplic.ExeEnvioMail <> "" Then
        cad = "OK"
    Else
        If vParamAplic.EnvioDesdeOutlook Then
            cad = "OK"
        Else
            cad = DevuelveDesdeBDNew(cAgro, "sparam", "smtphost", "codparam", "1", "N")
        End If
    End If
    Me.Command1.Enabled = (cad <> "")
    
    
    Me.Command1.visible = Opcion <> 6

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Opcion = 0
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Text1(0).Text = RecuperaValor(CadenaDevuelta, 2)
    Text1(1).Text = RecuperaValor(CadenaDevuelta, 4)
End Sub

Private Sub Image1_Click()
    'busqueda de clientes que tiene e-mail
Dim cadSel As String
Dim cadtabla As String
Dim cadCampo As String

    'seleccionar de que tabla vamos a leer los datos
    If Me.OptPara(0).Value Then
        'leer datos de clientes
        cadtabla = "clientes"
        
        'seleccionar a que e-mail vamos a enviar
        If Me.OptMail(0).Value Then 'mail de Administracion
            'seleccionar el mail1
            cadCampo = "maiclie1"
        Else 'mail comercial
            cadCampo = "maiclie2 "
        End If
        
    ElseIf Me.OptPara(1).Value Then
        'datos de proveedores
        cadtabla = "proveedor"
        
        'seleccionar a que e-mail vamos a enviar
        'seleccionar solo los proveedores que tiene valor en mail1 o mail2.
        If Me.OptMail(0).Value Then
            cadCampo = "maiprov1"
        Else
            cadCampo = "maiprov2"
        End If
    Else
        'destinos
        cadtabla = "destinos inner join clientes on (destinos.codclien = clientes.codclien)"
        cadCampo = "maidesti"
    End If

    cadSel = " (not isnull(" & cadCampo & ") and " & cadCampo & "<>'') "
    MandaBusquedaPrevia cadSel, cadtabla, cadCampo
    

'    Set frmC = New frmFacClientes
'    frmC.DatosADevolverBusqueda = "0|1"
''    frmC.ConfigurarBalances = 5  'NUEVO opcion
'    frmC.Show vbModal
'    Set frmC = Nothing
'    If Text1(0).Text <> "" Then PonerFoco Text1(2)
End Sub


'Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
'
'    Screen.MousePointer = vbHourglass
'    Text1(0).Tag = RecuperaValor(CadenaSeleccion, 1)
'    Text1(0).Text = RecuperaValor(CadenaSeleccion, 3)
'    'Si regresa con datos tengo k devolveer desde la bd el campo e-mail
'    Text1(1).Text = RecuperaValor(CadenaSeleccion, 4)
''    cad = DevuelveDesdeBDNew(conAri, "sclien", "maiclie1", "codclien", Text1(0).Tag, "T")
''    Text1(1).Text = cad
'    Screen.MousePointer = vbDefault
'End Sub
'
'Private Sub Image1_Click()
'    Set frmC = New frmFacClientes
'    frmC.DatosADevolverBusqueda = "0|1"
''    frmC.ConfigurarBalances = 5  'NUEVO opcion
'    frmC.Show vbModal
'    Set frmC = Nothing
'    If Text1(0).Text <> "" Then PonerFoco Text1(2)
'End Sub

Private Sub Option1_Click(Index As Integer)
    HabilitarText
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
'    With Text1(Index)
'        .SelStart = 0
'        .SelLength = Len(.Text)
'    End With
End Sub

Private Function DatosOk() As Boolean
Dim i As Integer

    DatosOk = False
    'If Opcion = 0 Or Opcion = 3 Then
    If Opcion <> 1 And Opcion <> 2 Then
                'Pocas cosas a comprobar
                For i = 0 To 2
                    Text1(i).Text = Trim(Text1(i).Text)
                    If Text1(i).Text = "" Then
                        MsgBox "El campo: " & Label1(i).Caption & " no puede estar vacio.", vbExclamation
                        Exit Function
                    End If
                Next i
                
                'EL del mail tiene k tener la arroba @
                i = InStr(1, Text1(1).Text, "@")
                If i = 0 Then
                    MsgBox "Direccion e-mail erronea", vbExclamation
                    Exit Function
                End If
    Else
        Text2.Text = Trim(Text2.Text)
        'SOPORTE
        If Trim(Text3.Text) = "" Then
            MsgBox "El mensaje no puede ir en blanco", vbExclamation
            Exit Function
        End If
        If Option1(2).Value Then
            If Text2.Text = "" Then
                MsgBox "El campo 'OTRO asunto' no puede ir en blanco", vbExclamation
                Exit Function
            End If
        End If
    End If
      
    'Llegados aqui OK
    DatosOk = True
        
End Function


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 3 Then Exit Sub 'Si estamos en el de datos nos salimos
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

'El procedimiento servira para ir buscando los vbcrlf y cambiarlos por </p><p>
Private Sub FijarTextoMensaje()
Dim i As Integer
Dim J As Integer

    J = 1
    Do
        i = InStr(J, Text1(3).Text, vbCrLf)
        If i > 0 Then
              cad = cad & Mid(Text1(3).Text, J, i - J) & "</P><P>"
        Else
            cad = cad & Mid(Text1(3).Text, J)
        End If
        J = i + 2
    Loop Until i = 0
End Sub

Private Sub HabilitarText()
    If Option1(2).Value Then
        Text2.Enabled = True
        Text2.BackColor = vbWhite
    Else
        Text2.Enabled = False
        Text2.BackColor = &H80000018
    End If
End Sub

Private Function RecuperarDatosEMAILAriadna() As Boolean
Dim NF As Integer

    RecuperarDatosEMAILAriadna = False
    NF = FreeFile
    Open App.path & "\soporte.dat" For Input As #NF
    Line Input #NF, cad
    Close #NF
    If cad <> "" Then RecuperarDatosEMAILAriadna = True
    
End Function

Private Function ObtenerValoresEnvioMail() As String
Dim miRsAux As ADODB.Recordset

    ObtenerValoresEnvioMail = ""
    Set miRsAux = New ADODB.Recordset
    cad = "Select diremail,SmtpHost, SmtpUser, SmtpPass  from sparam where"
    '####Descomentar
'    Cad = Cad & " fechaini='" & Format(vParam.fechaini, FormatoFecha) & "';"
    cad = cad & " codparam=1;"
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        cad = DBLet(miRsAux!Smtphost)
        cad = cad & "|" & DBLet(miRsAux!SmtpUser)
        cad = cad & "|" & DBLet(miRsAux!Smtppass)
        cad = cad & "|" & DBLet(miRsAux!DireMail) & "|"
        ObtenerValoresEnvioMail = cad
    End If
    miRsAux.Close
    Set miRsAux = Nothing
End Function


Private Sub MandaBusquedaPrevia(CadB As String, NomTabla As String, nomcampo As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim Tabla As String
Dim Titulo As String
Dim Conexion As Byte

    'Llamamos a al form
    '##A mano
    cad = ""
    Select Case NomTabla
        Case "clientes"
            cad = cad & "C�digo|clientes.codclien|N|000000|9�"
            cad = cad & "Nombre|clientes.nomclien|T||29�"
            cad = cad & "Domicilio|clientes.domclien|T||29�"
            cad = cad & "E-mail|clientes." & nomcampo & "|T||33�"
'            Tabla = NomTabla
            Titulo = "Clientes"
        Case "proveedor"
            cad = cad & "C�digo|proveedor.codprove|N|000000|9�"
            cad = cad & "Nombre|proveedor.nomprove|T||29�"
            cad = cad & "Nom.Comer.|proveedor.nomcomer|T||29�"
            cad = cad & "E-mail|proveedor." & nomcampo & "|T||33�"
'            Tabla = NomTabla
            Titulo = "Proveedores"

        Case "destinos inner join clientes on (destinos.codclien = clientes.codclien)"
            cad = cad & "C�digo|destinos.coddesti|N|0000|9�"
            cad = cad & "Nombre|destinos.nomdesti|T||44�"
            cad = cad & "Cliente|clientes.nomclien|T||30�"
            cad = cad & "E-mail|destinos." & nomcampo & "|T||44�"
            Tabla = NomTabla & " inner join clientes on destinos.codclien = clientes.codclien "
            Titulo = "Destinos"
    End Select
    Tabla = NomTabla
    Conexion = cAgro    'Conexi�n a BD: Ariagro
    
'    Select Case Val(Me.imgBuscar(0).Tag)
'        Case 5  'Cuenta Contable
'            'Se llama a Busqueda desde el campo Cuenta contable
'            '#A MANO: Porque busca en la tabla cuentas
'            'de la base de datos de Contabilidad
'            cad = cad & "C�digo|cuentas|codmacta|T||30�Denominacion|cuentas|nommacta|T||70�"
'            Tabla = "cuentas"
'            Titulo = "Cuentas Contables"
'            Conexion = conConta    'Conexi�n a BD: Conta
'        Case Else   'Registro de la tabla de cabeceras: sartic
'            cad = cad & ParaGrid(Text1(0), 10, "C�digo")
'            cad = cad & ParaGrid(Text1(1), 50, "Nombre")
'            cad = cad & ParaGrid(Text1(2), 40, "Nombre Comercial")
'            Tabla = "sclien"
'            Titulo = "Clientes"
'            Conexion = conAri    'Conexi�n a BD: Ariges
'    End Select
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vtabla = Tabla
        frmB.vSQL = CadB
'        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|1|2|3|"
        frmB.vTitulo = Titulo
        frmB.vSelElem = 1
'        frmB.vConexionGrid = Conexion
'        frmB.vCargaFrame = (Conexion = 2)
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
            PonerFoco Text1(2)
'        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub HacerMultiEnvio()
Dim cad As String
Dim RS As ADODB.Recordset
Dim i As Integer, Cont As Integer

On Error GoTo EMulti

        'Campos comunes
    'ENVIO MASIVO DE EMAILS
    Text1(2).Text = RecuperaValor(Me.DatosEnvio, 1)
    Text1(3).Text = RecuperaValor(Me.DatosEnvio, 2)
    
    Me.Refresh
    
    cad = "SELECT * from tmpMail WHERE codusu=" & vUsu.Codigo
    Set RS = New ADODB.Recordset
    RS.Open cad, conn, adOpenKeyset, adLockOptimistic, adCmdText

    Cont = 0
    While Not RS.EOF
        Cont = Cont + 1
        RS.MoveNext
    Wend
    RS.MoveFirst
    
    
    i = 1
    Me.Refresh
    While Not RS.EOF
        Screen.MousePointer = vbHourglass
        Text1(0).Text = RS!nomprove
        Text1(1).Text = RS!email
        Caption = "Enviar E-MAIL (" & i & " de " & Cont & ")"
        Me.Refresh
        
        'De momento volvemos a copiar el archivo como docum.pdf
        FileCopy App.path & "\temp\" & RS!codProve & ".pdf", App.path & "\docum.pdf"
        Me.Refresh
        NumRegElim = 0
        Enviar2 Nothing
        
        
'        If NumRegElim = 1 Then
'            'NO SE HA ENVIADO.
'            cad = "UPDATE tmp347 SET IMporte=0 WHERE codusu =" & vUsu.Codigo & " AND cliprov =0 AND cta='" & RS!cta & "'"
'            Conn.Execute cad
'        End If
        'Siguiente
        RS.MoveNext
        i = i + 1
    Wend
    RS.Close
    
EMulti:
    
End Sub


'MULTIE ENVIO FACTURACION
Private Sub HacerMultiEnvioFacturacion()
Dim cad As String
Dim RS As ADODB.Recordset
Dim i As Integer, Cont As Integer
Dim Lis As Collection
Dim ListaArchivos As Collection
Dim FormatoHtml As Boolean
Dim T1 As Single

On Error GoTo EMulti2

        'Campos comunes
    'ENVIO MASIVO DE EMAILS
    Text1(2).Text = RecuperaValor(Me.DatosEnvio, 1)
    
    
    Me.Refresh
    DoEvents
    cad = RecuperaValor(DatosEnvio, 4)
    'AGrupamos en el envio de facturas
    If Opcion = 4 Then cad = cad & " GROUP by codigo1"
    Set RS = New ADODB.Recordset
    RS.Open cad, conn, adOpenKeyset, adLockOptimistic, adCmdText

    Set Lis = New Collection
    While Not RS.EOF
        Lis.Add CStr(RS!Codigo1)
        RS.MoveNext
    Wend
    RS.Close
    
    FormatoHtml = False
    If vParamAplic.ExeEnvioMail <> "" Then
        FormatoHtml = True
    Else
        If Not vParamAplic.EnvioDesdeOutlook Then FormatoHtml = True
    End If
    
    T1 = Timer
    For i = 1 To Lis.Count
         Caption = "Enviar E-MAIL (" & i & " de " & Lis.Count & ")"
        DoEvents
        cad = RecuperaValor(DatosEnvio, 4)
        cad = cad & " and codigo1 =" & Lis.item(i)
        RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Screen.MousePointer = vbHourglass
        Text1(0).Text = RS!nomclien
        Text1(1).Text = RS!email
        'Los meteremos en una tabla
        If FormatoHtml Then
            cad = "<BR><BR><TABLE BORDER=""1"" CELLSPACING=1 CELLPADDING=0 WIDTH=576>"
            'Cuerpo del mensaje
            If Opcion = 4 Then
                cad = cad & "<TR><TD width=""274"" bgcolor=""#CCCCCC""><B>Factura</B></TD><TD width=""145"" bgcolor=""#CCCCCC""><B>Fecha</B></TD><TD width=""145"" bgcolor=""#CCCCCC""><B>Importe</B></td></TR>"
            Else
                cad = cad & "<TR><TD width=""640"" bgcolor=""#CCCCCC""><B>Documento</B></TD></TR>"
            End If
        Else
            If Opcion = 4 Then
                cad = " Factura             Fecha             Importe "
            Else
                cad = cad & "Documento "
            End If
            cad = vbCrLf & vbCrLf & vbCrLf & cad & vbCrLf & vbCrLf & String(40, "-") & vbCrLf & vbCrLf
        End If
        Text1(3).Text = RecuperaValor(Me.DatosEnvio, 2) & cad
        Set ListaArchivos = New Collection
        While Not RS.EOF
           
            Me.Refresh
            '
            'De momento volvemos a copiar el archivo como docum.pdf
            If Opcion = 4 Then
                'cad = App.Path & "\temp\" & RS!NumAlbar & Format(RS!codProve, "0000000") & Format(RS!codArtic, "0000000") & Format(RS!FechaAlb, "yymmdd") & ".pdf"
                cad = App.path & "\temp\" & RS!nombre1 & Format(RS!importe1, "0000000") & ".pdf" 'RS!importe1 & Format(RS!Codigo1, "0000000") & ".pdf"
            Else
                'Opcion5: Carta renovacion
                cad = App.path & "\temp\" & Format(RS!codProve, "0000000") & ".pdf"
            End If
            If Dir(cad, vbArchive) = "" Then
                'ERROR. El fichero ha sido eliminado
                MsgBox "No existe el fichero: " & cad & vbCrLf & "El proceso finalizara", vbExclamation
                RS.Close
                Exit Sub
            Else
                ListaArchivos.Add cad
                'En el asunto pondremos los archivos que enviamos
                If Opcion = 4 Then
                    If FormatoHtml Then
                        cad = "</div></TD><TD><div align=""right"">" & Format(RS!importe2, FormatoImporte) & "</div></TD></TR>"
                    Else
                        cad = Space(20) & Format(RS!importe2, FormatoImporte)
                    End If
                    
                    If FormatoHtml Then
                        cad = "</TD><TD><div align=""center"">" & Format(RS!fecha1, "dd/mm/yyyy") & cad
                    Else
                        cad = Space(15) & Format(RS!fecha1, "dd/mm/yyyy") & cad
                    End If
                    
        
                    cad = RS!nombre1 & Format(RS!importe1, "0000000") & cad
                                
                    If FormatoHtml Then
                        cad = "<TR><TD>" & cad
                    Else
                        cad = cad & vbCrLf
                    End If
                
                Else
                    'Opcion:5.  Carta renovacion
                    If FormatoHtml Then cad = "<TR><TD>"
                    cad = cad & "Documento" & Format(RS!codProve, "0000000")
                    If FormatoHtml Then
                        cad = cad & "</TD></TR>"
                    Else
                        cad = cad & vbCrLf
                    End If
                
                End If
                
                Text1(3).Text = Text1(3).Text & "    " & cad
            End If
            
            'Siguiente
            RS.MoveNext
            
        Wend
        RS.Close
        If FormatoHtml Then Text1(3).Text = Text1(3).Text & "</TABLE><BR><BR>"
        
        EnvioNuevo ListaArchivos
        
        Set ListaArchivos = Nothing
        
        T1 = Timer - T1
        If T1 < 3 Then
            T1 = 3 - T1
            espera T1
        End If
        T1 = Timer
        
    Next i
    Set Lis = Nothing
    Exit Sub
EMulti2:
    MuestraError Err.Number
End Sub




