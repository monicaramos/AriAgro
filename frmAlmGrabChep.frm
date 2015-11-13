VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAlmGrabChep 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6840
   Icon            =   "frmAlmGrabChep.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameCobros 
      Height          =   6165
      Left            =   90
      TabIndex        =   10
      Top             =   90
      Width           =   6555
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1755
         MaxLength       =   3
         TabIndex        =   2
         Top             =   2250
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1755
         MaxLength       =   3
         TabIndex        =   3
         Top             =   2625
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "Text5"
         Top             =   2250
         Width           =   3345
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "Text5"
         Top             =   2625
         Width           =   3345
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   3405
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Text5"
         Top             =   4485
         Width           =   2595
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   3405
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text5"
         Top             =   4110
         Width           =   2595
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   5
         Left            =   1740
         MaxLength       =   16
         TabIndex        =   7
         Top             =   4500
         Width           =   1590
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   4
         Left            =   1740
         MaxLength       =   16
         TabIndex        =   6
         Text            =   "1234567890123456"
         Top             =   4110
         Width           =   1590
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3540
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3180
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5070
         TabIndex        =   9
         Top             =   5460
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3990
         TabIndex        =   8
         Top             =   5460
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   0
         Top             =   1335
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1710
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   1335
         Width           =   3360
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   1710
         Width           =   3360
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   495
         TabIndex        =   20
         Top             =   5100
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   360
         Top             =   5370
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   9
         Left            =   885
         TabIndex        =   30
         Top             =   2250
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   10
         Left            =   885
         TabIndex        =   29
         Top             =   2625
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Destino"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   17
         Left            =   480
         TabIndex        =   28
         Top             =   2040
         Width           =   540
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1500
         MouseIcon       =   "frmAlmGrabChep.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar destino"
         Top             =   2250
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   1500
         MouseIcon       =   "frmAlmGrabChep.frx":015E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar destino"
         Top             =   2625
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1440
         MouseIcon       =   "frmAlmGrabChep.frx":02B0
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar artículo"
         Top             =   4485
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1440
         MouseIcon       =   "frmAlmGrabChep.frx":0402
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar artículo"
         Top             =   4110
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Artículo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   2
         Left            =   495
         TabIndex        =   25
         Top             =   3870
         Width           =   540
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   855
         TabIndex        =   24
         Top             =   4485
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   855
         TabIndex        =   23
         Top             =   4110
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Grabación Fichero Cheps"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   495
         TabIndex        =   19
         Top             =   450
         Width           =   5160
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   495
         TabIndex        =   18
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   855
         TabIndex        =   17
         Top             =   3180
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   855
         TabIndex        =   16
         Top             =   3540
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1440
         Picture         =   "frmAlmGrabChep.frx":0554
         ToolTipText     =   "Buscar fecha"
         Top             =   3180
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1440
         Picture         =   "frmAlmGrabChep.frx":05DF
         ToolTipText     =   "Buscar fecha"
         Top             =   3540
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   855
         TabIndex        =   15
         Top             =   1335
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   855
         TabIndex        =   14
         Top             =   1710
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   11
         Left            =   495
         TabIndex        =   13
         Top             =   1095
         Width           =   495
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1440
         MouseIcon       =   "frmAlmGrabChep.frx":066A
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   1335
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1440
         MouseIcon       =   "frmAlmGrabChep.frx":07BC
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   1710
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmAlmGrabChep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MANOLO +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmVar As frmManVariedad 'Variedad
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmCli As frmClientes 'Clientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmArt As frmManArtic ' articulos
Attribute frmArt.VB_VarHelpID = -1
Private WithEvents frmDes As frmDestCli 'Destinos de Clientes
Attribute frmDes.VB_VarHelpID = -1

Private WithEvents frmMensDestino As frmMensajes 'mensajes
Attribute frmMensDestino.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean

Dim Sql As String
Dim SqlDestino As String


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub frmDes_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Destinos
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMensDestino_DatoSeleccionado(CadenaSeleccion As String)

    If CadenaSeleccion <> "" Then
        SqlDestino = " and albaran.coddesti in (" & CadenaSeleccion & ")"
    Else
        SqlDestino = " and albaran.coddesti = -1 "
    End If

End Sub



Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim I As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String

Dim nRegs As Long
Dim b As Byte
Dim vsqlDestino As String


'    If Not DatosOK Then Exit Sub
    Sql = "SELECT  count(*) "
    Sql = Sql & " from albaran_envase, albaran, sartic "
    Sql = Sql & " where sartic.codtipar = 1 " ' el tipo de archivos es solamente cheps
    Sql = Sql & " and albaran.numalbar <> 0 "
    Sql = Sql & " and albaran_envase.tipomovi = 0 "
    
    If txtCodigo(0).Text <> "" Then Sql = Sql & " and albaran.codclien >= " & DBSet(txtCodigo(0).Text, "N")
    If txtCodigo(1).Text <> "" Then Sql = Sql & " and albaran.codclien <= " & DBSet(txtCodigo(1).Text, "N")
    
    '[Monica]07/01/2015: añadimos el destino
    'D/H Destino
    cDesde = Trim(txtCodigo(6).Text)
    cHasta = Trim(txtCodigo(7).Text)
    nDesde = txtNombre(6).Text
    nHasta = txtNombre(7).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        codigo = "{albaran.coddesti}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHDestino= """) Then Exit Sub
    End If
    
    vsqlDestino = ""
    If txtCodigo(6).Text <> "" Then vsqlDestino = vsqlDestino & " and destinos.coddesti >= " & DBSet(txtCodigo(6).Text, "N")
    If txtCodigo(7).Text <> "" Then vsqlDestino = vsqlDestino & " and destinos.coddesti <= " & DBSet(txtCodigo(7).Text, "N")
    
    If vsqlDestino <> "" And txtCodigo(6).Text <> txtCodigo(7).Text And txtCodigo(0).Text = txtCodigo(1).Text And txtCodigo(0).Text <> "" Then
        Set frmMensDestino = New frmMensajes

        frmMensDestino.OpcionMensaje = 21
        frmMensDestino.Label5 = "Destinos"
        frmMensDestino.cadWhere = vsqlDestino & " and destinos.codclien = " & txtCodigo(0).Text
        frmMensDestino.Show vbModal

        Set frmMensDestino = Nothing
        
        Sql = Sql & SqlDestino
    End If
    
    
    If txtCodigo(2).Text <> "" Then Sql = Sql & " and albaran.fechaalb >= " & DBSet(txtCodigo(2).Text, "F")
    If txtCodigo(3).Text <> "" Then Sql = Sql & " and albaran.fechaalb <= " & DBSet(txtCodigo(3).Text, "F")
    
    If txtCodigo(4).Text <> "" Then Sql = Sql & " and albaran_envase.codartic >= " & DBSet(txtCodigo(4).Text, "T")
    If txtCodigo(5).Text <> "" Then Sql = Sql & " and albaran_envase.codartic <= " & DBSet(txtCodigo(5).Text, "T")
    
    Sql = Sql & " and albaran.numalbar = albaran_envase.numalbar "
    Sql = Sql & " and albaran_envase.codartic = sartic.codartic "

    '[Monica]20/08/2012: quitamos los que no tienen codpalet o codcajas
    Sql = Sql & " and (albaran.codclien, albaran.coddesti)  in (select codclien, coddesti from destinos where trim(codpalet) <> '' or trim(codcajas) <> '') "


    nRegs = TotalRegistros(Sql)

    If nRegs <> 0 Then
        pb1.visible = True
        pb1.Max = nRegs
        pb1.Value = 0
        
        If GeneraFichero(nRegs) Then
            If CopiarFichero Then
                If MsgBox("¿Proceso realizado correctamente para actualizar?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    vParamAplic.NroLote = vParamAplic.NroLote + 1
                    b = vParamAplic.Modificar()
                    If b Then
                        MsgBox "Proceso realizado correctamente", vbExclamation
                    Else
                        MsgBox "No se ha actualizado el nro de lote correctamente.", vbExclamation
                    End If
                    cmdCancel_Click
                End If
            End If
        End If
    Else
        MsgBox "No hay datos entre esos límites.", vbExclamation
    End If

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(0)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim h As Integer, w As Integer, I As Integer
Dim List As Collection

    'Icono del formulario
    Me.Icon = frmPpal.Icon


    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
     For I = 0 To 3
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Next I
     For I = 6 To 7
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Next I

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    
    Me.pb1.visible = False
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = w + 70
    Me.Height = h + 350
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Articulos
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Variedades
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgFec_Click(Index As Integer)
'FEchas
    Dim esq, dalt As Long
    Dim obj As Object
    
    Set frmC = New frmCal

    esq = imgFec(Index).Left
    dalt = imgFec(Index).Top

    Set obj = imgFec(Index).Container

    While imgFec(Index).Parent.Name <> obj.Name
        esq = esq + obj.Left
        dalt = dalt + obj.Top
        Set obj = obj.Container
    Wend
       
    ' es desplega dalt i cap a la esquerra
    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + 420 + 30

    ' ***canviar l'index de imgFec pel 1r index de les imagens de buscar data***
    imgFec(2).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtCodigo(Index).Text <> "" Then frmC.NovaData = txtCodigo(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtCodigo(CByte(imgFec(2).Tag) + 2)
    ' ***************************
End Sub


Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1 'CLIENTE
            AbrirFrmClientes (Index)
        
        
        Case 2, 3 'ARTICULOS
            AbrirFrmArticulos (Index)
        
        Case 6, 7 'DESTINOS
            AbrirFrmDestinos (Index)
        
        
    End Select
    
    PonerFoco txtCodigo(indCodigo)
End Sub


Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'14/02/2007 antes
'    KEYpress KeyAscii
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'cliente desde
            Case 1: KEYBusqueda KeyAscii, 1 'cliente hasta
            Case 4: KEYBusqueda KeyAscii, 2 'articulo desde
            Case 5: KEYBusqueda KeyAscii, 3 'articulo hasta
            Case 2: KEYFecha KeyAscii, 2 'fecha desde
            Case 3: KEYFecha KeyAscii, 3 'fecha hasta
            Case 6: KEYBusqueda KeyAscii, 6 'destino desde
            Case 7: KEYBusqueda KeyAscii, 7 'destino hasta
            
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
            
        Case 0, 1 'CLIENTE
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "clientes", "nomclien", "codclien", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
        
            ' solo se puede introducir destino si cliente desde y hasta son iguales
            txtCodigo(6).Enabled = (txtCodigo(0).Text = txtCodigo(1).Text)
            imgBuscar(6).Enabled = txtCodigo(6).Enabled
            If Not txtCodigo(6).Enabled Then
                txtCodigo(6).Text = ""
                txtNombre(6).Text = ""
            End If
            txtCodigo(7).Enabled = (txtCodigo(0).Text = txtCodigo(1).Text)
            imgBuscar(7).Enabled = txtCodigo(7).Enabled
            If Not txtCodigo(7).Enabled Then
                txtCodigo(7).Text = ""
                txtNombre(7).Text = ""
            End If
            
            If Index = 1 Then
                If txtCodigo(6).Enabled Then
                    PonerFoco txtCodigo(6)
                Else
                    PonerFoco txtCodigo(16)
                End If
            End If
        
        
        
        
        Case 2, 3 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 4, 5 'ARTICULOS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "sartic", "nomartic", "codartic", "T")
'            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
    
        Case 6, 7 'DESTINO
            If txtCodigo(0).Text <> "" And txtCodigo(0).Text = txtCodigo(1).Text Then
                txtNombre(Index).Text = DevuelveDesdeBDNew(cAgro, "destinos", "nomdesti", "codclien", txtCodigo(0).Text, "N", , "coddesti", txtCodigo(Index).Text, "N")
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            End If
    
    
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 6720
        Me.FrameCobros.Width = 6555
        w = Me.FrameCobros.Width
        h = Me.FrameCobros.Height
    End If
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
End Sub

Private Function PonerDesdeHasta(codD As String, codH As String, nomD As String, nomH As String, param As String) As Boolean
'IN: codD,codH --> codigo Desde/Hasta
'    nomD,nomH --> Descripcion Desde/Hasta
'Añade a cadFormula y cadSelect la cadena de seleccion:
'       "(codigo>=codD AND codigo<=codH)"
' y añade a cadParam la cadena para mostrar en la cabecera informe:
'       "codigo: Desde codD-nomd Hasta: codH-nomH"
Dim devuelve As String
Dim devuelve2 As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(codD, codH, codigo, TipCod)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    If TipCod <> "F" Then 'Fecha
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        devuelve2 = CadenaDesdeHastaBD(codD, codH, codigo, TipCod)
        If devuelve2 = "Error" Then Exit Function
        If Not AnyadirAFormula(cadSelect, devuelve2) Then Exit Function
    End If
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            cadParam = cadParam & AnyadirParametroDH(param, codD, codH, nomD, nomH)
            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function

Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .EnvioEMail = False
        .Opcion = 0
        .Show vbModal
    End With
End Sub

Private Sub AbrirFrmClientes(indice As Integer)
    indCodigo = indice
    Set frmCli = New frmClientes
    frmCli.DatosADevolverBusqueda = "0|2|"
    frmCli.Show vbModal
    Set frmCli = Nothing
End Sub

Private Sub AbrirFrmArticulos(indice As Integer)
    indCodigo = indice + 2
    Set frmArt = New frmManArtic
    frmArt.DatosADevolverBusqueda = "0|1|"
    frmArt.Show vbModal
    Set frmArt = Nothing
End Sub

Private Sub AbrirFrmDestinos(indice As Integer)
    If txtCodigo(0).Text = "" Or txtCodigo(0).Text <> txtCodigo(1).Text Then Exit Sub

    indCodigo = indice
    Set frmDes = New frmDestCli
    frmDes.DatosADevolverBusqueda = "0|1|"
'    frmDes.DeConsulta = True
    frmDes.Cliente = txtCodigo(0).Text
    frmDes.CodigoActual = txtCodigo(indCodigo)
    frmDes.Show vbModal
    Set frmDes = Nothing
End Sub

 
Private Sub AbrirVisReport()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = cadFormula
        '.SoloImprimir = (Me.OptVisualizar(indFrame).Value = 1)
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        '##descomen
'        .MostrarTree = MostrarTree
'        .Informe = MIPATH & Nombre
'        .InfConta = InfConta
        '##
        
'        If NombreSubRptConta <> "" Then
'            .SubInformeConta = NombreSubRptConta
'        Else
'            .SubInformeConta = ""
'        End If
        '##descomen
'        .ConSubInforme = ConSubInforme
        '##
        .Opcion = ""
'        .ExportarPDF = (chkEMAIL.Value = 1)
        .Show vbModal
    End With
    
'    If Me.chkEMAIL.Value = 1 Then
'    '####Descomentar
'        If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
'    End If
    Unload Me
End Sub

Private Sub AbrirEMail()
    If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
End Sub


Private Function HayRegistros(cTabla As String, cWhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
Dim Rs As ADODB.Recordset

    Sql = "Select * FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    Sql = Sql & " group by 1 "
    Sql = Sql & " having sum(totalfac) > " & DBSet(txtCodigo(6).Text, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Rs.EOF Then
        MsgBox "No hay datos para mostrar en el Informe.", vbInformation
        HayRegistros = False
    Else
        HayRegistros = True
    End If

End Function

Private Function GeneraFichero(Numreg As Long) As Boolean
Dim NFich1 As Integer
Dim NFich2 As Integer
Dim Rs As ADODB.Recordset
Dim cad As String
Dim Sql As String
Dim AntLetraSer As String
Dim ActLetraSer As String
Dim AntNumfactu As Long
Dim ActNumfactu As Long
Dim v_Hayreg As Integer
Dim AntTarjet As Long
Dim ActTarjet As Long

Dim NomSocio As String
Dim NomArtic As String
Dim b As Boolean
Dim Mens As String
Dim NumLin As Long
Dim CanTotal As Long
Dim Almacen As Integer
Dim ClienteChep As String

Dim Longitud As Integer
Dim Insertar As Boolean

    On Error GoTo EGen
    
    GeneraFichero = False

    NFich1 = FreeFile
    Open App.path & "\" & Format(vParamAplic.NroFiche, "0000") & Format(vParamAplic.NroLote + 1, "0000") & ".txt" For Output As #NFich1

    Set Rs = New ADODB.Recordset
    
    'partimos de la tabla de lineas de albaranes
    Sql = "SELECT albaran.fechaalb, albaran_envase.cantidad, albaran.numalbar, albaran.coddesti, "
    Sql = Sql & " albaran.codclien, albaran_envase.codartic, sartic.codigoea, "
    Sql = Sql & " destinos.nomdesti, destinos.domdesti, destinos.pobdesti, destinos.codpobla, "
    Sql = Sql & " destinos.prodesti, paises.letraspais, "
    Sql = Sql & " albaran.codalmac, destinos.codpalet, destinos.codcajas "
    Sql = Sql & " from albaran_envase, albaran, sartic, destinos, paises "
    Sql = Sql & " where sartic.codtipar = '1' " ' el tipo de archivos es solamente cheps
    Sql = Sql & " and albaran.numalbar <> 0 "
    Sql = Sql & " and albaran_envase.tipomovi = 0 "
    
    If txtCodigo(0).Text <> "" Then Sql = Sql & " and albaran.codclien >= " & DBSet(txtCodigo(0).Text, "N")
    If txtCodigo(1).Text <> "" Then Sql = Sql & " and albaran.codclien <= " & DBSet(txtCodigo(1).Text, "N")
    
    If txtCodigo(2).Text <> "" Then Sql = Sql & " and albaran.fechaalb >= " & DBSet(txtCodigo(2).Text, "F")
    If txtCodigo(3).Text <> "" Then Sql = Sql & " and albaran.fechaalb <= " & DBSet(txtCodigo(3).Text, "F")
    
    If txtCodigo(4).Text <> "" Then Sql = Sql & " and albaran_envase.codartic >= " & DBSet(txtCodigo(4).Text, "T")
    If txtCodigo(5).Text <> "" Then Sql = Sql & " and albaran_envase.codartic <= " & DBSet(txtCodigo(5).Text, "T")
    
    '[Monica]07/01/2015: añadimos el destino
    If txtCodigo(6).Text <> "" Then Sql = Sql & " and albaran.coddesti >= " & DBSet(txtCodigo(6).Text, "N")
    If txtCodigo(7).Text <> "" Then Sql = Sql & " and albaran.coddesti <= " & DBSet(txtCodigo(7).Text, "N")
    
    Sql = Sql & " and albaran.numalbar = albaran_envase.numalbar "
    Sql = Sql & " and albaran_envase.codartic = sartic.codartic "
    Sql = Sql & " and albaran.codclien = destinos.codclien and albaran.coddesti = destinos.coddesti "
    Sql = Sql & " and destinos.codpaise = paises.codpaise "
    
    '[Monica]20/08/2012: quitamos los destinos que no tienen codpalet o codcajas
    Sql = Sql & " and (trim(destinos.codpalet) <> '' or trim(destinos.codcajas) <> '') "
    
    '[Monica]07/01/2015: añadimos el destino
    If SqlDestino <> "" Then Sql = Sql & SqlDestino
    
    
    Sql = Sql & " order by 1, 3"

    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    ' cabecera
    cad = "*****+FROM+CHEP-ES"                          '   column 0,"*****+FROM+CHEP-ES",
'    cad = cad & "5000448311"                            '   column 18,"5000448311",
    cad = cad & vParamAplic.NroCheps                     '   column 18,"5000448429",
    cad = cad & "+RCVD+"                                '   column 28,"+RCVD+",
    cad = cad & Format(Now, "yyyymmdd")                 '   column 34,today using "yyyymmdd", {fecha del fichero}
    cad = cad & "+FREF+"                                '   column 42,"+FREF+",
    cad = cad & Format(vParamAplic.NroLote + 1, "0000")   '   column 48,numelote using "&&&&",
    cad = cad & "+NORC+"                                '   column 52,"+NORC+",
    cad = cad & Format(Numreg, "000000")              '   column 58,v_numreg using "&&&&&&",
    cad = cad & "+SEPR+"                                '   column 64,"+SEPR+",
    cad = cad & "%"                                     '   column 70,"%",
    cad = cad & "+VERS+1.03+*****"                      '   column 71,"+VERS+1.03+*****",
        
    Print #NFich1, cad
    
    b = True
    NumLin = 0
    v_Hayreg = 0
    CanTotal = 0
    While Not Rs.EOF
        v_Hayreg = 1
        
        pb1.Value = pb1.Value + 1
        
        
        NumLin = NumLin + 1
        
            
        cad = "LI=%"                                    '   column 0,"LI=%",
        cad = cad & Format(NumLin, "00000")             '   column 4,v_contar using "&&&&&",
        cad = cad & "%1%ES%"                            '   column 9,"%1%ES%",
        cad = cad & "SA%"                               '   column 15,"SA%",
        
'--monica
'        cad = cad & "5000448311"                        '   column 18,"5000448311",
'++monica: cambiado por esto
'                                                        If origensa = 0 Then Column 18, "5000448779"
'                                                        If origensa = 1 Then Column 18, "5000448429"
        Almacen = DBLet(Rs.Fields(13).Value, "N")
        If Almacen <> 0 Then
            ClienteChep = DevuelveDesdeBDNew(cAgro, "salmpr", "nrocheps", "codalmac", CStr(Almacen), "N")
            '[Monica]22/08/2012: añadido el formato al cliente chep
            cad = cad & Format(ClienteChep, "0000000000")
        End If
'++
'        cad = cad & "%CC%"                              '   column 28,"%CC%",
        
'--monica
'        cad = cad & Format(DBLet(Rs.Fields(3).Value, "N"), "00000000")       '   column 32,v_destin using "&&&&&&&&",
'++ cambiado por esto:
        If DBLet(Rs.Fields(6).Value, "N") < 100 Then
            Longitud = Len(DBLet(Rs.Fields(14).Value, "T"))
            If Longitud <= 9 Then cad = cad & "%CC%"      '   column 28,"%CC%",
            If Longitud = 10 Then cad = cad & "%SA%"
'            cad = cad & RellenaABlancos(DBLet(RS.Fields(14).Value, "T"), True, 10)
            '[Monica]22/08/2012: dependiendo de la cooperativa sale con un formato u otro
            If vParamAplic.Cooperativa = 11 Then
                If Longitud >= 9 Then
                    cad = cad & "********"
                Else
                    cad = cad & Format(DBLet(Rs.Fields(14).Value, "N"), "00000000")
                End If
            Else
                cad = cad & Format(DBLet(Rs.Fields(14).Value, "N"), "0000000000")
            End If
        End If
        
        If DBLet(Rs.Fields(6).Value, "N") >= 100 Then
            Longitud = Len(DBLet(Rs.Fields(15).Value, "T"))
            If Longitud <= 9 Then cad = cad & "%CC%"      '   column 28,"%CC%",
            If Longitud = 10 Then cad = cad & "%SA%"
'            cad = cad & RellenaABlancos(DBLet(RS.Fields(15).Value, "T"), True, 10)
            '[Monica]22/08/2012: dependiendo de la cooperativa sale con un formato u otro
            If vParamAplic.Cooperativa = 11 Then
                If Longitud >= 9 Then
                    cad = cad & "********"
                Else
                    cad = cad & Format(DBLet(Rs.Fields(15).Value, "N"), "00000000")
                End If
            Else
                cad = cad & Format(DBLet(Rs.Fields(15).Value, "N"), "0000000000")
            End If
        End If
        
        cad = cad & "%91%"                              '               column 42,"%91%",
        cad = cad & Format(DBLet(Rs.Fields(6).Value, "N"), "00000")          '   column 46,codcheps using "&&&&&",
        cad = cad & "%"                                 '   column 51,"%",
        cad = cad & Format(DBLet(Rs.Fields(0).Value, "F"), "yyyymmdd")        '   column 52,fechamov using "yyyymmdd",
        cad = cad & "%"                                 '   column 60,"%",
        cad = cad & "%"                                 '    column 61,"%",
        cad = cad & Format(DBLet(Rs.Fields(1).Value, "N"), "0000")           '    column 62,canenvas using "&&&&",
        cad = cad & "%"                                 '   column 66,"%",
        cad = cad & Format(DBLet(Rs.Fields(2).Value, "N"), "000000") & Space(7) ' column 67,codexped using "&&&&&&" && "       ",
        cad = cad & "%"                                 '   column 80,"%",
        cad = cad & "%"                                 '   column 81,"%",
        cad = cad & "%"                                 '   column 82,"%",
        cad = cad & "%%%%%%"                            '   column 83,"%%%%%%",
        cad = cad & RellenaABlancos(DBLet(Rs.Fields(7).Value, "T"), True, 25) '   column 89,nomdesti,
        cad = cad & "%"                                 '   column 114,"%",
        cad = cad & RellenaABlancos(DBLet(Rs.Fields(8).Value, "T"), True, 30) '   column 115,nomdirec,
        cad = cad & "%"                                 '   column 145,"%",
        cad = cad & RellenaABlancos(DBLet(Rs.Fields(9).Value, "T"), True, 25) '   column 146,nomciuda,
        cad = cad & "%"                                 '   column 171,"%",
        cad = cad & RellenaABlancos(DBLet(Rs.Fields(10).Value, "T"), True, 10) '   column 172,codipost,
        cad = cad & "%"                                 '   column 182,"%",
        cad = cad & RellenaABlancos(DBLet(Rs.Fields(11).Value, "T"), True, 25) '   column 183,nomprovi,
        cad = cad & "%"                                 '   column 208,"%",
        cad = cad & RellenaABlancos(DBLet(Rs.Fields(12).Value, "T"), True, 2)  '   column 209,letraspa,
        cad = cad & "%"                                 '   column 211,"%",
        cad = cad & "%%"                                '   column 212,"%%", {telefonos}
        cad = cad & "<"                                 '   column 214,"<",
        
        Print #NFich1, cad
            
        CanTotal = CanTotal + DBLet(Rs.Fields(1).Value, "N")
        
        Rs.MoveNext
    Wend
   
    If v_Hayreg = 1 Then
        cad = "*****+NORC+"                             '   column  0,"*****+NORC+",
        cad = cad & Format(Numreg, "000000")            '   column 11,v_numreg using "&&&&&&",
        cad = cad & "+SQTY+"                            '   column 17,"+SQTY+",
        cad = cad & Format(CanTotal, "000000")          '   column 23,t_totale using "&&&&&&",
        cad = cad & "+EOF"                              '   column 29,"+EOF",
    
        Print #NFich1, cad
    End If
    
    Close (NFich1)
    Set Rs = Nothing
    
    pb1.visible = False
    
    GeneraFichero = True
    Exit Function
    
EGen:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description & vbCrLf
    End If
End Function


Private Function RellenaABlancos(cadena As String, PorLaDerecha As Boolean, Longitud As Integer) As String
Dim cad As String
    
    cad = Space(Longitud)
    If PorLaDerecha Then
        cad = cadena & cad
        RellenaABlancos = Left(cad, Longitud)
    Else
        cad = cad & cadena
        RellenaABlancos = Right(cad, Longitud)
    End If
    
End Function

Public Function CopiarFichero() As Boolean
Dim nomFich As String

On Error GoTo ecopiarfichero

    CopiarFichero = False
    ' abrimos el commondialog para indicar donde guardarlo
'    Me.CommonDialog1.InitDir = App.path

    Me.CommonDialog1.DefaultExt = "txt"
    
    CommonDialog1.Filter = "Archivos txt|txt|"
    CommonDialog1.FilterIndex = 1
    
    ' copiamos el primer fichero
    CommonDialog1.FileName = Format(vParamAplic.NroFiche, "0000") & Format(vParamAplic.NroLote + 1, "0000") & ".txt"
    Me.CommonDialog1.ShowSave
    
    
    
    If CommonDialog1.FileName <> "" Then
        FileCopy App.path & "\" & Format(vParamAplic.NroFiche, "0000") & Format(vParamAplic.NroLote + 1, "0000") & ".txt", CommonDialog1.FileName
    End If
    
    CopiarFichero = True
    Exit Function


ecopiarfichero:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    End If
    Err.Clear

End Function

