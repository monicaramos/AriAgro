VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListCyC 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6675
   Icon            =   "frmListCyC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   6675
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
      Height          =   5535
      Left            =   45
      TabIndex        =   9
      Top             =   0
      Width           =   6555
      Begin VB.CheckBox chkResumen 
         Caption         =   "S�lo resumen"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3780
         TabIndex        =   24
         Top             =   4140
         Width           =   2100
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   6
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   4050
         Width           =   1350
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   1710
         MaxLength       =   7
         TabIndex        =   1
         Tag             =   "N� Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   1485
         Width           =   870
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   1695
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "N� Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   1095
         Width           =   870
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   2550
         Width           =   1350
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   2190
         Width           =   1350
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   8
         Top             =   4950
         Width           =   1065
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3900
         TabIndex        =   7
         Top             =   4950
         Width           =   1065
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   4
         Top             =   3180
         Width           =   870
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   5
         Top             =   3555
         Width           =   870
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   3180
         Width           =   3630
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text5"
         Top             =   3555
         Width           =   3630
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   6
         Left            =   1440
         Picture         =   "frmListCyC.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   4050
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   3
         Left            =   450
         TabIndex        =   23
         Top             =   4050
         Width           =   1860
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   450
         TabIndex        =   22
         Top             =   4590
         Width           =   5325
      End
      Begin VB.Label Label1 
         Caption         =   "Facturas Pendientes"
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
         Left            =   450
         TabIndex        =   21
         Top             =   270
         Width           =   5160
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   780
         TabIndex        =   20
         Top             =   1485
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   780
         TabIndex        =   19
         Top             =   1125
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nro.Albar�n"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   240
         Index           =   2
         Left            =   450
         TabIndex        =   18
         Top             =   855
         Width           =   1185
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Albar�n"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   420
         TabIndex        =   17
         Top             =   1890
         Width           =   1860
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   780
         TabIndex        =   16
         Top             =   2190
         Width           =   600
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   780
         TabIndex        =   15
         Top             =   2550
         Width           =   645
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1440
         Picture         =   "frmListCyC.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   2190
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1440
         Picture         =   "frmListCyC.frx":0122
         ToolTipText     =   "Buscar fecha"
         Top             =   2550
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   780
         TabIndex        =   14
         Top             =   3180
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   780
         TabIndex        =   13
         Top             =   3555
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   240
         Index           =   11
         Left            =   465
         TabIndex        =   12
         Top             =   2895
         Width           =   720
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1440
         MouseIcon       =   "frmListCyC.frx":01AD
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   3180
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1440
         MouseIcon       =   "frmListCyC.frx":02FF
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   3555
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmListCyC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MANOLO +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar n� oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexi�n a BD Ariges  2.- Conexi�n a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmVar As frmManVariedad 'Variedad
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmCli As frmClientes 'Clientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmTra As frmManAgencias 'Agencias de transporte
Attribute frmTra.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadselect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'n� de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'C�digo para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String

Dim Sql As String

InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'A�adir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
     '======== FORMULA  ====================================
    'Seleccionar registros de la empresa conectada
'    Codigo = "{" & tabla & ".codempre}=" & vEmpresa.codEmpre
'    If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
'    If Not AnyadirAFormula(cadSelect, Codigo) Then Exit Sub
    
    
    'D/H Cliente
    cDesde = Trim(txtCodigo(0).Text)
    cHasta = Trim(txtCodigo(1).Text)
    nDesde = txtNombre(0).Text
    nHasta = txtNombre(1).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".codclien}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCliente= """) Then Exit Sub
    End If
    
    'D/H Nro de Albar�n
    cDesde = Trim(txtCodigo(4).Text)
    cHasta = Trim(txtCodigo(5).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{albaran.numalbar}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHAlbaran= """) Then Exit Sub
    End If
    
    'D/H Fecha Albar�n
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".fechaalb}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    '[Monica]20/02/2012: Son clientes asegurados los que tienen limiteriesgos  <> 0  (antes nroseguro)
'    If AnyadirAFormula(cadSelect, "clientes.nroseguro <> '' and clientes.nroseguro is not null ") = False Then Exit Sub

'    If AnyadirAFormula(cadSelect, "clientes.limiteriesgos <> 0 and clientes.limiteriesgos is not null ") = False Then Exit Sub
    
    cadTABLA = tabla & " INNER JOIN clientes ON albaran.codclien = clientes.codclien "
    
    cadParam = cadParam & "pSoloResumen=" & chkResumen.Value & "|"
    numParam = numParam + 1
    
    If HayRegistros(cadTABLA, cadselect) Then
        If CargarTemporal(cadTABLA, cadselect) Then
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            cadTitulo = "Facturas Pendientes"
            cadNombreRPT = "rFacturasPdtes.rpt"
            
            LlamarImprimir
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
     For H = 0 To 1
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Next H

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, H, W
    indFrame = 5
    tabla = "albaran"
    Me.Label2.Caption = ""
    Me.Refresh
        
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
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
    PonerFoco txtCodigo(CByte(imgFec(2).Tag))
    ' ***************************
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1 'CLIENTE
            AbrirFrmClientes (Index)
        
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub Optcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub OptNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
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
            Case 2: KEYFecha KeyAscii, 2 'fecha desde
            Case 3: KEYFecha KeyAscii, 3 'fecha hasta
            Case 6: KEYFecha KeyAscii, 6 'fecha de calculo
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
Dim Cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
            
        Case 0, 1 'CLIENTE
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "clientes", "nomclien", "codclien", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
        
        Case 2, 3, 6 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 4, 5 'ALBARANES
            If txtCodigo(Index).Text <> "" Then PonerFormatoEntero txtCodigo(Index)
        
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 5760
        Me.FrameCobros.Width = 6690
        W = Me.FrameCobros.Width
        H = Me.FrameCobros.Height
    End If
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadselect = ""
    cadParam = ""
    numParam = 0
End Sub

Private Function PonerDesdeHasta(codD As String, codH As String, nomD As String, nomH As String, param As String) As Boolean
'IN: codD,codH --> codigo Desde/Hasta
'    nomD,nomH --> Descripcion Desde/Hasta
'A�ade a cadFormula y cadSelect la cadena de seleccion:
'       "(codigo>=codD AND codigo<=codH)"
' y a�ade a cadParam la cadena para mostrar en la cabecera informe:
'       "codigo: Desde codD-nomd Hasta: codH-nomH"
Dim devuelve As String
Dim devuelve2 As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(codD, codH, Codigo, TipCod)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    If TipCod <> "F" Then 'Fecha
        If Not AnyadirAFormula(cadselect, devuelve) Then Exit Function
    Else
        devuelve2 = CadenaDesdeHastaBD(codD, codH, Codigo, TipCod)
        If devuelve2 = "Error" Then Exit Function
        If Not AnyadirAFormula(cadselect, devuelve2) Then Exit Function
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
        .FormulaSeleccion = "{tmpinformes.codusu} = " & vUsu.Codigo
        .OtrosParametros = cadParam
        .NumeroParametros = numParam + 1
        .SoloImprimir = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .EnvioEMail = False
        .Opcion = 0
        .ConSubInforme = True
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

Private Sub AbrirFrmVariedades(indice As Integer)
    indCodigo = indice
    Set frmVar = New frmManVariedad
    frmVar.DatosADevolverBusqueda = "0|1|"
    frmVar.DeConsulta = True
    frmVar.CodigoActual = txtCodigo(indCodigo)
    frmVar.Show vbModal
    Set frmVar = Nothing
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
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Rs.EOF Then
        MsgBox "No hay datos para mostrar en el Informe.", vbInformation
        HayRegistros = False
    Else
        HayRegistros = True
    End If

End Function

Private Function CargarTemporal(cTabla As String, cadwhere As String) As Boolean
Dim Sql As String
Dim SQL1 As String
Dim Sql2 As String
Dim Codiva As String
Dim PorcIva As String
Dim i As Integer
Dim HayReg As Integer
Dim b As Boolean
Dim Registro As String
Dim CADENA As String
Dim vCliente As CCliente
Dim Rs As ADODB.Recordset
Dim Importe As Currency
Dim ImporteFacturado As Currency
Dim Dias As Long

Dim Primero As Boolean
Dim MenError As String

'[Monica]23/05/2013: los calculos ahora se hacen con respecto a la fecha de factura si el albaran est� facturado
Dim FechaFactura As String
Dim vSQL As String

On Error GoTo eCargarTemporal

    HayReg = 0
    
    CargarTemporal = False
    
    conn.Execute "delete from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
        
    If cadwhere <> "" Then
        cadwhere = QuitarCaracterACadena(cadwhere, "{")
        cadwhere = QuitarCaracterACadena(cadwhere, "}")
        cadwhere = QuitarCaracterACadena(cadwhere, "_1")
    End If
                                  '(codusu, albaran,  linea, fecalbar,cliente, imp.factu,imp.provi,fec.pag,dias,   asterisco)
    Sql = "insert into tmpinformes (codusu, importe1, campo1, fecha1, codigo1, importe2, importe3, fecha2, campo2, nombre1) values "
    
    Sql2 = "select albaran.fechaalb, albaran.codclien, clientes.tipoiva, clientes.diasasegurados, albaran_variedad.* from albaran_variedad, albaran, clientes where albaran.numalbar = albaran_variedad.numalbar "
    Sql2 = Sql2 & " and albaran.codclien = clientes.codclien "
    If cadwhere <> "" Then Sql2 = Sql2 & " and " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CADENA = ""
    
    While Not Rs.EOF
        Registro = "(" & vUsu.Codigo & "," & DBSet(Rs!NumAlbar, "N") & "," & DBSet(Rs!NumLinea, "N") & ","
        Registro = Registro & DBSet(Rs!FechaAlb, "F") & "," & DBSet(Rs!CodClien, "N") & ","
        If AlbaranFacturado(CCur(Rs!NumAlbar), CCur(Rs!NumLinea)) Then
            FechaFactura = ""
           '[Monica]16/04/2010:antes FacturaCobradaTesoreria
           ' If FacturaCobradaTesoreria(CCur(RS!numalbar), CCur(RS!numlinea)) Then
            If AlbaranCobradoTesoreria(CCur(Rs!NumAlbar), CCur(Rs!NumLinea)) Then
            ' si la factura est� cobrada en tesoreria no hacemos nada.
                Registro = ""
'                Importe = CCur(ImporteAlbaranFacturado(CCur(Rs!numalbar), CCur(Rs!numlinea)))
'
'                Sql2 = "select codigiva from facturas_variedad where numalbar = " & DBSet(Rs!numalbar, "N")
'                Sql2 = Sql2 & " and numlinealbar = " & DBSet(Rs!numlinea, "N")
'                Codiva = DevuelveValor(Sql2)
'
'                PorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", Codiva, "N")
'
'                ' obtenemos el importe con iva
'                Importe = Importe + Round2(Importe * CCur(PorcIva) / 100, 2)
'
'                Registro = Registro & DBSet(Importe, "N") & ",0,"
            Else
                'ImporteAlbaranFacturado
                ' albaran facturado pero no cobrado
                '[Monica]26/09/2011:antes ImporteFacturado = CCur(ImporteAlbaranFacturado(CCur(RS!numalbar), CCur(RS!numlinea)))
                Dim Parcial As Boolean
                Parcial = False
                ImporteFacturado = CCur(ImporteAlbaranFacturadoNoCobrado(CCur(Rs!NumAlbar), CCur(Rs!NumLinea), Parcial))
                If Not Parcial Then
                    Sql2 = "select codigiva from facturas_variedad where numalbar = " & DBSet(Rs!NumAlbar, "N")
                    Sql2 = Sql2 & " and numlinealbar = " & DBSet(Rs!NumLinea, "N")
                    Codiva = DevuelveValor(Sql2)
    
                    PorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", Codiva, "N")
    
                    ' obtenemos el importe con iva
                    ImporteFacturado = ImporteFacturado + Round2(ImporteFacturado * CCur(PorcIva) / 100, 2)
                End If
                Importe = 0
                 
                vSQL = "select max(fecfactu) from facturas_variedad where numalbar = " & DBSet(Rs!NumAlbar, "N")
                vSQL = vSQL & " and numlinealbar = " & DBSet(Rs!NumLinea, "N")
                
                FechaFactura = DevuelveValor(vSQL)
                
'[Monica]:02/02/2010 o calculamos el importe facturado o el provisional
'                ' calculamos el importe provisional
'                Importe = Round2(DBLet(RS!Pesoneto, "N") * DBLet(RS!preciopro, "N"), 2)
'                Select Case RS!TipoIVA
'                    Case 0
'                        Codiva = DevuelveDesdeBDNew(cAgro, "variedades", "codigiva", "codvarie", RS!codvarie, "N")
'                    Case 1
'                        Codiva = vParamAplic.CodIvaExento
'                    Case 2
'                        Codiva = vParamAplic.CodIvaRecargo
'                End Select
'                PorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", Codiva, "N")
'
'                ' obtenemos el importe con iva
'                Importe = Importe + Round2(Importe * CCur(PorcIva) / 100, 2)
            
                Registro = Registro & DBSet(ImporteFacturado, "N") & "," & DBSet(Importe, "N") & ","
            End If
        Else
            FechaFactura = CStr(DBLet(Rs!FechaAlb, "F"))
            ' calculamos el importe provisional
            Importe = Round2(DBLet(Rs!Pesoneto, "N") * DBLet(Rs!preciopro, "N"), 2)
            Select Case Rs!TipoIva
                Case 0
                    Codiva = DevuelveDesdeBDNew(cAgro, "variedades", "codigiva", "codvarie", Rs!codvarie, "N")
                Case 1
                    Codiva = vParamAplic.CodIvaExento
                Case 2
                    Codiva = vParamAplic.CodIvaRecargo
            End Select
            PorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", Codiva, "N")
            
            ' obtenemos el importe con iva
            Importe = Importe + Round2(Importe * CCur(PorcIva) / 100, 2)
        
            Registro = Registro & "0," & DBSet(Importe, "N") & ","
        End If
        
        ' si hemos cargado algo lo metremos en la cadena
        If Registro <> "" Then
            'fecha de pago                                                              '[Monica]23/05/2013: cambiado RS!FechaAlb
            Registro = Registro & DBSet(DateAdd("d", CDbl(DBLet(Rs!Diasasegurados, "N")), CDate(FechaFactura)), "F") & ","
            'dias
            Dias = CDate(txtCodigo(6).Text) - CDate(FechaFactura) 'DBLet(RS!FechaAlb, "F")
            Registro = Registro & DBSet(Dias, "N") & ","
            '[Monica]28/02/2012: a�adido el igual en : DBLet(RS!Diasasegurados, "N") >= 0
            If Dias > DBLet(Rs!Diasasegurados, "N") And DBLet(Rs!Diasasegurados, "N") >= 0 Then
                Registro = Registro & "'*'),"
            Else
                Registro = Registro & "''),"
            End If
        
            CADENA = CADENA & Registro
        End If
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    ' quitamos la ultima coma
    If Len(CADENA) > 0 Then
        CADENA = Mid(CADENA, 1, Len(CADENA) - 1)
    
        conn.Execute Sql & CADENA
    End If
        
    '[Monica]06/10/2011
    ' quitamos las cantidades que esten duplicadas en el listado que se correspondan con facturas parcialmente cobradas
    '             cliente, imp.factu, albaran,  linea
    Sql = "select codigo1, importe2, importe1, campo1 from tmpinformes "
    Sql = Sql & " where codusu = " & vUsu.Codigo & " and (codigo1,importe2) in ( "
    Sql = Sql & " select codigo1, importe2 from (select codigo1,importe2,count(*)"
    Sql = Sql & " from tmpinformes where codusu = " & vUsu.Codigo & " and importe2 <> 0"
    Sql = Sql & " group by 1,2"
    Sql = Sql & " having count(*) > 1"
    Sql = Sql & " order by codigo1) aaaaa )"
    Sql = Sql & " order by 1,2,3,4"

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Primero = True
    MenError = ""
    
    While Not Rs.EOF And Not MenError <> ""
        If AlbaranDeFacturaCobradaParcialmente(Rs.Fields(2).Value, Rs.Fields(3).Value, MenError) Then
            If Primero Then
                Primero = False
            Else
                Sql = "update tmpinformes set importe2 = 0 where codusu = " & vUsu.Codigo & " and importe1 = " & DBSet(Rs.Fields(2).Value, "N")
                Sql = Sql & " and campo1 = " & DBSet(Rs.Fields(3).Value, "N")
                            
                conn.Execute Sql
            End If
        End If
        Rs.MoveNext
    Wend
        
    Set Rs = Nothing
    
    If MenError = "" Then
        CargarTemporal = True

        Exit Function
    End If
    
    
eCargarTemporal:
    MuestraError Err.Number, "Cargando Temporal", Err.Description
End Function

Private Function DatosOk() As Boolean

    DatosOk = False
        
    If txtCodigo(6).Text = "" Then
        MsgBox "Debe introducir valor para el campo Fecha.", vbExclamation
        PonerFoco txtCodigo(6)
        Exit Function
    End If
        
        
    DatosOk = True

End Function


Private Function AlbaranDeFacturaCobradaParcialmente(Albaran As Long, Linea As Long, MenError As String) As Boolean
Dim Sql As String

    On Error GoTo eAlbaranDeFacturaCobradaParcialmente

    If vParamAplic.ContabilidadNueva Then
        Sql = "select  sum(if(isnull(impcobro),0,impcobro)) from ariconta" & vParamAplic.NumeroConta & ".cobros where (numserie, numfactu, fecfactu) in ("
        Sql = Sql & " select stipom.letraser, facturas_variedad.numfactu, facturas_variedad.fecfactu "
        Sql = Sql & " from facturas_variedad, usuarios.stipom stipom "
        Sql = Sql & " where facturas_variedad.numalbar = " & Albaran
        Sql = Sql & " and facturas_variedad.numlinealbar = " & Linea
        Sql = Sql & " and facturas_variedad.codtipom = stipom.codtipom) "
    Else
        Sql = "select  sum(if(isnull(impcobro),0,impcobro)) from conta" & vParamAplic.NumeroConta & ".scobro where (numserie, codfaccl, fecfaccl) in ("
        Sql = Sql & " select stipom.letraser, facturas_variedad.numfactu, facturas_variedad.fecfactu "
        Sql = Sql & " from facturas_variedad, usuarios.stipom stipom "
        Sql = Sql & " where facturas_variedad.numalbar = " & Albaran
        Sql = Sql & " and facturas_variedad.numlinealbar = " & Linea
        Sql = Sql & " and facturas_variedad.codtipom = stipom.codtipom) "
    End If
    AlbaranDeFacturaCobradaParcialmente = (DevuelveValor(Sql) <> 0)

    Exit Function
    
eAlbaranDeFacturaCobradaParcialmente:
    MenError = "Actualizando Albaranes de Factura cobrada parcialmente"
End Function

