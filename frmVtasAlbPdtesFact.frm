VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmVtasAlbPdtesFact 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6600
   Icon            =   "frmVtasAlbPdtesFact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   6600
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
      Height          =   7590
      Left            =   120
      TabIndex        =   12
      Top             =   45
      Width           =   6375
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   2
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Tag             =   "Tipo Variedad|N|N|||variedades|tipovariedad||N|"
         Top             =   5130
         Width           =   1440
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Tag             =   "Tipo Variedad|N|N|||variedades|tipovariedad||N|"
         Top             =   4680
         Width           =   1440
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   4815
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   4260
         Width           =   1050
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "Tipo Variedad|N|N|||variedades|tipovariedad||N|"
         Top             =   4230
         Width           =   1050
      End
      Begin VB.Frame Frame1 
         Caption         =   "Orden"
         ForeColor       =   &H00972E0B&
         Height          =   735
         Left            =   540
         TabIndex        =   26
         Top             =   5670
         Width           =   2895
         Begin VB.OptionButton Option1 
            Caption         =   "Cliente"
            Height          =   195
            Index           =   1
            Left            =   1665
            TabIndex        =   28
            Top             =   315
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Variedad"
            Height          =   195
            Index           =   0
            Left            =   315
            TabIndex        =   27
            Top             =   315
            Width           =   1185
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Text5"
         Top             =   3675
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text5"
         Top             =   3300
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   5
         Top             =   3690
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1830
         MaxLength       =   6
         TabIndex        =   4
         Top             =   3300
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1695
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1335
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4920
         TabIndex        =   11
         Top             =   6915
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3735
         TabIndex        =   10
         Top             =   6915
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   2
         Top             =   2280
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   3
         Top             =   2655
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   2280
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text5"
         Top             =   2655
         Width           =   3135
      End
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   255
         Left            =   540
         TabIndex        =   32
         Top             =   6540
         Visible         =   0   'False
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Precio"
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
         Index           =   4
         Left            =   570
         TabIndex        =   34
         Top             =   5130
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Mercancia"
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
         Index           =   31
         Left            =   570
         TabIndex        =   33
         Top             =   4680
         Width           =   1065
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   0
         Left            =   5910
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   4260
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   6
         Left            =   4500
         Picture         =   "frmVtasAlbPdtesFact.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   4260
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
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
         Index           =   3
         Left            =   3360
         TabIndex        =   31
         Top             =   4290
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Albarán"
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
         Index           =   1
         Left            =   570
         TabIndex        =   30
         Top             =   4260
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Albaranes Pendientes de Facturar"
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
         Index           =   0
         Left            =   585
         TabIndex        =   29
         Top             =   450
         Width           =   5160
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1530
         MouseIcon       =   "frmVtasAlbPdtesFact.frx":0097
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   3675
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1545
         MouseIcon       =   "frmVtasAlbPdtesFact.frx":01E9
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   3300
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
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
         Left            =   585
         TabIndex        =   25
         Top             =   3060
         Width           =   630
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   945
         TabIndex        =   24
         Top             =   3675
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   945
         TabIndex        =   23
         Top             =   3300
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Albarán"
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
         Left            =   600
         TabIndex        =   20
         Top             =   1035
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   960
         TabIndex        =   19
         Top             =   1335
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   960
         TabIndex        =   18
         Top             =   1695
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1530
         Picture         =   "frmVtasAlbPdtesFact.frx":033B
         ToolTipText     =   "Buscar fecha"
         Top             =   1335
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1530
         Picture         =   "frmVtasAlbPdtesFact.frx":03C6
         ToolTipText     =   "Buscar fecha"
         Top             =   1695
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   960
         TabIndex        =   17
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   960
         TabIndex        =   16
         Top             =   2655
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
         Left            =   600
         TabIndex        =   15
         Top             =   2040
         Width           =   495
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1560
         MouseIcon       =   "frmVtasAlbPdtesFact.frx":0451
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1545
         MouseIcon       =   "frmVtasAlbPdtesFact.frx":05A3
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   2655
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmVtasAlbPdtesFact"
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
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim codigo As String 'Código para FormulaSelection de Crystal Report
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
Dim cadTabla As String, cOrden As String
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String

InicializarVbles
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
    numParam = numParam + 1
    
    cadParam = cadParam & "pUsu=" & vUsu.codigo & "|"
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
        codigo = "{" & tabla & ".codclien}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCliente= """) Then Exit Sub
    End If
    
    'D/H Variedad
    cDesde = Trim(txtCodigo(4).Text)
    cHasta = Trim(txtCodigo(5).Text)
    nDesde = txtNombre(4).Text
    nHasta = txtNombre(5).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        codigo = "{albaran_variedad.codvarie}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad= """) Then Exit Sub
    End If
    
    'D/H Fecha albaran
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        codigo = "{" & tabla & ".fechaalb}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    
'   cadTABLA = tabla & " INNER JOIN albaran_variedad ON albaran.numalbar = albaran_variedad.numalbar "
    
    cadFormula = "{tmpinformes.codusu} = {@pUsu}"
    
    If Option1(0).Value Then
        cadParam = cadParam & "pGrupo={albaran_variedad.codvarie}|"
        cadParam = cadParam & "pOrden=0|"
    Else
        cadParam = cadParam & "pGrupo={albaran.codclien}|"
        cadParam = cadParam & "pOrden=1|"
    End If
    numParam = numParam + 2
    
    '[Monica]09/07/2012: solo los clientes que tengan ese tipo de movimiento, si 0 es que son todos los tipos de movimiento
    If Combo1(1).ListIndex <> 0 Then
        If Not AnyadirAFormula(cadselect, "{albaran.codclien} in (select codclien from clientes where codtipalb = " & DBSet(Combo1(1).Text, "T") & ")") Then Exit Sub
    End If
      
    '[Monica]25/07/2012: solo los albaranes que estan marcados para facturar
    If Not AnyadirAFormula(cadselect, "{albaran_variedad.sefactura} = 1") Then Exit Sub
    
    
    '[Monica]03/04/2013: dependiendo del tipo de entrada seleccionamos las variedades
    Select Case Combo1(0).ListIndex
        Case 0 ' todas las variedades
        
        Case 1 ' cooperativa
            If Not AnyadirAFormula(cadselect, "{albaran_variedad.codvarie} in (select codvarie from variedades where tipovariedad = 0)") Then Exit Sub
        Case 2 ' ajenas
            If Not AnyadirAFormula(cadselect, "{albaran_variedad.codvarie} in (select codvarie from variedades where tipovariedad = 1)") Then Exit Sub
    End Select
    
    '[Monica]11/03/2014: dependiendo del tipo de precio seleccionamos las variedades
    Select Case Combo1(2).ListIndex
        Case 0 ' todas los precios
        
        Case 1 ' provisional
            If Not AnyadirAFormula(cadselect, "not isnull({albaran_variedad.preciopro}) and {albaran_variedad.preciopro} <> 0  and (isnull({albaran_variedad.preciodef}) or {albaran_variedad.preciodef} = 0)") Then Exit Sub
        Case 2 ' denitivo
            If Not AnyadirAFormula(cadselect, "not isnull({albaran_variedad.preciodef}) and {albaran_variedad.preciodef} <> 0 ") Then Exit Sub
    End Select
    
    
    If ProcesarCambios(cadselect) Then
          'Nombre fichero .rpt a Imprimir
          cadTitulo = "Albaranes Pendientes de Facturar"
          
          If txtCodigo(6).Text = "" Then
              cadNombreRPT = "rAlbPdtesFact.rpt"
          Else
              cadNombreRPT = "rAlbPdtesFact1.rpt"
          End If
      LlamarImprimir
      'AbrirVisReport
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
Dim i As Integer

    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(2)
        
        For i = 0 To Combo1(1).ListCount - 1
            If Combo1(1).List(i) = vParamAplic.CodTipomAlb Then
                Combo1(1).ListIndex = i
                Exit For
            End If
        Next i
        '[Monica]03/04/2013: el tipo de variedad por defecto es todas
        Combo1(0).ListIndex = 0
        
        '[Monica]11/03/2014: el tipo de precio por defecto es todos
        Combo1(2).ListIndex = 0
        
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
    Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Me.imgBuscar(1).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Me.imgBuscar(4).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Me.imgBuscar(5).Picture = frmPpal.imgListImages16.ListImages(1).Picture

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, H, W
    indFrame = 5
    tabla = "albaran"
    
    CargaCombo
    
    Me.Option1(0).Value = True
    
    imgAyuda(0).Picture = frmPpal.ImageListB.ListImages(10).Picture
    
    
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

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Variedades
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Agencias de transporte
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           ' "____________________________________________________________"
            vCadena = "Si indicamos una fecha, sacará un informe en el que se detalla  " & vbCrLf & _
                      "por variedad que kilos e importe quedan por facturar y que kilos" & vbCrLf & _
                      "importe se han facturado a partir de la fecha de factura que se" & vbCrLf & _
                      "indica. Sacaría un informe de periodificación." & vbCrLf & vbCrLf & _
                      "En caso contrario saca el informe de albaranes pendientes de " & _
                      "facturar. "
                      
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"
    
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
        
        Case 2, 3 'AGENCIAS DE TRANSPORTE
            AbrirFrmAgencias (Index)
        
        Case 4, 5 'VARIEDADES
            AbrirFrmVariedades (Index)
        
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
            Case 4: KEYBusqueda KeyAscii, 4 'variedad desde
            Case 5: KEYBusqueda KeyAscii, 5 'variedad hasta
            Case 6: KEYBusqueda KeyAscii, 4 'agencia desde
            Case 7: KEYBusqueda KeyAscii, 5 'agencia hasta
            Case 2: KEYFecha KeyAscii, 2 'fecha desde
            Case 3: KEYFecha KeyAscii, 3 'fecha hasta
            Case 6: KEYFecha KeyAscii, 6 'desde fecha factura
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
        
        Case 2, 3, 6 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 4, 5 'VARIEDAD
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "00")
            
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 7590
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
        If Not AnyadirAFormula(cadselect, devuelve) Then Exit Function
    Else
        devuelve2 = CadenaDesdeHastaBD(codD, codH, codigo, TipCod)
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

Private Sub AbrirFrmVariedades(indice As Integer)
    indCodigo = indice
    Set frmVar = New frmManVariedad
    frmVar.DatosADevolverBusqueda = "0|1|"
    frmVar.DeConsulta = True
    frmVar.CodigoActual = txtCodigo(indCodigo)
    frmVar.Show vbModal
    Set frmVar = Nothing
End Sub

Private Sub AbrirFrmAgencias(indice As Integer)
    indCodigo = indice + 4
    Set frmTra = New frmManAgencias
    frmTra.DatosADevolverBusqueda = "0|1|"
    frmTra.Show vbModal
    Set frmTra = Nothing
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

Private Function ProcesarCambios(cadwhere As String) As Boolean
Dim Sql As String
Dim Sql1 As String
Dim i As Integer
Dim HayReg As Integer
Dim b As Boolean
Dim Rs As ADODB.Recordset
Dim Importe As Currency
Dim Kilos As Currency
Dim Nregs As Integer

On Error GoTo eProcesarCambios

    HayReg = 0
    
    conn.Execute "delete from tmpinformes where codusu = " & DBSet(vUsu.codigo, "N")
        
    If cadwhere <> "" Then
        cadwhere = QuitarCaracterACadena(cadwhere, "{")
        cadwhere = QuitarCaracterACadena(cadwhere, "}")
        cadwhere = QuitarCaracterACadena(cadwhere, "_1")
    End If
        
    Sql = "insert into tmpinformes (codusu, codigo1, campo1, campo2, importe2) select " & DBSet(vUsu.codigo, "N")
    Sql = Sql & ", albaran.numalbar, albaran_variedad.numlinea, 0, albaran_variedad.pesoneto from albaran inner join albaran_variedad on albaran.numalbar = albaran_variedad.numalbar   "
    Sql = Sql & " where (albaran.numalbar,albaran_variedad.numlinea) not in (select facturas_variedad.numalbar, facturas_variedad.numlinealbar from facturas_variedad) "
    '[Monica]16/05/2012: no incluimos los albaranes de venta a socio
    Sql = Sql & " and albaran.codsocio is null "
    
    If cadwhere <> "" Then Sql = Sql & " and  " & cadwhere
    
    conn.Execute Sql
    
    '[Monica]07/02/2013: si hay fecha desde de factura añadimos los albaranes que se hayan facturado posteriormente a la fecha indicada
    If txtCodigo(6).Text <> "" Then
        Sql = "insert into tmpinformes (codusu, codigo1, campo1, campo2) select " & DBSet(vUsu.codigo, "N")
        Sql = Sql & ", albaran.numalbar, albaran_variedad.numlinea, 1 from albaran inner join albaran_variedad on albaran.numalbar = albaran_variedad.numalbar   "
        Sql = Sql & " where (albaran.numalbar,albaran_variedad.numlinea) in (select facturas_variedad.numalbar, facturas_variedad.numlinealbar from facturas_variedad where fecfactu >= " & DBSet(txtCodigo(6).Text, "F") & ") "
        '[Monica]16/05/2012: no incluimos los albaranes de venta a socio
        Sql = Sql & " and albaran.codsocio is null "
        
        If cadwhere <> "" Then Sql = Sql & " and  " & cadwhere
    
        conn.Execute Sql
        
        
        Sql = "select count(*) from tmpinformes where codusu = " & vUsu.codigo
        Nregs = TotalRegistros(Sql)
        Me.pb1.visible = True
        
        CargarProgresNew Me.pb1, Nregs
        DoEvents
        
        ' cargamos el importe de los albaranes facturados posteriormente a la fecha
        Sql = "select codigo1, campo1 from tmpinformes where codusu = " & vUsu.codigo & " and campo2 = 1"
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs.EOF
        
            IncrementarProgresNew Me.pb1, 1
            DoEvents
            
            Sql1 = "select sum(impornet) from facturas_variedad where numalbar = " & DBSet(Rs!Codigo1, "N") & " and numlinealbar = " & DBSet(Rs!campo1, "N")
            Sql1 = Sql1 & " and facturas_variedad.fecfactu >= " & DBSet(txtCodigo(6).Text, "F")
            Importe = DevuelveValor(Sql1)
            
            Sql1 = "select pesoneto from albaran_variedad where numalbar = " & DBSet(Rs!Codigo1, "N") & " and numlinea  = " & DBSet(Rs!campo1, "N")
            Kilos = DevuelveValor(Sql1)
            
            Sql1 = "update tmpinformes set importe1 = " & DBSet(Importe, "N")
            Sql1 = Sql1 & ", importe2 = " & DBSet(Kilos, "N")
            Sql1 = Sql1 & " where codusu = " & vUsu.codigo
            Sql1 = Sql1 & " and codigo1 = " & DBSet(Rs!Codigo1, "N")
            Sql1 = Sql1 & " and campo1 = " & DBSet(Rs!campo1, "N")
            
            conn.Execute Sql1
            
            Rs.MoveNext
        Wend
        Set Rs = Nothing
        
        ' cargamos el importe de los albaranes que estan pendientes de facturar con el precio provisional o definitivo
        Sql = "select codigo1, campo1 from tmpinformes where codusu = " & vUsu.codigo & " and campo2 = 0"
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs.EOF
            Sql1 = "select round((if(preciodef is null or preciodef = 0, if(preciopro is null or preciopro = 0, 0, preciopro),preciodef)) * pesoneto, 2) from albaran_variedad where numalbar = " & DBSet(Rs!Codigo1, "N") & " and numlinea= " & DBSet(Rs!campo1, "N")
            Importe = DevuelveValor(Sql1)
            
            IncrementarProgresNew Me.pb1, 1
            DoEvents
    
            
            Sql1 = "update tmpinformes set importe1 = " & DBSet(Importe, "N")
            Sql1 = Sql1 & " where codusu = " & vUsu.codigo
            Sql1 = Sql1 & " and codigo1 = " & DBSet(Rs!Codigo1, "N")
            Sql1 = Sql1 & " and campo1 = " & DBSet(Rs!campo1, "N")
            
            conn.Execute Sql1
            
            Rs.MoveNext
        Wend
        Set Rs = Nothing
        
        
    End If
        
    Me.pb1.visible = False
        
    ProcesarCambios = HayRegistros("tmpinformes", "codusu = " & vUsu.codigo)

eProcesarCambios:
    Me.pb1.visible = False
    If Err.Number <> 0 Then
        ProcesarCambios = False
    End If
End Function


Private Sub InsertaLineaEnTemporal(ByRef ItmX As ListItem)
Dim Sql As String
Dim Codmacta As String
Dim Rs As ADODB.Recordset
Dim Sql1 As String

        Sql1 = "insert into tmpinformes(codusu, codigo1) values ("
        Sql1 = Sql1 & DBSet(vUsu.codigo, "N") & "," & DBSet(ItmX.Text, "N") & ")"

        conn.Execute Sql1
    
End Sub

' ********* si n'hi han combos a la capçalera ************
Private Sub CargaCombo()
Dim i As Integer
Dim Sql As String
Dim Rs As ADODB.Recordset


    Combo1(0).Clear

    Combo1(0).AddItem "Todas"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Cooperativa"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "Ajena"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    Combo1(1).Clear
    
    Sql = "select distinct codtipalb from clientes order by 1 "
    
    Combo1(1).AddItem "Todos"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    i = 1
    While Not Rs.EOF
        Combo1(1).AddItem Rs!codtipalb
        Combo1(1).ItemData(Combo1(1).NewIndex) = i
        i = i + 1
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    ' Tipo de precio
    Combo1(2).Clear

    Combo1(2).AddItem "Todos"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 0
    Combo1(2).AddItem "Provisional"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 1
    Combo1(2).AddItem "Definitivo"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 2
    
End Sub



