VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmVtasPaletsPrev 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gestión de Palets"
   ClientHeight    =   11070
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   17460
   Icon            =   "frmVtasPaletsPrev.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11070
   ScaleWidth      =   17460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   12330
      TabIndex        =   27
      Top             =   945
      Width           =   4965
      Begin VB.TextBox Text1 
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
         Left            =   135
         MaxLength       =   15
         TabIndex        =   30
         Text            =   "Text3"
         Top             =   450
         Width           =   1530
      End
      Begin VB.TextBox Text1 
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
         MaxLength       =   15
         TabIndex        =   29
         Text            =   "Text3"
         Top             =   450
         Width           =   1530
      End
      Begin VB.TextBox Text1 
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
         Index           =   7
         Left            =   3330
         MaxLength       =   15
         TabIndex        =   28
         Text            =   "Text3"
         Top             =   450
         Width           =   1530
      End
      Begin VB.Label Label1 
         Caption         =   "Cajas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   14
         Left            =   135
         TabIndex        =   33
         Top             =   180
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "Peso Bruto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   1710
         TabIndex        =   32
         Top             =   180
         Width           =   1500
      End
      Begin VB.Label Label1 
         Caption         =   "Peso Neto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   3330
         TabIndex        =   31
         Top             =   180
         Width           =   1545
      End
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   4860
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   405
      Width           =   2070
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   7020
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   405
      Width           =   2070
   End
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3735
      TabIndex        =   19
      Top             =   90
      Width           =   795
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   210
         TabIndex        =   20
         Top             =   180
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Informe Palets en Cámaras"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   90
      TabIndex        =   17
      Top             =   90
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   18
         Top             =   180
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Nuevo"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Eliminar"
               Object.Tag             =   "2"
               Object.Width           =   1e-4
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Buscar"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver Todos"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Imprimir"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Salir"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1110
      Left            =   135
      TabIndex        =   9
      Top             =   945
      Width           =   12135
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
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
         Left            =   180
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Nº Palet|N|S|||palets|numpalet|0000000|S|"
         Text            =   "Text1 7"
         Top             =   450
         Width           =   980
      End
      Begin VB.TextBox Text1 
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
         Left            =   8055
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "Cod.Forfait|T|S|||palets_variedad|codforfait|||"
         Text            =   "Text1"
         Top             =   450
         Width           =   960
      End
      Begin VB.TextBox Text2 
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
         Index           =   4
         Left            =   9045
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   15
         Text            =   "Text2"
         Top             =   450
         Width           =   2910
      End
      Begin VB.TextBox Text1 
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
         Left            =   1350
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Confeccion|F|N|||palets|fechaconf|dd/mm/yyyy||"
         Top             =   450
         Width           =   1425
      End
      Begin VB.TextBox Text1 
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
         Left            =   4185
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Variedad|N|N|||palets_variedad|codvarie|000000||"
         Top             =   450
         Width           =   870
      End
      Begin VB.TextBox Text1 
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
         Left            =   3060
         TabIndex        =   2
         Tag             =   "Nº Pedido|N|S|||palets|numpedid|000000||"
         Text            =   "Text3"
         Top             =   450
         Width           =   990
      End
      Begin VB.TextBox Text2 
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
         Index           =   3
         Left            =   5085
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   11
         Text            =   "Text2"
         Top             =   450
         Width           =   2910
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   9090
         ToolTipText     =   "Buscar Cámara"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Forfait"
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
         Index           =   19
         Left            =   8055
         TabIndex        =   16
         Top             =   180
         Width           =   810
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   3795
         ToolTipText     =   "Buscar Pedidos sin albarán"
         Top             =   210
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   2700
         Picture         =   "frmVtasPaletsPrev.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F.Confección"
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
         Index           =   13
         Left            =   1350
         TabIndex        =   14
         Top             =   180
         Width           =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "Variedad"
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
         Index           =   29
         Left            =   4185
         TabIndex        =   13
         Top             =   180
         Width           =   870
      End
      Begin VB.Label Label1 
         Caption         =   "Pedido"
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
         Index           =   6
         Left            =   3060
         TabIndex        =   12
         Top             =   180
         Width           =   675
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   5130
         ToolTipText     =   "Buscar Palet"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Palet"
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
         Index           =   28
         Left            =   225
         TabIndex        =   10
         Top             =   180
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   525
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   10395
      Width           =   2175
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   8
         Top             =   180
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
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
      Left            =   16200
      TabIndex        =   6
      Top             =   10395
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
      Left            =   15030
      TabIndex        =   5
      Top             =   10410
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc Data3 
      Height          =   330
      Left            =   3000
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   16785
      TabIndex        =   21
      Top             =   180
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ayuda"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lw1 
      Height          =   7995
      Left            =   135
      TabIndex        =   26
      Top             =   2115
      Width           =   17115
      _ExtentX        =   30189
      _ExtentY        =   14102
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Cargando datos...."
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
      Index           =   0
      Left            =   2520
      TabIndex        =   34
      Top             =   10575
      Visible         =   0   'False
      Width           =   4545
   End
   Begin VB.Label Label1 
      Caption         =   "Estado"
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
      Index           =   5
      Left            =   4860
      TabIndex        =   25
      Top             =   135
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Cámaras"
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
      Index           =   27
      Left            =   7020
      TabIndex        =   23
      Top             =   135
      Width           =   1515
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver Todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         HelpContextID   =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         HelpContextID   =   2
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnLineas 
         Caption         =   "&Lineas"
         Enabled         =   0   'False
         HelpContextID   =   2
         Shortcut        =   ^L
         Visible         =   0   'False
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnInfCamaras 
         Caption         =   "Informe Palets en Cámaras"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmVtasPaletsPrev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'Si se llama de la busqueda en el frmAlmMovimArticulos se accede
'a las tablas del Albaran o de Facturas de movimiento seleccionado (solo consulta)
Public hcoCodMovim As String 'cod. movim
Public hcoCodTipoM As String 'Codigo detalle de Movimiento(ALC)
Public hcoFechaMov As String 'fecha del movimiento

'========== VBLES PRIVADAS ====================
Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmLPal As frmVtasLinPalets 'Lineas de variedades de palets
Attribute frmLPal.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmMen As frmMensajes 'Pedidos que no tienen asociado un nro de albaran
Attribute frmMen.VB_VarHelpID = -1
Private WithEvents frmVar As frmManVariedad 'variedades
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmFor As frmManForfaits 'forfaits
Attribute frmFor.VB_VarHelpID = -1

Private WithEvents frmMPal As frmManPaleConf 'Form Mto de Palets de confeccion
Attribute frmMPal.VB_VarHelpID = -1
Private WithEvents frmMCam As frmManCamara 'Form Mto de Camaras
Attribute frmMCam.VB_VarHelpID = -1
Private WithEvents frmBas  As frmBasico ' Lineas de confeccion
Attribute frmBas.VB_VarHelpID = -1
Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Mantenimiento Lineas
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim CodTipoMov As String
'Codigo tipo de movimiento en función del valor en la tabla de parámetros: stipom

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean

Dim EsCabecera As Boolean
'Para saber en MandaBusquedaPrevia si busca en la tabla scapla o en la tabla sdirec


Dim EsDeVarios As Boolean
'Si el cliente mostrado es de Varios o No

'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1


Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos
Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal
Dim indice As Byte

Dim nomColumna As String
Dim columna As Integer
Dim Orden As Integer

Dim CadB As String
Dim FiltroCamara As String

Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim i As Integer

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda

    End Select
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1 'Busqueda
            CargaListview nomColumna, CadB, False

    End Select
End Sub


Private Sub BotonBuscar()
Dim anc As Single

    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        lw1.ListItems.Clear
        
        PonerModo 1
        
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbLightBlue 'vbYellow
    End If
End Sub


Private Sub BotonVerTodos()
    LimpiarCampos
    
    nomColumna = "fechaconf"
    columna = 2
    Orden = 1
    
    CargaListview nomColumna, "", False
    
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Facturas (scafac)
' y los registros correspondientes de las tablas cab. albaranes (scafac1)
' y las lineas de la factura (slifac)
Dim Cad As String
Dim NroAlbar As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If lw1.SelectedItem Is Nothing Then Exit Sub
    
    NroAlbar = NroAlbaranAsignado(lw1.SelectedItem.Text, 0)
    If NroAlbar <> "" Then
        Cad = "El pedido asociado a este palet se encuentra asignado al albarán " & NroAlbar & "." & vbCrLf
        Cad = Cad & "         ¿ Desea continuar ?"
        If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Sub
        End If
    End If

    Cad = "Cabecera de Palets." & vbCrLf
    Cad = Cad & "-------------------------------------      " & vbCrLf & vbCrLf
    Cad = Cad & "Va a eliminar el Palet:            "
    Cad = Cad & vbCrLf & "Nº Palet:  " & Format(lw1.SelectedItem.Text, "0000000")
    Cad = Cad & vbCrLf & "Fecha:  " & Format(lw1.SelectedItem.SubItems(1), "dd/mm/yyyy")
    Cad = Cad & vbCrLf & vbCrLf & " ¿Desea Eliminarlo? "

    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = lw1.SelectedItem
        
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            CargaListview nomColumna, CadB, False
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
EEliminar:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminar Albaran", Err.Description
End Sub


Private Sub Combo1_Change(Index As Integer)
    Select Case Index
        Case 0
            If Combo1(1).ListIndex > 0 Then
                cadfiltro1 = "palets.codcamara = " & Combo1(1).ListIndex
            Else
                cadfiltro1 = ""
            End If
    End Select
    
    CargaListview columna1, CadB, False
    
   
End Sub

Private Sub Combo1_Click(Index As Integer)
    Select Case Index
        Case 0
            If Combo1(0).ListIndex > 0 Then
                FiltroCamara = "palets.codcamara = " & Combo1(0).ListIndex
            Else
                FiltroCamara = ""
            End If
    End Select
    
    If Not PrimeraVez Then
        CargaListview nomColumna, CadB, False
    End If
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
    If PrimeraVez Then
        Combo1(0).ListIndex = 0
        Combo1(1).ListIndex = 0
        FiltroCamara = ""
        BotonVerTodos
        PrimeraVez = False
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
Dim i As Integer

    PrimeraVez = True

    'Icono del formulario
'    Me.Icon = frmPpal.Icon
    
     'Icono de busqueda
    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next kCampo

    ' ICONITOS DE LA BARRA
    btnPrimero = 16
    
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'l'1 i el 2 son separadors
        .Buttons(5).Image = 1   'Buscar
        .Buttons(6).Image = 2   'Totss
        'el 5 i el 6 son separadors
        .Buttons(1).Image = 3   'Insertar
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Borrar
        'el 10  son separadors
        .Buttons(8).Image = 10  'Imprimir
    End With
    
    With Me.Toolbar5
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26 'Busqueda Avanzada
    End With
    
    ' La Ayuda
'    With Me.ToolbarAyuda
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 12
'    End With
    
    
    LimpiarCampos   'Limpia los campos TextBox
    CargaCombo
    
    CodTipoMov = "PAL" 'hcoCodTipoM
    VieneDeBuscar = False
    
    '## A mano
    NombreTabla = "palets"
    NomTablaLineas = "palets_variedad" 'Tabla lineas de variedades
    Ordenacion = " ORDER BY palets.numpalet"
    
    CargarColumnas
    
    CargaCombo
    
End Sub


Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim i As Byte
Dim miRsAux As ADODB.Recordset
    
    
    For i = 0 To Combo1.Count - 1
        Combo1(i).Clear
    Next i
    
    Set miRsAux = New ADODB.Recordset
    
    ' camaras
    Sql = "Select * from camaras order by codcamara"
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    Combo1(0).AddItem "Todas"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    
    While Not miRsAux.EOF
        Combo1(0).AddItem miRsAux!nomcamara
        Combo1(0).ItemData(Combo1(0).NewIndex) = miRsAux!Codcamara
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    
    ' estado
    Combo1(1).AddItem "Todos"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    
    Combo1(1).AddItem "Sin Asignar"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    
    Combo1(1).AddItem "En Pedido"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 2
    
    Combo1(1).AddItem "Salido"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 3
    
End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Modo = 4 Then TerminaBloquear
End Sub


Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Clientes
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1)  'Cod Clien
End Sub

Private Sub frmBas_DatoSeleccionado(CadenaSeleccion As String)
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) ' codigo de linea de confeccion
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    Text1(indice).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Variedades
    If CadenaSeleccion <> "" Then
        Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codvariedad
        Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
    End If
End Sub

Private Sub frmFor_DatoSeleccionado(CadenaSeleccion As String)
'Forfaits
    If CadenaSeleccion <> "" Then
        Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codforfait
        Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
    End If
End Sub

Private Sub frmMCam_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Palet
        Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Palets
    End If
End Sub

Private Sub frmMen_DatoSeleccionado(CadenaSeleccion As String)
    Text1(5).Text = CadenaSeleccion
End Sub

Private Sub frmMPal_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Palets de confecciones
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Palet
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Palets
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 3) 'Peso Palet confeccion
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(indice).Text = vCampo
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim Cad As String

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Variedad
            indice = 3
            Set frmVar = New frmManVariedad
            frmVar.DatosADevolverBusqueda = "0|1|"
            frmVar.CodigoActual = Text1(indice).Text
            frmVar.Show vbModal
            Set frmVar = Nothing
            PonerFoco Text1(indice)
            
        Case 1 'Ayuda de pedidos que no tengan asignado nro de albaran
            'mostramos los palets asociados al pedido
            Set frmMen = New frmMensajes
            
            Cad = "select * from pedidos, clientes, destinos where numalbar is null "
            Cad = Cad & " and pedidos.codclien = clientes.codclien and "
            Cad = Cad & " pedidos.codclien = destinos.codclien and pedidos.coddesti = destinos.coddesti"
            
            frmMen.cadwhere = Cad
            
            frmMen.OpcionMensaje = 20 'Pedidos que no tienen asociados un nro de albaran
            frmMen.Show vbModal
            Set frmMen = Nothing
            
        Case 2 ' forfait
            indice = 4
            Set frmFor = New frmManForfaits
            frmFor.DatosADevolverBusqueda = "0|1|"
            frmFor.CodigoActual = Text1(4).Text
            frmFor.Show vbModal
            Set frmFor = Nothing
            PonerFoco Text1(4)
        
    End Select
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub imgFec_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
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
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40

    imgFec(0).Tag = Index '<===
    Select Case Index
        Case 0
            indice = Index + 1
    End Select
    ' *** repasar si el camp es txtAux o Text1 ***
    If Text1(indice).Text <> "" Then frmC.NovaData = Text1(indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco Text1(indice) '<===
    ' ********************************************
End Sub

Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        indice = 7
        frmZ.pTitulo = "Observaciones del Palet"
        frmZ.pValor = Text1(indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco Text1(indice)
    End If
End Sub




Private Sub lw1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim campo2 As Integer


    Select Case ColumnHeader
        Case "Palet", "Palet v"
            nomColumna = "numpalet"
            campo2 = 1
        Case "Fecha Conf.", "Fecha Conf.v"
            nomColumna = "fechaconf"
            campo2 = 2
        Case "Pedido", "Pedido v"
            nomColumna = "numpedid"
            campo2 = 3
        Case "Variedad", "Variedad v"
            nomColumna = "codvarie"
            campo2 = 4
        Case "Nombre Variedad", "Nombre Variedad v"
            nomColumna = "nomvarie"
            campo2 = 5
        Case "Forfait", "Forfait v"
            nomColumna = "codforfait"
            campo2 = 6
        Case "Nombre Forfait", "Nombre Forfait v"
            nomColumna = "nomconfe"
            campo2 = 7
        Case "Cajas", "Cajas v"
            nomColumna = "numcajas"
            campo2 = 8
        Case "Peso Bruto", "Peso Bruto v"
            nomColumna = "pesobrut"
            campo2 = 9
        Case "Peso Neto", "Peso Neto v"
            nomColumna = "pesoneto"
            campo2 = 10
    End Select
    
    If campo2 = columna Then
        If Orden = lvwAscending Then
            'nomColumna = nomColumna & " DESC"
            Orden = lvwDescending
        Else
            Orden = lvwAscending
        End If
    Else
        columna = campo2
    End If
    
    CargaListview nomColumna, CadB, True
 
End Sub



Private Sub lw1_ItemClick(ByVal item As MSComctlLib.ListItem)
    lblIndicador.Caption = PonerContRegistrosLw(lw1, item)
End Sub

Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Pedido
'         BotonEliminarLinea
    Else   'Eliminar Pedido
         BotonEliminar
    End If
End Sub


Private Sub mnImprimir_Click()
'Imprimir Factura
    
    If lw1.SelectedItem = 0 Then Exit Sub
    
    BotonImprimir
End Sub

Private Sub mnBusquedaAvanzada_Click()
Dim frmPal As frmVtasPalets
    
    Set frmPal = New frmVtasPalets
    
    frmPal.pModo = 1
    frmPal.Show vbModal
    
    Set frmPal = Nothing
End Sub

Private Sub mnNuevo_Click()
Dim frmPal As frmVtasPalets
    
    Set frmPal = New frmVtasPalets
    
    frmPal.pModo = 3
    frmPal.Show vbModal
    
    Set frmPal = Nothing
    
    CargaListview nomColumna, CadB, False
    
End Sub

Private Sub mnLineas_Click()
    BotonMtoLineas 1, "Facturas"
End Sub


Private Sub mnModificar_Click()
Dim frmPal As frmVtasPalets
    
    Set frmPal = New frmVtasPalets
    
    frmPal.pModo = 4
    frmPal.DatosADevolverBusqueda = lw1.SelectedItem.Text
    frmPal.Show vbModal
    
    Set frmPal = Nothing

    CargaListview nomColumna, CadB, False


End Sub


Private Function BloqueaAlbxFac() As Boolean
'bloquea todos los albaranes de la factura
Dim Sql As String

    On Error GoTo EBloqueaAlb
    
    BloqueaAlbxFac = False
    'bloquear cabecera albaranes x factura
    Sql = "select * FROM scafac1 "
    Sql = Sql & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute Sql, , adCmdText
    BloqueaAlbxFac = True

EBloqueaAlb:
    If Err.Number <> 0 Then BloqueaAlbxFac = False
End Function


Private Function BloqueaLineasFac() As Boolean
'bloquea todas las lineas de la factura
Dim Sql As String

    On Error GoTo EBloqueaLin

    BloqueaLineasFac = False
    'bloquear cabecera albaranes x factura
    Sql = "select * FROM slifac "
    Sql = Sql & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute Sql, , adCmdText
    BloqueaLineasFac = True

EBloqueaLin:
    If Err.Number <> 0 Then BloqueaLineasFac = False
End Function


Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    If (Modo = 5) Then 'Modo 5: Mto Lineas
        '1:Insertar linea, 2: Modificar
        If ModificaLineas = 1 Or ModificaLineas = 2 Then cmdCancelar_Click
        cmdRegresar_Click
        Exit Sub
    End If
    Unload Me
End Sub


Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub


Private Sub Text1_Change(Index As Integer)
    If Index = 9 Then HaCambiadoCP = True 'Cod. Postal
End Sub

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    If Index = 9 Then HaCambiadoCP = False 'CPostal
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 7 Or (Index = 7 And Text1(7).Text = "") Then KEYpress KeyAscii
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)
Dim devuelve As String
Dim cadMen As String
Dim Sql As String
        
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 0 'numero de palet
            PonerFormatoEntero Text1(Index)
        
        Case 1 'Fecha de confeccion
            If Text1(Index).Text <> "" Then
                PonerFormatoFecha Text1(Index), True
            End If
                
        
        Case 3 'Variedad
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = DevuelveDesdeBDNew(cAgro, "variedades", "nomvarie", "codvarie", Text1(Index).Text, "N")
            Else
                Text2(Index).Text = ""
            End If
                
        Case 4 'Forfait
            If Text1(Index).Text <> "" Then
                Text2(Index) = PonerNombreDeCod(Text1(Index), "forfaits", "nomconfe")
            Else
                Text2(Index).Text = ""
            End If
        
    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadAux As String
    
    
    CadB = ObtenerBusqueda(Me) ' antes obtenerbusqueda3(me,false)
        
    nomColumna = "fechaconf"
    columna = 11 ' 1
    Orden = 1
    
    CargaListview nomColumna, CadB, False

End Sub


Private Sub PonerCampos()
Dim BrutoFac As Single

    On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 2, "Frame2"
    Text1(9).Text = Mid(Text1(8).Text, 12, 8)
    Text1(10).Text = Mid(Text1(6).Text, 12, 8)
    Text1(12).Text = Mid(Text1(14).Text, 12, 8)
    Text1(11).Text = Mid(Text1(15).Text, 12, 8)
    
    Text2(3).Text = DevuelveDesdeBDNew(cAgro, "confpale", "pesopale", "codpalet", Text1(4).Text, "N")
    If Text2(3).Text <> "" Then PonerFormatoDecimal Text2(3), 4
    
'    FormatoDatosTotales
    
    'poner descripcion campos
    Modo = 4
    Text2(4) = PonerNombreDeCod(Text1(4), "confpale", "nompalet", "codpalet", "N") 'palet de confeccion
    Text2(20) = PonerNombreDeCod(Text1(20), "camaras", "nomcamara", "codcamara", "N") 'camara
    Modo = 2
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Ofertas
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario

    If Err.Number <> 0 Then Err.Clear
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte, Numreg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo

    'Actualiza Iconos Insertar,Modificar,Eliminar
    '## No tiene el boton modificar y no utiliza la funcion general
'    ActualizarToolbar Modo, Kmodo
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    If Modo = 2 Then
        lblIndicador.Caption = PonerContRegistrosLw(lw1, lw1.SelectedItem)
    End If
    
    
    ' el frame de busqueda está activo unicamente en busqueda normal
    Frame2.Enabled = (Modo = 1)
    
    '---------------------------------------------
    b = (Modo = 1)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
       
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub DesplazamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim b As Boolean
Dim Sql As String

    On Error GoTo EDatosOK

    DatosOk = False
    
'    ComprobarDatosTotales

    'concatenamos en el text1(6) y text1(8) la fechahora
    Text1(8).Text = Format(Text1(2).Text, "dd/mm/yyyy") & " " & Format(Text1(9).Text, "HH:MM:SS")
    If Text1(3).Text <> "" And Text1(10).Text <> "" Then
        Text1(6).Text = Format(Text1(3).Text, "dd/mm/yyyy") & " " & Format(Text1(10).Text, "HH:MM:SS")
    Else
        Text1(6).Text = ""
    End If
    
    If Text1(13).Text <> "" And Text1(12).Text <> "" Then
        Text1(14).Text = Format(Text1(13).Text, "dd/mm/yyyy") & " " & Format(Text1(12).Text, "HH:MM:SS")
    Else
        Text1(14).Text = ""
    End If
    
    If Text1(13).Text <> "" And Text1(11).Text <> "" Then
        Text1(15).Text = Format(Text1(13).Text, "dd/mm/yyyy") & " " & Format(Text1(11).Text, "HH:MM:SS")
    Else
        Text1(15).Text = ""
    End If
    
    'comprobamos datos OK de la tabla palets
    b = CompForm2(Me, 2, "Frame2") ' , 1) 'Comprobar formato datos ok de la cabecera: opcion=1
    If Not b Then Exit Function
    
    
    ' comprobamos los rangos de fechas
    If b And Text1(3).Text <> "" Then
        If CDate(Text1(2).Text) > CDate(Text1(3).Text) Then
            MsgBox "La fecha de inicio no puede ser superior a la fecha fin. Revise.", vbExclamation
            b = False
            PonerFoco Text1(9)
        End If
    End If
    
    If b And Text1(6).Text <> "" Then
        If CDate(Text1(8).Text) > CDate(Text1(6).Text) Then
            MsgBox "La hora de inicio no puede ser superior a la de fin. Revise.", vbExclamation
            b = False
            PonerFoco Text1(9)
        End If
    End If
    
    If b And Text1(15).Text <> "" Then
        If CDate(Text1(14).Text) > CDate(Text1(15).Text) Then
            MsgBox "La hora de inicio de confección no puede ser superior a la de fin. Revise.", vbExclamation
            b = False
            PonerFoco Text1(12)
        End If
    End If
    
    
    
    'comprobamos que el numero de pedido existe si no es nulo
    If b And Text1(5).Text <> "" Then
        Sql = ""
        Sql = DevuelveDesdeBDNew(cAgro, "pedidos", "numpedid", "numpedid", Text1(5), "N")
        If Sql = "" Then
            MsgBox "El número de pedido no existe en la tabla de pedidos. Reintroduzca.", vbExclamation
            Text1(5).Text = ""
            b = False
            PonerFoco Text1(5)
        End If
    End If
    
    
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
Dim b As Boolean
Dim i As Byte

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    b = True

    For i = 0 To txtAux.Count - 1
        If i = 4 Or i = 6 Or i = 7 Then
            If txtAux(i).Text = "" Then
                MsgBox "El campo " & txtAux(i).Tag & " no puede ser nulo", vbExclamation
                b = False
                PonerFoco txtAux(i)
                Exit Function
            End If
        End If
    Next i
            
    DatosOkLinea = b
    
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 16 And KeyCode = 40 Then 'campo Amliacion Linea y Flecha hacia abajo
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 16 And KeyAscii = 13 Then 'campo Amliacion Linea y ENTER
        PonerFocoBtn Me.cmdAceptar
    End If
End Sub


Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
Dim NroAlbar As String
Dim Cad As String

    NroAlbar = NroAlbaranAsignado(Data1.Recordset!numpalet, 0)
    If NroAlbar <> "" Then
        Cad = "El pedido asociado a este palet se encuentra asignado al albarán " & NroAlbar & "." & vbCrLf
        Cad = Cad & "                     ¿ Desea continuar ?"
        If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Sub
        End If
    End If

    If BloqueaRegistro(NombreTabla, "numpalet = " & Data1.Recordset!numpalet) Then
'    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
        Select Case Button.Index
            Case 1 'añadir variedad
                Set frmLPal = New frmVtasLinPalets
                
                frmLPal.ModoExt = 3
                frmLPal.Palet = Data1.Recordset.Fields(0).Value
                frmLPal.Show vbModal
            
                Set frmLPal = Nothing
            Case 2 'modificar variedad
                Set frmLPal = New frmVtasLinPalets
                
                frmLPal.ModoExt = 4
                frmLPal.Palet = Data3.Recordset.Fields(0).Value
                frmLPal.Linea = Data3.Recordset.Fields(1).Value
                frmLPal.Show vbModal
                
                Set frmLPal = Nothing
                
            Case 3 ' boton eliminar linea de variedades
                BotonEliminarLinea
            Case Else
        End Select
    End If
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Añadir
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
        Case 5  'Buscar
            mnBuscar_Click
        Case 6  'Todos
            BotonVerTodos
        Case 8 'Imprimir Albaran
            mnImprimir_Click
    End Select
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub





Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
           mnBusquedaAvanzada_Click
    End Select
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

Private Sub TxtAux_Change(Index As Integer)
    If Index = 6 And ModificaLineas = 2 Then 'Precio y Modo Borrar Lineas
        txtAux(5).Text = "M"
    End If
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)

    'Quitar espacios en blanco
    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    Select Case Index
        Case 4 'Precio
             'Tipo 2: Decimal(10,4)
             If txtAux(Index).Text <> "" Then PonerFormatoDecimal txtAux(Index), 2
            
        Case 6, 7 'Descuentos
            PonerFormatoDecimal txtAux(Index), 4 'Tipo 4: Decimal(4,2)
            If Index = 7 Then PonerFoco Me.Text2(16)
            
        Case 8 'Importe Linea
            PonerFormatoDecimal txtAux(Index), 3 'Tipo 3: Decimal(10,2)
    End Select
    
'    If (Index = 3 Or Index = 4 Or Index = 6 Or Index = 7) Then 'Cant., Precio, Dto1, Dto2
'        If txtAux(1).Text = "" Then Exit Sub
'        txtAux(8).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(6).Text, txtAux(7).Text, vParamAplic.TipoDtos)
'        PonerFormatoDecimal txtAux(8), 1
'    End If
End Sub


Private Function Eliminar() As Boolean
Dim Sql As String, LEtra As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
    
    On Error GoTo FinEliminar

    b = False
    If lw1.SelectedItem Is Nothing Then Exit Function
        
    conn.BeginTrans
        

    'Eliminar en tablas de factura de Ariges
    '------------------------------------------
    Sql = " " & ObtenerWhereCP(True)

    'Lineas de calibres (palets_calibre)
    conn.Execute "Delete from palets_calibre " & Sql

    'Lineas de variedades
    conn.Execute "Delete from palets_variedad " & Sql
    
    'Cabecera de palets (palets)
    conn.Execute "Delete from " & NombreTabla & Sql
    
    'Decrementar contador si borramos el ult. palet
    Set vTipoMov = New CTiposMov
    vTipoMov.DevolverContador "PAL", Val(Text1(0).Text)
    Set vTipoMov = Nothing
    
    b = True
    
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Palet", Err.Description
        b = False
    End If
    If Not b Then
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function

Private Function EliminarLinea() As Boolean
Dim Sql As String, LEtra As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
    
    On Error GoTo FinEliminar

    b = False
    If Data3.Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        

    'Eliminar en tablas de paltes_variedad y palets_calibre
    '------------------------------------------
    Sql = " where numpalet = " & Data3.Recordset.Fields(0)
    Sql = Sql & " and numlinea = " & Data3.Recordset.Fields(1)

    'Lineas de calibres (palets_calibre)
    conn.Execute "Delete from palets_calibre " & Sql

    'Lineas de variedades
    conn.Execute "Delete from palets_variedad " & Sql
    
    b = True
    
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Variedad del Palet", Err.Description
        b = False
    End If
    If Not b Then
        conn.RollbackTrans
        EliminarLinea = False
    Else
        conn.CommitTrans
        EliminarLinea = True
    End If
End Function

Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next

    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Function ObtenerWhereCP(conWhere As Boolean) As String
Dim Sql As String

    On Error Resume Next
    
    Sql = " numpalet= " & lw1.SelectedItem
    If conWhere Then Sql = " WHERE " & Sql
    ObtenerWhereCP = Sql
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo cadena WHERE.", Err.Description
End Function


Private Function MontaSQLCarga(enlaza As Boolean, Opcion As Byte) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim Sql As String
    
    If Opcion = 1 Then
        Sql = "SELECT numpalet, numlinea, numline1, palets_calibre.codvarie, palets_calibre.codcalib, nomcalib, numcajas "
        Sql = Sql & " FROM palets_calibre, calibres WHERE palets_calibre.codvarie = calibres.codvarie and "
        Sql = Sql & " palets_calibre.codcalib = calibres.codcalib "
    ElseIf Opcion = 2 Then
        Sql = "SELECT palets_variedad.numpalet, numlinea, palets_variedad.codvarie, a.nomvarie as nomvarie1, palets_variedad.codvarco, "
        Sql = Sql & " b.nomvarie as nomvarie2, palets_variedad.codmarca, marcas.nommarca, palets_variedad.codforfait, forfaits.nomconfe, "
        Sql = Sql & " categori, numcajas, pesobrut, pesoneto "
        Sql = Sql & " FROM palets_variedad, variedades a, variedades b, marcas, forfaits " 'lineas de variedades del palet
        Sql = Sql & " WHERE palets_variedad.codvarie = a.codvarie "
        Sql = Sql & " and palets_variedad.codvarco = b.codvarie"
        Sql = Sql & " and palets_variedad.codmarca = marcas.codmarca "
        Sql = Sql & " and palets_variedad.codforfait = forfaits.codforfait "
    End If
    
    If enlaza Then
        Sql = Sql & " and " & ObtenerWhereCP(False)
        If Opcion = 1 Then Sql = Sql & " AND numlinea=" & Data3.Recordset.Fields!NumLinea
    Else
        Sql = Sql & " and numpalet = -1"
    End If
    Sql = Sql & " ORDER BY numpalet"
    If Opcion = 1 Then Sql = Sql & ", numlinea "
    MontaSQLCarga = Sql
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean, bAux As Boolean
Dim i As Integer

        b = (Modo = 2) Or (Modo = 0) 'Or (Modo = 5 And ModificaLineas = 0)
        'Buscar
        Toolbar1.Buttons(5).Enabled = b
        Me.mnBuscar.Enabled = b
        'Vore Tots
        Toolbar1.Buttons(6).Enabled = b
        Me.mnVerTodos.Enabled = b
        'Añadir
        Toolbar1.Buttons(1).Enabled = b
        Me.mnModificar.Enabled = b
        
        If Not lw1.SelectedItem Is Nothing Then
            b = (Modo = 2 And lw1.SelectedItem <> 0)
        Else
            b = False
        End If
        'Modificar
        Toolbar1.Buttons(2).Enabled = b
        Me.mnModificar.Enabled = b
        'eliminar
        Toolbar1.Buttons(3).Enabled = (Modo = 2)
        Me.mnEliminar.Enabled = (Modo = 2)
            
        b = (Modo = 2)
        'Imprimir
        Toolbar1.Buttons(8).Enabled = b
        Me.mnImprimir.Enabled = b

End Sub




Private Sub BotonImprimir()
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadselect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim Sql As String

    If lw1.SelectedItem = 0 Then
        MsgBox "Debe seleccionar un Palet para Imprimir.", vbInformation
        Exit Sub
    End If
    
    cadFormula = ""
    cadParam = ""
    cadselect = ""
    numParam = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    indRPT = 5 'Impresion de Palet
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
      
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de palet
    '---------------------------------------------------
    If lw1.SelectedItem <> "" Then
        'Nº palet
        devuelve = "{" & NombreTabla & ".numpalet}=" & Val(lw1.SelectedItem)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        devuelve = "numpalet = " & Val(lw1.SelectedItem)
        If Not AnyadirAFormula(cadselect, devuelve) Then Exit Sub
    End If
    
    cadParam = cadParam & "|pImprimeBarras=""1""|"
    numParam = numParam + 1
    
    Sql = ""
    Sql = ClientePalet(lw1.SelectedItem)
    
    cadParam = cadParam & "|pCliente=""" & Trim(Sql) & """|"
    numParam = numParam + 1
   
    If Not HayRegParaInforme(NombreTabla, cadselect) Then Exit Sub
     
     With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .ConSubInforme = True
            .Opcion = 0
            .Titulo = "Impresión de Palet"
            .Show vbModal
    End With
End Sub


Private Sub BotonImprimirTicket()
Dim MIPATH As String
Dim cadImpresion As String, Sql As String
Dim NomImpre As String
Dim NomImpTi As String
Dim bImpre As Boolean

    cadImpresion = "{scafac.codtipom}='" & Text1(1).Text & "' and {scafac.numfactu}=" & Text1(0).Text
    Sql = cadImpresion & " and {scafac.fecfactu}=" & DBSet(Text1(2).Text, "F")
    cadImpresion = cadImpresion & " and {scafac.fecfactu}=Date(" & Year(CDate(Text1(2).Text)) & "," & Month(CDate(Text1(2).Text)) & "," & Day(CDate(Text1(2).Text)) & ")"
    
    If Not HayRegParaInforme("scafac", Sql) Then Exit Sub
    
'    'Obtener que terminal es
'     'Terminal con el que trabajaremos, leemos el nombre del ordenador
'    SQL = ComputerName 'Nombre PC conectado por Terminal Server / local
'    SQL = DevuelveDesdeBDNew(conAri, "spatpvt", "numtermi", "nombrepc", SQL, "T")
'    If Not IsNumeric(SQL) Then
'        MsgBox "No se ha podido establecer la impresora de ticket." & vbCrLf & "Debe configurar primero los parámetros del TPV.", vbExclamation
'    Else
'        bImpre = True
'    End If
'
'    If bImpre Then
'         'Establecemos la impresora de ticket
'         NomImpTi = NombreImpresoraTicket(CInt(SQL))
'         If NomImpTi <> "" Then
'            If Printer.DeviceName <> NomImpTi Then
'                'guardamos la impresora que habia
'                NomImpre = Printer.DeviceName
'                'establecemos la de ticket
'                EstablecerImpresora NomImpTi
'            End If
'        End If
'    End If


    


    MIPATH = App.path & "\Informes\"
'    cadImpresion = cadImpresion & " and {scafac.fecfactu}=Date(" & Year(RSVenta!fecventa) & "," & Month(RSVenta!fecventa) & "," & Day(RSVenta!fecventa) & ")"
    With frmVisReport
        .FormulaSeleccion = cadImpresion
        .SoloImprimir = False
        .OtrosParametros = ""
        .NumeroParametros = 0
        .MostrarTree = False
        .Informe = MIPATH & "rTPVTicket.rpt"
        .ConSubInforme = False
        .Opcion = 93
        .ExportarPDF = False
        .Show vbModal
   End With
   
'   If bImpre Then
'        'volver la impresora a la predeterminada
'        EstablecerImpresora NomImpre
'   End If
   
End Sub

Private Sub TxtAux3_GotFocus(Index As Integer)
    ConseguirFoco txtAux3(Index), Modo
End Sub

Private Sub TxtAux3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index <> 0 And KeyCode <> 38 Then KEYdown KeyCode
End Sub

Private Sub TxtAux3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub TxtAux3_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux3(Index), Modo) Then Exit Sub
End Sub


Private Function ObtenerSelFactura() As String
Dim Cad As String
Dim Rs As ADODB.Recordset

    On Error Resume Next

    Cad = ""
    '******************************************************
    'laura: esto se puede comentar, ya no hay movimiento FTI en la smoval
    If hcoCodTipoM = "FTI" Then
        'no hay albaran directamente va a factura de ticket
        
        'ver si lo encontramos como factura: codtipom, numfactu,fecfactu
        Cad = "SELECT COUNT(*) FROM scafac "
        Cad = Cad & " WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
        If RegistrosAListar(Cad) > 0 Then
            Cad = " WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
        Else
            Cad = ""
        End If
    End If
    '******************************************************
        
    If Cad = "" Then
        'En la smoval estaba e mov. de ALbaran
        Cad = "SELECT codtipom,numfactu,fecfactu FROM scafac1 "
        Cad = Cad & " WHERE codtipoa=" & DBSet(hcoCodTipoM, "T") & " AND numalbar=" & hcoCodMovim & " AND fechaalb=" & DBSet(hcoFechaMov, "F")
        
        Set Rs = New ADODB.Recordset
        Rs.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then 'where para la factura
            Cad = " WHERE codtipom='" & Rs!codTipoM & "' AND numfactu= " & Rs!NumFactu & " AND fecfactu=" & DBSet(Rs!FecFactu, "F")
        Else
            Cad = " WHERE numfactu=-1"
        End If
        Rs.Close
        Set Rs = Nothing
    End If
    ObtenerSelFactura = Cad
End Function


Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim Sql As String

    On Error GoTo EInsertarCab
    
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
        Sql = CadenaInsertarDesdeForm(Me)
        If Sql <> "" Then
            If InsertarOferta(Sql, vTipoMov) Then
                CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                PonerCadenaBusqueda
                PonerModo 2
                'Ponerse en Modo Insertar Lineas
'                BotonMtoLineas 0, "Variedades"
'                BotonAnyadirLinea
                Set frmLPal = New frmVtasLinPalets
                
                frmLPal.ModoExt = 3
                frmLPal.Palet = Text1(0).Text
                frmLPal.Show vbModal
                
                Set frmLPal = Nothing
            End If
        End If
        Text1(0).Text = Format(Text1(0).Text, "0000000")
        CalcularTaraEnvase 1
    End If
    Set vTipoMov = Nothing
    
EInsertarCab:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Function InsertarOferta(vSQL As String, vTipoMov As CTiposMov) As Boolean
Dim MenError As String
Dim bol As Boolean, Existe As Boolean
Dim cambiaSQL As Boolean
Dim devuelve As String

    On Error GoTo EInsertarOferta
    
    bol = True
    
    cambiaSQL = False
    'Comprobar si mientras tanto se incremento el contador de Pedidos
    'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
    Do
        devuelve = DevuelveDesdeBDNew(cAgro, NombreTabla, "numpalet", "numpalet", Text1(0).Text, "N")
        If devuelve <> "" Then
            'Ya existe el contador incrementarlo
            Existe = True
            vTipoMov.IncrementarContador (CodTipoMov)
            Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
            cambiaSQL = True
        Else
            Existe = False
        End If
    Loop Until Not Existe
    If cambiaSQL Then vSQL = CadenaInsertarDesdeForm(Me)
    
    
    'Aqui empieza transaccion
    conn.BeginTrans
    MenError = "Error al insertar en la tabla Cabecera de Palets (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
'    'Actualizar los datos del cliente si es de varios
'    If EsDeVarios Then
'        'Si es cliente de varios actualizar datos cliente en tabla:sclvar
'        MenError = "Modificando datos cliente varios"
'        bol = ActualizarClienteVarios(Text1(4).Text, Text1(6).Text)
'    End If
'
'    If bol Then
'        'Actualizar el campo fechamov (ult. movimiento) de la tabla de clientes (sclien)
'        MenError = "Actualizando Fecha Movimiento del Cliente."
'        bol = ActualizarFecMovCliente
        
        MenError = "Error al actualizar el contador del Palets."
    '    bol = vTipoMov.IncrementarContador("REG")
        vTipoMov.IncrementarContador (CodTipoMov)
'    End If
    
EInsertarOferta:
        If Err.Number <> 0 Then
            MenError = "Insertando Albaran." & vbCrLf & "----------------------------" & vbCrLf & MenError
            MuestraError Err.Number, MenError, Err.Description
            bol = False
        End If
        If bol Then
            conn.CommitTrans
            InsertarOferta = True
        Else
            conn.RollbackTrans
            InsertarOferta = False
        End If
End Function

Private Sub CalcularTaraEnvase(NumLinea As String)
Dim Valor As Currency
Dim TotalCajas As Currency
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim TaraEnvase As String
Dim Forfaits As String
Dim PesoCaja As String

    If CCur(NumLinea) < 1 Then
        Text2(1).Text = ""
        Text2(2).Text = ""
        Exit Sub
    End If

'    'total importes de envases para ese forfait
'    Sql = "select sum(numcajas) "
'    Sql = Sql & " from palets_calibre where numpalet = " & Data1.Recordset.Fields(0)
'    Sql = Sql & " and palets_calibre.numlinea = " & numlinea
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    TotalCajas = 0
'    If Not Rs.EOF Then
'        If Rs.Fields(0).Value > 0 Then TotalCajas = Rs.Fields(0).Value
'    End If
'
'    Rs.Close
'    Set Rs = Nothing
    
    Sql = ""
    Sql = DevuelveDesdeBDNew(cAgro, "palets_variedad", "numcajas", "numpalet", Data1.Recordset.Fields(0), "N", , "numlinea", NumLinea, "N")
    If Sql = "" Then
        TotalCajas = 0
    Else
        TotalCajas = CLng(Sql)
    End If
    
    Forfaits = DevuelveDesdeBDNew(cAgro, "palets_variedad", "codforfait", "numpalet", Data1.Recordset.Fields(0), "N", , "numlinea", NumLinea, "N")
    
    Sql = ""
    Sql = DevuelveDesdeBDNew(cAgro, "forfaits", "pesocaja", "codforfait", Forfaits, "N")
    PesoCaja = ""
    If Sql <> "" Then
        PesoCaja = Format(TransformaPuntosComas(Sql), "###,###,##0.00")
    End If
        
    If PesoCaja <> "" Then
       Text2(0).Text = PesoCaja
    Else
       Text2(0).Text = ""
       PesoCaja = "0"
    End If
    
    Text2(1).Text = Format(TotalCajas, "###,###,##0")
    Valor = Round(CCur(TransformaPuntosComas(PesoCaja)) * TotalCajas, 2)
    If Valor <> 0 Then
        Text2(2).Text = Format(Valor, "###,###,##0.00")
    Else
        Text2(2).Text = ""
    End If


    'Calculo de totales
    Sql = "select palets_variedad.numlinea, round(sum(palets_calibre.numcajas) * forfaits.pesocaja  ,2) "
    Sql = Sql & " from palets_variedad, forfaits, palets_calibre "
    Sql = Sql & " where palets_variedad.numpalet = " & Data1.Recordset.Fields(0) & " and "
    Sql = Sql & " palets_variedad.numpalet = palets_calibre.numpalet and "
    Sql = Sql & " palets_variedad.numlinea = palets_calibre.numlinea and "
    Sql = Sql & " palets_variedad.codforfait = forfaits.codforfait "
    Sql = Sql & " group by 1"
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    TaraEnvase = 0
    While Not Rs.EOF
        TaraEnvase = TaraEnvase + DBLet(Rs.Fields(1).Value, "N")
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    Text2(5).Text = Format(TaraEnvase, "###,###,##0.00")
    
    Text2(6).Text = Format(TaraEnvase + CCur(TransformaPuntosComas(ComprobarCero(Text2(3).Text))), "###,###,##0.00")


End Sub


Private Function ClientePalet(Palet As String) As String
Dim Rs As ADODB.Recordset
Dim Sql As String

    On Error GoTo eClientePalet

    ClientePalet = ""
    Sql = "select pedidos.codclien, clientes.nomclien from palets, pedidos, clientes "
    Sql = Sql & " where palets.numpalet = " & DBSet(Palet, "N")
    Sql = Sql & " and palets.numpedid = pedidos.numpedid "
    Sql = Sql & " and pedidos.codclien = clientes.codclien "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    If Not Rs.EOF Then
        ClientePalet = "Cliente : " & Format(DBLet(Rs.Fields(0).Value, "N"), "000000") & " " & DBLet(Rs.Fields(1).Value, "T")
    End If
    
    Set Rs = Nothing
    Exit Function
    
eClientePalet:
    MuestraError Err.Number, "Cliente del pedido asociado"
End Function


Private Sub CargarColumnas()
    
    lw1.ColumnHeaders.Clear

    If columna = 1 Then
        If Orden = 0 Then
            lw1.ColumnHeaders.Add , , "Palet", 1400
        Else
            lw1.ColumnHeaders.Add , , "Palet v", 1400
        End If
    Else
        lw1.ColumnHeaders.Add , , "Palet", 1400
    End If
    If columna = 2 Then
        If Orden = 0 Then
            lw1.ColumnHeaders.Add , , "Fecha Conf.", 1500
        Else
            lw1.ColumnHeaders.Add , , "Fecha Conf.v", 1500
        End If
    Else
        lw1.ColumnHeaders.Add , , "Fecha Conf.", 1500
    End If
    If columna = 3 Then
        If Orden = 0 Then
            lw1.ColumnHeaders.Add , , "Pedido", 1100, 0
        Else
            lw1.ColumnHeaders.Add , , "Pedido v", 1100, 0
        End If
    Else
        lw1.ColumnHeaders.Add , , "Pedido", 1100, 0
    End If
    If columna = 4 Then
        If Orden = 0 Then
            lw1.ColumnHeaders.Add , , "Variedad", 1300, 1
        Else
            lw1.ColumnHeaders.Add , , "Variedad v", 1300, 1
        End If
    Else
        lw1.ColumnHeaders.Add , , "Variedad", 1300, 1
    End If
    If columna = 5 Then
        If Orden = 0 Then
            lw1.ColumnHeaders.Add , , "Nombre Variedad", 2100, 0
        Else
            lw1.ColumnHeaders.Add , , "Nombre Variedad v", 2100, 0
        End If
    Else
        lw1.ColumnHeaders.Add , , "Nombre Variedad", 2100, 0
    End If
    If columna = 6 Then
        If Orden = 0 Then
            lw1.ColumnHeaders.Add , , "Forfait", 1400, 0
        Else
            lw1.ColumnHeaders.Add , , "Forfait v", 1400, 0
        End If
    Else
        lw1.ColumnHeaders.Add , , "Forfait", 1400, 0
    End If
    If columna = 7 Then
        If Orden = 0 Then
            lw1.ColumnHeaders.Add , , "Nombre Forfait", 2500, 0
        Else
            lw1.ColumnHeaders.Add , , "Nombre Forfait v", 2500, 0
        End If
    Else
        lw1.ColumnHeaders.Add , , "Nombre Forfait", 2500, 0
    End If
    If columna = 8 Then
        If Orden = 0 Then
            lw1.ColumnHeaders.Add , , "Cajas", 1400, 1
        Else
            lw1.ColumnHeaders.Add , , "Cajas v", 1400, 1
        End If
    Else
        lw1.ColumnHeaders.Add , , "Cajas", 1500, 1
    End If
    If columna = 9 Then
        If Orden = 0 Then
            lw1.ColumnHeaders.Add , , "Peso Bruto", 2000, 1
        Else
            lw1.ColumnHeaders.Add , , "Peso Bruto v", 2000, 1
        End If
    Else
        lw1.ColumnHeaders.Add , , "Peso Bruto", 2000, 1
    End If
    If columna = 10 Then
        If Orden = 0 Then
            lw1.ColumnHeaders.Add , , "Peso Neto", 2000, 1
        Else
            lw1.ColumnHeaders.Add , , "Peso Neto v", 2000, 1
        End If
    Else
        lw1.ColumnHeaders.Add , , "Peso Neto", 2000, 1
    End If
    If columna = 11 Then
        If Orden = 0 Then
            lw1.ColumnHeaders.Add , , "Fecha", 0, 1
        Else
            lw1.ColumnHeaders.Add , , "Fecha v", 0, 1
        End If
    Else
        lw1.ColumnHeaders.Add , , "Fecha", 0, 1
    End If
    
    lw1.SmallIcons = frmPpal.imgListPpal


End Sub



Private Sub CargaListview(scolumna1 As String, cadwhere As String, Refrescar As Boolean)
Dim ItmX As ListItem
Dim CampoOrden As String
Dim Descen As String
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim IT As ListItem
Dim TotalArray As Long

Dim TotCajas As Long
Dim TotBruto As Long
Dim TotNeto As Long
Dim Estado As Integer


    CargarColumnas

    If Not Refrescar Then
        Label1(0).visible = True
        DoEvents
    End If

    Sql = "Select palets.numpalet, palets.fechaconf fecha, palets.numpedid, palets_variedad.codvarie, variedades.nomvarie,  "
    Sql = Sql & " palets_variedad.codforfait, forfaits.nomconfe, "
    Sql = Sql & " palets_variedad.numcajas, palets_variedad.pesobrut, palets_variedad.pesoneto "
    Sql = Sql & ", concat(year(palets.fechaconf),right(concat('00',month(palets.fechaconf)),2),right(concat('00',day(palets.fechaconf)),2)) fechaconf"
    Sql = Sql & " FROM ((palets left join palets_variedad on palets.numpalet = palets_variedad.numpalet) "
    Sql = Sql & " LEFT JOIN variedades ON palets_variedad.codvarie = variedades.codvarie)"
    Sql = Sql & " LEFT JOIN forfaits ON palets_variedad.codforfait = forfaits.codforfait "
    
    Sql = Sql & " where (1=1) "
    
    If cadwhere <> "" Then Sql = Sql & " and " & cadwhere
    
    If FiltroCamara <> "" Then Sql = Sql & " and " & FiltroCamara
    
    If scolumna1 <> "" Then
        Sql = Sql & " order by "
        If scolumna1 <> "" Then Sql = Sql & scolumna1

        If Orden = 1 Then Sql = Sql & " desc "
    End If

    If Refrescar Then
        If Orden = 0 Then
            lw1.SortOrder = lvwAscending
        Else
            lw1.SortOrder = lvwDescending
        End If
        Orden = lw1.SortOrder
        If columna = 2 Then
            lw1.SortKey = 10
        Else
            lw1.SortKey = columna - 1
        End If
        lw1.Sorted = True
    Else
        lw1.ListItems.Clear
        
        '[Monica]11/07/2018: limpiamos los totales
        Text1(5).Text = ""
        Text1(6).Text = ""
        Text1(7).Text = ""
        TotCajas = 0
        TotBruto = 0
        TotNeto = 0
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not Rs.EOF
            If PaletEnAlbaran(DBLet(Rs!numpalet, "T"), DBLet(Rs!numpedid, "T")) Then
                Estado = 3
            Else
                If DBLet(Rs!numpedid, "N") <> 0 Then
                    Estado = 2
                Else
                    Estado = 1
                End If
            End If
        
            If Combo1(1).ListIndex = 0 Or (Combo1(1).ListIndex = 1 And Estado = 1) Or (Combo1(1).ListIndex = 2 And Estado = 2) Or (Combo1(1).ListIndex = 3 And Estado = 3) Then
        
                Set IT = lw1.ListItems.Add
                
                IT.Text = Format(DBLet(Rs!numpalet, "N"), "0000000")
                IT.SubItems(1) = DBLet(Rs!fecha, "F")
                IT.SubItems(2) = Format(DBLet(Rs!numpedid, "N"), "0000000")
                IT.SubItems(3) = Format(DBLet(Rs!codvarie, "N"), "000000")
                IT.SubItems(4) = DBLet(Rs!nomvarie, "T")
                IT.ListSubItems.item(4).ToolTipText = DBLet(Rs!nomvarie, "T")
                IT.SubItems(5) = DBLet(Rs!codforfait, "T")
                IT.SubItems(6) = DBLet(Rs!nomconfe, "T")
                IT.ListSubItems.item(6).ToolTipText = DBLet(Rs!nomconfe, "T")
                IT.SubItems(7) = Format(DBLet(Rs!NumCajas, "N"), "###,###,###,##0")
                IT.SubItems(8) = Format(DBLet(Rs!pesobrut, "N"), "###,###,###,##0")
                IT.SubItems(9) = Format(DBLet(Rs!Pesoneto, "N"), "###,###,###,##0")
                IT.SubItems(10) = DBLet(Rs!fechaconf, "T")
                        
                TotCajas = TotCajas + DBLet(Rs!NumCajas, "N")
                TotBruto = TotBruto + DBLet(Rs!pesobrut, "N")
                TotNeto = TotNeto + DBLet(Rs!Pesoneto, "N")
                        
                If Estado = 3 Then
                    IT.ForeColor = vbDarkBlue
                    IT.ListSubItems.item(1).ForeColor = vbDarkBlue
                    IT.ListSubItems.item(2).ForeColor = vbDarkBlue
                    IT.ListSubItems.item(3).ForeColor = vbDarkBlue
                    IT.ListSubItems.item(4).ForeColor = vbDarkBlue
                    IT.ListSubItems.item(5).ForeColor = vbDarkBlue
                    IT.ListSubItems.item(6).ForeColor = vbDarkBlue
                    IT.ListSubItems.item(7).ForeColor = vbDarkBlue
                    IT.ListSubItems.item(8).ForeColor = vbDarkBlue
                    IT.ListSubItems.item(9).ForeColor = vbDarkBlue
                Else
                    If Estado = 2 Then
                        IT.ForeColor = vbRed
                        IT.ListSubItems.item(1).ForeColor = vbRed
                        IT.ListSubItems.item(2).ForeColor = vbRed
                        IT.ListSubItems.item(3).ForeColor = vbRed
                        IT.ListSubItems.item(4).ForeColor = vbRed
                        IT.ListSubItems.item(5).ForeColor = vbRed
                        IT.ListSubItems.item(6).ForeColor = vbRed
                        IT.ListSubItems.item(7).ForeColor = vbRed
                        IT.ListSubItems.item(8).ForeColor = vbRed
                        IT.ListSubItems.item(9).ForeColor = vbRed
                    End If
                End If
                
                TotalArray = TotalArray + 1
                If TotalArray > 300 Then
                    TotalArray = 0
                    DoEvents
                End If
                
            End If
            Rs.MoveNext
        Wend
        
        lw1.Refresh
        
        Rs.Close
        Set Rs = Nothing
    
        ' cargamos los totales
        Text1(5).Text = Format(TotCajas, "###,###,###,##0")
        Text1(6).Text = Format(TotBruto, "###,###,###,##0")
        Text1(7).Text = Format(TotNeto, "###,###,###,##0")
    End If
    
    PonerModo 2

    Label1(0).visible = False
    DoEvents
    
End Sub

Private Function PaletEnAlbaran(NPalet As String, Optional NPedido As String) As Boolean
Dim Sql As String

    If ComprobarCero(NPedido) <> "" Then
        Sql = "select numalbar from albaran where numpedid = " & DBSet(NPedido, "N")
        
        PaletEnAlbaran = (DevuelveValor(Sql) <> 0)
    Else
        Sql = "select palets.numpedid from palets, albaran where palets.numpalet = " & DBSet(NPalet, "N")
        Sql = Sql & " and palets.numpedid = albaran.numpedid "
        
        PaletEnAlbaran = (DevuelveValor(Sql) = 0)
    End If

End Function



