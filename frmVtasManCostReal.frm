VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmVtasManCostReal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Costes Reales"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   Icon            =   "frmVtasManCostReal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1530
      Index           =   0
      Left            =   90
      TabIndex        =   17
      Top             =   495
      Width           =   6975
      Begin VB.Frame Frame3 
         Height          =   825
         Left            =   1530
         TabIndex        =   24
         Top             =   630
         Width           =   5145
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   5
            Left            =   3465
            MaxLength       =   16
            TabIndex        =   28
            Tag             =   "Imp.Coste Calculado|N|N|||cabcostreal|impcostecalc|###,###,##0.00||"
            Text            =   "1234"
            Top             =   405
            Width           =   1500
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   1755
            MaxLength       =   16
            TabIndex        =   3
            Tag             =   "Imp.Coste Real|N|N|||cabcostreal|impcostereal|###,###,##0.00||"
            Text            =   "1234"
            Top             =   405
            Width           =   1500
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   135
            MaxLength       =   16
            TabIndex        =   25
            Tag             =   "Imp.Coste Albaranes|N|N|||cabcostreal|impcostealb|###,###,##0.00||"
            Text            =   "1234"
            Top             =   405
            Width           =   1500
         End
         Begin VB.Label Label6 
            Caption         =   "Coste Calculado"
            Height          =   255
            Index           =   4
            Left            =   3465
            TabIndex        =   29
            Top             =   180
            Width           =   1320
         End
         Begin VB.Label Label6 
            Caption         =   "Coste Real"
            Height          =   255
            Index           =   3
            Left            =   1755
            TabIndex        =   27
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label6 
            Caption         =   "Coste Albaranes"
            Height          =   255
            Index           =   2
            Left            =   135
            TabIndex        =   26
            Top             =   180
            Width           =   1590
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   1035
         MaxLength       =   2
         TabIndex        =   0
         Tag             =   "Cod. Coste|N|N|0|99|cabcostreal|codcoste|00|S|"
         Text            =   "Text1"
         Top             =   270
         Width           =   780
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1875
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   22
         Text            =   "Text2"
         Top             =   270
         Width           =   4815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   630
         MaxLength       =   2
         TabIndex        =   1
         Tag             =   "Mes|N|N|1|12|cabcostreal|mescoste|00|S|"
         Text            =   "12"
         Top             =   675
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   630
         MaxLength       =   4
         TabIndex        =   2
         Tag             =   "A�o|N|N|||cabcostreal|anocoste||S|"
         Text            =   "1234"
         Top             =   1035
         Width           =   510
      End
      Begin VB.Label Label6 
         Caption         =   "A�o"
         Height          =   255
         Index           =   0
         Left            =   225
         TabIndex        =   23
         Top             =   1035
         Width           =   690
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   720
         ToolTipText     =   "Buscar transportista"
         Top             =   300
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Coste"
         Height          =   255
         Index           =   0
         Left            =   225
         TabIndex        =   19
         Top             =   300
         Width           =   510
      End
      Begin VB.Label Label6 
         Caption         =   "Mes"
         Height          =   255
         Index           =   1
         Left            =   225
         TabIndex        =   18
         Top             =   675
         Width           =   690
      End
   End
   Begin VB.Frame FrameAux0 
      Caption         =   "Variedades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3360
      Left            =   90
      TabIndex        =   10
      Top             =   2115
      Width           =   7005
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   6
         Left            =   6030
         MaxLength       =   11
         TabIndex        =   33
         Tag             =   "Imp.Coste|N|N|||lincostreal|impcoste|###,###,##0.00||"
         Text            =   "Imp.Coste"
         Top             =   2565
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   5
         Left            =   4860
         MaxLength       =   8
         TabIndex        =   32
         Tag             =   "Coste Kg|N|N|||lincostreal|costekg|#,##0.0000||"
         Text            =   "Coste "
         Top             =   2565
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   4
         Left            =   3645
         MaxLength       =   16
         TabIndex        =   31
         Tag             =   "Peso Neto|N|N|||lincostreal|pesoneto|###,##0||"
         Text            =   "Peso N"
         Top             =   2565
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   2
         Left            =   855
         MaxLength       =   4
         TabIndex        =   30
         Tag             =   "A�o|N|N|||lincostreal|anocoste|0000|S|"
         Text            =   "a�o"
         Top             =   2565
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   0
         Left            =   90
         MaxLength       =   2
         TabIndex        =   14
         Tag             =   "Codigo|N|N|||lincostreal|codcoste|00|S|"
         Text            =   "codigo"
         Top             =   2565
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   3
         Left            =   1125
         MaxLength       =   6
         TabIndex        =   6
         Tag             =   "Variedad|N|N|||lincostreal|codvarie|00|S|"
         Text            =   "Va"
         Top             =   2565
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1755
         TabIndex        =   13
         Top             =   2565
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton btnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   300
         Index           =   0
         Left            =   1530
         MaskColor       =   &H00000000&
         TabIndex        =   12
         ToolTipText     =   "Buscar Envase"
         Top             =   2565
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   1
         Left            =   540
         MaxLength       =   2
         TabIndex        =   11
         Tag             =   "Mes|N|N|1|12|lincostreal|mescoste|00|S|"
         Text            =   "mes"
         Top             =   2565
         Visible         =   0   'False
         Width           =   240
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   0
         Left            =   135
         TabIndex        =   15
         Top             =   225
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Nuevo"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Eliminar"
               Object.Tag             =   "2"
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc AdoAux 
         Height          =   375
         Index           =   0
         Left            =   3720
         Top             =   225
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
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
         Caption         =   "AdoAux(0)"
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
      Begin MSDataGridLib.DataGrid DataGridAux 
         Bindings        =   "frmVtasManCostReal.frx":000C
         Height          =   2595
         Index           =   0
         Left            =   135
         TabIndex        =   16
         Top             =   630
         Width           =   6600
         _ExtentX        =   11642
         _ExtentY        =   4577
         _Version        =   393216
         AllowUpdate     =   0   'False
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            AllowFocus      =   0   'False
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4830
      TabIndex        =   4
      Tag             =   "   "
      Top             =   5730
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   5730
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   5985
      TabIndex        =   9
      Top             =   5760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   90
      TabIndex        =   7
      Top             =   5535
      Width           =   2385
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   40
         TabIndex        =   8
         Top             =   240
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   4275
      Top             =   3960
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
      Caption         =   "Data1"
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Expandir operaciones"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Generar Costes"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "�ltimo"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   5535
         TabIndex        =   21
         Top             =   90
         Width           =   1215
      End
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
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
      Begin VB.Menu mnGenerar 
         Caption         =   "&Generar Costes"
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
Attribute VB_Name = "frmVtasManCostReal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MANOLO                   -+-+
' +-+- Men�: CLIENTES                  -+-+
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

' +-+-+-+- DISSENY +-+-+-+-
' 1. Posar tots els controls al formulari
' 2. Posar els index correlativament
' 3. Si n'hi han botons de buscar repasar el ToolTipText
' 4. Alliniar els camps num�rics a la dreta i el resto a l'esquerra
' 5. Posar els TAGs
' (si es INTEGER: si PK => m�nim 1; si no PK => m�nim 0; m�xim => 99; format => 00)
' (si es DECIMAL; m�nim => 0; m�xim => 99.99; format => #,###,###,##0.00)
' (si es DATE; format => dd/mm/yyyy)
' 6. Posar els MAXLENGTHs
' 7. Posar els TABINDEXs

Option Explicit

'Dim T1 As Single

Public DatosADevolverBusqueda As String    'Tindr� el n� de text que vol que torne, empipat
Public Event DatoSeleccionado(CadenaSeleccion As String)
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean

' *** declarar els formularis als que vaig a cridar ***
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmVar As frmManVariedad  'variedades
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmCost As frmManNomCoste 'nombre de costes
Attribute frmCost.VB_VarHelpID = -1
'*****************************************************
Private Modo As Byte
'*************** MODOS ********************
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la b�squeda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edici� del camp
'   3.-  Inserci� de nou registre
'   4.-  Modificar
'   5.-  Manteniment Llinies

'+-+-Variables comuns a tots els formularis+-+-+

Dim ModoLineas As Byte
'1.- Afegir,  2.- Modificar,  3.- Borrar,  0.-Passar control a Ll�nies

Dim NumTabMto As Integer 'Indica quin n� de Tab est� en modo Mantenimient
Dim TituloLinea As String 'Descripci� de la ll�nia que est� en Mantenimient
Dim PrimeraVez As Boolean

Private CadenaConsulta As String 'SQL de la taula principal del formulari
Private Ordenacion As String
Private NombreTabla As String  'Nom de la taula
Private NomTablaLineas As String 'Nom de la Taula de ll�nies del Mantenimient en que estem

Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

'Private VieneDeBuscar As Boolean
'Per a quan torna 2 poblacions en el mateix codi Postal. Si ve de pulsar prismatic
'de b�squeda posar el valor de poblaci� seleccionada i no tornar a recuperar de la Base de Datos

Dim btnPrimero As Byte 'Variable que indica el n� del Bot� PrimerRegistro en la Toolbar1
'Dim CadAncho() As Boolean  'array, per a quan cridem al form de ll�nies
Dim indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim CadB As String

Private Sub btnBuscar_Click(Index As Integer)
Dim numNivel As String

    TerminaBloquear
    Select Case Index
        Case 0 'Cuentas contables
            Set frmVar = New frmManVariedad
            frmVar.DatosADevolverBusqueda = "0|1|"
            frmVar.CodigoActual = txtAux(1).Text
            frmVar.Show vbModal
            Set frmVar = Nothing
            PonerFoco txtAux(1)
    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub


Private Sub cmdAceptar_Click()
Dim v As Integer
Dim b As Boolean

    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1

    Select Case Modo
        Case 1  'B�SQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If AnyadirRegistro Then
                    Data1.RecordSource = "Select * from " & NombreTabla & Ordenacion
                    PosicionarData
                    CargaGrid 0, True
                    PonerModo 2
                End If
            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaRegistro Then
                    TerminaBloquear
                    PosicionarData
                    CargaGrid 0, True
                    PonerModo 2
                End If
            Else
                ModoLineas = 0
            End If
        ' *** si n'hi han ll�nies ***
        Case 5 'LL�NIES
            Select Case ModoLineas
                Case 2 'modificar ll�nies
                    If ModificarLinea Then
                        ModoLineas = 0
                        
                        v = AdoAux(NumTabMto).Recordset.Fields(3) 'el 2 es el n� de llinia
                        
                        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCab(True) & Ordenacion
                        PonerCadenaBusqueda
                        b = BLOQUEADesdeFormulario2(Me, Data1, 1)
                        
                        CargaGrid NumTabMto, True
                        
                        PonerFocoGrid Me.DataGridAux(NumTabMto)
                        AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(3).Name & " =" & v)
                        
                        LLamaLineas NumTabMto, 0
                        
                        TerminaBloquear
                        '++monica
                        BloqueaRegistro "cabcostreal", ObtenerWhereCab(False)
                        PosicionarData
                    Else
                        PonerFoco txtAux(12)
                    End If
            End Select
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
'    If PrimeraVez Then PrimeraVez = False
    If PrimeraVez Then
        PrimeraVez = False
        If DatosADevolverBusqueda = "" Then
            PonerModo 0
        Else
            If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                BotonAnyadir
            Else
                PonerModo 1 'b�squeda
                ' *** posar de groc els camps visibles de la clau primaria de la cap�alera ***
                Text1(0).BackColor = vbYellow 'codforfait
                ' ****************************************************************************
            End If
        End If
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    If Modo = 4 Or Modo = 5 Then TerminaBloquear
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim i As Integer

    PrimeraVez = True
    
    ' ICONETS DE LA BARRA
    btnPrimero = 16 'index del bot� "primero"
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'l'1 i el 2 son separadors
        .Buttons(3).Image = 1   'Buscar
        .Buttons(4).Image = 2   'Totss
        'el 5 i el 6 son separadors
        .Buttons(7).Image = 3   'Insertar
        .Buttons(8).Image = 4   'Modificar
        .Buttons(9).Image = 5   'Borrar
        .Buttons(11).Image = 19   'Expandir A�adir, Borrar y Modificar
        'el 10 i el 11 son separadors
        .Buttons(12).Image = 26  'Imprimir
        .Buttons(13).Image = 11  'Eixir
        'el 13 i el 14 son separadors
        .Buttons(btnPrimero).Image = 6  'Primer
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Seg�ent
        .Buttons(btnPrimero + 3).Image = 9 '�ltim
    End With
    
    ' ******* si n'hi han ll�nies *******
    'ICONETS DE LES BARRES ALS TABS DE LL�NIA
    For i = 0 To ToolAux.Count - 1
        With Me.ToolAux(i)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
        End With
    Next i
    ' ***********************************
    Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    
    LimpiarCampos   'Neteja els camps TextBox
    ' ******* si n'hi han ll�nies *******
    DataGridAux(0).ClearFields
    
    '*** canviar el nom de la taula i l'ordenaci� de la cap�alera ***
    NombreTabla = "cabcostreal"
    Ordenacion = " ORDER BY codcoste"
    
    'Mirem com est� guardat el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    Data1.ConnectionString = conn
    '***** cambiar el nombre de la PK de la cabecera *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where codcoste=-1"
    Data1.Refresh
       
    CargaGrid 0, False
       
    ModoLineas = 0
       
    
'    If DatosADevolverBusqueda = "" Then
'        PonerModo 0
'    Else
'        PonerModo 1 'b�squeda
'        ' *** posar de groc els camps visibles de la clau primaria de la cap�alera ***
'        Text1(0).BackColor = vbYellow 'codforfait
'        ' ****************************************************************************
'    End If
End Sub

Private Sub LimpiarCampos()
    On Error Resume Next
    
    limpiar Me   'M�tode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub LimpiarCamposLin(frameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, frameAux  'M�tode general: Neteja els controls TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO s'habiliten, o no, els diversos camps del
'   formulari en funci� del modo en que anem a treballar
Private Sub PonerModo(Kmodo As Byte, Optional indFrame As Integer)
Dim i As Integer, Numreg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo
 
    'Actualisa Iconos Insertar,Modificar,Eliminar
    'ActualizarToolbar Modo, Kmodo
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo, ModoLineas
       
    'Modo 2. N'hi han datos i estam visualisant-los
    '=========================================
    'Posem visible, si es formulari de b�squeda, el bot� "Regresar" quan n'hi han datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = (Modo = 2)
    Else
        cmdRegresar.visible = False
    End If
    
    '=======================================
    b = (Modo = 2)
    'Posar Fleches de desplasament visibles
    Numreg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then Numreg = 2 'Nom�s es per a saber que n'hi ha + d'1 registre
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, Numreg
    '---------------------------------------------
    
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
       
    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
    'Si estem en Insertar a m�s neteja els camps Text1
    BloquearText1 Me, Modo
    BloquearImgBuscar Me, Modo
    
    ' ***** bloquejar tots els controls visibles de la clau primaria de la cap�alera ***
'    If Modo = 4 Then _
'        BloquearTxt Text1(0), True 'si estic en  modificar, bloqueja la clau primaria
    ' **********************************************************************************
    
    Text1(3).Enabled = (Modo = 1)
    Text1(5).Enabled = (Modo = 1)
    
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    For i = 0 To txtAux.Count - 1
        txtAux(i).visible = False
        txtAux(i).Enabled = False
    Next i
    
    btnBuscar(0).Enabled = False
    btnBuscar(0).visible = False
    
    PonerLongCampos

    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 0, False
    End If
    
    b = (Modo = 4) Or (Modo = 2)
    DataGridAux(0).Enabled = b
      
    Frame2(0).Enabled = Not (Modo = 5)
      
    PonerModoOpcionesMenu (Modo) 'Activar opcions men� seg�n modo
    PonerOpcionesMenu   'Activar opcions de men� seg�n nivell
                        'de permisos de l'usuari

EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los TEXT1
    PonerLongCamposGnral Me, Modo, 1
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub

Private Sub PonerModoOpcionesMenu(Modo)
'Actives unes Opcions de Men� i Toolbar seg�n el modo en que estem
Dim b As Boolean, bAux As Boolean
Dim i As Byte
    
    'Barra de CAP�ALERA
    '------------------------------------------
    'b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    b = (Modo = 2 Or Modo = 0)
    'Buscar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnBuscar.Enabled = b
    'Vore Tots
    Toolbar1.Buttons(4).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    'Insertar
    Toolbar1.Buttons(7).Enabled = b And Not DeConsulta
    Me.mnNuevo.Enabled = b And Not DeConsulta
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(9).Enabled = b
    Me.mnEliminar.Enabled = b
    
    'Expandir operaciones
    Toolbar1.Buttons(11).Enabled = True And Not DeConsulta
    'Imprimir
    'Toolbar1.Buttons(12).Enabled = (b Or Modo = 0)
    Toolbar1.Buttons(12).Enabled = b And Not DeConsulta
       
    ' *** si n'hi han ll�nies que tenen grids (en o sense tab) ***
'++monica: si insertamos lo he quitado
'    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
    b = (Modo = 4 Or Modo = 2) And Not DeConsulta
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.AdoAux(i).Recordset.RecordCount > 0)
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i
    
End Sub

Private Sub Desplazamiento(Index As Integer)
'Botons de Despla�ament; per a despla�ar-se pels registres de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index
    PonerCampos
End Sub

Private Function MontaSQLCarga(Index As Integer, enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basant-se en la informaci� proporcionada pel vector de camps
'   crea un SQl per a executar una consulta sobre la base de datos que els
'   torne.
' Si ENLAZA -> Enla�a en el data1
'           -> Si no el carreguem sense enlla�ar a cap camp
'--------------------------------------------------------------------
Dim Sql As String
Dim Tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
               
        Case 0 ' gastos por variedades
            Sql = "SELECT lincostreal.codcoste, lincostreal.mescoste, lincostreal.anocoste, lincostreal.codvarie, "
            Sql = Sql & " variedades.nomvarie, lincostreal.pesoneto, lincostreal.costekg, lincostreal.impcoste "
            Sql = Sql & " FROM lincostreal, variedades "
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE lincostreal.codcoste = -1"
            End If
            Sql = Sql & " and lincostreal.codvarie = variedades.codvarie "
            Sql = Sql & " ORDER BY lincostreal.codvarie "
               
    End Select
    
    MontaSQLCarga = Sql
End Function

Private Sub frmCost_DatoSeleccionado(CadenaSeleccion As String)
'Costes
    Text1(0).Text = RecuperaValor(CadenaSeleccion, 1) 'codcoste
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Variedades
    txtAux(3).Text = RecuperaValor(CadenaSeleccion, 1) 'codvarie
    txtAux2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim Aux As String
    
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabem quins camps son els que mos torna
        'Creem una cadena consulta i posem els datos
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadB = Aux
        '   Com la clau principal es �nica, en posar el sql apuntant
        '   al valor retornat sobre la clau ppal es suficient
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        ' **********************************
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0 'Coste
            AbrirFrmCostes (Index)
        
    End Select
    PonerFoco Text1(Index)
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub


Private Sub mnEliminar_Click()
    BotonEliminar
End Sub


Private Sub mnGenerar_Click()
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonGenerarCostes
End Sub

Private Sub mnModificar_Click()
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************
    
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 3  'B�scar
           mnBuscar_Click
        Case 4  'Tots
            mnVerTodos_Click
        Case 7  'Nou
            mnNuevo_Click
        Case 8  'Modificar
            mnModificar_Click
        Case 9  'Borrar
            mnEliminar_Click
        Case 11  'Expandir operaciones
        
        Case 12 'Imprimir
            mnGenerar_Click
        Case 13    'Eixir
            mnSalir_Click
            
        Case btnPrimero To btnPrimero + 3 'Fleches Despla�ament
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub BotonBuscar()
Dim i As Integer
' ***** Si la clau primaria de la cap�alera no es Text1(0), canviar-ho en <=== *****
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        PonerFoco Text1(0) ' <===
        Text1(0).BackColor = vbYellow ' <===
        ' *** si n'hi han combos a la cap�alera ***
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
' ******************************************************************************
End Sub

Private Sub HacerBusqueda()

    CadB = ObtenerBusqueda2(Me, 1)
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        ' *** foco al 1r camp visible de la cap�alera que siga clau primaria ***
        PonerFoco Text1(0)
        ' **********************************************************************
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
    Dim cad As String
        
    'Cridem al form
    ' **************** arreglar-ho per a vore lo que es desije ****************
    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
    
'Montamos al final: "Cod Diag.|idDiag|N|10�"
    
    
    cad = ""
    cad = cad & "Codigo|cabcostreal.codcoste|N||10�"
    cad = cad & "Descripcion|denominacion|T||60�"
    cad = cad & ParaGrid(Text1(1), 10, "Mes")
    cad = cad & ParaGrid(Text1(2), 10, "A�o")
    
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = NombreTabla & " inner join nombcoste on cabcostreal.codcoste = nombcoste.codcoste "
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        frmB.vDevuelve = "0|2|3|" '*** els camps que volen que torne ***
        frmB.vTitulo = "Costes por meses" ' ***** repasa a��: t�tol de BuscaGrid *****
        frmB.vSelElem = 1

        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha posat valors i tenim que es formulari de b�squeda llavors
        'tindrem que tancar el form llan�ant l'event
        If HaDevueltoDatos Then
            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                cmdRegresar_Click
        Else   'de ha retornat datos, es a decir NO ha retornat datos
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub

Private Sub cmdRegresar_Click()
Dim cad As String
Dim Aux As String
Dim i As Integer
Dim J As Integer

    If Data1.Recordset.EOF Then
        MsgBox "Ning�n registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    cad = ""
    i = 0
    Do
        J = i + 1
        i = InStr(J, DatosADevolverBusqueda, "|")
        If i > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, i - J)
            J = Val(Aux)
            cad = cad & Text1(J).Text & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        PonerModo 2
        'Data1.Recordset.MoveLast
        Data1.Recordset.MoveFirst
        PonerCampos
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub

EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub BotonVerTodos()
'Vore tots
    LimpiarCampos 'Neteja els Text1
    CadB = ""
    
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

Private Sub BotonAnyadir()
Dim NumF As String
    LimpiarCampos 'Huida els TextBox
    PonerModo 3
    
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la cap�alera *******
'    text1(0).Text = SugerirCodigoSiguienteStr("forfaits", "codforfait")
'    FormateaCampo text1(0)
    '******************** canviar taula i camp **************************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        NumF = NuevoCodigo
    Else
        NumF = ""
    End If
    '********************************************************************
       
    Text1(0) = NumF
    PonerFoco Text1(0) '*** 1r camp visible que siga PK ***
    
    Text1(3).Text = 0
    Text1(4).Text = 0
    Text1(5).Text = 0
    ' *** si n'hi han camps de descripci� a la cap�alera ***
    'PosarDescripcions

End Sub

Private Sub BotonGenerarCostes()
Dim Sql As String
Dim Sql1 As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim ImpCosteVariedad As Currency
Dim Coste As Currency
Dim CosteTot As Currency
Dim Diferencia As Currency
Dim NumAlb As Currency
Dim NumLin As Currency

    On Error GoTo eBotonGenerarCostes

    If EstaCalculado Then
        Sql = "El Coste Real para el concepto, mes y a�o ya est� calculado. �Desea volverlo a recalcular?"
        If MsgBox(Sql, vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Sub
        End If
    End If
    
    
    
    If Not PodemosGenerarCostes Then
        Sql = "No podemos generar los Costes Reales. " & vbCrLf & vbCrLf
        Sql = Sql & "La difencia entre �stos y los calculados es superior al porcentaje de desviacion "
        Sql = Sql & vParamAplic.PorcDesvCostes & "." & vbCrLf & vbCrLf
        Sql = Sql & "Revise."
        MsgBox Sql, vbExclamation
        Exit Sub
    End If
    
    conn.BeginTrans

    Sql = "delete from albaran_costreal where codcoste = " & DBSet(Data1.Recordset!codCoste, "N")
    Sql = Sql & " and numalbar in (select numalbar from albaran where month(fechaalb) = "
    Sql = Sql & DBSet(Data1.Recordset!mescoste, "N") & " and year(fechaalb) = "
    Sql = Sql & DBSet(Data1.Recordset!anocoste, "N") & ")"

    conn.Execute Sql
    
    Sql1 = "select lincostreal.* from lincostreal where codcoste = " & DBSet(Data1.Recordset!codCoste, "N")
    Sql1 = Sql1 & " and mescoste = " & DBSet(Data1.Recordset!mescoste, "N")
    Sql1 = Sql1 & " and anocoste = " & DBSet(Data1.Recordset!anocoste, "N")
    Sql1 = Sql1 & " order by codvarie "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        ImpCosteVariedad = DBLet(Rs!ImpCoste, "N")
        CosteTot = 0
        
        Sql = "select albaran_variedad.numalbar, albaran_variedad.numlinea, albaran_variedad.pesoneto "
        Sql = Sql & " from albaran_variedad inner join albaran on albaran_variedad.numalbar = albaran.numalbar "
        Sql = Sql & " and year(albaran.fechaalb) = " & DBSet(Data1.Recordset!anocoste, "N")
        Sql = Sql & " and month(albaran.fechaalb) = " & DBSet(Data1.Recordset!mescoste, "N")
        Sql = Sql & " and albaran_variedad.codvarie = " & DBSet(Rs!codvarie, "N")
        
        Set Rs1 = New ADODB.Recordset
        Rs1.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        Sql2 = ""
        While Not Rs1.EOF
            Coste = 0
            If DBLet(Rs!Pesoneto, "N") <> 0 Then
                Coste = Round2(DBLet(Rs1!Pesoneto, "N") * DBLet(Rs!ImpCoste, "N") / DBLet(Rs!Pesoneto, "N"), 4)
            End If
            
            Sql2 = Sql2 & "(" & DBSet(Rs1!NumAlbar, "N") & "," & DBSet(Rs1!NumLinea, "N") & "," & DBSet(Rs!codCoste, "N") & ","
            Sql2 = Sql2 & DBSet(Coste, "N") & "),"
            
            NumAlb = DBLet(Rs1!NumAlbar, "N")
            NumLin = DBLet(Rs1!NumLinea, "N")
            
            CosteTot = CosteTot + Coste
            Rs1.MoveNext
        Wend
        
        ' quitamos la ultima coma de sql2
        If Sql2 <> "" Then
            Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
            
            Sql2 = "insert into albaran_costreal(numalbar, numlinea, codcoste, impcoste) values " & Sql2
            conn.Execute Sql2
        
            ' en el ultimo registro vamos a poner el resto de ImpCosteVariedad - CosteTot
            Diferencia = ImpCosteVariedad - CosteTot
            
            If Diferencia <> 0 Then
                Sql2 = "update albaran_costreal set impcoste = impcoste + " & DBSet(Diferencia, "N")
                Sql2 = Sql2 & " where numalbar = " & DBSet(NumAlb, "N") & " and numlinea = "
                Sql2 = Sql2 & DBSet(NumLin, "N")
            
                conn.Execute Sql2
            End If
        End If
        
        Set Rs1 = Nothing
        
        Rs.MoveNext
    Wend

    Set Rs = Nothing

eBotonGenerarCostes:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Error generando los costes reales.", vbExclamation
        conn.RollbackTrans
    Else
        conn.CommitTrans
        MsgBox "Proceso terminado correctamente.", vbExclamation
    End If
End Sub

Private Sub BotonModificar()

    PonerModo 4

    ' *** bloquejar els camps visibles de la clau primaria de la cap�alera ***
    BloquearTxt Text1(0), True
    
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco Text1(1)
End Sub

Private Sub BotonEliminar()
Dim cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
    ' ***************************************************************************

    ' *************** canviar la pregunta ****************
    cad = "�Seguro que desea eliminar el Coste?"
    cad = cad & vbCrLf & "C�digo: " & Data1.Recordset.Fields(0)
    cad = cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(1)
    
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Coste", Err.Description
End Sub

Private Sub PonerCampos()
Dim i As Integer
Dim codPobla As String, desPobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la cap�alera
    
    ' *** si n'hi han ll�nies en datagrids ***
    'For i = 0 To DataGridAux.Count - 1
    CargaGrid 0, True
    If Not AdoAux(0).Recordset.EOF Then
        PonerCamposForma2 Me, AdoAux(0), 2, "FrameAux" & 0
    End If
    
    ' ************* configurar els camps de les descripcions de la cap�alera *************
    Text2(0).Text = PonerNombreDeCod(Text1(0), "nombcoste", "denominacion")
    ' ********************************************************************************
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu
End Sub

Private Sub cmdCancelar_Click()
Dim i As Integer
Dim v

    Select Case Modo
        Case 1, 3 'B�squeda, Insertar
                LimpiarCampos
                If Data1.Recordset.EOF Then
                    PonerModo 0
                Else
                    PonerModo 2
                    PonerCampos
                End If
                ' *** foco al primer camp visible de la cap�alera ***
                PonerFoco Text1(0)

        Case 4  'Modificar
                TerminaBloquear
                PonerModo 2
                PonerCampos
                ' *** primer camp visible de la cap�alera ***
                PonerFoco Text1(0)
        
        Case 5 'LL�NIES
            Select Case ModoLineas
                Case 1 'afegir ll�nia
                    ModoLineas = 0
                    ' *** les ll�nies que tenen datagrid (en o sense tab) ***
                    If NumTabMto = 0 Or NumTabMto = 1 Or NumTabMto = 2 Or NumTabMto = 4 Then
                        DataGridAux(NumTabMto).AllowAddNew = False
                        ' **** repasar si es diu Data1 l'adodc de la cap�alera ***
                        'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'Modificar
                        LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                        DataGridAux(NumTabMto).Enabled = True
                        DataGridAux(NumTabMto).SetFocus

                        ' *** si n'hi han camps de descripci� dins del grid, els neteje ***
                        'txtAux2(2).text = ""

                    End If
                    
'                    ' *** si n'hi han tabs ***
'                    SituarTab (NumTabMto + 1)
                    
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        AdoAux(NumTabMto).Recordset.MoveFirst
                    End If

                Case 2 'modificar ll�nies
                    ModoLineas = 0
                    
                    ' *** si n'hi han tabs ***
'                    SituarTab (NumTabMto + 1)
                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                    PonerModo 4
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        ' *** l'Index de Fields es el que canvie de la PK de ll�nies ***
                        v = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el n� de llinia
                        AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & v)
                        ' ***************************************************************
                    End If

            End Select
            
            PosicionarData
            
            ' *** si n'hi han ll�nies en grids i camps fora d'estos ***
            If Not AdoAux(NumTabMto).Recordset.EOF Then
                DataGridAux_RowColChange NumTabMto, 1, 1
            Else
                LimpiarCamposFrame NumTabMto
            End If
    End Select
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Sql As String
'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm2(Me, 1)
    If Not b Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
    If (Modo = 3) Then 'insertar
        'comprobar si existe ya el cod. del campo clave primaria
        Sql = ""
        Sql = DevuelveDesdeBDNew(cAgro, "cabcostreal", "codcoste", "codcoste", Text1(0).Text, "N", , "mescoste", Text1(1).Text, "N", "anocoste", Text1(2).Text, "N")
        If Sql <> "" Then b = False
    End If
    ' ************************************************************************************
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Sub PosicionarData()
Dim cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la cap�alera, no llevar els () ***
    cad = "(codcoste=" & DBSet(Text1(0).Text, "T") & ") and mescoste = " & DBSet(Text1(1).Text, "N")
    cad = cad & " and anocoste = " & DBSet(Text1(2).Text, "N")
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(Data1, cad, Indicador) Then
    If SituarDataMULTI(Data1, cad, Indicador, False) Then
        If ModoLineas <> 1 Then PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
End Sub

Private Function Eliminar() As Boolean
Dim vWhere As String
Dim Sql As String

    On Error GoTo FinEliminar

    conn.BeginTrans
    ' ***** canviar el nom de la PK de la cap�alera, repasar codEmpre *******
    vWhere = " WHERE codcoste=" & DBSet(Data1.Recordset!codCoste, "N") & " and mescoste = " & _
             DBSet(Data1.Recordset!mescoste, "N") & " and anocoste = " & DBSet(Data1.Recordset!anocoste, "N")
    
    ' *** elimina los registros de albaran_costreal *** '
    Sql = "delete from albaran_costreal where codcoste = " & DBSet(Data1.Recordset!codCoste, "N")
    Sql = Sql & " and numalbar in (select numalbar from albaran where month(fechaalb)=" & DBSet(Data1.Recordset!mescoste, "N")
    Sql = Sql & " and year(fechaalb) = " & DBSet(Data1.Recordset!anocoste, "N") & ") "
    
    conn.Execute Sql
        
    ' ***** elimina les ll�nies ****
    conn.Execute "DELETE FROM lincostreal " & vWhere
        
    'Eliminar la CAP�ALERA
    conn.Execute "Delete from " & NombreTabla & vWhere
       
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    
    ' ***************** configurar els LostFocus dels camps de la cap�alera *****************
    Select Case Index
        Case 0 'codigo de coste
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "nombcoste", "denominacion")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Coste: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "�Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCost = New frmManNomCoste
                        frmCost.DatosADevolverBusqueda = "0|1|"
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmCost.Show vbModal
                        Set frmCost = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            End If
        
        Case 1, 2 ' mes, a�o
            PonerFormatoEntero Text1(Index)
        
        Case 4 ' coste real
            PonerFormatoDecimal Text1(Index), 1
        
        
        
    End Select
        ' ***************************************************************************
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvan�ar/Retrocedir els camps en les fleches de despla�ament del teclat.
    KEYdown KeyCode
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
End Sub



'************* LLINIES: ****************************
Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
'-- pon el bloqueo aqui
    'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
    Select Case Button.Index
        Case 1
            BotonAnyadirLinea Index
        Case 2
            BotonModificarLinea Index
        Case 3
            BotonEliminarLinea Index
        Case Else
    End Select
    'End If
End Sub

Private Sub BotonEliminarLinea(Index As Integer)
Dim Sql As String
Dim vWhere As String
Dim Eliminar As Boolean

    On Error GoTo Error2

    ModoLineas = 3 'Posem Modo Eliminar Ll�nia
    
    If Modo = 4 Then 'Modificar Cap�alera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index

    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar(Index) Then Exit Sub
    NumTabMto = Index
    Eliminar = False
   
    vWhere = ObtenerWhereCab(True)
    
    ' ***** independentment de si tenen datagrid o no,
    ' canviar els noms, els formats i el DELETE *****
    Select Case Index
        Case 0 'envases
            Sql = "�Seguro que desea eliminar la Variedad?"
            Sql = Sql & vbCrLf & "Codigo : " & AdoAux(Index).Recordset!codvarie
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM lincostreal "
                Sql = Sql & vWhere & " AND codvarie= " & AdoAux(Index).Recordset!codvarie
            End If
            
    End Select

    If Eliminar Then
        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        conn.Execute Sql
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        CargaGrid Index, True
        If Not SituarDataTrasEliminar(AdoAux(Index), NumRegElim, True) Then
'            PonerCampos
            
        End If
        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
        ' *** si n'hi han tabs ***
'        SituarTab (NumTabMto + 1)
    End If
    
    ModoLineas = 0
    PosicionarData
    
    Exit Sub
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando linea", Err.Description
End Sub

Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vTabla As String
Dim anc As Single
Dim i As Integer
    
    ModoLineas = 1 'Posem Modo Afegir Ll�nia
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Cap�alera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index
    
    ' *** bloquejar la clau primaria de la cap�alera ***
    BloquearTxt Text1(0), True

    ' *** posar el nom del les distintes taules de ll�nies ***
    Select Case Index
        Case 0: vTabla = "lincostreal"
    End Select
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 0, 1 ' *** pose els index dels tabs de ll�nies que tenen datagrid ***
            ' *** canviar la clau primaria de les ll�nies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            
            If Index = 0 Then NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", vWhere)

            AnyadirLinea DataGridAux(Index), AdoAux(Index)
    
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
            
            LLamaLineas Index, ModoLineas, anc
        
            Select Case Index
                ' *** valor per defecte a l'insertar i formateig de tots els camps ***
                Case 0 'envases
                    txtAux(0).Text = Text1(0).Text 'codcoste
                    txtAux(1).Text = NumF 'numlinea
                    txtAux(2).Text = ""
                    txtAux2(2).Text = ""
                    BloquearTxt txtAux(2), False
                    BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux0"
                    PonerFoco txtAux(2)
            End Select
            
    End Select
End Sub

Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim i As Integer
    Dim J As Integer
    
    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub
    
    ModoLineas = 2 'Modificar ll�nia
       
    If Modo = 4 Then 'Modificar Cap�alera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index
    ' *** bloqueje la clau primaria de la cap�alera ***
    BloquearTxt Text1(0), True
    BloquearTxt Text1(1), True
    BloquearTxt Text1(2), True
  
    Text1(0).Text = AdoAux(0).Recordset!codCoste
    Text1(1).Text = AdoAux(0).Recordset!mescoste
    Text1(2).Text = AdoAux(0).Recordset!anocoste
  
    Select Case Index
        Case 0 ' *** pose els index de ll�nies que tenen datagrid (en o sense tab) ***
            If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
                i = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
                DataGridAux(Index).Scroll 0, i
                DataGridAux(Index).Refresh
            End If
              
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
    End Select
    
    Select Case Index
        ' *** valor per defecte al modificar dels camps del grid ***
        Case 0 ' variedades
        
            For J = 0 To 3
                txtAux(J).Text = DataGridAux(Index).Columns(J).Text
            Next J
            txtAux2(3).Text = DataGridAux(Index).Columns(4).Text
            txtAux(4).Text = DataGridAux(Index).Columns(5).Text
            txtAux(5).Text = DataGridAux(Index).Columns(6).Text
            txtAux(6).Text = DataGridAux(Index).Columns(7).Text
            
            
'            BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux0"
            
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    PonerFoco txtAux(5)
    ' ***************************************************************************************
End Sub

Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim b As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    DeseleccionaGrid DataGridAux(Index)
       
    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Ll�nies
    Select Case Index
        Case 0 'cuentas contables
             For jj = 5 To 6
                txtAux(jj).visible = b
                txtAux(jj).Enabled = b
                txtAux(jj).Top = alto
            Next jj
    End Select
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
Dim Pesoneto As String
Dim CosteUnit As String
Dim ImpCoste As String

    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    

    ' ******* configurar el LostFocus dels camps de ll�nies (dins i fora grid) ********
    Select Case Index
        Case 5 'coste unitario
            PonerFormatoDecimal txtAux(Index), 7
            CalcularImportes txtAux(4).Text, txtAux(5).Text, txtAux(6).Text, 1
            
        Case 6 ' importe de coste
            PonerFormatoDecimal txtAux(Index), 1
            CalcularImportes txtAux(4).Text, txtAux(5).Text, txtAux(6).Text, 0
    End Select
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
   If Not txtAux(Index).MultiLine Then ConseguirFocoLin txtAux(Index)
End Sub


Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not txtAux(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not txtAux(Index).MultiLine Then
        If KeyAscii = teclaBuscar Then
            If Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) Then
                Select Case Index
                    Case 1: 'codvarie
                        KeyAscii = 0
                        btnBuscar_Click (0)
                End Select
            End If
        Else
            KEYpress KeyAscii
        End If
    End If
End Sub

Private Function DatosOkLlin(nomframe As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim b As Boolean
Dim Cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte

    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False
        
    b = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not b Then Exit Function
    
    ' ******************************************************************************
    DatosOkLlin = b
    
EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function SepuedeBorrar(ByRef Index As Integer) As Boolean
    SepuedeBorrar = False
    
    ' *** si cal comprovar alguna cosa abans de borrar ***
'    Select Case Index
'        Case 0 'cuentas bancarias
'            If AdoAux(Index).Recordset!ctaprpal = 1 Then
'                MsgBox "No puede borrar una Cuenta Principal. Seleccione antes otra cuenta como Principal", vbExclamation
'                Exit Function
'            End If
'    End Select
    ' ****************************************************
    
    SepuedeBorrar = True
End Function


Private Sub DataGridAux_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
Dim i As Byte

    If ModoLineas <> 1 Then
        Select Case Index
            Case 0 'cuentas bancarias
                If DataGridAux(Index).Columns.Count > 2 Then
'                    txtAux(11).Text = DataGridAux(Index).Columns("direccio").Text
'                    txtAux(12).Text = DataGridAux(Index).Columns("observac").Text
                End If
                
            Case 1 'departamentos
                If DataGridAux(Index).Columns.Count > 2 Then
'                    txtAux(21).Text = DataGridAux(Index).Columns(5).Text
'                    txtAux(22).Text = DataGridAux(Index).Columns(6).Text
'                    txtAux(23).Text = DataGridAux(Index).Columns(8).Text
'                    txtAux(24).Text = DataGridAux(Index).Columns(15).Text
'                    txtAux2(22).Text = DataGridAux(Index).Columns(7).Text
                End If
                
        End Select
        
    Else 'vamos a Insertar
        Select Case Index
            Case 0 'cuentas bancarias
'                txtAux(11).Text = ""
'                txtAux(12).Text = ""
            Case 1 'departamentos
                For i = 21 To 24
'                   txtAux(i).Text = ""
                Next i
'               txtAux2(22).Text = ""
            Case 2 'Tarjetas
'               txtAux(50).Text = ""
'               txtAux(51).Text = ""
        End Select
    End If
End Sub


Private Sub CargaFrame(Index As Integer, enlaza As Boolean)
Dim tip As Integer
Dim i As Byte

    AdoAux(Index).ConnectionString = conn
    AdoAux(Index).RecordSource = MontaSQLCarga(Index, enlaza)
    AdoAux(Index).CursorType = adOpenDynamic
    AdoAux(Index).LockType = adLockPessimistic
    AdoAux(Index).Refresh
    
    If Not AdoAux(Index).Recordset.EOF Then
        PonerCamposForma2 Me, AdoAux(Index), 2, "FrameAux" & Index
    Else
        ' *** si n'hi han tabs sense datagrids, li pose els valors als camps ***
        NetejaFrameAux "FrameAux3" 'neteja nom�s lo que te TAG
    End If
End Sub

' *** si n'hi han tabs sense datagrids ***
Private Sub NetejaFrameAux(nom_frame As String)
Dim Control As Object
    
    For Each Control In Me.Controls
        If (Control.Tag <> "") Then
            If (Control.Container.Name = nom_frame) Then
                If TypeOf Control Is TextBox Then
                    Control.Text = ""
                ElseIf TypeOf Control Is ComboBox Then
                    Control.ListIndex = -1
                End If
            End If
        End If
    Next Control

End Sub

Private Sub CargaGrid(Index As Integer, enlaza As Boolean)
Dim b As Boolean
Dim i As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)

    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
    
    Select Case Index
        Case 0 'cuentas contables
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;N||||0|;" 'codcoste, mes, a�o
            tots = tots & "S|txtAux(3)|T|Cod.|600|;S|btnBuscar(0)|B|||;S|txtAux2(3)|T|Variedad|2400|;"
            tots = tots & "S|txtAux(4)|T|Peso Neto|1000|;S|txtAux(5)|T|Coste|1000|;S|txtAux(6)|T|Imp.Coste|1000|;"
            
            arregla tots, DataGridAux(Index), Me
        
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
    End Select
    
    DataGridAux(Index).ScrollBars = dbgAutomatic
      
    ' **** si n'hi han ll�nies en grids i camps fora d'estos ****
'    If Not AdoAux(Index).Recordset.EOF Then
'        DataGridAux_RowColChange Index, 1, 1
'    Else
''        LimpiarCamposFrame Index
'    End If
      
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub

Private Sub InsertarLinea()
'Inserta registre en les taules de Ll�nies
Dim nomframe As String
Dim b As Boolean

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'envases
    End Select
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomframe) Then
            b = BLOQUEADesdeFormulario2(Me, Data1, 1)
            Select Case NumTabMto
                Case 0 ' *** els index de les llinies en grid (en o sense tab) ***
                     CargaGrid NumTabMto, True
                    If b Then BotonAnyadirLinea NumTabMto
            End Select
           
'            SituarTab (NumTabMto + 1)
        End If
    End If
End Sub

Private Function ModificarLinea() As Boolean
'Modifica registre en les taules de Ll�nies
Dim nomframe As String
Dim v As Integer
Dim Importe As Currency
Dim Sql As String

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    ModificarLinea = False
    If DatosOkLlin("FrameAux0") Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, "FrameAux0") Then
            ModoLineas = 0
            
            v = AdoAux(NumTabMto).Recordset.Fields(3) 'el 2 es el n� de llinia
            
            'calcular el coste calculado
            Importe = CosteCalculado(Data1.Recordset!codCoste, Data1.Recordset!mescoste, Data1.Recordset!anocoste)
            Sql = "update cabcostreal set impcostecalc = " & DBSet(Importe, "N") & " where codcoste= "
            Sql = Sql & DBSet(Data1.Recordset!codCoste, "N") & " and mescoste = " & DBSet(Data1.Recordset!mescoste, "N")
            Sql = Sql & " and anocoste = " & DBSet(Data1.Recordset!anocoste, "N")
            
            conn.Execute Sql
            
            CargaGrid 0, True
            
            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
            PonerFocoGrid Me.DataGridAux(NumTabMto)
            AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(3).Name & " =" & v)
            
            LLamaLineas NumTabMto, 0
            ModificarLinea = True
        End If
    End If
End Function

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la cap�alera ***
    vWhere = vWhere & " codcoste=" & DBSet(Me.Data1.Recordset!codCoste, "N")
    vWhere = vWhere & " and mescoste = " & DBSet(Me.Data1.Recordset!mescoste, "N")
    vWhere = vWhere & " and anocoste = " & DBSet(Me.Data1.Recordset!anocoste, "N")
    
    ObtenerWhereCab = vWhere
End Function

'' *** neteja els camps dels tabs de grid que
''estan fora d'este, i els camps de descripci� ***
Private Sub LimpiarCamposFrame(Index As Integer)
    On Error Resume Next
 
'    Select Case Index
'        Case 0 'Cuentas Bancarias
'            txtAux(11).Text = ""
'            txtAux(12).Text = ""
'        Case 1 'Departamentos
'            txtAux(21).Text = ""
'            txtAux(22).Text = ""
'            txtAux2(22).Text = ""
'            txtAux(23).Text = ""
'            txtAux(24).Text = ""
'        Case 2 'Tarjetas
'            txtAux(50).Text = ""
'            txtAux(51).Text = ""
'        Case 4 'comisiones
'            txtAux2(2).Text = ""
'    End Select
'
    If Err.Number <> 0 Then Err.Clear
End Sub

'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del rat�n.
'Private Sub DataGridAux_GotFocus(Index As Integer)
'  WheelHook DataGridAux(Index)
'End Sub
'Private Sub DataGridAux_LostFocus(Index As Integer)
'  WheelUnHook
'End Sub

Private Sub printNou()
    With frmImprimir2
        .cadTabla2 = "nombcoste"
        .Informe2 = "rManNomCoste.rpt"
        If CadB <> "" Then
            '.cadRegSelec = Replace(SQL2SF(CadB), "clientes", "clientes_1")
            .cadRegSelec = SQL2SF(CadB)
        Else
            .cadRegSelec = ""
        End If
        ' *** repasar el nom de l'adodc ***
        '.cadRegActua = Replace(POS2SF(Data1, Me), "clientes", "clientes_1")
        .cadRegActua = POS2SF(Data1, Me)
        ' *** repasar codEmpre ***
        .cadTodosReg = ""
        '.cadTodosReg = "{itinerar.codempre} = " & codEmpre
        ' *** repasar si li pose ordre o no ****
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomEmpre & "'|pOrden={nombcoste.codcoste}|"
        '.OtrosParametros2 = "pEmpresa='" & vEmpresa.NomEmpre & "'|"
        ' *** posar el n� de par�metres que he posat en OtrosParametros2 ***
        '.NumeroParametros2 = 1
        .NumeroParametros2 = 2
        ' ******************************************************************
        .MostrarTree2 = False
        .InfConta2 = False
        .ConSubInforme2 = False
        .SubInformeConta = "rep0"
        .Show vbModal
    End With
End Sub


Private Sub AbrirFrmCostes(indice As Integer)
    Set frmCost = New frmManNomCoste
    frmCost.DatosADevolverBusqueda = "0|1|"
    frmCost.Show vbModal
    Set frmCost = Nothing
End Sub

Private Function AnyadirRegistro() As Boolean
Dim Sql As String
Dim b As Boolean
    On Error GoTo eAnyadirRegistro

    conn.BeginTrans

    b = SumaCosteAlbaran


    b = InsertarDesdeForm2(Me, 1)
    If b Then
        Sql = "insert into lincostreal (codcoste, mescoste, anocoste, codvarie, pesoneto, costekg, impcoste) "
        Sql = Sql & "select " & DBSet(Text1(0).Text, "N") & "," & DBSet(Text1(1).Text, "N") & "," & DBSet(Text1(2).Text, "N") & ",albaran_variedad.codvarie, sum(albaran_variedad.pesoneto),0,0 "
        Sql = Sql & " from albaran_variedad inner join albaran on albaran_variedad.numalbar = albaran.numalbar "
        Sql = Sql & " where month(fechaalb) = " & DBSet(Text1(1).Text, "N") & " and year(fechaalb) = " & DBSet(Text1(2).Text, "N")
        Sql = Sql & " group by 1,2,3,4,6,7 order by 1,2,3,4 "
        
        conn.Execute Sql
    End If
        
eAnyadirRegistro:
    If Err.Number <> 0 Or Not b Then
        conn.RollbackTrans
        AnyadirRegistro = False
        MsgBox "Error en a�adir registro.", vbExclamation
    Else
        conn.CommitTrans
        AnyadirRegistro = True
    End If
End Function


Private Function CosteCalculado(Coste As String, mes As String, ano As String) As Currency
Dim Sql As String
Dim Rs As ADODB.Recordset

    On Error GoTo eCosteCalculado

    CosteCalculado = 0

    Sql = "select sum(impcoste) from lincostreal where codcoste = " & DBSet(Coste, "N") & " and "
    Sql = Sql & " mescoste = " & DBSet(mes, "N") & " and anocoste = " & DBSet(ano, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then CosteCalculado = DBLet(Rs.Fields(0).Value, "N")
    
    Exit Function
    
eCosteCalculado:
    MuestraError Err.Number, "Coste calculado"
End Function

Private Function ModificaRegistro() As Boolean
Dim Sql As String
Dim b As Boolean
    On Error GoTo eModificaRegistro

    conn.BeginTrans

    b = ModificaDesdeFormulario2(Me, 1)
    If b Then
    
    
    
    End If
        
eModificaRegistro:
    If Err.Number <> 0 Or Not b Then
        conn.RollbackTrans
        ModificaRegistro = False
        MsgBox "Error en modificar registro.", vbExclamation
    Else
        conn.CommitTrans
        ModificaRegistro = True
    End If
End Function

Private Sub CalcularImportes(Pesoneto As String, Costelin As String, Importe As String, tipo As Byte)
'tipo 1: me dan el costelinea
'tipo 0: me dan el importe

    Pesoneto = ComprobarCero(Pesoneto)
    Importe = ComprobarCero(Importe)
    Costelin = ComprobarCero(Costelin)
    
    Select Case tipo
        Case 0
            If CCur(Pesoneto) <> 0 Then
                txtAux(5) = Round2(CCur(ImporteSinFormato(Importe)) / CCur(ImporteSinFormato(Pesoneto)), 4)
            Else
                txtAux(5).Text = "0"
            End If
        Case 1
            txtAux(6) = Round2(CCur(ImporteSinFormato(Costelin)) * CCur(ImporteSinFormato(Pesoneto)), 2)
    End Select
End Sub

Private Function PodemosGenerarCostes() As Boolean
Dim CosteTotal As Currency
Dim SumaCostes As Currency
Dim Valor As Currency
Dim Sql As String

    On Error GoTo ePodemosGenerarCostes

    PodemosGenerarCostes = False

    Sql = "select sum(impcoste) from lincostreal where codcoste = " & DBSet(Data1.Recordset!codCoste, "N")
    Sql = Sql & " and mescoste = " & DBSet(Data1.Recordset!mescoste, "N")
    Sql = Sql & " and anocoste = " & DBSet(Data1.Recordset!anocoste, "N")
    
    SumaCostes = DevuelveValor(Sql)

    Sql = "select impcostereal from cabcostreal where codcoste = " & DBSet(Data1.Recordset!codCoste, "N")
    Sql = Sql & " and mescoste = " & DBSet(Data1.Recordset!mescoste, "N")
    Sql = Sql & " and anocoste = " & DBSet(Data1.Recordset!anocoste, "N")
    
    CosteTotal = DevuelveValor(Sql)
    Valor = 0
    If CosteTotal <> 0 Then
        Valor = Round2((CosteTotal - SumaCostes) * 100 / CosteTotal, 2)
    End If
    
    If Valor < 0 Then Valor = Valor * (-1)
    
    PodemosGenerarCostes = (Valor <= vParamAplic.PorcDesvCostes)
    Exit Function
    
ePodemosGenerarCostes:
     If Err.Number <> 0 Then
         MuestraError Err.Number, "Podemos Generar Costes"
     End If
End Function


Private Function SumaCosteAlbaran() As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset

    On Error GoTo eSumaCosteAlbaran

    SumaCosteAlbaran = False
    
    Sql = "select sum(impcoste) from albaran_costes inner join albaran on albaran_costes.numalbar = albaran.numalbar "
    Sql = Sql & " where month(albaran.fechaalb) = " & DBSet(Text1(1).Text, "N") & " and "
    Sql = Sql & " year(albaran.fechaalb) = " & DBSet(Text1(2).Text, "N")
    Sql = Sql & " and albaran_costes.tipogasto = 0"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Text1(3).Text = ""
    If Not Rs.EOF Then
        Text1(3).Text = DBLet(Rs.Fields(0).Value, "N")
        PonerFormatoDecimal Text1(3), 1
    End If
        
    SumaCosteAlbaran = True
    Exit Function

eSumaCosteAlbaran:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Suma de Costes de Albar�n."
    End If
End Function

Private Function EstaCalculado() As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset

    On Error GoTo eEstaCalculado

    EstaCalculado = False
    
    Sql = "select count(*) from cabcostreal where codcoste = " & DBSet(Text1(0).Text, "N")
    Sql = Sql & " and mescoste = " & DBSet(Text1(1).Text, "N")
    Sql = Sql & " and anocoste = " & DBSet(Text1(2).Text, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If DBLet(Rs.Fields(0).Value, "N") <> 0 Then EstaCalculado = True
    End If

eEstaCalculado:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Est� calculado " & Err.Description
    End If
End Function
