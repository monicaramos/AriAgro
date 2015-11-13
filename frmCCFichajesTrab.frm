VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCCFichajesTrab 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fichajes Trabajador"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13845
   Icon            =   "frmCCFichajesTrab.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   13845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAux 
      BackColor       =   &H80000005&
      Height          =   255
      Index           =   1
      Left            =   10650
      TabIndex        =   22
      Tag             =   "Procesado|N|N|0|1|cctrabaconf|sinacabar|||"
      Top             =   6150
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtAux 
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   4
      Left            =   6570
      MaxLength       =   4
      TabIndex        =   5
      Tag             =   "Código Coste|N|S|||cctrabaconf|codcoste|0000|N|"
      Text            =   "1234567"
      Top             =   6150
      Width           =   765
   End
   Begin VB.TextBox txtAux2 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   290
      Index           =   4
      Left            =   7740
      TabIndex        =   20
      Text            =   "12345678901234567890"
      Top             =   6150
      Width           =   2190
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   2160
      TabIndex        =   1
      Tag             =   "Fecha|F|N|||cctrabaconf|fecha|dd/mm/yyyy||"
      Text            =   "1234567890"
      ToolTipText     =   "Fecha|F|N|||cctrabaconf|fecha|dd/mm/yyyy||"
      Top             =   6150
      Width           =   1005
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   300
      Index           =   0
      Left            =   3300
      MaskColor       =   &H00000000&
      TabIndex        =   19
      ToolTipText     =   "Buscar fecha"
      Top             =   6150
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   300
      Index           =   1
      Left            =   7410
      MaskColor       =   &H00000000&
      TabIndex        =   18
      ToolTipText     =   "Buscar concepto"
      Top             =   6120
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   300
      Index           =   2
      Left            =   1080
      MaskColor       =   &H00000000&
      TabIndex        =   15
      ToolTipText     =   "Buscar trabajador"
      Top             =   6120
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1350
      TabIndex        =   14
      Top             =   6150
      Width           =   705
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   270
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "Trabajador|N|N|||cctrabaconf|codtraba|000000|S|"
      Text            =   "trabajado"
      Top             =   6120
      Width           =   750
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   3630
      MaxLength       =   8
      TabIndex        =   2
      Text            =   "12345678"
      Top             =   6150
      Width           =   795
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   4560
      MaxLength       =   8
      TabIndex        =   3
      Text            =   "12345678"
      Top             =   6150
      Width           =   885
   End
   Begin VB.CheckBox chkAux 
      BackColor       =   &H80000005&
      Height          =   255
      Index           =   0
      Left            =   10200
      TabIndex        =   6
      Tag             =   "Procesado|N|N|0|1|cctrabaconf|procesado|||"
      Top             =   6150
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   7
      Left            =   5550
      MaxLength       =   10
      TabIndex        =   4
      Tag             =   "Linea Coste|N|N|||cctrabaconf|codlinconf|000||"
      Text            =   "1234567890"
      Top             =   6150
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   90
      TabIndex        =   7
      Top             =   7410
      Width           =   2865
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
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   12210
      TabIndex        =   9
      Top             =   7590
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   11040
      TabIndex        =   8
      Top             =   7590
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Index           =   2
      Left            =   750
      Top             =   6030
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   13845
      _ExtentX        =   24421
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   21
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
            Object.ToolTipText     =   "Modificacion Masiva"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Borrado Masivo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Procesos Varios"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Primero"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Último"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Index           =   0
         Left            =   8550
         TabIndex        =   13
         Top             =   60
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   12210
      TabIndex        =   11
      Top             =   7590
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSDataGridLib.DataGrid DataGridAux 
      Bindings        =   "frmCCFichajesTrab.frx":000C
      Height          =   6810
      Index           =   2
      Left            =   90
      TabIndex        =   21
      Top             =   540
      Width           =   13530
      _ExtentX        =   23865
      _ExtentY        =   12012
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
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   5
      Left            =   10620
      MaxLength       =   20
      TabIndex        =   16
      Tag             =   "Fecha Ini|FH|N|||cctrabaconf|fechaini|yyyy-mm-dd hh:mm:ss|S|"
      Text            =   "f.ini"
      Top             =   6150
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   6
      Left            =   11430
      MaxLength       =   20
      TabIndex        =   17
      Tag             =   "Fecha Fin|FH|N|||cctrabaconf|fechafin|yyyy-mm-dd hh:mm:ss|S|"
      Text            =   "f.fin"
      Top             =   6150
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgFec 
      Height          =   240
      Index           =   0
      Left            =   2700
      Picture         =   "frmCCFichajesTrab.frx":0024
      ToolTipText     =   "Buscar fecha"
      Top             =   7410
      Width           =   240
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
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
      Begin VB.Menu mnModMasiva 
         Caption         =   "Modificación Masiva"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnBorMasivo 
         Caption         =   "Borrado Masivo"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnProcesos 
         Caption         =   "&Procesos Varios"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
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
Attribute VB_Name = "frmCCFichajesTrab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MANOLO                   -+-+
' +-+- Menú: CLIENTES                  -+-+
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

' +-+-+-+- DISSENY +-+-+-+-
' 1. Posar tots els controls al formulari
' 2. Posar els index correlativament
' 3. Si n'hi han botons de buscar repasar el ToolTipText
' 4. Alliniar els camps numérics a la dreta i el resto a l'esquerra
' 5. Posar els TAGs
' (si es INTEGER: si PK => mínim 1; si no PK => mínim 0; màxim => 99; format => 00)
' (si es DECIMAL; mínim => 0; màxim => 99.99; format => #,###,###,##0.00)
' (si es DATE; format => dd/mm/yyyy)
' 6. Posar els MAXLENGTHs
' 7. Posar els TABINDEXs

Option Explicit

'Dim T1 As Single

Public DatosADevolverBusqueda As String    'Tindrà el nº de text que vol que torne, empipat
Public Event DatoSeleccionado(CadenaSeleccion As String)
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean

' *** declarar els formularis als que vaig a cridar ***
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmC1 As frmCal 'calendario fecha
Attribute frmC1.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1

Private WithEvents frmCon As frmCCManConcep 'conceptos
Attribute frmCon.VB_VarHelpID = -1

Private WithEvents frmZon As frmBasico 'zonas
Attribute frmZon.VB_VarHelpID = -1
Private WithEvents frmTra As frmBasico 'trabajadores ( de recoleccion )
Attribute frmTra.VB_VarHelpID = -1
Private WithEvents frmCat As frmBasico 'salarios o categorias ( de recoleccion )
Attribute frmCat.VB_VarHelpID = -1

Private WithEvents frmFor As frmManForfaits 'confeccion
Attribute frmFor.VB_VarHelpID = -1

'*****************************************************
Private Modo As Byte
'*************** MODOS ********************
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'   5.-  Manteniment Llinies

'+-+-Variables comuns a tots els formularis+-+-+

Dim ModoLineas As Byte
'1.- Afegir,  2.- Modificar,  3.- Borrar,  0.-Passar control a Llínies

Dim NumTabMto As Integer 'Indica quin nº de Tab està en modo Mantenimient
Dim TituloLinea As String 'Descripció de la llínia que està en Mantenimient
Dim PrimeraVez As Boolean

Private CadenaConsulta As String 'SQL de la taula principal del formulari
Private Ordenacion As String
Private NombreTabla As String  'Nom de la taula
Private NomTablaLineas As String 'Nom de la Taula de llínies del Mantenimient en que estem

Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

'Private VieneDeBuscar As Boolean
'Per a quan torna 2 poblacions en el mateix codi Postal. Si ve de pulsar prismatic
'de búsqueda posar el valor de població seleccionada i no tornar a recuperar de la Base de Datos

Dim btnPrimero As Byte 'Variable que indica el nº del Botó PrimerRegistro en la Toolbar1
'Dim CadAncho() As Boolean  'array, per a quan cridem al form de llínies
Dim indice As Byte 'Index del txtaux on es poses els datos retornats des d'atres Formularis de Mtos
Dim CadB As String

Dim CodTipoMov As String

Dim FecIni As String
Dim FecFin As String

Private BuscaChekc As String


Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 0 ' Fecha de coste
            btnFec (0)
            
        Case 1 ' Concepto
            indice = 4
            
            Set frmCon = New frmCCManConcep
            frmCon.DatosADevolverBusqueda = "0|1|"
            frmCon.CodigoActual = txtAux(4).Text
            frmCon.Show vbModal
            Set frmCon = Nothing
            PonerFoco txtAux(4)
            
        Case 2 ' Trabajadores
            indice = 0
            Set frmTra = New frmBasico
            AyudaTrabajadores frmTra, txtAux(indice)
            Set frmTra = Nothing
            PonerFoco txtAux(0)
            
    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, AdoAux(2), 1
End Sub


Private Sub btnFec(Index As Integer)
Dim esq As Long
Dim dalt As Long
Dim menu As Long
Dim obj As Object

    Set frmC = New frmCal
    
    esq = btnBuscar(Index).Left
    dalt = btnBuscar(Index).Top
        
    Set obj = btnBuscar(Index).Container
      
    While btnBuscar(Index).Parent.Name <> obj.Name
           esq = esq + obj.Left
           dalt = dalt + obj.Top
           Set obj = obj.Container
    Wend
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    frmC.Left = esq + btnBuscar(Index).Parent.Left + 30
    frmC.Top = dalt + btnBuscar(Index).Parent.Top + btnBuscar(Index).Height + menu - 40

    btnBuscar(0).Tag = Index '<===
    Select Case Index
        Case 0
            indice = 1
    End Select
    ' *** repasar si el camp es txtAux o txtaux ***
    If txtAux(indice).Text <> "" Then frmC.NovaData = txtAux(indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o txtaux ***
    PonerFoco txtAux(indice) '<===
    ' ********************************************
End Sub

Private Sub chkAux_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "chkAux(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "chkAux(" & Index & ")|"
    End If
End Sub

Private Sub chkAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    
    Select Case Modo
        Case 1  'BÚSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                txtAux(1).Tag = ""
                txtAux(2).Tag = ""
                txtAux(3).Tag = ""
                
                If InsertarDesdeForm2(Me, 1) Then
                    TerminaBloquear
                    CargaGrid
                    PosicionarData
                    DataGridAux(2).AllowAddNew = False
                    BotonAnyadir
                End If
            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                txtAux(1).Tag = ""
                txtAux(2).Tag = ""
                txtAux(3).Tag = ""
                'If ModificaDesdeFormulario2(Me, 1) Then
                If ModificaRegistro Then
                    TerminaBloquear
                    CargaGrid
                    PosicionarData
                End If
            Else
                ModoLineas = 0
            End If
            
        ' **************************
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Function ModificaRegistro() As Boolean
Dim Sql As String

    On Error GoTo eModificarRegistro

    ModificaRegistro = False

    Sql = "update cctrabaconf set codlinconf = " & DBSet(txtAux(7).Text, "N")
    Sql = Sql & ", codcoste = " & DBSet(txtAux(4).Text, "N")
    Sql = Sql & ", fechaini = " & DBSet(txtAux(5).Text, "FH")
    Sql = Sql & ", fechafin = " & DBSet(txtAux(6).Text, "FH")
    Sql = Sql & " where fechaini = " & DBSet(FecIni, "FH")
    Sql = Sql & " and fechafin = " & DBSet(FecFin, "FH")
    Sql = Sql & " and codtraba = " & DBSet(txtAux(0).Text, "N")
    conn.Execute Sql
    
    ModificaRegistro = True
    Exit Function
    
eModificarRegistro:
    MuestraError Err.Number, "Modifica Registro", Err.Descripc
End Function


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
'    If PrimeraVez Then PrimeraVez = False
    If PrimeraVez Then
        PrimeraVez = False
        If DatosADevolverBusqueda = "" Then
            PonerModo 2
        Else
            If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                BotonAnyadir
            Else
                PonerModo 1 'búsqueda
                ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
                txtAux(0).BackColor = vbYellow 'codforfait
                ' ****************************************************************************
            End If
        End If
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia(0).Value
    If Modo = 4 Then TerminaBloquear
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim I As Integer

    PrimeraVez = True
    
    ' ICONETS DE LA BARRA
    btnPrimero = 17 'index del botó "primero"
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
        .Buttons(11).Image = 26  'Modificacion masiva
        .Buttons(12).Image = 20  'Borrado masivo
        .Buttons(13).Image = 19  'Procesos varios
        
        'el 10 i el 11 son separadors
        .Buttons(14).Image = 10  'Imprimir
        .Buttons(15).Image = 11  'Eixir
        'el 13 i el 14 son separadors
'        .Buttons(btnPrimero).Image = 6  'Primer
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Següent
'        .Buttons(btnPrimero + 3).Image = 9 'Últim
    End With
    
    
    
    LimpiarCampos   'Neteja els camps TextBox
    ' ******* si n'hi han llínies *******
    DataGridAux(2).ClearFields
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "cctrabaconf"
    Ordenacion = " ORDER BY codtraba, fechaini, fechafin"
    
    'Mirem com està guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    CadenaConsulta = "Select cctrabaconf.codtraba, straba.nomtraba, date(fechaini), time(fechaini) horaini, time(fechafin) horafin, "
    CadenaConsulta = CadenaConsulta & "cctrabaconf.codlinconf, cctrabaconf.codcoste, ccconcostes.nomcoste,  procesado, IF(procesado=1,'*','') as dprocesado, fechaini, fechafin, sinacabar, IF(sinacabar=1,'*','') as dsinacabar from (" & NombreTabla
    CadenaConsulta = CadenaConsulta & " left join straba on cctrabaconf.codtraba = straba.codtraba) left join ccconcostes on cctrabaconf.codcoste = ccconcostes.codcoste where (1=1) "
    
    AdoAux(2).ConnectionString = conn
    AdoAux(2).RecordSource = CadenaConsulta
    AdoAux(2).Refresh
       
    CargaGrid ""
    
    ModoLineas = 0
    PonerCampos

'    PonerModo 2
       
End Sub

Private Sub LimpiarCampos()
    On Error Resume Next
    
    limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub LimpiarCamposLin(frameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, frameAux  'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO s'habiliten, o no, els diversos camps del
'   formulari en funció del modo en que anem a treballar
Private Sub PonerModo(Kmodo As Byte, Optional indFrame As Integer)
Dim I As Integer, NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo
 
    BuscaChekc = ""
 
    Modo = Kmodo
    
    
    b = (Modo = 2)
    If b Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    For I = 0 To txtAux.Count - 1
        txtAux(I).visible = Not b
    Next I
    
    txtAux2(0).visible = Not b
    txtAux2(4).visible = Not b
    btnBuscar(0).visible = Not b
    btnBuscar(1).visible = Not b
    btnBuscar(2).visible = Not b
    
    cmdAceptar.visible = Not b
    cmdCancelar.visible = Not b
    Me.DataGridAux(2).Enabled = b
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = b
    
    'Si estamos modo Modificar bloquear clave primaria
    BloquearTxt txtAux(0), (Modo = 4)
'    BloquearTxt txtAux(1), (Modo = 4)
'    BloquearTxt txtAux(2), (Modo = 4)
'    BloquearTxt txtAux(3), (Modo = 4)
    
    btnBuscar(0).Enabled = (Modo = 1 Or Modo = 3)
    btnBuscar(2).Enabled = (Modo = 1 Or Modo = 3)
    chkAux(0).visible = (Modo = 1)
    chkAux(1).visible = (Modo = 1)
    
    PonerLongCampos
    PonerModoOpcionesMenu Kmodo 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu  'En funcion del usuario
    
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub PonerContRegIndicador()
'si estamos en modo ver registros muestra el numero de registro en el que estamos
'situados del total de registros mostrados: 1 de 24
Dim cadReg As String

    If (Modo = 2 Or Modo = 0) Then
        cadReg = PonerContRegistros(Me.AdoAux(0))
        If CadB = "" Then
            lblIndicador.Caption = cadReg
        Else
            lblIndicador.Caption = "BUSQUEDA: " & cadReg
        End If
    End If
End Sub



Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los txtaux
    PonerLongCamposGnral Me, Modo, 3
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub

Private Sub PonerModoOpcionesMenu(Modo)
'Actives unes Opcions de Menú i Toolbar según el modo en que estem
Dim b As Boolean, bAux As Boolean
Dim I As Byte
    
    'Barra de CAPÇALERA
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
    
    b = (Modo = 2 And AdoAux(2).Recordset.RecordCount > 0) And Not DeConsulta
    If b Then b = b And AdoAux(2).Recordset!procesado = 0
    'Modificar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(9).Enabled = b
    Me.mnEliminar.Enabled = b
    'modificacion masiva
    Toolbar1.Buttons(11).Enabled = b
    Me.mnModMasiva.Enabled = b
    'borrado masivo
    Toolbar1.Buttons(12).Enabled = b
    Me.mnBorMasivo.Enabled = b
    'procesos varios
    Toolbar1.Buttons(13).Enabled = b
    Me.mnProcesos.Enabled = b
    
    'Imprimir
    Toolbar1.Buttons(14).Enabled = True And Not DeConsulta
       
End Sub

Private Sub Desplazamiento(Index As Integer)
'Botons de Desplaçament; per a desplaçar-se pels registres de control Data
    If AdoAux(2).Recordset.EOF Then Exit Sub
    DesplazamientoData AdoAux(2), Index
    PonerCampos
End Sub

Private Function MontaSQLCarga(Index As Integer, enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basant-se en la informació proporcionada pel vector de camps
'   crea un SQl per a executar una consulta sobre la base de datos que els
'   torne.
' Si ENLAZA -> Enlaça en el adoaux(2)
'           -> Si no el carreguem sense enllaçar a cap camp
'--------------------------------------------------------------------
Dim Sql As String
Dim Tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
               
        Case 0 'TRABAJADORES
            Sql = "SELECT cclindia1.fecha, cclindia1.codcoste, cclindia1.numlinea, "
            Sql = Sql & " cclindia1.codtraba, straba.nomtraba, time(cclindia1.fechaini) horaini, "
            Sql = Sql & " time(cclindia1.fechafin) horafin, cclindia1.horas, cclindia1.fechaini, cclindia1.fechafin "
            Sql = Sql & " FROM cclindia1 INNER JOIN straba ON cclindia1.codtraba = straba.codtraba "
            
            If enlaza Then
                Sql = Sql & Replace(ObtenerWhereCab(True), "cccabdia", "cclindia1")
            Else
                Sql = Sql & " WHERE cclindia1.fecha is null "
            End If
            Sql = Sql & " ORDER BY cclindia1.fecha, cclindia1.codcoste, cclindia1.numlinea"
               
        Case 1 'CATEGORIAS
            Sql = "SELECT cclindia2.fecha, cclindia2.codcoste, cclindia2.numlinea, "
            Sql = Sql & " cclindia2.codcateg, salarios.nomcateg, cclindia2.horas "
            Sql = Sql & " FROM cclindia2 INNER JOIN  salarios ON cclindia2.codcateg = salarios.codcateg "
            
            If enlaza Then
                Sql = Sql & Replace(ObtenerWhereCab(True), "cccabdia", "cclindia2")
            Else
                Sql = Sql & " WHERE cclindia2.fecha is null "
            End If
            Sql = Sql & " ORDER BY cclindia2.fecha, cclindia2.codcoste, cclindia2.numlinea"
        
        Case 2 'CABECERA
            Sql = "SELECT cccabdia.fecha, cccabdia.codcoste, ccconcostes.nomcoste, cccabdia.observac"
            Sql = Sql & " FROM cccabdia INNER JOIN ccconcostes ON cccabdia.codcoste = ccconcostes.codcoste "
            
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE cccabdia.fecha is null "
            End If
            Sql = Sql & " ORDER BY cccabdia.fecha, cccabdia.codcoste"
            
    End Select
    
    MontaSQLCarga = Sql
End Function

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim Aux As String
    
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabem quins camps son els que mos torna
        'Creem una cadena consulta i posem els datos
        CadB = ""
        Aux = ValorDevueltoFormGrid(txtAux(0), CadenaDevuelta, 1)
        CadB = Aux
        Aux = "cccabdia." & ValorDevueltoFormGrid(txtAux(1), CadenaDevuelta, 2)
        CadB = CadB & " and " & Aux
        '   Com la clau principal es única, en posar el sql apuntant
        '   al valor retornat sobre la clau ppal es suficient
'        CadenaConsulta = "select * from " & NombreTabla & "  WHERE " & CadB & " " & Ordenacion
        
        CadenaConsulta = "Select cctrabaconf.codtraba, straba.nomtraba, date(fechaini), time(fechaini) horaini, time(fechafin) horafin, "
        CadenaConsulta = CadenaConsulta & "cctrabaconf.codlinconf, cctrabaconf.codcoste, ccconcostes.nomcoste,  procesado, IF(procesado=1,'*','') as dprocesado, fechaini, fechafin from (" & NombreTabla
        CadenaConsulta = CadenaConsulta & " left join straba on cctrabaconf.codtraba = straba.codtraba) left join ccconcostes on cctrabaconf.codcoste = ccconcostes.codcoste where (1=1) "
        
        ' **********************************
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o txtaux ***
    txtAux(indice).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmC1_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o txtaux ***
    txtAux(indice).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmCat_DatoSeleccionado(CadenaSeleccion As String)
'Categorias
    txtAux(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00") 'codigo de categoria
    txtAux2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmCon_DatoSeleccionado(CadenaSeleccion As String)
'Concepto de coste
    txtAux(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000") 'concepto de coste
    txtAux2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub


Private Sub frmZ_Actualizar(vCampo As String)
     txtAux(indice).Text = vCampo
End Sub

Private Sub frmZon_DatoSeleccionado(CadenaSeleccion As String)
'Zonas
    txtAux(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'codigo de zona
    txtAux2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
'Trabajadores
    txtAux(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000") 'codigo de trabajador
    txtAux2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
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
            indice = 1
        Case 1
            indice = 8
        
    End Select
    ' *** repasar si el camp es txtAux o txtaux ***
    If txtAux(indice).Text <> "" Then frmC.NovaData = txtAux(indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o txtaux ***
    PonerFoco txtAux(indice) '<===
    ' ********************************************

End Sub

Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        indice = 10
        frmZ.pTitulo = "Observaciones de la Orden de Confección"
        frmZ.pValor = txtAux(indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco txtAux(indice)
    End If
End Sub

Private Sub mnBorMasivo_Click()
    frmCCListados.CadBusqueda = ""
    frmCCListados.OpcionListado = 4
    frmCCListados.Show vbModal
    
    CargaGrid
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    frmCCListados.OpcionListado = 5
    frmCCListados.Show vbModal
End Sub

Private Sub mnModificar_Click()
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(adoaux(2).Recordset.Fields(0).Value), FormatoCampo(txtaux(0))) Then Exit Sub
    ' ***************************************************************************
    
    If BLOQUEADesdeFormulario2(Me, AdoAux(2), 1) Then BotonModificar
End Sub

Private Sub mnModMasiva_Click()
    If CadB = "" Then
        MsgBox "Debe introducir previamente un criterio de búsqueda.", vbExclamation
    Else
        frmCCListados.OpcionListado = 3
        frmCCListados.CadBusqueda = CadB
        frmCCListados.Show vbModal
        
        CargaGrid CadB
        
    End If
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnProcesos_Click()
        
    frmCCListados.OpcionListado = 7
    frmCCListados.CadBusqueda = CadB
    frmCCListados.Show vbModal

    CargaGrid CadB
    
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 3  'Búscar
           mnBuscar_Click
        Case 4  'Tots
            mnVerTodos_Click
        Case 7  'Nou
            mnNuevo_Click
        Case 8  'Modificar
            mnModificar_Click
        Case 9  'Borrar
            mnEliminar_Click
            
        Case 11 ' Modificacion masiva
            mnModMasiva_Click
        Case 12 ' Borrado masivo
            mnBorMasivo_Click
        Case 13 ' Procesos varios
            mnProcesos_Click
        Case 14 'Imprimir
            mnImprimir_Click
        Case 15    'Eixir
            mnSalir_Click
            
'        Case btnPrimero To btnPrimero + 3 'Fleches Desplaçament
'            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub BotonBuscar()
Dim I As Integer
Dim anc As Single

    ' ***************** canviar per la clau primaria ********
    CargaGrid "cctrabaconf.codtraba = -1"
    '*******************************************************************************
    'Buscar
    For I = 0 To txtAux.Count - 1
        txtAux(I).Text = ""
    Next I
    txtAux2(0).Text = ""
    txtAux2(4).Text = ""
    
    chkAux(0).Value = 0
    chkAux(1).Value = 0
    
    LLamaLineas 2, 1, DataGridAux(2).Top + 206
    PonerFoco txtAux(0)



'' ***** Si la clau primaria de la capçalera no es txtaux(0), canviar-ho en <=== *****
'    If Modo <> 1 Then
'        LimpiarCampos
'
'        anc = DataGridAux(2).Top
'        If DataGridAux(2).Row < 0 Then
'            anc = anc + 210
'        Else
'            anc = anc + DataGridAux(2).RowTop(DataGridAux(2).Row) + 5
'        End If
'
'        LLamaLineas 2, 1, anc
'
'        PonerModo 1
'        PonerFoco txtAux(0) ' <===
'        txtAux(0).BackColor = vbYellow ' <===
'        ' *** si n'hi han combos a la capçalera ***
'    Else
'        HacerBusqueda
'        If AdoAux(2).Recordset.EOF Then
'            txtAux(kCampo).Text = ""
'            txtAux(kCampo).BackColor = vbYellow
'            PonerFoco txtAux(kCampo)
'        End If
'    End If
'' ******************************************************************************
End Sub

Private Sub HacerBusqueda()
    
    txtAux(1).Tag = ""
    txtAux(2).Tag = ""
    txtAux(3).Tag = ""
    
    If txtAux(1).Text <> "" And txtAux(2).Text <> "" Then
        txtAux(5).Text = txtAux(1).Text & " " & txtAux(2).Text
        txtAux(6).Tag = ""
    End If
    If txtAux(1).Text <> "" And txtAux(3).Text <> "" Then
        txtAux(6).Text = txtAux(1).Text & " " & txtAux(3).Text
        txtAux(5).Tag = ""
    End If
    If txtAux(1).Text <> "" And txtAux(2).Text = "" And txtAux(3).Text = "" Then
        txtAux(5).Text = txtAux(1).Text
        txtAux(5).Tag = Replace(txtAux(5).Tag, "FH", "FHF")
        txtAux(6).Tag = ""
    End If
    If txtAux(1).Text = "" And txtAux(2).Text <> "" Then
        txtAux(5).Text = txtAux(2).Text
        txtAux(5).Tag = Replace(txtAux(5).Tag, "FH", "FHH")
        txtAux(6).Tag = ""
    End If
    If txtAux(1).Text = "" And txtAux(3).Text <> "" Then
        txtAux(6).Text = txtAux(3).Text
        txtAux(6).Tag = Replace(txtAux(6).Tag, "FH", "FHH")
        txtAux(5).Tag = ""
    End If
    
    'cadB = ObtenerBusqueda2(Me, 1)
    CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    
    txtAux(1).Tag = txtAux(1).ToolTipText
    
    txtAux(5).Tag = "Fecha Ini|FH|N|||cctrabaconf|fechaini|yyyy-mm-dd hh:mm:ss|S|"
    txtAux(6).Tag = "Fecha Ini|FH|N|||cctrabaconf|fechafin|yyyy-mm-dd hh:mm:ss|S|"
    
    
    If chkVistaPrevia(0) = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        CadenaConsulta = "Select cctrabaconf.codtraba, straba.nomtraba, date(fechaini), time(fechaini) horaini, time(fechafin) horafin, "
        CadenaConsulta = CadenaConsulta & "cctrabaconf.codlinconf, cctrabaconf.codcoste, ccconcostes.nomcoste,  procesado, IF(procesado=1,'*','') as dprocesado, fechaini, fechafin,  sinacabar, IF(sinacabar=1,'*','') as dsinacabar from (" & NombreTabla
        CadenaConsulta = CadenaConsulta & " left join straba on cctrabaconf.codtraba = straba.codtraba) left join ccconcostes on cctrabaconf.codcoste = ccconcostes.codcoste where (1=1) "
        CadenaConsulta = CadenaConsulta & " and " & CadB
        PonerCadenaBusqueda
    Else
        ' *** foco al 1r camp visible de la capçalera que siga clau primaria ***
        PonerFoco txtAux(0)
        ' **********************************************************************
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
    Dim cad As String
        
    'Cridem al form
    ' **************** arreglar-ho per a vore lo que es desije ****************
    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
    cad = ""
    cad = cad & "Codigo|cctrabaconf.codtraba|T|000000|10·"
    cad = cad & "Trabajador|straba.nomtraba|T||25·"
    cad = cad & "Fecha|date(fechaini)|T|dd/mm/yyyy|15·"
    cad = cad & "Hora Inicio|time(fechaini)|T|hh:mm:ss|10·"
    cad = cad & "Hora Fin|time(fechafin)|T|hh:mm:ss|10·"
    cad = cad & "Código|cctrabaconf.codcoste|N|000000|10·"
    cad = cad & "Denominacion|ccconcostes.nomcoste|T||20·"
    
    If cad <> "" Then
        
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        cad = "(" & NombreTabla & " left join straba on cctrabaconf.codtraba = straba.codtraba) " & _
               " left join ccconcostes on cctrabaconf.codcoste = ccconcostes.codcoste "
             
        frmB.vtabla = cad 'NombreTabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        frmB.vDevuelve = "0|1|" '*** els camps que volen que torne ***
        frmB.vTitulo = "Fichajes Trabajador" ' ***** repasa açò: títol de BuscaGrid *****
        frmB.vSelElem = 0

        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha posat valors i tenim que es formulari de búsqueda llavors
        'tindrem que tancar el form llançant l'event
        If HaDevueltoDatos Then
            If (Not AdoAux(2).Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                cmdRegresar_Click
        Else   'de ha retornat datos, es a decir NO ha retornat datos
            PonerFoco txtAux(kCampo)
        End If
    End If
End Sub

Private Sub cmdRegresar_Click()
Dim cad As String
Dim Aux As String
Dim I As Integer
Dim J As Integer

    If AdoAux(2).Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    cad = ""
    I = 0
    Do
        J = I + 1
        I = InStr(J, DatosADevolverBusqueda, "|")
        If I > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, I - J)
            J = Val(Aux)
            cad = cad & txtAux(J).Text & "|"
        End If
    Loop Until I = 0
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    AdoAux(2).RecordSource = CadenaConsulta
    AdoAux(2).Refresh
    CargaGrid
    If AdoAux(2).Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        PonerModo 2
        'adoaux(2).Recordset.MoveLast
        AdoAux(2).Recordset.MoveFirst
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
    LimpiarCampos 'Neteja els txtaux
    CadB = ""
    
    If chkVistaPrevia(0).Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select cctrabaconf.codtraba, straba.nomtraba, date(fechaini), time(fechaini) horaini, time(fechafin) horafin, "
        CadenaConsulta = CadenaConsulta & "cctrabaconf.codlinconf, cctrabaconf.codcoste, ccconcostes.nomcoste,  procesado, IF(procesado=1,'*','') as dprocesado, fechaini, fechafin, sinacabar, IF(sinacabar=1,'*','') as dsinacabar from (" & NombreTabla
        CadenaConsulta = CadenaConsulta & " left join straba on cctrabaconf.codtraba = straba.codtraba) left join ccconcostes on cctrabaconf.codcoste = ccconcostes.codcoste where (1=1) "
        PonerCadenaBusqueda
    End If
End Sub

Private Sub BotonAnyadir()
Dim NumF As String
Dim anc As Single
Dim I As Integer

    
    PonerModo 3
    
    LimpiarCampos 'Huida els TextBox
    
            
    AnyadirLinea DataGridAux(2), AdoAux(2)

    anc = DataGridAux(2).Top
    If DataGridAux(2).Row < 0 Then
        anc = anc + 206
    Else
        anc = anc + DataGridAux(2).RowTop(DataGridAux(2).Row) + 5
    End If
    For I = 0 To txtAux.Count - 1
        txtAux(I).Text = ""
    Next I
    txtAux2(0).Text = ""
    txtAux2(4).Text = ""
    chkAux(0).Value = 0
    chkAux(1).Value = 0
    
    LLamaLineas 2, Modo, anc
    
    
    PonerFoco txtAux(0) '*** 1r camp visible que siga PK ***
    
    ' *** si n'hi han camps de descripció a la capçalera ***
    'PosarDescripcions

End Sub

Private Sub BotonModificar()
Dim I As Integer
Dim anc As Single

    PonerModo 4

    ' *** bloquejar els camps visibles de la clau primaria de la capçalera ***
    For I = 0 To 0
        BloquearTxt txtAux(I), True
    Next I
    
    
    If DataGridAux(2).Bookmark < DataGridAux(2).FirstRow Or DataGridAux(2).Bookmark > (DataGridAux(2).FirstRow + DataGridAux(2).VisibleRows - 1) Then
        I = DataGridAux(2).Bookmark - DataGridAux(2).FirstRow
        DataGridAux(2).Scroll 0, I
        DataGridAux(2).Refresh
    End If
    
    If DataGridAux(2).Row < 0 Then
        anc = 320
    Else
        anc = DataGridAux(2).RowTop(DataGridAux(2).Row) + 545
    End If
    txtAux(0).Text = DataGridAux(2).Columns(0).Text
    txtAux2(0).Text = DataGridAux(2).Columns(1).Text
    txtAux(1).Text = DataGridAux(2).Columns(2).Text
    txtAux(2).Text = DataGridAux(2).Columns(3).Text
    txtAux(3).Text = DataGridAux(2).Columns(4).Text
    txtAux(7).Text = DataGridAux(2).Columns(5).Text
    txtAux(4).Text = DataGridAux(2).Columns(6).Text
    txtAux2(4).Text = DataGridAux(2).Columns(7).Text
    
    Me.chkAux(0).Value = AdoAux(2).Recordset!procesado
    Me.chkAux(1).Value = AdoAux(2).Recordset!sinacabar
     
    txtAux(5).Text = DataGridAux(2).Columns(10).Text
    txtAux(6).Text = DataGridAux(2).Columns(11).Text
     
    FecIni = txtAux(5).Text
    FecFin = txtAux(6).Text
     
    LLamaLineas 2, 4, anc
    
    ' *** foco al 1r camp visible que NO siga clau primaria ***
'    PonerFoco txtAux(7)
    PonerFoco txtAux(1)

End Sub

Private Sub BotonEliminar()
Dim cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If AdoAux(2).Recordset.EOF Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(adoaux(2).Recordset.Fields(0).Value), FormatoCampo(txtaux(0))) Then Exit Sub
    ' ***************************************************************************

    ' *************** canviar la pregunta ****************
    cad = "¿Seguro que desea eliminar el Registro del Trabajador?"
    cad = cad & vbCrLf & "Trabajador: " & AdoAux(2).Recordset.Fields(0)
    cad = cad & vbCrLf & "Fecha Inicio: " & AdoAux(2).Recordset.Fields(10)
    cad = cad & vbCrLf & "Fecha Fin: " & AdoAux(2).Recordset.Fields(11)
    
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = AdoAux(2).Recordset.AbsolutePosition
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf SituarDataTrasEliminar(AdoAux(2), NumRegElim, True) Then
'            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Coste diario", Err.Description
End Sub

Private Sub PonerCampos()
Dim I As Integer
Dim codPobla As String, desPobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If AdoAux(2).Recordset.EOF Then Exit Sub
    
    CargaGrid
    
    ' ************* configurar els camps de les descripcions de la capçalera *************
'    text2(3).Text = PonerNombreDeCod(txtaux(3), "variedades", "nomvarie")
'    text2(4).Text = PonerNombreDeCod(txtaux(4), "forfaits", "nomconfe")
    ' ********************************************************************************
    
'    CalcularTotales
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = AdoAux(2).Recordset.AbsolutePosition & " de " & AdoAux(2).Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu
End Sub

Private Sub cmdCancelar_Click()
Dim I As Integer
Dim V

  Select Case Modo
        Case 1 'búsqueda
            CargaGrid
        Case 3 'insertar
            Me.DataGridAux(2).AllowAddNew = False
            'CargaGrid
            If Not Me.AdoAux(2).Recordset.EOF Then AdoAux(2).Recordset.MoveFirst
        Case 4 'modificar
            TerminaBloquear
    End Select
    
    PonerModo 2

'    PonerFocoGrid Me.DataGridAux(2)
    If Err.Number <> 0 Then Err.Clear

End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Sql As String

    On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm2(Me, 1)
    If Not b Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
    If (Modo = 3) Then 'insertar
        'comprobar si existe ya el cod. del campo clave primaria
'        Sql = DevuelveDesdeBDNew(cAgro, "cccabdia", "fecha", "fecha", txtAux(0).Text, "F", , "codcoste", txtAux(1).Text, "N")
'        If Sql <> "" Then
'            MsgBox "Ya existe el concepto de coste para esta fecha. Modifique.", vbExclamation
'            b = False
'        End If
    End If
    
    If b And Modo = 4 Then
        ' si me han cambiado la fecha de inicio o la fecha de fin comprobamos que no exista ese registro
        If CDate(txtAux(5).Text) <> FecIni Or CDate(txtAux(6).Text) <> FecFin Then
            Sql = "select count(*) from cctrabaconf where codtraba =" & DBSet(txtAux(0).Text, "N")
            Sql = Sql & " and fechaini = " & DBSet(txtAux(5).Text, "FH")
            Sql = Sql & " and fechafin = " & DBSet(txtAux(6).Text, "FH")
            If DevuelveValor(Sql) <> 0 Then
                MsgBox "Ya existe este registro. Revise.", vbExclamation
                b = False
            End If
        End If
    End If
    
'    If b And (Modo = 3 Or Modo = 4) Then
'        ' vemos si se solapa con otro registro
'        If SeSolapaConOtroRegistro Then
'            MsgBox "Se solapa con otro registro del trabajador. Revise.", vbExclamation
'            b = False
'        End If
'    End If
    
    ' ************************************************************************************
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function SeSolapaConOtroRegistro() As Boolean
Dim Sql As String

    On Error Resume Next
    
    SeSolapaConOtroRegistro = False
    
    Sql = "select count(*) from cctrabaconf where codtraba = " & DBSet(txtAux(0).Text, "N") & " and "
    
    If Modo = 3 Then
        Sql = Sql & " ((fechaini < " & DBSet(txtAux(5).Text, "FH") & " and " & DBSet(txtAux(5).Text, "FH") & " < fechafin) or "
        Sql = Sql & "  (fechaini < " & DBSet(txtAux(6).Text, "FH") & " and " & DBSet(txtAux(6).Text, "FH") & " < fechafin))  "
    Else
        Sql = Sql & " ((fechaini < " & DBSet(FecIni, "FH") & " and " & DBSet(FecIni, "FH") & " < fechafin) or "
        Sql = Sql & "  (fechaini < " & DBSet(FecFin, "FH") & " and " & DBSet(FecFin, "FH") & " < fechafin))  "
    End If
    SeSolapaConOtroRegistro = (TotalRegistros(Sql) <> 0)
    
End Function


Private Sub PosicionarData()
Dim cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    cad = "(codtraba = " & DBSet(txtAux(0).Text, "N") & " and fechaini=" & DBSet(txtAux(5).Text, "FH") & " and fechafin = " & DBSet(txtAux(6).Text, "FH") & " )"
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(adoaux(2), cad, Indicador) Then
    If SituarDataMULTI(AdoAux(2), cad, Indicador, True) Then
        If ModoLineas <> 1 Then PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
End Sub

Private Function Eliminar() As Boolean
Dim vWhere As String
Dim vTipoMov As CTiposMov

    On Error GoTo FinEliminar

    ' ***** canviar el nom de la PK de la capçalera, repasar codEmpre *******
    vWhere = " WHERE codtraba=" & AdoAux(2).Recordset!codtraba & " and fechaini = " & DBSet(AdoAux(2).Recordset!FechaIni, "FH")
    vWhere = vWhere & " and fechafin = " & DBSet(AdoAux(2).Recordset!FechaFin, "FH")
        
    conn.Execute "Delete from " & NombreTabla & vWhere
       
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        Eliminar = False
    Else
        Eliminar = True
    End If
    CargaGrid
End Function

Private Sub txtAux_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco txtAux(Index), Modo
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
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




Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim b As Boolean

    DeseleccionaGrid Me.DataGridAux(Index)
    PonerModo xModo
    
    'Fijamos el ancho
    For jj = 0 To txtAux.Count - 1
        txtAux(jj).Top = alto
    Next jj
    
    ' ### [Monica] 12/09/2006
    txtAux2(0).Top = alto
    txtAux2(4).Top = alto
    
    btnBuscar(0).Top = alto - 15
    btnBuscar(1).Top = alto - 15
    btnBuscar(2).Top = alto - 15
    
    Me.chkAux(0).Top = alto
    Me.chkAux(1).Top = alto

End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
Dim Sql As String
    
    
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 1 'fecha de coste
            PonerFormatoFecha txtAux(Index)
    
        Case 4 'concepto de coste
            If Modo = 1 Then Exit Sub
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), "ccconcostes", "nomcoste")
                If txtAux2(Index).Text = "" Then
                    cadMen = "No existe el Concepto de Coste: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCon = New frmCCManConcep
                        frmCon.DatosADevolverBusqueda = "0|1|"
                        frmCon.NuevoCodigo = txtAux(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmCon.Show vbModal
                        Set frmCon = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, AdoAux(2), 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                End If
            Else
                txtAux2(Index).Text = ""
            End If
    
        Case 0 ' trabajador
            If Modo = 1 Then Exit Sub
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(0).Text = PonerNombreDeCod(txtAux(Index), "straba", "nomtraba")
                If txtAux2(0).Text = "" Then
                    cadMen = "No existe el Trabajador: " & txtAux(Index).Text & ". Reintroduzca." & vbCrLf
                    MsgBox cadMen, vbExclamation
                    PonerFoco txtAux(Index)
                End If
            Else
                txtAux2(0).Text = ""
            End If
        
        Case 2, 3 ' hora inicio y hora fin
            PonerFormatoHora txtAux(Index)
        
        Case 7 ' linea de coste
            If Modo = 1 Then Exit Sub
            If PonerFormatoEntero(txtAux(Index)) Then
                Sql = DevuelveDesdeBDNew(cAgro, "cclinconf", "nomlinconf", "codlinconf", txtAux(7).Text, "N")
                If Sql = "" Then
                    cadMen = "No existe la Línea de Coste: " & txtAux(Index).Text & vbCrLf
                    MsgBox cadMen, vbInformation
                    PonerFoco txtAux(Index)
                End If
            End If
    End Select
    
    If Index = 1 Or Index = 2 Then
        txtAux(5).Text = Trim(Format(txtAux(1).Text, "yyyy-mm-dd") & " " & Format(txtAux(2).Text, "hh:mm:ss"))
    End If
    
    If Index = 1 Or Index = 3 Then
        txtAux(6).Text = Trim(Format(txtAux(1).Text, "yyyy-mm-dd") & " " & Format(txtAux(3).Text, "hh:mm:ss"))
    End If

End Sub



Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not txtAux(Index).MultiLine Then
        If KeyAscii = teclaBuscar Then
            If Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) Then
                Select Case Index
                    Case 1: 'articulo
                        KeyAscii = 0
                        btnBuscar_Click (0)
                    Case 9: 'coste
                        KeyAscii = 0
                        btnBuscar_Click (1)
                End Select
            End If
        Else
            KEYpress KeyAscii
        End If
    End If
End Sub

Private Function DatosOkLlin(nomframe As String) As Boolean
Dim Rs As adodb.Recordset
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
    
    If b And nomframe = "FrameAux0" Then
        If txtAux(4).Text = "" Then
            MsgBox "El valor de Fecha Inicio no puede ser nulo.", vbExclamation
            b = False
        End If
        If b And txtAux(5).Text = "" Then
            MsgBox "El valor de Hora Inicio no puede ser nulo.", vbExclamation
            b = False
        End If
        If b And txtAux(6).Text = "" Then
            MsgBox "El valor de Fecha Fin no puede ser nulo.", vbExclamation
            b = False
        End If
        If b And txtAux(7).Text = "" Then
            MsgBox "El valor de Hora Fin no puede ser nulo.", vbExclamation
            b = False
        End If
    End If
    
    If b And nomframe = "FrameAux1" Then
        If Modo = 3 Or Modo = 4 Then
            Sql = "select count(*) from cclindia2 where codorden= " & DBSet(txtAux(0).Text, "N")
            Sql = Sql & " and codzona = " & DBSet(txtAux(42).Text, "N")
            Sql = Sql & " and codcateg = " & DBSet(txtAux(43).Text, "N")
            Sql = Sql & " and numlinea <> " & DBSet(txtAux(1).Text, "N")
            
            If TotalRegistros(Sql) <> 0 Then
                MsgBox "Existe un registro en esta orden para esta zona y categoria. Revise.", vbExclamation
                b = False
            End If
        End If
    End If
    
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
Dim I As Byte

    If Index = 2 Then
        If Modo = 3 Then
        Else
            If DataGridAux(Index).Columns.Count > 1 Then
'               PonerCampos
                lblIndicador.Caption = AdoAux(2).Recordset.AbsolutePosition & " de " & AdoAux(2).Recordset.RecordCount

                PonerModoOpcionesMenu (Modo)
                PonerOpcionesMenu

            End If
        End If
    End If

End Sub

' ***** si n'hi han varios nivells de tabs *****
'Private Sub SituarTab(numTab As Integer)
'    On Error Resume Next
'
'    SSTab1.Tab = numTab
'
'    If Err.Number <> 0 Then Err.Clear
'End Sub
' **********************************************

Private Sub CargaFrame(Index As Integer, enlaza As Boolean)
Dim tip As Integer
Dim I As Byte

    AdoAux(Index).ConnectionString = conn
    AdoAux(Index).RecordSource = MontaSQLCarga(Index, enlaza)
    AdoAux(Index).CursorType = adOpenDynamic
    AdoAux(Index).LockType = adLockPessimistic
    AdoAux(Index).Refresh
    
    If Not AdoAux(Index).Recordset.EOF Then
        PonerCamposForma2 Me, AdoAux(Index), 2, "FrameAux" & Index
    Else
        ' *** si n'hi han tabs sense datagrids, li pose els valors als camps ***
        NetejaFrameAux "FrameAux3" 'neteja només lo que te TAG
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

Private Sub CargaGrid(Optional vSQL As String)
    Dim Sql As String
    Dim tots As String
    Dim b As Boolean
    
    txtAux(2).Tag = "Hora I|FHH|N|||cctrabaconf|fechaini|hh:mm:ss||"
    txtAux(3).Tag = "Hora F|FHH|N|||cctrabaconf|fechafin|hh:mm:ss||"
    
'    adodc1.ConnectionString = Conn
    If vSQL <> "" Then
        Sql = CadenaConsulta & " AND " & vSQL
    Else
        Sql = CadenaConsulta
    End If
    '********************* canviar el ORDER BY *********************++
    Sql = Sql & " ORDER BY cctrabaconf.codtraba, cctrabaconf.fechaini"
    '**************************************************************++
    
    CargaGridGnral Me.DataGridAux(2), Me.AdoAux(2), Sql, PrimeraVez
    
    
    ' *******************canviar els noms i si fa falta la cantitat********************
    tots = "S|txtAux(0)|T|Código|1200|;S|btnBuscar(2)|B||195|;S|txtAux2(0)|T|Trabajador|3300|;"
    tots = tots & "S|txtAux(1)|T|Fecha|1100|;S|btnBuscar(0)|B||195|;S|txtAux(2)|T|Hora I|900|;S|txtAux(3)|T|Hora F|900|;S|txtAux(7)|T|Línea|900|;"
    tots = tots & "S|txtAux(4)|T|Concepto|900|;S|btnBuscar(1)|B||195|;S|txtAux2(4)|T|Descripcion|3000|;"
    tots = tots & "N||||0|;S|chkAux(0)|CB|Pr|360|;N||||0|;N||||0|;N||||0|;S|chkAux(1)|CB|SA|360|;"
    
    arregla tots, DataGridAux(2), Me
    
    DataGridAux(2).ScrollBars = dbgAutomatic
    
    DataGridAux(2).Columns(1).Alignment = dbgLeft
    DataGridAux(2).Columns(2).Alignment = dbgLeft
    DataGridAux(2).Columns(3).Alignment = dbgLeft
    
    DataGridAux(2).Columns(5).Alignment = dbgLeft
    DataGridAux(2).Columns(6).Alignment = dbgLeft
    
    DataGridAux(2).Columns(8).Alignment = dbgCenter
    
    
    b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
        
End Sub



Private Sub CargaGrid2(Index As Integer, enlaza As Boolean)
Dim b As Boolean
Dim I As Byte
Dim tots As String

'    On Error GoTo ECarga
'
'    tots = MontaSQLCarga(Index, enlaza)
'
'    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
'
'    Select Case Index
'        Case 0 'trabajadores
'            txtAux(2).Tag = "Fecha Ini|FHH|N|||cctrabaconf|fechaini|hh:mm:ss||"
'            txtAux(3).Tag = "Fecha fin|FHH|N|||cctrabaconf|fechafin|hh:mm:ss||"
'
'            'si es visible|control|tipo campo|nombre campo|ancho control|
'            tots = "" 'codtraba, fechainicio, fechafin
'            tots = tots & "S|txtAux(3)|T|Código|800|;S|btnBuscar(2)|B|||;"
'            tots = tots & "S|txtAux2(3)|T|Trabajador|3300|;S|txtAux(2)|T|H.Inicio|800|;"
'            tots = tots & "S|txtAux(5)|T|H.Fin|800|;S|txtAux(5)|T|H.Fin|800|;S|txtAux(6)|T|Horas|1000|;N||||0|;N||||0|;"
'
'            arregla tots, DataGridAux(Index), Me
'
''            DataGridAux(0).Columns(6).NumberFormat = "dd/mm/yyyy"
''            DataGridAux(0).Columns(8).NumberFormat = "dd/mm/yyyy"
'
'            DataGridAux(0).Columns(2).Alignment = dbgLeft
'            DataGridAux(0).Columns(3).Alignment = dbgLeft
'            DataGridAux(0).Columns(4).Alignment = dbgLeft
'            DataGridAux(0).Columns(5).Alignment = dbgLeft
'
'            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
'
'            txtAux(4).Tag = ""
'            txtAux(5).Tag = ""
'
'    End Select
'
'    DataGridAux(Index).ScrollBars = dbgAutomatic
'
''    PonerModoOpcionesMenu Modo
'
'    ' **** si n'hi han llínies en grids i camps fora d'estos ****
''    If Not AdoAux(Index).Recordset.EOF Then
''        DataGridAux_RowColChange Index, 1, 1
''    Else
'''        LimpiarCamposFrame Index
''    End If
'ECarga:
'    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub

Private Sub ModificarCategorias(Eliminar As Boolean)
Dim Sql As String
Dim Sql2 As String
Dim Categoria As Integer
Dim NumF As String
Dim Horas As String
Dim Rs As adodb.Recordset

    Sql = "select * from cclindia1 where fecha = " & DBSet(txtAux(0).Text, "F")
    Sql = Sql & " and codcoste = " & DBSet(txtAux(1).Text, "N")
    
    Set Rs = New adodb.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    While Not Rs.EOF
    
        Sql = "select codcateg from straba where codtraba = " & DBSet(Rs!codtraba, "N")
        Categoria = DevuelveValor(Sql)
    
        Sql = "select count(*) from cclindia2 where fecha = " & DBSet(txtAux(0).Text, "F")
        Sql = Sql & " and codcoste = " & DBSet(Rs!codCoste, "N")
        Sql = Sql & " and codcateg = " & DBSet(Categoria, "N")
        
        If TotalRegistros(Sql) = 0 Then
            NumF = SugerirCodigoSiguienteStr("cclindia2", "numlinea", "fecha = " & DBSet(txtAux(0).Text, "F") & " and codcoste = " & DBSet(txtAux(1).Text, "N"))
        
            Sql2 = "insert into cclindia2 (fecha,codcoste,numlinea,codcateg,horas) values ("
            Sql2 = Sql2 & DBSet(txtAux(0).Text, "F") & "," & DBSet(txtAux(1).Text, "N") & "," & DBSet(NumF, "N") & ","
            Sql2 = Sql2 & DBSet(Categoria, "N") & "," & DBSet(Rs!Horas, "N") & ")"
            
            conn.Execute Sql2
        
        Else
            Sql2 = "select sum(horas) from cclindia1 where fecha = " & DBSet(txtAux(0).Text, "F")
            Sql2 = Sql2 & " and codcoste = " & DBSet(txtAux(1).Text, "N")
            Sql2 = Sql2 & " and codtraba in ( select codtraba from straba where codcateg = " & DBSet(Categoria, "N") & ")"
            
            Horas = DevuelveValor(Sql2)
        
            Sql2 = "update cclindia2 set horas = " & DBSet(ImporteSinFormato(Horas), "N")
            Sql2 = Sql2 & " where fecha = " & DBSet(txtAux(0).Text, "F")
            Sql2 = Sql2 & " and codcoste = " & DBSet(txtAux(1).Text, "N")
            Sql2 = Sql2 & " and codcateg = " & DBSet(Categoria, "N")
                
            conn.Execute Sql2
                
            Sql2 = "delete from cclindia2 where fecha = " & DBSet(txtAux(0).Text, "F")
            Sql2 = Sql2 & " and codcoste = " & DBSet(txtAux(1).Text, "N")
            Sql2 = Sql2 & " and horas = 0"
            
            conn.Execute Sql2
        End If
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
End Sub



Private Function ModificarLinea() As Boolean
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim V As Integer
    
    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'trabajadores
        Case 1: nomframe = "FrameAux1" 'categorias
    End Select
    ModificarLinea = False
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
            ' solo en el caso de que estemos en trabajadores y añadamos una nueva linea hemos de modificar las lineas de categoria
            If NumTabMto = 0 Then
                ModificarCategorias False
                CargaGrid2 1, True
            End If
            
            ModoLineas = 0
            
            Select Case NumTabMto
                Case 0
                    V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
                Case 1
                    V = AdoAux(NumTabMto).Recordset.Fields(2) 'el 2 es el nº de llinia
            End Select
            CargaGrid2 NumTabMto, True
            
            ' *** si n'hi han tabs ***
'            SituarTab (NumTabMto + 1)

            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
            PonerFocoGrid Me.DataGridAux(NumTabMto)
            AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
            
            LLamaLineas NumTabMto, 0
            ModificarLinea = True
        End If
    End If
End Function

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " fecha=" & DBSet(Me.AdoAux(2).Recordset!fecha, "F") & " and cccabdia.codcoste = " & Me.AdoAux(2).Recordset!codCoste
'    vWhere = vWhere & " fecha=" & DBSet(txtaux(0).Text, "F") & " and cccabdia.codcoste = " & DBSet(txtaux(1).Text, "N")
    ObtenerWhereCab = vWhere
End Function

'' *** neteja els camps dels tabs de grid que
''estan fora d'este, i els camps de descripció ***
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

'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGridAux_GotFocus(Index As Integer)
'  WheelHook DataGridAux(Index)
'End Sub
'Private Sub DataGridAux_LostFocus(Index As Integer)
'  WheelUnHook
'End Sub


