VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAlmTraspaso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Traspaso Almacenes"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   11715
   Icon            =   "frmAlmTraspaso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar ToolAux 
      Height          =   390
      Index           =   0
      Left            =   225
      TabIndex        =   31
      Top             =   1890
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
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Index           =   6
      Left            =   5115
      TabIndex        =   29
      Tag             =   "Hora|H|N|||scatra|hormovim|hh:mm:ss|N|"
      Text            =   "Text1"
      Top             =   735
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox chkImpresion 
      Caption         =   "Impreso"
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   4800
      TabIndex        =   28
      Tag             =   "Situaci�n Impresi�n|N|N|||scatra|situacio||N|"
      Top             =   735
      Width           =   855
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Left            =   2160
      TabIndex        =   26
      ToolTipText     =   "Buscar art�culo"
      Top             =   5040
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   3
      Left            =   6000
      MaxLength       =   50
      TabIndex        =   15
      Text            =   "observac"
      Top             =   5040
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   2
      Left            =   4920
      MaxLength       =   16
      TabIndex        =   14
      Text            =   "cantidad"
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   320
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   13
      Text            =   "nombre artic"
      Top             =   5040
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   1200
      MaxLength       =   16
      TabIndex        =   12
      Text            =   "codartic"
      Top             =   5040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8640
      TabIndex        =   6
      Top             =   5730
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9915
      TabIndex        =   7
      Top             =   5730
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   9900
      TabIndex        =   25
      Top             =   5715
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   240
      TabIndex        =   23
      Top             =   5580
      Width           =   3000
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
         Left            =   240
         TabIndex        =   24
         Top             =   180
         Width           =   2595
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   1
      Left            =   2560
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "Text2"
      Top             =   1575
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   0
      Left            =   2560
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "Text2"
      Top             =   1230
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   1155
      Index           =   5
      Left            =   6360
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   5
      Tag             =   "Observaciones|T|S|||scatra|observa1||N|"
      Text            =   "frmAlmTraspaso.frx":000C
      Top             =   860
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Index           =   4
      Left            =   6435
      MaxLength       =   4
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1530
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   3
      Tag             =   "Almacen Destino|N|N|0|999|scatra|almadest|000|N|"
      Text            =   "Text1"
      Top             =   1575
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   2
      Tag             =   "Almacen Origen|N|N|0|999|scatra|almaorig|000|N|"
      Text            =   "Text1"
      Top             =   1230
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   3360
      MaxLength       =   10
      TabIndex        =   1
      Tag             =   "Fecha|F|N|||scatra|fechatra|dd/mm/yyyy|N|"
      Text            =   "Text1"
      Top             =   690
      Width           =   1095
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver Todos"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Lineas"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Actualizar"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "�ltimo"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   315
         Left            =   6600
         TabIndex        =   22
         Top             =   135
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   8280
      Top             =   480
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
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FEF7E4&
      Height          =   315
      Index           =   0
      Left            =   1320
      MaxLength       =   7
      TabIndex        =   0
      Tag             =   "N� Traspaso|N|S|0||scatra|codtrasp|0000000|S|"
      Text            =   "Text1"
      Top             =   690
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   9720
      Top             =   480
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   3480
      TabIndex        =   27
      Top             =   5670
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAlmTraspaso.frx":0012
      Height          =   3240
      Left            =   240
      TabIndex        =   8
      Top             =   2340
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   5715
      _Version        =   393216
      AllowUpdate     =   0   'False
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      Caption         =   "Hora"
      Height          =   210
      Left            =   4725
      TabIndex        =   30
      Top             =   720
      Width           =   375
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   1500
      ToolTipText     =   "Buscar almacen"
      Top             =   1575
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   1500
      ToolTipText     =   "Buscar almacen"
      Top             =   1230
      Width           =   240
   End
   Begin VB.Image imgFec 
      Height          =   240
      Index           =   0
      Left            =   3045
      Picture         =   "frmAlmTraspaso.frx":0027
      ToolTipText     =   "Buscar fecha"
      Top             =   720
      Width           =   240
   End
   Begin VB.Label Label6 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   6360
      TabIndex        =   19
      Top             =   630
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Almacen Destino"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   1575
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Almacen Origen"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   1230
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   2535
      TabIndex        =   16
      Top             =   690
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "N� Traspaso"
      Height          =   255
      Left            =   225
      TabIndex        =   11
      Top             =   690
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Cargando datos ........."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   8220
      Visible         =   0   'False
      Width           =   3495
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
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmAlmTraspaso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public EsHistorico As Boolean 'Si es true abrir el formulario con la tabla de
                              'historico schmov, y solo en modo de consulta

'Si se llama de la busqueda en el frmAlmMovimArticulos se accede
'a las tablas del hist�rico del traspaso seleccionado (solo consulta)
Public hcoCodMovim As Long 'cod. traspaso del historico
Public hcoFechaMovim As Date 'Fecha del historico

'--------------------------------------------------------------------------

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'Calendario de Fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmA As frmManAlmProp 'Almacen Origen/Destino
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmArt As frmManArtic   'Form Articulos
Attribute frmArt.VB_VarHelpID = -1


Dim NombreTabla As String
Dim NomTablaLineas As String
Dim Ordenacion As String

Private Modo As Byte
Private ModoAnterior As Byte
Dim kCampo As Integer

Dim btnAnyadir As Byte
'Variable que indica el n�mero del Boton  Anyadir en la Toolbar1
Dim btnPrimero As Byte
'Variable que indica el n�mero del Boton  PrimerRegistro en la Toolbar1

Dim ModificaLineas As Byte
'1.- A�adir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim CodTipoMov As String
'Codigo tipo de movimiento en funci�n del valor en la tabla de par�metros: stipom

Dim CadenaConsulta As String
Dim cadSeleccion As String 'Cadena de seleccion para FormulaSelection del Informe

Private HaDevueltoDatos As Boolean



Private Sub chkVistaPrevia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    Select Case Modo
    Case 1 'BUSQUEDA
        Text1(kCampo).BackColor = vbWhite
        cadSeleccion = ""
        HacerBusqueda
        
    Case 3 'INSERTAR
        If DatosOk(True) Then InsertarCabecera
        
    Case 4 'MODIFICAR
        If DatosOk(True) Then
             If ModificaDesdeFormulario(Me) Then
                 TerminaBloquear
                 PosicionarData
             End If
         End If
            
    Case 5 'LINEAS Traspaso Almacenes
        If InsertarModificarLinea Then
            'Reestablecemos los campos
            'y ponemos el grid
            DataGrid1.AllowAddNew = False
            If ModificaLineas = 1 Then 'Insertar
                CargaGrid True
                ModificaLineas = 0
                BotonAnyadirLineas
            ElseIf ModificaLineas = 2 Then 'Modificar
                TerminaBloquear
'                Data2.Recordset.Find (Data2.Recordset.Fields(1).Name & " =" & CInt(Me.cmdAceptar.Tag))
                ModificaLineas = 0
'--monica:rollo
'                PonerBotonCabecera True
                CargaTxtAux False, False
                Me.lblIndicador.Caption = ""
'++monica:rollo
                CargaGrid True
                Data2.Recordset.Find (Data2.Recordset.Fields(1).Name & " =" & CInt(Me.cmdAceptar.Tag))
                DataGrid1.Enabled = True
                DataGrid1.SetFocus
'++monica: rollo
                PonerModo 2
                PonerCampos
                
            End If
        End If
    End Select
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdAux_Click()
    Set frmArt = New frmManArtic
    frmArt.DatosADevolverBusqueda = "0|1|" 'Poner en Modo busqueda
    frmArt.Show vbModal
    Set frmArt = Nothing
    PonerFoco txtAux(0)
End Sub

Private Sub cmdCancelar_Click()
On Error GoTo ECancelar
    Select Case Modo
        Case 1 'Buscar
            LimpiarCampos
            PonerModo 0
        Case 3 'Insertar
            If ModoAnterior = 0 Then
                LimpiarCampos
                PonerModo 0
            Else
                PonerModo 2
                PonerCampos
            End If
                
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
            
        Case 5 'Mantenimiento Lineas traspasos
            CargaTxtAux False, False
            DataGrid1.Enabled = True
            DataGrid1.AllowAddNew = False
            If Not ModificaLineas = 2 Then 'Modificar
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            End If
            ModificaLineas = 0
'--monica:rollo toolbar
'             PonerBotonCabecera True
            DataGrid1.Refresh
            PonerFocoBtn Me.cmdRegresar
            PonerModo 2
            
            PonerCampos
            
    End Select
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdRegresar_Click()
'Este es el boton Cabecera

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then 'modo 5: Mantenimiento Lineas
        PonerModo 2
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid Me.DataGrid1
            DataGrid1.Bookmark = 1
        End If
        Me.cmdRegresar.visible = False
    End If
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    If Modo = 5 And KeyAscii = 27 Then 'ESC 'Modo Lineas
        cmdRegresar_Click
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim I As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    'ICONOS de La toolbar
    btnAnyadir = 5 'Posicion del boton A�adir en la toolbar1
    btnPrimero = 15 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
    With Toolbar1
        .ImageList = frmPpal.imgListComun
        .DisabledImageList = frmPpal.imgListComun_BN
        'ASignamos botones
        .Buttons(1).Image = 1   'Buscar
        .Buttons(2).Image = 2 'Ver Todos
        .Buttons(5).Image = 3 'A�adir
        .Buttons(6).Image = 4 'Modificar
        .Buttons(7).Image = 5 'Eliminar
        .Buttons(9).Image = 21 'Mantenimiento L�neas
        .Buttons(10).Image = 16 'Actualizar
        .Buttons(12).Image = 10 'Imprimir
        .Buttons(13).Image = 11 'Salir
        .Buttons(btnPrimero).Image = 6 'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Ultimo
    End With
    ' ******* si n'hi han ll�nies *******
    'ICONETS DE LES BARRES ALS TABS DE LL�NIA
    For I = 0 To ToolAux.Count - 1
        With Me.ToolAux(I)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
        End With
    Next I
    ' ***********************************
    
    'cargar IMAGES de busqueda
    For I = 0 To imgBuscar.Count - 1
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
    
    LimpiarCampos   'Limpia los campos TextBox
       
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    CodTipoMov = "TRA"
    'campo situacio solo en tabla scatra
    Me.chkImpresion(0).visible = Not EsHistorico
    'campo Hora solo en tabla hist. schtra
    Me.Label7.visible = EsHistorico
    Me.Text1(6).visible = EsHistorico
    
    cadSeleccion = ""
    
     If Not EsHistorico Then
        NombreTabla = "scatra"
        NomTablaLineas = "slitra" 'Tabla lineas de Traspasos Almacen
        Me.Caption = "Traspaso de Almacen"
    Else
        NombreTabla = "schtra"
        NomTablaLineas = "slhtra"
        CargarTagsHco Me, "scatra", NombreTabla
        Me.Caption = "Hist�rico Traspaso de Almacen"
    End If
    
    Ordenacion = " ORDER BY codtrasp"
    CadenaConsulta = "Select * from " & NombreTabla
    If hcoCodMovim <> -1 Then
    'Se llama desde Dobleclick en frmAlmMovimArticulos
        CadenaConsulta = CadenaConsulta & " where codtrasp=" & hcoCodMovim & " and fechatra= '" & Format(hcoFechaMovim, "yyyy-mm-dd") & "'"
    Else
         CadenaConsulta = CadenaConsulta & " WHERE codtrasp = -1"
    End If
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    If Not Data1.Recordset.EOF Then 'Se llama desde DblClick frmAlmMovimArticulos
                                    'Se carga con el valor del registro del DblClick
        Data1.Recordset.MoveFirst
        Me.Text1(0).Text = Format(Data1.Recordset!codtrasp, "0000000")
        Me.Text1(1).Text = Data1.Recordset!fechatra
        Me.Text1(6).Text = Format(Data1.Recordset!hormovim, "hh:mm:ss")
        'Almacen Origen
        Me.Text1(2).Text = Format(Data1.Recordset!almaorig, "000")
        Me.Text2(0).Text = PonerNombreDeCod(Text1(2), "salmpr", "nomalmac", "codalmac")
        'Almacen Destino
        Me.Text1(3).Text = Format(Data1.Recordset!almadest, "000")
        Me.Text2(1).Text = PonerNombreDeCod(Text1(3), "salmpr", "nomalmac", "codalmac")
        'Cod. Trabajador
        Me.Text1(4).Text = Format(Data1.Recordset!codtraba, "0000")
        Me.Text2(2).Text = PonerNombreDeCod(Text1(4), "straba", "nomtraba")
        Text1(5).Text = Data1.Recordset!observa1
        CargaGrid True
        Toolbar1.Buttons(5).Enabled = True 'Imprimir
    Else
        CargaGrid False '(Modo = 2) 'False
    End If
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim b As Boolean
Dim I As Byte
Dim SQL As String

    On Error GoTo ECarga

    b = DataGrid1.Enabled
    
    SQL = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data2, SQL, False
      
    DataGrid1.Columns(0).visible = False 'Cod. trasp
    DataGrid1.Columns(1).visible = False 'Numlinea
    
    I = 2
    'Cod. Art�culo
    DataGrid1.Columns(I).Caption = "Cod. Articulo"
    DataGrid1.Columns(I).Width = 1800
    
    'Nombre Art�culo
    I = I + 1
    DataGrid1.Columns(I).Caption = "Nombre Articulo"
    DataGrid1.Columns(I).Width = 3200
    
    'Cantidad
    I = I + 1
    DataGrid1.Columns(I).Caption = "Cantidad"
    DataGrid1.Columns(I).Width = 1400
    DataGrid1.Columns(I).Alignment = dbgRight
    DataGrid1.Columns(I).NumberFormat = FormatoImporte & " "
    
    'Observaciones
    I = I + 1
    DataGrid1.Columns(I).Caption = "Observaciones"
    DataGrid1.Columns(I).Width = 4300
       
    For I = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(I).AllowSizing = False
    Next I

'--monica: rollo toolbar
'    DataGrid1.Enabled = b
    DataGrid1.ScrollBars = dbgAutomatic
    
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posici�n adecuada
'    limpiar: si es true vaciar los txtAux
Dim I As Byte
Dim alto As Single

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For I = 0 To txtAux.Count - 1
            txtAux(I).Top = 290
        Next I
        Me.cmdAux.Top = 290
    Else
        DeseleccionaGrid Me.DataGrid1
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            For I = 0 To txtAux.Count - 1
                txtAux(I).Text = ""
            Next I
        End If
        
        If ModificaLineas = 1 Then 'Insertar
            For I = 0 To txtAux.Count - 1
'                If i <> 1 Then txtAux(i).Locked = False
                'LAURA 19/10/2006
                If I <> 1 Then BloquearTxt txtAux(I), False
            Next I
            cmdAux.Enabled = True
        ElseIf ModificaLineas = 2 Then
            'Poner valor a los txtAux
            For I = 0 To txtAux.Count - 1
                txtAux(I).Text = DataGrid1.Columns(I + 2).Text
            Next I
            BloquearTxt txtAux(0), True
            cmdAux.Enabled = False
            BloquearTxt txtAux(2), False
            BloquearTxt txtAux(3), False
        End If
        
        If DataGrid1.Row < 0 Then
            alto = DataGrid1.Top + 220
        Else
            alto = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + 10
        End If
        
        
        'Fijamos altura y posici�n Top
        For I = 0 To txtAux.Count - 1
            txtAux(I).Top = alto
            txtAux(I).Height = DataGrid1.RowHeight
        Next I
        Me.cmdAux.Top = alto
        Me.cmdAux.Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        txtAux(0).Left = DataGrid1.Left + 340 'codartic
        txtAux(0).Width = DataGrid1.Columns(2).Width - 200
        cmdAux.Left = txtAux(0).Left + txtAux(0).Width
        txtAux(1).Left = cmdAux.Left + cmdAux.Width + 10 'Nom artic
        txtAux(1).Width = DataGrid1.Columns(3).Width - 25
        For I = 2 To txtAux.Count - 1 'Cantidad y Observacion
            txtAux(I).Left = txtAux(I - 1).Left + txtAux(I - 1).Width + 25
            txtAux(I).Width = DataGrid1.Columns(I + 2).Width - 35
        Next I
    End If

    'Los ponemos Visibles o No
    For I = 0 To txtAux.Count - 1
        txtAux(I).visible = visible
    Next I
    cmdAux.visible = visible
End Sub



Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Almacenes Propios
Dim Indice As Byte
    Indice = CByte(Me.imgBuscar(0).Tag)
    Text1(Indice + 2).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Articulos
    txtAux(0).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Artic
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda
Dim cadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        
        If Modo <> 5 Then 'Estamos en Cabecera
            'Recupera todo el registro de Traspaso Almacenes
            'Sabemos que campos son los que nos devuelve
            'Creamos una cadena consulta y ponemos los datos
            cadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            cadB = Aux
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
        Else 'Estamos en Lineas
            'Llamamos desde el boton auxiliar de Art�culos
            txtAux(0).Text = RecuperaValor(CadenaDevuelta, 1)
            txtAux(1).Text = RecuperaValor(CadenaDevuelta, 2)
            PonerFoco txtAux(2)
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmC_Selec(vFecha As Date)
'Calendario de Fecha
    Text1(1).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgBuscar_Click(Index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
 
    Screen.MousePointer = vbHourglass
    imgBuscar(0).Tag = Index
    
    Select Case Index
        Case 0, 1 'Codigo Almacen Origen/Destino
            Set frmA = New frmManAlmProp
            frmA.DatosADevolverBusqueda = "0"
            frmA.Show vbModal
            Set frmA = Nothing
    End Select
    
    PonerFoco Text1(Index + 2)
    Screen.MousePointer = vbDefault
End Sub

Private Sub imgFec_Click(Index As Integer)
Dim Indice As Byte

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
    
    menu = Me.Height - Me.ScaleHeight 'ac� tinc el heigth del men� i de la toolbar

    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40

    Select Case Index
        Case 0
            imgFec(0).Tag = 1 '<===
        
            ' *** repasar si el camp es Text3 o Text1 ***
            If Text1(1).Text <> "" Then frmC.NovaData = Text1(1).Text
            ' ********************************************
    End Select
    

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es Text3 o Text1 ***
    PonerFoco Text1(CByte(imgFec(0).Tag)) '<===
    ' ********************************************
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    If Modo = 5 Then   'Eliminar lineas Traspaso Almacenes
        BotonEliminarLinea
    Else 'Eliminar Cabecera Traspaso Almacenes
        BotonEliminar
    End If
End Sub

Private Sub mnModificar_Click()
    If Modo = 5 Then  'Modificar lineas Traspaso Almacenes
        BotonModificarLinea
    Else 'Modificar Cabecera Traspaso Almacenes
        If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub

Private Sub mnNuevo_Click()
    If Modo = 5 Then  'A�adir lineas Traspaso Almacenes
        BotonAnyadirLineas
    Else 'A�adir Cabecera Traspaso Almacenes
        BotonAnyadir
    End If
End Sub

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

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    If Index <> 5 Then ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 And Index = 5 And Modo = 1 Then
        PonerFocoBtn cmdAceptar
    Else
        KEYpress KeyAscii
    End If
End Sub


Private Sub text1_LostFocus(Index As Integer)
    
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    'Bloquear el contador si no estamos en busquedas
    If (Modo <> 1) And (Index = 0) Then BloquearTxt Text1(0), True, True
    
    Select Case Index
        Case 0 'Codigo Traspaso Almacen
            If Text1(Index).Text <> "" Then Text1(Index).Text = Format(Text1(Index).Text, "0000000")
        Case 1 'Fecha
            If Text1(Index).Text <> "" And Modo <> 1 Then PonerFormatoFecha Text1(Index)
        Case 2, 3 'Codigo Almacen Origen/Destino
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index - 2).Text = PonerNombreDeCod(Text1(Index), "salmpr", "nomalmac", "codalmac")
                'no existe el almacen
                If Text2(Index - 2).Text = "" Then PonerFoco Text1(Index)
            Else
                Text2(Index - 2).Text = ""
                
            End If
'        Case 4  'Codigo Trabajador
'            If PonerFormatoEntero(Text1(Index)) Then
'                Text2(Index - 2).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba")
'            Else
'                Text2(Index - 2).Text = ""
'            End If
        Case 5 'Observaciones
            If Text1(Index).Text <> "" Then Text1(Index).Text = QuitarCaracterEnter(Text1(Index).Text)
    End Select
End Sub


'++monica : rollo toolbar
Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
Dim vWhere As String

'-- pon el bloqueo aqui
    'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
    Select Case Button.Index
        Case 1
            BotonAnyadirLineas
        Case 2
            vWhere = ObtenerWhereCP(False) & " and numlinea=" & Me.Data2.Recordset.Fields(1)
            If BloqueaRegistro(NomTablaLineas, vWhere) Then BotonModificarLinea
        Case 3
            BotonEliminarLinea
        Case Else
    End Select
    'End If
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 3 And KeyCode = 40 Then
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 3 And KeyAscii = 13 Then
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYpress KeyAscii
    End If
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim devuelve As String
    
    'Quitar espacios en blanco por los lados
    txtAux(Index).Text = Trim(txtAux(Index).Text)
    
    Select Case Index
        Case 0 'Cod. Articulo
            If txtAux(Index).Text = "" Then
                txtAux(Index + 1).Text = ""
            ElseIf ModificaLineas = 1 Then 'Insertando linea
                'Comprobamos si ya existe una linea con el art�culo, solo si estamos insertando (ModificaLineas=1)
                'conAri: conexion a BD Ariges
                devuelve = DevuelveDesdeBDNew(cAgro, NomTablaLineas, "codtrasp", "codtrasp", Text1(0).Text, "N", , "codartic", txtAux(0).Text, "T")
                If devuelve <> "" Then
                    devuelve = "Ya hay una l�nea con ese Art�culo: " & vbCrLf
                    devuelve = devuelve & "Codigo: " & txtAux(0).Text & vbCrLf
                    MsgBox devuelve, vbExclamation
                    PonerFoco txtAux(Index)
                Else
                    PonerArticulo txtAux(0), txtAux(1), Text1(2).Text, CodTipoMov, ModificaLineas
                End If
            End If
            
        Case 2 'Cantidad (Comprobamos formato como si fuera un Importe)
            'Formato tipo 1: Decimal(12,2)
            If txtAux(Index) <> "" Then PonerFormatoDecimal txtAux(Index), 1
    End Select
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Busqueda
            mnBuscar_Click
        Case 2 'Ver Todos
            mnVerTodos_Click
            
        Case 5 'Nuevo
            mnNuevo_Click
        Case 6  'Modificar
            mnModificar_Click
        Case 7 'Eliminar
            mnEliminar_Click
        Case 9 'Mantenimiento Lineas
            BotonLineas
        Case 10 'Actualizar
            BotonActualizar
        Case 12 'Imprimir
            BotonImprimir
            
        Case 13  'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas de Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte
Dim b As Boolean
Dim Numreg As Byte

    'Actualiza Iconos Insertar,Modificar,Eliminar
'--monica: rollo toolbar
'   ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo

    'Modo 2. Hay datos y estamos visualizandolos
    '-------------------------------------------
    b = (Kmodo = 2)
    'Poner Flechas de desplazamiento visibles
    Numreg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then Numreg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, Numreg
    
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1
    BloquearText1 Me, Modo
              
    'Como el campo 0 es clave primaria, NO se puede modificar
    BloquearTxt Text1(0), (Modo <> 1), True
    
    'Modo 1:Busqueda / Modo 3: Insertar / Modo 4: Modificar
    '-------------------------------------------------------
'--monica:rollo toolbar
'    b = (Modo = 3 Or Modo = 4 Or Modo = 1)
'++monica:rollo toolbar
    b = (Modo <> 0 And Modo <> 2)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    For I = 0 To Me.imgFec.Count - 1
        Me.imgFec(I).Enabled = b
    Next I
    
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Enabled = b
    Next I
    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar opciones de menu seg�n MODO
    PonerOpcionesMenu   'Activar opciones de menu seg�n NIVEL
                        'de permisos del usuario
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean, bAux As Boolean
Dim I As Byte
    
    'Si visualizamos el historico no mostrar botones de Mantenimiento, solo es consulta
    For I = 5 To 10
        '++monica: rollo toolbar he puesto condicion
        If I <> 9 Then Toolbar1.Buttons(I).visible = Not EsHistorico
    Next I
    Me.mnNuevo.visible = Not EsHistorico
    Me.mnModificar.visible = Not EsHistorico
    Me.mnEliminar.visible = Not EsHistorico
    Me.mnBarra2.visible = Not EsHistorico
    
    If Not EsHistorico Then
         b = (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
        'Insertar
        Toolbar1.Buttons(5).Enabled = (b Or Modo = 0)
        Me.mnNuevo.Enabled = (b Or Modo = 0)
        'Modificar
        Toolbar1.Buttons(6).Enabled = b
        Me.mnModificar.Enabled = b
        'eliminar
        Toolbar1.Buttons(7).Enabled = b
        Me.mnEliminar.Enabled = b
        
        '--------------------------------
        b = (Modo = 2)
        
        'Lineas Traspaso Almacenes
'--monica: rollo
'        Toolbar1.Buttons(9).Enabled = b
        'Actualizar
        Toolbar1.Buttons(10).Enabled = b
        'Imprimir
        Toolbar1.Buttons(12).Enabled = b
            
        '-------------------------------
        b = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(1).Enabled = Not b
        Me.mnBuscar.Enabled = Not b
        'VerTodos
        Toolbar1.Buttons(2).Enabled = Not b
        Me.mnVerTodos.Enabled = Not b
    End If
    
    '++monica: rollo toolbar
    ' *** si n'hi han ll�nies que tenen grids (en o sense tab) ***
    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not EsHistorico
    For I = 0 To ToolAux.Count - 1
        ToolAux(I).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.Data2.Recordset.RecordCount > 0)
        ToolAux(I).Buttons(2).Enabled = bAux
        ToolAux(I).Buttons(3).Enabled = bAux
    Next I
    
    
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    Me.chkImpresion(0).Value = 0
    'Aqui va el especifico de cada form es
    '### a mano
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones(Flechas) de Desplazamiento de Registros de la Toolbar

    Select Case Modo
        Case 5 'Modo Mantenimiento de Almacenes (Lineas)
            If Data2.Recordset.EOF Then Exit Sub
            DesplazamientoData Data2, Index
        Case Else 'Datos de Cabecera
            If Data1.Recordset.EOF Then Exit Sub
            DesplazamientoData Data1, Index
            PonerCampos
    End Select
End Sub


Private Function MontaSQLCarga(enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Bas�ndose en la informaci�n proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
Dim Tabla As String
On Error GoTo EMontaSQL
 
    Tabla = NomTablaLineas

    SQL = "SELECT " & Tabla & ".codtrasp, "
    SQL = SQL & Tabla & ".numlinea, " & Tabla & ".codartic, Articulos.nomartic, "
    SQL = SQL & Tabla & ".cantidad, " & Tabla & ".observa2 "
    SQL = SQL & " FROM ((" & Tabla & " LEFT JOIN sartic AS Articulos ON " & Tabla & ".codartic ="
    SQL = SQL & " Articulos.codartic))"
    If enlaza Then
        SQL = SQL & ObtenerWhereCP(True)  '" WHERE codtrasp = " & Data1.Recordset!codtrasp
    Else
        SQL = SQL & " WHERE codtrasp = -1"
    End If
    SQL = SQL & " ORDER BY " & Tabla & ".numlinea"
    MontaSQLCarga = SQL
    
EMontaSQL:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub BotonBuscar()
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False

        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
        PonerFoco Text1(0)
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub BotonLineas()
On Error GoTo ErrorLineas

    Screen.MousePointer = vbHourglass
    PonerModo 5
    ModificaLineas = 0
    PonerBotonCabecera True
    CargaGrid True
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorLineas:
    If Err.Number <> 0 Then MuestraError Err.Number, "Lineas"
    Screen.MousePointer = vbDefault
End Sub


Private Sub BotonAnyadir()
Dim NomTraba As String

    LimpiarCampos 'Vac�a los TextBox
    
    'A�adiremos el boton de aceptar y demas objetos para insertar
    ModoAnterior = Modo 'Para el bot�n Cancelar en Modo Insertar
    PonerModo 3
        
    'Ponemos el grid lineas Traspaso enlazando a ningun sitio
    CargaGrid False
    
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    'Poner Trabajador por defecto el trabajador conectado
'    Text1(4).Text = PonerTrabajadorConectado(NomTraba)
'    Text2(2).Text = NomTraba
    
    PonerFoco Text1(1)
End Sub


Private Sub BotonAnyadirLineas()
Dim vWhere As String
    
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
    
    ModificaLineas = 1
    
    '++monica: rollo toolbar
    PonerModo (5)
    DataGrid1.Enabled = True
    Me.DataGrid1.SetFocus
    '++monica
    
    
    vWhere = ObtenerWhereCP(False)
    cmdAceptar.Tag = SugerirCodigoSiguienteStr("slitra", "numlinea", vWhere)

'--monica: rollo toolbar
'    PonerBotonCabecera False
'    lblIndicador.Caption = "INSERTAR"
    
    'Situamos el grid al final
    AnyadirLinea DataGrid1, Data2

    CargaTxtAux True, True
    PonerFoco txtAux(0)
End Sub


Private Sub BotonModificar()
    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    'Como el campo 0 es clave primaria, NO se puede modificar
    BloquearTxt Text1(0), True, True
    PonerFoco Text1(1)
End Sub

Private Sub BotonModificarLinea()
Dim I As Integer

    If Data2.Recordset.EOF Then Exit Sub
    If Data2.Recordset.RecordCount < 1 Then Exit Sub

    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub
    
    ModificaLineas = 2 'Modificar
    
    '++monica: rollo toolbar
    PonerModo (5)
    DataGrid1.Enabled = True
    Me.DataGrid1.SetFocus
    '++monica

    Screen.MousePointer = vbHourglass
    
'--monica: rollo toolbar
'    PonerBotonCabecera False
'    Me.lblIndicador.Caption = "MODIFICAR"
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    cmdAceptar.Tag = Data2.Recordset!numlinea

    CargaTxtAux True, False
    PonerFoco txtAux(2) 'Poner el foco
    Screen.MousePointer = vbDefault
    Me.DataGrid1.Enabled = False
End Sub


Private Sub BotonEliminar()
Dim SQL As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    SQL = "Cabecera de Traspaso Almacenes." & vbCrLf
    SQL = SQL & "------------------------------------------" & vbCrLf & vbCrLf
    
    SQL = SQL & "Va a eliminar el Traspaso:" & vbCrLf
    SQL = SQL & vbCrLf & "N� Traspaso   : " & Text1(0).Text
    SQL = SQL & vbCrLf & "Fecha Trasp.  : " & CStr(Data1.Recordset.Fields(1))
    SQL = SQL & vbCrLf & "Almac. Origen : " & Text1(2).Text
    SQL = SQL & vbCrLf & "Almac. Destino: " & Text1(3).Text
    SQL = SQL & vbCrLf & vbCrLf & " �Desea continuar ? "
    
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        If Not Eliminar Then Exit Sub
'
'        'Devolvemos contador, si no estamos actualizando
'        Set vTipoMov = New CTiposMov
'        NumRegElim = Data1.Recordset.Fields(0)
'        vTipoMov.DevolverContador CodTipoMov, NumRegElim
'        Set vTipoMov = Nothing
    
        NumRegElim = Data1.Recordset.AbsolutePosition
        DataGrid1.Enabled = False
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
'++monica: rollo
            PonerModo 2
            PonerCampos
        Else 'Solo habia un registro
            LimpiarCampos
            CargaGrid False
            PonerModo 0
        End If
    End If
     
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        MsgBox Err.Number & ": " & Err.Description, vbExclamation
        Data1.Recordset.CancelUpdate
    End If
End Sub


Private Function Eliminar() As Boolean
Dim SQL As String
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
On Error GoTo FinEliminar
    
    conn.BeginTrans
    SQL = ObtenerWhereCP(True)  '" WHERE  codtrasp=" & Data1.Recordset!codtrasp
    
    'Lineas
    conn.Execute "Delete  from " & NomTablaLineas & SQL
    
    'Cabeceras
    conn.Execute "Delete  from " & NombreTabla & SQL
                      
                      
  'Devolvemos contador, si no estamos actualizando
    Set vTipoMov = New CTiposMov
    vTipoMov.DevolverContador CodTipoMov, Data1.Recordset.Fields(0)
    Set vTipoMov = Nothing
                      
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


Private Sub BotonEliminarLinea()
Dim SQL As String
On Error GoTo Error2
    'Ciertas comprobaciones
    If Data2.Recordset.EOF Then Exit Sub

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar
       
    ModificaLineas = 3 'Eliminar
    
    '++monica: rollo toolbar
    PonerModo (5)
    DataGrid1.Enabled = True
    Me.DataGrid1.SetFocus
    '++monica
    
    '### a mano
    SQL = "Seguro que desea eliminar la l�nea del Art�culo:"
    SQL = SQL & vbCrLf & "C�digo: " & Data2.Recordset!codArtic
    SQL = SQL & vbCrLf & "Descripci�n: " & Data2.Recordset.Fields(3)
    
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        SQL = "Delete from slitra where codtrasp=" & Data2.Recordset!codtrasp
        SQL = SQL & " and numlinea=" & Data2.Recordset!numlinea
        SQL = SQL & " and codartic=" & DBSet(Data2.Recordset!codArtic, "T")
        
        '++ monica: rollo
        NumRegElim = Data2.Recordset.AbsolutePosition
        TerminaBloquear
        
        conn.Execute SQL
'--monica: rollo toolbar
'        CancelaADODC Me.Data2
        CargaGrid True
'--monica: rollo toolbar
'        CancelaADODC Me.Data2
    
'++monica: rollo
        If Not SituarDataTrasEliminar(Data2, NumRegElim, True) Then
'            PonerCampos
            
        End If
        PonerModo 2
    
    End If
    ModificaLineas = 0
Error2:
        Screen.MousePointer = vbDefault
        ModificaLineas = 0
        If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar L�nea de Art�culo de Traspaso Almacenes", Err.Description
End Sub


Private Function DatosOk(Optional cabecera As Boolean) As Boolean
Dim b As Boolean

    DatosOk = False
    b = CompForm(Me)
    If Not b Then Exit Function

    'Comprobar que almacen origen y destino son distintos
    If Trim(Text1(2).Text) = Trim(Text1(3).Text) Then
        MsgBox "Almacen Origen y Destino no pueden ser el mismo.", vbExclamation
        b = False
        Exit Function
    End If
    
    If Not cabecera Then b = ComprobarStocksLineas
    
    DatosOk = b
End Function



Private Function ComprobarStocksLineas() As Boolean
'Comprobar para todas las lineas del traspaso que:
' - todos los Art�culos entan en el almacen origen
' - Comprobar que hay suficiente stock en el Almacen Origen de ese Articulo
Dim b As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String

    On Error GoTo ErrStock
    
    
    '---- Laura: 27/09/2006
    If Data2 Is Nothing Then Exit Function
    
    SQL = Data2.RecordSource
    If SQL = "" Then Exit Function
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    'para cada linea comprabar stock del articulo en almacen
    b = True
    While Not RS.EOF And b
        b = ComprobarStock(RS!codArtic, Data1.Recordset!almaorig, RS!Cantidad, CodTipoMov)
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    ComprobarStocksLineas = b
    Exit Function
    '----

    '## ANTES
'    If Not Data2.Recordset.EOF Then  'Si hay lineas
'        Data2.Recordset.MoveFirst
'        b = True
'
'        While Not Data2.Recordset.EOF And b
'            b = ComprobarStock(Data2.Recordset!codArtic, Data1.Recordset!almaorig, Data2.Recordset!Cantidad, CodTipoMov)
'            Data2.Recordset.MoveNext
'        Wend
'    End If
'    ComprobarStocksLineas = b
    '##
    
ErrStock:
    ComprobarStocksLineas = False
    MuestraError Err.Number, "Comprobar stock.", Err.Description
End Function


Private Function DatosOkLinea() As Boolean
Dim b As Boolean

    DatosOkLinea = False
    b = True
        
    If txtAux(0).Text = "" Then
        MsgBox "El campo Cod. Art�culo no puede ser nulo", vbExclamation
        b = False
        Exit Function
    End If
        
    'Comprobamos el campo Cantidad
    If txtAux(2).Text = "" Then
         MsgBox "El campo Cantidad no puede ser nulo", vbExclamation, "Art�culos"
         b = False
    ElseIf Not IsNumeric(txtAux(2).Text) Then
        MsgBox "El campo Cantidad debe ser num�rico", vbExclamation
        b = False
    End If
    If Not b Then
        PonerFoco txtAux(2)
        Exit Function
    End If
    
    'b = ComprobarStock(txtAux(0).Text, txtAux(1).Text, txtAux(2).Text, CodTipoMov)
    b = ComprobarStock(txtAux(0).Text, Text1(2).Text, txtAux(2).Text, CodTipoMov)
         
    DatosOkLinea = b
End Function


Private Sub PonerBotonCabecera(b As Boolean)
On Error Resume Next
    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdRegresar.visible = b
    Me.cmdRegresar.Caption = "Cabecera"
    If b Then
        Me.lblIndicador.Caption = "LINEAS DETALLE"
    Else
        Me.lblIndicador.Caption = ""
    End If
     'Habilitar las opciones correctas del menu seg�n Modo
    PonerModoOpcionesMenu
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu seg�n Nivel de Acceso
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Function InsertarModificarLinea() As Boolean
Dim SQL As String
On Error GoTo EInsertarModificarLinea
    
    SQL = ""
    InsertarModificarLinea = False
    Select Case ModificaLineas
    Case 1 'Insertar
        If DatosOkLinea() Then 'INSERTAR
            SQL = "INSERT INTO slitra (codtrasp,numlinea,codartic,cantidad,observa2) "
            SQL = SQL & " VALUES (" & Val(Text1(0).Text) & ", "
            SQL = SQL & cmdAceptar.Tag & ", "
            SQL = SQL & DBSet(txtAux(0).Text, "T") & ", "
            SQL = SQL & DBSet(txtAux(2).Text, "N") & ","
            SQL = SQL & DBSet(txtAux(3).Text, "T") & ") "
        Else
'            PonerFoco txtAux(3)
        End If
    Case 2 'Modificar
        If DatosOkLinea() Then
            SQL = "UPDATE slitra Set cantidad = " & DBSet(txtAux(2).Text, "N")
            SQL = SQL & ", observa2 = " & DBSet(txtAux(3).Text, "T")
            SQL = SQL & ObtenerWhereCP(True) & " AND " '" WHERE codtrasp =" & Val(Text1(0).Text) & " AND "
            SQL = SQL & " numlinea =" & cmdAceptar.Tag
        End If
    End Select
            
    If SQL <> "" Then
        conn.Execute SQL
        InsertarModificarLinea = True
    End If
    Exit Function
EInsertarModificarLinea:
    MuestraError Err.Number, "Insertar/Modificar Lineas Traspaso Almacenes" & vbCrLf & Err.Description
End Function


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim Tabla As String
Dim Titulo As String

    'Llamamos a al form
    Cad = ""
    If Modo <> 5 Then 'Estamos en Modo de Cabeceras
    'Registro de la tabla de cabeceras: scatra
        Cad = Cad & ParaGrid(Text1(0), 12, "N� Trasp.")
        Cad = Cad & ParaGrid(Text1(1), 15, "Fecha")
        Cad = Cad & ParaGrid(Text1(2), 7, "Orig.")
        Cad = Cad & "Desc. Alm. Orig|salmpr.nomalmac|T||30�"
        Cad = Cad & ParaGrid(Text1(3), 7, "Dest.")
        Cad = Cad & "Alm. Dest|AlmDestino.nomalmac as almdest|T||29�"
        
        Tabla = "(" & NombreTabla & " LEFT JOIN salmpr ON " & NombreTabla & ".almaorig=salmpr.codalmac" & ")"
        Tabla = Tabla & " LEFT JOIN salmpr AS AlmDestino ON " & NombreTabla & ".almadest=AlmDestino.codalmac "
        'tabla = tabla & NombreTabla & ".coddirec=sdirec.coddirec"
        
        ' tabla = "scatra"
        Titulo = Me.Caption
    Else 'Estamos en modo Lineas
        Cad = Cad & "C�digo|sartic|codartic|T||30�Denominacion|sartic|nomartic|T||70�"
        Tabla = "sartic"
        Titulo = "Articulos"
    End If
           
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vtabla = Tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = Titulo
        frmB.vSelElem = 0
'        frmB.vConexionGrid = cAgro 'Conexion a BD Ariagro
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
'            PonerFoco Text1(kCampo)
'        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub HacerBusqueda()
Dim cadB As String
    
    cadB = ObtenerBusqueda3(Me, False)
    cadSeleccion = ObtenerBusqueda(Me, True) 'Para la consulta de report

    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    Else
        'Se muestran en el mismo form
        If cadB <> "" Then
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & Ordenacion
            PonerCadenaBusqueda
        Else
'            CadenaConsulta = "select * from " & NombreTabla & Ordenacion
'            PonerCadenaBusqueda
            MsgBox "Introducir criterios de b�squeda", vbExclamation
            PonerFoco Text1(0)
        End If
    End If
End Sub


Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        If Modo = 1 Then 'Busqueda
             MsgBox "No hay ning�n registro en la tabla " & NombreTabla & " para ese criterio de B�squeda.", vbInformation
             PonerFoco Text1(0)
        Else
            MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        PonerModo 2
        Data1.Recordset.MoveFirst
        PonerCampos
        Me.DataGrid1.Enabled = True
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
EEPonerBusq:
    If Err.Number <> 0 Then MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub PonerCampos()
On Error GoTo EPonerCampos

    If Data1.Recordset.EOF Then Exit Sub
    
    PonerCamposForma Me, Data1
    Text2(0).Text = PonerNombreDeCod(Text1(2), "salmpr", "nomalmac", "codalmac")
    Text2(1).Text = PonerNombreDeCod(Text1(3), "salmpr", "nomalmac", "codalmac")
'    Text2(2).Text = PonerNombreDeCod(Text1(4), "straba", "nomtraba")
    CargaGrid True
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu
    PonerOpcionesMenu

EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Function ActualizarStocks() As Boolean
Dim SQL As String
Dim Cantidad As Single
Dim devuelve As String
Dim RS As ADODB.Recordset

    On Error GoTo EActualizarStock

    ActualizarStocks = False
    
    '---- Laura: 27/09/2006
    'sustituir el data2 por el RecordSEt
    Set RS = New ADODB.Recordset
    RS.Open Data2.RecordSource, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
'    While Not Data2.Recordset.EOF

        'Actualizar el stock si el articulo tiene control de stock
        devuelve = DevuelveDesdeBDNew(cAgro, "sartic", "ctrstock", "codartic", RS!codArtic, "T")
        If Val(devuelve) = 1 Then

            Cantidad = CSng(RS!Cantidad) 'Cant a traspasar
            
            '==== Almacen Origen
            SQL = "UPDATE salmac Set canstock = canstock - " & DBSet(Cantidad, "N")
            SQL = SQL & " WHERE codartic =" & DBSet(RS!codArtic, "T") & " AND "
            SQL = SQL & " codalmac =" & Data1.Recordset!almaorig
            conn.Execute SQL
        
            '==== Almacen Destino
            'Comprobar que existe el articulo en Almacen Destino
            devuelve = DevuelveDesdeBDNew(cAgro, "salmac", "codalmac", "codartic", RS!codArtic, "T", , "codalmac", Text1(3).Text, "N")
            If devuelve = "" Then 'No hay de ese art�culo en Destino
                SQL = "INSERT INTO salmac (codartic,codalmac,canstock,stockmin,puntoped,stockmax,stockinv,fechainv,horainve,statusin)"
                SQL = SQL & " VALUES (" & DBSet(RS!codArtic, "T") & "," & Val(Text1(3).Text) & "," & DBSet(Cantidad, "N") & ",0,0,0,0,NULL,NULL,0)"
            Else 'Existe el artic en almac. Dest -> Aumentar stock
                SQL = "UPDATE salmac Set canstock = canstock + " & DBSet(Cantidad, "N")
                SQL = SQL & " WHERE codartic =" & DBSet(RS!codArtic, "T") & " AND "
                SQL = SQL & " codalmac =" & Data1.Recordset!almadest
            End If
            
            conn.Execute SQL
        End If
        'Data2.Recordset.MoveNext
        RS.MoveNext
    Wend
    
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
        ActualizarStocks = False
    Else
        ActualizarStocks = True
    End If
    
EActualizarStock:
End Function


Private Sub BotonActualizar()
'Actualizar Traspaso Almacen
Dim SQL As String

    If Data1.Recordset.EOF Then
        MsgBox "Ning�n Traspaso para actualizar.", vbExclamation
        Exit Sub
    End If
    
    If Data2 Is Nothing Then Exit Sub
    
    If Data2.Recordset.EOF Then
        MsgBox "No hay lineas insertadas para este N� de Traspaso", vbExclamation
        Exit Sub
    End If
    
    If Data1.Recordset!situacio = 0 Then 'Informe No Impreso
        SQL = "Actualizaci�n Traspaso Almacenes." & vbCrLf
        SQL = SQL & "------------------------------------------" & vbCrLf & vbCrLf
        SQL = SQL & "NO ESTA IMPRESO EL TRASPASO:" & vbCrLf
        SQL = SQL & vbCrLf & "N� Trasp.     :  " & Format(Data1.Recordset.Fields(0), "0000000")
        SQL = SQL & vbCrLf & "Fecha Trasp.  :  " & CStr(Data1.Recordset.Fields(1))
        SQL = SQL & vbCrLf & "Almac. Origen :  " & Format(Data1.Recordset.Fields(2), "000") & " - " & Text2(0).Text & "     "
        SQL = SQL & vbCrLf & "Almac. Destino:  " & Format(Data1.Recordset.Fields(3), "000") & " - " & Text2(1).Text & "     "
        SQL = SQL & vbCrLf & vbCrLf & " �Desea continuar ? "
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then
            Exit Sub
        End If
    Else 'Informe Impreso
        SQL = "Actualizaci�n Traspaso Almacenes." & vbCrLf
        SQL = SQL & "-----------------------------------------" & vbCrLf & vbCrLf
        SQL = SQL & "Va a Actualizar el Traspaso:" & vbCrLf
        SQL = SQL & vbCrLf & "   N� Trasp.   : " & Format(Data1.Recordset.Fields(0), "0000000")
        SQL = SQL & vbCrLf & "   Fecha Trasp.: " & CStr(Data1.Recordset.Fields(1))
        SQL = SQL & vbCrLf & "   Almac. Orig.: " & Format(Data1.Recordset.Fields(2), "000") & " - " & Text2(0).Text & "     "
        SQL = SQL & vbCrLf & "   Almac. Dest.: " & Format(Data1.Recordset.Fields(3), "000") & " - " & Text2(1).Text & "     "
        SQL = SQL & vbCrLf & vbCrLf & " �Desea continuar ? "
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then
            Exit Sub
        End If
    End If
    
    
    'bloqueamos el registro que vamos a traspasar
    If Not BLOQUEADesdeFormulario(Me) Then Exit Sub
    
    
    'realizamos el traspaso de almacenes
    Me.ProgressBar1.visible = True
    NumRegElim = Data1.Recordset.AbsolutePosition
    
    If ActualizarTraspaso Then
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else 'Solo habia un registro
            LimpiarCampos
            CargaGrid False
            PonerModo 0
            Me.Refresh
        End If
    End If
    Me.ProgressBar1.visible = False
    TerminaBloquear
End Sub


Private Function ActualizarTraspaso() As Boolean
Dim Donde As String
Dim devuelve As String
Dim bol As Boolean

    On Error GoTo EActualizarTraspaso
    
    'Comprobamos que no existe en historico de traspasos
    devuelve = DevuelveDesdeBDNew(cAgro, "schtra", "codtrasp", "codtrasp", Data1.Recordset!codtrasp, "N", , "fechatra", Data1.Recordset!fechatra, "F")
    If Trim(devuelve <> "") Then
        devuelve = "Ya existe en el hist�rico el traspaso:" & vbCrLf
        devuelve = devuelve & " N�: " & Data1.Recordset!codtrasp & vbCrLf
        devuelve = devuelve & " Fecha: " & Data1.Recordset!fechatra
        MsgBox devuelve, vbExclamation
        Exit Function
    End If
    
    'Comprobar que en almacen origen existe la cantidad a traspasar
    'de cada linea de articulo
    If Not ComprobarStocksLineas Then Exit Function
    
    'Aqui empieza transaccion
    conn.BeginTrans
    bol = ActualizarElTraspaso(Donde)

EActualizarTraspaso:
        If Err.Number <> 0 Then
'            devuelve = "Actualizar Traspaso." & vbCrLf & "----------------------------" & vbCrLf
'            devuelve = devuelve & Donde
'            MuestraError Err.Number, devuelve, Err.Description
            devuelve = Err.Description & ": " & Err.Number
            bol = False
        Else
            devuelve = ""
        End If
        If bol Then
            conn.CommitTrans
            ActualizarTraspaso = True
        Else
            conn.RollbackTrans
            devuelve = "Actualizar Traspaso." & vbCrLf & "----------------------------" & vbCrLf & "ERROR: " & Donde & vbCrLf & devuelve
            MsgBox devuelve, vbExclamation
        End If
End Function


Private Function ActualizarElTraspaso(ByRef ADonde As String) As Boolean
Dim cadError As String

    ActualizarElTraspaso = False
    
    'Insertamos en cabeceras Historico
    ADonde = "Insertando datos en historico cabeceras traspaso almacenes"
    If Not InsertarCabeceraHistorico Then Exit Function
    IncrementarProgres 2
     
    'Insertamos en lineas Historico
    ADonde = "Insertando datos en Historico lineas Traspaso Almacenes"
    If Not InsertarLineasHistorico(cadError) Then
        ADonde = ADonde & vbCrLf & cadError
        Exit Function
    End If
    IncrementarProgres 2
    
    'Modificar stock (Tabla: salmac)
    ADonde = "Actualizando Stocks Almacenes"
    If Not ActualizarStocks() Then Exit Function
    IncrementarProgres 2
    
    'Insertamos en Movimientos Art�culos (Tabla: smoval)
    ADonde = "Insertando datos en Movimientos de Articulos"
    If Not InsertarMovimArticulos(cadError) Then
        ADonde = ADonde & vbCrLf & cadError
        Exit Function
    End If
    IncrementarProgres 2

    
    'Borramos cabeceras y lineas del TRaspaso
    ADonde = "Borrar cabeceras y lineas en Traspaso Almacenes"
    If Not BorrarTraspaso(cadError) Then
        ADonde = ADonde & vbCrLf & cadError
        Exit Function
    End If
    IncrementarProgres 2
    ActualizarElTraspaso = True
End Function


Private Function InsertarCabeceraHistorico() As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
On Error GoTo EInsertarCab

    SQL = "SELECT codtrasp,fechatra,almaorig,almadest,observa1 from scatra "
    SQL = SQL & ObtenerWhereCP(True)
    SQL = SQL & " AND fechatra='" & Format(Data1.Recordset!fechatra, "yyyy-mm-dd") & "'"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        SQL = "INSERT INTO schtra (codtrasp, fechatra,hormovim,almaorig,almadest,observa1) "
        SQL = SQL & " VALUES (" & RS.Fields(0).Value & ", '" & Format(RS.Fields(1).Value, "yyyy-mm-dd") & "', '"
        SQL = SQL & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', " & RS.Fields(2).Value & ", " & RS.Fields(3).Value & ", "
        SQL = SQL & DBSet(RS.Fields(4).Value, "T") & ")"
    End If
    RS.Close
    Set RS = Nothing
    
    conn.Execute SQL
    
EInsertarCab:
    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        InsertarCabeceraHistorico = False
    Else
        InsertarCabeceraHistorico = True
    End If
End Function


Private Function InsertarLineasHistorico(MenError As String) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
On Error GoTo EInsertarLineas

    SQL = "SELECT codtrasp, numlinea, codartic, cantidad, observa2 from slitra "
    SQL = SQL & ObtenerWhereCP(True)
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    RS.MoveFirst
    While Not RS.EOF
        SQL = "INSERT INTO slhtra (codtrasp, fechamov, numlinea, codartic, cantidad, observa2)"
        SQL = SQL & " VALUES (" & RS.Fields(0).Value & ", '" & Format(Data1.Recordset!fechatra, FormatoFecha) & "', "
        SQL = SQL & RS.Fields(1).Value & ", " & DBSet(RS.Fields(2).Value, "T") & ", "
        SQL = SQL & DBSet(RS.Fields(3).Value, "N") & ", " & DBSet(RS.Fields(4).Value, "T") & ")"
        conn.Execute SQL
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
EInsertarLineas:
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
        InsertarLineasHistorico = False
        RS.Close
        Set RS = Nothing
        MenError = Err.Number & ": " & Err.Description
    Else
        MenError = ""
        InsertarLineasHistorico = True
    End If
End Function


Private Sub IncrementarProgres(Veces As Integer)
On Error Resume Next
    Me.ProgressBar1.Value = Me.ProgressBar1.Value + (Veces * 10)
    If Err.Number <> 0 Then Err.Clear
    Me.Refresh
End Sub


Private Function BorrarTraspaso(MenError As String) As Boolean
Dim SQL As String

    BorrarTraspaso = False
    
    'Borramos las lineas
    SQL = "Delete from "
    SQL = SQL & "slitra"
    SQL = SQL & " WHERE codtrasp = " & Data1.Recordset!codtrasp
    conn.Execute SQL
    
    'La cabecera
    SQL = "Delete from "
    SQL = SQL & "scatra"
    SQL = SQL & " WHERE codtrasp =" & Data1.Recordset!codtrasp
    conn.Execute SQL
    
    If Err.Number <> 0 Then
        BorrarTraspaso = False
        MenError = Err.Number & ": " & Err.Description
    Else
        BorrarTraspaso = True
        MenError = ""
    End If
End Function

Public Sub ActualizarSituacionImpresion()
Dim Cad As String
Dim Indicador As String
On Error GoTo EImpresion
   
        Cad = "(codtrasp=" & Val(Text1(0).Text) & ")"
        If SituarData(Data1, Cad, Indicador) Then
            If Modo <> 5 Then
                PonerModo 2
            Else
                PonerModo 5
            End If
            PonerCampos
            lblIndicador.Caption = Indicador
        Else
            PonerModo 0
        End If
EImpresion:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Function InsertarMovimArticulos(MenError As String) As Boolean
Dim SQL As String, Cad As String
Dim RS As ADODB.Recordset
Dim vImporte As Single, vPrecioVenta As String
Dim vTipoMov As CTiposMov
Dim bol As Boolean
    
    On Error GoTo EInsertar

    bol = True
    Set vTipoMov = New CTiposMov
    If vTipoMov.leer(CodTipoMov) Then
        'Se han cargado correctamente los valores de la clase
        SQL = "SELECT scatra.codtrasp, fechatra, almaorig, almadest, numlinea, codartic, cantidad "
        SQL = SQL & " FROM scatra LEFT JOIN slitra ON scatra.codtrasp=slitra.codtrasp "
        SQL = SQL & " WHERE scatra.codtrasp =" & Data1.Recordset!codtrasp
    
        Set RS = New ADODB.Recordset
        RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not RS.EOF
             'Obtener el precio de venta del articulo, si tiene control de stock
            Cad = "ctrstock"
            '++monica a�adido el tipo de precio antes solo era el pmp
            If vParamAplic.TipoPrecio = 0 Then 'precio medio ponderado
                vPrecioVenta = DevuelveDesdeBDNew(cAgro, "sartic", "preciomp", "codartic", RS.Fields!codArtic, "T", Cad)
            Else 'ultimo precio
                vPrecioVenta = DevuelveDesdeBDNew(cAgro, "sartic", "preciouc", "codartic", RS.Fields!codArtic, "T", Cad)
            End If
            If vPrecioVenta <> "" Then
                vImporte = Round2(RS.Fields!Cantidad * CSng(vPrecioVenta), 2)
            Else
                vImporte = 0
            End If
            If Val(Cad) = 1 Then
                'Insertar Movimiento de Salida en Almacen Origen
                SQL = "INSERT INTO smoval (codartic, codalmac, fechamov, horamovi, tipomovi, detamovi, cantidad, impormov, codigope, letraser, document, numlinea) "
                SQL = SQL & " VALUES (" & DBSet(RS.Fields!codArtic, "T") & ", " & RS.Fields!almaorig & ", '" & Format(RS.Fields!fechatra, "yyyy-mm-dd") & "', '"
                SQL = SQL & Format(RS.Fields!fechatra & " " & Time, "yyyy-mm-dd hh:mm:ss") & "', 0" & ", '" & vTipoMov.TipoMovimiento & "', " & DBSet(RS.Fields!Cantidad, "N") & ", " & DBSet(vImporte, "N") & ", 0, " '& RS.Fields!codtraba & ", "
                SQL = SQL & DBSet(vTipoMov.LetraSerie, "T") & ", " & RS.Fields!codtrasp & ", " & RS.Fields!numlinea & ")"
                conn.Execute SQL
                
                'Insertar Movimiento de Entrada en Almacen Destino
                SQL = "INSERT INTO smoval (codartic, codalmac, fechamov, horamovi, tipomovi, detamovi, cantidad, impormov, codigope, letraser, document, numlinea) "
                SQL = SQL & " VALUES (" & DBSet(RS.Fields!codArtic, "T") & ", " & RS.Fields!almadest & ", '" & Format(RS.Fields!fechatra, "yyyy-mm-dd") & "', '"
                SQL = SQL & Format(RS.Fields!fechatra & " " & Time, "yyyy-mm-dd hh:mm:ss") & "', 1" & ", '" & vTipoMov.TipoMovimiento & "', " & DBSet(RS.Fields!Cantidad, "N") & ", " & DBSet(vImporte, "N") & ", 0, " '& RS.Fields!codtraba & ", "
                SQL = SQL & DBSet(vTipoMov.LetraSerie, "T") & ", " & RS.Fields!codtrasp & ", " & RS.Fields!numlinea & ")"
                conn.Execute SQL
            End If
            RS.MoveNext
        Wend
    Else
        bol = False
    End If
    Set vTipoMov = Nothing
    RS.Close
    Set RS = Nothing
    
EInsertar:
    If Err.Number <> 0 Then
        Set vTipoMov = Nothing
        RS.Close
        Set RS = Nothing
        MenError = Err.Number & ": " & Err.Description
    End If
    
    If Err.Number <> 0 Or Not bol Then
        InsertarMovimArticulos = False
    Else
        InsertarMovimArticulos = True
        MenError = ""
    End If
End Function


Private Sub BotonImprimir()
    frmListado.NumCod = Text1(0).Text
    
    If Not EsHistorico Then
        AbrirListado (7) '7: Informe Traspaso de Almacen
        ActualizarSituacionImpresion
    Else
        BotonImprimirHco
    End If
End Sub


Private Sub BotonImprimirHco()
Dim indRPT As Byte
Dim cadParam As String
Dim Cad As String
Dim numParam As Byte
Dim nomDocu As String

    cadParam = "|"
    numParam = 0
    If Not PonerParamEmpresa(cadParam, numParam) Then Exit Sub
    
    indRPT = 2 '2: Historico Traspaso de Almacen
    If PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then
        With frmImprimir
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .NombreRPT = nomDocu
            .EnvioEMail = False
            .Opcion = 7
            .Titulo = "Hist. Traspaso Alm."
            If cadSeleccion <> "" Then
                .FormulaSeleccion = cadSeleccion
            Else
                'Se Llama desde dobleclick en frmAlmMovimArticulos
                'o estamos en Historico
                Cad = "{schtra.codtrasp}= " & Data1.Recordset!codtrasp
                Cad = Cad & " and {schtra.fechatra}= Date(" & Year(Data1.Recordset!fechatra) & "," & Month(Data1.Recordset!fechatra) & "," & Day(Data1.Recordset!fechatra) & ")" & ""
                .FormulaSeleccion = Cad
            End If
            .Show vbModal
        End With
    End If
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Function InsertarTraspaso(vSQL As String, vTipoMov As CTiposMov) As Boolean
Dim MenError As String
Dim bol As Boolean
On Error GoTo EInsertarMovimiento
    
    bol = True
    InsertarTraspaso = False
    
    'Aqui empieza transaccion
    conn.BeginTrans
    MenError = "Error al insertar en la tabla de Traspasos(scatra)."
    conn.Execute vSQL, , adCmdText
    
    MenError = "Error al actualizar el contador del recibo."
    vTipoMov.IncrementarContador (CodTipoMov)
    

EInsertarMovimiento:
        If Err.Number <> 0 Then
            MenError = "Insertando Traspaso." & vbCrLf & "----------------------------" & vbCrLf & MenError
            MuestraError Err.Number, MenError, Err.Description
            bol = False
        End If
        If bol Then
            conn.CommitTrans
            InsertarTraspaso = True
        Else
            conn.RollbackTrans
            InsertarTraspaso = False
        End If
End Function


Private Function ObtenerWhereCP(conWhere As Boolean) As String
On Error Resume Next
    If conWhere Then
        ObtenerWhereCP = " WHERE codtrasp= " & Val(Text1(0).Text)
    Else
        ObtenerWhereCP = " codtrasp= " & Val(Text1(0).Text)
    End If
End Function


Private Sub PosicionarData()
'Despues de hacer refresh del Data, volver a situar el Data en el registro que estaba
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP(False) & ")"
         If SituarData(Data1, vWhere, Indicador) Then
             PonerModo 2
             PonerCampos
             lblIndicador.Caption = Indicador
        Else
             LimpiarCampos
             'Poner los grid sin apuntar a nada
             LimpiarDataGrids
             PonerModo 0
         End If
    End If
End Sub


Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim SQL As String

    Set vTipoMov = New CTiposMov
    
    If vTipoMov.leer(CodTipoMov) Then
        Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
        Text1(0).Text = Format(Text1(0).Text, "0000000")
        cmdCancelar.Caption = "Cancelar"
        SQL = CadenaInsertarDesdeForm(Me)
        If SQL <> "" Then
            If InsertarTraspaso(SQL, vTipoMov) Then
                CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                PonerCadenaBusqueda
                PonerModo 2
                 'Ponerse en Modo Insertar Lineas
'--monica: rollo toolbar
'                BotonLineas
                BotonAnyadirLineas
            End If
        End If
    End If
    Set vTipoMov = Nothing
End Sub


Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ning�n registro
On Error Resume Next
    CargaGrid False
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub AbrirListado(numero As Byte)
    Screen.MousePointer = vbHourglass
    frmListado2.OpcionListado = numero
    frmListado2.Show vbModal
    Screen.MousePointer = vbDefault
End Sub
