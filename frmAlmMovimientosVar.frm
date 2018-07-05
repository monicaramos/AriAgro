VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAlmMovimientosVar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Movimientos de Servicios Varios"
   ClientHeight    =   10020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13770
   Icon            =   "frmAlmMovimientosVar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10020
   ScaleWidth      =   13770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkVistaPrevia 
      Caption         =   "Vista previa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10935
      TabIndex        =   45
      Top             =   285
      Width           =   1605
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   4635
      TabIndex        =   43
      Top             =   90
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   44
         Top             =   180
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Primero"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Anterior"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Siguiente"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Último"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3780
      TabIndex        =   41
      Top             =   90
      Width           =   795
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   210
         TabIndex        =   42
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
               Object.ToolTipText     =   "Actualizar"
               Object.Tag             =   "2"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   135
      TabIndex        =   39
      Top             =   90
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   40
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
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   5
      Left            =   6120
      MaxLength       =   16
      TabIndex        =   18
      Text            =   "importe"
      Top             =   5850
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   5040
      MaxLength       =   16
      TabIndex        =   17
      Text            =   "precio"
      Top             =   5850
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   9075
      MaxLength       =   12
      TabIndex        =   8
      Tag             =   "Matricula|T|S|||scaser|matriveh||N|"
      Top             =   1500
      Width           =   1305
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
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
      Left            =   2265
      Locked          =   -1  'True
      TabIndex        =   38
      Text            =   "Text2"
      Top             =   1740
      Width           =   4620
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
      Left            =   1395
      MaxLength       =   6
      TabIndex        =   5
      Tag             =   "Cod.Socio|N|S|||scaser|codsocio|000000|N|"
      Text            =   "Text1"
      Top             =   1740
      Width           =   870
   End
   Begin VB.CheckBox chkImpresion 
      Caption         =   "Impreso"
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
      Height          =   255
      Index           =   0
      Left            =   4860
      TabIndex        =   3
      Tag             =   "Situación Impresión|N|N|||scamov|situacio||N|"
      Top             =   1320
      Width           =   1215
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
      Left            =   150
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Tag             =   "Tipo|N|N|||scaser|clisoc||N|"
      Top             =   1230
      Width           =   1860
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
      Left            =   1395
      MaxLength       =   6
      TabIndex        =   6
      Tag             =   "Cod. Cliente|N|S|||scaser|codclien|000000|N|"
      Text            =   "Text1"
      Top             =   2130
      Width           =   870
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
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
      Left            =   2265
      Locked          =   -1  'True
      TabIndex        =   34
      Text            =   "Text2"
      Top             =   2130
      Width           =   4620
   End
   Begin MSComctlLib.Toolbar ToolAux 
      Height          =   390
      Index           =   0
      Left            =   135
      TabIndex        =   32
      Top             =   2925
      Width           =   1110
      _ExtentX        =   1958
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
      Left            =   4830
      MaxLength       =   8
      TabIndex        =   4
      Tag             =   "Hora|H|N|||scaser|hormovim|hh:mm:ss|N|"
      Text            =   "Text1"
      Top             =   1230
      Width           =   855
   End
   Begin VB.ComboBox cboAux 
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
      ItemData        =   "frmAlmMovimientosVar.frx":000C
      Left            =   7200
      List            =   "frmAlmMovimientosVar.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Tag             =   "Situación Impresión|N|N|||scaser|situacio||N|"
      Top             =   5850
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Left            =   1200
      TabIndex        =   29
      ToolTipText     =   "Buscar artículo"
      Top             =   5850
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   8730
      MaxLength       =   50
      TabIndex        =   20
      Text            =   "observac"
      Top             =   5850
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   3960
      MaxLength       =   16
      TabIndex        =   16
      Text            =   "cantidad"
      Top             =   5850
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
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
      Height          =   330
      Index           =   1
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   21
      Text            =   "nombre artic"
      Top             =   5850
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   240
      MaxLength       =   16
      TabIndex        =   15
      Text            =   "codartic"
      Top             =   5850
      Visible         =   0   'False
      Width           =   855
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
      Left            =   11220
      TabIndex        =   10
      Top             =   9465
      Width           =   1065
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
      Left            =   12495
      TabIndex        =   11
      Top             =   9465
      Width           =   1065
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
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
      Left            =   12510
      TabIndex        =   28
      Top             =   9495
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   9390
      Width           =   3000
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
         TabIndex        =   27
         Top             =   180
         Width           =   2340
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
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
      Left            =   2265
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "Text2"
      Top             =   2520
      Width           =   4620
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   4
      Left            =   6945
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Tag             =   "Observaciones|T|S|||scaser|observa1||N|"
      Text            =   "frmAlmMovimientosVar.frx":0010
      Top             =   2130
      Width           =   6630
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
      Left            =   1395
      MaxLength       =   3
      TabIndex        =   7
      Tag             =   "Cod. Almacen|N|N|0|999|scaser|codalmac|000|N|"
      Text            =   "Text1"
      Top             =   2520
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
      Index           =   1
      Left            =   3420
      MaxLength       =   10
      TabIndex        =   2
      Tag             =   "Fecha|F|N|||scaser|fecmovim|dd/mm/yyyy|N|"
      Text            =   "Text1"
      Top             =   1230
      Width           =   1350
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAlmMovimientosVar.frx":0016
      Height          =   5985
      Left            =   120
      TabIndex        =   12
      Top             =   3390
      Width           =   13425
      _ExtentX        =   23680
      _ExtentY        =   10557
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
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
      Left            =   2040
      MaxLength       =   7
      TabIndex        =   1
      Tag             =   "Nº Movimiento|N|S|0||scaser|codservi|0000000|S|"
      Text            =   "Text1"
      Top             =   1230
      Width           =   1275
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
      Left            =   3600
      TabIndex        =   30
      Top             =   8505
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   12975
      TabIndex        =   46
      Top             =   225
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
   Begin VB.Label Label9 
      Caption         =   "Matrícula del Camión"
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
      Left            =   6930
      TabIndex        =   37
      Top             =   1530
      Width           =   2145
   End
   Begin VB.Label Label7 
      Caption         =   "Socio"
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
      Left            =   180
      TabIndex        =   36
      Top             =   1740
      Width           =   705
   End
   Begin VB.Label Label8 
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
      Height          =   255
      Left            =   180
      TabIndex        =   35
      Top             =   2130
      Width           =   705
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   2
      Left            =   1095
      ToolTipText     =   "Buscar cliente"
      Top             =   2145
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   1095
      ToolTipText     =   "Buscar socio"
      Top             =   1755
      Width           =   240
   End
   Begin VB.Label Label5 
      Caption         =   "Tipo Movimiento"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   33
      Top             =   930
      Width           =   1695
   End
   Begin VB.Image imgFec 
      Height          =   240
      Index           =   0
      Left            =   4515
      Picture         =   "frmAlmMovimientosVar.frx":002B
      ToolTipText     =   "Buscar fecha"
      Top             =   930
      Width           =   240
   End
   Begin VB.Label Label4 
      Caption         =   "Hora"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4860
      TabIndex        =   31
      Top             =   930
      Width           =   735
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   1095
      ToolTipText     =   "Buscar almacen"
      Top             =   2535
      Width           =   240
   End
   Begin VB.Label Label6 
      Caption         =   "Observaciones"
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
      Left            =   6930
      TabIndex        =   24
      Top             =   1830
      Width           =   1500
   End
   Begin VB.Label Label3 
      Caption         =   "Almacen"
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
      Left            =   210
      TabIndex        =   23
      Top             =   2520
      Width           =   915
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   3435
      TabIndex        =   22
      Top             =   945
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "Número"
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
      Left            =   2070
      TabIndex        =   14
      Top             =   960
      Width           =   1440
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
      TabIndex        =   13
      Top             =   8490
      Visible         =   0   'False
      Width           =   3495
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
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmAlmMovimientosVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public EsHistorico As Boolean 'Si es true abrir el formulario con la tabla de
                              'historico schser, y solo en modo de consulta

'Si se llama de la busqueda en el frmAlmMovimArticulos se accede
'a las tablas del histórico de movimiento seleccionado (solo consulta)
Public hcoCodMovim As Long 'cod. movim del historico
Public hcoCliSoc As Byte
Public hcoFechaMovim As Date 'Fecha del historico


'-----------------------------------------------------------------------

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1

Private WithEvents frmA As frmManAlmProp 'Almacen Origen/Destino
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmArt As frmManArtic   'Form Articulos
Attribute frmArt.VB_VarHelpID = -1
Private WithEvents frmSoc As frmBasico  'Form accedemos a recoleccion
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmCli As frmClientes  'Form Clientes
Attribute frmCli.VB_VarHelpID = -1


Dim NombreTabla As String
Dim NomTablaLineas As String
Dim Ordenacion As String

Private Modo As Byte
Private ModoAnterior As Byte
Dim kCampo As Integer

Dim btnAnyadir As Byte
'Variable que indica el número del Boton  Anyadir en la Toolbar1

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim CodTipoMov As String
'Codigo tipo de movimiento en función del valor en la tabla de parámetros: stipom

Dim CadenaConsulta As String
Dim cadSeleccion As String 'Cadena de seleccion para FormulaSelection del Informe


Private HaDevueltoDatos As Boolean

Dim PrimeraVez As Boolean


Private Sub cboAux_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboImpresion_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkVistaPrevia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim Cad As String, Indicador As String
On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    
    Select Case Modo
    Case 1 'BUSQUEDA
        cadSeleccion = ""
        HacerBusqueda
        
    Case 3 'INSERTAR
        If DatosOk Then InsertarCabecera

    Case 4 'MODIFICAR
        If DatosOk Then
            If ModificaDesdeFormulario(Me) Then
                TerminaBloquear
                Cad = "(" & ObtenerWhereCP(False) & ")"
                If SituarDataMULTI(Data1, Cad, Indicador) Then
                    PonerModo 2
                    lblIndicador.Caption = Indicador
                Else
                    PonerModo 0
                End If
            End If
        End If
            
    Case 5 'Lineas Movimientos Almacenes
        If InsertarModificarLinea Then
            'Reestablecemos los campos y ponemos el grid
            DataGrid1.AllowAddNew = False
            If ModificaLineas = 1 Then 'Insertar
'++monica: rollo
                'TerminaBloquear
                CargaGrid True
                ModificaLineas = 0
                If Me.Data2.Recordset.RecordCount < 10 Then
                    BotonAnyadirLineas
                Else
                    CargaTxtAux False, False
                    DataGrid1.AllowAddNew = False
                    DataGrid1.Refresh
                    DataGrid1.Enabled = True
                
                    PonerModo 2
                    PonerCampos
                End If
            ElseIf ModificaLineas = 2 Then 'Modificar
                TerminaBloquear
'                Data2.Recordset.Find (Data2.Recordset.Fields(1).Name & " =" & CInt(Me.cmdAceptar.Tag))
                ModificaLineas = 0
'--monica: rollo toolbar
'                PonerBotonCabecera True
                CargaTxtAux False, False
                Me.lblIndicador.Caption = ""
                CargaGrid True
                Data2.Recordset.Find (Data2.Recordset.Fields(1).Name & " =" & CInt(Me.cmdAceptar.Tag))
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
    frmArt.DatosADevolverBusqueda = "0|1|" 'Abre en Modo busqueda
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
            TerminaBloquear
            
            CargaTxtAux False, False
            DataGrid1.AllowAddNew = False
            If Not ModificaLineas = 2 Then '2 = Modificar
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            End If
            ModificaLineas = 0
'--monica: rollo toolbar
'            PonerBotonCabecera True
            DataGrid1.Refresh
            DataGrid1.Enabled = True
            PonerModo 2
            
            PonerCampos
    End Select
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdRegresar_Click()
Dim Cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then 'modo 5: Mantenimiento Lineas
'--monica: rollo toolbar
'        PonerBotonCabecera False
        PonerModo 2
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid Me.DataGrid1
            DataGrid1.Bookmark = 1
        End If
    
    Else 'Se llama desde algún Prismatico de otro Form al Mantenimiento de Trabajadores
        If Data1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        Cad = Data1.Recordset.Fields(0) & "|"
        Cad = Cad & Data1.Recordset.Fields(1) & "|"
'        RaiseEvent DatoSeleccionado(cad)
        Unload Me
    End If
End Sub


Private Sub cmdRegresar_KeyPress(KeyAscii As Integer)
    If Modo = 5 And KeyAscii = 27 Then 'ESC 'Modo Lineas
        cmdRegresar_Click
    End If
End Sub


Private Sub Combo1_Change(Index As Integer)
    PonerSocioVisible
    If Modo <> 1 Then PonerFoco Text1(1)
End Sub


Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    If Combo1(Index).BackColor = vbYellow Then Combo1(Index).BackColor = vbWhite
    PonerSocioVisible
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    If Modo = 5 And KeyAscii = 27 Then 'ESC 'Modo Lineas
        cmdRegresar_Click
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub Form_Activate()
'    Screen.MousePointer = vbDefault

    If PrimeraVez Then
        PrimeraVez = False
        'Viene de DblClick en frmAlmMovimArticulos y carga el form con los valores
        If hcoCodMovim <> "" And Not Data1.Recordset.EOF Then PonerCadenaBusqueda
    End If
        
End Sub

Private Sub Form_Load()
Dim i As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    'ICONOS de La toolbar
    btnAnyadir = 5 'Posicion del boton Añadir en la toolbar1
    btnPrimero = 15 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
'    With Toolbar1
'        .ImageList = frmPpal.imgListComun
'        .DisabledImageList = frmPpal.imgListComun_BN
'       'ASignamos botones
'        .Buttons(1).Image = 1   'Buscar
'        .Buttons(2).Image = 2 'Ver Todos
'        .Buttons(5).Image = 3 'Añadir
'        .Buttons(6).Image = 4 'Modificar
'        .Buttons(7).Image = 5 'Eliminar
'        .Buttons(9).Image = 21 '10 'Mantenimiento Líneas
'        .Buttons(10).Image = 16 '39 'Actualizar
'        .Buttons(12).Image = 10 'Imprimir
'        .Buttons(13).Image = 11 'Salir
'        .Buttons(btnPrimero).Image = 6 'Primero
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
'        .Buttons(btnPrimero + 3).Image = 9 'Ultimo
'    End With

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
        .Buttons(1).Image = 16  'Actualizar
    End With
    
    ' desplazamiento
    With Me.ToolbarDes
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 6
        .Buttons(2).Image = 7
        .Buttons(3).Image = 8
        .Buttons(4).Image = 9
    End With
    
' comentado de momento
'    ' La Ayuda
'    With Me.ToolbarAyuda
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 12
'    End With




    ' ******* si n'hi han llínies *******
    'ICONETS DE LES BARRES ALS TABS DE LLÍNIA
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
    
    'cargar IMAGES de busqueda
    For i = 0 To imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    
    CargaCombo
    
    LimpiarCampos   'Limpia los campos TextBox
    DataGrid1.ClearFields
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    CodTipoMov = "SES"
    
    'campo situacio solo en tabla scaser
    Me.chkImpresion(0).visible = Not EsHistorico
    'Campo Hora solo en el Historico
    Me.Label4.visible = EsHistorico
    Me.Text1(5).visible = EsHistorico
    
    cadSeleccion = ""
   
    If Not EsHistorico Then
        NombreTabla = "scaser"
        NomTablaLineas = "sliser" 'Tabla lineas de Movimientos
        Me.Caption = "Movimientos de Servicios Varios"
    Else
        NombreTabla = "schser"
        NomTablaLineas = "slhser"
        CargarTagsHco Me, "scaser", NombreTabla
        Me.Caption = "Histórico Movimientos de Servicios Varios"
    End If
    Ordenacion = " ORDER BY codservi"
    
    CadenaConsulta = "Select * from " & NombreTabla
    If hcoCodMovim <> -1 Then
    'Se llama desde Dobleclick en frmAlmMovimArticulos
        CadenaConsulta = CadenaConsulta & " where codservi=" & hcoCodMovim & " and fecmovim= """ & Format(hcoFechaMovim, "yyyy-mm-dd") & """" & " and clisoc = " & DBSet(hcoCliSoc, "N")
    Else
        CadenaConsulta = CadenaConsulta & " WHERE codservi = -1"
    End If
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    If Not Data1.Recordset.EOF Then 'Se llama desde DblClick frmAlmMovimArticulos
                                    'Se carga con el valor del registro del DblClick
        Data1.Recordset.MoveFirst
        Me.Text1(0).Text = Format(Data1.Recordset!codservi, "0000000")
        Me.Text1(1).Text = Data1.Recordset!fecmovim
        Me.Text1(5).Text = Format(Data1.Recordset!hormovim, "hh:mm:ss")
        'Cod. Almacen
        Me.Text1(6).Text = Format(Data1.Recordset!codAlmac, "000")
        Me.Text2(6).Text = PonerNombreDeCod(Text1(6), "salmpr", "nomalmac", "codalmac")
        'Cod. Socio
        Me.Text1(2).Text = Format(Data1.Recordset!CodSocio, "000000")
        Me.Text2(2).Text = PonerNombreDeCod(Text1(2), "rsocios", "nomsocio", "codsocio")
        'Cod. Cliente
        Me.Text1(3).Text = Format(Data1.Recordset!CodClien, "000000")
        Me.Text2(3).Text = PonerNombreDeCod(Text1(3), "clientes", "nomclien", "codclien")
        
        'Observaciones
        Text1(4).Text = DBLet(Data1.Recordset!observa1, "T")
        CargaGrid True
    Else
        CargaGrid False '(Modo = 2) 'False
    End If
    If hcoCodMovim <> -1 Then
        PonerModo 2
    Else
        PonerModo 0
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim b As Boolean
Dim i As Byte
Dim Sql As String
On Error GoTo ECarga

    b = DataGrid1.Enabled
    
    Sql = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data2, Sql, False
    
    DataGrid1.Columns(0).visible = False 'Cod. Movim
    DataGrid1.Columns(1).visible = False 'clisoc
    DataGrid1.Columns(2).visible = False 'Numlinea
    i = 3
    
    'Cod. Artículo
    DataGrid1.Columns(i).Caption = "Cod. Articulo"
    DataGrid1.Columns(i).Width = 1700
    
    'Nombre Artículo
    i = i + 1
    DataGrid1.Columns(i).Caption = "Nombre Articulo"
    DataGrid1.Columns(i).Width = 3000
    
    'Cantidad
    i = i + 1
    DataGrid1.Columns(i).Caption = "Cantidad"
    DataGrid1.Columns(i).Width = 1000
    DataGrid1.Columns(i).Alignment = dbgRight
    DataGrid1.Columns(i).NumberFormat = FormatoImporte
    
    
    'Precio
    i = i + 1
    DataGrid1.Columns(i).Caption = "Precio"
    DataGrid1.Columns(i).Width = 1000
    DataGrid1.Columns(i).Alignment = dbgRight
    DataGrid1.Columns(i).NumberFormat = FormatoDec8d4
    
    'Importe
    i = i + 1
    DataGrid1.Columns(i).Caption = "Importe"
    DataGrid1.Columns(i).Width = 1300
    DataGrid1.Columns(i).Alignment = dbgRight
    DataGrid1.Columns(i).NumberFormat = FormatoImporte
    
    
    'tipo Movimiento
    i = i + 1
    DataGrid1.Columns(i).Caption = "T.Mov."
    DataGrid1.Columns(i).Width = 700
    DataGrid1.Columns(i).Alignment = dbgCenter
    
    
    
    'Observaciones
    i = i + 1
    DataGrid1.Columns(i).Caption = "Observaciones"
    DataGrid1.Columns(i).Width = 4050
       
    For i = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(i).AllowSizing = False
    Next i
'--monica: rollo toolbar
'    DataGrid1.Enabled = b
    DataGrid1.ScrollBars = dbgAutomatic
    
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub

'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim i As Byte
Dim alto As Single

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For i = 0 To txtAux.Count - 1
            txtAux(i).Top = 290
        Next i
        Me.cmdAux.Top = 290
        Me.cboAux.Top = 290
    Else
        DeseleccionaGrid Me.DataGrid1
        CargarComboAux
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            For i = 0 To txtAux.Count - 1
                txtAux(i).Text = ""
                If i <> 1 Then txtAux(i).Locked = False
            Next i
            cmdAux.Enabled = True
            cboAux.Enabled = True
            cboAux.ListIndex = -1
        Else  'Poner valor a los txtAux
            For i = 0 To 2
                txtAux(i).Text = DataGrid1.Columns(i + 3).Text
            Next i
            Select Case DataGrid1.Columns(8).Value
                Case "S"
                    Me.cboAux.ListIndex = 0
                Case "E"
                    Me.cboAux.ListIndex = 1
            End Select
'            PosicionarCombo Me.Combo1, Me.cboAux.ListIndex
            txtAux(3).Text = DataGrid1.Columns(9).Text
            txtAux(0).Locked = True
            cmdAux.Enabled = False
            cboAux.Enabled = True
            txtAux(2).Locked = False
            txtAux(3).Locked = False
            
            txtAux(4).Locked = False
            txtAux(5).Locked = False
            
            txtAux(4).Text = DataGrid1.Columns(6).Text
            txtAux(5).Text = DataGrid1.Columns(7).Text
        End If
        
        If DataGrid1.Row < 0 Then
            alto = DataGrid1.Top + 220
        Else
            alto = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + 10
        End If
        
        'Fijamos altura y posición Top
        For i = 0 To txtAux.Count - 1
            txtAux(i).Top = alto
            txtAux(i).Height = DataGrid1.RowHeight
        Next i
        Me.cmdAux.Top = alto
        Me.cmdAux.Height = DataGrid1.RowHeight
        cboAux.Top = alto - 5
        
        'Fijamos anchura y posicion Left
        txtAux(0).Left = DataGrid1.Left + 340 'codartic
        txtAux(0).Width = DataGrid1.Columns(3).Width - 200
        cmdAux.Left = txtAux(0).Left + txtAux(0).Width
        txtAux(1).Left = cmdAux.Left + cmdAux.Width  'Nombre Artic
        txtAux(1).Width = DataGrid1.Columns(4).Width - 35
        i = 2 'Cantidad
        txtAux(i).Left = txtAux(i - 1).Left + txtAux(i - 1).Width + 25
        txtAux(i).Width = DataGrid1.Columns(i + 3).Width - 20
        ' precio
        i = 3
        txtAux(4).Left = txtAux(2).Left + txtAux(2).Width + 25
        txtAux(4).Width = DataGrid1.Columns(i + 3).Width - 20
        
        ' importe
        i = 4
        txtAux(5).Left = txtAux(4).Left + txtAux(4).Width + 25
        txtAux(5).Width = DataGrid1.Columns(i + 3).Width - 20
        
        'Tipo Movimiento
        cboAux.Left = txtAux(5).Left + txtAux(5).Width + 20
        cboAux.Width = DataGrid1.Columns(8).Width  '+ 10
        
        i = 3 'Observac
        txtAux(i).Left = cboAux.Left + cboAux.Width + 30
        txtAux(i).Width = DataGrid1.Columns(9).Width - 60
    End If

    'Los ponemos Visibles o No
    For i = 0 To txtAux.Count - 1
        txtAux(i).visible = visible
    Next i
    cmdAux.visible = visible
    cboAux.visible = visible
End Sub

Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Almacen Propios
Dim indice As Byte
    indice = CByte(Me.imgBuscar(0).Tag)
    Text1(indice + 6).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    Text2(indice + 6).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Articulos
    txtAux(0).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Artic
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        
        If Modo <> 5 Then 'Estamos en Cabecera
            'Recupera todo el registro de Traspaso Almacenes
            'Sabemos que campos son los que nos devuelve
            'Creamos una cadena consulta y ponemos los datos
            CadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            CadB = Aux
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
            PonerCadenaBusqueda
        Else 'Estamos en Lineas
            'Llamamos desde el boton auxiliar de Artículos
            txtAux(0).Text = RecuperaValor(CadenaDevuelta, 1)
            txtAux(1).Text = RecuperaValor(CadenaDevuelta, 2)
            PonerFoco txtAux(2)
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub




Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de clientes
    Text1(3).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod cliente
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom cliente
End Sub

Private Sub frmF_Selec(vFecha As Date)
Dim indice As Byte
    indice = CByte(Me.imgFec(0).Tag) + 1
    Text1(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de socio
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod socio
    Text2(2).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom socio
End Sub

Private Sub imgBuscar_Click(Index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
 
    Screen.MousePointer = vbHourglass
    imgBuscar(0).Tag = Index
    
    Select Case Index
        Case 0 'Codigo Almacen
            Set frmA = New frmManAlmProp
            frmA.DatosADevolverBusqueda = "0|1|"
            frmA.Show vbModal
            Set frmA = Nothing
        
        Case 1  'Cod. socio
            Set frmSoc = New frmBasico
            
            AyudaSocios frmSoc, Text1(2).Text
            
            Set frmSoc = Nothing
            PonerFoco Text1(2)
        
        Case 2 ' cod. cliente
            Set frmCli = New frmClientes
            frmCli.DatosADevolverBusqueda = "0|2|"
            frmCli.Show vbModal
            Set frmCli = Nothing
            PonerFoco Text1(3)
        
    End Select
    PonerFoco Text1(Index)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFec_Click(Index As Integer)
Dim indice As Byte
   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   
   '++monica
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim obj As Object

    Set frmF = New frmCal
    
    esq = imgFec(Index).Left
    dalt = imgFec(Index).Top
    
    Set obj = imgFec(Index).Container

    While imgFec(Index).Parent.Name <> obj.Name
        esq = esq + obj.Left
        dalt = dalt + obj.Top
        Set obj = obj.Container
    Wend
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    frmF.Left = esq + imgFec(Index).Parent.Left + 30
    frmF.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40

   
    frmF.NovaData = Now
    indice = Index + 1
    Me.imgFec(0).Tag = Index
   
    PonerFormatoFecha Text1(indice)
    If Text1(indice).Text <> "" Then frmF.NovaData = CDate(Text1(indice).Text)

    Screen.MousePointer = vbDefault
    frmF.Show vbModal
    Set frmF = Nothing
    PonerFoco Text1(indice)
End Sub


Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    If Modo = 5 Then   'Eliminar lineas Movimiento Almacenes
        BotonEliminarLinea
    Else 'Eliminar Cabecera Movimiento Almacenes
        BotonEliminar
    End If
End Sub

Private Sub mnModificar_Click()
Dim vWhere As String

    If Modo = 5 Then  'Modificar LINEAS
        vWhere = ObtenerWhereCP(False) & " and numlinea=" & Me.Data2.Recordset.Fields(1)
        If BloqueaRegistro(NomTablaLineas, vWhere) Then BotonModificarLinea
    Else 'Modificar Cabecera
       If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub


Private Sub mnNuevo_Click()
     If Modo = 5 Then  'Añadir lineas Movimiento Almacenes
        BotonAnyadirLineas
    Else 'Añadir Cabecera Movimiento Almacenes
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
    If Index <> 4 Then ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 And Index = 3 And Modo = 1 Then
        PonerFocoBtn cmdAceptar
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    'Bloquear el contador si no estamos en busquedas
    If (Modo <> 1) And (Index = 0) Then BloquearTxt Text1(0), True, True

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub



    Select Case Index
        Case 0 'Codigo Movimiento Almacen
            Text1(Index).Text = Format(Text1(Index).Text, "0000000")
            
        Case 1 'Fecha
            PonerFormatoFecha Text1(Index)
            
        Case 2  'Codigo socio
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rsocios", "nomsocio", "codsocio", "N")
            
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Socio: " & Text1(Index).Text & ". Revise" & vbCrLf
                    MsgBox cadMen, vbExclamation
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 3  'Codigo cliente
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "clientes", "nomclien", "codclien", "N")
                
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Cliente: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCli = New frmClientes
                        frmCli.DatosADevolverBusqueda = "0|1|"
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmCli.Show vbModal
                        Set frmCli = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 6 'Codigo Almacen
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "salmpr", "nomalmac", "codalmac")
            Else
                Text2(Index).Text = ""
            End If
            
        Case 4 'Observaciones
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

Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Actualizar
           BotonActualizar
    End Select
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
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

    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub

    Select Case Index
        Case 0 'Cod ARTICULO
            If txtAux(Index).Text = "" Then
                txtAux(Index + 1).Text = ""
            Else
                 PonerArticulo txtAux(0), txtAux(1), Text1(6).Text, CodTipoMov, ModificaLineas
                 
                 PonerPrecio txtAux(0), Text1(3).Text
            End If
            
        Case 2 'CANTIDAD (Comprobamos formato como si fuera un Importe)
            'Formato tipo 1: Decimal(12,2)
            PonerFormatoDecimal txtAux(Index), 1
            
        Case 4 ' precio
            PonerFormatoDecimal txtAux(Index), 7
            
        Case 5 ' importe
            PonerFormatoDecimal txtAux(Index), 3
            
    End Select
    
    If Index = 2 Or Index = 4 Then
        If txtAux(2).Text <> "" And txtAux(4).Text <> "" Then
            txtAux(5).Text = CalcularImporte(txtAux(2).Text, txtAux(4).Text, "", "", 0, 0)
        End If
    End If
End Sub

Private Sub PonerPrecio(Articulo As String, Cliente As String)
Dim Sql As String
Dim Precio As Currency

    On Error Resume Next

    Sql = "select precioar from clientes_precio where codclien =  " & DBSet(Cliente, "N")
    Sql = Sql & " and codartic = " & DBSet(Articulo, "T")
    
    If TotalRegistrosConsulta(Sql) = 0 Then
        Sql = "select preciove from sartic where codartic = " & DBSet(Articulo, "T")
    End If
    
    Precio = DevuelveValor(Sql)
    
    txtAux(4).Text = Format(Precio, "###,##0.0000")

End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 5 'Busqueda
           mnBuscar_Click
        Case 6 'Ver Todos
           mnVerTodos_Click
        Case 1 'Nuevo
           mnNuevo_Click
        Case 2  'Modificar
           mnModificar_Click
        Case 3 'Eliminar
           mnEliminar_Click
        Case 8 'Imprimir
           BotonImprimir
           
'        Case 9 'Mantenimiento Lineas
'           BotonLineas
'        Case 10 'Actualizar
'           BotonActualizar
'        Case 12 'Imprimir
'           BotonImprimir
'        Case 13  'Salir
'           mnSalir_Click
'        Case btnPrimero To btnPrimero + 3 'Flechas de Desplazamiento
'           Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte, Numreg As Byte
Dim b As Boolean
    
    'Actualiza Iconos Insertar,Modificar,Eliminar
'--monica: rollo toolbar
'    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo, ModificaLineas
    
    '--------------------------------------------
    b = (Kmodo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If hcoCodMovim <> -1 Then
        cmdRegresar.visible = b
    Else
        cmdRegresar.visible = False
    End If
    
    Numreg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then Numreg = 2 'Solo es para saber q hay + de 1 registro
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, Numreg
    DesplazamientoVisible b And Data1.Recordset.RecordCount > 1


    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    
    'Como el campo 0 es clave primaria, NO se puede modificar, es contador
    BloquearTxt Text1(0), (Modo <> 1), True
    Text1(0).Enabled = (Modo = 1)
    Combo1(0).Enabled = (Modo = 1) Or (Modo = 3)
    
    '=================================================
    b = Modo <> 0 And Modo <> 2 '--monica: rollo toolbar And Modo <> 5
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    For i = 0 To Me.imgFec.Count - 1
        Me.imgFec(i).Enabled = b And Modo <> 5
    Next i
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = b And Modo <> 5
    Next i

    Me.chkVistaPrevia.Enabled = (Modo <= 2)

    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu 'Activar opciones de menu según Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
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

Private Sub PonerModoOpcionesMenu()
Dim b As Boolean, bAux As Boolean
Dim i As Byte

    'Si visualizamos el historico no mostrar botones de Mantenimiento, solo es consulta
    For i = 1 To 3
        '++monica: rollo toolbar he puesto condicion
        If i <> 9 Then Toolbar1.Buttons(i).Enabled = Not EsHistorico
    Next i
    Me.mnNuevo.visible = Not EsHistorico
    Me.mnModificar.visible = Not EsHistorico
    Me.mnEliminar.visible = Not EsHistorico
    Me.mnBarra2.visible = Not EsHistorico
    
    
    If Not EsHistorico Then
        'Modo 2. Hay datos y estamos visualizandolos
        b = (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
        'Insertar
        Toolbar1.Buttons(1).Enabled = (b Or Modo = 0)
        Me.mnNuevo.Enabled = (b Or Modo = 0)
        'Modificar
        Toolbar1.Buttons(2).Enabled = b
        Me.mnModificar.Enabled = b
        'eliminar
        Toolbar1.Buttons(3).Enabled = b
        Me.mnEliminar.Enabled = b
        
        '--------------------------------
        b = (Modo = 2)
'--monica: rollo toolbar
'        'Lineas Movimientos Almacenes
'        Toolbar1.Buttons(9).Enabled = b

        'Actualizar
        Toolbar5.Buttons(1).Enabled = b
        
        
        b = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(5).Enabled = Not b
        Me.mnBuscar.Enabled = Not b
        'Ver Todos
        Toolbar1.Buttons(6).Enabled = Not b
        Me.mnVerTodos.Enabled = Not b
    End If
    
    '++monica: rollo toolbar
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not EsHistorico
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = b 'And Me.Data2.Recordset.RecordCount < 10
        If b Then bAux = (b And Me.Data2.Recordset.RecordCount > 0)
        ToolAux(i).Buttons(2).Enabled = (Modo = 3 Or Modo = 4 Or Modo = 2) And Me.Data2.Recordset.RecordCount > 0 'bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i
    
End Sub



Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
    Me.chkImpresion(0).Value = 0
    Me.Combo1(0).ListIndex = -1
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
    
    Select Case Modo
        Case 5 'Modo Mantenimiento de Almacenes (Lineas)
            If Data2.Recordset.EOF Then Exit Sub
            DesplazamientoData Data2, Index, True
        Case Else 'Datos de Cabecera
            If Data1.Recordset.EOF Then Exit Sub
            DesplazamientoData Data1, Index, True
            PonerCampos
    End Select
End Sub


Private Function MontaSQLCarga(enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim Sql As String
Dim tabla As String
    
    tabla = NomTablaLineas

    Sql = "SELECT " & tabla & ".codservi, " & tabla & ".clisoc, "
    Sql = Sql & tabla & ".numlinea, " & tabla & ".codartic, Articulos.nomartic, "
    Sql = Sql & tabla & ".cantidad, " & tabla & ".precioar, " & tabla & ".importel,  if(" & tabla & ".tipomovi=0,""S"",""E"") as tipomovi, "
    Sql = Sql & tabla & ".motimovi "
    Sql = Sql & " FROM ((" & tabla & " LEFT JOIN sartic AS Articulos ON " & tabla & ".codartic ="
    Sql = Sql & " Articulos.codartic))"
    If enlaza Then
        Sql = Sql & " WHERE codservi = " & Data1.Recordset!codservi & " and clisoc = " & Data1.Recordset!clisoc
    Else
        Sql = Sql & " WHERE codservi = -1"
    End If
    Sql = Sql & " ORDER BY " & tabla & ".numlinea"
    MontaSQLCarga = Sql
End Function


Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False

        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFocoCmb Combo1(0)
        Combo1(0).BackColor = vbYellow
        Combo1(0).ListIndex = -1
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
'Ver todos
    
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
    
End Sub


Private Sub BotonLineas()
On Error GoTo ErrorLineas

    Screen.MousePointer = vbHourglass
    PonerModo (5)
    ModificaLineas = 0
    PonerBotonCabecera True
    CargaGrid True
    DataGrid1.Enabled = True
    Me.DataGrid1.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorLineas:
    If Err.Number <> 0 Then MuestraError Err.Number, "Lineas"
    Screen.MousePointer = vbDefault
End Sub


Private Sub BotonAnyadir()
Dim NomTraba As String

    LimpiarCampos 'Vacía los TextBox
    ModoAnterior = Modo 'Para el botón Cancelar en Modo Insertar
    PonerModo 3
           
    'Ponemos el grid lineas Traspaso enlazando a ningun sitio
    CargaGrid False
    
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    'Poner Trabajador por defecto el trabajador conectado

'**quitado
'    Text1(3).Text = PonerTrabajadorConectado(NomTraba)
'    Text2(1).Text = NomTraba
    Combo1(0).ListIndex = 1
    PonerFocoCmb Combo1(0)
    Combo1_Change (0)
    
    PonerFoco Text1(1)
    
    Text1(6).Text = vParamAplic.Almacen
    If Text1(6).Text <> "" Then
        Text2(6).Text = PonerNombreDeCod(Text1(6), "salmpr", "nomalmac")
    End If
End Sub


Private Sub BotonAnyadirLineas()
Dim vWhere As String
    
    If Me.Data2.Recordset.RecordCount >= 10 Then
        MsgBox "Sólo se permiten un máximo de 10 líneas por albarán para que quepa en la impresión." & vbCrLf & vbCrLf & "Cree un nuevo albarán con el resto de movimientos.", vbExclamation
        Exit Sub
    End If
    
    
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
    cmdAceptar.Tag = SugerirCodigoSiguienteStr("sliser", "numlinea", vWhere)
'--monica: rollo toolbar
'    PonerBotonCabecera False
'    lblIndicador.Caption = "INSERTAR"
    
    'Situamos el grid al final
    AnyadirLinea DataGrid1, Data2

    DataGrid1.Enabled = False
    CargaTxtAux True, True
    PonerFoco txtAux(0)
End Sub


Private Sub BotonModificar()
    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    
    'Como el campo 1 es clave primaria, NO se puede modificar
    BloquearTxt Text1(0), True, True
    PonerFoco Text1(1)
End Sub


Private Sub BotonModificarLinea()
Dim i As Integer


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
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    cmdAceptar.Tag = Data2.Recordset!NumLinea
    
    CargaTxtAux True, False
    DataGrid1.Enabled = False
    PonerFoco txtAux(2) 'Poner el foco
    Screen.MousePointer = vbDefault
End Sub


Private Sub BotonEliminar()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim Sql As String
Dim Nombre As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    Sql = "Cabecera de Movimiento Servicios Varios." & vbCrLf
    Sql = Sql & "----------------------------------------" & vbCrLf & vbCrLf
    
    Sql = Sql & "Va a eliminar el Movimiento:"
    Sql = Sql & vbCrLf & " Nº Movim. : " & Text1(0).Text
    Sql = Sql & vbCrLf & " Fecha Mov.: " & CStr(Data1.Recordset.Fields(3))
    Sql = Sql & vbCrLf & " Almacen   : " & Text1(6).Text
    If Data1.Recordset!clisoc = 0 Then
        Sql = Sql & vbCrLf & " Socio     : " & Format(Data1.Recordset!CodSocio, "000000") & " " & Text2(2).Text
    Else
        Nombre = PonerNombreDeCod(Text1(3), "clientes", "nomclien", Data1.Recordset!CodClien, "N")
        Sql = Sql & vbCrLf & " Cliente   : " & Format(Data1.Recordset!CodClien, "000000") & " " & Text2(3).Text
    End If
    
    Sql = Sql & vbCrLf & vbCrLf & " ¿Desea continuar ? "
    
    If MsgBox(Sql, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        If Not Eliminar Then Exit Sub
    
        'Devolvemos contador, si no estamos actualizando
        Set vTipoMov = New CTiposMov
        NumRegElim = Data1.Recordset.Fields(0)
        vTipoMov.DevolverContador CodTipoMov, NumRegElim
        Set vTipoMov = Nothing
        
        NumRegElim = Data1.Recordset.AbsolutePosition
        DataGrid1.Enabled = False
        
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
'++monica: rollo
            PonerModo 2
            PonerCampos
        Else 'solo habia un registro
            LimpiarCampos
            CargaGrid False
            PonerModo 0
        End If
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Movimiento", Err.Description
        Data1.Recordset.CancelUpdate
    End If
End Sub


Private Function Eliminar() As Boolean
Dim Sql As String
On Error GoTo FinEliminar
        
        conn.BeginTrans
        Sql = " WHERE  codservi=" & Data1.Recordset!codservi & " and clisoc = " & Data1.Recordset!clisoc
        
        'Lineas
        conn.Execute "Delete  from sliser " & Sql
        
        'Cabeceras
        conn.Execute "Delete  from scaser " & Sql
                      
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
Dim Sql As String
On Error GoTo Error2
    
    
    'Ciertas comprobaciones
    If Data2.Recordset.EOF Then Exit Sub
    
    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar
       
    ModificaLineas = 3 'Eliminar
    
    '++monica: rollo toolbar
'    PonerModo (5)
    DataGrid1.Enabled = True
    Me.DataGrid1.SetFocus
    '++monica
    
    '### a mano
    Sql = "Seguro que desea eliminar la línea del Artículo:"
    Sql = Sql & vbCrLf & "Código: " & Data2.Recordset!CodArtic
    Sql = Sql & vbCrLf & "Descripción: " & Data2.Recordset.Fields(3)
    
    If MsgBox(Sql, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        Sql = "Delete from sliser where codservi=" & Data2.Recordset!codservi
        Sql = Sql & " and clisoc = " & Data2.Recordset!clisoc
        Sql = Sql & " and numlinea=" & Data2.Recordset!NumLinea
        Sql = Sql & " and codartic=" & DBSet(Data2.Recordset!CodArtic, "T")
        '++ monica: rollo
        NumRegElim = Data2.Recordset.AbsolutePosition
        TerminaBloquear
        
        conn.Execute Sql
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
    
    PonerCampos
    
Error2:
    Screen.MousePointer = vbDefault
    ModificaLineas = 0
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Línea de Artículo de Movimiento Almacen", Err.Description
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
'Dim vStock As String
'Dim vstockOrig As Single  'Stock en el almacen Origen
'Dim SQL As String, devuelve As String

    DatosOk = False
    b = CompForm(Me)
    If Not b Then Exit Function
    
    If Modo = 3 Then
        If Combo1(0).ListIndex = 0 Or Combo1(0).ListIndex = 2 Then
            If Text1(2).Text = "" Then
                MsgBox "Debe introducir un Socio para el movimiento. Revise."
                PonerFoco Text1(2)
                Exit Function
            Else
                Text2(2).Text = DevuelveDesdeBDNew(cAgro, "rsocios", "nomsocio", "codsocio", Text1(2).Text, "N")
                If Text2(2).Text = "" Then
                    MsgBox "Socio no existe. Revise.", vbExclamation
                    PonerFoco Text1(2)
                    Exit Function
                End If
            End If
        Else
            If Text1(3).Text = "" Then
                MsgBox "Debe introducir un Cliente para el movimiento. Revise."
                PonerFoco Text1(3)
                Exit Function
            Else
                Text2(3).Text = DevuelveDesdeBDNew(cAgro, "clientes", "nomclien", "codclien", Text1(3).Text, "N")
                If Text2(3).Text = "" Then
                    MsgBox "Cliente existe. Revise.", vbExclamation
                    PonerFoco Text1(3)
                    Exit Function
                End If
            End If
        End If
    End If

    'Comprobar que todos los Artículos estan en el nuevo almacen
    If Modo = 4 Then 'Modificando
        b = ComprobarStocksLineas
    End If

    DatosOk = True
End Function



Private Function ComprobarStocksLineas() As Boolean
'Comprobar para todas las lineas del traspaso que:
' - todos los Artículos entan en el almacen origen
' - Comprobar que hay suficiente stock en el Almacen Origen de ese Articulo
Dim b As Boolean

    If Not Data2.Recordset.EOF Then  'Si hay lineas
        Data2.Recordset.MoveFirst
        b = True
        
        While Not Data2.Recordset.EOF And b
            If Data2.Recordset!tipomovi = "S" Then 'Mov. de salida
                b = ComprobarStock(Data2.Recordset!CodArtic, Text1(6).Text, Data2.Recordset!Cantidad, CodTipoMov)
            End If
            Data2.Recordset.MoveNext
        Wend
        Data2.Recordset.MoveFirst
    End If
    ComprobarStocksLineas = b
End Function




Private Function DatosOkLinea() As Boolean
Dim b As Boolean
Dim devuelve As String

    DatosOkLinea = False
    b = True
        
    If txtAux(0).Text = "" Then
        MsgBox "El campo Cod. Artículo no puede ser nulo", vbExclamation
        b = False
        Exit Function
    End If
        
    'Comprobamos el campo Cantidad
    If txtAux(2).Text = "" Then
         MsgBox "El campo Cantidad no puede ser nulo", vbExclamation, "Artículos"
         b = False
    ElseIf Not IsNumeric(txtAux(2).Text) Then
        MsgBox "El campo Cantidad debe ser numérico", vbExclamation
        b = False
    End If
    If Not b Then
        PonerFoco txtAux(2)
        Exit Function
    End If
     
    'Comprobamos si ya existe una linea con el artículo, solo si estamos insertando (ModificaLineas=1)
    'BD 1: conexion a BD Ariges
    If ModificaLineas = 1 Then
        devuelve = DevuelveDesdeBDNew(cAgro, "sliser", "codservi", "codservi", Text1(0).Text, "N", , "clisoc", Combo1(0).ListIndex, "N", "codartic", txtAux(0).Text, "T")
        If devuelve <> "" Then
            b = False
            devuelve = "Ya hay una línea con ese Artículo: " & vbCrLf
            devuelve = devuelve & "Codigo: " & txtAux(0).Text & vbCrLf
            devuelve = devuelve & "Descripción: " & txtAux(1).Text
            MsgBox devuelve, vbExclamation
        End If
        
        'Comprobamos si existe el artículo, solo si estamos insertando (ModificaLineas=1)
        If Trim(txtAux(1).Text) = "" Then
            b = False
            devuelve = "No existe el Artículo " & vbCrLf
            devuelve = devuelve & "Codigo: " & txtAux(0).Text & vbCrLf
            devuelve = devuelve & "Descripción: " & txtAux(1).Text
            MsgBox devuelve, vbExclamation
        End If
    End If
    If Not b Then Exit Function
    
    'Comprobar que hay suficiente stock en el Almacen
    'Si es movimiento de Salida
    If Me.cboAux.ListIndex = 0 Then
        b = ComprobarStock(txtAux(0).Text, Text1(6).Text, txtAux(2).Text, CodTipoMov)
    End If
    DatosOkLinea = b
End Function


Private Sub PonerBotonCabecera(b As Boolean)
On Error Resume Next
    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdRegresar.visible = b
    Me.cmdRegresar.Caption = "Cabecera"
    If b Then
        Me.lblIndicador.Caption = "Lineas Detalle"
        PonerFocoBtn Me.cmdRegresar
    Else
        Me.lblIndicador.Caption = ""
    End If
    'Habilitar las opciones correctas del menu según Modo
    PonerModoOpcionesMenu
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu según Nivel de Acceso
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Function InsertarModificarLinea() As Boolean
Dim Sql As String, Cad As String
On Error GoTo EInsertarModificarLinea
    
    Sql = ""
    InsertarModificarLinea = False
    
    Select Case ModificaLineas
    Case 1 'Insertar
        If DatosOkLinea Then 'INSERTAR
            Sql = "INSERT INTO sliser (codservi,clisoc,numlinea,codartic,cantidad,tipomovi,motimovi, precioar, importel) "
            Sql = Sql & " VALUES (" & Val(Text1(0).Text) & ", "
            Sql = Sql & Combo1(0).ListIndex & ","
            Sql = Sql & cmdAceptar.Tag & ", "
            Sql = Sql & DBSet(txtAux(0).Text, "T") & ", "
            Sql = Sql & DBSet(txtAux(2).Text, "N") & ", "
            If cboAux.ListIndex = -1 Then
                Cad = ValorNulo
            Else
                 Cad = cboAux.ItemData(cboAux.ListIndex)
            End If
            Sql = Sql & CSng(Cad) & ","
            Sql = Sql & DBSet(txtAux(3).Text, "T") & ", "
            Sql = Sql & DBSet(txtAux(4).Text, "N") & ", "
            Sql = Sql & DBSet(txtAux(5).Text, "N") & ") "
        End If
    Case 2 'Modificar
        If DatosOkLinea Then
            If Not EsHistorico Then
                Sql = "UPDATE sliser Set cantidad = " & DBSet(txtAux(2).Text, "N")
                Sql = Sql & ", tipomovi = " & cboAux.ItemData(cboAux.ListIndex)
                Sql = Sql & ", motimovi = " & DBSet(txtAux(3).Text, "T")
                Sql = Sql & ", precioar = " & DBSet(txtAux(4).Text, "N")
                Sql = Sql & ", importel = " & DBSet(txtAux(5).Text, "N")
                Sql = Sql & " WHERE codservi =" & Val(Text1(0).Text) & " AND "
                Sql = Sql & " clisoc = " & Combo1(0).ListIndex & " AND "
                Sql = Sql & " numlinea =" & Val(cmdAceptar.Tag)
            Else
                InsertarModificarLinea = ModificarLineaHco
                Exit Function
            
'                SQL = "UPDATE slhser Set cantidad = " & DBSet(txtAux(2).Text, "N")
'                SQL = SQL & ", tipomovi = " & cboAux.ItemData(cboAux.ListIndex)
'                SQL = SQL & ", motimovi = " & DBSet(txtAux(3).Text, "T")
'                SQL = SQL & ", precioar = " & DBSet(txtAux(4).Text, "N")
'                SQL = SQL & ", importel = " & DBSet(txtAux(5).Text, "N")
'                SQL = SQL & " WHERE codservi =" & Val(text1(0).Text) & " AND "
'                SQL = SQL & " clisoc = " & Combo1(0).ListIndex & " AND "
'                SQL = SQL & " numlinea =" & Val(cmdAceptar.Tag)
            End If
        End If
    End Select
            
    If Sql <> "" Then
        conn.Execute Sql
        InsertarModificarLinea = True
    End If
    Exit Function
EInsertarModificarLinea:
    MuestraError Err.Number, "Insertar/Modificar Lineas de Servicios" & vbCrLf & Err.Description
End Function


Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim tabla As String
Dim Titulo As String

    'Llamamos a al form
    Cad = ""
    'Registro de la tabla de cabeceras: scaser
    Cad = Cad & ParaGrid(Text1(0), 15, "Nº Mov.")
    
    If NombreTabla = "scaser" Then
        Cad = Cad & "Tipo|if(scaser.clisoc=0,'Socio','Cliente')|N||10·"
    Else
        Cad = Cad & "Tipo|if(schser.clisoc=0,'Socio','Cliente')|N||10·"
    End If
    
    Cad = Cad & ParaGrid(Text1(1), 20, "Fecha")
    Cad = Cad & "Almacen|salmpr.codalmac|N||10·"
    Cad = Cad & "Desc. Alm. Orig|nomalmac|T||40·"

    tabla = "(" & NombreTabla & " LEFT JOIN salmpr ON " & NombreTabla & ".codalmac=salmpr.codalmac" & ") "
    Titulo = Me.Caption

    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vtabla = tabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = Titulo
        frmB.vSelElem = 0
'**quitado
'        frmB.vConexionGrid = cAgro 'Conexion a BD Ariges

'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If HaDevueltoDatos Then
''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            If Modo = 5 Then
'                PonerFoco txtAux(0)
'            Else
                PonerFoco Text1(kCampo)
'            End If
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub HacerBusqueda()
Dim CadB As String

    CadB = ObtenerBusqueda(Me, False)
    cadSeleccion = ObtenerBusqueda3(Me, True) 'Para la consulta de report

    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then 'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        MsgBox "Introducir criterios de búsqueda", vbExclamation
        PonerFoco Text1(0)
    End If
    
End Sub


Private Sub PonerCadenaBusqueda()
On Error GoTo EEPonerBusq
    Screen.MousePointer = vbHourglass

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        PonerFoco Text1(0)
        Exit Sub
    Else
        PonerModo 2
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

Private Sub PonerCampos()
On Error GoTo EPonerCampos

    If Data1.Recordset.EOF Then Exit Sub
    
    PonerCamposForma Me, Data1
    Text2(6).Text = PonerNombreDeCod(Text1(6), "salmpr", "nomalmac")
    Text2(2).Text = PonerNombreDeCod(Text1(2), "rsocios", "nomsocio")
    Text2(3).Text = PonerNombreDeCod(Text1(3), "clientes", "nomclien")
'    Text2(1).Text = PonerNombreDeCod(Text1(3), conAri, "straba", "nomtraba")
    CargaGrid True
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerSocioVisible
    
    PonerModoOpcionesMenu
    PonerOpcionesMenu
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Function ActualizarStocks() As Boolean
Dim Sql As String, EnAlmDest As String
Dim Cantidad As Single, vStock As Single
Dim devuelve As String
    
    On Error GoTo EActualizarStock

    ActualizarStocks = False
    While Not Data2.Recordset.EOF
        'Actualizar el stock si el articulo tiene control de stock
        devuelve = DevuelveDesdeBDNew(cAgro, "sartic", "ctrstock", "codartic", Data2.Recordset!CodArtic, "T")
        If Val(devuelve) = 1 Then 'Hay control de stock

            Cantidad = Data2.Recordset!Cantidad 'Cant a traspasar
            
            If Data2.Recordset!tipomovi = "E" Then 'Mov. de Entrada
                '==== Aumentar el stock en el Almacen
                'Comprobar que existe el articulo en Almacen Destino
                EnAlmDest = DevuelveDesdeBDNew(cAgro, "salmac", "codartic", "codartic", Data2.Recordset!CodArtic, "T", , "codalmac", Text1(6).Text, "N")
                If EnAlmDest = "" Then 'No hay de ese artículo en Almacen
                    Sql = "INSERT INTO salmac (codartic,codalmac,ubialmac,canstock,stockmin,puntoped,stockmax,stockinv,fechainv,horainve,statusin)"
                    Sql = Sql & " VALUES (" & DBSet(Data2.Recordset!CodArtic, "T") & "," & Val(Text1(6).Text) & ",''," & DBSet(Cantidad, "N") & ",0,0,0,0,NULL,NULL,0)"
                Else 'Existe el artic en almac. Dest -> Aumentar stock
                    Sql = "UPDATE salmac Set canstock = canstock + " & DBSet(Cantidad, "N") ' ++monica:añadido el dbset, fallaba en decimales
                    Sql = Sql & " WHERE codartic =" & DBSet(Data2.Recordset!CodArtic, "T") & " AND "
                    Sql = Sql & " codalmac =" & Data1.Recordset!codAlmac
                End If
                
            Else 'Mov. de Salida
                '==== Disminuir Stock en Almacen Origen
                EnAlmDest = DevuelveDesdeBDNew(cAgro, "salmac", "canstock", "codartic", Data2.Recordset!CodArtic, "T", , "codalmac", Text1(6).Text, "N")
                If EnAlmDest = "" Then 'No hay de ese artículo en Almacen
                    devuelve = "No existe en el Almacen: " & Data1.Recordset!codAlmac & vbCrLf
                    devuelve = devuelve & "El Artículo: " & Data2.Recordset!CodArtic
                    MsgBox devuelve, vbExclamation
                Else 'Existe el artic en almac. Dest -> Disminuir stock
                    vStock = CLng(EnAlmDest)
                    If ComprobarHayStock(vStock, Cantidad, Data2.Recordset!CodArtic, Data2.Recordset!NomArtic, CodTipoMov) Then
                        Sql = "UPDATE salmac Set canstock = canstock - " & DBSet(Cantidad, "N") '++monica:añadido el dbset, fallaba en decimales
                        Sql = Sql & " WHERE codartic =" & DBSet(Data2.Recordset!CodArtic, "T") & " AND "
                        Sql = Sql & " codalmac =" & Data1.Recordset!codAlmac
                    End If
                End If
            End If
            
            conn.Execute Sql
        End If
        Data2.Recordset.MoveNext
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
Dim Sql As String

    If Data1.Recordset.EOF Then
        MsgBox "Ningún Movimiento para actualizar.", vbExclamation
        Exit Sub
    End If
    
    If Data2 Is Nothing Then Exit Sub
    If Data2.Recordset.EOF Then
        MsgBox "No hay lineas insertadas para este Nº de Movimiento", vbExclamation
        Exit Sub
    End If
    
    If Not CBool(Data1.Recordset.Fields(7).Value) And (Data1.Recordset.Fields(1).Value = 0 Or Data1.Recordset.Fields(1).Value = 1) Then 'Informe No Impreso
        Sql = "Actualización Movimientos Servicios Varios." & vbCrLf
        Sql = Sql & "-------------------------------------------" & vbCrLf & vbCrLf
        Sql = Sql & "NO ESTA IMPRESO EL MOVIMIENTO:" & vbCrLf
        Sql = Sql & vbCrLf & "Nº Movim. : " & Format(Data1.Recordset.Fields(0), "0000000")
        Sql = Sql & vbCrLf & "Fecha        : " & CStr(Data1.Recordset.Fields(3))
        Sql = Sql & vbCrLf & "Almacen    : " & Format(Data1.Recordset.Fields(2), "000") & " - " & Text2(6).Text
        Sql = Sql & vbCrLf & " Tipo  : "
        Select Case Data1.Recordset.Fields(1).Value
            Case 0
                Sql = Sql & " Socio " & Format(Data1.Recordset!CodSocio, "000000")
            Case 1
                Sql = Sql & " Cliente " & Format(Data1.Recordset!CodClien, "000000")
            Case 2
                Sql = Sql & " Reg.Socio " & Format(Data1.Recordset!CodSocio, "000000")
            Case 3
                Sql = Sql & " Reg.Cliente " & Format(Data1.Recordset!CodClien, "000000")
        End Select
        Sql = Sql & vbCrLf & vbCrLf & " ¿Desea continuar ? "
        If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub

    Else 'Informe Impreso
        Sql = "Actualización Movimientos Servicios Varios." & vbCrLf
        Sql = Sql & "--------------------------------------------" & vbCrLf & vbCrLf
        
        Sql = Sql & "Va a Actualizar el Movimiento:"
        Sql = Sql & vbCrLf & " Nº Movim.  :  " & Format(Data1.Recordset.Fields(0), "0000000")
        Sql = Sql & vbCrLf & " Fecha Mov.:  " & CStr(Data1.Recordset.Fields(3))
        Sql = Sql & vbCrLf & " Almacen     :  " & CStr(Format(Data1.Recordset.Fields(2), "000"))
        Sql = Sql & vbCrLf & " Tipo  : "
        Select Case Data1.Recordset.Fields(1).Value
            Case 0
                Sql = Sql & " Socio " & Format(Data1.Recordset!CodSocio, "000000")
            Case 1
                Sql = Sql & " Cliente " & Format(Data1.Recordset!CodClien, "000000")
            Case 2
                Sql = Sql & " Reg.Socio " & Format(Data1.Recordset!CodSocio, "000000")
            Case 3
                Sql = Sql & " Reg.Cliente " & Format(Data1.Recordset!CodClien, "000000")
        End Select
        Sql = Sql & vbCrLf & vbCrLf & " ¿Desea continuar ? "
        If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then
            Exit Sub
        End If
    End If
    
    Me.ProgressBar1.visible = True
    Me.ProgressBar1.Value = 0
    
    NumRegElim = Data1.Recordset.AbsolutePosition
    If ActualizarTraspaso Then
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
            PonerModo 2
        Else 'Solo habia un registro
            LimpiarCampos
            CargaGrid False
            PonerModo 0
            espera 0.3
            Me.Refresh
        End If
    
    End If
    Me.ProgressBar1.visible = False
End Sub


Private Function ActualizarTraspaso() As Boolean
Dim Donde As String
Dim devuelve As String
Dim bol As Boolean
On Error GoTo EActualizarTraspaso
    
    'Comprobamos que no existe en historico
    devuelve = DevuelveDesdeBDNew(cAgro, "schser", "codservi", "codservi", Data1.Recordset!codservi, "N", , "clisoc", Data1.Recordset!clisoc, "N", "fecmovim", Data1.Recordset!fecmovim, "F")
    If Trim(devuelve) <> "" Then
        Select Case Data1.Recordset!clisoc
            Case 0
                devuelve = "Ya existe en el histórico el movimiento de Socio:" & vbCrLf
            Case 1
                devuelve = "Ya existe en el histórico el movimiento de Cliente:" & vbCrLf
            Case 2
                devuelve = "Ya existe en el histórico el movimiento de Reg.Socio:" & vbCrLf
            Case 3
                devuelve = "Ya existe en el histórico el movimiento de Reg.Cliente:" & vbCrLf
        End Select
        devuelve = devuelve & " Nº: " & Data1.Recordset!codservi & vbCrLf
        devuelve = devuelve & " Fecha: " & Data1.Recordset!fecmovim
        
        Select Case Data1.Recordset!clisoc
            Case 0
                devuelve = devuelve & " Socio: " & Format(Data1.Recordset!CodSocio, "000000")
            Case 1
                devuelve = devuelve & " Cliente: " & Format(Data1.Recordset!CodClien, "000000")
            Case 2
                devuelve = devuelve & " Socio: " & Format(Data1.Recordset!CodSocio, "000000")
            Case 3
                devuelve = devuelve & " Cliente: " & Format(Data1.Recordset!CodClien, "000000")
        End Select
        
        MsgBox devuelve, vbExclamation
        Exit Function
    End If
    
    If Data1.Recordset!clisoc = 0 Or Data1.Recordset!clisoc = 1 Then
        If Not ComprobarStocksLineas Then Exit Function
    End If
    
    
    'Aqui empieza transaccion
    conn.BeginTrans
    Donde = ""
    bol = ActualizarElTraspaso(Donde)

EActualizarTraspaso:
    If Err.Number <> 0 Or Donde <> "" Then
        devuelve = "Actualizar Movimiento." & vbCrLf & "----------------------------" & vbCrLf
        devuelve = devuelve & Donde
        MuestraError Err.Number, devuelve, Err.Description
        bol = False
    End If
    If bol Then
        conn.CommitTrans
        ActualizarTraspaso = True
    Else
        conn.RollbackTrans
        MuestraError Err.Number, devuelve, Err.Description
    End If
End Function


Private Function ActualizarElTraspaso(ByRef ADonde As String) As Boolean

    ActualizarElTraspaso = False
    
    'Insertamos en cabeceras Historico
    ADonde = "Insertando datos en historico cabeceras movimientos Servicios Varios"
    If Not InsertarCabeceraHistorico Then Exit Function
    IncrementarProgres 2
     
    'Insertamos en lineas Historico
    ADonde = "Insertando datos en Historico lineas Movimientos Servicios Varios"
    If Not InsertarLineasHistorico Then Exit Function
    IncrementarProgres 2
    
    
     'Modificar stock
    If Data1.Recordset!clisoc = 0 Or Data1.Recordset!clisoc = 1 Then
        ADonde = "Actualizando Stocks Almacenes"
        If Not ActualizarStocks() Then Exit Function
        IncrementarProgres 2
    End If
    
    'Insertamos en Movimientos Artículos
    ADonde = "Insertando datos en Movimientos de Articulos"
    If Not InsertarMovimArticulos Then Exit Function
    IncrementarProgres 2
   
    
    'Borramos cabeceras y lineas del asiento
    ADonde = "Borrar cabeceras y lineas en Movimientos Almacen"
    If Not BorrarTraspaso(False) Then Exit Function
    IncrementarProgres 2
    
    ActualizarElTraspaso = True
    ADonde = ""
End Function


Private Function InsertarCabeceraHistorico() As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
On Error GoTo EInsertarCab

'    SQL = "SELECT codmovim,codalmac,fecmovim,codtraba,observa1 from scaser where "
    Sql = "SELECT codservi,clisoc,codalmac,fecmovim,codclien, codsocio, observa1, matriveh from scaser where "
    Sql = Sql & " codservi =" & Data1.Recordset!codservi
    Sql = Sql & " AND clisoc = " & Data1.Recordset!clisoc
    Sql = Sql & " AND fecmovim='" & Format(Data1.Recordset!fecmovim, "yyyy-mm-dd") & "'"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
'        SQL = "INSERT INTO schser (codmovim, fecmovim,hormovim,codalmac,codtraba,observa1) "
        Sql = "INSERT INTO schser (codservi, clisoc, fecmovim,hormovim,codalmac,codclien, codsocio, observa1, matriveh) "
        Sql = Sql & " VALUES (" & Rs.Fields(0).Value & "," & Rs.Fields(1).Value & ", '" & Format(Rs.Fields(3).Value, "yyyy-mm-dd") & "','"
        Sql = Sql & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', " & Rs.Fields(2).Value & ","
        Sql = Sql & DBSet(Rs.Fields(4).Value, "N") & "," & DBSet(Rs.Fields(5).Value, "N")
        Sql = Sql & ", " & DBSet(Rs.Fields(6).Value, "T") & ","
        Sql = Sql & DBSet(Rs.Fields(7).Value, "T") & ")"
    End If
    Rs.Close
    Set Rs = Nothing
    conn.Execute Sql
   
EInsertarCab:
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
        InsertarCabeceraHistorico = False
    Else
        InsertarCabeceraHistorico = True
    End If
End Function


Private Function InsertarLineasHistorico() As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
On Error GoTo EInsertarLineas

    Sql = "SELECT codservi, numlinea, codartic, cantidad, tipomovi, motimovi, clisoc, precioar, importel from sliser where "
    Sql = Sql & " codservi =" & Data1.Recordset!codservi
    Sql = Sql & " and clisoc = " & Data1.Recordset!clisoc
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Rs.MoveFirst
    While Not Rs.EOF
        Sql = "INSERT INTO slhser (codservi, fecmovim, numlinea, codartic, cantidad, tipomovi, motimovi, clisoc, precioar, importel)"
        Sql = Sql & " VALUES (" & Rs.Fields(0).Value & ", '" & Format(Data1.Recordset!fecmovim, "yyyy-mm-dd") & "', "
        Sql = Sql & Rs.Fields(1).Value & ", " & DBSet(Rs.Fields(2).Value, "T") & ", "
        Sql = Sql & DBSet(Rs.Fields(3).Value, "N") & ", " & Rs.Fields(4).Value
        Sql = Sql & ", '" & Rs.Fields(5).Value & "'," & Rs.Fields!clisoc & ","
        Sql = Sql & DBSet(Rs.Fields(7).Value, "N") & ","
        Sql = Sql & DBSet(Rs.Fields(8).Value, "N") & ")"
        conn.Execute Sql
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
EInsertarLineas:
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
        Rs.Close
        Set Rs = Nothing
        InsertarLineasHistorico = False
    Else
        InsertarLineasHistorico = True
    End If
End Function


Private Function InsertarMovimArticulos() As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim vImporte As Single, vPrecioVenta As String
Dim vTipoMov As CTiposMov
Dim bol As Boolean
Dim Cad As String
On Error GoTo EInsertar

    bol = True
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        'Se han cargado correctamente los valores de la clase
'        SQL = "SELECT scaser.codmovim, codalmac, fecmovim, codtraba, numlinea, codartic, cantidad, tipomovi "
        Sql = "SELECT scaser.codservi, codalmac, fecmovim, numlinea, codartic, cantidad, tipomovi, codsocio, codclien, precioar, importel "
        Sql = Sql & " from scaser LEFT JOIN sliser on scaser.codservi=sliser.codservi and scaser.clisoc=sliser.clisoc "
        Sql = Sql & " WHERE scaser.codservi =" & Data1.Recordset!codservi
        Sql = Sql & " and scaser.clisoc = " & Data1.Recordset!clisoc
        Sql = Sql & " AND fecmovim='" & Format(Data1.Recordset!fecmovim, "yyyy-mm-dd") & "'"
    
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not Rs.EOF
            'Obtener el precio de venta del articulo, si tiene control de stock
            Cad = "ctrstock"
            
'[Monica]22/11/2011: ponemos el importe de la linea
'            '++monica antes solo el pmp
'            If vParamAplic.TipoPrecio = 0 Then 'precio medio ponderado
'                vPrecioVenta = DevuelveDesdeBDNew(cAgro, "sartic", "preciomp", "codartic", Rs.Fields!codArtic, "T", cad)
'            Else 'precio ultima compra
'                vPrecioVenta = DevuelveDesdeBDNew(cAgro, "sartic", "preciouc", "codartic", Rs.Fields!codArtic, "T", cad)
'            End If
'
'            If vPrecioVenta <> "" Then
'                vImporte = Rs.Fields!Cantidad * CSng(vPrecioVenta)
'            Else
'                vImporte = 0
'            End If
            
            vPrecioVenta = DevuelveDesdeBDNew(cAgro, "sartic", "preciomp", "codartic", Rs.Fields!CodArtic, "T", Cad)
            
            If Val(Cad) = 1 Then
                Sql = "INSERT INTO smoval (codartic, codalmac, fechamov, horamovi, tipomovi, detamovi, cantidad, impormov, codigope, letraser, document, numlinea) "
                Sql = Sql & " VALUES (" & DBSet(Rs.Fields!CodArtic, "T") & ", " & Rs.Fields!codAlmac & ", '" & Format(Rs.Fields!fecmovim, "yyyy-mm-dd") & "', '"
                Sql = Sql & Format(Rs.Fields!fecmovim & " " & Time, "yyyy-mm-dd hh:mm:ss") & "', " & Rs.Fields!tipomovi & ", '" & vTipoMov.TipoMovimiento & "', " & DBSet(Rs.Fields!Cantidad, "N") & ", " & DBSet(Rs.Fields!ImporteL, "N") & ", "
                If Data1.Recordset!clisoc = 0 Or Data1.Recordset!clisoc = 2 Then
                    Sql = Sql & DBSet(Rs.Fields!CodSocio, "N") & ",'"
                Else
                    Sql = Sql & DBSet(Rs.Fields!CodClien, "N") & ",'"
                End If
                Sql = Sql & vTipoMov.LetraSerie & "', '" & Format(DBSet(Rs.Fields!codservi, "N"), "0000000") & "', " & Rs.Fields!NumLinea & ")"
                conn.Execute Sql
            End If
            Rs.MoveNext
        Wend
    Else
        bol = False
    End If
    Set vTipoMov = Nothing
    Rs.Close
    Set Rs = Nothing
    
EInsertar:
    If Err.Number <> 0 Then
        Set vTipoMov = Nothing
        Rs.Close
        Set Rs = Nothing
    End If
    If Err.Number <> 0 Or Not bol Then
         'Hay error , almacenamos y salimos
        InsertarMovimArticulos = False
    Else
        InsertarMovimArticulos = True
    End If
End Function



Private Sub IncrementarProgres(Veces As Integer)
On Error Resume Next
    Me.ProgressBar1.Value = Me.ProgressBar1.Value + (Veces * 10)
    If Err.Number <> 0 Then Err.Clear
    Me.Refresh
End Sub


Private Function BorrarTraspaso(EnHistorico As Boolean) As Boolean
'Si EnHistorico=true borra de las tablas de historico: "schtra" y "slhtra"
'Si EnHistorico=false borra de las tablas de traspaso: "scatra" y "slitra"
Dim Sql As String

    BorrarTraspaso = False
    
    'Borramos las lineas
    Sql = "Delete from "
    If EnHistorico Then
        Sql = Sql & "slhser"
        Sql = Sql & " WHERE codservi = " & Data1.Recordset!codservi
        Sql = Sql & " AND fecmovim = '" & Data1.Recordset!fecmovim & "'"
        Sql = Sql & " and clisoc = " & Data1.Recordset!clisoc
    Else
        Sql = Sql & "sliser"
        Sql = Sql & " WHERE codservi = " & Data1.Recordset!codservi
        Sql = Sql & " and clisoc = " & Data1.Recordset!clisoc
    End If
    conn.Execute Sql
    
    'La cabecera
    Sql = "Delete from "
    If EnHistorico Then
        Sql = Sql & "schser"
        Sql = Sql & " WHERE codservi =" & Data1.Recordset!codservi
        Sql = Sql & " AND fecmovim='" & Data1.Recordset!fecmovim & "'"
        Sql = Sql & " and clisoc = " & Data1.Recordset!clisoc
    Else
        Sql = Sql & "scaser"
        Sql = Sql & " WHERE codservi =" & Data1.Recordset!codservi
        Sql = Sql & " and clisoc = " & Data1.Recordset!clisoc
    End If
    conn.Execute Sql
    
    If Err.Number <> 0 Then
        BorrarTraspaso = False
    Else
        BorrarTraspaso = True
    End If
End Function


Private Sub CargarComboAux()
'### Combo Tipo Movimiento
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Entrada, 1-Salida

    cboAux.Clear
    cboAux.AddItem "S"
    cboAux.ItemData(cboAux.NewIndex) = 0
    
    cboAux.AddItem "E"
    cboAux.ItemData(cboAux.NewIndex) = 1
        
End Sub


Public Sub ActualizarSituacionImpresion()
Dim Cad As String, Indicador As String
On Error GoTo EImpresion
   
    Cad = "(" & ObtenerWhereCP(False) & ")"
    If SituarDataMULTI(Data1, Cad, Indicador) Then
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


Private Sub BotonImprimir()
        If Text1(0).Text = "" Then Exit Sub
        frmListado2.NumCod = Text1(0).Text
        frmListado2.TipoMto = Combo1(0).ListIndex
        If Not EsHistorico Then
            AbrirListado2 (10) '10: Informe Movimientos de Servicios
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

    indRPT = 80 ' Historico Movimientos de Servicios
    If PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then
        With frmImprimir
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .NombreRPT = nomDocu
            .EnvioEMail = False
            .Opcion = 8
            .Titulo = "Hist. Movimientos Servicios"
            .ConSubInforme = True
            If cadSeleccion <> "" Then
                .FormulaSeleccion = cadSeleccion
            Else
                'Se Llama desde dobleclick en frmAlmMovimArticulos
                Cad = "{schser.codservi}= " & Data1.Recordset!codservi
                Cad = Cad & " and {schser.fecmovim}= Date(" & Year(Data1.Recordset!fecmovim) & "," & Month(Data1.Recordset!fecmovim) & "," & Day(Data1.Recordset!fecmovim) & ")" & ""
                Cad = Cad & " and {schser.clisoc} = " & Combo1(0).ListIndex
                .FormulaSeleccion = Cad
            End If
            .Show vbModal
        End With
    End If
End Sub



Private Function InsertarMovimiento(vSQL As String, vTipoMov As CTiposMov) As Boolean
Dim MenError As String
Dim bol As Boolean
On Error GoTo EInsertarMovim
    
    bol = True
    
    'Aqui empieza transaccion
    conn.BeginTrans
    
    MenError = "Error al insertar en la tabla de Movimientos(smoval)."
    conn.Execute vSQL, , adCmdText
    
    MenError = "Error al actualizar el contador del recibo."
    bol = vTipoMov.IncrementarContador(CodTipoMov)

EInsertarMovim:
        If Err.Number <> 0 Then
            MenError = "Insertando Movimiento." & vbCrLf & "----------------------------" & vbCrLf & MenError
            MuestraError Err.Number, MenError, Err.Description
            bol = False
        End If
        If bol Then
            conn.CommitTrans
            InsertarMovimiento = True
        Else
            conn.RollbackTrans
            InsertarMovimiento = False
        End If
End Function


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Function ObtenerWhereCP(conWhere As Boolean) As String
'Obtiene la sentencia WHERE para seleccionar registros de la tabla por Clave Primaria
On Error Resume Next
    If conWhere Then
        ObtenerWhereCP = " WHERE codservi= " & Val(Text1(0).Text) & " and clisoc = " & DBSet(Combo1(0).ListIndex, "N")
    Else
        ObtenerWhereCP = " codservi= " & Val(Text1(0).Text) & " and clisoc = " & DBSet(Combo1(0).ListIndex, "N")
    End If
    If Err.Number <> 0 Then Err.Clear
End Function


Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim Sql As String

    Set vTipoMov = New CTiposMov
    
    If vTipoMov.Leer(CodTipoMov) Then
        Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
        Text1(0).Text = Format(Text1(0).Text, "0000000")
        cmdCancelar.Caption = "Cancelar"
        Sql = CadenaInsertarDesdeForm(Me)
        
        If Sql <> "" Then
            If InsertarMovimiento(Sql, vTipoMov) Then
                CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                PonerCadenaBusqueda
                PonerModo 2
                 'Ponerse en Modo Insertar Lineas
                '--monica : rollo toolbar
                'BotonLineas
                BotonAnyadirLineas
            End If
        End If
    End If
    Set vTipoMov = Nothing
End Sub


Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim i As Byte
    
    Combo1(0).Clear
    
    Combo1(0).AddItem "Socios"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    
    Combo1(0).AddItem "Clientes"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    
    Combo1(0).AddItem "Regulariz.Socios"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    
    Combo1(0).AddItem "Regulariz.Clientes"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 3
    
End Sub

Private Sub PonerSocioVisible()
    imgBuscar(1).visible = (Combo1(0).ListIndex = 0 Or Combo1(0).ListIndex = 2)
    imgBuscar(1).Enabled = (Combo1(0).ListIndex = 0 Or Combo1(0).ListIndex = 2)
    Text1(2).Enabled = (Combo1(0).ListIndex = 0 Or Combo1(0).ListIndex = 2)
    
    If Combo1(0).ListIndex = 0 Or Combo1(0).ListIndex = 2 Then ' socios --> limpio clientes
        Text1(3).Text = ""
        Text2(3).Text = ""
    End If
    
    imgBuscar(2).visible = (Combo1(0).ListIndex = 1 Or Combo1(0).ListIndex = 3)
    imgBuscar(2).Enabled = (Combo1(0).ListIndex = 1 Or Combo1(0).ListIndex = 3)
    Text1(3).Enabled = (Combo1(0).ListIndex = 1 Or Combo1(0).ListIndex = 3)
    
    If Combo1(0).ListIndex = 1 Or Combo1(0).ListIndex = 3 Then ' cliente --> limpio socios
        Text1(2).Text = ""
        Text2(2).Text = ""
    End If
    
    Select Case Combo1(0).ListIndex
        Case 0
            CodTipoMov = "SES"
        Case 1
            CodTipoMov = "SEC"
        Case 2
            CodTipoMov = "RES"
        Case 3
            CodTipoMov = "REC"
    End Select
    
End Sub


Private Function ModificarLineaHco() As Boolean
'Modifica un registro en la tabla de lineas de Albaran: slialb
Dim Sql As String, vWhere As String
Dim vCStock As CStock
Dim vArtic As CArticulo
Dim b As Boolean
Dim MenError As String
Dim dentroTRANSAC As Boolean
Dim cadNumLote As String
Dim vEntSal As String

'--monica
'Dim cLote As CNumLote

    On Error GoTo eModificarLinea

    ModificarLineaHco = False
    Sql = ""
    dentroTRANSAC = False
    b = True

    Set vArtic = New CArticulo
    If Not vArtic.LeerDatos(txtAux(0).Text) Then Exit Function
    Set vCStock = New CStock

'   If DatosOkLinea() Then
    'sql para actualizar la linea de los servicios

    Sql = "UPDATE slhser Set cantidad = " & DBSet(txtAux(2).Text, "N")
    Sql = Sql & ", tipomovi = " & cboAux.ItemData(cboAux.ListIndex)
    Sql = Sql & ", motimovi = " & DBSet(txtAux(3).Text, "T")
    Sql = Sql & ", precioar = " & DBSet(txtAux(4).Text, "N")
    Sql = Sql & ", importel = " & DBSet(txtAux(5).Text, "N")
    Sql = Sql & " WHERE codservi =" & Val(Text1(0).Text) & " AND "
    Sql = Sql & " clisoc = " & Combo1(0).ListIndex & " AND "
    Sql = Sql & " numlinea =" & Val(cmdAceptar.Tag)

    If Sql <> "" Then
        dentroTRANSAC = True
        conn.BeginTrans

        MenError = "Actualizando Lineas Servicios"
        conn.Execute Sql

        'Actualizar Stocks de los articulos y movimientos
        '===================================================
        If b Then
            If cboAux.ItemData(cboAux.ListIndex) = 0 Then
                vEntSal = "S"
            Else
                vEntSal = "E"
            End If
            
            MenError = "Actualizando stocks y movimientos almacen"
            'si no se ha modificado el almacen reestablecemos cantidad y precio
            'deshacer el movimiento para el almacen anterior y devolver stock
            
            b = InicializarCStock(vCStock, vEntSal)
            If Data2.Recordset!clisoc = 0 Or Data2.Recordset!clisoc = 1 Then
                If b Then b = vCStock.ModificarStockServicios(DBLet(Data2.Recordset!Cantidad, "N"))
            Else
                ' solo modificamos el movimiento de la smoval pq son regularizaciones que no mueven stock
                If b Then b = vCStock.ModificarMovimArticulosServicios
            End If
        End If

        If b Then
            conn.CommitTrans
        Else
            conn.RollbackTrans
        End If
        ModificarLineaHco = b
    End If

    Set vCStock = Nothing
    Set vArtic = Nothing
    Exit Function

eModificarLinea:
    If dentroTRANSAC Then conn.RollbackTrans
    If Not vArtic Is Nothing Then Set vArtic = Nothing
    If Not vCStock Is Nothing Then Set vCStock = Nothing
    ModificarLineaHco = False
    MuestraError Err.Number, "Modificar Lineas de Servicios" & vbCrLf & MenError & vbCrLf & Err.Description
End Function


Private Function InicializarCStock(ByRef vCStock As CStock, TipoM As String, Optional NumLinea As String) As Boolean
'On Error Resume Next
On Error Resume Next

    Select Case DBLet(Data2.Recordset!clisoc, "N")
        Case 0
            CodTipoMov = "SES"
            vCStock.Trabajador = CLng(Text1(2).Text) 'En smoval guardamos el socio
        Case 1
            CodTipoMov = "SEC"
            vCStock.Trabajador = CLng(Text1(3).Text) 'En smoval guardamos el cliente
        Case 2
            CodTipoMov = "RES"
            vCStock.Trabajador = CLng(Text1(2).Text) 'En smoval guardamos el socio
        Case 3
            CodTipoMov = "REC"
            vCStock.Trabajador = CLng(Text1(3).Text) 'En smoval guardamos el cliente
    End Select

    vCStock.tipoMov = TipoM 'Movimiento de Entrada o Salida
    vCStock.DetaMov = CodTipoMov '"ALC=Albaran de Compra"
    vCStock.Fechamov = Text1(1).Text
    vCStock.Documento = Text1(0).Text
    
    vCStock.CodArtic = Data2.Recordset!CodArtic
    vCStock.codAlmac = CInt(Data1.Recordset!codAlmac)
    vCStock.Cantidad = CSng(ImporteSinFormato(txtAux(2).Text))
    vCStock.Importe = CCur(ImporteSinFormato(txtAux(5).Text))
    vCStock.LineaDocu = CInt(Data2.Recordset!NumLinea)
    
    If Err.Number <> 0 Then
        MsgBox "No se han podido inicializar la clase para actualizar Stock", vbExclamation
        InicializarCStock = False
    Else
        InicializarCStock = True
    End If
End Function

