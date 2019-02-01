VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCCCostesDiarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entrada Costes Diarios"
   ClientHeight    =   10050
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   15915
   Icon            =   "frmCCCostesDiarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10050
   ScaleWidth      =   15915
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
      Index           =   0
      Left            =   13185
      TabIndex        =   50
      Top             =   270
      Width           =   1605
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   90
      TabIndex        =   48
      Top             =   135
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   49
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
   Begin VB.Frame FrameAux1 
      Caption         =   "Categorias"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   4770
      Left            =   9270
      TabIndex        =   20
      Top             =   4440
      Width           =   6530
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
         Index           =   41
         Left            =   780
         MaxLength       =   4
         TabIndex        =   44
         Tag             =   "Código Coste|N|S|||cccabdia|codcoste||S|"
         Text            =   "cost"
         Top             =   2610
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Index           =   3
         Left            =   4275
         TabIndex        =   40
         Text            =   "Text3"
         Top             =   4245
         Width           =   1200
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Height          =   285
         Index           =   1
         Left            =   1710
         TabIndex        =   36
         Text            =   "Text3"
         Top             =   4245
         Width           =   1300
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
         Index           =   43
         Left            =   1770
         MaxLength       =   2
         TabIndex        =   22
         Tag             =   "Categoria|N|N|||cclindia2|codcateg|00||"
         Text            =   "cate"
         Top             =   2610
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtAux2 
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
         Index           =   43
         Left            =   2550
         TabIndex        =   33
         Text            =   "nombre"
         Top             =   2610
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.CommandButton btnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "+"
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
         Left            =   2340
         MaskColor       =   &H00000000&
         TabIndex        =   32
         ToolTipText     =   "Buscar categoria"
         Top             =   2580
         Visible         =   0   'False
         Width           =   195
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
         Index           =   42
         Left            =   1260
         MaxLength       =   16
         TabIndex        =   31
         Tag             =   "Linea|N|N|||cclindia2|numlinea|00|S|"
         Text            =   "lin"
         Top             =   2610
         Visible         =   0   'False
         Width           =   240
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
         Index           =   44
         Left            =   4320
         MaxLength       =   7
         TabIndex        =   23
         Tag             =   "Horas|N|N|||cclindia2|horas|###0.00||"
         Text            =   "horas"
         Top             =   2610
         Visible         =   0   'False
         Width           =   795
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
         Index           =   40
         Left            =   180
         MaxLength       =   16
         TabIndex        =   21
         Tag             =   "Fecha|F|N|||cclindia2|fecha|dd/mm/yyyy|S|"
         Text            =   "orden"
         Top             =   2610
         Visible         =   0   'False
         Width           =   555
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   1
         Left            =   90
         TabIndex        =   24
         Top             =   285
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
      Begin MSDataGridLib.DataGrid DataGridAux 
         Bindings        =   "frmCCCostesDiarios.frx":000C
         Height          =   3375
         Index           =   1
         Left            =   90
         TabIndex        =   25
         Top             =   690
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   5953
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
      Begin MSAdodcLib.Adodc AdoAux 
         Height          =   375
         Index           =   1
         Left            =   1350
         Top             =   660
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
         Caption         =   "AdoAux(1)"
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
      Begin VB.Label Label11 
         Caption         =   "Total Horas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   3
         Left            =   3120
         TabIndex        =   41
         Top             =   4245
         Width           =   1440
      End
      Begin VB.Label Label11 
         Caption         =   "TOTAL COSTES: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   37
         Top             =   4245
         Width           =   1740
      End
   End
   Begin VB.Frame FrameAux0 
      Caption         =   "Trabajadores"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   4770
      Left            =   90
      TabIndex        =   11
      Top             =   4440
      Width           =   9135
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Height          =   285
         Index           =   0
         Left            =   2205
         TabIndex        =   42
         Text            =   "Text3"
         Top             =   4290
         Width           =   1300
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Index           =   2
         Left            =   5700
         TabIndex        =   38
         Text            =   "Text3"
         Top             =   4290
         Width           =   1200
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
         Index           =   8
         Left            =   6240
         MaxLength       =   20
         TabIndex        =   35
         Tag             =   "Fecha Fin|FH|N|||cclindia1|fechafin|yyyy-mm-dd hh:mm:ss||"
         Text            =   "f.fin"
         Top             =   2550
         Visible         =   0   'False
         Width           =   720
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
         Index           =   7
         Left            =   5400
         MaxLength       =   20
         TabIndex        =   34
         Tag             =   "Fecha Ini|FH|N|||cclindia1|fechaini|yyyy-mm-dd hh:mm:ss||"
         Text            =   "f.ini"
         Top             =   2550
         Visible         =   0   'False
         Width           =   720
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
         Index           =   6
         Left            =   4920
         MaxLength       =   7
         TabIndex        =   17
         Tag             =   "Horas|N|N|||cclindia1|horas|###0.00||"
         Text            =   "horas"
         Top             =   2550
         Visible         =   0   'False
         Width           =   390
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
         Left            =   4020
         MaxLength       =   8
         TabIndex        =   16
         Text            =   "h.fin"
         Top             =   2550
         Visible         =   0   'False
         Width           =   750
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
         Left            =   3090
         MaxLength       =   8
         TabIndex        =   15
         Text            =   "h.inic"
         Top             =   2550
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.CommandButton btnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "+"
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
         Left            =   2040
         MaskColor       =   &H00000000&
         TabIndex        =   30
         ToolTipText     =   "Buscar trabajador"
         Top             =   2520
         Visible         =   0   'False
         Width           =   195
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
         Left            =   930
         MaxLength       =   16
         TabIndex        =   29
         Tag             =   "Linea|N|N|||cclindia1|numlinea|0000|S|"
         Text            =   "lin"
         Top             =   2550
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.TextBox txtAux2 
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
         Index           =   3
         Left            =   2265
         TabIndex        =   28
         Top             =   2550
         Visible         =   0   'False
         Width           =   705
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
         Index           =   3
         Left            =   1260
         MaxLength       =   6
         TabIndex        =   14
         Tag             =   "Trabajador|N|N|||cclindia1|codtraba|000000||"
         Text            =   "trabajado"
         Top             =   2550
         Visible         =   0   'False
         Width           =   750
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
         Index           =   1
         Left            =   450
         MaxLength       =   4
         TabIndex        =   13
         Tag             =   "Código Coste|N|S|||cccabdia|codcoste||S|"
         Text            =   "cost"
         Top             =   2565
         Visible         =   0   'False
         Width           =   435
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
         Index           =   0
         Left            =   45
         MaxLength       =   16
         TabIndex        =   12
         Tag             =   "Fecha|F|N|||cclindia1|fecha|dd/mm/yyyy|S|"
         Text            =   "Fecha"
         Top             =   2565
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   0
         Left            =   135
         TabIndex        =   18
         Top             =   285
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
         Bindings        =   "frmCCCostesDiarios.frx":0024
         Height          =   3375
         Index           =   0
         Left            =   135
         TabIndex        =   19
         Top             =   690
         Width           =   8760
         _ExtentX        =   15452
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   19
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
            AllowFocus      =   0   'False
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label11 
         Caption         =   "TOTAL COSTES: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   405
         TabIndex        =   43
         Top             =   4290
         Width           =   1995
      End
      Begin VB.Label Label11 
         Caption         =   "Total Horas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   2
         Left            =   4245
         TabIndex        =   39
         Top             =   4290
         Width           =   1440
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3420
      Index           =   0
      Left            =   75
      TabIndex        =   7
      Top             =   930
      Width           =   15695
      Begin VB.CommandButton btnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "+"
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
         Left            =   2340
         MaskColor       =   &H00000000&
         TabIndex        =   47
         ToolTipText     =   "Buscar concepto"
         Top             =   2580
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton btnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "+"
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
         Left            =   1260
         MaskColor       =   &H00000000&
         TabIndex        =   46
         ToolTipText     =   "Buscar fecha"
         Top             =   2580
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
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
         Left            =   180
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Fecha|F|N|||cccabdia|fecha|dd/mm/yyyy|S|"
         Text            =   "1234567890"
         Top             =   2610
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox text1 
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
         Left            =   4860
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   2
         Tag             =   "Observaciones|T|S|||cabdia|observac|||"
         Top             =   2580
         Visible         =   0   'False
         Width           =   8490
      End
      Begin VB.TextBox text2 
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
         Left            =   2610
         TabIndex        =   26
         Text            =   "12345678901234567890"
         Top             =   2580
         Visible         =   0   'False
         Width           =   2190
      End
      Begin VB.TextBox text1 
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
         Index           =   1
         Left            =   1590
         MaxLength       =   7
         TabIndex        =   1
         Tag             =   "Código Coste|N|S|||cccabdia|codcoste|000000|S|"
         Text            =   "1234567"
         Top             =   2610
         Visible         =   0   'False
         Width           =   765
      End
      Begin MSDataGridLib.DataGrid DataGridAux 
         Bindings        =   "frmCCCostesDiarios.frx":003C
         Height          =   2910
         Index           =   2
         Left            =   180
         TabIndex        =   45
         Top             =   300
         Width           =   13170
         _ExtentX        =   23230
         _ExtentY        =   5133
         _Version        =   393216
         AllowUpdate     =   0   'False
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   19
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
            AllowFocus      =   0   'False
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1260
         Picture         =   "frmCCCostesDiarios.frx":0054
         ToolTipText     =   "Buscar fecha"
         Top             =   2580
         Width           =   240
      End
      Begin VB.Label Label10 
         Caption         =   "Concepto"
         Height          =   255
         Left            =   2100
         TabIndex        =   27
         Top             =   2790
         Width           =   690
      End
      Begin VB.Label Label29 
         Caption         =   "Observaciones"
         Height          =   255
         Left            =   195
         TabIndex        =   10
         Top             =   2820
         Width           =   1125
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   1365
         ToolTipText     =   "Zoom descripción"
         Top             =   2820
         Width           =   240
      End
      Begin VB.Label Label18 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   3000
         TabIndex        =   9
         Top             =   2820
         Width           =   930
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   90
      TabIndex        =   5
      Top             =   9270
      Width           =   2865
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
         Left            =   120
         TabIndex        =   6
         Top             =   180
         Width           =   2655
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
      Left            =   14715
      TabIndex        =   4
      Top             =   9315
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
      Left            =   13545
      TabIndex        =   3
      Top             =   9315
      Width           =   1065
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
      Left            =   14715
      TabIndex        =   8
      Top             =   9315
      Visible         =   0   'False
      Width           =   1065
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   15300
      TabIndex        =   51
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
Attribute VB_Name = "frmCCCostesDiarios"
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
Dim indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim CadB As String

Dim CodTipoMov As String



Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 0 ' Fecha de coste
            btnFec (0)
            
        Case 1 ' Concepto
            indice = 1
            
            Set frmCon = New frmCCManConcep
            frmCon.DatosADevolverBusqueda = "0|1|"
            frmCon.CodigoActual = Text1(0).Text
            frmCon.Show vbModal
            Set frmCon = Nothing
            PonerFoco Text1(1)
            
        Case 2 ' Trabajadores
            indice = 3
            Set frmTra = New frmBasico
            AyudaTrabajadores frmTra, txtAux(3)
            Set frmTra = Nothing
            PonerFoco txtAux(3)
            
        Case 3 ' Categoria
            indice = 43
            Set frmCat = New frmBasico
            AyudaCategorias frmCat, txtAux(43)
            Set frmCat = Nothing
            PonerFoco txtAux(43)
            
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
            indice = 0
        Case 3
            indice = 4
        Case 4
            indice = 6
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






Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BÚSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm2(Me, 1) Then
                    TerminaBloquear
                    CargaGrid
                    PosicionarData
                    DataGridAux(2).AllowAddNew = False
                    BotonAnyadirLinea 0
                End If
            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario2(Me, 1) Then
                    TerminaBloquear
                    PosicionarData
                End If
            Else
                ModoLineas = 0
            End If
            
        ' *** si n'hi han llínies ***
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 1 'afegir llínia
                    InsertarLinea
                Case 2 'modificar llínies
                    If ModificarLinea Then
                        PosicionarData
                    Else
                        PonerFoco txtAux(12)
                    End If
            End Select
            'nuevo calculamos los totales de lineas
            CalcularTotales
        ' **************************
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
            PonerModo 2
        Else
            If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                BotonAnyadir
            Else
                PonerModo 1 'búsqueda
                ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
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
    CheckValueGuardar Me.Name, Me.chkVistaPrevia(0).Value
    If Modo = 4 Then TerminaBloquear
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim i As Integer

    PrimeraVez = True
    
    ' ICONETS DE LA BARRA
    btnPrimero = 17 'index del botó "primero"
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
'        .Buttons(11).Image = 19   'Expandir Añadir, Borrar y Modificar
'        .Buttons(12).Image = 26   'cambio de costes de confeccion
        
        'el 10 i el 11 son separadors
        .Buttons(8).Image = 10  'Imprimir
    End With
    
    
    
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
   
    'carga IMAGES de mail
'    For i = 0 To Me.imgMail.Count - 1
'        Me.imgMail(i).Picture = frmPpal.imgListImages16.ListImages(2).Picture
'    Next i
    
    'IMAGES para zoom
    For i = 0 To Me.imgZoom.Count - 1
        Me.imgZoom(i).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    Next i
    
    LimpiarCampos   'Neteja els camps TextBox
    ' ******* si n'hi han llínies *******
    DataGridAux(2).ClearFields
    DataGridAux(0).ClearFields
    DataGridAux(1).ClearFields
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "cccabdia"
    Ordenacion = " ORDER BY fecha, codcoste"
    
    'Mirem com està guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    CadenaConsulta = "Select fecha, cccabdia.codcoste, ccconcostes.nomcoste, cccabdia.observac from " & NombreTabla
    CadenaConsulta = CadenaConsulta & ", ccconcostes where cccabdia.codcoste = ccconcostes.codcoste "
    
    AdoAux(2).ConnectionString = conn
    AdoAux(2).RecordSource = CadenaConsulta
    AdoAux(2).Refresh
       
    CargaGrid ""
    CargaGrid2 0, False
    CargaGrid2 1, False

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

Private Sub LimpiarCamposLin(FrameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, FrameAux  'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO s'habiliten, o no, els diversos camps del
'   formulari en funció del modo en que anem a treballar
Private Sub PonerModo(Kmodo As Byte, Optional indFrame As Integer)
Dim i As Integer, NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo
 
    'Actualisa Iconos Insertar,Modificar,Eliminar
    'ActualizarToolbar Modo, Kmodo
    Modo = Kmodo
'    PonerIndicador lblIndicador, Modo, ModoLineas
    
    b = (Modo = 2)
    If b Then
        If Not AdoAux(2).Recordset.EOF Then
            Text1(0).Text = Format(AdoAux(2).Recordset!fecha, "dd/mm/yyyy")
            Text1(1).Text = AdoAux(2).Recordset!codCoste
            
            PonerContRegIndicador Me.lblIndicador, AdoAux(2), ""
            
            LLamaLineas 2, Modo
            CargaGrid2 0, True
            CargaGrid2 1, True
        End If
    Else
        PonerIndicador lblIndicador, Modo, ModoLineas
    End If
           
               
    'Modo 2. N'hi han datos i estam visualisant-los
    '=========================================
    'Posem visible, si es formulari de búsqueda, el botó "Regresar" quan n'hi han datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = (Modo = 2)
    Else
        cmdRegresar.visible = False
    End If
    
'    text1(5).Enabled = True
    '=======================================
'    b = (Modo = 2)
'    'Posar Fleches de desplasament visibles
'    NumReg = 1
'    If Not AdoAux(2).Recordset.EOF Then
'        If AdoAux(2).Recordset.RecordCount > 1 Then NumReg = 2 'Només es per a saber que n'hi ha + d'1 registre
'    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
'    '---------------------------------------------
'
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
       
    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
    'Si estem en Insertar a més neteja els camps Text1
'    BloquearText1 Me, Modo
    
    b = (Modo = 3 Or Modo = 4 Or Modo = 1) '06/09/2005, lleve el modo 5 per a que no es puga modificar la capçalera mentre treballe en les llínies
    If b Then
        For i = 0 To 2
            Text1(i).Locked = Not b  '((Not b) And (Modo <> 1))
            If (Modo = 3 Or Modo = 4) Then
                 Text1(i).BackColor = vbWhite
            Else
'                 text1(I).BackColor = &H80000018 'groc
            End If
        Next i
    End If
    
    b = (Modo <> 1)
    'Campos Nº Pedido bloqueado y en azul
'    BloquearTxt text1(0), b, True
    
    '*** si n'hi han combos a la capçalera ***
    '**************************
    
    ' ***** bloquejar tots els controls visibles de la clau primaria de la capçalera ***
'    If Modo = 4 Then _
'        BloquearTxt Text1(0), True 'si estic en  modificar, bloqueja la clau primaria
    ' **********************************************************************************
    
    ' **** si n'hi han imagens de buscar en la capçalera *****
'    BloquearImgBuscar Me, Modo, ModoLineas
    BloquearImgZoom Me, Modo, ModoLineas
    ' ********************************************************
    BloquearImgFec Me, 0, Modo
    BloquearImgFec Me, 1, Modo
        
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    PonerLongCampos

    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid2 0, False
        CargaGrid2 1, False
    End If
    
    b = (Modo = 4) Or (Modo = 2)
    DataGridAux(0).Enabled = b
    DataGridAux(1).Enabled = b
      
    ' ****** si n'hi han combos a la capçalera ***********************
    ' ****************************************************************
    
    PonerModoOpcionesMenu (Modo) 'Activar opcions menú según modo
    PonerOpcionesMenu   'Activar opcions de menú según nivell
                        'de permisos de l'usuari

EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los TEXT1
    PonerLongCamposGnral Me, Modo, 1
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub

Private Sub PonerModoOpcionesMenu(Modo)
'Actives unes Opcions de Menú i Toolbar según el modo en que estem
Dim b As Boolean, bAux As Boolean
Dim i As Byte
    
    'Barra de CAPÇALERA
    '------------------------------------------
    'b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    b = (Modo = 2 Or Modo = 0)
    'Buscar
    Toolbar1.Buttons(5).Enabled = b
    Me.mnBuscar.Enabled = b
    'Vore Tots
    Toolbar1.Buttons(6).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    'Insertar
    Toolbar1.Buttons(1).Enabled = b And Not DeConsulta
    Me.mnNuevo.Enabled = b And Not DeConsulta
    
    b = (Modo = 2 And AdoAux(2).Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnEliminar.Enabled = b
    
    'Imprimir
    Toolbar1.Buttons(8).Enabled = True And Not DeConsulta
       
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
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
Dim tabla As String
    
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
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadB = Aux
        Aux = "cccabdia." & ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
        CadB = CadB & " and " & Aux
        '   Com la clau principal es única, en posar el sql apuntant
        '   al valor retornat sobre la clau ppal es suficient
'        CadenaConsulta = "select * from " & NombreTabla & "  WHERE " & CadB & " " & Ordenacion
        
        CadenaConsulta = "Select fecha, cccabdia.codcoste, ccconcostes.nomcoste, cccabdia.observac from " & NombreTabla
        CadenaConsulta = CadenaConsulta & ", ccconcostes where cccabdia.codcoste = ccconcostes.codcoste and " & CadB
        
        ' **********************************
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    Text1(indice).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmC1_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
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
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000") 'concepto de coste
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub


Private Sub frmZ_Actualizar(vCampo As String)
     Text1(indice).Text = vCampo
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
        indice = 10
        frmZ.pTitulo = "Observaciones de la Orden de Confección"
        frmZ.pValor = Text1(indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco Text1(indice)
    End If
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
'    Screen.MousePointer = vbHourglass
'    frmListConfeccion.Show vbModal
'    Screen.MousePointer = vbDefault
End Sub

Private Sub mnModificar_Click()
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(adoaux(2).Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************
    
    If BLOQUEADesdeFormulario2(Me, AdoAux(2), 1) Then BotonModificarLinea 2
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadirLinea 2
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 1  'Nou
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
        Case 5  'Búscar
           mnBuscar_Click
        Case 6  'Tots
            mnVerTodos_Click
        Case 8 'Imprimir
            mnImprimir_Click
    End Select
End Sub

Private Sub BotonBuscar()
Dim i As Integer
Dim anc As Single

' ***** Si la clau primaria de la capçalera no es Text1(0), canviar-ho en <=== *****
    If Modo <> 1 Then
        LimpiarCampos
        
        anc = DataGridAux(2).Top
        If DataGridAux(2).Row < 0 Then
            anc = anc + 240
        Else
            anc = anc + DataGridAux(2).RowTop(DataGridAux(2).Row) + 5
        End If
        
        LLamaLineas 2, 1, anc
        
        PonerModo 1
        PonerFoco Text1(0) ' <===
        Text1(0).BackColor = vbYellow ' <===
        ' *** si n'hi han combos a la capçalera ***
    Else
        HacerBusqueda
        If AdoAux(2).Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
' ******************************************************************************
End Sub

Private Sub HacerBusqueda()
    
    
    CadB = ObtenerBusqueda2(Me, 1)
    
    If chkVistaPrevia(0) = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        ' *** foco al 1r camp visible de la capçalera que siga clau primaria ***
        PonerFoco Text1(0)
        ' **********************************************************************
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
    Dim Cad As String
        
    'Cridem al form
    ' **************** arreglar-ho per a vore lo que es desije ****************
    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
    Cad = ""
    Cad = Cad & "Fecha|fecha|T|dd/mm/yyyy|15·"
    Cad = Cad & "Código|cccabdia.codcoste|N|000000|10·"
    Cad = Cad & "Denominacion|ccconcostes.nomcoste|T||36·"
    Cad = Cad & "Observaciones|cccabdia.observac|T||39·"
    
    If Cad <> "" Then
        
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        Cad = NombreTabla & " left join ccconcostes on cccabdia.codcoste = ccconcostes.codcoste "
        frmB.vtabla = Cad 'NombreTabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        frmB.vDevuelve = "0|1|" '*** els camps que volen que torne ***
        frmB.vTitulo = "Costes Diarios" ' ***** repasa açò: títol de BuscaGrid *****
        frmB.vSelElem = 0

        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha posat valors i tenim que es formulari de búsqueda llavors
        'tindrem que tancar el form llançant l'event
        If HaDevueltoDatos Then
            If (Not AdoAux(2).Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                cmdRegresar_Click
        Else   'de ha retornat datos, es a decir NO ha retornat datos
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String
Dim Aux As String
Dim i As Integer
Dim J As Integer

    If AdoAux(2).Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    Cad = ""
    i = 0
    Do
        J = i + 1
        i = InStr(J, DatosADevolverBusqueda, "|")
        If i > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, i - J)
            J = Val(Aux)
            Cad = Cad & Text1(J).Text & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(Cad)
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
    LimpiarCampos 'Neteja els Text1
    CadB = ""
    
    If chkVistaPrevia(0).Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select fecha, cccabdia.codcoste, ccconcostes.nomcoste, cccabdia.observac from " & NombreTabla
        CadenaConsulta = CadenaConsulta & ", ccconcostes where cccabdia.codcoste = ccconcostes.codcoste "
        PonerCadenaBusqueda
    End If
End Sub

Private Sub BotonAnyadir()
Dim NumF As String
    
    LimpiarCampos 'Huida els TextBox
    PonerModo 3
    
    
    
    LLamaLineas 2, Modo
    
    
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la capçalera *******
'    text1(0).Text = SugerirCodigoSiguienteStr("forfaits", "codforfait")
'    FormateaCampo text1(0)
    '******************** canviar taula i camp **************************
'    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
'        NumF = NuevoCodigo
'    Else
'        NumF = ""
'    End If
    '********************************************************************

'    text1(0) = NumF
    PonerFoco Text1(1) '*** 1r camp visible que siga PK ***
    
    ' *** si n'hi han camps de descripció a la capçalera ***
    'PosarDescripcions

End Sub

Private Sub BotonModificar()

    PonerModo 4

    ' *** bloquejar els camps visibles de la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    BloquearTxt Text1(1), True
    
    LLamaLineas 2, 4
    
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco Text1(2)
End Sub

Private Sub BotonEliminar()
Dim Cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If AdoAux(2).Recordset.EOF Then Exit Sub

    Text1(0).Text = Format(AdoAux(2).Recordset!fecha, "dd/mm/yyyy")
    Text1(1).Text = AdoAux(2).Recordset!codCoste


    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(adoaux(2).Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
    ' ***************************************************************************

    ' *************** canviar la pregunta ****************
    Cad = "¿Seguro que desea eliminar el Coste Diario?"
    Cad = Cad & vbCrLf & "Fecha: " & AdoAux(2).Recordset.Fields(0)
    Cad = Cad & vbCrLf & "Coste: " & AdoAux(2).Recordset.Fields(1) & " " & AdoAux(2).Recordset.Fields(2)
    
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = AdoAux(2).Recordset.AbsolutePosition
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf SituarDataTrasEliminar(AdoAux(2), NumRegElim, True) Then
            PonerCampos
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
Dim i As Integer
Dim codpobla As String, despobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If AdoAux(2).Recordset.EOF Then Exit Sub
    
    CargaGrid
    
    ' *** si n'hi han llínies en datagrids ***
    'For i = 0 To DataGridAux.Count - 1
    For i = 0 To 1
        CargaGrid2 i, True
        If Not AdoAux(i).Recordset.EOF Then _
            PonerCamposForma2 Me, AdoAux(i), 2, "FrameAux" & i
    Next i
    
    ' ************* configurar els camps de les descripcions de la capçalera *************
'    text2(3).Text = PonerNombreDeCod(Text1(3), "variedades", "nomvarie")
'    text2(4).Text = PonerNombreDeCod(Text1(4), "forfaits", "nomconfe")
    ' ********************************************************************************
    
    CalcularTotales
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = AdoAux(2).Recordset.AbsolutePosition & " de " & AdoAux(2).Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu
End Sub

Private Sub cmdCancelar_Click()
Dim i As Integer
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
            
        Case 5 ' lineas
            Select Case ModoLineas
                Case 1 'afegir llínia
                    ModoLineas = 0
                    ' *** les llínies que tenen datagrid (en o sense tab) ***
                    If NumTabMto = 0 Or NumTabMto = 1 Or NumTabMto = 2 Or NumTabMto = 4 Then
                        DataGridAux(NumTabMto).AllowAddNew = False
                        ' **** repasar si es diu Data1 l'adodc de la capçalera ***
                        'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'Modificar
                        LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                        DataGridAux(NumTabMto).Enabled = True
                        DataGridAux(NumTabMto).SetFocus


                    End If
                    
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        AdoAux(NumTabMto).Recordset.MoveFirst
                    End If

                Case 2 'modificar llínies
                    ModoLineas = 0
                    
                    ' *** si n'hi han tabs ***
'                    SituarTab (NumTabMto + 1)
                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                    PonerModo 4
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        ' *** l'Index de Fields es el que canvie de la PK de llínies ***
                        V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
                        AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
                        ' ***************************************************************
                    End If

            End Select
            
            TerminaBloquear
'            CargaGrid
'            PosicionarData
            
            ' *** si n'hi han llínies en grids i camps fora d'estos ***
            If Not AdoAux(NumTabMto).Recordset.EOF Then
                DataGridAux_RowColChange NumTabMto, 1, 1
            Else
                LimpiarCamposFrame NumTabMto
            End If
        
    End Select
    
    PonerModo 2

    PonerFocoGrid Me.DataGridAux(2)
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
        Sql = DevuelveDesdeBDNew(cAgro, "cccabdia", "fecha", "fecha", Text1(0).Text, "F", , "codcoste", Text1(1).Text, "N")
        If Sql <> "" Then
            MsgBox "Ya existe el concepto de coste para esta fecha. Modifique.", vbExclamation
            b = False
        End If
    End If
    
    If Modo = 3 Or Modo = 4 Then
    End If
    ' ************************************************************************************
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    Cad = "(fecha=" & DBSet(Text1(0).Text, "F") & " and codcoste = " & DBSet(Text1(1).Text, "N") & " )"
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(adoaux(2), cad, Indicador) Then
    If SituarDataMULTI(AdoAux(2), Cad, Indicador, True) Then
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

    conn.BeginTrans
    ' ***** canviar el nom de la PK de la capçalera, repasar codEmpre *******
    vWhere = " WHERE fecha=" & DBSet(AdoAux(2).Recordset!fecha, "F") & " and codcoste = " & AdoAux(2).Recordset!codCoste
        
    ' ***** elimina les llínies ****
    conn.Execute "DELETE FROM cclindia1 " & vWhere
        
    conn.Execute "DELETE FROM cclindia2 " & vWhere
        
    'Eliminar la CAPÇALERA
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
    CargaGrid
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
    
    ' ***************** configurar els LostFocus dels camps de la capçalera *****************
    Select Case Index
        Case 0 'fecha de coste
            PonerFormatoFecha Text1(Index)
        
        Case 1 'concepto de coste
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "ccconcostes", "nomcoste")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Concepto de Coste: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCon = New frmCCManConcep
                        frmCon.DatosADevolverBusqueda = "0|1|"
                        frmCon.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmCon.Show vbModal
                        Set frmCon = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, AdoAux(2), 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
    End Select
        ' ***************************************************************************
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 0: KEYBusqueda KeyAscii, 1 'concepto de coste
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
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



'************* LLINIES: ****************************
Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
'-- pon el bloqueo aqui
    'If BLOQUEADesdeFormulario2(Me, adoaux(2), 1) Then
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

    ModoLineas = 3 'Posem Modo Eliminar Llínia
    
    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index

    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar(Index) Then Exit Sub
    
    Text1(0).Text = Format(AdoAux(2).Recordset!fecha, "dd/mm/yyyy")
    Text1(1).Text = AdoAux(2).Recordset!codCoste
    
    NumTabMto = Index
    Eliminar = False
   
    vWhere = ObtenerWhereCab(True)
    
    ' ***** independentment de si tenen datagrid o no,
    ' canviar els noms, els formats i el DELETE *****
    Select Case Index
        Case 0 'trabajadores
            Sql = "¿ Seguro que desea eliminar el trabajador del coste diario ?"
            Sql = Sql & vbCrLf & "Trabajador: " & AdoAux(Index).Recordset!codtraba & " " & AdoAux(Index).Recordset!NomTraba
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM cclindia1 "
                Sql = Sql & Replace(vWhere, "cccabdia", "cclindia1") & " AND numlinea= " & AdoAux(Index).Recordset!NumLinea
            End If
            
        Case 1 'categoria
            Sql = "¿ Seguro que desea eliminar la categoria ?"
            Sql = Sql & vbCrLf & "Nombre: " & AdoAux(Index).Recordset!nomcateg
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM cclindia2 "
                Sql = Sql & Replace(vWhere, "cccabdia", "cclindia2") & " AND numlinea= " & AdoAux(Index).Recordset!NumLinea
            End If
            
    End Select

    If Eliminar Then
        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        conn.Execute Sql
        
        ' solo en el caso de que estemos en trabajadores y añadamos una nueva linea hemos de modificar las lineas de categoria
        If Index = 0 Then
            ModificarCategorias (True)
            CargaGrid2 1, True
        End If
        
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        CargaGrid2 Index, True
        If Not SituarDataTrasEliminar(AdoAux(Index), NumRegElim, True) Then
'            PonerCampos
            
        End If
        CalcularTotales
        If BLOQUEADesdeFormulario2(Me, AdoAux(2), 1) Then BotonModificar
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
Dim vWhere As String, vtabla As String
Dim anc As Single
Dim i As Integer
    
    If Index = 2 Then
        NumTabMto = Index
        PonerModo 3, Index
    Else
    
        Text1(0).Text = Format(Me.AdoAux(2).Recordset!fecha, "dd/mm/yyyy")
        Text1(1).Text = Me.AdoAux(2).Recordset!codCoste
    
    
        ModoLineas = 1 'Posem Modo Afegir Llínia
        
        If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Capçalera
            cmdAceptar_Click
            If ModoLineas = 0 Then Exit Sub
        End If
           
        NumTabMto = Index
        PonerModo 5, Index
        
        ' *** bloquejar la clau primaria de la capçalera ***
        BloquearTxt Text1(0), True

    End If

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case Index
        Case 0: vtabla = "cclindia1"
                vWhere = Replace(ObtenerWhereCab(False), "cccabdia", "cclindia1")
        Case 1: vtabla = "cclindia2"
                vWhere = Replace(ObtenerWhereCab(False), "cccabdia", "cclindia2")
        Case 2: vtabla = "cccabdia"
    End Select
    
    
    
    Select Case Index
        Case 0, 1 ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            
            NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", vWhere)

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
                Case 0 'trabajadores
                    txtAux(0).Text = Text1(0).Text 'fecha
                    txtAux(1).Text = Text1(1).Text 'concepto de coste
                    txtAux(2).Text = NumF 'numlinea
                    For i = 3 To 8
                        txtAux(i).Text = ""
                    Next i
                    txtAux2(3).Text = ""
                    
                    txtAux(7).Text = Text1(1).Text
                    txtAux(8).Text = Text1(1).Text
                    
                    BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux0"
                    PonerFoco txtAux(3)
                
                Case 1 'categorias
                    txtAux(40).Text = Text1(0).Text 'fecha
                    txtAux(41).Text = Text1(1).Text 'concepto de coste
                    txtAux(42).Text = NumF
                    txtAux(43).Text = ""
                    txtAux2(43).Text = ""
                    txtAux(44).Text = ""
                    
                    BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux1"
                    PonerFoco txtAux(43)
                    
                    
            End Select
            
        Case 2
            
            AnyadirLinea DataGridAux(Index), AdoAux(Index)
            
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
            
            LLamaLineas Index, 3, anc
        
            Text1(0).Text = ""
            Text1(1).Text = ""
            Text1(2).Text = ""
            Text2(1).Text = ""
            
'                    BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux1"
            PonerFoco Text1(0)
            
    End Select
End Sub

Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim i As Integer
    Dim J As Integer
    
    If Index = 2 Then
        Text1(0).Text = Format(Me.AdoAux(2).Recordset!fecha, "dd/mm/yyyy")
    
        NumTabMto = Index
        PonerModo 4, Index
    
    
        If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
            i = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
            DataGridAux(Index).Scroll 0, i
            DataGridAux(Index).Refresh
        End If
          
        anc = DataGridAux(Index).Top
        If DataGridAux(Index).Row < 0 Then
            anc = anc + 240
        Else
            anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
        End If
        LLamaLineas 2, 4, anc
    
        For J = 0 To 1
            Text1(J).Text = DataGridAux(Index).Columns(J).Text
        Next J
        Text2(1).Text = DataGridAux(Index).Columns(2).Text
        Text1(2).Text = DataGridAux(Index).Columns(3).Text
    
        PonerFoco Text1(2)
    Else
        
          If AdoAux(Index).Recordset.EOF Then Exit Sub
          If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub
          
          ModoLineas = 2 'Modificar llínia
             
          If Modo = 4 Then 'Modificar Capçalera
              cmdAceptar_Click
              If ModoLineas = 0 Then Exit Sub
          End If
             
          NumTabMto = Index
          PonerModo 5, Index
          ' *** bloqueje la clau primaria de la capçalera ***
          BloquearTxt Text1(0), True
        
          Select Case Index
              Case 0, 1 ' *** pose els index de llínies que tenen datagrid (en o sense tab) ***
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
              Case 0 ' trabajadores
              
                  For J = 0 To 2
                      txtAux(J).Text = DataGridAux(Index).Columns(J).Text
                  Next J
                  txtAux(3).Text = DataGridAux(Index).Columns(3).Text
                  txtAux2(3).Text = DataGridAux(Index).Columns(4).Text
                  txtAux(4).Text = DataGridAux(Index).Columns(5).Text
                  txtAux(5).Text = DataGridAux(Index).Columns(6).Text
                  txtAux(6).Text = DataGridAux(Index).Columns(7).Text
                  txtAux(7).Text = DataGridAux(Index).Columns(8).Text
                  txtAux(8).Text = DataGridAux(Index).Columns(9).Text
                  
                  BloquearbtnBuscar Me, Modo, 1, "FrameAux0"
                  
              Case 1 'categorias
                  For J = 40 To 42
                      txtAux(J).Text = DataGridAux(Index).Columns(J - 40).Text
                  Next J
                  txtAux(43).Text = DataGridAux(Index).Columns(3).Text
                  txtAux2(43).Text = DataGridAux(Index).Columns(4).Text
                  txtAux(44).Text = DataGridAux(Index).Columns(5).Text
                  
                  BloquearbtnBuscar Me, Modo, 1, "FrameAux1"
                  
          End Select
          
          LLamaLineas Index, ModoLineas, anc
         
          ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
          Select Case Index
              Case 0 'trabajadores
                  PonerFoco txtAux(3)
              Case 1 'categorias
                  PonerFoco txtAux(43)
          End Select
          ' ***************************************************************************************
          
    End If

End Sub

Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim b As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    DeseleccionaGrid DataGridAux(Index)
       
    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
    Select Case Index
        Case 0 'trabajadores
             For jj = 3 To 6
                txtAux(jj).visible = b
                txtAux(jj).Top = alto
            Next jj
            For jj = 3 To 3
                txtAux2(jj).visible = b
                txtAux2(jj).Top = alto
            Next jj
            btnBuscar(2).visible = b
            btnBuscar(2).Top = alto
            
        Case 1 'categorias
            For jj = 43 To 44
                txtAux(jj).visible = b
                txtAux(jj).Top = alto
            Next jj
            For jj = 43 To 43
                txtAux2(jj).visible = b
                txtAux2(jj).Top = alto
            Next jj
            
            btnBuscar(3).visible = b
            btnBuscar(3).Top = alto
            
        Case 2 ' cabecera
            b = (xModo = 1 Or xModo = 3)
            For jj = 0 To 1
                Text1(jj).visible = b
                Text1(jj).Top = alto
            Next jj
            For jj = 0 To 1
                btnBuscar(jj).visible = b
                btnBuscar(jj).Top = alto
            Next jj
            Text2(1).visible = b
            Text2(1).Top = alto
            
            Text1(2).visible = b Or Modo = 4
            Text1(2).Top = alto
    End Select
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
    
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    

    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 3 ' trabajador
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(3).Text = PonerNombreDeCod(txtAux(Index), "straba", "nomtraba")
                If txtAux2(3).Text = "" Then
                    cadMen = "No existe el Trabajador: " & txtAux(Index).Text & ". Reintroduzca." & vbCrLf
                    PonerFoco txtAux(Index)
                End If
            Else
                txtAux2(3).Text = ""
            End If
        
        Case 7, 8 ' fecha inicio y fecha fin
            PonerFormatoFecha txtAux(Index)
            
        Case 4, 5 ' hora inicio y hora fin
            If PonerFormatoHora(txtAux(Index)) Then
                If Index = 5 Then Me.cmdAceptar.SetFocus
            End If
        
        Case 44 ' horas
            If PonerFormatoDecimal(txtAux(Index), 9) Then cmdAceptar.SetFocus
            
        Case 43 ' categoria
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), "salarios", "nomcateg")
                If txtAux2(Index).Text = "" Then
                    cadMen = "No existe la Categoria: " & txtAux(Index).Text & ". Reintroduzca." & vbCrLf
                    PonerFoco txtAux(Index)
                End If
            Else
                txtAux2(Index).Text = ""
            End If
            
    End Select
    
    If Index = 4 Then
        txtAux(7).Text = Trim(Format(Text1(0).Text, "yyyy-mm-dd") & " " & Format(txtAux(4).Text, "hh:mm:ss"))
    End If
    
    If Index = 5 Then
        txtAux(8).Text = Trim(Format(Text1(0).Text, "yyyy-mm-dd") & " " & Format(txtAux(5).Text, "hh:mm:ss"))
    End If
    
    If txtAux(4).Text <> "" And txtAux(5).Text <> "" And txtAux(7).Text <> "" And txtAux(8).Text <> "" Then
        txtAux(6).Text = Round2(DateDiff("n", CDate(txtAux(7).Text), CDate(txtAux(8).Text)) / 60, 2)
    End If
    
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

Private Function DatosOkLlin(nomFrame As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim b As Boolean
Dim Cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte

    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False
        
    b = CompForm2(Me, 2, nomFrame) 'Comprovar formato datos ok
    If Not b Then Exit Function
    
    If b And nomFrame = "FrameAux0" Then
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
    
    If b And nomFrame = "FrameAux1" Then
        If Modo = 3 Or Modo = 4 Then
            Sql = "select count(*) from cclindia2 where codorden= " & DBSet(Text1(0).Text, "N")
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
Dim i As Byte

    If Index = 2 Then
        If Modo = 3 Then
            CargaGrid2 0, False
            CargaGrid2 1, False
        Else
            If DataGridAux(Index).Columns.Count > 1 Then
               CargaGrid2 0, True
               CargaGrid2 1, True
               
               Text1(0).Text = Format(Me.AdoAux(2).Recordset!fecha, "dd/mm/yyyy")
               Text1(1).Text = Me.AdoAux(2).Recordset!codCoste
               
               CalcularTotales
               '-- Esto permanece para saber donde estamos
               lblIndicador.Caption = AdoAux(2).Recordset.AbsolutePosition & " de " & AdoAux(2).Recordset.RecordCount
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
    
'    adodc1.ConnectionString = Conn
    If vSQL <> "" Then
        Sql = CadenaConsulta & " AND " & vSQL
    Else
        Sql = CadenaConsulta
    End If
    '********************* canviar el ORDER BY *********************++
    Sql = Sql & " ORDER BY cccabdia.fecha, cccabdia.codcoste"
    '**************************************************************++
    
    CargaGridGnral Me.DataGridAux(2), Me.AdoAux(2), Sql, PrimeraVez
    
    ' *******************canviar els noms i si fa falta la cantitat********************
    tots = "S|text1(0)|T|Fecha|1200|;S|btnBuscar(0)|B||195|;"
    tots = tots & "S|text1(1)|T|Concepto|1100|;S|btnBuscar(1)|B||195|;S|text2(1)|T|Descripcion|2800|;"
    tots = tots & "S|text1(2)|T|Observaciones|7500|;"
    
    arregla tots, DataGridAux(2), Me
    
    DataGridAux(2).ScrollBars = dbgAutomatic
    
    DataGridAux(2).Columns(1).Alignment = dbgLeft
    DataGridAux(2).Columns(2).Alignment = dbgLeft
    DataGridAux(2).Columns(3).Alignment = dbgLeft
    
    b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
    
End Sub


Private Sub CargaGrid2(Index As Integer, enlaza As Boolean)
Dim b As Boolean
Dim i As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)

    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
    
    Select Case Index
        Case 0 'trabajadores
            txtAux(4).Tag = "Fecha Ini|FHH|N|||cclindia1|fechaini|hh:mm:ss||"
            txtAux(5).Tag = "Fecha fin|FHH|N|||cclindia1|fechafin|hh:mm:ss||"
        
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;N||||0|;" 'fecha,codorden, numlinea
            tots = tots & "S|txtAux(3)|T|Código|1000|;S|btnBuscar(2)|B|||;"
            tots = tots & "S|txtAux2(3)|T|Trabajador|4000|;S|txtAux(4)|T|H.Inicio|1000|;"
            tots = tots & "S|txtAux(5)|T|H.Fin|1000|;S|txtAux(6)|T|Horas|1300|;N||||0|;N||||0|;"
            
            arregla tots, DataGridAux(Index), Me, 350
        
'            DataGridAux(0).Columns(6).NumberFormat = "dd/mm/yyyy"
'            DataGridAux(0).Columns(8).NumberFormat = "dd/mm/yyyy"
            
            DataGridAux(0).Columns(2).Alignment = dbgLeft
            DataGridAux(0).Columns(3).Alignment = dbgLeft
            DataGridAux(0).Columns(4).Alignment = dbgLeft
            DataGridAux(0).Columns(5).Alignment = dbgLeft

            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
            txtAux(4).Tag = ""
            txtAux(5).Tag = ""
            
        Case 1 'categorias
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;N||||0|;" 'fecha,codorden,numlinea
            tots = tots & "S|txtAux(43)|T|Categoria|1000|;S|btnBuscar(3)|B||195|;"
            tots = tots & "S|txtAux2(43)|T|Denominación|3380|;"
            tots = tots & "S|txtAux(44)|T|Horas|1200|;"
            
            arregla tots, DataGridAux(Index), Me
            
            DataGridAux(1).Columns(2).Alignment = dbgLeft
            DataGridAux(1).Columns(4).Alignment = dbgLeft
            
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
        Case 2 ' cabecera de costes
            tots = "S|text1(0)|T|Fecha|1200|;S|btnBuscar(0)|B||195|;"
            tots = tots & "S|text1(1)|T|Concepto|2200|;S|btnBuscar(1)|B||195|;S|text2(1)|T|Descripcion|3450|;"
            tots = tots & "S|text1(2)|T|Observaciones|7500|;"
            
            arregla tots, DataGridAux(Index), Me
            
            DataGridAux(2).Columns(2).Alignment = dbgLeft
            DataGridAux(2).Columns(3).Alignment = dbgLeft
            
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
        
    End Select
    
    DataGridAux(Index).ScrollBars = dbgAutomatic
    
'    PonerModoOpcionesMenu Modo
    
    ' **** si n'hi han llínies en grids i camps fora d'estos ****
'    If Not AdoAux(Index).Recordset.EOF Then
'        DataGridAux_RowColChange Index, 1, 1
'    Else
''        LimpiarCamposFrame Index
'    End If
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub

Private Sub ModificarCategorias(Eliminar As Boolean)
Dim Sql As String
Dim Sql2 As String
Dim Categoria As Integer
Dim NumF As String
Dim Horas As String
Dim Rs As ADODB.Recordset

    Sql = "select * from cclindia1 where fecha = " & DBSet(Text1(0).Text, "F")
    Sql = Sql & " and codcoste = " & DBSet(Text1(1).Text, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    While Not Rs.EOF
    
        Sql = "select codcateg from straba where codtraba = " & DBSet(Rs!codtraba, "N")
        Categoria = DevuelveValor(Sql)
    
        Sql = "select count(*) from cclindia2 where fecha = " & DBSet(Text1(0).Text, "F")
        Sql = Sql & " and codcoste = " & DBSet(Rs!codCoste, "N")
        Sql = Sql & " and codcateg = " & DBSet(Categoria, "N")
        
        If TotalRegistros(Sql) = 0 Then
            NumF = SugerirCodigoSiguienteStr("cclindia2", "numlinea", "fecha = " & DBSet(Text1(0).Text, "F") & " and codcoste = " & DBSet(Text1(1).Text, "N"))
        
            Sql2 = "insert into cclindia2 (fecha,codcoste,numlinea,codcateg,horas) values ("
            Sql2 = Sql2 & DBSet(Text1(0).Text, "F") & "," & DBSet(Text1(1).Text, "N") & "," & DBSet(NumF, "N") & ","
            Sql2 = Sql2 & DBSet(Categoria, "N") & "," & DBSet(Rs!Horas, "N") & ")"
            
            conn.Execute Sql2
        
        Else
            Sql2 = "select sum(horas) from cclindia1 where fecha = " & DBSet(Text1(0).Text, "F")
            Sql2 = Sql2 & " and codcoste = " & DBSet(Text1(1).Text, "N")
            Sql2 = Sql2 & " and codtraba in ( select codtraba from straba where codcateg = " & DBSet(Categoria, "N") & ")"
            
            Horas = DevuelveValor(Sql2)
        
            Sql2 = "update cclindia2 set horas = " & DBSet(ImporteSinFormato(Horas), "N")
            Sql2 = Sql2 & " where fecha = " & DBSet(Text1(0).Text, "F")
            Sql2 = Sql2 & " and codcoste = " & DBSet(Text1(1).Text, "N")
            Sql2 = Sql2 & " and codcateg = " & DBSet(Categoria, "N")
                
            conn.Execute Sql2
                
            Sql2 = "delete from cclindia2 where fecha = " & DBSet(Text1(0).Text, "F")
            Sql2 = Sql2 & " and codcoste = " & DBSet(Text1(1).Text, "N")
            Sql2 = Sql2 & " and horas = 0"
            
            conn.Execute Sql2
        End If
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
End Sub


Private Sub InsertarLinea()
'Inserta registre en les taules de Llínies
Dim nomFrame As String
Dim b As Boolean

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomFrame = "FrameAux0" 'trabajadores
        Case 1: nomFrame = "FrameAux1" 'categorias
    End Select
    
    
    If DatosOkLlin(nomFrame) Then
        If NumTabMto = 0 Then
            txtAux(7).Text = Format(Text1(0).Text, "yyyy-mm-dd") & " " & Format(txtAux(4).Text, "hh:mm:ss")
            txtAux(8).Text = Format(Text1(0).Text, "yyyy-mm-dd") & " " & Format(txtAux(5).Text, "hh:mm:ss")
        End If
        
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomFrame) Then
            
            ' solo en el caso de que estemos en trabajadores y añadamos una nueva linea hemos de modificar las lineas de categoria
            If NumTabMto = 0 Then
                ModificarCategorias False
                CargaGrid2 1, True
            End If
        
            b = BLOQUEADesdeFormulario2(Me, AdoAux(2), 1)
            Select Case NumTabMto
                Case 0, 1 ' *** els index de les llinies en grid (en o sense tab) ***
                     CargaGrid2 NumTabMto, True
                    If b Then BotonAnyadirLinea NumTabMto
            End Select
           
'            SituarTab (NumTabMto + 1)
        End If
    End If
End Sub

Private Function ModificarLinea() As Boolean
'Modifica registre en les taules de Llínies
Dim nomFrame As String
Dim V As Integer
    
    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomFrame = "FrameAux0" 'trabajadores
        Case 1: nomFrame = "FrameAux1" 'categorias
    End Select
    ModificarLinea = False
    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomFrame) Then
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
'    vWhere = vWhere & " fecha=" & DBSet(text1(0).Text, "F") & " and cccabdia.codcoste = " & DBSet(text1(1).Text, "N")
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

Private Sub CalcularPrecio()
    txtAux2(1).Text = ""
    If txtAux(2).Text <> "" And txtAux2(0).Text <> "" Then
        txtAux2(1).Text = Round2(ImporteSinFormato(txtAux(2).Text) * ImporteSinFormato(txtAux2(0).Text), 4)
    End If
End Sub

Private Sub VisualizaPrecio()
    Select Case vParamAplic.TipoPrecio
        Case 0
            txtAux2(0).Text = DevuelveDesdeBDNew(cAgro, "sartic", "preciomp", "codartic", txtAux(1), "T")
        Case 1
            txtAux2(0).Text = DevuelveDesdeBDNew(cAgro, "sartic", "preciouc", "codartic", txtAux(1), "T")
    End Select
End Sub

Private Sub CalcularTotales()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim TotalEnvases As String
Dim TotalCostes As String
Dim Valor As Currency

    On Error Resume Next

    ' Coste de trabajadores

    Sql = "select sum(total) from ("
    Sql = Sql & "select salarios.codcateg, round(sum(cclindia1.horas) * salarios.impsalar,2) total "
    Sql = Sql & " from (cclindia1 inner join straba on cclindia1.codtraba = straba.codtraba) "
    Sql = Sql & " inner join salarios on straba.codcateg = salarios.codcateg "
    Sql = Sql & " where cclindia1.fecha = " & DBSet(Text1(0).Text, "F")
    Sql = Sql & " and cclindia1.codcoste = " & DBSet(Text1(1).Text, "N")
    Sql = Sql & " group by 1) aaaaa "
    
    Text3(0).Text = DevuelveValor(Sql)
    
    Valor = CCur(TransformaPuntosComas(Text3(0).Text))
    If Valor <> 0 Then
        Text3(0).Text = Format(Valor, "###,###,##0.00")
    Else
        Text3(0).Text = ""
    End If
    
    ' Horas de trabajadores
    
    Sql = "select sum(cclindia1.horas) "
    Sql = Sql & " from cclindia1 where fecha = " & DBSet(Text1(0).Text, "F")
    Sql = Sql & " and codcoste = " & DBSet(Text1(1).Text, "N")
    
    Text3(2).Text = DevuelveValor(Sql)
    
    Valor = CCur(TransformaPuntosComas(Text3(2).Text))
    If Valor <> 0 Then
        Text3(2).Text = Format(Valor, "###,###,##0.00")
    Else
        Text3(2).Text = ""
    End If
    
    
    ' Coste de categorias

    Sql = "select round(sum(cclindia2.horas) * salarios.impsalar,2) "
    Sql = Sql & " from cclindia2 inner join salarios on cclindia2.codcateg = salarios.codcateg "
    Sql = Sql & " where cclindia2.fecha = " & DBSet(Text1(0).Text, "F")
    Sql = Sql & " and cclindia2.codcoste = " & DBSet(Text1(1).Text, "N")
    
    Text3(1).Text = DevuelveValor(Sql)
    
    Valor = CCur(TransformaPuntosComas(Text3(1).Text))
    If Valor <> 0 Then
        Text3(1).Text = Format(Valor, "###,###,##0.00")
    Else
        Text3(1).Text = ""
    End If
    
    ' Horas de trabajadores
    
    Sql = "select sum(cclindia2.horas) from cclindia2 "
    Sql = Sql & " where cclindia2.fecha = " & DBSet(Text1(0).Text, "F")
    Sql = Sql & " and cclindia2.codcoste = " & DBSet(Text1(1).Text, "N")
    
    Text3(3).Text = DevuelveValor(Sql)
    
    Valor = CCur(TransformaPuntosComas(Text3(3).Text))
    If Valor <> 0 Then
        Text3(3).Text = Format(Valor, "###,###,##0.00")
    Else
        Text3(3).Text = ""
    End If
    
    If Err.Number <> 0 Then
        Err.Clear
    End If

End Sub

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
                CadenaConsulta = "Select * from " & NombreTabla & " where codorden = " & DBSet(Text1(0).Text, "N") & Ordenacion
                PonerCadenaBusqueda
                PonerModo 2
                'Ponerse en Modo Insertar Lineas
                BotonAnyadirLinea 0
            End If
        End If
        Text1(0).Text = Format(Text1(0).Text, "0000000")
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
    'Comprobar si mientras tanto se incremento el contador de Albaranes
    'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
    Do
        devuelve = DevuelveDesdeBDNew(cAgro, NombreTabla, "codorden", "codorden", Text1(0).Text, "N")
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
    MenError = "Error al insertar en la tabla Cabecera de Ordenes de Confeccion (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
        
    MenError = "Error al actualizar el contador del Albarán."
    vTipoMov.IncrementarContador (CodTipoMov)
    
EInsertarOferta:
        If Err.Number <> 0 Then
            MenError = "Insertando Orden." & vbCrLf & "----------------------------" & vbCrLf & MenError
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

