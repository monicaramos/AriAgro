VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCCOrdenConfeccion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordenes de Confección"
   ClientHeight    =   10320
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   13575
   Icon            =   "frmCCOrdenConfeccion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10320
   ScaleWidth      =   13575
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
      Left            =   10890
      TabIndex        =   99
      Top             =   315
      Width           =   1605
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   3735
      TabIndex        =   97
      Top             =   90
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   98
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
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   45
      TabIndex        =   95
      Top             =   90
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   96
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
   Begin VB.Frame FrameResumen 
      Caption         =   "RESUMEN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   7065
      Left            =   45
      TabIndex        =   65
      Top             =   2520
      Visible         =   0   'False
      Width           =   13445
      Begin VB.Frame FrameAux3 
         Caption         =   "Forfaits"
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
         Height          =   2655
         Left            =   1200
         TabIndex        =   82
         Top             =   3450
         Width           =   11420
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
            Height          =   360
            Index           =   7
            Left            =   7605
            TabIndex        =   90
            Text            =   "Text3"
            Top             =   2160
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
            Height          =   360
            Index           =   6
            Left            =   4410
            TabIndex        =   89
            Text            =   "Text3"
            Top             =   2160
            Width           =   1300
         End
         Begin VB.TextBox txtAux1 
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
            Left            =   3810
            MaxLength       =   10
            TabIndex        =   88
            Tag             =   "Kilos|N|N|||cclinorden4|kilosnet|##,###,##0||"
            Text            =   "kilos"
            Top             =   1860
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.TextBox txtAux1 
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
            Index           =   41
            Left            =   750
            MaxLength       =   16
            TabIndex        =   87
            Tag             =   "Linea|N|N|||cclinorden4|numlinea|00|S|"
            Text            =   "lin"
            Top             =   1830
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
            Index           =   4
            Left            =   1710
            TabIndex        =   86
            Text            =   "nombre"
            Top             =   1830
            Visible         =   0   'False
            Width           =   2040
         End
         Begin VB.TextBox txtAux1 
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
            Left            =   1110
            MaxLength       =   16
            TabIndex        =   85
            Tag             =   "Codigo Forfait|T|N|||cclinorden4|codforfait|||"
            Text            =   "forfait"
            Top             =   1830
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtAux1 
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
            Left            =   5760
            MaxLength       =   10
            TabIndex        =   84
            Tag             =   "Cajas|N|N|||cclinorden4|numcajon|##,###,##0||"
            Text            =   "cajas"
            Top             =   1860
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.TextBox txtAux1 
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
            TabIndex        =   83
            Tag             =   "Codigo Orden|N|N|||cclinorden4|codorden||S|"
            Text            =   "orden"
            Top             =   1830
            Visible         =   0   'False
            Width           =   555
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   3
            Left            =   90
            TabIndex        =   91
            Top             =   330
            Visible         =   0   'False
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
            Enabled         =   0   'False
         End
         Begin MSDataGridLib.DataGrid DataGridAux 
            Bindings        =   "frmCCOrdenConfeccion.frx":000C
            Height          =   1710
            Index           =   3
            Left            =   180
            TabIndex        =   92
            Top             =   360
            Width           =   10880
            _ExtentX        =   19182
            _ExtentY        =   3016
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
         Begin MSAdodcLib.Adodc AdoAux 
            Height          =   375
            Index           =   3
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
            Caption         =   "Total Cajas"
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
            Height          =   330
            Index           =   7
            Left            =   6150
            TabIndex        =   94
            Top             =   2160
            Width           =   1440
         End
         Begin VB.Label Label11 
            Caption         =   "TOTAL KILOS: "
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
            Index           =   6
            Left            =   2610
            TabIndex        =   93
            Top             =   2160
            Width           =   3420
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
         Height          =   2655
         Left            =   1200
         TabIndex        =   66
         Top             =   630
         Width           =   11420
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
            TabIndex        =   77
            Tag             =   "Codigo Orden|N|N|||cclinorden2|codorden||S|"
            Text            =   "orden"
            Top             =   1830
            Visible         =   0   'False
            Width           =   555
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
            Left            =   6360
            MaxLength       =   7
            TabIndex        =   76
            Tag             =   "Horas|N|N|||cclinorden2|horas|###0.00||"
            Text            =   "horas"
            Top             =   1860
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
            Index           =   42
            Left            =   1110
            MaxLength       =   4
            TabIndex        =   75
            Tag             =   "Codigo Coste|N|N|||cclinorden2|codcoste|0000|S|"
            Text            =   "Cost"
            Top             =   1830
            Visible         =   0   'False
            Width           =   375
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
            Index           =   42
            Left            =   1710
            TabIndex        =   74
            Text            =   "nombre"
            Top             =   1830
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
            Index           =   1
            Left            =   1560
            MaskColor       =   &H00000000&
            TabIndex        =   73
            ToolTipText     =   "Buscar zona"
            Top             =   1800
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
            Index           =   41
            Left            =   750
            MaxLength       =   16
            TabIndex        =   72
            Tag             =   "Linea|N|N|||cclinorden2|numlinea|00|S|"
            Text            =   "lin"
            Top             =   1830
            Visible         =   0   'False
            Width           =   240
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
            Index           =   5
            Left            =   4380
            MaskColor       =   &H00000000&
            TabIndex        =   71
            ToolTipText     =   "Buscar categoria"
            Top             =   1830
            Visible         =   0   'False
            Width           =   195
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
            Left            =   4590
            TabIndex        =   70
            Text            =   "nombre"
            Top             =   1860
            Visible         =   0   'False
            Width           =   2040
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
            Left            =   3810
            MaxLength       =   2
            TabIndex        =   69
            Tag             =   "Categoria|N|N|||cclinorden2|codcateg|00||"
            Text            =   "cate"
            Top             =   1860
            Visible         =   0   'False
            Width           =   555
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
            Height          =   360
            Index           =   1
            Left            =   4410
            TabIndex        =   68
            Text            =   "Text3"
            Top             =   2160
            Width           =   1300
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
            Height          =   360
            Index           =   3
            Left            =   7605
            TabIndex        =   67
            Text            =   "Text3"
            Top             =   2160
            Width           =   1200
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   1
            Left            =   90
            TabIndex        =   78
            Top             =   330
            Visible         =   0   'False
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
            Enabled         =   0   'False
         End
         Begin MSDataGridLib.DataGrid DataGridAux 
            Bindings        =   "frmCCOrdenConfeccion.frx":0024
            Height          =   1710
            Index           =   1
            Left            =   180
            TabIndex        =   79
            Top             =   360
            Width           =   10880
            _ExtentX        =   19182
            _ExtentY        =   3016
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
            Left            =   2610
            TabIndex        =   81
            Top             =   2160
            Width           =   3420
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
            Height          =   330
            Index           =   3
            Left            =   6150
            TabIndex        =   80
            Top             =   2160
            Width           =   1440
         End
      End
   End
   Begin VB.Frame FrameAux2 
      Caption         =   "Variedades/Forfaits"
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
      Height          =   3240
      Left            =   60
      TabIndex        =   48
      Top             =   5820
      Width           =   11445
      Begin VB.TextBox txtAux1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   0
         Left            =   45
         MaxLength       =   16
         TabIndex        =   57
         Tag             =   "CodigoOrden|N|N|||cclinorden3|codorden||S|"
         Text            =   "Orden"
         Top             =   2565
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtAux1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   2
         Left            =   765
         MaxLength       =   6
         TabIndex        =   56
         Tag             =   "Codigo Variedad|N|N|||cclinorden3|codvarie|000000|N|"
         Text            =   "varie"
         Top             =   2565
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtAux1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   3
         Left            =   2040
         MaxLength       =   16
         TabIndex        =   58
         Tag             =   "Forfait|T|N|||cclinorden3|codforfait|||"
         Text            =   "forfai"
         Top             =   2550
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   55
         Top             =   2550
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.CommandButton btnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   300
         Index           =   6
         Left            =   1200
         MaskColor       =   &H00000000&
         TabIndex        =   54
         ToolTipText     =   "Buscar concepto coste"
         Top             =   2550
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
         Left            =   3045
         TabIndex        =   53
         Top             =   2550
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox txtAux1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   1
         Left            =   450
         MaxLength       =   16
         TabIndex        =   52
         Tag             =   "Linea|N|N|||cclinorden3|numlinea|0000|S|"
         Text            =   "lin"
         Top             =   2565
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.CommandButton btnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   300
         Index           =   7
         Left            =   2820
         MaskColor       =   &H00000000&
         TabIndex        =   51
         ToolTipText     =   "Buscar trabajador"
         Top             =   2520
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   4
         Left            =   3810
         MaxLength       =   10
         TabIndex        =   60
         Tag             =   "Kilos|N|N|||cclinorden3|kilosnet|##,###,##0||"
         Text            =   "kilos"
         Top             =   2550
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.TextBox txtAux1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   5
         Left            =   4830
         MaxLength       =   10
         TabIndex        =   62
         Tag             =   "Kilos|N|N|||cclinorden3|numcajon|##,###,##0||"
         Text            =   "cajas"
         Top             =   2520
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   9645
         TabIndex        =   50
         Text            =   "Text3"
         Top             =   2850
         Width           =   1380
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   6150
         TabIndex        =   49
         Text            =   "Text3"
         Top             =   2850
         Width           =   1300
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   2
         Left            =   135
         TabIndex        =   59
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
         Index           =   2
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
         Bindings        =   "frmCCOrdenConfeccion.frx":003C
         Height          =   2010
         Index           =   2
         Left            =   135
         TabIndex        =   61
         Top             =   720
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   3545
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
      Begin VB.Label Label11 
         Caption         =   "Total Cajas"
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
         Index           =   5
         Left            =   8190
         TabIndex        =   64
         Top             =   2850
         Width           =   1440
      End
      Begin VB.Label Label11 
         Caption         =   "TOTAL KILOS: "
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
         Index           =   4
         Left            =   4350
         TabIndex        =   63
         Top             =   2850
         Width           =   1995
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1500
      Index           =   0
      Left            =   60
      TabIndex        =   11
      Top             =   900
      Width           =   13445
      Begin VB.TextBox text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   6
         Left            =   6585
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Tag             =   "Observaciones|T|S|||caborden|observac|||"
         Top             =   540
         Width           =   6045
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   7590
         MaxLength       =   6
         TabIndex        =   47
         Text            =   "123456"
         Top             =   780
         Width           =   720
      End
      Begin VB.TextBox text1 
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
         Left            =   1365
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "1234567890"
         Top             =   540
         Width           =   1350
      End
      Begin VB.TextBox text2 
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
         Left            =   2340
         TabIndex        =   30
         Text            =   "12345678901234567890"
         Top             =   960
         Width           =   4095
      End
      Begin VB.TextBox text1 
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
         Left            =   1365
         MaxLength       =   6
         TabIndex        =   5
         Tag             =   "Cliente|N|S|||cccaborden|codclien|000000||"
         Text            =   "123456"
         Top             =   960
         Width           =   900
      End
      Begin VB.TextBox text1 
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
         Left            =   2775
         MaxLength       =   8
         TabIndex        =   2
         Text            =   "12345678"
         Top             =   540
         Width           =   1110
      End
      Begin VB.TextBox text1 
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
         Index           =   9
         Left            =   5340
         MaxLength       =   8
         TabIndex        =   4
         Top             =   540
         Width           =   1110
      End
      Begin VB.TextBox text1 
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
         Index           =   8
         Left            =   3945
         MaxLength       =   10
         TabIndex        =   3
         Text            =   "1234567890"
         Top             =   540
         Width           =   1350
      End
      Begin VB.TextBox text1 
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
         Left            =   270
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Código Orden|N|S|||cccaborden|codorden||S|"
         Text            =   "1234567"
         Top             =   540
         Width           =   945
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   8250
         MaxLength       =   20
         TabIndex        =   39
         Tag             =   "Fecha Inicio|FH|N|||cccaborden|fechaini|yyyy-mm-dd hh:mm:ss||"
         Text            =   "1234567890"
         Top             =   780
         Width           =   1875
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   7440
         MaxLength       =   20
         TabIndex        =   40
         Tag             =   "Fecha Fin|FH|S|||cccaborden|fechafin|yyyy-mm-dd hh:mm:ss||"
         Text            =   "1234567890"
         Top             =   900
         Width           =   1845
      End
      Begin VB.Image imgDoc 
         Height          =   345
         Index           =   1
         Left            =   12180
         ToolTipText     =   "Mostrar Resumen"
         Top             =   150
         Width           =   405
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   5040
         Picture         =   "frmCCOrdenConfeccion.frx":0054
         ToolTipText     =   "Buscar fecha"
         Top             =   270
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   2475
         Picture         =   "frmCCOrdenConfeccion.frx":00DF
         ToolTipText     =   "Buscar fecha"
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label10 
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
         Left            =   270
         TabIndex        =   31
         Top             =   960
         Width           =   720
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1095
         ToolTipText     =   "Buscar Variedad"
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label5 
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
         Height          =   255
         Left            =   5340
         TabIndex        =   29
         Top             =   300
         Width           =   825
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Fin"
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
         Left            =   3945
         TabIndex        =   28
         Top             =   300
         Width           =   1065
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   2805
         TabIndex        =   27
         Top             =   300
         Width           =   780
      End
      Begin VB.Label Label29 
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
         Left            =   6570
         TabIndex        =   15
         Top             =   300
         Width           =   1485
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   8100
         ToolTipText     =   "Zoom descripción"
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label18 
         Caption         =   "Fec.Inicio"
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
         Left            =   1395
         TabIndex        =   14
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
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
         Left            =   270
         TabIndex        =   12
         Top             =   300
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   90
      TabIndex        =   9
      Top             =   9630
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
         TabIndex        =   10
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
      Left            =   12375
      TabIndex        =   8
      Top             =   9765
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
      Left            =   11160
      TabIndex        =   7
      Top             =   9765
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   660
      Top             =   3630
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
      Left            =   12375
      TabIndex        =   13
      Top             =   9765
      Visible         =   0   'False
      Width           =   1065
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
      Height          =   3810
      Left            =   45
      TabIndex        =   16
      Top             =   2565
      Width           =   13445
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
         Height          =   360
         Index           =   0
         Left            =   6150
         TabIndex        =   45
         Text            =   "Text3"
         Top             =   3375
         Width           =   1300
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
         Height          =   360
         Index           =   2
         Left            =   9645
         TabIndex        =   43
         Text            =   "Text3"
         Top             =   3375
         Width           =   1380
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
         Index           =   10
         Left            =   8820
         MaxLength       =   20
         TabIndex        =   42
         Tag             =   "Fecha Fin|FH|N|||cclinorden1|fechafin|yyyy-mm-dd hh:mm:ss||"
         Text            =   "f.fin"
         Top             =   2520
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
         Index           =   9
         Left            =   7950
         MaxLength       =   20
         TabIndex        =   41
         Tag             =   "Fecha Ini|FH|N|||cclinorden1|fechaini|yyyy-mm-dd hh:mm:ss||"
         Text            =   "f.ini"
         Top             =   2520
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
         Index           =   8
         Left            =   7350
         MaxLength       =   7
         TabIndex        =   24
         Tag             =   "Horas|N|N|||cclinorden1|horas|###0.00||"
         Text            =   "horas"
         Top             =   2490
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
         Index           =   7
         Left            =   6540
         MaxLength       =   8
         TabIndex        =   23
         Text            =   "h.fin"
         Top             =   2490
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
         Index           =   4
         Left            =   6300
         MaskColor       =   &H00000000&
         TabIndex        =   38
         ToolTipText     =   "Buscar fecha"
         Top             =   2490
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
         Index           =   6
         Left            =   5580
         MaxLength       =   10
         TabIndex        =   22
         Text            =   "f.fin"
         Top             =   2520
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
         Index           =   5
         Left            =   4800
         MaxLength       =   8
         TabIndex        =   21
         Text            =   "h.inic"
         Top             =   2520
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
         Index           =   3
         Left            =   4590
         MaskColor       =   &H00000000&
         TabIndex        =   37
         ToolTipText     =   "Buscar fecha"
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
         Index           =   4
         Left            =   3810
         MaxLength       =   10
         TabIndex        =   20
         Text            =   "f.inic"
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
         Left            =   2820
         MaskColor       =   &H00000000&
         TabIndex        =   36
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
         Index           =   1
         Left            =   450
         MaxLength       =   16
         TabIndex        =   35
         Tag             =   "Linea|N|N|||cclinorden1|numlinea|0000|S|"
         Text            =   "lin"
         Top             =   2565
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
         Left            =   3045
         TabIndex        =   34
         Top             =   2550
         Visible         =   0   'False
         Width           =   705
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
         Left            =   1200
         MaskColor       =   &H00000000&
         TabIndex        =   33
         ToolTipText     =   "Buscar concepto coste"
         Top             =   2550
         Visible         =   0   'False
         Width           =   195
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
         Index           =   2
         Left            =   1440
         TabIndex        =   32
         Top             =   2550
         Visible         =   0   'False
         Width           =   525
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
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   19
         Tag             =   "Trabajador|N|N|||cclinorden1|codtraba|000000||"
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
         Index           =   2
         Left            =   765
         MaxLength       =   4
         TabIndex        =   18
         Tag             =   "Concepto Coste|N|N|||cclinorden1|codcoste|0000|N|"
         Text            =   "Cost"
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
         TabIndex        =   17
         Tag             =   "CodigoOrden|N|N|||cclinorden1|codorden||S|"
         Text            =   "Orden"
         Top             =   2565
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   0
         Left            =   135
         TabIndex        =   25
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
         Bindings        =   "frmCCOrdenConfeccion.frx":016A
         Height          =   2610
         Index           =   0
         Left            =   135
         TabIndex        =   26
         Top             =   690
         Width           =   13160
         _ExtentX        =   23204
         _ExtentY        =   4604
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
         Left            =   4350
         TabIndex        =   46
         Top             =   3420
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
         Left            =   8190
         TabIndex        =   44
         Top             =   3420
         Width           =   1440
      End
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   13005
      TabIndex        =   100
      Top             =   270
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
         Enabled         =   0   'False
         Shortcut        =   ^I
         Visible         =   0   'False
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
Attribute VB_Name = "frmCCOrdenConfeccion"
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

Private WithEvents frmVar As frmManVariedad 'variedades
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmZon As frmBasico 'zonas
Attribute frmZon.VB_VarHelpID = -1
Private WithEvents frmTra As frmBasico 'trabajadores ( de recoleccion )
Attribute frmTra.VB_VarHelpID = -1
Private WithEvents frmCat As frmBasico 'salarios o categorias ( de recoleccion )
Attribute frmCat.VB_VarHelpID = -1
Private WithEvents frmCos As frmCCManConcep
Attribute frmCos.VB_VarHelpID = -1
Private WithEvents frmCli As frmClientes 'clientes
Attribute frmCli.VB_VarHelpID = -1

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
        Case 0 ' Concepto de coste
            indice = 2
            
            Set frmCos = New frmCCManConcep
            frmCos.CodigoActual = txtAux(indice).Text
            frmCos.DatosADevolverBusqueda = "0|1"
            frmCos.Show vbModal
            Set frmCos = Nothing
            PonerFoco txtAux(2)
            
        Case 2 ' Trabajadores
            indice = 3
            Set frmTra = New frmBasico
            AyudaTrabajadores frmTra, txtAux(3)
            Set frmTra = Nothing
            PonerFoco txtAux(3)
            
        Case 3 ' Fecha Inicio
            btnFec (3)
            
        Case 4 ' Fecha Fin
            btnFec (4)
            
        Case 1 ' Zona de Categoria
            indice = 42
            Set frmZon = New frmBasico
            AyudaZonasCC frmZon, txtAux(42)
            Set frmZon = Nothing
            PonerFoco txtAux(42)
        
        Case 5 ' Categoria
            indice = 43
            Set frmCat = New frmBasico
            AyudaCategorias frmCat, txtAux(43)
            Set frmCat = Nothing
            PonerFoco txtAux(43)
            
        Case 6 ' Variedades
            indice = 2
            Set frmVar = New frmManVariedad
            frmVar.DatosADevolverBusqueda = "0|1|"
            frmVar.CodigoActual = txtAux1(indice).Text
            frmVar.Show vbModal
            Set frmVar = Nothing
            PonerFoco txtAux1(indice)
        
        Case 7 ' Forfaits
            indice = 3
            Set frmFor = New frmManForfaits
            frmFor.DatosADevolverBusqueda = "0|1|"
            frmFor.CodigoActual = txtAux1(indice).Text
            frmFor.Show vbModal
            Set frmFor = Nothing
            PonerFoco txtAux1(indice)
            
    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub


Private Sub btnFec(Index As Integer)
Dim esq As Long
Dim dalt As Long
Dim menu As Long
Dim obj As Object

    Set frmC1 = New frmCal
    
    esq = btnBuscar(Index).Left
    dalt = btnBuscar(Index).Top
        
    Set obj = btnBuscar(Index).Container
      
    While btnBuscar(Index).Parent.Name <> obj.Name
           esq = esq + obj.Left
           dalt = dalt + obj.Top
           Set obj = obj.Container
    Wend
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    frmC1.Left = esq + btnBuscar(Index).Parent.Left + 30
    frmC1.Top = dalt + btnBuscar(Index).Parent.Top + btnBuscar(Index).Height + menu - 40

    btnBuscar(0).Tag = Index '<===
    Select Case Index
        Case 3
            indice = 4
        Case 4
            indice = 6
    End Select
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtAux(indice).Text <> "" Then frmC1.NovaData = txtAux(indice).Text
    ' ********************************************

    frmC1.Show vbModal
    Set frmC1 = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtAux(indice) '<===
    ' ********************************************
End Sub






Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BÚSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            text1(4).Text = Format(text1(1).Text, "yyyy-mm-dd") & " " & Format(text1(2).Text, "hh:mm:ss")
            text1(5).Text = ""
            If text1(8).Text <> "" Then
                text1(5).Text = Format(text1(8).Text, "yyyy-mm-dd") & " " & Format(text1(9).Text, "hh:mm:ss")
            End If
        
            
            If DatosOk Then InsertarCabecera
'            if datosok then
'                If InsertarDesdeForm2(Me, 1) Then
'
'                    Data1.RecordSource = "Select * from " & NombreTabla & Ordenacion
'                    PosicionarData
'                End If
'            Else
'                ModoLineas = 0
'            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                text1(4).Text = text1(1).Text & " " & Format(text1(2).Text, "hh:mm:ss")
                If text1(8).Text <> "" Or text1(9).Text <> "" Then
                    text1(5).Text = text1(8).Text & " " & Format(text1(9).Text, "hh:mm:ss")
                End If
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
                        PonerFoco txtAux(5)
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
            PonerModo 0
        Else
            If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                BotonAnyadir
            Else
                PonerModo 1 'búsqueda
                ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
                text1(0).BackColor = vbYellow 'codforfait
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
'        .Buttons(14).Image = 11  'Eixir
'        'el 13 i el 14 son separadors
'        .Buttons(btnPrimero).Image = 6  'Primer
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Següent
'        .Buttons(btnPrimero + 3).Image = 9 'Últim
    End With
    
    ' ******* si n'hi han llínies *******
    'ICONETS DE LES BARRES ALS TABS DE LLÍNIA
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
    For I = 3 To 3
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
   
    Me.imgDoc(1).Picture = frmPpal.imgListPpal.ListImages(16).Picture
   
   
   
    'carga IMAGES de mail
'    For i = 0 To Me.imgMail.Count - 1
'        Me.imgMail(i).Picture = frmPpal.imgListImages16.ListImages(2).Picture
'    Next i
    
    'IMAGES para zoom
    For I = 0 To Me.imgZoom.Count - 1
        Me.imgZoom(I).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    Next I
    
    CodTipoMov = "ORD" 'hcoCodTipoM
    
    LimpiarCampos   'Neteja els camps TextBox
    ' ******* si n'hi han llínies *******
    DataGridAux(0).ClearFields
    DataGridAux(1).ClearFields
    DataGridAux(2).ClearFields
    DataGridAux(3).ClearFields
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "cccaborden"
    Ordenacion = " ORDER BY codorden"
    
    'Mirem com està guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    Data1.ConnectionString = conn
    '***** cambiar el nombre de la PK de la cabecera *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where codorden is null"
    Data1.Refresh
       
    CargaGrid 0, False
    CargaGrid 1, False
    CargaGrid 2, False
    CargaGrid 3, False
       
    ModoLineas = 0
       
    ' *** si n'hi han combos (capçalera o llínies) ***
    
'    If DatosADevolverBusqueda = "" Then
'        PonerModo 0
'    Else
'        PonerModo 1 'búsqueda
'        ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
'        Text1(0).BackColor = vbYellow 'codforfait
'        ' ****************************************************************************
'    End If
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
Dim I As Integer, NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo
 
    'Actualisa Iconos Insertar,Modificar,Eliminar
    'ActualizarToolbar Modo, Kmodo
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo, ModoLineas
       
    'Modo 2. N'hi han datos i estam visualisant-los
    '=========================================
    'Posem visible, si es formulari de búsqueda, el botó "Regresar" quan n'hi han datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = (Modo = 2)
    Else
        cmdRegresar.visible = False
    End If
    
    text1(5).Enabled = True
    
    
    '=======================================
    b = (Modo = 2)
    'Posar Fleches de desplasament visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Només es per a saber que n'hi ha + d'1 registre
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    DesplazamientoVisible b And Data1.Recordset.RecordCount > 1
    '---------------------------------------------
    
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
       
    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
    'Si estem en Insertar a més neteja els camps Text1
    BloquearText1 Me, Modo
    
    b = (Modo = 3 Or Modo = 4 Or Modo = 1) '06/09/2005, lleve el modo 5 per a que no es puga modificar la capçalera mentre treballe en les llínies
        
    For I = 1 To 9
        If I = 1 Or I = 2 Or I = 8 Or I = 9 Or I = 4 Or I = 5 Then
            text1(I).Locked = Not b  '((Not b) And (Modo <> 1))
            If b Then
                 text1(I).BackColor = vbWhite
            Else
                 text1(I).BackColor = &H80000018 'groc
            End If
            If Modo = 3 Then text1(I).Text = "" 'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)
        End If
    Next I
    
    b = (Modo <> 1)
    'Campos Nº Pedido bloqueado y en azul
    BloquearTxt text1(0), b, True
    
    b = (Modo = 3 Or Modo = 4 Or Modo = 2)
    Me.imgDoc(1).visible = b
    Me.imgDoc(1).Enabled = b
    
    
    '*** si n'hi han combos a la capçalera ***
    '**************************
    
    ' ***** bloquejar tots els controls visibles de la clau primaria de la capçalera ***
'    If Modo = 4 Then _
'        BloquearTxt Text1(0), True 'si estic en  modificar, bloqueja la clau primaria
    ' **********************************************************************************
    
    For I = 2 To 3
        BloquearTxt txtAux(I), True
        txtAux(I).Enabled = False
    Next I
    For I = 2 To 3
        BloquearTxt txtAux(I), (Modo <> 1)
        txtAux(I).Enabled = (Modo = 1)
    Next I
    
    
    
    ' **** si n'hi han imagens de buscar en la capçalera *****
    BloquearImgBuscar Me, Modo, ModoLineas
    BloquearImgZoom Me, Modo, ModoLineas
    ' ********************************************************
    BloquearImgFec Me, 0, Modo
    BloquearImgFec Me, 1, Modo
        
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    PonerLongCampos

    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 0, False
        CargaGrid 1, False
        CargaGrid 2, False
        CargaGrid 3, False
    End If
    
    b = (Modo = 4) Or (Modo = 2)
    DataGridAux(0).Enabled = b
    DataGridAux(1).Enabled = b
    DataGridAux(2).Enabled = b
    DataGridAux(3).Enabled = b
      
    ' ****** si n'hi han combos a la capçalera ***********************
    ' ****************************************************************
    
    PonerModoOpcionesMenu (Modo) 'Activar opcions menú según modo
    PonerOpcionesMenu   'Activar opcions de menú según nivell
                        'de permisos de l'usuari

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
Dim I As Byte
    
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
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0) And Not DeConsulta
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
    For I = 0 To ToolAux.Count - 1
        ToolAux(I).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.AdoAux(I).Recordset.RecordCount > 0)
        ToolAux(I).Buttons(2).Enabled = bAux
        ToolAux(I).Buttons(3).Enabled = bAux
    Next I
    
End Sub

Private Sub Desplazamiento(Index As Integer)
'Botons de Desplaçament; per a desplaçar-se pels registres de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index, True
    PonerCampos
End Sub

Private Function MontaSQLCarga(Index As Integer, enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basant-se en la informació proporcionada pel vector de camps
'   crea un SQl per a executar una consulta sobre la base de datos que els
'   torne.
' Si ENLAZA -> Enlaça en el data1
'           -> Si no el carreguem sense enllaçar a cap camp
'--------------------------------------------------------------------
Dim Sql As String
Dim tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
               
        Case 0 'TRABAJADORES
            Sql = "SELECT cclinorden1.codorden, cclinorden1.numlinea, cclinorden1.codcoste, ccconcostes.nomcoste,  "
            Sql = Sql & " cclinorden1.codtraba, straba.nomtraba, date(cclinorden1.fechaini), time(cclinorden1.fechaini) horaini, "
            Sql = Sql & " date(cclinorden1.fechafin), time(cclinorden1.fechafin) horafin, cclinorden1.horas, cclinorden1.fechaini, cclinorden1.fechafin "
            Sql = Sql & " FROM (cclinorden1 INNER JOIN ccconcostes On cclinorden1.codcoste = ccconcostes.codcoste)  "
            Sql = Sql & " INNER JOIN straba ON cclinorden1.codtraba = straba.codtraba "
            
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE cclinorden1.codorden is null "
            End If
            Sql = Sql & " ORDER BY cclinorden1.numlinea"
               
        Case 1 'CATEGORIAS
            Sql = "SELECT cclinorden2.codorden, cclinorden2.numlinea, cclinorden2.codcoste, ccconcostes.nomcoste,  "
            Sql = Sql & " cclinorden2.codcateg, salarios.nomcateg, cclinorden2.horas "
            Sql = Sql & " FROM (cclinorden2 INNER JOIN ccconcostes ON cclinorden2.codcoste = ccconcostes.codcoste) "
            Sql = Sql & " INNER JOIN salarios ON cclinorden2.codcateg = salarios.codcateg "
            
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE cclinorden2.codorden is null "
            End If
            Sql = Sql & " ORDER BY cclinorden2.numlinea"
            
        Case 2 'VARIEDADES/FORFAITS
            Sql = "SELECT cclinorden3.codorden, cclinorden3.numlinea, cclinorden3.codvarie, variedades.nomvarie,  "
            Sql = Sql & " cclinorden3.codforfait, forfaits.nomconfe, cclinorden3.kilosnet, cclinorden3.numcajon "
            Sql = Sql & " FROM (cclinorden3 INNER JOIN variedades On cclinorden3.codvarie = variedades.codvarie)  "
            Sql = Sql & " INNER JOIN forfaits ON cclinorden3.codforfait = forfaits.codforfait "
            
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE cclinorden3.codorden is null "
            End If
            Sql = Sql & " ORDER BY cclinorden3.numlinea"
            
        Case 3 'RESUMEN POR FORFAITS
            Sql = "SELECT cclinorden4.codorden, cclinorden4.numlinea, cclinorden4.codforfait, forfaits.nomconfe,  "
            Sql = Sql & " cclinorden4.kilosnet, cclinorden4.numcajon "
            Sql = Sql & " FROM cclinorden4 INNER JOIN forfaits ON cclinorden4.codforfait = forfaits.codforfait "
            
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE cclinorden4.codorden is null "
            End If
            Sql = Sql & " ORDER BY cclinorden4.numlinea"
    
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
        Aux = ValorDevueltoFormGrid(text1(0), CadenaDevuelta, 1)
        CadB = Aux
        '   Com la clau principal es única, en posar el sql apuntant
        '   al valor retornat sobre la clau ppal es suficient
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        ' **********************************
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    text1(indice).Text = Format(vFecha, "dd/mm/yyyy") '<===
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

Private Sub frmCos_DatoSeleccionado(CadenaSeleccion As String)
'Concepto de coste
    txtAux1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo
    txtAux2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'denominacion
End Sub

Private Sub frmFor_DatoSeleccionado(CadenaSeleccion As String)
'Forfaits
    txtAux1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codforfait
    txtAux2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'denominacion
End Sub


Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Variedades
    txtAux1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000") 'codvariedad
    txtAux2(1).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     text1(indice).Text = vCampo
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



Private Sub imgDoc_Click(Index As Integer)
    
    Select Case Index
        Case 1 ' Mostramos los resumenes por categoria y por variedad
            FrameResumen.visible = Not FrameResumen.visible
            FrameAux0.visible = Not FrameResumen.visible
            FrameAux2.visible = Not FrameResumen.visible
            
            If FrameResumen.visible = True Then
                Me.imgDoc(1).Picture = frmPpal.imgListPpal.ListImages(26).Picture
                Me.imgDoc(1).ToolTipText = "Volver de Resumen"
            Else
                Me.imgDoc(1).Picture = frmPpal.imgListPpal.ListImages(16).Picture
                Me.imgDoc(1).ToolTipText = "Mostrar Resumen"
            End If

    End Select
    
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
    If text1(indice).Text <> "" Then frmC.NovaData = text1(indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco text1(indice) '<===
    ' ********************************************

End Sub

Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        indice = 6
        frmZ.pTitulo = "Observaciones de la Orden de Confección"
        frmZ.pValor = text1(indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco text1(indice)
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
        Case 5  'Búscar
           mnBuscar_Click
        Case 6  'Tots
            mnVerTodos_Click
        Case 1  'Nou
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
        Case 8 'Imprimir
            mnImprimir_Click
    End Select
End Sub

Private Sub BotonBuscar()
Dim I As Integer
Dim anc As Single

' ***** Si la clau primaria de la capçalera no es Text1(0), canviar-ho en <=== *****
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        
        'poner los txtaux para buscar por lineas
        anc = DataGridAux(0).Top
        If DataGridAux(0).Row < 0 Then
            anc = anc + 210
        Else
            anc = anc + DataGridAux(0).RowTop(DataGridAux(0).Row) + 5
        End If
        
        LLamaLineas 0, 1, anc
        
'        anc = DataGridAux(1).Top
'        If DataGridAux(1).Row < 0 Then
'            anc = anc + 210
'        Else
'            anc = anc + DataGridAux(1).RowTop(DataGridAux(1).Row) + 5
'        End If
'
'        LLamaLineas 1, 1, anc
        
        
        PonerFoco text1(0) ' <===
        text1(0).BackColor = vbYellow ' <===
        ' *** si n'hi han combos a la capçalera ***
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            text1(kCampo).Text = ""
            text1(kCampo).BackColor = vbYellow
            PonerFoco text1(kCampo)
        End If
    End If
' ******************************************************************************
End Sub

Private Sub HacerBusqueda()
    
    
    text1(4).ToolTipText = text1(4).Tag
    text1(5).ToolTipText = text1(5).Tag
    
    If text1(1).Text <> "" Then
        If text1(2).Text = "" Then
            text1(4).Text = text1(1).Text
            text1(4).Tag = Replace(text1(4).Tag, "FH", "FHF")
'        Else
'            text1(4).Text = Format(text1(1).Text, "yyyy-mm-dd") & " " & text1(2).Text
        End If
    Else
        If text1(2).Text <> "" Then
            text1(4).Text = text1(2).Text
            text1(4).Tag = Replace(text1(4).Tag, "FH", "FHH")
        End If
    End If

    If text1(8).Text <> "" Then
        If text1(9).Text = "" Then
            text1(5).Text = text1(8).Text
            text1(5).Tag = Replace(text1(5).Tag, "FH", "FHF")
'        Else
'            text1(5).Text = Format(text1(8).Text, "yyyy-mm-dd") & " " & text1(9).Text
        End If
    Else
        If text1(9).Text <> "" Then
            text1(5).Text = text1(9).Text
            text1(5).Tag = Replace(text1(5).Tag, "FH", "FHH")
        End If
    End If

    CadB = ObtenerBusqueda2(Me, 1)
    
    text1(4).Tag = text1(4).ToolTipText
    text1(5).Tag = text1(5).ToolTipText
    
    
    If chkVistaPrevia(0) = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
'        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        CadenaConsulta = "select " & NombreTabla & ".* from " & NombreTabla & " LEFT JOIN cclinorden1 ON cccaborden.codorden=cclinorden1.codorden "
        CadenaConsulta = CadenaConsulta & " WHERE " & CadB & " GROUP BY cccaborden.codorden " & Ordenacion
        
        PonerCadenaBusqueda
    Else
        ' *** foco al 1r camp visible de la capçalera que siga clau primaria ***
        PonerFoco text1(0)
        ' **********************************************************************
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
    Dim Cad As String
        
    'Cridem al form
    ' **************** arreglar-ho per a vore lo que es desije ****************
    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
    Cad = ""
    Cad = Cad & ParaGrid(text1(0), 10, "Codigo")
    Cad = Cad & "F.Inicio|fechaini|T|dd/mm/yyyy|15·"
    Cad = Cad & "H.Inicio|fechaini|T|hh:mm:ss|10·"
    Cad = Cad & "F.Fin|fechafin|T|dd/mm/yyyy|15·"
    Cad = Cad & "H.Fin|fechafin|T|hh:mm:ss|10·"
    Cad = Cad & "Codigo|cccaborden.codvarie|T|000000|10·"
    
'    cad = cad & ParaGrid(text1(2), 60, "Descripción")
    Cad = Cad & "Variedad|nomvarie|T||28·"
    If Cad <> "" Then
        
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        Cad = NombreTabla & " left join variedades on cccaborden.codvarie = variedades.codvarie "
        frmB.vtabla = Cad 'NombreTabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        frmB.vDevuelve = "0|" '*** els camps que volen que torne ***
        frmB.vTitulo = "Ordenes Confeccion" ' ***** repasa açò: títol de BuscaGrid *****
        frmB.vSelElem = 0

        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha posat valors i tenim que es formulari de búsqueda llavors
        'tindrem que tancar el form llançant l'event
        If HaDevueltoDatos Then
            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                cmdRegresar_Click
        Else   'de ha retornat datos, es a decir NO ha retornat datos
            PonerFoco text1(kCampo)
        End If
    End If
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String
Dim Aux As String
Dim I As Integer
Dim J As Integer

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    Cad = ""
    I = 0
    Do
        J = I + 1
        I = InStr(J, DatosADevolverBusqueda, "|")
        If I > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, I - J)
            J = Val(Aux)
            Cad = Cad & text1(J).Text & "|"
        End If
    Loop Until I = 0
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub

Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        PonerModo 2
        
        LLamaLineas 0, 0, 0
        
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
    
    If chkVistaPrevia(0).Value = 1 Then
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
    ' codEmpre i quins camps tenen la PK de la capçalera *******
'    text1(0).Text = SugerirCodigoSiguienteStr("forfaits", "codforfait")
'    FormateaCampo text1(0)
    '******************** canviar taula i camp **************************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        NumF = NuevoCodigo
    Else
        NumF = ""
    End If
    '********************************************************************

'    text1(0) = NumF
    PonerFoco text1(1) '*** 1r camp visible que siga PK ***
    
    ' *** si n'hi han camps de descripció a la capçalera ***
    'PosarDescripcions

End Sub

Private Sub BotonModificar()

    PonerModo 4

    ' *** bloquejar els camps visibles de la clau primaria de la capçalera ***
    BloquearTxt text1(0), True
    
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco text1(1)
End Sub

Private Sub BotonEliminar()
Dim Cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
    ' ***************************************************************************

    ' *************** canviar la pregunta ****************
    Cad = "¿Seguro que desea eliminar la Orden?"
    Cad = Cad & vbCrLf & "Código: " & Data1.Recordset.Fields(0)
    Cad = Cad & vbCrLf & "Fecha: " & Data1.Recordset.Fields(1)
    
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
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
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Proveedor", Err.Description
End Sub

Private Sub PonerCampos()
Dim I As Integer
Dim codpobla As String, despobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la capçalera
    
    
    text1(1).Text = Format(text1(4).Text, "dd/mm/yyyy")
    text1(2).Text = Format(text1(4).Text, "hh:mm:ss")
    
    If text1(5).Text <> "" Then
        text1(8).Text = Format(text1(5).Text, "dd/mm/yyyy")
        text1(9).Text = Format(text1(5).Text, "hh:mm:ss")
    End If
    
    ' *** si n'hi han llínies en datagrids ***
    'For i = 0 To DataGridAux.Count - 1
    For I = 0 To 3
        CargaGrid I, True
        If Not AdoAux(I).Recordset.EOF Then _
            PonerCamposForma2 Me, AdoAux(I), 2, "FrameAux" & I
    Next I

    
    ' ************* configurar els camps de les descripcions de la capçalera *************
    text2(3).Text = PonerNombreDeCod(text1(3), "clientes", "nomclien")
    ' ********************************************************************************
    
    CalcularTotales
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu
End Sub

Private Sub cmdCancelar_Click()
Dim I As Integer
Dim V

    Select Case Modo
        Case 1, 3 'Búsqueda, Insertar
                LimpiarCampos
                If Data1.Recordset.EOF Then
                    PonerModo 0
                Else
                    PonerModo 2
                    PonerCampos
                End If
                LLamaLineas 0, 0, 0
                ' *** foco al primer camp visible de la capçalera ***
                PonerFoco text1(0)

        Case 4  'Modificar
                TerminaBloquear
                PonerModo 2
                PonerCampos
                ' *** primer camp visible de la capçalera ***
                PonerFoco text1(0)
        
        Case 5 'LLÍNIES
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

                        ' *** si n'hi han camps de descripció dins del grid, els neteje ***
                        'txtAux2(2).text = ""

                    End If
                    
'                    ' *** si n'hi han tabs ***
'                    SituarTab (NumTabMto + 1)
                    
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
            
            PosicionarData
            
            ' *** si n'hi han llínies en grids i camps fora d'estos ***
            If Not AdoAux(NumTabMto).Recordset.EOF Then
                DataGridAux_RowColChange NumTabMto, 1, 1
            Else
                LimpiarCamposFrame NumTabMto
            End If
    End Select
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm2(Me, 1)
    If Not b Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
    If (Modo = 3) Then 'insertar
        'comprobar si existe ya el cod. del campo clave primaria
        If ExisteCP(text1(0)) Then b = False
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
    Cad = "(codorden=" & DBSet(text1(0).Text, "N") & ")"
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(Data1, cad, Indicador) Then
    If SituarData(Data1, Cad, Indicador) Then
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
    vWhere = " WHERE codorden=" & Data1.Recordset!CodOrden
        
    ' ***** elimina les llínies ****
    conn.Execute "DELETE FROM cclinorden1 " & vWhere
        
    conn.Execute "DELETE FROM cclinorden2 " & vWhere
        
    conn.Execute "DELETE FROM cclinorden3 " & vWhere
        
    conn.Execute "DELETE FROM cclinorden4 " & vWhere
    
    conn.Execute "DELETE FROM cclinorden5 " & vWhere
    
    'Eliminar la CAPÇALERA
    conn.Execute "Delete from " & NombreTabla & vWhere
       
    'Decrementar contador si borramos la ultima orden de confeccion
    Set vTipoMov = New CTiposMov
    vTipoMov.DevolverContador CodTipoMov, Val(text1(0).Text) ' "ORD", Val(Text1(0).Text)
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

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco text1(Index), Modo
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean

    If Not PerderFocoGnral(text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    ' ***************** configurar els LostFocus dels camps de la capçalera *****************
    Select Case Index
        Case 0 'codigo de orden
            text1(Index).Text = UCase(text1(Index).Text)
        
        Case 3 'cliente
            If PonerFormatoEntero(text1(Index)) Then
                text2(Index).Text = PonerNombreDeCod(text1(Index), "clientes", "nomclien")
                If text2(Index).Text = "" Then
                    cadMen = "No existe el Cliente: " & text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCli = New frmClientes
                        frmCli.DatosADevolverBusqueda = "0|1|"
                        text1(Index).Text = ""
                        TerminaBloquear
                        frmCli.Show vbModal
                        Set frmCli = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        text1(Index).Text = ""
                    End If
                    PonerFoco text1(Index)
                End If
            Else
                text2(Index).Text = ""
            End If
            
        Case 4 'Forfait
            
        Case 1, 8 'fecha inicio y fin
            PonerFormatoFecha text1(Index)
            
        Case 2, 9 ' hora inicio y fin
            PonerFormatoHora text1(Index)
            
    End Select
        ' ***************************************************************************
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 3: KEYBusqueda KeyAscii, 3 'variedad
                Case 4: KEYBusqueda KeyAscii, 2 'confeccion
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
    imgBuscar_Click (indice)
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

    ModoLineas = 3 'Posem Modo Eliminar Llínia
    
    If Modo = 4 Then 'Modificar Capçalera
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
        Case 0 'trabajadores
            Sql = "¿ Seguro que desea eliminar el trabajador ?"
            Sql = Sql & vbCrLf & "Trabajador: " & AdoAux(Index).Recordset!codtraba & " " & AdoAux(Index).Recordset!NomTraba
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM cclinorden1 "
                Sql = Sql & vWhere & " AND numlinea= " & AdoAux(Index).Recordset!NumLinea
            End If
            
        Case 2 'variedades
            Sql = "¿ Seguro que desea eliminar la variedad ?"
            Sql = Sql & vbCrLf & "Nombre: " & AdoAux(Index).Recordset!nomvarie
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM cclinorden3 "
                Sql = Sql & vWhere & " AND numlinea= " & AdoAux(Index).Recordset!NumLinea
            End If
            
    End Select

    If Eliminar Then
        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        conn.Execute Sql
        
        ' solo en el caso de que estemos en trabajadores y añadamos una nueva linea hemos de modificar las lineas de categoria
        If Index = 0 Then
            ModificarCategorias (True)
            CargaGrid 1, True
        End If
        
        If Index = 2 Then
            ModificarForfaits (True)
            CargaGrid 3, True
        End If
        
        
        
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        CargaGrid Index, True
        If Not SituarDataTrasEliminar(AdoAux(Index), NumRegElim, True) Then
'            PonerCampos
            
        End If
        CalcularTotales
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
Dim vWhere As String, vtabla As String
Dim anc As Single
Dim I As Integer
    
    ModoLineas = 1 'Posem Modo Afegir Llínia
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index
    
    ' *** bloquejar la clau primaria de la capçalera ***
    BloquearTxt text1(0), True

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case Index
        Case 0: vtabla = "cclinorden1"
        Case 2: vtabla = "cclinorden3"
    End Select
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 0, 2 ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            
            NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", vWhere)

            AnyadirLinea DataGridAux(Index), AdoAux(Index)
    
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 240
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
            
            LLamaLineas Index, ModoLineas, anc
        
            Select Case Index
                ' *** valor per defecte a l'insertar i formateig de tots els camps ***
                Case 0 'trabajadores
                    txtAux(0).Text = text1(0).Text 'codorden
                    txtAux(1).Text = NumF 'numlinea
                    For I = 2 To 8
                        txtAux(I).Text = ""
                    Next I
                    txtAux2(2).Text = ""
                    txtAux2(3).Text = ""
                    
                    txtAux(4).Text = text1(1).Text
                    txtAux(6).Text = text1(1).Text
                    
                    BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux0"
                    btnBuscar(4).visible = False
                    btnBuscar(4).Enabled = False
                    
                    PonerFoco txtAux(2)
                    
                
                Case 2 'variedades
                    txtAux1(0).Text = text1(0).Text 'codorden
                    txtAux1(1).Text = NumF 'numlinea
                    txtAux1(2).Text = ""
                    txtAux2(1).Text = ""
                    txtAux1(3).Text = ""
                    txtAux2(0).Text = ""
                    txtAux1(4).Text = ""
                    txtAux1(5).Text = ""
                    
                    BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux2"
                    PonerFoco txtAux1(2)
            End Select
            
    End Select
End Sub

Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim I As Integer
    Dim J As Integer
    
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
    BloquearTxt text1(0), True
  
    Select Case Index
        Case 0, 2 ' *** pose els index de llínies que tenen datagrid (en o sense tab) ***
            If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
                I = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
                DataGridAux(Index).Scroll 0, I
                DataGridAux(Index).Refresh
            End If
              
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 240
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
            txtAux2(2).Text = DataGridAux(Index).Columns(3).Text
            txtAux(3).Text = DataGridAux(Index).Columns(4).Text
            txtAux2(3).Text = DataGridAux(Index).Columns(5).Text
            txtAux(4).Text = DataGridAux(Index).Columns(6).Text
            txtAux(5).Text = DataGridAux(Index).Columns(7).Text
            txtAux(6).Text = DataGridAux(Index).Columns(8).Text
            txtAux(7).Text = DataGridAux(Index).Columns(9).Text
            txtAux(8).Text = DataGridAux(Index).Columns(10).Text
            txtAux(9).Text = DataGridAux(Index).Columns(11).Text
            txtAux(10).Text = DataGridAux(Index).Columns(12).Text
            
            BloquearbtnBuscar Me, Modo, 1, "FrameAux0"
            btnBuscar(4).visible = False
            btnBuscar(4).Enabled = False
            
        Case 2 'variedades
            For J = 0 To 2
                txtAux1(J).Text = DataGridAux(Index).Columns(J).Text
            Next J
            txtAux2(2).Text = DataGridAux(Index).Columns(3).Text
            txtAux1(3).Text = DataGridAux(Index).Columns(4).Text
            txtAux2(0).Text = DataGridAux(Index).Columns(5).Text
            txtAux1(4).Text = DataGridAux(Index).Columns(6).Text
            txtAux1(5).Text = DataGridAux(Index).Columns(7).Text
            
            BloquearbtnBuscar Me, Modo, 1, "FrameAux2"
            
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 0 'trabajadores
            PonerFoco txtAux(2)
        Case 2 'variedades
            PonerFoco txtAux1(2)
    End Select
    ' ***************************************************************************************
End Sub

Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim b As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    DeseleccionaGrid DataGridAux(Index)
       
    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
    Select Case Index
        Case 0 'trabajadores
             For jj = 2 To 8
                If jj <> 6 Then
                    txtAux(jj).visible = b
                    txtAux(jj).Top = alto
                End If
            Next jj
            For jj = 2 To 3
                txtAux2(jj).visible = b
                txtAux2(jj).Top = alto
            Next jj
            btnBuscar(0).visible = b
            btnBuscar(0).Top = alto
            For jj = 2 To 3
                btnBuscar(jj).visible = b
                btnBuscar(jj).Top = alto
            Next jj
            
        Case 2 'variedades
            For jj = 2 To 5
                txtAux1(jj).visible = b
                txtAux1(jj).Top = alto
            Next jj
            For jj = 0 To 1
                txtAux2(jj).visible = b
                txtAux2(jj).Top = alto
            Next jj
            
            btnBuscar(6).visible = b
            btnBuscar(6).Top = alto
            btnBuscar(7).visible = b
            btnBuscar(7).Top = alto
            
    End Select
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
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
        Case 2 ' codigo de coste dentro de trabajadores
            If txtAux(Index) <> "" Then
                txtAux2(2).Text = PonerNombreDeCod(txtAux(Index), "ccconcostes", "nomcoste")
                If txtAux2(2).Text = "" Then
                    cadMen = "No existe el concepto de coste: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCos = New frmCCManConcep
                        
                        indice = Index
                        txtAux(Index).Text = ""
                        
                        frmCos.DatosADevolverBusqueda = "0|1|"
                        frmCos.NuevoCodigo = txtAux(Index).Text

                        TerminaBloquear
                        frmCos.Show vbModal
                        
                        Set frmCos = Nothing
        
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                Else
                    If Not EsConceptoDirecto(txtAux(Index).Text) Then
                        MsgBox "El Concepto de coste no es directo. Reintroduzca.", vbExclamation
                        PonerFoco txtAux(Index)
                    End If
                End If
            Else
                txtAux2(2).Text = ""
            End If
        
        Case 42 ' zona
        
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
        
        Case 4, 6 ' fecha inicio y fecha fin
            PonerFormatoFecha txtAux(Index)
            
        Case 5, 7 ' hora inicio y hora fin
            PonerFormatoHora txtAux(Index)
        
        Case 8, 44 ' horas
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
            
        
        Case 2, 10 ' cantidad e importes
            If txtAux(Index).Text <> "" Then
                If PonerFormatoDecimal(txtAux(Index), 7) Then
                    CalcularPrecio
                    cmdAceptar.SetFocus
                End If
            End If
    End Select
    
    If Index = 4 Or Index = 5 Then
        txtAux(9).Text = Trim(Format(txtAux(4).Text, "yyyy-mm-dd") & " " & Format(txtAux(5).Text, "hh:mm:ss"))
    End If
    
    If Index = 6 Or Index = 7 Then
        txtAux(10).Text = Trim(Format(txtAux(6).Text, "yyyy-mm-dd") & " " & Format(txtAux(7).Text, "hh:mm:ss"))
    End If
    
    If txtAux(4).Text <> "" And txtAux(5).Text <> "" And txtAux(6).Text <> "" And txtAux(7).Text <> "" Then
        txtAux(8).Text = Round2(DateDiff("n", CDate(Format(txtAux(9).Text, "dd/mm/yyyy") & " " & Format(txtAux(9).Text, "hh:mm:ss")), CDate(Format(txtAux(10).Text, "dd/mm/yyyy") & " " & Format(txtAux(10).Text, "hh:mm:ss"))) / 60, 2)
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
            Sql = "select count(*) from cclinorden2 where codorden= " & DBSet(text1(0).Text, "N")
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

Private Sub imgBuscar_Click(Index As Integer)
    TerminaBloquear
    
     Select Case Index
        Case 2 'forfaits
            indice = 4
            
            Set frmFor = New frmManForfaits
            frmFor.DatosADevolverBusqueda = "0|1|"
            frmFor.CodigoActual = text1(4).Text
            frmFor.Show vbModal
            Set frmFor = Nothing
            PonerFoco text1(4)
            
        Case 3 'Variedades
            indice = 3
            
            Set frmVar = New frmManVariedad
            frmVar.DatosADevolverBusqueda = "0|1|"
            frmVar.CodigoActual = text1(3).Text
            frmVar.Show vbModal
            Set frmVar = Nothing
            PonerFoco text1(3)
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub

Private Sub DataGridAux_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
Dim I As Byte

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
                For I = 21 To 24
'                   txtAux(i).Text = ""
                Next I
'               txtAux2(22).Text = ""
            Case 2 'Tarjetas
'               txtAux(50).Text = ""
'               txtAux(51).Text = ""
        End Select
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

Private Sub CargaGrid(Index As Integer, enlaza As Boolean)
Dim b As Boolean
Dim I As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)

    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
    
    Select Case Index
        Case 0 'trabajadores
            txtAux(5).Tag = "Fecha Ini|FHH|N|||cclinorden1|fechaini|hh:mm:ss||"
            txtAux(7).Tag = "Fecha Ini|FHH|N|||cclinorden1|fechaini|hh:mm:ss||"
        
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;" 'codorden, numlinea
            tots = tots & "S|txtAux(2)|T|Coste|600|;S|btnBuscar(0)|B|||;"
            tots = tots & "S|txtAux2(2)|T|Denominación|2500|;S|txtAux(3)|T|Código|800|;S|btnBuscar(2)|B|||;"
            tots = tots & "S|txtAux2(3)|T|Trabajador|3200|;S|txtAux(4)|T|Fecha|1100|;S|btnBuscar(3)|B|||;"
            tots = tots & "S|txtAux(5)|T|Hora In.|800|;N||||0|;N|btnBuscar(4)|B|||;S|txtAux(7)|T|Hora F.|800|;"
            tots = tots & "S|txtAux(8)|T|Horas|800|;N||||0|;N||||0|;"
            
            arregla tots, DataGridAux(Index), Me
        
            DataGridAux(0).Columns(6).NumberFormat = "dd/mm/yyyy"
            DataGridAux(0).Columns(8).NumberFormat = "dd/mm/yyyy"
            
            DataGridAux(0).Columns(2).Alignment = dbgLeft
            DataGridAux(0).Columns(3).Alignment = dbgLeft
            DataGridAux(0).Columns(4).Alignment = dbgLeft
            DataGridAux(0).Columns(5).Alignment = dbgLeft

            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
            txtAux(5).Tag = ""
            txtAux(7).Tag = ""
            
        Case 1 'categorias
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;" 'codorden,numlinea
            tots = tots & "S|txtAux(42)|T|Coste|600|;S|btnBuscar(1)|B||195|;"
            tots = tots & "S|txtAux2(42)|T|Denominación|3050|;"
            tots = tots & "S|txtAux(43)|T|Cat.|500|;S|btnBuscar(5)|B||195|;"
            tots = tots & "S|txtAux2(43)|T|Denominación|2950|;"
            tots = tots & "S|txtAux(44)|T|Horas|1200|;"
            
            arregla tots, DataGridAux(Index), Me
            
            DataGridAux(1).Columns(2).Alignment = dbgLeft
            DataGridAux(1).Columns(4).Alignment = dbgLeft
            
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
            
        Case 2 'variedades/forfaits
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;" 'codorden, numlinea
            tots = tots & "S|txtAux1(2)|T|Codigo|1000|;S|btnBuscar(6)|B|||;"
            tots = tots & "S|txtAux2(1)|T|Variedad|2600|;S|txtAux1(3)|T|Forfait|1500|;S|btnBuscar(7)|B|||;"
            tots = tots & "S|txtAux2(0)|T|Descripcion|3200|;S|txtAux1(4)|T|Kilos|1140|;"
            tots = tots & "S|txtAux1(5)|T|Cajas|1140|;"
            
            arregla tots, DataGridAux(Index), Me
        
            
            DataGridAux(2).Columns(2).Alignment = dbgLeft
            DataGridAux(2).Columns(3).Alignment = dbgLeft

            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
        Case 3 'resumen por forfaits
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;" 'codorden, numlinea
            tots = tots & "S|txtAux1(42)|T|Codigo|1500|;"
            tots = tots & "S|txtAux2(4)|T|Forfait|4400|;"
            tots = tots & "S|txtAux1(43)|T|Kilos|1200|;"
            tots = tots & "S|txtAux1(44)|T|Cajas|1200|;"
            
            arregla tots, DataGridAux(Index), Me
            
            DataGridAux(3).Columns(2).Alignment = dbgLeft
            DataGridAux(3).Columns(3).Alignment = dbgLeft
            
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
    End Select
    
    DataGridAux(Index).ScrollBars = dbgAutomatic
      
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub

Private Function ModificarCategorias(Eliminar As Boolean) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Categoria As Integer
Dim NumF As String
Dim Horas As String
Dim Rs As ADODB.Recordset

    On Error GoTo eModificarCategorias


    ModificarCategorias = True

    Sql = "delete from cclinorden2 where codorden = " & text1(0).Text
    conn.Execute Sql

    Sql = "select * from cclinorden1 where codorden = " & text1(0).Text
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    While Not Rs.EOF
    
        Sql = "select codcateg from straba where codtraba = " & DBSet(Rs!codtraba, "N")
        Categoria = DevuelveValor(Sql)
    
        Sql = "select count(*) from cclinorden2 where codorden = " & DBSet(text1(0).Text, "N")
        Sql = Sql & " and codcoste = " & DBSet(Rs!codCoste, "N")
        Sql = Sql & " and codcateg = " & DBSet(Categoria, "N")
        
        If TotalRegistros(Sql) = 0 Then
            NumF = SugerirCodigoSiguienteStr("cclinorden2", "numlinea", "codorden = " & DBSet(text1(0).Text, "N"))
        
            Sql2 = "insert into cclinorden2 (codorden,numlinea,codcoste,codcateg,horas) values ("
            Sql2 = Sql2 & DBSet(text1(0).Text, "N") & "," & DBSet(NumF, "N") & "," & DBSet(Rs!codCoste, "N") & ","
            Sql2 = Sql2 & DBSet(Categoria, "N") & "," & DBSet(Rs!Horas, "N") & ")"
            
            conn.Execute Sql2
        
        Else
            Sql2 = "update cclinorden2 set horas = horas + " & DBSet(Rs!Horas, "N")
            Sql2 = Sql2 & " where codorden = " & DBSet(text1(0).Text, "N")
            Sql2 = Sql2 & " and codcoste = " & DBSet(Rs!codCoste, "N")
            Sql2 = Sql2 & " and codcateg = " & DBSet(Categoria, "N")
                
            conn.Execute Sql2
        End If
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    Exit Function
    
eModificarCategorias:
    ModificarCategorias = False
    
End Function


Private Function ModificarForfaits(Eliminar As Boolean) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Categoria As Integer
Dim NumF As String
Dim Horas As String
Dim Rs As ADODB.Recordset

    On Error GoTo eModificarForfaits

    ModificarForfaits = True
    
    Sql = "delete from cclinorden4 where codorden = " & text1(0).Text
    conn.Execute Sql
    
    Sql = "select codforfait, sum(kilosnet) kilosnet, sum(numcajon) numcajon  from cclinorden3 where codorden = " & text1(0).Text
    Sql = Sql & " group by 1 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    While Not Rs.EOF
    
        NumF = SugerirCodigoSiguienteStr("cclinorden4", "numlinea", "codorden = " & DBSet(text1(0).Text, "N"))
    
        Sql2 = "insert into cclinorden4 (codorden,numlinea,codforfait,kilosnet,numcajon) values ("
        Sql2 = Sql2 & DBSet(text1(0).Text, "N") & "," & DBSet(NumF, "N") & "," & DBSet(Rs!codforfait, "T") & ","
        Sql2 = Sql2 & DBSet(Rs!KilosNet, "N") & "," & DBSet(Rs!numcajon, "N") & ")"
        
        conn.Execute Sql2
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing

eModificarForfaits:
    ModificarForfaits = False
End Function


Private Sub InsertarLinea()
'Inserta registre en les taules de Llínies
Dim nomFrame As String
Dim b As Boolean

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomFrame = "FrameAux0" 'trabajadores
        Case 2: nomFrame = "FrameAux2" 'variedades
    End Select
    
    
    If DatosOkLlin(nomFrame) Then
        If NumTabMto = 0 Then
            txtAux(9).Text = Format(txtAux(4).Text, "yyyy-mm-dd") & " " & Format(txtAux(5).Text, "hh:mm:ss")
            txtAux(10).Text = Format(txtAux(6).Text, "yyyy-mm-dd") & " " & Format(txtAux(7).Text, "hh:mm:ss")
        End If
        
        TerminaBloquear
        
        
        If InsertarDesdeForm2(Me, 2, nomFrame) Then
            
            ' solo en el caso de que estemos en trabajadores y añadamos una nueva linea hemos de modificar las lineas de categoria
            If NumTabMto = 0 Then
                b = ModificarCategorias(False)
                CargaGrid 1, True
            End If
        
        
            ' solo en el caso de que estemos en variedades hemos cuando insertemos hemos de modificar el resumen por forfaits
            If NumTabMto = 2 Then
                b = ModificarForfaits(False)
                CargaGrid 3, True
            End If
        
            b = BLOQUEADesdeFormulario2(Me, Data1, 1)
            Select Case NumTabMto
                Case 0, 2 ' *** els index de les llinies en grid (en o sense tab) ***
                     CargaGrid NumTabMto, True
                    If b Then BotonAnyadirLinea NumTabMto
            End Select
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
        Case 2: nomFrame = "FrameAux2" 'variedades
    End Select
    ModificarLinea = False
    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomFrame) Then
            ' solo en el caso de que estemos en trabajadores y añadamos una nueva linea hemos de modificar las lineas de categoria
            If NumTabMto = 0 Then
                ModificarCategorias False
                CargaGrid 1, True
            End If
            
            If NumTabMto = 2 Then
                ModificarForfaits False
                CargaGrid 3, True
            End If
            
            ModoLineas = 0
            
            Select Case NumTabMto
                Case 0
                    V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
                Case 2
                    V = AdoAux(NumTabMto).Recordset.Fields(2) 'el 2 es el nº de llinia
            End Select
            CargaGrid NumTabMto, True
            
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
    vWhere = vWhere & " codorden=" & Me.Data1.Recordset!CodOrden
    
    ObtenerWhereCab = vWhere
End Function

'' *** neteja els camps dels tabs de grid que
''estan fora d'este, i els camps de descripció ***
Private Sub LimpiarCamposFrame(Index As Integer)
    On Error Resume Next
 
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

Private Sub CalcularTotales()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim TotalEnvases As String
Dim TotalCostes As String
Dim Valor As Currency

    On Error Resume Next

    ' Coste de trabajadores

    Sql = "select sum(total) from ("
    Sql = Sql & "select cclinorden1.codtraba, round(sum(cclinorden1.horas) * straba.prhoracoste,2) total "
    Sql = Sql & " from (cclinorden1 inner join straba on cclinorden1.codtraba = straba.codtraba) "
    Sql = Sql & " where cclinorden1.codorden = " & text1(0).Text
    Sql = Sql & " group by 1) aaaaa "
    
    Text3(0).Text = DevuelveValor(Sql)
    
    Valor = CCur(TransformaPuntosComas(Text3(0).Text))
    If Valor <> 0 Then
        Text3(0).Text = Format(Valor, "###,###,##0.00")
    Else
        Text3(0).Text = ""
    End If
    
    ' Horas de trabajadores
    
    Sql = "select sum(cclinorden1.horas) "
    Sql = Sql & " from cclinorden1 where codorden = " & text1(0).Text
    
    Text3(2).Text = DevuelveValor(Sql)
    
    Valor = CCur(TransformaPuntosComas(Text3(2).Text))
    If Valor <> 0 Then
        Text3(2).Text = Format(Valor, "###,###,##0.00")
    Else
        Text3(2).Text = ""
    End If
    
    
    ' Coste de categorias
    Sql = "select sum(importe) from ("
    Sql = Sql & "select cclinorden1.codtraba, if (round(sum(cclinorden1.horas) * straba.prhoracoste,2) is null,0,round(sum(cclinorden1.horas) * straba.prhoracoste,2)) importe "
    Sql = Sql & " from cclinorden1 inner join on straba cclinorden1.codtraba = straba.codtraba "
    Sql = Sql & " where cclinorden1.codorden = " & DBSet(text1(0).Text, "N")
    Sql = Sql & " and straba.codcateg in (select cclinorden2.codcateg from cclinorden2 where cclinorden2.codorden = " & DBSet(text1(0).Text, "N") & ")"
    Sql = Sql & " and cclinorden1.codorden = " & DBSet(text1(0).Text, "N") & " group by 1) aaaaa "
    
    Text3(1).Text = DevuelveValor(Sql)
    
    Valor = CCur(TransformaPuntosComas(Text3(1).Text))
    If Valor <> 0 Then
        Text3(1).Text = Format(Valor, "###,###,##0.00")
    Else
        Text3(1).Text = ""
    End If
    
    ' Horas de trabajadores
    
    Sql = "select sum(cclinorden2.horas) "
    Sql = Sql & " from cclinorden2 where codorden = " & text1(0).Text
    
    Text3(3).Text = DevuelveValor(Sql)
    
    Valor = CCur(TransformaPuntosComas(Text3(3).Text))
    If Valor <> 0 Then
        Text3(3).Text = Format(Valor, "###,###,##0.00")
    Else
        Text3(3).Text = ""
    End If
    
    
    ' Kilos totales
    Sql = "select sum(if (kilosnet is null, 0,kilosnet)), sum(if(numcajon is null,0,numcajon)) from cclinorden3 "
    Sql = Sql & " where cclinorden3.codorden = " & text1(0).Text
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Text3(4).Text = DBLet(Rs.Fields(0).Value, "N")
        Text3(5).Text = DBLet(Rs.Fields(1).Value, "N")
        
    Else
        Text3(4).Text = ""
        Text3(5).Text = ""
    End If
    
    ' Kilos totales de forfaits
    Sql = "select sum(if (kilosnet is null, 0,kilosnet)), sum(if(numcajon is null,0,numcajon)) from cclinorden4 "
    Sql = Sql & " where cclinorden4.codorden = " & text1(0).Text
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Text3(6).Text = DBLet(Rs.Fields(0).Value, "N")
        Text3(7).Text = DBLet(Rs.Fields(1).Value, "N")
        
    Else
        Text3(6).Text = ""
        Text3(7).Text = ""
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
        text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
        Sql = CadenaInsertarDesdeForm(Me)
        If Sql <> "" Then
            If InsertarOferta(Sql, vTipoMov) Then
                CadenaConsulta = "Select * from " & NombreTabla & " where codorden = " & DBSet(text1(0).Text, "N") & Ordenacion
                PonerCadenaBusqueda
                PonerModo 2
                'Ponerse en Modo Insertar Lineas
                BotonAnyadirLinea 0
            End If
        End If
        text1(0).Text = Format(text1(0).Text, "0000000")
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
        devuelve = DevuelveDesdeBDNew(cAgro, NombreTabla, "codorden", "codorden", text1(0).Text, "N")
        If devuelve <> "" Then
            'Ya existe el contador incrementarlo
            Existe = True
            vTipoMov.IncrementarContador (CodTipoMov)
            text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
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


Private Sub txtAux1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
    
    If Not PerderFocoGnral(txtAux1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    

    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 2 ' variedad
            If txtAux1(Index) <> "" Then
                txtAux2(1).Text = PonerNombreDeCod(txtAux1(Index), "variedades", "nomvarie")
                If txtAux2(1).Text = "" Then
                    cadMen = "No existe la variedad " & txtAux1(Index).Text & ". Reintroduzca." & vbCrLf
                    MsgBox cadMen, vbExclamation
                    PonerFoco txtAux1(Index)
                End If
            Else
                txtAux2(1).Text = ""
            End If
        
        Case 3 ' forfait
            If txtAux1(Index) <> "" Then
                txtAux2(0).Text = PonerNombreDeCod(txtAux1(Index), "forfaits", "nomconfe")
                If txtAux2(0).Text = "" Then
                    cadMen = "No existe la Confección " & txtAux1(Index).Text & ". Reintroduzca." & vbCrLf
                    MsgBox cadMen, vbExclamation
                    PonerFoco txtAux1(Index)
                End If
            Else
                txtAux2(0).Text = ""
            End If
        
        
        Case 4, 5 ' kilos, cajas
            PonerFormatoEntero txtAux1(Index)
            
    End Select
    
End Sub

Private Sub txtAux1_GotFocus(Index As Integer)
   If Not txtAux1(Index).MultiLine Then ConseguirFocoLin txtAux1(Index)
End Sub


Private Sub TxtAux1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not txtAux1(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub txtAux1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not txtAux1(Index).MultiLine Then
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


Private Function EsConceptoDirecto(Concepto As String) As Boolean
Dim Sql As String

    Sql = "select tipocoste from ccconcostes where codcoste = " & DBSet(Concepto, "N")
    EsConceptoDirecto = (DevuelveValor(Sql) = 0)

End Function



Private Sub CalcularTotalesOld()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim TotalEnvases As String
Dim TotalCostes As String
Dim Valor As Currency

    On Error Resume Next

    ' Coste de trabajadores

    Sql = "select sum(total) from ("
    Sql = Sql & "select salarios.codcateg, round(sum(cclinorden1.horas) * salarios.impsalar,2) total "
    Sql = Sql & " from (cclinorden1 inner join straba on cclinorden1.codtraba = straba.codtraba) "
    Sql = Sql & " inner join salarios on straba.codcateg = salarios.codcateg "
    Sql = Sql & " where cclinorden1.codorden = " & text1(0).Text
    Sql = Sql & " group by 1) aaaaa "
    
    Text3(0).Text = DevuelveValor(Sql)
    
    Valor = CCur(TransformaPuntosComas(Text3(0).Text))
    If Valor <> 0 Then
        Text3(0).Text = Format(Valor, "###,###,##0.00")
    Else
        Text3(0).Text = ""
    End If
    
    ' Horas de trabajadores
    
    Sql = "select sum(cclinorden1.horas) "
    Sql = Sql & " from cclinorden1 where codorden = " & text1(0).Text
    
    Text3(2).Text = DevuelveValor(Sql)
    
    Valor = CCur(TransformaPuntosComas(Text3(2).Text))
    If Valor <> 0 Then
        Text3(2).Text = Format(Valor, "###,###,##0.00")
    Else
        Text3(2).Text = ""
    End If
    
    
    ' Coste de categorias

    Sql = "select round(sum(cclinorden2.horas) * salarios.impsalar,2) "
    Sql = Sql & " from cclinorden2 inner join salarios on cclinorden2.codcateg = salarios.codcateg "
    Sql = Sql & " where cclinorden2.codorden = " & text1(0).Text
    
    Text3(1).Text = DevuelveValor(Sql)
    
    Valor = CCur(TransformaPuntosComas(Text3(1).Text))
    If Valor <> 0 Then
        Text3(1).Text = Format(Valor, "###,###,##0.00")
    Else
        Text3(1).Text = ""
    End If
    
    ' Horas de trabajadores
    
    Sql = "select sum(cclinorden2.horas) "
    Sql = Sql & " from cclinorden2 where codorden = " & text1(0).Text
    
    Text3(3).Text = DevuelveValor(Sql)
    
    Valor = CCur(TransformaPuntosComas(Text3(3).Text))
    If Valor <> 0 Then
        Text3(3).Text = Format(Valor, "###,###,##0.00")
    Else
        Text3(3).Text = ""
    End If
    
    
    ' Kilos totales
    Sql = "select sum(if (kilosnet is null, 0,kilosnet)), sum(if(numcajon is null,0,numcajon)) from cclinorden3 "
    Sql = Sql & " where cclinorden3.codorden = " & text1(0).Text
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Text3(4).Text = DBLet(Rs.Fields(0).Value, "N")
        Text3(5).Text = DBLet(Rs.Fields(1).Value, "N")
        
    Else
        Text3(4).Text = ""
        Text3(5).Text = ""
    End If
    
    ' Kilos totales de forfaits
    Sql = "select sum(if (kilosnet is null, 0,kilosnet)), sum(if(numcajon is null,0,numcajon)) from cclinorden4 "
    Sql = Sql & " where cclinorden4.codorden = " & text1(0).Text
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Text3(6).Text = DBLet(Rs.Fields(0).Value, "N")
        Text3(7).Text = DBLet(Rs.Fields(1).Value, "N")
        
    Else
        Text3(6).Text = ""
        Text3(7).Text = ""
    End If
    
    If Err.Number <> 0 Then
        Err.Clear
    End If

End Sub

