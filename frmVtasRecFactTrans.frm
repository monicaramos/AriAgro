VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmVtasRecFactTrans 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturas de Transporte / Comisi�n"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   14250
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVtasRecFactTrans.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   14250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   135
      TabIndex        =   81
      Top             =   90
      Width           =   2010
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   82
         Top             =   180
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Pedir Datos"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver Albaranes"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Generar factura"
            EndProperty
         EndProperty
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5280
      Left            =   45
      TabIndex        =   60
      Top             =   3150
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   9313
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Albaranes"
      TabPicture(0)   =   "frmVtasRecFactTrans.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ListView1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Portes de Vuelta"
      TabPicture(1)   =   "frmVtasRecFactTrans.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "FrameAux0"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame FrameAux0 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   135
         TabIndex        =   62
         Top             =   360
         Width           =   7540
         Begin VB.TextBox txtaux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   5
            Left            =   1170
            MaxLength       =   35
            TabIndex        =   76
            Tag             =   "Fec.Factura|F|N|||tmpportesv|fecfactu|dd/mm/yyyy|S|"
            Text            =   "F.Factura"
            Top             =   3150
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.TextBox txtaux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   4
            Left            =   810
            MaxLength       =   35
            TabIndex        =   75
            Tag             =   "Factura|T|N|||tmpportesv|numfactu||S|"
            Text            =   "Factura"
            Top             =   3150
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.CommandButton btnBuscar 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   300
            Index           =   1
            Left            =   2250
            MaskColor       =   &H00000000&
            TabIndex        =   74
            ToolTipText     =   "Buscar Cta Contable"
            Top             =   3150
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   2430
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   73
            Text            =   "Text2"
            Top             =   3150
            Width           =   2520
         End
         Begin VB.Frame Frame1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   0
            Left            =   180
            TabIndex        =   71
            Top             =   3960
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
               TabIndex        =   72
               Top             =   180
               Width           =   2655
            End
         End
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "&Aceptar"
            Height          =   375
            Left            =   5175
            TabIndex        =   67
            Top             =   4050
            Width           =   1065
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "&Cancelar"
            Height          =   375
            Left            =   6375
            TabIndex        =   68
            Top             =   4050
            Width           =   1065
         End
         Begin VB.TextBox txtaux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   330
            Index           =   3
            Left            =   5085
            MaxLength       =   35
            TabIndex        =   66
            Tag             =   "Importe|N|N|||tmpportesv|importel|###,###0.00||"
            Text            =   "Importe"
            Top             =   3150
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtaux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   330
            Index           =   2
            Left            =   1650
            MaxLength       =   40
            TabIndex        =   65
            Tag             =   "Cta.Contable|T|N|||tmpportesv|codmacta||S|"
            Text            =   "Cta"
            Top             =   3150
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtaux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   0
            Left            =   90
            MaxLength       =   6
            TabIndex        =   64
            Tag             =   "Usuario|N|N|||tmpinformes|codusu|00000|S|"
            Text            =   "Usu"
            Top             =   3135
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtaux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   1
            Left            =   450
            MaxLength       =   3
            TabIndex        =   63
            Tag             =   "Transportista|N|N|||tmpportesv|codtrans|0000|S|"
            Text            =   "Tr"
            Top             =   3135
            Visible         =   0   'False
            Width           =   330
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   0
            Left            =   135
            TabIndex        =   69
            Top             =   0
            Width           =   1290
            _ExtentX        =   2275
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
            Left            =   3840
            Top             =   705
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
            Bindings        =   "frmVtasRecFactTrans.frx":0044
            Height          =   3510
            Index           =   0
            Left            =   135
            TabIndex        =   70
            Top             =   450
            Width           =   7185
            _ExtentX        =   12674
            _ExtentY        =   6191
            _Version        =   393216
            AllowUpdate     =   0   'False
            BorderStyle     =   0
            HeadLines       =   1
            RowHeight       =   19
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
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
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4545
         Left            =   -74910
         TabIndex        =   61
         Top             =   450
         Width           =   8140
         _ExtentX        =   14367
         _ExtentY        =   8017
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
   End
   Begin VB.Frame FrameIntro 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Left            =   135
      TabIndex        =   7
      Top             =   855
      Width           =   14045
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   27
         Left            =   6750
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Fecha Factura|F|N|||tcafpc|fecfactu|dd/mm/yyyy|S|"
         Top             =   585
         Width           =   1350
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         Index           =   0
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Tag             =   "Tipo|N|N|||tcafpc|tipofact|||"
         Top             =   585
         Width           =   1890
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   26
         Left            =   8415
         MaxLength       =   6
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   585
         Width           =   865
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   360
         Index           =   26
         Left            =   9315
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   58
         Text            =   "Text2"
         Top             =   585
         Width           =   4290
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   360
         Index           =   3
         Left            =   2250
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   14
         Text            =   "Text2"
         Top             =   1320
         Width           =   4185
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   2
         Left            =   5085
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Recepci�n|F|N|||tcafpc|fecrecep|dd/mm/yyyy|N|"
         Top             =   585
         Width           =   1350
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   360
         Index           =   5
         Left            =   9300
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   12
         Text            =   "Text2"
         Top             =   1320
         Width           =   4290
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   5
         Left            =   8400
         MaxLength       =   5
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1320
         Width           =   865
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   3
         Left            =   495
         MaxLength       =   3
         TabIndex        =   5
         Tag             =   "Cod. Transportista|N|N|0|999|tcafpc|codtrans|000|S|"
         Text            =   "Text1"
         Top             =   1320
         Width           =   1500
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   1
         Left            =   3600
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Factura|F|N|||tcafpc|fecfactu|dd/mm/yyyy|S|"
         Top             =   585
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   0
         Left            =   2265
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "N� Factura|T|N|||tcafpc|numfactu||S|"
         Text            =   "Text1 7"
         Top             =   585
         Width           =   1245
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Tesoreria"
         Height          =   375
         Index           =   0
         Left            =   4050
         TabIndex        =   46
         Top             =   1350
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2190
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Contabiliz."
         Height          =   375
         Index           =   1
         Left            =   4455
         TabIndex        =   47
         Top             =   945
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label1 
         Caption         =   "Desde F.Alb"
         Height          =   345
         Index           =   18
         Left            =   6750
         TabIndex        =   84
         Top             =   270
         Width           =   1185
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   7965
         Picture         =   "frmVtasRecFactTrans.frx":005C
         ToolTipText     =   "Buscar fecha"
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Factura"
         Height          =   255
         Index           =   17
         Left            =   135
         TabIndex        =   77
         Top             =   270
         Width           =   1860
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   210
         Index           =   16
         Left            =   8415
         TabIndex        =   59
         Top             =   270
         Width           =   720
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   9135
         ToolTipText     =   "Buscar cliente"
         Top             =   270
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   6390
         Picture         =   "frmVtasRecFactTrans.frx":00E7
         ToolTipText     =   "Buscar fecha"
         Top             =   285
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   4680
         Picture         =   "frmVtasRecFactTrans.frx":0172
         ToolTipText     =   "Buscar fecha"
         Top             =   270
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   10620
         ToolTipText     =   "Buscar banco propio"
         Top             =   1035
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   135
         ToolTipText     =   "Buscar transportista"
         Top             =   1350
         Width           =   285
      End
      Begin VB.Label Label1 
         Caption         =   "F.Recepci�n"
         Height          =   255
         Index           =   3
         Left            =   5085
         TabIndex        =   13
         Top             =   270
         Width           =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta Prevista Pago"
         Height          =   255
         Index           =   2
         Left            =   8400
         TabIndex        =   11
         Top             =   1035
         Width           =   2790
      End
      Begin VB.Label Label1 
         Caption         =   "Transportista / Comisionista"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   10
         Top             =   1035
         Width           =   2790
      End
      Begin VB.Label Label1 
         Caption         =   "F.Factura"
         Height          =   255
         Index           =   29
         Left            =   3585
         TabIndex        =   9
         Top             =   285
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "N� Factura"
         Height          =   255
         Index           =   28
         Left            =   2265
         TabIndex        =   8
         Top             =   285
         Width           =   1095
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   2160
      MaxLength       =   4
      TabIndex        =   49
      Text            =   "Text1"
      Top             =   1110
      Width           =   660
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   2880
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   48
      Text            =   "Text2"
      Top             =   1110
      Width           =   3615
   End
   Begin VB.Frame FrameFactura 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   8460
      TabIndex        =   15
      Top             =   3060
      Width           =   5695
      Begin VB.CommandButton CmdCan 
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   435
         TabIndex        =   80
         Top             =   4635
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "&Generar"
         Height          =   375
         Left            =   1605
         TabIndex        =   79
         Top             =   4635
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   360
         Index           =   25
         Left            =   3510
         MaxLength       =   15
         TabIndex        =   53
         Tag             =   "Importe Retencion|N|N|0||scafac|imporete|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   3885
         Width           =   1710
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   24
         Left            =   900
         MaxLength       =   5
         TabIndex        =   52
         Tag             =   "% reten|N|S|0|99.90|scafac|porcereten|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   3885
         Width           =   660
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   23
         Left            =   1695
         MaxLength       =   15
         TabIndex        =   51
         Tag             =   "Base Imponible 1|N|N|0||scafac|baseimp1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   3885
         Width           =   1620
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   360
         Index           =   9
         Left            =   3495
         MaxLength       =   15
         TabIndex        =   40
         Tag             =   "Importe IVA 1|N|N|0||scafac|imporiv1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   1395
         Width           =   1710
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   8
         Left            =   3495
         MaxLength       =   15
         TabIndex        =   39
         Tag             =   "Base Imponible 3|N|N|0||scafac|baseimp3|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   945
         Width           =   1710
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   7
         Left            =   3495
         MaxLength       =   15
         TabIndex        =   38
         Tag             =   "Base Imponible 2 |N|N|0||scafac|baseimp2|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   570
         Width           =   1710
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   6
         Left            =   3510
         MaxLength       =   15
         TabIndex        =   36
         Tag             =   "Base Imponible 1|N|N|0||scafac|baseimp1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   180
         Width           =   1710
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   12
         Left            =   360
         MaxLength       =   5
         TabIndex        =   34
         Tag             =   "% IVA 3|N|S|0|99|scafac|porciva3|00|N|"
         Text            =   "Text1 7"
         Top             =   3060
         Width           =   500
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   11
         Left            =   360
         MaxLength       =   5
         TabIndex        =   33
         Tag             =   "& IVA 2|N|S|0|99|scafac|porciva2|00|N|"
         Text            =   "Text1 7"
         Top             =   2685
         Width           =   500
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   10
         Left            =   360
         MaxLength       =   5
         TabIndex        =   32
         Tag             =   "% IVA 1|N|S|0|99|scafac|porciva1|00|N|"
         Text            =   "Text1 7"
         Top             =   2295
         Width           =   500
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   16
         Left            =   1665
         MaxLength       =   15
         TabIndex        =   25
         Tag             =   "Base Imponible 1|N|N|0||scafac|baseimp1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2325
         Width           =   1620
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   13
         Left            =   885
         MaxLength       =   5
         TabIndex        =   24
         Tag             =   "% IVA 1|N|S|0|99.90|scafac|porciva1|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   2295
         Width           =   705
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   360
         Index           =   19
         Left            =   3495
         MaxLength       =   15
         TabIndex        =   23
         Tag             =   "Importe IVA 1|N|N|0||scafac|imporiv1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2325
         Width           =   1710
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   17
         Left            =   1665
         MaxLength       =   15
         TabIndex        =   22
         Tag             =   "Base Imponible 2 |N|N|0||scafac|baseimp2|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2685
         Width           =   1620
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   14
         Left            =   885
         MaxLength       =   5
         TabIndex        =   21
         Tag             =   "& IVA 2|N|S|0|99.90|scafac|porciva2|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   2685
         Width           =   705
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   360
         Index           =   20
         Left            =   3495
         MaxLength       =   15
         TabIndex        =   20
         Tag             =   "Importe IVA 2|N|N|0||scafac|imporiv2|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2685
         Width           =   1710
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   18
         Left            =   1665
         MaxLength       =   15
         TabIndex        =   19
         Tag             =   "Base Imponible 3|N|N|0||scafac|baseimp3|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   3060
         Width           =   1620
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   15
         Left            =   900
         MaxLength       =   5
         TabIndex        =   18
         Tag             =   "% IVA 3|N|S|0|99.90|scafac|porciva3|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   3060
         Width           =   705
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   360
         Index           =   21
         Left            =   3495
         MaxLength       =   15
         TabIndex        =   17
         Tag             =   "Importe IVA 3|N|N|0||scafac|imporiv3|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   3060
         Width           =   1710
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   22
         Left            =   3180
         MaxLength       =   15
         TabIndex        =   16
         Tag             =   "Total Factura|N|N|0||scafac|totalfac|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   4665
         Width           =   2055
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   90
         ToolTipText     =   "Buscar codigo iva"
         Top             =   3075
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   90
         ToolTipText     =   "Buscar codigo iva"
         Top             =   2700
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   90
         ToolTipText     =   "Buscar codigo iva"
         Top             =   2295
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "% Ret"
         Height          =   255
         Index           =   15
         Left            =   900
         TabIndex        =   57
         Top             =   3630
         Width           =   720
      End
      Begin VB.Label Label1 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   14
         Left            =   3375
         TabIndex        =   56
         Top             =   3870
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Retenci�n"
         Height          =   255
         Index           =   13
         Left            =   3510
         TabIndex        =   55
         Top             =   3630
         Width           =   1560
      End
      Begin VB.Label Label1 
         Caption         =   "Base Imponible"
         Height          =   255
         Index           =   12
         Left            =   1695
         TabIndex        =   54
         Top             =   3630
         Width           =   1530
      End
      Begin VB.Line Line3 
         X1              =   2475
         X2              =   5235
         Y1              =   3510
         Y2              =   3510
      End
      Begin VB.Label Label1 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   3345
         TabIndex        =   45
         Top             =   945
         Width           =   135
      End
      Begin VB.Line Line2 
         X1              =   3195
         X2              =   5225
         Y1              =   1335
         Y2              =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Imp.Dto.General"
         Height          =   255
         Index           =   10
         Left            =   1665
         TabIndex        =   44
         Top             =   945
         Width           =   1710
      End
      Begin VB.Label Label1 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   3345
         TabIndex        =   43
         Top             =   570
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "Imp.Dto.PPago"
         Height          =   255
         Index           =   8
         Left            =   1665
         TabIndex        =   42
         Top             =   570
         Width           =   1935
      End
      Begin VB.Line Line1 
         X1              =   2460
         X2              =   5220
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label1 
         Caption         =   "Base Imponible"
         Height          =   255
         Index           =   7
         Left            =   1665
         TabIndex        =   41
         Top             =   1395
         Width           =   1710
      End
      Begin VB.Label Label1 
         Caption         =   "Bruto Factura"
         Height          =   255
         Index           =   6
         Left            =   1665
         TabIndex        =   37
         Top             =   240
         Width           =   1830
      End
      Begin VB.Label Label1 
         Caption         =   "Cod."
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   35
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Base Imponible"
         Height          =   255
         Index           =   4
         Left            =   1665
         TabIndex        =   31
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Importe IVA"
         Height          =   255
         Index           =   33
         Left            =   3495
         TabIndex        =   30
         Top             =   2070
         Width           =   1560
      End
      Begin VB.Label Label1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   37
         Left            =   3330
         TabIndex        =   29
         Top             =   2385
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   36
         Left            =   11880
         TabIndex        =   28
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "TOTAL FACTURA"
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
         Height          =   255
         Index           =   39
         Left            =   3195
         TabIndex        =   27
         Top             =   4410
         Width           =   2025
      End
      Begin VB.Label Label1 
         Caption         =   "% IVA"
         Height          =   255
         Index           =   41
         Left            =   885
         TabIndex        =   26
         Top             =   2025
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   7080
      Top             =   5640
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
   Begin VB.Label Label1 
      Caption         =   "Cargando Albaranes"
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
      Height          =   255
      Index           =   19
      Left            =   270
      TabIndex        =   83
      Top             =   8685
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Label Label1 
      Caption         =   "Operador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1845
      TabIndex        =   50
      Top             =   1035
      Width           =   735
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   1845
      Picture         =   "frmVtasRecFactTrans.frx":01FD
      ToolTipText     =   "Buscar trabajador"
      Top             =   1125
      Width           =   240
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnPedirDatos 
         Caption         =   "&Pedir Datos"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnVerAlbaran 
         Caption         =   "&Ver Albaranes"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnGenerarFac 
         Caption         =   "&Generar Factura"
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Shortcut        =   ^I
         Visible         =   0   'False
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmVtasRecFactTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)


'========== VBLES PRIVADAS ====================
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1

Private WithEvents frmTrans As frmManAgencias
Attribute frmTrans.VB_VarHelpID = -1
Private WithEvents frmCli As frmClientes  'Form Mto clientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmBanPr As frmManBanco 'Mto de Bancos propios
Attribute frmBanPr.VB_VarHelpID = -1
Private WithEvents frmCtas As frmCtasConta 'Cuentas contables
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmTipIva As frmTipIVAConta  'Tipos de IVA de la contabilidad
Attribute frmTipIva.VB_VarHelpID = -1

Private WithEvents frmMens As frmMensajes  'Para introducir el importe de portes y de comision
Attribute frmMens.VB_VarHelpID = -1

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

'cadena donde se almacena la WHERE para la seleccion de los albaranes
'marcados para facturar
Dim cadwhere As String

'Cuando vuelve del formulario de ver los albaranes seleccionados hay que volver
'a cargar los datos de los albaranes
'Dim VerAlbaranes As Boolean

Dim PrimeraVez As Boolean
Dim Bloquear As Boolean
Dim indice As Integer

Dim indCodigo As Integer


'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer
'-------------------------------------------------------------------------

Private vTrans As CTransportista

Dim ModoLineas As Byte
'1.- Afegir,  2.- Modificar,  3.- Borrar,  0.-Passar control a Ll�nies
Dim NumTabMto As Integer 'Indica quin n� de Tab est� en modo Mantenimient

Dim vWhere As String
Dim PorlineaAlbaran As Boolean
Dim FechaAnt As String

Dim mCantidad As String


Private Sub btnBuscar_Click(Index As Integer)
    If vParamAplic.NumeroConta = 0 Then Exit Sub
    
    indice = Index + 1
    Set frmCtas = New frmCtasConta
    frmCtas.NumDigit = 0
    frmCtas.DatosADevolverBusqueda = "0|1|"
    frmCtas.CodigoActual = Text1(indice).Text
    frmCtas.Show vbModal
    Set frmCtas = Nothing
    PonerFoco Text1(indice)
End Sub

Private Sub CmdCan_Click()
Dim i As Integer

    FrameIntro.Enabled = False
    SSTab1.Enabled = True
    FrameFactura.Enabled = False
    
    
    BloquearTxt Text1(10), True
    BloquearTxt Text1(11), True
    BloquearTxt Text1(12), True
    
    
    For i = 4 To 6
        imgBuscar(i).Enabled = False
        imgBuscar(i).visible = False
    Next i
    
    
    Me.CmdCan.Enabled = False
    Me.CmdCan.visible = False
    Me.cmdGenerar.Enabled = False
    Me.cmdGenerar.visible = False


End Sub

Private Sub cmdGenerar_Click()
Dim i As Integer

    FrameIntro.Enabled = False
    SSTab1.Enabled = True
    FrameFactura.Enabled = False


    For i = 4 To 6
        imgBuscar(i).Enabled = False
        imgBuscar(i).visible = False
    Next i
    
    
    Me.CmdCan.Enabled = False
    Me.CmdCan.visible = False
    Me.cmdGenerar.Enabled = False
    Me.cmdGenerar.visible = False


    BotonFacturar
    Set vTrans = Nothing
    Screen.MousePointer = vbDefault

End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    SSTab1.TabEnabled(1) = (Combo1(0).ListIndex = 0)
    
    mnVerAlbaran.Enabled = (vParamAplic.Cooperativa = 5)
    Me.Toolbar1.Buttons(2).Enabled = mnVerAlbaran.Enabled
    
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
'    If VerAlbaranes Then RefrescarAlbaranes
'    VerAlbaranes = False
End Sub

Private Sub Form_Load()
Dim i As Integer
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .DisabledImageList = frmPpal.imgListComun_BN

        .Buttons(1).Image = 1   'Pedir Datos
        .Buttons(2).Image = 3   'Ver albaranes
        .Buttons(3).Image = 15   'Generar FActura
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
    
    'cargar IMAGES de busqueda
    Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    For i = 2 To 6
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    
    Me.FrameFactura.Enabled = False
    
    LimpiarCampos   'Limpia los campos TextBox
    InicializarListView
   
    '## A mano
    NombreTabla = "albaran" 'cabecera albaranes de venta
    
    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    CadenaConsulta = "Select * from " & NombreTabla
    CadenaConsulta = CadenaConsulta & " where numalbar=-1"
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
        
    CargaCombo
        
    Combo1(0).ListIndex = -1
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    End If
    Me.SSTab1.Tab = 0
    CargaGrid 0, False
    
    '[Monica]14/12/2018: asignamos la fecha de inicio de campa�a a la fecha inicio de campa�a
    Text1(27).Text = DateAdd("d", 1, DateAdd("yyyy", -1, vParam.FecFinCam))
    FechaAnt = Text1(27).Text
    
    PrimeraVez = False
End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me
    'Aqui va el especifico de cada form es
    '### a mano
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    DesBloqueoManual "FACTRA"
    TerminaBloquear
'    DesBloqueoManual ("scaalp")
End Sub


Private Sub frmBanPr_DatoSeleccionado(CadenaSeleccion As String)
    'Form de Mantenimiento de Bancos Propios
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 1)
    Text1(5).Text = Format(Text1(5).Text, "0000")
    Text2(5).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Clientes
    Text1(26).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod cliente
    FormateaCampo Text1(26)
    Text2(26).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom cliente
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
    txtaux(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cuenta
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmF_Selec(vFecha As Date)
Dim indice As Byte
    indice = CByte(Me.imgFecha(0).Tag)
    Text1(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
    mCantidad = CadenaSeleccion
End Sub

Private Sub frmTipIva_DatoSeleccionado(CadenaSeleccion As String)
'Tipos de IVA (de la Contabilidad)
    Text1(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codigiva
    FormateaCampo Text1(indCodigo)
    Text1(indCodigo + 3).Text = RecuperaValor(CadenaSeleccion, 3) '% iva
    RecalcularDatosFactura
End Sub

Private Sub frmTrans_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Agencias de transporte
Dim indice As Byte
    indice = 3
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Proveedor
    FormateaCampo Text1(indice)
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom proveedor
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
'    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. Proveedor
            Set frmTrans = New frmManAgencias
            frmTrans.DatosADevolverBusqueda = "0|1|"
            frmTrans.Show vbModal
            Set frmTrans = Nothing
            indice = 3
       
       Case 2 'Bancos Propios
            indice = 5
            Set frmBanPr = New frmManBanco
            frmBanPr.DatosADevolverBusqueda = "0|1|"
            frmBanPr.Show vbModal
            Set frmBanPr = Nothing
            
       Case 3 'Clientes
            Set frmCli = New frmClientes
            frmCli.DatosADevolverBusqueda = "0|2|"
            frmCli.Show vbModal
            Set frmCli = Nothing
            indice = 26
       
       Case 4, 5, 6
            indCodigo = Index + 6
        
            Set frmTipIva = New frmTipIVAConta
            frmTipIva.DeConsulta = True
            frmTipIva.DatosADevolverBusqueda = "0|1|2|"
            frmTipIva.CodigoActual = Text1(indCodigo).Text
            frmTipIva.Show vbModal
            Set frmTipIva = Nothing
       
    End Select
    
    PonerFoco Text1(indice)
'    Screen.MousePointer = vbDefault
End Sub

Private Sub imgFecha_Click(Index As Integer)
Dim indice As Byte
Dim esq As Long
Dim dalt As Long
Dim menu As Long
Dim obj As Object

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
    
   Set frmF = New frmCal
    
   esq = imgFecha(Index).Left
   dalt = imgFecha(Index).Top
    
   Set obj = imgFecha(Index).Container

   While imgFecha(Index).Parent.Name <> obj.Name
       esq = esq + obj.Left
       dalt = dalt + obj.Top
       Set obj = obj.Container
   Wend
    
   menu = Me.Height - Me.ScaleHeight 'ac� tinc el heigth del men� i de la toolbar

   frmF.Left = esq + imgFecha(Index).Parent.Left + 30
   frmF.Top = dalt + imgFecha(Index).Parent.Top + imgFecha(Index).Height + menu - 40
   
   frmF.NovaData = Now
   indice = Index + 1
   
   '[Monica]14/12/2018: ponemos la fecha para coger los albaranes
   If Index = 2 Then indice = 27
   
   Me.imgFecha(0).Tag = indice
   
   PonerFormatoFecha Text1(indice)
   If Text1(indice).Text <> "" Then frmF.NovaData = CDate(Text1(indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(indice)
End Sub


Private Sub ListView1_ItemClick(ByVal item As MSComctlLib.ListItem)
Dim Cantidad As String
Dim palets As String
Dim Sql As String
Dim Valor As Currency
Dim valor2 As Currency
Dim i As Long
Dim b As Boolean

    If Modo <> 5 Then Exit Sub

    If Bloquear = True Then
        ListView1.SetFocus
        item.EnsureVisible
        Exit Sub
    End If
    
    mCantidad = ""
    
    '[Monica]14/12/2018: ponemos en frame el input de importes
    Select Case Combo1(0).ListIndex
        Case 0 ' transportista
            'Cantidad = InputBox("Introduzca el Porte para el albar�n:" & item.Text, "Portes", , 5000, 4000)
            Set frmMens = New frmMensajes
            frmMens.CADENA = "Introduzca el Porte para el albar�n: " & item.Text
            frmMens.OpcionMensaje = 31
            frmMens.Show vbModal
            Set frmMens = Nothing
            
        Case 1 ' comisionista
            'Cantidad = InputBox("Introduzca la Comisi�n para el albar�n:" & item.Text, "Comisi�n", , 5000, 4000)
            Set frmMens = New frmMensajes
            frmMens.CADENA = "Introduzca la Comisi�n para el albar�n: " & item.Text
            frmMens.OpcionMensaje = 31
            frmMens.Show vbModal
            Set frmMens = Nothing
    End Select
    
    Cantidad = mCantidad
    If Cantidad = "" Then Exit Sub
    
    b = True
    If vParamAplic.PortesKiloCaja = 1 And Combo1(0).ListIndex = 0 Then
'        palets = InputBox("Introduzca el N�mero de Palets para el albar�n:" & item.Text, "Palets", , 5000, 4000)
        
        mCantidad = ""
        Set frmMens = New frmMensajes
        frmMens.CADENA = "Introduzca Palets para el albar�n: " & item.Text
        frmMens.OpcionMensaje = 31
        frmMens.Label2(29).Caption = "Palets"
        frmMens.Show vbModal
        Set frmMens = Nothing
        
        palets = mCantidad
        
        If palets = "" Then Exit Sub
        
        If EsNumerico(palets) Then
            If InStr(1, palets, ",") > 0 Then
                valor2 = ImporteFormateado(palets)
            Else
                valor2 = CCur(TransformaPuntosComas(palets))
            End If
        Else
            b = False
        End If
    End If
    
    If b And EsNumerico(Cantidad) Then
        If InStr(1, Cantidad, ",") > 0 Then
            Valor = ImporteFormateado(Cantidad)
        Else
            Valor = CCur(TransformaPuntosComas(Cantidad))
        End If
    
        Select Case Combo1(0).ListIndex
            Case 0 ' factura de transportista
                Sql = "update albaran set portespag = " & DBSet(CStr(Valor), "N")
                '[Monica]14/12/2018: para no volver a cargar el listview
                ListView1.SelectedItem.SubItems(5) = Format(Valor, "###,###,##0.00")
                
                If vParamAplic.PortesKiloCaja = 1 Then
                    Sql = Sql & ", paletspag = " & DBSet(CStr(valor2), "N")
                
                    '[Monica]14/12/2018: para no volver a cargar el listview
                    ListView1.SelectedItem.SubItems(3) = valor2
                End If
                Sql = Sql & " where numalbar = " & DBSet(item.Text, "N")
                
            
            
            Case 1 ' factura de comisionista
                Sql = "update albaran set comisionespag = " & DBSet(CStr(Valor), "N")
                Sql = Sql & " where numalbar = " & DBSet(item.Text, "N")
        
                '[Monica]14/12/2018: para no volver a cargar el listview
                ListView1.SelectedItem.SubItems(5) = Format(Valor, "###,###,##0.00")
        
        
        End Select
        conn.Execute Sql
        
        i = ListView1.SelectedItem.Index
        
'[Monica]14/12/2018: quito esto para que no tarde tanto en refrescar
'        CargarAlbaranes vWhere, Combo1(0).ListIndex
        CalcularDatosFactura
        
        ' Crea una variable ListItem.
        ' Establece la variable al elemento encontrado.
        If i < ListView1.ListItems.Count Then
            ListView1.SelectedItem = ListView1.ListItems.item(i + 1)
        Else
            ListView1.SelectedItem = ListView1.ListItems.item(i)
        End If
        ListView1.SetFocus
        item.EnsureVisible
        
    End If

End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)

    Bloquear = True

End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
Dim item As MSComctlLib.ListItem

    Set item = ListView1.FindItem(ListView1.SelectedItem.Text, , 1, lvwPartial)

    Bloquear = True
    If KeyAscii = 13 Then Bloquear = False

    ListView1_ItemClick item

End Sub


'Private Sub ListView1_ItemCheck(ByVal item As MSComctlLib.ListItem)
''Cuando se selecciona un albaran de la lista
'Dim i As Integer
'Dim cad As String
'Dim TipoFP As Integer 'Forma de pago
'Dim TipoDtoPP As Currency 'descuento pronto pago
'Dim tipoDtoGn As Currency 'descuento general
'
'    Screen.MousePointer = vbHourglass
'
'    Set ListView1.SelectedItem = item
'
'    'Inicializamos a cero
'    TipoFP = 0
'    TipoDtoPP = 0
'    tipoDtoGn = 0
'
'    'cuando seleccionamos un check vemos si lo podemos seleccionar
'    'ya que si ya habia algun albaran selecionado tendremos que comprobar
'    'que son de la misma forpa, dtoppago y dtognral.
'    'si esto no se cumple no se pueden agrupar en la misma factura
'    For i = 1 To ListView1.ListItems.Count
'        If item.Index <> i Then
'            If ListView1.ListItems(i).Checked Then
'                'ya habia otro albaran seleccionado
'                TipoFP = ListView1.ListItems(i).SubItems(2)
'                TipoDtoPP = CCur(ListView1.ListItems(i).SubItems(4))
'                tipoDtoGn = CCur(ListView1.ListItems(i).SubItems(5))
'                Exit For
'            End If
'        End If
'    Next i
'
'    If Not (TipoFP = 0 And TipoDtoPP = 0 And tipoDtoGn = 0) Then
'    'si ya habia un albaran seleccionado, comprobar que es del mismo tipo
'        If item.SubItems(2) <> TipoFP Or item.SubItems(4) <> TipoDtoPP Or item.SubItems(5) <> tipoDtoGn Then
'            MsgBox "Se debe seleccionar albaranes de la misma Forma de Pago y Descuentos", vbExclamation
'            ListView1.SelectedItem.Checked = False
'            Screen.MousePointer = vbDefault
'            ListView1.SetFocus
'            Exit Sub
'        End If
'    Else
'    End If
'
'    ' Calculamos los datos de factura
'    If Not VerAlbaranes Then CalcularDatosFactura
'    Screen.MousePointer = vbDefault
'End Sub


Private Sub mnGenerarFac_Click()
Dim i As Integer

    FrameIntro.Enabled = False
    SSTab1.Enabled = False
    FrameFactura.Enabled = True
    
    For i = 6 To 22
        BloquearTxt Text1(i), True
    Next i
    
    For i = 6 To 22
        Text1(i).Enabled = False
    Next i
    
    
    BloquearTxt Text1(10), (Text1(16).Text = "")
    BloquearTxt Text1(11), (Text1(17).Text = "")
    BloquearTxt Text1(12), (Text1(18).Text = "")
    
    imgBuscar(4).Enabled = (Text1(16).Text <> "")
    imgBuscar(5).Enabled = (Text1(17).Text <> "")
    imgBuscar(6).Enabled = (Text1(18).Text <> "")
    imgBuscar(4).visible = (Text1(16).Text <> "")
    imgBuscar(5).visible = (Text1(17).Text <> "")
    imgBuscar(6).visible = (Text1(18).Text <> "")
    
    Me.CmdCan.Enabled = True
    Me.CmdCan.visible = True
    Me.cmdGenerar.Enabled = True
    Me.cmdGenerar.visible = True
    
    PonerFoco Text1(10)


'    BotonFacturar
'    Set vTrans = Nothing
'    Screen.MousePointer = vbDefault
End Sub

Private Sub mnPedirDatos_Click()
    BotonPedirDatos
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub


Private Sub mnVerAlbaran_Click()
Dim Albaranes As String

    If CargarLineasAlbaranes Then
        Set frmVtasRecFactTrans2.lw1 = Me.ListView1
        frmVtasRecFactTrans2.Show vbModal
        
'        If Not ImportesLineasCuadrados(Albaranes) Then
'            MsgBox "No coinciden los importes por albar�n con las l�neas en los siguientes albaranes: " & vbCrLf & Albaranes, vbExclamation
'        End If
        
        PorlineaAlbaran = True
    End If
    

End Sub

Private Function ImportesLineasCuadrados(ByRef Albaranes As String) As Boolean
Dim Sql As String
Dim i As Long


    Albaranes = ""

    For i = 1 To Me.ListView1.ListItems.Count
        If CCur(ListView1.ListItems(i).SubItems(5)) <> 0 Then
            Sql = "select sum(portespag) from tmpalbaranes where codusu = " & vUsu.Codigo
            Sql = Sql & " and numalbar = " & DBSet(ListView1.ListItems(i).Text, "N")
            
            If CCur(DevuelveValor(Sql)) <> CCur(Me.ListView1.ListItems(i).SubItems(5)) Then
                Albaranes = Albaranes & Me.ListView1.ListItems(i).Text & ","
            End If
        End If
    Next i
    
    If Albaranes <> "" Then Albaranes = Mid(Albaranes, 1, Len(Albaranes) - 1)
    
    ImportesLineasCuadrados = (Albaranes = "")

End Function


Private Function CargarLineasAlbaranes() As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim NetoTotal As Long
Dim RestoPortes As Currency
Dim PortesLinea As Currency
Dim Linea As Integer
Dim PaletsLinea As Currency
Dim i As Long

    On Error GoTo eCargarLineasAlbaranes


    CargarLineasAlbaranes = False

    Sql = "delete from tmpalbaranes where codusu = " & DBSet(vUsu.Codigo, "N")
    conn.Execute Sql
    
    For i = 1 To Me.ListView1.ListItems.Count
        If ListView1.ListItems(i).SubItems(5) <> 0 Then
    
            If vParamAplic.PortesKiloCaja = 0 Then ' caso de que se calculen por kilos
                Sql2 = "select sum(pesoneto) from albaran_variedad where numalbar = " & DBSet(ListView1.ListItems(i).Text, "N")
                Set Rs1 = New ADODB.Recordset
                Rs1.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                NetoTotal = 0
                If Not Rs1.EOF Then NetoTotal = DBLet(Rs1.Fields(0).Value, "N")
                Set Rs1 = Nothing
                
                If NetoTotal <> 0 Then
                    RestoPortes = ListView1.ListItems(i).SubItems(5)
                    
                    Sql2 = "select numlinea, pesoneto from albaran_variedad where numalbar = " & DBSet(ListView1.ListItems(i).Text, "N")
                    Sql2 = Sql2 & " order by numlinea "
                    
                    Set Rs1 = New ADODB.Recordset
                    Rs1.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    While Not Rs1.EOF
                        PortesLinea = Round2(DBLet(Rs1!Pesoneto, "N") * ListView1.ListItems(i).SubItems(5) / NetoTotal, 2)
                        RestoPortes = RestoPortes - PortesLinea
                        
                        Linea = DBLet(Rs1!NumLinea, "N")
                        
                        Sql3 = "select count(*) from tmpalbaranes where numalbar = " & DBSet(ListView1.ListItems(i).Text, "N")
                        Sql3 = Sql3 & " and numlinea = " & DBSet(Rs1!NumLinea, "N")
                        Sql3 = Sql3 & " and codusu = " & DBSet(vUsu.Codigo, "N")
                        
                        If TotalRegistros(Sql3) = 0 Then
                            Sql = "INSERT INTO tmpalbaranes(codusu, numalbar, numlinea, portespag) "
                            Sql = Sql & " VALUES (" & DBSet(vUsu.Codigo, "N") & "," & DBSet(ListView1.ListItems(i).Text, "N") & "," & DBSet(Rs1!NumLinea, "N") & ","
                            Sql = Sql & DBSet(PortesLinea, "N") & ")"
                        
                        Else
                            Sql = "update tmpalbaranes set portespag = portespag + " & DBSet(PortesLinea, "N")
                            Sql = Sql & " where numalbar = " & DBSet(ListView1.ListItems(i).Text, "N")
                            Sql = Sql & " and numlinea = " & DBSet(Rs1!NumLinea, "N")
                            Sql = Sql & " and codusu = " & DBSet(vUsu.Codigo, "N")
                        End If
                        
                        conn.Execute Sql
                    
                        Rs1.MoveNext
                    Wend
                    Set Rs1 = Nothing
                    
                    If RestoPortes <> 0 Then
                        Sql = "update tmpalbaranes set portespag = portespag + " & DBSet(RestoPortes, "N")
                        Sql = Sql & " where numalbar = " & DBSet(ListView1.ListItems(i).Text, "N") & " and numlinea = " & DBSet(Linea, "N")
                        Sql = Sql & " and codusu = " & DBSet(vUsu.Codigo, "N")
                        
                        conn.Execute Sql
                        
                    End If
                End If
            Else
                '14/12/2018
            
                ' caso de que se calculen por cajas
                Sql2 = "select albaran_variedad.numlinea, albaran_variedad.numcajas, albaran_variedad.codforfait, forfaits.cajaspalet, albaran.paletspag from albaran_variedad, forfaits, albaran where albaran_variedad.numalbar = " & DBSet(ListView1.ListItems(i).Text, "N")
                Sql2 = Sql2 & " and albaran_variedad.codforfait = forfaits.codforfait "
                Sql2 = Sql2 & " and albaran_variedad.numalbar = albaran.numalbar "
            
                Set Rs1 = New ADODB.Recordset
                Rs1.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
                '[Monica]06/09/2018: daba error el campo rs!portespag
                RestoPortes = ListView1.ListItems(i).SubItems(5) 'DBLet(Rs!portespag, "N")
                
                While Not Rs1.EOF
                    PortesLinea = 0
                    Linea = DBLet(Rs1.Fields(0).Value, "N")
                    
                    If DBLet(Rs1.Fields(3).Value, "N") <> 0 Then
                        PaletsLinea = Round2(DBLet(Rs1.Fields(1).Value, "N") / DBLet(Rs1.Fields(3).Value, "N"), 2)
                        
                        If ListView1.ListItems(i).SubItems(3) <> 0 Then
                            '[Monica]28/01/2019: Portes por linea, se redondea a 2 (antes estaba a 4)
                            PortesLinea = Round2(ListView1.ListItems(i).SubItems(5) * PaletsLinea / DBLet(Rs1!paletspag, "N"), 2)
                        End If
                        ' --monica:cambiado por lo de arriba pq ahora tenemos en cuenta el numero de palets de la linea de albaran.
                        'PortesLinea = Round2(DBLet(Rs!portespag, "N") / DBLet(Rs1.Fields(3).Value, "N") * DBLet(Rs1.Fields(1).Value, "N"), 4)
                        RestoPortes = RestoPortes - PortesLinea
                    End If
                    
                    Sql3 = "select count(*) from tmpalbaranes where numalbar = " & DBSet(ListView1.ListItems(i).Text, "N")
                    Sql3 = Sql3 & " and numlinea = " & DBSet(Rs1!NumLinea, "N")
                    Sql3 = Sql3 & " and codusu = " & DBSet(vUsu.Codigo, "N")
                    
                    If TotalRegistros(Sql3) = 0 Then
                        '[Monica]06/09/2018: daba error el campo es portespag
                        Sql = "INSERT INTO tmpalbaranes(codusu, numalbar, numlinea, portespag) " 'impcoste, importes) "
                        Sql = Sql & " VALUES (" & DBSet(vUsu.Codigo, "N") & "," & DBSet(ListView1.ListItems(i).Text, "N") & "," & DBSet(Rs1!NumLinea, "N") & ","
                        Sql = Sql & DBSet(PortesLinea, "N") & ")" '& "," & DBSet(PortesLinea, "N") & ")"
                    Else
                        '[Monica]06/09/2018: daba error el campo impcoste
                        Sql = "update tmpalbaranes set portespag = portespag + " & DBSet(PortesLinea, "N") ' antes impcoste
'                        Sql = Sql & ", importes = importes + " & DBSet(PortesLinea, "N")
                        Sql = Sql & " where numalbar = " & DBSet(ListView1.ListItems(i).Text, "N")
                        Sql = Sql & " and numlinea = " & DBSet(Rs1!NumLinea, "N")
                        Sql = Sql & " and codusu = " & vUsu.Codigo
                    End If
                    conn.Execute Sql
                    
                    
                    Rs1.MoveNext
                Wend
                
                Set Rs1 = Nothing
                    
                If RestoPortes <> 0 Then
                    Sql = "update tmpalbaranes set portespag = portespag + " & DBSet(RestoPortes, "N") ' antes impcoste
 '                   Sql = Sql & ", importes = importes + " & DBSet(RestoPortes, "N")
                    Sql = Sql & " where numalbar = " & DBSet(ListView1.ListItems(i).Text, "N") & " and numlinea = " & DBSet(Linea, "N")
                    Sql = Sql & " and codusu = " & DBSet(vUsu.Codigo, "N")
                    
                    conn.Execute Sql
                    
                End If
            End If ' del caso en el que los calculos de portes se hagan por cajas
    
        End If
    Next i
    
    
    
    
'    Sql = "select numalbar, portespag, paletspag from albaran where portespag <> 0"
'    Set Rs = New ADODB.Recordset
'    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    While Not Rs.EOF
'        If vParamAplic.PortesKiloCaja = 0 Then ' caso de que se calculen por kilos
'            Sql2 = "select sum(pesoneto) from albaran_variedad where numalbar = " & DBLet(Rs!numalbar, "N")
'            Set Rs1 = New ADODB.Recordset
'            Rs1.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'            NetoTotal = 0
'            If Not Rs1.EOF Then NetoTotal = DBLet(Rs1.Fields(0).Value, "N")
'            Set Rs1 = Nothing
'
'            If NetoTotal <> 0 Then
'                RestoPortes = DBLet(Rs!portespag, "N")
'
'                Sql2 = "select numlinea, pesoneto from albaran_variedad where numalbar = " & DBLet(Rs!numalbar, "N")
'                Sql2 = Sql2 & " order by numlinea "
'
'                Set Rs1 = New ADODB.Recordset
'                Rs1.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'                While Not Rs1.EOF
'                    PortesLinea = Round2(DBLet(Rs1!Pesoneto, "N") * DBLet(Rs!portespag, "N") / NetoTotal, 2)
'                    RestoPortes = RestoPortes - PortesLinea
'
'                    Linea = DBLet(Rs1!numlinea, "N")
'
'                    Sql3 = "select count(*) from tmpalbaranes where numalbar = " & DBLet(Rs!numalbar, "N")
'                    Sql3 = Sql3 & " and numlinea = " & DBSet(Rs1!numlinea, "N")
'                    Sql3 = Sql3 & " and codusu = " & DBSet(vUsu.codigo, "N")
'
'                    If TotalRegistros(Sql3) = 0 Then
'                        Sql = "INSERT INTO tmpalbaranes(codusu, numalbar, numlinea, impcoste, importes) "
'                        Sql = Sql & " VALUES (" & DBSet(vUsu.codigo, "N") & "," & DBSet(Rs!numalbar, "N") & "," & DBSet(Rs1!numlinea, "N") & ","
'                        Sql = Sql & DBSet(PortesLinea, "N") & "," & DBSet(PortesLinea, "N") & ")"
'
'                    Else
'                        Sql = "update tmpalbaranes set impcoste = impcoste + " & DBSet(PortesLinea, "N")
'                        Sql = Sql & ", importes = importes + " & DBSet(PortesLinea, "N")
'                        Sql = Sql & " where numalbar = " & DBSet(Rs!numalbar, "N")
'                        Sql = Sql & " and numlinea = " & DBSet(Rs1!numlinea, "N")
'                        Sql = Sql & " and codusu = " & DBSet(vUsu.codigo, "N")
'                    End If
'
'                    conn.Execute Sql
'
'                    Rs1.MoveNext
'                Wend
'                Set Rs1 = Nothing
'
'                If RestoPortes <> 0 Then
'                    Sql = "update tmpalbaranes set impcoste = impcoste + " & DBSet(RestoPortes, "N")
'                    Sql = Sql & ", importes = importes + " & DBSet(RestoPortes, "N")
'                    Sql = Sql & " where numalbar = " & DBLet(Rs!numalbar, "N") & " and numlinea = " & DBSet(Linea, "N")
'                    Sql = Sql & " and codusu = " & DBSet(vUsu.codigo, "N")
'
'                    conn.Execute Sql
'
'                End If
'            End If
'        Else
'            ' caso de que se calculen por cajas
'            Sql2 = "select albaran_variedad.numlinea, albaran_variedad.numcajas, albaran_variedad.codforfait, forfaits.cajaspalet from albaran_variedad, forfaits where numalbar = " & DBLet(Rs!numalbar, "N")
'            Sql2 = Sql2 & " and albaran_variedad.codforfait = forfaits.codforfait "
'
'            Set Rs1 = New ADODB.Recordset
'            Rs1.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'            RestoPortes = DBLet(Rs!portespag, "N")
'
'            While Not Rs1.EOF
'                PortesLinea = 0
'                Linea = DBLet(Rs1.Fields(0).Value, "N")
'
'                If DBLet(Rs1.Fields(3).Value, "N") <> 0 Then
'                    PaletsLinea = Round2(DBLet(Rs1.Fields(1).Value, "N") / DBLet(Rs1.Fields(3).Value, "N"), 2)
'
'                    If DBLet(Rs!paletspag, "N") <> 0 Then
'                        PortesLinea = Round2(DBLet(Rs!portespag, "N") * PaletsLinea / DBLet(Rs!paletspag, "N"), 4)
'                    End If
'                    ' --monica:cambiado por lo de arriba pq ahora tenemos en cuenta el numero de palets de la linea de albaran.
'                    'PortesLinea = Round2(DBLet(Rs!portespag, "N") / DBLet(Rs1.Fields(3).Value, "N") * DBLet(Rs1.Fields(1).Value, "N"), 4)
'                    RestoPortes = RestoPortes - PortesLinea
'                End If
'
'                Sql3 = "select count(*) from tmpalbaranes where numalbar = " & DBLet(Rs!numalbar, "N")
'                Sql3 = Sql3 & " and numlinea = " & DBSet(Rs1!numlinea, "N")
'                Sql3 = Sql3 & " and codusu = " & DBSet(vUsu.codigo, "N")
'
'                If TotalRegistros(Sql3) = 0 Then
'                    Sql = "INSERT INTO tmpalbaranes(codusu, numalbar, numlinea, impcoste, importes) "
'                    Sql = Sql & " VALUES (" & DBSet(vUsu.codigo, "N") & "," & DBSet(Rs!numalbar, "N") & "," & DBSet(Rs1!numlinea, "N") & ","
'                    Sql = Sql & DBSet(PortesLinea, "N") & "," & DBSet(PortesLinea, "N") & ")"
'                Else
'                    Sql = "update tmpalbaranes set impcoste = impcoste + " & DBSet(PortesLinea, "N")
'                    Sql = Sql & ", importes = importes + " & DBSet(PortesLinea, "N")
'                    Sql = Sql & " where numalbar = " & DBSet(Rs!numalbar, "N")
'                    Sql = Sql & " and numlinea = " & DBSet(Rs1!numlinea, "N")
'                    Sql = Sql & " and codusu = " & vUsu.codigo
'                End If
'                conn.Execute Sql
'
'
'                Rs1.MoveNext
'            Wend
'
'            Set Rs1 = Nothing
'
'            If RestoPortes <> 0 Then
'                Sql = "update tmpalbaranes set impcoste = impcoste + " & DBSet(RestoPortes, "N")
'                Sql = Sql & ", importes = importes + " & DBSet(RestoPortes, "N")
'                Sql = Sql & " where numalbar = " & DBLet(Rs!numalbar, "N") & " and numlinea = " & DBSet(Linea, "N")
'                Sql = Sql & " and codusu = " & DBSet(vUsu.codigo, "N")
'
'                conn.Execute Sql
'
'            End If
'        End If ' del caso en el que los calculos de portes se hagan por cajas
'        Rs.MoveNext
'    Wend
'    Set Rs = Nothing


    CargarLineasAlbaranes = True
    Exit Function
    
eCargarLineasAlbaranes:
    MuestraError Err.Number, "Cargar Lineas de albaran", Err.Description
End Function
    
    
    
    
'Private Sub mnVerAlbaran_Click()
'    BotonVerAlbaranes
'End Sub

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
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
Dim Sql3 As String

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1, 2 'Fecha factura, fecha recepcion
            PonerFormatoFecha Text1(Index)
            If Text1(Index) <> "" Then
                ' No debe existir el n�mero de factura para el proveedor en hco
                If ExisteFacturaEnHco Then
                    InicializarListView
                End If
            End If
            
            '[Monica]19/06/2018: compruebo la fecha
            If Index = 2 Then
                If Text1(2).Text <> "" Then
                    'Comprobar que la fecha de RECEPCION esta dentro de los ejercicios contables
                    If vParamAplic.NumeroConta <> 0 Then
                        ResultadoFechaContaOK = EsFechaOKConta(CDate(Text1(2).Text))
                        If ResultadoFechaContaOK > 0 Then
                            If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
                            PonerFoco Text1(2)
                        End If
                    End If
                End If
            End If
            
        Case 27 ' fecha de inicio de albaranes
            PonerFormatoFecha Text1(Index)
            
        Case 3 'Cod Transportista
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "agencias", "nomtrans", "codtrans")
                ' No debe existir el n�mero de factura para el proveedor en hco
                If ExisteFacturaEnHco Then
                    InicializarListView
                Else
                    'comprobamos que no haya nadie recepcionando facturas de ese proveedor
'                    DesBloqueoManual ("FACTRA")
'                    If Not Bloqu�eoManual("FACTRA", Text1(3).Text) Then
                    Select Case Combo1(0).ListIndex
                        Case 0 ' transportista
                            '[Monica]16/04/2019: el codigo de transporte est� en lineas (albaran_transporte)
                            'vWhere = "codtrans = " & DBSet(Text1(3).Text, "N")
                            vWhere = "albaran.numalbar in (select distinct numalbar from albaran_transporte where codtrans = " & DBSet(Text1(3).Text, "N") & ")"
                        Case 1 ' comisionista
                            vWhere = "codcomis = " & DBSet(Text1(3).Text, "N")
                    End Select
                    If Text1(26).Text <> "" Then vWhere = vWhere & " and codclien = " & DBSet(Text1(26).Text, "N")
                    
                    '[Monica]14/12/2018: si hay fecha cogemos todos los albaranes que sean de fecha mayor o igual
                    If Text1(27).Text <> "" Then vWhere = vWhere & " and albaran.fechaalb >= " & DBSet(Text1(27).Text, "F")
                    
                    If Not BloqueaRegistro("albaran", vWhere) Then
                        Select Case Combo1(0).ListIndex
                            Case 0
                                MsgBox "No se puede recepcionar factura de ese transportista. Hay otro usuario recepcionando.", vbExclamation
                            Case 1
                                MsgBox "No se puede recepcionar factura de ese comisionista. Hay otro usuario recepcionando.", vbExclamation
                        End Select
                        BotonPedirDatos
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    Else
                        '--monica:080908
                        'PonerModo 5
                        '--
                        '[Monica]06/02/2019: a�ado control de que tiene que tener cuenta contable
                        If AgenciaSinCtaContable(Text1(3)) Then
                            Text1(3).Text = ""
                            PonerFoco Text1(3)
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If
                        
                        
                        Label1(19).visible = True
                        Label1(19).Caption = "Inicializando albaranes"
                        DoEvents
                        Screen.MousePointer = vbHourglass
                        
                        Select Case Combo1(0).ListIndex
                            Case 0
                                Sql3 = "update albaran set portespag = 0 where " & vWhere
                                conn.Execute Sql3
                            Case 1
                                Sql3 = "update albaran set comisionespag = 0 where " & vWhere
                                conn.Execute Sql3
                        End Select
                        
                        Label1(19).visible = False
                        DoEvents
                        Screen.MousePointer = vbDefault
                        
                        
                        CargarAlbaranes vWhere, Combo1(0).ListIndex
                        CalcularDatosFactura
                    End If
                    
                End If
                
            Else
                Text2(Index).Text = ""
            End If
            
        Case 5 'Cta Prevista de PAgo
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "banpropi", "nombanpr", "codbanpr")
                Text1(Index).Text = Format(Text1(Index).Text, "0000")
            Else
                Text2(Index).Text = ""
            End If
            
            '++monica:080908
            If Not ExisteFacturaEnHco And Text1(3) <> "" Then
                PonerModo 5
                PonerFocoLw Me.ListView1
            End If
            
        Case 26 'Cliente
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "clientes", "nomclien", "codclien")
                Text1(Index).Text = Format(Text1(Index).Text, "000000")
            Else
                Text2(Index).Text = ""
            End If
    
        Case 10, 11, 12 ' codigo de iva
            If PonerFormatoEntero(Text1(Index)) Then
                Text1(Index + 3).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", Text1(Index).Text, "N")
            Else
                Text1(Index + 3).Text = ""
            End If
        
        
            RecalcularDatosFactura
    
    End Select
End Sub


Private Function AgenciaSinCtaContable(vcodigo As String) As Boolean
Dim vSQL As String

    AgenciaSinCtaContable = True

    vSQL = DevuelveDesdeBDNew(cAgro, "agencias", "codmacta", "codtrans", vcodigo, "N")
    If vSQL <> "" Then
        vSQL = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", vSQL, "T")
        If vSQL = "" Then
            MsgBox "Cuenta contable no existe en contabilidad.", vbExclamation
        Else
            AgenciaSinCtaContable = False
        End If
    Else
        MsgBox "La Agencia de transporte/comisionista no tiene asignada una cuenta contable. Revise.", vbExclamation
    End If

End Function


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte, NumReg As Byte
Dim b As Boolean
On Error GoTo EPonerModo

    Modo = Kmodo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
        
                 
'    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
'    'Si estamos en Insertar adem�s limpia los campos Text1
'    'si estamos en modificar bloquea las compos que son clave primaria
'    BloquearText1 Me, Modo
    
    For i = 0 To Text1.Count - 1
        BloquearTxt Text1(i), (Modo <> 3)
    Next i
    
    'Importes siempre bloqueados
    For i = 6 To 25
        BloquearTxt Text1(i), True
    Next i
    
    If Combo1(0).ListIndex = -1 Then BloquearCombo Me, Modo
    
    'Campo B.Imp y Imp. IVA siempre en azul
    Text1(9).BackColor = &HFFFFC0 'Base imponible
    Text1(19).BackColor = &HFFFFC0 'Total Iva 1
    Text1(20).BackColor = &HFFFFC0 'Iva 2
    Text1(21).BackColor = &HFFFFC0 'IVa 3
    Text1(22).BackColor = &HC0C0FF 'Total factura
    Text1(25).BackColor = &HFFFFC0 'Imp.Retencion
    
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2 And Modo <> 5)
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = b
    Next i
                    
'    Me.chkVistaPrevia.Enabled = (Modo <= 2)
    
    For i = 0 To txtaux.Count - 1
        BloquearTxt txtaux(i), True
        txtaux(i).visible = False
    Next i
        
    Me.FrameIntro.Enabled = (Modo = 3)
    Me.FrameAux0.Enabled = (Modo = 5)
       
    BloquearBtn btnBuscar(1), True
    Text2(2).visible = False
       
    'Poner el tama�o de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
'    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu seg�n modo
    PonerOpcionesMenu 'Activar opciones de menu seg�n nivel de permisos del usuario
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
'Comprobar que los datos del frame de introduccion son correctos antes de cargar datos
Dim vtag As CTag
Dim Cad As String
Dim i As Byte

    On Error GoTo EDatosOK
    DatosOk = False
    
    ' deben de introducirse todos los datos del frame
    For i = 0 To 5
        If Text1(i).Text = "" And i <> 4 Then
            If Text1(i).Tag <> "" Then
                Set vtag = New CTag
                If vtag.Cargar(Text1(i)) Then
                    Cad = vtag.Nombre
                Else
                    Cad = "Campo"
                End If
                Set vtag = Nothing
            Else
                Cad = "Campo"
                If i = 5 Then Cad = "Cta. Prev. Pago"
            End If
            MsgBox Cad & " no puede estar vacio. Reintroduzca", vbExclamation
            PonerModo 3
            PonerFoco Text1(i)
            Exit Function
        End If
    Next i
        
    'comprobar que la fecha de la factura sea anterior a la fecha de recepcion
    If Not EsFechaIgualPosterior(Text1(1).Text, Text1(2).Text, True, "La fecha de recepci�n debe ser igual o posterior a la fecha de la factura.") Then
        Exit Function
    End If
    
    
    'Comprobar que la fecha de RECEPCION esta dentro de los ejercicios contables
    If vParamAplic.NumeroConta <> 0 Then
        ResultadoFechaContaOK = EsFechaOKConta(CDate(Text1(2).Text))
        If ResultadoFechaContaOK > 0 Then
            If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
            Exit Function
        End If
    End If
    


'--monica:03/12/2008
'    'comprobar que se han seleccionado lineas para facturar
'    If cadWHERE = "" Then
'        MsgBox "Debe seleccionar albaranes para facturar.", vbExclamation
'        Exit Function
'    End If
    
'++monica:03/12/2008
    'comprobamos que hay lineas para facturar: o albaranes o portes de vuelta
    If cadwhere = "" Then
        If AdoAux(0).Recordset.RecordCount = 0 Then
            MsgBox "No hay albaranes ni portes de vuelta para incluir en la factura. Revise.", vbExclamation
            Exit Function
        End If
    End If
    
    ' No debe existir el n�mero de factura para el transportista en hco
    If ExisteFacturaEnHco Then Exit Function
    
'--monica
'    'todos los albaranes seleccionados deben tener la misma: forma pago, dto ppago, dto gnral
'    cad = "select count(distinct codforpa,dtoppago,dtognral) from scaalp "
'    cad = cad & " WHERE " & Replace(cadWHERE, "slialp", "scaalp")
'    If RegistrosAListar(cad) > 1 Then
'        MsgBox "No se puede facturar albaranes con distintas: forma de pago, dto gral, dto ppago.", vbExclamation
'        Exit Function
'    End If
    
    
    'Si la forpa es TRANSFERENCIA entonces compruebo la si tiene cta bancaria
'    cad = "select distinct (codforpa) from scaalp "
'    cad = cad & " WHERE " & Replace(cadWHERE, "slialp", "scaalp")
    Set miRsAux = New ADODB.Recordset
'    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    cad = miRsAux.Fields(0)
'    miRsAux.Close
    
    
    
    'Ahora buscamos el tipforpa del codforpa
    Cad = "Select tipoforp from forpago where codforpa=" & vTrans.ForPago 'cad
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    If miRsAux.EOF Then
        MsgBox "Error en el TIPO de forma de pago", vbExclamation
    Else
        i = 1
        Cad = miRsAux.Fields(0)
        If Val(Cad) = vbFPTransferencia Then
            'Compruebo que la forpa es transferencia
            i = 2
        End If
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    
    
    If i = 2 Then
        'La forma de pago es transferencia. Debo comprobar que existe la cuenta bancaria
        'del proveedor
        If vTrans.CuentaBan = "" Or vTrans.DigControl = "" Or vTrans.Sucursal = "" Or vTrans.Banco = "" Then
            Cad = "Cuenta bancaria incorrecta. Forma de pago: transferencia.    �Continuar?"
            If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then i = 0
        End If
    End If
    
    'Si i=0 es que o esta mal la forpa o no quiere seguir pq no tiene cuenta bancaria
    If i > 0 Then DatosOk = True
    Exit Function
    
EDatosOK:
    DatosOk = False
    MuestraError Err.Number, "Comprobar datos correctos", Err.Description
End Function



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Pedir datos
             mnPedirDatos_Click
            
        Case 2 ' Ver albaranes
            mnVerAlbaran_Click
            
        Case 3 'Generar Factura
            mnGenerarFac_Click
    End Select
End Sub


Private Sub PonerOpcionesMenu()
Dim J As Byte

    PonerOpcionesMenuGeneral Me
    
    J = Val(Me.mnPedirDatos.HelpContextID)
    If J < vUsu.Nivel Then Me.mnPedirDatos.Enabled = False
    
    J = Val(Me.mnGenerarFac.HelpContextID)
    If J < vUsu.Nivel Then Me.mnGenerarFac.Enabled = False
    
    '[Monica]13/12/2018: se a�ade a natural
    Me.mnVerAlbaran.Enabled = (vParamAplic.Cooperativa = 5 Or vParamAplic.Cooperativa = 12 Or vParamAplic.Cooperativa = 9)
    Me.Toolbar1.Buttons(2).Enabled = mnVerAlbaran.Enabled
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

 
Private Sub BotonPedirDatos()
Dim Nombre As String

    
    PorlineaAlbaran = False
    
    TerminaBloquear

    FrameIntro.Enabled = True
    SSTab1.Enabled = False
    FrameFactura.Enabled = False

'    '[Monica]14/12/2018: en el caso de no queramos inicializarlo todo
'    If Text1(3).Text <> "" Then
'        If MsgBox("� Desea inicializar los datos ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
'            PonerModo 3
'            Exit Sub
'        End If
'    End If

    'Vaciamos todos los Text
    LimpiarCampos
    
    Text1(27).Text = FechaAnt

    'Vaciamos el ListView
    InicializarListView
    CargaGrid 0, False

    InicializarTemporal
    SSTab1.Tab = 0
    'Como no habr� albaranes seleccionados vaciamos la cadwhere
    cadwhere = ""
    
    PonerModo 3
    
    'fecha recepcion
    Text1(2).Text = Format(Now, "dd/mm/yyyy")
    
    
    'desbloquear los registros de la saalp (si hay bloquedos)
    TerminaBloquear
    
    'si vamos
    'desBloqueo Manual de las tablas
'    DesBloqueoManual ("scaalp")
    
    ' si no hay nada
    If Combo1(0).ListIndex = -1 Then Combo1(0).ListIndex = 0 ' inicialmente es transportista
    
    PonerFoco Text1(0)
End Sub


Private Sub CargarAlbaranes(cadwhere As String, Tipo As Byte)
'Tipo = 0 transportista
'Tipo = 1 comisionista
'Recupera de la BD y muestra en el Listview todos los albaranes de compra
'que tiene el proveedor introducido.
Dim Sql As String, Sql3 As String
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
On Error GoTo ECargar

    ListView1.ListItems.Clear
    
    'si no hay trasportista/comisionista salir
    If Text1(3).Text = "" Then Exit Sub
    
    Select Case Tipo
        Case 0 ' factura de transportista
        
            'sum(if (coalesce(ac.tipogasto,0)=2,ac.impcoste,0))
        
            If vParamAplic.PortesKiloCaja = 0 Then
'                SQL = "SELECT albaran.numalbar,albaran.fechaalb,albaran.matriveh,albaran.matrirem,albaran.portespag "
                Sql = "SELECT albaran.numalbar,albaran.fechaalb,albaran.matriveh,albaran.matrirem,sum(albaran_costes.impcoste) as portesac ,albaran.portespag "
            Else                                                                'monica:080908 visualizaba el totpalet
'                SQL = "SELECT albaran.numalbar,albaran.fechaalb,albaran.matriveh,albaran.paletspag as totpalet, albaran.portespag "
                Sql = "SELECT albaran.numalbar,albaran.fechaalb,albaran.matriveh,albaran.paletspag as totpalet, sum(albaran_costes.impcoste) as portesac, albaran.portespag "
            End If
            '[Monica]11/12/2018: tiene que ser la condicion en el ON obligatoriamente pq sino no devuelve registros
            Sql = Sql & " FROM (albaran LEFT JOIN albaran_costes ON albaran.numalbar = albaran_costes.numalbar and albaran_costes.tipogasto = 2) LEFT JOIN albaran_transporte on albaran.numalbar = albaran_transporte.numalbar  "
            Sql = Sql & " WHERE " & cadwhere '& Text1(3).Text
            
'[Monica] 04/01/2010 : un mismo albaran puede ser recepcionado varias veces por lo que el coste ir� aumentando
'            SQL = SQL & " and albaran.numalbar not in (select numalbar from albaran_costes where tipogasto = 2) "
'[Monica] 04/01/2010 : los albaranes ir�n ordenados por el importe pagado (primero los que no han sido pagados
            Sql = Sql & " group by 1,2,3,4,6 "
            Sql = Sql & " order by portesac asc, albaran.numalbar , albaran.fechaalb "
        
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
            InicializarListView
            
            Label1(19).visible = True
            Label1(19).Caption = "Cargando albaranes"
            DoEvents
            Screen.MousePointer = vbHourglass
            
            While Not Rs.EOF
                Set ItmX = ListView1.ListItems.Add()
                ItmX.Text = Rs!NumAlbar
                ItmX.SubItems(1) = Format(DBLet(Rs!FechaAlb, "F"), "dd/mm/yyyy")
                ItmX.SubItems(2) = DBLet(Rs!matriveh, "T")
                If vParamAplic.PortesKiloCaja = 0 Then
                    ItmX.SubItems(3) = DBLet(Rs!matrirem, "T")
                Else
                    ItmX.SubItems(3) = Format(DBLet(Rs!TotPalet, "N"), "#,##0")
                End If
                ItmX.SubItems(4) = Format(DBLet(Rs!portesac, "N"), "###,##0.00")
                ItmX.SubItems(5) = Format(DBLet(Rs!portespag, "N"), "###,##0.00")
                'Sig
                Rs.MoveNext
            Wend
            
            Rs.Close
            Set Rs = Nothing
            
      Case 1 ' factura de comisionista
            
            Sql = "SELECT albaran.numalbar,albaran.fechaalb,albaran.matriveh,albaran.matrirem,sum(albaran_costes.impcoste) as comisionesac, albaran.comisionespag "
            Sql = Sql & " FROM albaran LEFT JOIN albaran_costes ON albaran.numalbar = albaran_costes.numalbar and albaran_costes.tipogasto = 3 "
            Sql = Sql & " WHERE " & cadwhere
            'SQL = SQL & " and albaran.numalbar not in (select numalbar from albaran_costes where tipogasto = 2) "
            
            Sql = Sql & " group by 1,2,3,4,6"
            Sql = Sql & " order by comisionesac, numalbar, fechaalb "
        
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
            InicializarListView
            
            Label1(19).visible = True
            Label1(19).Caption = "Cargando albaranes"
            DoEvents
            Screen.MousePointer = vbHourglass
            
            While Not Rs.EOF
                Set ItmX = ListView1.ListItems.Add()
                ItmX.Text = Rs!NumAlbar
                ItmX.SubItems(1) = Format(DBLet(Rs!FechaAlb, "F"), "dd/mm/yyyy")
                ItmX.SubItems(2) = DBLet(Rs!matriveh, "T")
                ItmX.SubItems(3) = DBLet(Rs!matrirem, "T")
                ItmX.SubItems(4) = Format(DBLet(Rs!comisionesac, "N"), "###,##0.00")
                ItmX.SubItems(5) = Format(DBLet(Rs!comisionespag, "N"), "###,##0.00")
                'Sig
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
        
      
   End Select
   
   SSTab1.Enabled = True
   
ECargar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando Albaranes", Err.Description

    Screen.MousePointer = vbDefault
    Label1(19).visible = False
    DoEvents


End Sub


Private Sub InicializarListView()
'Inicializa las columnas del List view

    ListView1.ListItems.Clear
    
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Albaran", 1100 ' 1100
    ListView1.ColumnHeaders.Add , , "Fecha", 1400, 2 '1270, 2
    ListView1.ColumnHeaders.Add , , "Mat.Vehiculo", 1500 '"Mat.Veh�culo", 1300
    
    Select Case Combo1(0).ListIndex
        Case 0
            If vParamAplic.PortesKiloCaja = 0 Then
                ListView1.ColumnHeaders.Add , , "Mat.Remolque", 1200 ' "Mat.Remolque", 1300
            Else
                ListView1.ColumnHeaders.Add , , "Nro.Palets", 1200 '1300
            End If
    
            ListView1.ColumnHeaders.Add , , "Portes Ac.", 1500, 1
            ListView1.ColumnHeaders.Add , , "Portes", 1100, 1 '1300, 1
            
        Case 1
            ListView1.ColumnHeaders.Add , , "Mat.Remolque", 1400 '"Mat.Remolque", 1300
            
            ListView1.ColumnHeaders.Add , , "Comisi�n Ac.", 1350, 1
            ListView1.ColumnHeaders.Add , , "Comisi�n", 1050, 1 ' 1300, 1

    End Select
End Sub

Private Sub InicializarTemporal()
'Inicializa la tmpinformes
Dim Sql As String
    
    Sql = "delete from tmpportesv where codusu = " & vUsu.Codigo
    conn.Execute Sql
    
End Sub


Private Sub CalcularDatosFactura()
Dim i As Integer
Dim Sql As String
Dim cadAux As String
Dim ImpBruto As Currency
Dim impiva As Currency
Dim vFactu As CFacturaTra
Dim Rs As ADODB.Recordset
Dim CadTabla As String

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For i = 6 To 25
         Text1(i).Text = ""
    Next i


    cadAux = ""
    cadwhere = ""
    ImpBruto = 0
    
    For i = 1 To ListView1.ListItems.Count
        If DBLet(ListView1.ListItems(i).SubItems(5), "N") <> 0 Then
        'para cada albaran que tenga importe de portes
            ImpBruto = ImpBruto + DBSet(ListView1.ListItems(i).SubItems(5), "N")
            
            Sql = "(albaran.numalbar=" & DBSet(ListView1.ListItems(i).Text, "N") & ") "
            If cadAux = "" Then
                cadAux = Sql
            Else
                cadAux = cadAux & " OR " & Sql
            End If
        End If
    Next i
    
    Sql = "select sum(importel) from tmpportesv where codusu = " & vUsu.Codigo & " and codtrans = " & DBSet(Text1(3).Text, "N")
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    If Not Rs.EOF Then ImpBruto = ImpBruto + DBLet(Rs.Fields(0).Value, "N")
    
'    InicializarListView
    
'    While Not rs.EOF
    
    
    If cadAux <> "" Then
    'se han seleccionado albaranes para facturar
    'Esta el la cadena WHERE de los albaranes seleccionados para obtener
    'el bruto de las lineas de los albaranes agrupadas por tipo de iva
    
        Select Case Combo1(0).ListIndex
            Case 0
    
                cadwhere = "albaran_transporte.codtrans=" & Val(Text1(3).Text)
                CadTabla = "albaran left join albaran_transporte on albaran.numalbar = albaran_transporte.numalbar"
            Case 1
            
                cadwhere = "albaran.codcomis=" & Val(Text1(3).Text)
                CadTabla = "albaran"
        End Select
        
        cadwhere = cadwhere & " AND (" & cadAux & ")"
    
        If Not SeleccionaRegistros(Combo1(0).ListIndex) Then Exit Sub
        
        If Not BloqueaRegistro(CadTabla, cadwhere) Then
            Select Case Combo1(0).ListIndex
                Case 0
        
                    conn.Execute "update albaran, albaran_transporte set portespag = 0 where albaran.numalbar = albaran_transporte.numalbar and " & cadwhere
                    
                    
                Case 1
                    
                    conn.Execute "update albaran set comisionespag = 0 where " & cadwhere
                    
            End Select
            
                
            CargarAlbaranes vWhere, Combo1(0).ListIndex
            
        End If
    
        Set vFactu = New CFacturaTra
        vFactu.DtoPPago = 0
        vFactu.DtoGnral = 0
        vFactu.EsComisionista = Combo1(0).ListIndex
        If vFactu.CalcularDatosFactura(cadwhere, Text1(3).Text) Then
            Text1(6).Text = vFactu.BrutoFac
            Text1(7).Text = vFactu.ImpPPago
            Text1(8).Text = vFactu.ImpGnral
            Text1(9).Text = vFactu.BaseImp
            Text1(10).Text = vFactu.TipoIVA1
            Text1(11).Text = vFactu.TipoIVA2
            Text1(12).Text = vFactu.TipoIVA3
            Text1(13).Text = vFactu.PorceIVA1
            Text1(14).Text = vFactu.PorceIVA2
            Text1(15).Text = vFactu.PorceIVA3
            Text1(16).Text = vFactu.BaseIVA1
            Text1(17).Text = vFactu.BaseIVA2
            Text1(18).Text = vFactu.BaseIVA3
            Text1(19).Text = vFactu.ImpIVA1
            Text1(20).Text = vFactu.ImpIVA2
            Text1(21).Text = vFactu.ImpIVA3
            Text1(22).Text = vFactu.TotalFac
            Text1(24).Text = vFactu.PorcReten
            Text1(23).Text = vFactu.BaseImp
            Text1(25).Text = vFactu.ImpReten
            
            For i = 6 To 26
                FormateaCampo Text1(i)
            Next i
            'Quitar ceros de linea IVA 2
            If Val(Text1(14).Text) = 0 And Val(Text1(11).Text) = 0 Then
                For i = 11 To 20 Step 3
                    Text1(i).Text = QuitarCero(CCur(Text1(i).Text))
                Next i
            End If
            'Quitar ceros de linea IVA 3
            If Val(Text1(15).Text) = 0 And Val(Text1(12).Text) = 0 Then
                For i = 12 To 21 Step 3
                    Text1(i).Text = QuitarCero(CCur(Text1(i).Text))
                Next i
            End If
            
        Else
            MuestraError Err.Number, "Calculando Factura", Err.Description
        End If
        Set vFactu = Nothing
    
    
    
    Else
        Set vFactu = New CFacturaTra
        vFactu.DtoPPago = 0
        vFactu.DtoGnral = 0
        vFactu.EsComisionista = Combo1(0).ListIndex
        If vFactu.CalcularDatosFacturaSinAlbaran(cadwhere, Text1(3).Text) Then
            Text1(6).Text = vFactu.BrutoFac
            Text1(7).Text = vFactu.ImpPPago
            Text1(8).Text = vFactu.ImpGnral
            Text1(9).Text = vFactu.BaseImp
            Text1(10).Text = vFactu.TipoIVA1
            Text1(11).Text = vFactu.TipoIVA2
            Text1(12).Text = vFactu.TipoIVA3
            Text1(13).Text = vFactu.PorceIVA1
            Text1(14).Text = vFactu.PorceIVA2
            Text1(15).Text = vFactu.PorceIVA3
            Text1(16).Text = vFactu.BaseIVA1
            Text1(17).Text = vFactu.BaseIVA2
            Text1(18).Text = vFactu.BaseIVA3
            Text1(19).Text = vFactu.ImpIVA1
            Text1(20).Text = vFactu.ImpIVA2
            Text1(21).Text = vFactu.ImpIVA3
            Text1(22).Text = vFactu.TotalFac
            Text1(24).Text = vFactu.PorcReten
            Text1(23).Text = vFactu.BaseImp
            Text1(25).Text = vFactu.ImpReten
            
            For i = 6 To 26
                FormateaCampo Text1(i)
            Next i
            'Quitar ceros de linea IVA 2
            If Val(Text1(14).Text) = 0 And Val(Text1(11).Text) = 0 Then
                For i = 11 To 20 Step 3
                    Text1(i).Text = QuitarCero(CCur(Text1(i).Text))
                Next i
            End If
            'Quitar ceros de linea IVA 3
            If Val(Text1(15).Text) = 0 And Val(Text1(12).Text) = 0 Then
                For i = 12 To 21 Step 3
                    Text1(i).Text = QuitarCero(CCur(Text1(i).Text))
                Next i
            End If
            
        Else
            MuestraError Err.Number, "Calculando Factura", Err.Description
        End If
        Set vFactu = Nothing
    End If
    
End Sub

Private Function SeleccionaRegistros(Tipo As Byte) As Boolean
'Comprueba que se seleccionan albaranes en la base de datos
'es decir que hay albaranes marcados
'cuando se van marcando albaranes se van a�adiendo el la cadena cadWhere
Dim Sql As String

    On Error GoTo ESel
    SeleccionaRegistros = False
    
    If cadwhere = "" Then Exit Function
    
    
    If Tipo = 0 Then
        Sql = "Select count(*) FROM albaran left join albaran_transporte on albaran.numalbar = albaran_transporte.numalbar "
    Else
        Sql = "Select count(*) FROM albaran  "
    End If
    
    Sql = Sql & " WHERE " & cadwhere
    If RegistrosAListar(Sql) <> 0 Then SeleccionaRegistros = True
    Exit Function
    
ESel:
    SeleccionaRegistros = False
    MuestraError Err.Number, "No hay seleccionados Albaranes", Err.Description
End Function


Private Sub BotonFacturar()
Dim vFactu As CFacturaTra
Dim Cad As String

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    
    Cad = ""
    If Text1(3).Text = "" Then
        Select Case Combo1(0).ListIndex
            Case 0
                Cad = "Falta transportista"
            Case 1
                Cad = "Falta comisionista"
            Case Else
                Cad = "Faltan datos"
        End Select
    Else
        If Not IsNumeric(Text1(3).Text) Then Cad = "Campo transportista/comisionista debe ser num�rico"
    End If
    If Cad <> "" Then
        MsgBox Cad, vbExclamation
        Exit Sub
    End If
        
        
        
    Set vTrans = New CTransportista
    
    'Tiene que ller los datos del transportista
    If Not vTrans.LeerDatos(Text1(3).Text) Then Exit Sub
    
    If Not DatosOk Then Exit Sub

        'Pasar los Albaranes seleccionados con cadWHERE a una factura
        Set vFactu = New CFacturaTra
        vFactu.Transportista = Text1(3).Text
        vFactu.NumFactu = Text1(0).Text
        vFactu.FecFactu = Text1(1).Text
        vFactu.FecRecep = Text1(2).Text
        vFactu.Trabajador = Text1(4).Text
        vFactu.BancoPr = Text1(5).Text
        vFactu.BrutoFac = ImporteFormateado(Text1(6).Text)
        vFactu.ForPago = vTrans.ForPago
        vFactu.DtoPPago = 0
        vFactu.DtoGnral = 0
        vFactu.ImpPPago = ImporteFormateado(Text1(7).Text)
        vFactu.ImpGnral = ImporteFormateado(Text1(8).Text)
        vFactu.BaseIVA1 = ImporteFormateado(Text1(16).Text)
        vFactu.BaseIVA2 = ImporteFormateado(Text1(17).Text)
        vFactu.BaseIVA3 = ImporteFormateado(Text1(18).Text)
        vFactu.TipoIVA1 = ComprobarCero(Text1(10).Text)
        vFactu.TipoIVA2 = ComprobarCero(Text1(11).Text)
        vFactu.TipoIVA3 = ComprobarCero(Text1(12).Text)
        vFactu.PorceIVA1 = ComprobarCero(Text1(13).Text)
        vFactu.PorceIVA2 = ComprobarCero(Text1(14).Text)
        vFactu.PorceIVA3 = ComprobarCero(Text1(15).Text)
        vFactu.ImpIVA1 = ImporteFormateado(Text1(19).Text)
        vFactu.ImpIVA2 = ImporteFormateado(Text1(20).Text)
        vFactu.ImpIVA3 = ImporteFormateado(Text1(21).Text)
        vFactu.TotalFac = ImporteFormateado(Text1(22).Text)
        vFactu.PorcReten = ImporteFormateado(Text1(24).Text)
        vFactu.ImpReten = ImporteFormateado(Text1(25).Text)
        vFactu.EsComisionista = Combo1(0).ListIndex
        
        'Si el proveedor tiene CTA BANCARIA se la asigno
        vFactu.CCC_Entidad = vTrans.Banco
        vFactu.CCC_Oficina = vTrans.Sucursal
        vFactu.CCC_CC = vTrans.DigControl
        vFactu.CCC_CTa = vTrans.CuentaBan
        '[Monica]22/11/2013: Tema iban
        vFactu.CCC_Iban = vTrans.Iban
        
        If cadwhere <> "" Then
            If vFactu.TraspasoAlbaranesAFactura(cadwhere, PorlineaAlbaran) Then BotonPedirDatos
        Else
            If vFactu.TraspasoPortesVueltaAFactura(cadwhere) Then BotonPedirDatos
        End If
        Set vFactu = Nothing
    
    
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Function ExisteFacturaEnHco() As Boolean
'Comprobamos si la factura ya existe en la tabla de Facturas a Proveedor: scafpc
Dim Cad As String

    ExisteFacturaEnHco = False
    'Tiene que tener valor los 3 campos de clave primaria antes de comprobar
    If Not (Text1(0).Text <> "" And Text1(1).Text <> "" And Text1(3).Text <> "") Then Exit Function
    
    ' No debe existir el n�mero de factura para el proveedor en hco
    Cad = "SELECT count(*) FROM tcafpc "
    Cad = Cad & " WHERE codtrans=" & Text1(3).Text & " AND numfactu=" & DBSet(Text1(0).Text, "T") & " AND year(fecfactu)=" & Year(Text1(1).Text)
    If RegistrosAListar(Cad) > 0 Then
        Select Case Combo1(0).ListIndex
            Case 0 'transportista
                MsgBox "Factura de transportista ya existente. Reintroduzca.", vbExclamation
            Case 1 'comisionista
                MsgBox "Factura de comisionista ya existente. Reintroduzca.", vbExclamation
        End Select
        
        ExisteFacturaEnHco = True
        Exit Function
    End If
End Function

'Private Sub RefrescarAlbaranes()
'Dim i As Integer
'Dim Sql As String
'Dim Itm As ListItem
'Dim RS As ADODB.Recordset
'
'
'    For i = 1 To ListView1.ListItems.Count
'        Sql = "SELECT albaran.numalbar,albaran.fechaalb,albaran.codclien,clientes.nomclien, albaran.imporrec "
'        Sql = Sql & " FROM albaran INNER JOIN clientes ON albaran.codclien=clientes.codclien "
'        Sql = Sql & " WHERE albaran.codtrans =" & Text1(3).Text & " AND albaran.numalbar=" & DBSet(ListView1.ListItems(i).Text, "T") & " AND albaran.fechaalb=" & DBSet(ListView1.ListItems(i).SubItems(1), "F")
'        Sql = Sql & " ORDER BY albaran.numalbar"
'
'        Set RS = New ADODB.Recordset
'        RS.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'        If Not RS.EOF Then 'Actualizamos los datos de este item en el list
'            ListView1.ListItems(i).SubItems(1) = Format(RS!FechaAlb, "dd/mm/yyyy")
'            ListView1.ListItems(i).SubItems(2) = RS!nomclien
'            ListView1.ListItems(i).SubItems(3) = RS!matriveh
'            ListView1.ListItems(i).SubItems(4) = Format(RS!portespag, "###,##0.00")
'
'        End If
'
'        RS.Close
'        Set RS = Nothing
'    Next i
'
''--monica
''    'recalcular el total de la factura
''     For i = 1 To ListView1.ListItems.Count
''        If ListView1.ListItems(i).Checked Then
''            CalcularDatosFactura
''            Exit For
''        End If
''     Next i
'End Sub
'


'****************************************

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
    PonerModo 5

    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar(Index) Then Exit Sub
    NumTabMto = Index
    Eliminar = False
   
    vWhere = ObtenerWhereCab(True)
    
    ' ***** independentment de si tenen datagrid o no,
    ' canviar els noms, els formats i el DELETE *****
    Select Case Index
        Case 0 'destino
            Sql = "�Seguro que desea eliminar el importe?"
            Sql = Sql & vbCrLf & "Cuenta: " & AdoAux(Index).Recordset!Codmacta & " - Importe: " & AdoAux(Index).Recordset!ImporteL
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM tmpportesv "
                Sql = Sql & vWhere & " AND codtrans = " & AdoAux(Index).Recordset!codTrans
                Sql = Sql & " AND numfactu = " & AdoAux(Index).Recordset!NumFactu
                Sql = Sql & " AND fecfactu = " & DBSet(AdoAux(Index).Recordset!FecFactu, "F")
                Sql = Sql & " AND codmacta = " & DBSet(AdoAux(Index).Recordset!Codmacta, "T")
            End If
            
    End Select

    If Eliminar Then
        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
'        TerminaBloquear
        conn.Execute Sql
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        If Index <> 3 Then _
            CargaGrid Index, True
        ' ***************************************************
        If Not SituarDataTrasEliminar(AdoAux(Index), NumRegElim, True) Then
'            PonerCampos
            
        End If
        ' *** si n'hi han tabs sense datagrid ***
        If Index = 3 Then CargaFrame 3, True
        ' ***************************************
'        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
        CalcularDatosFactura
        ' *** si n'hi han tabs ***
'        SituarTab (NumTabMto)
        ' ************************
    End If
    
    ModoLineas = 0
'    PosicionarData
    
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
    
    ModoLineas = 1 'Posem Modo Afegir Ll�nia
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Cap�alera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5
    
    ' *** bloquejar la clau primaria de la cap�alera ***
    BloquearTxt Text1(0), True
    ' **************************************************

    ' *** posar el nom del les distintes taules de ll�nies ***
    Select Case Index
        Case 0: vtabla = "tmpportesv"
    End Select
    ' ********************************************************
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 0 ' *** pose els index dels tabs de ll�nies que tenen datagrid ***
            ' *** canviar la clau primaria de les ll�nies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
'            If Index <> 4 Then ' *** els index als que no volem sugerir-li un codi ***
'                NumF = SugerirCodigoSiguienteStr(vtabla, "tmpportesv.codmacta", vWhere)
'            Else
'                NumF = ""
'            End If
            ' ***************************************************************

            AnyadirLinea DataGridAux(Index), AdoAux(Index)
    
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
            
            LLamaLineas Index, ModoLineas, anc
            
            BloquearTxt txtaux(2), False
            BloquearTxt txtaux(3), False
            
            Select Case Index
                ' *** valor per defecte a l'insertar i formateig de tots els camps ***
                Case 0 'cuentas
                    For i = 0 To txtaux.Count - 1
                        txtaux(i).Text = ""
                    Next i
                    txtaux(0).Text = vUsu.Codigo
                    txtaux(1).Text = Text1(3).Text 'tranportista
                    txtaux(4).Text = Text1(0).Text 'numfactu
                    txtaux(5).Text = Text1(1).Text 'fecfactu
                    
                    For i = 2 To 3
                        txtaux(i).Text = ""
                    Next i
                    Text2(2).Text = ""
                    PonerFoco txtaux(2)
                    
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
    PonerModo 5
    ' *** bloqueje la clau primaria de la cap�alera ***
    BloquearTxt Text1(0), True
    ' *********************************
  
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
        Case 0 'importes
            txtaux(0).Text = DataGridAux(Index).Columns(0).Text
            txtaux(1).Text = DataGridAux(Index).Columns(1).Text
            txtaux(4).Text = DataGridAux(Index).Columns(2).Text
            txtaux(5).Text = DataGridAux(Index).Columns(3).Text
            txtaux(2).Text = DataGridAux(Index).Columns(4).Text
            Text2(2).Text = DataGridAux(Index).Columns(5).Text
            txtaux(3).Text = DataGridAux(Index).Columns(6).Text
                       
            BloquearTxt txtaux(2), True
            BloquearTxt txtaux(3), False
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 0 'importes
            PonerFoco txtaux(3)
    End Select
    ' ***************************************************************************************
End Sub

Private Sub CargaGrid(Index As Integer, enlaza As Boolean)
Dim b As Boolean
Dim i As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)

    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
    
    
    Select Case Index
        Case 0 'tmpportesv
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtaux(2)|T|Cta.Contable|1400|;S|btnBuscar(1)|B|||;S|Text2(2)|T|Nombre|2800|;"
            tots = tots & "S|txtaux(3)|T|Importe|1500|;"
            arregla tots, DataGridAux(Index), Me, 350
        
            DataGridAux(0).Columns(1).Alignment = dbgLeft
            DataGridAux(0).Columns(2).Alignment = dbgRight
        
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            BloquearTxt txtaux(2), Not b
            BloquearTxt txtaux(3), Not b

            If (enlaza = True) And (Not AdoAux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
'                txtAux(2).Text = DataGridAux(Index).Columns(5).Text
'                txtAux(3).Text = DataGridAux(Index).Columns(6).Text
            Else
                txtaux(2).Text = ""
                txtaux(3).Text = ""
            End If
            
    End Select
    DataGridAux(Index).ScrollBars = dbgAutomatic
      
'    ' **** si n'hi han ll�nies en grids i camps fora d'estos ****
'    If Not AdoAux(Index).Recordset.EOF Then
'        DataGridAux_RowColChange Index, 1, 1
'    Else
'        LimpiarCamposFrame Index
'    End If
'    ' **********************************************************
      
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
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
Dim tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
        Case 0 'DESTINOS
            tabla = "tmpportesv"
            Sql = "SELECT tmpportesv.codusu, tmpportesv.codtrans, tmpportesv.numfactu, tmpportesv.fecfactu, tmpportesv.codmacta, "
            If vParamAplic.ContabilidadNueva Then
                Sql = Sql & " if(" & vParamAplic.NumeroConta & "=0,'',ariconta" & vParamAplic.NumeroConta & ".cuentas.nommacta), tmpportesv.importel"
            Else
                Sql = Sql & " if(" & vParamAplic.NumeroConta & "=0,'',conta" & vParamAplic.NumeroConta & ".cuentas.nommacta), tmpportesv.importel"
            End If
            Sql = Sql & " FROM " & tabla
            
            If vParamAplic.NumeroConta <> 0 Then
                If vParamAplic.ContabilidadNueva Then
                    Sql = Sql & ", ariconta" & vParamAplic.NumeroConta & ".cuentas "
                Else
                    Sql = Sql & ", conta" & vParamAplic.NumeroConta & ".cuentas "
                End If
            End If
            
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE codusu = -1"
            End If
            
            If vParamAplic.NumeroConta <> 0 Then
                If vParamAplic.ContabilidadNueva Then
                    Sql = Sql & " and tmpportesv.codmacta = ariconta" & vParamAplic.NumeroConta & ".cuentas.codmacta"
                Else
                    Sql = Sql & " and tmpportesv.codmacta = conta" & vParamAplic.NumeroConta & ".cuentas.codmacta"
                End If
            End If
            
            Sql = Sql & " ORDER BY " & tabla & ".codusu,  " & tabla & ".codmacta "
            
    End Select
    ' ********************************************************************************
    
    MontaSQLCarga = Sql
End Function


Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    If Text1(0).Text = "" Then Exit Sub
    Select Case Modo
        Case 5 'LL�NIES
            Select Case ModoLineas
                Case 1 'afegir ll�nia
                    InsertarLinea
                    If Not AdoAux(0).Recordset.EOF Then _
                        Me.lblIndicador.Caption = AdoAux(0).Recordset.AbsolutePosition & " de " & AdoAux(0).Recordset.RecordCount
            

                Case 2 'modificar ll�nies
                    ModificarLinea
                    PosicionarData
                    If Not AdoAux(0).Recordset.EOF Then _
                        Me.lblIndicador.Caption = AdoAux(0).Recordset.AbsolutePosition & " de " & AdoAux(0).Recordset.RecordCount
            End Select
            
        CalcularDatosFactura
        ' **************************
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub cmdCancelar_Click()
Dim i As Integer
Dim V
    
    Select Case Modo
        Case 5 'LL�NIES
            Select Case ModoLineas
                Case 1 'afegir ll�nia
                    ModoLineas = 0
                    ' *** les ll�nies que tenen datagrid (en o sense tab) ***
                    If NumTabMto = 0 Then
                        DataGridAux(NumTabMto).AllowAddNew = False
                        ' **** repasar si es diu Data1 l'adodc de la cap�alera ***
                        'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'Modificar
                        ' ********************************************************
                        LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                        DataGridAux(NumTabMto).Enabled = True
                        DataGridAux(NumTabMto).SetFocus
                        If Not AdoAux(0).Recordset.EOF Then _
                            Me.lblIndicador.Caption = AdoAux(0).Recordset.AbsolutePosition & " de " & AdoAux(0).Recordset.RecordCount

                    End If
                    
                    ' *** si n'hi han tabs ***
                    'SSTab1.Tab = 1
                    'SSTab2.Tab = NumTabMto
                    ' ************************
                    
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        AdoAux(NumTabMto).Recordset.MoveFirst
                    End If
                    PosicionarData
                    If Not AdoAux(0).Recordset.EOF Then _
                         Me.lblIndicador.Caption = AdoAux(0).Recordset.AbsolutePosition & " de " & AdoAux(0).Recordset.RecordCount

                Case 2 'modificar ll�nies
                    ModoLineas = 0
                    
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        ' *** l'Index de Fields es el que canvie de la PK de ll�nies ***
                        V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el n� de llinia
                        AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
                        ' ***************************************************************
                    End If
                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                    PosicionarData
                    If Not AdoAux(0).Recordset.EOF Then _
                         Me.lblIndicador.Caption = AdoAux(0).Recordset.AbsolutePosition & " de " & AdoAux(0).Recordset.RecordCount
                    End Select
            
'            PosicionarData
'            If Not AdoAux(0).Recordset.EOF Then _
'                 Me.lblIndicador.Caption = AdoAux(0).Recordset.AbsolutePosition & " de " & AdoAux(0).Recordset.RecordCount
            
    End Select
End Sub

Private Function SepuedeBorrar(ByRef Index As Integer) As Boolean
    SepuedeBorrar = False
    
'    ' *** si cal comprovar alguna cosa abans de borrar ***
'    Select Case Index
'        Case 0 'cuentas bancarias
'            If AdoAux(Index).Recordset!ctaprpal = 1 Then
'                MsgBox "No puede borrar una Cuenta Principal. Seleccione antes otra cuenta como Principal", vbExclamation
'                Exit Function
'            End If
'    End Select
'    ' ****************************************************
    
    SepuedeBorrar = True
End Function

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la cap�alera ***
    vWhere = vWhere & " codusu=" & Val(txtaux(0).Text)
    ' *******************************************************
    
    ObtenerWhereCab = vWhere
End Function

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
        ' *** si n'hi han tabs sense datagrids, li pose els valors als camps ***
        If (Index = 3) Then 'datos facturacion
'??monica
'            tip = AdoAux(Index).Recordset!tipclien
'            If (tip = 1) Then 'persona
'                txtAux2(27).Text = AdoAux(Index).Recordset!ape_raso & "," & AdoAux(Index).Recordset!Nom_Come
'            ElseIf (tip = 2) Then 'empresa
'                txtAux2(27).Text = AdoAux(Index).Recordset!Nom_Come
'            End If
'            txtAux2(28).Text = DBLet(AdoAux(Index).Recordset!desforpa, "T")
'            txtAux2(29).Text = DBLet(AdoAux(Index).Recordset!desrutas, "T")
'            'txtAux2(31).Text = DBLet(AdoAux(Index).Recordset!comision, "T") & " %"
'            txtAux2(32).Text = DBLet(AdoAux(Index).Recordset!nomrapel, "T")
'            'Descripcion cuentas contables de la Contabilidad
'            For i = 35 To 38
'                txtAux2(i).Text = PonerNombreDeCod(txtaux(i), "cuentas", "nommacta", "codmacta", , cConta)
'            Next i
        End If
        ' ************************************************************************
    Else
        ' *** si n'hi han tabs sense datagrids, li pose els valors als camps ***
        
    End If
End Sub

Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la cap�alera, no llevar els () ***
    Cad = "(codusu=" & vUsu.Codigo & " and codtrans = " & DBSet(Text1(3).Text, "N")
    Cad = Cad & " and numfactu = " & DBSet(Text1(0).Text, "T") & " and fecfactu = " & DBSet(Text1(1).Text, "F")
    Cad = Cad & " and codmacta = " & DBSet(AdoAux(0).Recordset!Codmacta, "T") & ")"
    ' ***************************************
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    If SituarDataMULTI(AdoAux(0), Cad, Indicador) Then
    'If SituarData(AdoAux(0), cad, Indicador, True) Then
        lblIndicador.Caption = Indicador
'    Else
'       LimpiarCampos
'       PonerModo 0
    End If
    ' ***********************************************************************************
End Sub


Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim b As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    If Index <> 3 Then DeseleccionaGrid DataGridAux(Index)
    ' ***************************************************
       
    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Ll�nies
    Select Case Index
        Case 0 'destinos
            txtaux(0).visible = False
            For jj = 1 To 3
                txtaux(jj).visible = b
                txtaux(jj).Top = alto
            Next jj
            Me.btnBuscar(1).visible = b
            Me.btnBuscar(1).Enabled = (xModo = 1)
            Me.btnBuscar(1).Top = alto
            
            Text2(2).visible = b
            Text2(2).Top = alto
            
'            txtaux2(0).visible = b
'            txtaux2(0).Top = alto
'
'            btnBuscar(0).visible = b
'            btnBuscar(0).Top = txtaux(3).Top
'            btnBuscar(0).Height = txtaux(3).Height
            
    End Select
End Sub

Private Sub InsertarLinea()
'Inserta registre en les taules de Ll�nies
Dim nomFrame As String
Dim b As Boolean

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomFrame = "FrameAux0" 'cuentas contables
    End Select
    ' ***************************************************************
    
    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomFrame) Then
            ' *** si n'hi ha que fer alguna cosa abas d'insertar
'            b = BLOQUEADesdeFormulario2(Me, Data1, 1)
            b = True
            If cadwhere <> "" Then b = BloqueaRegistro("albaran", cadwhere)
            Select Case NumTabMto
                Case 0 ' *** els index de les llinies en grid (en o sense tab) ***
                    CargaGrid NumTabMto, True
                    If b Then CalcularDatosFactura
                    BotonAnyadirLinea NumTabMto
            End Select
           
'            SituarTab (NumTabMto)
        End If
    End If
End Sub


Private Sub ModificarLinea()
'Modifica registre en les taules de Ll�nies
Dim nomFrame As String
Dim V As Integer
Dim Cad As String
    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomFrame = "FrameAux0" 'cuentas Bancarias
    End Select
    ' **************************************************************

    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomFrame) Then
'??monica
'            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
            If cadwhere <> "" Then BloqueaRegistro "albaran", cadwhere

            ModoLineas = 0

            If NumTabMto <> 3 Then
                V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el n� de llinia
                CargaGrid NumTabMto, True
            End If

            ' *** si n'hi han tabs ***
'            SituarTab (NumTabMto)

            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
            If NumTabMto <> 3 Then
                DataGridAux(NumTabMto).SetFocus
                AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
            End If
            ' ***********************************************************

            LLamaLineas NumTabMto, 0
        End If
    End If
        
'    'Cridem al form
'    ' **************** arreglar-ho per a vore lo que es desije ****************
'    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
'    Cad = ""
'    Cad = Cad & ParaGrid(text1(0), 15, "C�d.")
'    Cad = Cad & ParaGrid(text1(2), 60, "Nombre")
'    Cad = Cad & ParaGrid(text1(3), 25, "N.I.F.")
'    If Cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = Cad
'        frmB.vtabla = NombreTabla
'        frmB.vSQL = CadB
'        HaDevueltoDatos = False
'        frmB.vDevuelve = "0|1|2|" '*** els camps que volen que torne ***
'        frmB.vTitulo = "Clientes" ' ***** repasa a��: t�tol de BuscaGrid *****
'        frmB.vSelElem = 1
'
'        frmB.Show vbModal
'        Set frmB = Nothing
'        'Si ha posat valors i tenim que es formulari de b�squeda llavors
'        'tindrem que tancar el form llan�ant l'event
'        If HaDevueltoDatos Then
'            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                cmdRegresar_Click
'        Else   'de ha retornat datos, es a decir NO ha retornat datos
'            PonerFoco text1(kCampo)
'        End If
'    End If
End Sub

Private Sub PonerCampos()
Dim i As Integer
Dim codpobla As String, despobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la cap�alera
    
    ' *** si n'hi han ll�nies en datagrids ***
    'For i = 0 To DataGridAux.Count - 1
'        If i <> 3 Then
            CargaGrid i, True
            If Not AdoAux(i).Recordset.EOF Then _
                PonerCamposForma2 Me, AdoAux(i), 2, "FrameAux" & i
'        End If
    ' *******************************************

    ' *** si n'hi han ll�nies sense datagrid ***
'    CargaFrame 3, True
    ' ***************************************
    
    ' ************* configurar els camps de les descripcions de la cap�alera *************
'    txtAux2(22).Text = PonerNombreDeCod(txtAux(22), "poblacio", "despobla", "codpobla", "N")

'    PosarDescripcions

'    codPobla = DBLet(Data1.Recordset!codPobla, "T")
'    DatosPoblacion codPobla, desPobla, CPostal, desProvi, desPais
'    text1(5).Text = codPobla 'Devuelve el campo formateado
'    text2(5).Text = desPobla
''    text1(8).Text = CPostal
'    text2(1).Text = desProvi
'    text2(2).Text = desPais
'
'    text2(7).Text = PonerNombreDeCod(text1(7), "activida", "desactiv")
'    text2(8).Text = PonerNombreDeCod(text1(8), "grupempr", "desgrupo", "codgrupo", "N")
    ' ********************************************************************************
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
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
    
    If ModoLineas = 1 Then
        'comprobar si existe ya el cod. del campo clave primaria
        Sql = "select * from tmpportesv where codusu = " & vUsu.Codigo & " and codtrans = " & DBSet(Text1(3).Text, "N")
        Sql = Sql & " and numfactu = " & DBSet(Text1(0).Text, "T") & " and fecfactu = " & DBSet(Text1(1).Text, "F")
        Sql = Sql & " and codmacta = " & DBSet(txtaux(2).Text, "T")
        If RegistrosAListar(Sql, cAgro) = 1 Then
            MsgBox "Ya existe la cuenta contable para la factura. Modifique.", vbExclamation
            b = False
        End If
    End If
    
    DatosOkLlin = b

EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub PonerModoOpcionesMenu(Modo)
'Actives unes Opcions de Men� i Toolbar seg�n el modo en que estem
Dim b As Boolean, bAux As Boolean
Dim i As Byte
    
    ' *** si n'hi han ll�nies que tenen grids (en o sense tab) ***
    b = (Modo = 3 Or Modo = 4 Or Modo = 2)
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.AdoAux(i).Recordset.RecordCount > 0)
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i
    ' ****************************************
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFoco txtaux(Index), Modo
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index <> 0 And KeyCode <> 38 Then KEYdown KeyCode
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim TipoDto As Byte
Dim ImpDto As String
Dim Unidades As String
Dim Cantidad As String

    'Quitar espacios en blanco
    If Not PerderFocoGnralLineas(txtaux(Index), ModoLineas) Then Exit Sub
    
    Select Case Index
        Case 2 'Cta Contable
            If txtaux(Index).Text = "" Then Exit Sub
            Text2(2).Text = PonerNombreCuenta(txtaux(Index), Modo)
            If Text2(2).Text = "" Then
                PonerFoco txtaux(Index)
            End If
            
        Case 3 ' Importe
            If txtaux(Index).Text <> "" Then PonerFormatoDecimal txtaux(Index), 1
            
    End Select
End Sub



Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim i As Byte
    
    Combo1(0).Clear
    
    Combo1(0).AddItem "Transportista"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    
    Combo1(0).AddItem "Comisionista"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    
End Sub


Private Sub RecalcularDatosFactura()
Dim i As Integer
Dim Sql As String
Dim cadAux As String
Dim TotalFactura As Currency
    
Dim ImpBImIVA As Currency
Dim impiva As Currency
Dim ImpIVA1 As Currency
Dim ImpIVA2 As Currency
Dim ImpIVA3 As Currency
    
Dim BaseReten As Currency
Dim PorcReten As Currency
Dim ImpReten As Currency

    On Error GoTo eRecalcularDatosFactura
    
    '[Monica]23/04/2019: la tabla de transportista es albaran_transportista
    If Not SeleccionaRegistros(Combo1(0).ListIndex) Then Exit Sub
    
    If Combo1(0).ListIndex = 0 Then
        If Not BloqueaRegistro("albaran left join albaran_transporte on albaran.numalbar = albaran_transporte.numalbar", cadwhere) Then
            ListView1.SelectedItem.Checked = False
        End If
    Else
        If Not BloqueaRegistro("albaran", cadwhere) Then
            ListView1.SelectedItem.Checked = False
        End If
    End If
    TotalFactura = 0
    If Text1(16).Text <> "" Then
        cadAux = Text1(13).Text
        ImpBImIVA = CCur(ImporteSinFormato(Text1(16).Text))
        If cadAux = "" Then cadAux = "0"
        impiva = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
        
        ImpIVA1 = impiva
        
        
        'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
        'los vamos acumulando
        TotalFactura = TotalFactura + ImpBImIVA + impiva
    End If
    
    If Text1(17).Text <> "" Then
        cadAux = Text1(14).Text
        ImpBImIVA = CCur(ImporteSinFormato(Text1(17).Text))
        If cadAux = "" Then cadAux = "0"
        impiva = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
        
        ImpIVA2 = impiva
        
        
        'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
        'los vamos acumulando
        TotalFactura = TotalFactura + ImpBImIVA + impiva
    End If
    
    
    If Text1(18).Text <> "" Then
        cadAux = Text1(15).Text
        ImpBImIVA = CCur(ImporteSinFormato(Text1(18).Text))
        If cadAux = "" Then cadAux = "0"
        impiva = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
        
        ImpIVA3 = impiva
        
        
        'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
        'los vamos acumulando
        TotalFactura = TotalFactura + ImpBImIVA + impiva
    End If
        
'        Text1(6).Text = vFactu.BrutoFac
'        Text1(7).Text = vFactu.ImpPPago
'        Text1(8).Text = vFactu.ImpGnral
'        Text1(9).Text = vFactu.BaseImp
'        Text1(10).Text = vFactu.TipoIVA1
'        Text1(11).Text = vFactu.TipoIVA2
'        Text1(12).Text = vFactu.TipoIVA3
'        Text1(13).Text = vFactu.PorceIVA1
'        Text1(14).Text = vFactu.PorceIVA2
'        Text1(15).Text = vFactu.PorceIVA3
'        Text1(16).Text = vFactu.BaseIVA1
'        Text1(17).Text = vFactu.BaseIVA2
'        Text1(18).Text = vFactu.BaseIVA3
        
        
        Text1(19).Text = ImpIVA1
        Text1(20).Text = ImpIVA2
        Text1(21).Text = ImpIVA3
        
        
    If Text1(24).Text <> "" Then
        PorcReten = CCur(ImporteSinFormato(Text1(24).Text))
        BaseReten = 0
        If Text1(23).Text <> "" Then BaseReten = CCur(ImporteSinFormato(Text1(23).Text))
        
        ImpReten = CalcularPorcentaje(BaseReten, PorcReten, 2)
    
        Text1(25).Text = Format(ImpReten, "###,###,##0.00")
    
        'TOTAL de la factura
        TotalFactura = TotalFactura - ImpReten
    End If
    
        
        Text1(22).Text = TotalFactura
        
        For i = 19 To 22
            FormateaCampo Text1(i)
        Next i
        'Quitar ceros de linea IVA 2
        If Val(Text1(14).Text) = 0 And Val(Text1(11).Text) = 0 Then
            For i = 11 To 20 Step 3
                Text1(i).Text = ""
            Next i
        End If
        'Quitar ceros de linea IVA 3
        If Val(Text1(15).Text) = 0 And Val(Text1(12).Text) = 0 Then
            For i = 12 To 21 Step 3
                Text1(i).Text = ""
            Next i
        End If
        Exit Sub
        
   
eRecalcularDatosFactura:
    MuestraError Err.Number, "Recalculando Datos de Factura", Err.Description
End Sub



