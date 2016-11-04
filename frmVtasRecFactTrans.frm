VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmVtasRecFactTrans 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturas de Transporte / Comisión"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   12315
   Icon            =   "frmVtasRecFactTrans.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmVtasRecFactTrans.frx":000C
   ScaleHeight     =   7110
   ScaleWidth      =   12315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4740
      Left            =   45
      TabIndex        =   61
      Top             =   2070
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   8361
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Albaranes"
      TabPicture(0)   =   "frmVtasRecFactTrans.frx":0A0E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Portes de Vuelta"
      TabPicture(1)   =   "frmVtasRecFactTrans.frx":0A2A
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "FrameAux0"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame FrameAux0 
         BorderStyle     =   0  'None
         Height          =   4200
         Left            =   90
         TabIndex        =   63
         Top             =   405
         Width           =   6540
         Begin VB.TextBox txtaux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   5
            Left            =   1170
            MaxLength       =   35
            TabIndex        =   77
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
            TabIndex        =   76
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
            TabIndex        =   75
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
            TabIndex        =   74
            Text            =   "Text2"
            Top             =   3150
            Width           =   2520
         End
         Begin VB.Frame Frame1 
            Height          =   555
            Index           =   0
            Left            =   180
            TabIndex        =   72
            Top             =   3510
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
               TabIndex        =   73
               Top             =   180
               Width           =   2655
            End
         End
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "&Aceptar"
            Height          =   375
            Left            =   4095
            TabIndex        =   68
            Top             =   3600
            Width           =   1035
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "&Cancelar"
            Height          =   375
            Left            =   5295
            TabIndex        =   69
            Top             =   3600
            Width           =   1035
         End
         Begin VB.TextBox txtaux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   3
            Left            =   5085
            MaxLength       =   35
            TabIndex        =   67
            Tag             =   "Importe|N|N|||tmpportesv|importel|###,###0.00||"
            Text            =   "Importe"
            Top             =   3150
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtaux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   2
            Left            =   1650
            MaxLength       =   40
            TabIndex        =   66
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
            TabIndex        =   65
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
            TabIndex        =   64
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
            TabIndex        =   70
            Top             =   225
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
            Bindings        =   "frmVtasRecFactTrans.frx":0A46
            Height          =   2745
            Index           =   0
            Left            =   135
            TabIndex        =   71
            Top             =   675
            Width           =   6180
            _ExtentX        =   10901
            _ExtentY        =   4842
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
      Begin MSComctlLib.ListView ListView1 
         Height          =   4095
         Left            =   -74955
         TabIndex        =   62
         Top             =   450
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   7223
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame FrameIntro 
      Height          =   1550
      Left            =   135
      TabIndex        =   8
      Top             =   495
      Width           =   12045
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Tag             =   "Tipo|N|N|||tcafpc|tipofact|||"
         Top             =   405
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   26
         Left            =   6930
         MaxLength       =   6
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   435
         Width           =   660
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   26
         Left            =   7650
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   59
         Text            =   "Text2"
         Top             =   435
         Width           =   3930
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Contabiliz."
         Height          =   375
         Index           =   1
         Left            =   5535
         TabIndex        =   48
         Top             =   1080
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Tesoreria"
         Height          =   375
         Index           =   0
         Left            =   5535
         TabIndex        =   47
         Top             =   720
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   1530
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   15
         Text            =   "Text2"
         Top             =   1000
         Width           =   4095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   4365
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Recepción|F|N|||tcafpc|fecrecep|dd/mm/yyyy|N|"
         Top             =   400
         Width           =   1305
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   5
         Left            =   7635
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   13
         Text            =   "Text2"
         Top             =   1000
         Width           =   3930
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   6915
         MaxLength       =   5
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1000
         Width           =   660
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   450
         MaxLength       =   3
         TabIndex        =   4
         Tag             =   "Cod. Transportista|N|N|0|999|tcafpc|codtrans|000|S|"
         Text            =   "Text1"
         Top             =   1000
         Width           =   960
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   2925
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Factura|F|N|||tcafpc|fecfactu|dd/mm/yyyy|S|"
         Top             =   405
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   1545
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Nº Factura|T|N|||tcafpc|numfactu||S|"
         Text            =   "Text1 7"
         Top             =   400
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo"
         Height          =   255
         Index           =   17
         Left            =   135
         TabIndex        =   78
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   210
         Index           =   16
         Left            =   6615
         TabIndex        =   60
         Top             =   225
         Width           =   1215
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   6615
         ToolTipText     =   "Buscar cliente"
         Top             =   450
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   5400
         Picture         =   "frmVtasRecFactTrans.frx":0A5E
         ToolTipText     =   "Buscar fecha"
         Top             =   150
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   3960
         Picture         =   "frmVtasRecFactTrans.frx":0AE9
         ToolTipText     =   "Buscar fecha"
         Top             =   135
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   6600
         ToolTipText     =   "Buscar banco propio"
         Top             =   1035
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   135
         ToolTipText     =   "Buscar transportista"
         Top             =   1035
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Recep."
         Height          =   255
         Index           =   3
         Left            =   4365
         TabIndex        =   14
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Cta. Prev. Pago"
         Height          =   255
         Index           =   2
         Left            =   6600
         TabIndex        =   12
         Top             =   795
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Transportista / Comisionista"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   11
         Top             =   810
         Width           =   2160
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Factura"
         Height          =   255
         Index           =   29
         Left            =   2910
         TabIndex        =   10
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Factura"
         Height          =   255
         Index           =   28
         Left            =   1545
         TabIndex        =   9
         Top             =   195
         Width           =   1095
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   2160
      MaxLength       =   4
      TabIndex        =   50
      Text            =   "Text1"
      Top             =   1110
      Width           =   660
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   4
      Left            =   2880
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   49
      Text            =   "Text2"
      Top             =   1110
      Width           =   3615
   End
   Begin VB.Frame FrameFactura 
      Height          =   4830
      Left            =   7470
      TabIndex        =   16
      Top             =   2025
      Width           =   4695
      Begin VB.CommandButton CmdCan 
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   120
         TabIndex        =   81
         Top             =   4320
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "&Generar"
         Height          =   375
         Left            =   1290
         TabIndex        =   80
         Top             =   4320
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   25
         Left            =   3015
         MaxLength       =   15
         TabIndex        =   54
         Tag             =   "Importe Retencion|N|N|0||scafac|imporete|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   3750
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   24
         Left            =   720
         MaxLength       =   5
         TabIndex        =   53
         Tag             =   "% reten|N|S|0|99.90|scafac|porcereten|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   3750
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   23
         Left            =   1335
         MaxLength       =   15
         TabIndex        =   52
         Tag             =   "Base Imponible 1|N|N|0||scafac|baseimp1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   3735
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   9
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   41
         Tag             =   "Importe IVA 1|N|N|0||scafac|imporiv1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   1350
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   40
         Tag             =   "Base Imponible 3|N|N|0||scafac|baseimp3|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   900
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   39
         Tag             =   "Base Imponible 2 |N|N|0||scafac|baseimp2|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   570
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   3015
         MaxLength       =   15
         TabIndex        =   37
         Tag             =   "Base Imponible 1|N|N|0||scafac|baseimp1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   180
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   360
         MaxLength       =   5
         TabIndex        =   35
         Tag             =   "% IVA 3|N|S|0|99|scafac|porciva3|00|N|"
         Text            =   "Text1 7"
         Top             =   2925
         Width           =   500
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   360
         MaxLength       =   5
         TabIndex        =   34
         Tag             =   "& IVA 2|N|S|0|99|scafac|porciva2|00|N|"
         Text            =   "Text1 7"
         Top             =   2595
         Width           =   500
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   360
         MaxLength       =   5
         TabIndex        =   33
         Tag             =   "% IVA 1|N|S|0|99|scafac|porciva1|00|N|"
         Text            =   "Text1 7"
         Top             =   2250
         Width           =   500
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   26
         Tag             =   "Base Imponible 1|N|N|0||scafac|baseimp1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2280
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   885
         MaxLength       =   5
         TabIndex        =   25
         Tag             =   "% IVA 1|N|S|0|99.90|scafac|porciva1|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   2250
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   19
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   24
         Tag             =   "Importe IVA 1|N|N|0||scafac|imporiv1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2280
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   23
         Tag             =   "Base Imponible 2 |N|N|0||scafac|baseimp2|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2595
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   14
         Left            =   885
         MaxLength       =   5
         TabIndex        =   22
         Tag             =   "& IVA 2|N|S|0|99.90|scafac|porciva2|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   2595
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   20
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   21
         Tag             =   "Importe IVA 2|N|N|0||scafac|imporiv2|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2595
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   20
         Tag             =   "Base Imponible 3|N|N|0||scafac|baseimp3|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2925
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   15
         Left            =   900
         MaxLength       =   5
         TabIndex        =   19
         Tag             =   "% IVA 3|N|S|0|99.90|scafac|porciva3|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   2925
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   21
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   18
         Tag             =   "Importe IVA 3|N|N|0||scafac|imporiv3|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2925
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   22
         Left            =   2460
         MaxLength       =   15
         TabIndex        =   17
         Tag             =   "Total Factura|N|N|0||scafac|totalfac|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   4395
         Width           =   2055
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   90
         ToolTipText     =   "Buscar codigo iva"
         Top             =   2940
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   90
         ToolTipText     =   "Buscar codigo iva"
         Top             =   2610
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   90
         ToolTipText     =   "Buscar codigo iva"
         Top             =   2250
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "% Ret"
         Height          =   255
         Index           =   15
         Left            =   720
         TabIndex        =   58
         Top             =   3555
         Width           =   495
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
         Index           =   14
         Left            =   3615
         TabIndex        =   57
         Top             =   3150
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Retención"
         Height          =   255
         Index           =   13
         Left            =   3015
         TabIndex        =   56
         Top             =   3540
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Base Imponible"
         Height          =   210
         Index           =   12
         Left            =   1335
         TabIndex        =   55
         Top             =   3555
         Width           =   1215
      End
      Begin VB.Line Line3 
         X1              =   135
         X2              =   2895
         Y1              =   3390
         Y2              =   3390
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
         Left            =   2760
         TabIndex        =   46
         Top             =   900
         Width           =   135
      End
      Begin VB.Line Line2 
         X1              =   2520
         X2              =   4550
         Y1              =   1250
         Y2              =   1250
      End
      Begin VB.Label Label1 
         Caption         =   "Imp. dto. gnral."
         Height          =   255
         Index           =   10
         Left            =   1440
         TabIndex        =   45
         Top             =   900
         Width           =   1215
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
         Left            =   2760
         TabIndex        =   44
         Top             =   570
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "Imp. dto. ppago"
         Height          =   255
         Index           =   8
         Left            =   1440
         TabIndex        =   43
         Top             =   570
         Width           =   1215
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   2880
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label1 
         Caption         =   "Base Imponible"
         Height          =   255
         Index           =   7
         Left            =   1440
         TabIndex        =   42
         Top             =   1350
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Bruto Factura"
         Height          =   255
         Index           =   6
         Left            =   1440
         TabIndex        =   38
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Cod."
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   36
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Base Imponible"
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   32
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Importe IVA"
         Height          =   255
         Index           =   33
         Left            =   3000
         TabIndex        =   31
         Top             =   2070
         Width           =   1335
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
         Left            =   3600
         TabIndex        =   30
         Top             =   1680
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
         TabIndex        =   29
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "TOTAL FACTURA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   39
         Left            =   2970
         TabIndex        =   28
         Top             =   4125
         Width           =   1530
      End
      Begin VB.Label Label1 
         Caption         =   "% IVA"
         Height          =   255
         Index           =   41
         Left            =   885
         TabIndex        =   27
         Top             =   2070
         Width           =   495
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pedir Datos"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Ver Albaranes"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Generar Factura"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8400
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Operador"
      Height          =   255
      Index           =   1
      Left            =   1845
      TabIndex        =   51
      Top             =   900
      Width           =   735
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   1845
      Picture         =   "frmVtasRecFactTrans.frx":0B74
      ToolTipText     =   "Buscar trabajador"
      Top             =   1125
      Width           =   240
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnPedirDatos 
         Caption         =   "&Pedir Datos"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnVerAlbaran 
         Caption         =   "&Ver Albaranes"
         Enabled         =   0   'False
         Shortcut        =   ^A
         Visible         =   0   'False
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
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
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
Dim cadWhere As String

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
'1.- Afegir,  2.- Modificar,  3.- Borrar,  0.-Passar control a Llínies
Dim NumTabMto As Integer 'Indica quin nº de Tab està en modo Mantenimient

Dim vWhere As String



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
Dim I As Integer

    FrameIntro.Enabled = False
    SSTab1.Enabled = True
    FrameFactura.Enabled = False
    
    
    BloquearTxt Text1(10), True
    BloquearTxt Text1(11), True
    BloquearTxt Text1(12), True
    
    
    For I = 4 To 6
        imgBuscar(I).Enabled = False
        imgBuscar(I).visible = False
    Next I
    
    
    Me.CmdCan.Enabled = False
    Me.CmdCan.visible = False
    Me.cmdGenerar.Enabled = False
    Me.cmdGenerar.visible = False


End Sub

Private Sub cmdGenerar_Click()
Dim I As Integer

    FrameIntro.Enabled = False
    SSTab1.Enabled = True
    FrameFactura.Enabled = False


    For I = 4 To 6
        imgBuscar(I).Enabled = False
        imgBuscar(I).visible = False
    Next I
    
    
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
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
'    If VerAlbaranes Then RefrescarAlbaranes
'    VerAlbaranes = False
End Sub

Private Sub Form_Load()
Dim I As Integer
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Pedir Datos
        .Buttons(2).Image = 3   'Ver albaranes
        .Buttons(3).Image = 15   'Generar FActura
        .Buttons(6).Image = 11   'Salir
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
    Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    For I = 2 To 6
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
    
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
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
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
    txtAux(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cuenta
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmF_Selec(vFecha As Date)
Dim indice As Byte
    indice = CByte(Me.imgFecha(0).Tag)
    Text1(indice).Text = Format(vFecha, "dd/mm/yyyy")
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
    
   menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

   frmF.Left = esq + imgFecha(Index).Parent.Left + 30
   frmF.Top = dalt + imgFecha(Index).Parent.Top + imgFecha(Index).Height + menu - 40
   
   frmF.NovaData = Now
   indice = Index + 1
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
Dim I As Long
Dim b As Boolean

    If Modo <> 5 Then Exit Sub

    If Bloquear = True Then
        ListView1.SetFocus
        item.EnsureVisible
        Exit Sub
    End If
    
    Select Case Combo1(0).ListIndex
        Case 0 ' transportista
            Cantidad = InputBox("Introduzca el Porte para el albarán:" & item.Text, "Portes", , 5000, 4000)
            
        Case 1 ' comisionista
            Cantidad = InputBox("Introduzca la Comisión para el albarán:" & item.Text, "Comisión", , 5000, 4000)
    
    End Select
    
    If Cantidad = "" Then Exit Sub
    b = True
    If vParamAplic.PortesKiloCaja = 1 And Combo1(0).ListIndex = 0 Then
        palets = InputBox("Introduzca el Número de Palets para el albarán:" & item.Text, "Palets", , 5000, 4000)
        
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
                If vParamAplic.PortesKiloCaja = 1 Then
                    Sql = Sql & ", paletspag = " & DBSet(CStr(valor2), "N")
                End If
                Sql = Sql & " where numalbar = " & DBSet(item.Text, "N")
                
            Case 1 ' factura de comisionista
                Sql = "update albaran set comisionespag = " & DBSet(CStr(Valor), "N")
                Sql = Sql & " where numalbar = " & DBSet(item.Text, "N")
                
        End Select
        conn.Execute Sql
        
        I = ListView1.SelectedItem.Index
        
        CargarAlbaranes vWhere, Combo1(0).ListIndex
        CalcularDatosFactura
        
        ' Crea una variable ListItem.
        ' Establece la variable al elemento encontrado.
        If I < ListView1.ListItems.Count Then
            ListView1.SelectedItem = ListView1.ListItems.item(I + 1)
        Else
            ListView1.SelectedItem = ListView1.ListItems.item(I)
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
Dim I As Integer

    FrameIntro.Enabled = False
    SSTab1.Enabled = False
    FrameFactura.Enabled = True
    
    For I = 6 To 22
        BloquearTxt Text1(I), True
    Next I
    
    For I = 6 To 22
        Text1(I).Enabled = False
    Next I
    
    
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
                ' No debe existir el número de factura para el proveedor en hco
                If ExisteFacturaEnHco Then
                    InicializarListView
                End If
            End If
            
        Case 3 'Cod Transportista
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "agencias", "nomtrans", "codtrans")
                ' No debe existir el número de factura para el proveedor en hco
                If ExisteFacturaEnHco Then
                    InicializarListView
                Else
                    'comprobamos que no haya nadie recepcionando facturas de ese proveedor
'                    DesBloqueoManual ("FACTRA")
'                    If Not BloqueoManual("FACTRA", Text1(3).Text) Then
                    Select Case Combo1(0).ListIndex
                        Case 0 ' transportista
                            vWhere = "codtrans = " & DBSet(Text1(3).Text, "N")
                        Case 1 ' comisionista
                            vWhere = "codcomis = " & DBSet(Text1(3).Text, "N")
                    End Select
                    If Text1(26).Text <> "" Then vWhere = vWhere & " and codclien = " & DBSet(Text1(26).Text, "N")
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
                        Select Case Combo1(0).ListIndex
                            Case 0
                                Sql3 = "update albaran set portespag = 0 where " & vWhere
                                conn.Execute Sql3
                            Case 1
                                Sql3 = "update albaran set comisionespag = 0 where " & vWhere
                                conn.Execute Sql3
                        End Select
                        
                        
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
            If Not ExisteFacturaEnHco Then
                PonerModo 5
                Me.ListView1.SetFocus
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


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte, NumReg As Byte
Dim b As Boolean
On Error GoTo EPonerModo

    Modo = Kmodo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
        
                 
'    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
'    'Si estamos en Insertar además limpia los campos Text1
'    'si estamos en modificar bloquea las compos que son clave primaria
'    BloquearText1 Me, Modo
    
    For I = 0 To Text1.Count - 1
        BloquearTxt Text1(I), (Modo <> 3)
    Next I
    
    'Importes siempre bloqueados
    For I = 6 To 25
        BloquearTxt Text1(I), True
    Next I
    
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
    
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Enabled = b
    Next I
                    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
    
    For I = 0 To txtAux.Count - 1
        BloquearTxt txtAux(I), True
        txtAux(I).visible = False
    Next I
        
    Me.FrameIntro.Enabled = (Modo = 3)
    Me.FrameAux0.Enabled = (Modo = 5)
       
    BloquearBtn btnBuscar(1), True
    Text2(2).visible = False
       
       
       
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
'    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
'Comprobar que los datos del frame de introduccion son correctos antes de cargar datos
Dim vtag As CTag
Dim cad As String
Dim I As Byte

    On Error GoTo EDatosOK
    DatosOk = False
    
    ' deben de introducirse todos los datos del frame
    For I = 0 To 5
        If Text1(I).Text = "" And I <> 4 Then
            If Text1(I).Tag <> "" Then
                Set vtag = New CTag
                If vtag.Cargar(Text1(I)) Then
                    cad = vtag.Nombre
                Else
                    cad = "Campo"
                End If
                Set vtag = Nothing
            Else
                cad = "Campo"
                If I = 5 Then cad = "Cta. Prev. Pago"
            End If
            MsgBox cad & " no puede estar vacio. Reintroduzca", vbExclamation
            PonerModo 3
            PonerFoco Text1(I)
            Exit Function
        End If
    Next I
        
    'comprobar que la fecha de la factura sea anterior a la fecha de recepcion
    If Not EsFechaIgualPosterior(Text1(1).Text, Text1(2).Text, True, "La fecha de recepción debe ser igual o posterior a la fecha de la factura.") Then
        Exit Function
    End If
    
    'Comprobar que la fecha de RECEPCION esta dentro de los ejercicios contables
    If vParamAplic.NumeroConta <> 0 Then
        I = EsFechaOKConta(CDate(Text1(2).Text))
        If I > 0 Then
            'If i = 1 Then
                MsgBox "Fecha fuera ejercicios contables", vbExclamation
                Exit Function
           ' Else
           '     cad = "La fecha es superior al ejercico contable siguiente. ¿Desea continuar?"
           '     If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
           ' End If
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
    If cadWhere = "" Then
        If AdoAux(0).Recordset.RecordCount = 0 Then
            MsgBox "No hay albaranes ni portes de vuelta para incluir en la factura. Revise.", vbExclamation
            Exit Function
        End If
    End If
    
    ' No debe existir el número de factura para el transportista en hco
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
    cad = "Select tipoforp from forpago where codforpa=" & vTrans.ForPago 'cad
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    If miRsAux.EOF Then
        MsgBox "Error en el TIPO de forma de pago", vbExclamation
    Else
        I = 1
        cad = miRsAux.Fields(0)
        If Val(cad) = vbFPTransferencia Then
            'Compruebo que la forpa es transferencia
            I = 2
        End If
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    
    
    If I = 2 Then
        'La forma de pago es transferencia. Debo comprobar que existe la cuenta bancaria
        'del proveedor
        If vTrans.CuentaBan = "" Or vTrans.DigControl = "" Or vTrans.Sucursal = "" Or vTrans.Banco = "" Then
            cad = "Cuenta bancaria incorrecta. Forma de pago: transferencia.    ¿Continuar?"
            If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then I = 0
        End If
    End If
    
    'Si i=0 es que o esta mal la forpa o no quiere seguir pq no tiene cuenta bancaria
    If I > 0 Then DatosOk = True
    Exit Function
    
EDatosOK:
    DatosOk = False
    MuestraError Err.Number, "Comprobar datos correctos", Err.Description
End Function



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Pedir datos
             mnPedirDatos_Click
            
        Case 3 'Generar Factura
            mnGenerarFac_Click

        Case 6    'Salir
            mnSalir_Click
    End Select
End Sub


Private Sub PonerOpcionesMenu()
Dim j As Byte

    PonerOpcionesMenuGeneral Me
    
    j = Val(Me.mnPedirDatos.HelpContextID)
    If j < vUsu.Nivel Then Me.mnPedirDatos.Enabled = False
    
    j = Val(Me.mnGenerarFac.HelpContextID)
    If j < vUsu.Nivel Then Me.mnGenerarFac.Enabled = False
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

 
Private Sub BotonPedirDatos()
Dim Nombre As String

    TerminaBloquear

    FrameIntro.Enabled = True
    SSTab1.Enabled = False
    FrameFactura.Enabled = False


    'Vaciamos todos los Text
    LimpiarCampos
    'Vaciamos el ListView
    InicializarListView
    CargaGrid 0, False
    
    InicializarTemporal
    SSTab1.Tab = 0
    'Como no habrá albaranes seleccionados vaciamos la cadwhere
    cadWhere = ""
    
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

'Private Sub BotonVerAlbaranes()
'
'    If Not SeleccionaRegistros Then Exit Sub
'
'    VerAlbaranes = True
'
'    frmVtasAlbaranes.cadSelAlbaranes = cadWHERE
'    frmVtasAlbaranes.Show vbModal
'    frmVtasAlbaranes.cadSelAlbaranes = ""
'End Sub
'


Private Sub CargarAlbaranes(cadWhere As String, tipo As Byte)
'Tipo = 0 transportista
'Tipo = 1 comisionista
'Recupera de la BD y muestra en el Listview todos los albaranes de compra
'que tiene el proveedor introducido.
Dim Sql As String, Sql3 As String
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
On Error GoTo ECargar

    ListView1.ListItems.Clear
'    If VerAlbaranes = False Then cadWHERE = ""
    
    'si no hay trasportista/comisionista salir
    If Text1(3).Text = "" Then Exit Sub
    
    Select Case tipo
        Case 0 ' factura de transportista
            If vParamAplic.PortesKiloCaja = 0 Then
'                SQL = "SELECT albaran.numalbar,albaran.fechaalb,albaran.matriveh,albaran.matrirem,albaran.portespag "
                Sql = "SELECT albaran.numalbar,albaran.fechaalb,albaran.matriveh,albaran.matrirem,sum(albaran_costes.impcoste) as portesac ,albaran.portespag "
            Else                                                                'monica:080908 visualizaba el totpalet
'                SQL = "SELECT albaran.numalbar,albaran.fechaalb,albaran.matriveh,albaran.paletspag as totpalet, albaran.portespag "
                Sql = "SELECT albaran.numalbar,albaran.fechaalb,albaran.matriveh,albaran.paletspag as totpalet, sum(albaran_costes.impcoste) as portesac, albaran.portespag "
            End If
            Sql = Sql & " FROM albaran LEFT JOIN albaran_costes ON albaran.numalbar = albaran_costes.numalbar and albaran_costes.tipogasto = 2 "
        '    SQL = SQL & " FROM albaran INNER JOIN clientes ON albaran.codclien=clientes.codclien "
            Sql = Sql & " WHERE " & cadWhere '& Text1(3).Text
            
'[Monica] 04/01/2010 : un mismo albaran puede ser recepcionado varias veces por lo que el coste irá aumentando
'            SQL = SQL & " and albaran.numalbar not in (select numalbar from albaran_costes where tipogasto = 2) "
'[Monica] 04/01/2010 : los albaranes irán ordenados por el importe pagado (primero los que no han sido pagados
            Sql = Sql & " group by 1,2,3,4,6 "
            Sql = Sql & " order by portesac asc, albaran.numalbar , albaran.fechaalb "
        
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
            InicializarListView
            
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
            Sql = Sql & " WHERE " & cadWhere
            'SQL = SQL & " and albaran.numalbar not in (select numalbar from albaran_costes where tipogasto = 2) "
            Sql = Sql & " group by 1,2,3,4,6"
            Sql = Sql & " order by comisionesac, numalbar, fechaalb "
        
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
            InicializarListView
            
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
End Sub


Private Sub InicializarListView()
'Inicializa las columnas del List view

    ListView1.ListItems.Clear
    
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Albaran", 900 ' 1100
    ListView1.ColumnHeaders.Add , , "Fecha", 1100, 2 '1270, 2
    ListView1.ColumnHeaders.Add , , "Mat.Vehiculo", 1300 '"Mat.Vehículo", 1300
    
    Select Case Combo1(0).ListIndex
        Case 0
            If vParamAplic.PortesKiloCaja = 0 Then
                ListView1.ColumnHeaders.Add , , "Mat.Remolque", 1300 ' "Mat.Remolque", 1300
            Else
                ListView1.ColumnHeaders.Add , , "Nro.Palets", 1300 '1300
            End If
    
            ListView1.ColumnHeaders.Add , , "Portes Ac.", 1100, 1
            ListView1.ColumnHeaders.Add , , "Portes", 1100, 1 '1300, 1
            
        Case 1
            ListView1.ColumnHeaders.Add , , "Mat.Remolque", 1300 '"Mat.Remolque", 1300
            
            ListView1.ColumnHeaders.Add , , "Comisión Ac.", 1100, 1
            ListView1.ColumnHeaders.Add , , "Comisión", 1100, 1 ' 1300, 1

    End Select
End Sub

Private Sub InicializarTemporal()
'Inicializa la tmpinformes
Dim Sql As String
    
    Sql = "delete from tmpportesv where codusu = " & vUsu.codigo
    conn.Execute Sql
    
End Sub


Private Sub CalcularDatosFactura()
Dim I As Integer
Dim Sql As String
Dim cadAux As String
Dim ImpBruto As Currency
Dim ImpIVA As Currency
Dim vFactu As CFacturaTra
Dim Rs As ADODB.Recordset

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For I = 6 To 25
         Text1(I).Text = ""
    Next I


    cadAux = ""
    cadWhere = ""
    ImpBruto = 0
    
    For I = 1 To ListView1.ListItems.Count
        If DBLet(ListView1.ListItems(I).SubItems(5), "N") <> 0 Then
        'para cada albaran que tenga importe de portes
            ImpBruto = ImpBruto + DBSet(ListView1.ListItems(I).SubItems(5), "N")
            
            Sql = "(albaran.numalbar=" & DBSet(ListView1.ListItems(I).Text, "N") & ") "
            If cadAux = "" Then
                cadAux = Sql
            Else
                cadAux = cadAux & " OR " & Sql
            End If
        End If
    Next I
    
    Sql = "select sum(importel) from tmpportesv where codusu = " & vUsu.codigo & " and codtrans = " & DBSet(Text1(3).Text, "N")
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
    
                cadWhere = "albaran.codtrans=" & Val(Text1(3).Text)
                
            Case 1
            
                cadWhere = "albaran.codcomis=" & Val(Text1(3).Text)
                
        End Select
        
        cadWhere = cadWhere & " AND (" & cadAux & ")"
    
        If Not SeleccionaRegistros Then Exit Sub
        
        If Not BloqueaRegistro("albaran", cadWhere) Then
            Select Case Combo1(0).ListIndex
                Case 0
        
                    conn.Execute "update albaran set portespag = 0 where " & cadWhere
                    
                    
                Case 1
                    
                    conn.Execute "update albaran set comisionespag = 0 where " & cadWhere
                    
            End Select
            
                
            CargarAlbaranes vWhere, Combo1(0).ListIndex
            
        End If
    
        Set vFactu = New CFacturaTra
        vFactu.DtoPPago = 0
        vFactu.DtoGnral = 0
        vFactu.EsComisionista = Combo1(0).ListIndex
        If vFactu.CalcularDatosFactura(cadWhere, Text1(3).Text) Then
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
            
            For I = 6 To 26
                FormateaCampo Text1(I)
            Next I
            'Quitar ceros de linea IVA 2
            If Val(Text1(14).Text) = 0 And Val(Text1(11).Text) = 0 Then
                For I = 11 To 20 Step 3
                    Text1(I).Text = QuitarCero(CCur(Text1(I).Text))
                Next I
            End If
            'Quitar ceros de linea IVA 3
            If Val(Text1(15).Text) = 0 And Val(Text1(12).Text) = 0 Then
                For I = 12 To 21 Step 3
                    Text1(I).Text = QuitarCero(CCur(Text1(I).Text))
                Next I
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
        If vFactu.CalcularDatosFacturaSinAlbaran(cadWhere, Text1(3).Text) Then
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
            
            For I = 6 To 26
                FormateaCampo Text1(I)
            Next I
            'Quitar ceros de linea IVA 2
            If Val(Text1(14).Text) = 0 And Val(Text1(11).Text) = 0 Then
                For I = 11 To 20 Step 3
                    Text1(I).Text = QuitarCero(CCur(Text1(I).Text))
                Next I
            End If
            'Quitar ceros de linea IVA 3
            If Val(Text1(15).Text) = 0 And Val(Text1(12).Text) = 0 Then
                For I = 12 To 21 Step 3
                    Text1(I).Text = QuitarCero(CCur(Text1(I).Text))
                Next I
            End If
            
        Else
            MuestraError Err.Number, "Calculando Factura", Err.Description
        End If
        Set vFactu = Nothing
    End If
    
End Sub

Private Function SeleccionaRegistros() As Boolean
'Comprueba que se seleccionan albaranes en la base de datos
'es decir que hay albaranes marcados
'cuando se van marcando albaranes se van añadiendo el la cadena cadWhere
Dim Sql As String

    On Error GoTo ESel
    SeleccionaRegistros = False
    
    If cadWhere = "" Then Exit Function
    
    Sql = "Select count(*) FROM albaran"
    Sql = Sql & " WHERE " & cadWhere
    If RegistrosAListar(Sql) <> 0 Then SeleccionaRegistros = True
    Exit Function
    
ESel:
    SeleccionaRegistros = False
    MuestraError Err.Number, "No hay seleccionados Albaranes", Err.Description
End Function


Private Sub BotonFacturar()
Dim vFactu As CFacturaTra
Dim cad As String

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    
    cad = ""
    If Text1(3).Text = "" Then
        Select Case Combo1(0).ListIndex
            Case 0
                cad = "Falta transportista"
            Case 1
                cad = "Falta comisionista"
            Case Else
                cad = "Faltan datos"
        End Select
    Else
        If Not IsNumeric(Text1(3).Text) Then cad = "Campo transportista/comisionista debe ser numérico"
    End If
    If cad <> "" Then
        MsgBox cad, vbExclamation
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
        
        If cadWhere <> "" Then
            If vFactu.TraspasoAlbaranesAFactura(cadWhere) Then BotonPedirDatos
        Else
            If vFactu.TraspasoPortesVueltaAFactura(cadWhere) Then BotonPedirDatos
        End If
        Set vFactu = Nothing
    
    
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Function ExisteFacturaEnHco() As Boolean
'Comprobamos si la factura ya existe en la tabla de Facturas a Proveedor: scafpc
Dim cad As String

    ExisteFacturaEnHco = False
    'Tiene que tener valor los 3 campos de clave primaria antes de comprobar
    If Not (Text1(0).Text <> "" And Text1(1).Text <> "" And Text1(3).Text <> "") Then Exit Function
    
    ' No debe existir el número de factura para el proveedor en hco
    cad = "SELECT count(*) FROM tcafpc "
    cad = cad & " WHERE codtrans=" & Text1(3).Text & " AND numfactu=" & DBSet(Text1(0).Text, "T") & " AND year(fecfactu)=" & Year(Text1(1).Text)
    If RegistrosAListar(cad) > 0 Then
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

    ModoLineas = 3 'Posem Modo Eliminar Llínia
    
    If Modo = 4 Then 'Modificar Capçalera
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
            Sql = "¿Seguro que desea eliminar el importe?"
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
Dim I As Integer
    
    ModoLineas = 1 'Posem Modo Afegir Llínia
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5
    
    ' *** bloquejar la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    ' **************************************************

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case Index
        Case 0: vtabla = "tmpportesv"
    End Select
    ' ********************************************************
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 0 ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
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
            
            BloquearTxt txtAux(2), False
            BloquearTxt txtAux(3), False
            
            Select Case Index
                ' *** valor per defecte a l'insertar i formateig de tots els camps ***
                Case 0 'cuentas
                    For I = 0 To txtAux.Count - 1
                        txtAux(I).Text = ""
                    Next I
                    txtAux(0).Text = vUsu.codigo
                    txtAux(1).Text = Text1(3).Text 'tranportista
                    txtAux(4).Text = Text1(0).Text 'numfactu
                    txtAux(5).Text = Text1(1).Text 'fecfactu
                    
                    For I = 2 To 3
                        txtAux(I).Text = ""
                    Next I
                    Text2(2).Text = ""
                    PonerFoco txtAux(2)
                    
            End Select
    End Select
End Sub


Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim I As Integer
    Dim j As Integer
    
    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub
    
    ModoLineas = 2 'Modificar llínia
       
    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5
    ' *** bloqueje la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    ' *********************************
  
    Select Case Index
        Case 0 ' *** pose els index de llínies que tenen datagrid (en o sense tab) ***
            If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
                I = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
                DataGridAux(Index).Scroll 0, I
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
            txtAux(0).Text = DataGridAux(Index).Columns(0).Text
            txtAux(1).Text = DataGridAux(Index).Columns(1).Text
            txtAux(4).Text = DataGridAux(Index).Columns(2).Text
            txtAux(5).Text = DataGridAux(Index).Columns(3).Text
            txtAux(2).Text = DataGridAux(Index).Columns(4).Text
            Text2(2).Text = DataGridAux(Index).Columns(5).Text
            txtAux(3).Text = DataGridAux(Index).Columns(6).Text
                       
            BloquearTxt txtAux(2), True
            BloquearTxt txtAux(3), False
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 0 'importes
            PonerFoco txtAux(3)
    End Select
    ' ***************************************************************************************
End Sub

Private Sub CargaGrid(Index As Integer, enlaza As Boolean)
Dim b As Boolean
Dim I As Byte
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
            arregla tots, DataGridAux(Index), Me
        
            DataGridAux(0).Columns(1).Alignment = dbgLeft
            DataGridAux(0).Columns(2).Alignment = dbgRight
        
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            BloquearTxt txtAux(2), Not b
            BloquearTxt txtAux(3), Not b

            If (enlaza = True) And (Not AdoAux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
'                txtAux(2).Text = DataGridAux(Index).Columns(5).Text
'                txtAux(3).Text = DataGridAux(Index).Columns(6).Text
            Else
                txtAux(2).Text = ""
                txtAux(3).Text = ""
            End If
            
    End Select
    DataGridAux(Index).ScrollBars = dbgAutomatic
      
'    ' **** si n'hi han llínies en grids i camps fora d'estos ****
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
'   Basant-se en la informació proporcionada pel vector de camps
'   crea un SQl per a executar una consulta sobre la base de datos que els
'   torne.
' Si ENLAZA -> Enlaça en el data1
'           -> Si no el carreguem sense enllaçar a cap camp
'--------------------------------------------------------------------
Dim Sql As String
Dim Tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
        Case 0 'DESTINOS
            Tabla = "tmpportesv"
            Sql = "SELECT tmpportesv.codusu, tmpportesv.codtrans, tmpportesv.numfactu, tmpportesv.fecfactu, tmpportesv.codmacta, "
            Sql = Sql & " if(" & vParamAplic.NumeroConta & "=0,'',conta" & vParamAplic.NumeroConta & ".cuentas.nommacta), tmpportesv.importel"
            Sql = Sql & " FROM " & Tabla
            If vParamAplic.NumeroConta <> 0 Then
                Sql = Sql & ", conta" & vParamAplic.NumeroConta & ".cuentas "
            End If
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE codusu = -1"
            End If
            
            If vParamAplic.NumeroConta <> 0 Then
                Sql = Sql & " and tmpportesv.codmacta = conta" & vParamAplic.NumeroConta & ".cuentas.codmacta"
            End If
            
            Sql = Sql & " ORDER BY " & Tabla & ".codusu,  " & Tabla & ".codmacta "
            
    End Select
    ' ********************************************************************************
    
    MontaSQLCarga = Sql
End Function


Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    If Text1(0).Text = "" Then Exit Sub
    Select Case Modo
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 1 'afegir llínia
                    InsertarLinea
                    If Not AdoAux(0).Recordset.EOF Then _
                        Me.lblIndicador.Caption = AdoAux(0).Recordset.AbsolutePosition & " de " & AdoAux(0).Recordset.RecordCount
            

                Case 2 'modificar llínies
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
Dim I As Integer
Dim V
    
    Select Case Modo
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 1 'afegir llínia
                    ModoLineas = 0
                    ' *** les llínies que tenen datagrid (en o sense tab) ***
                    If NumTabMto = 0 Then
                        DataGridAux(NumTabMto).AllowAddNew = False
                        ' **** repasar si es diu Data1 l'adodc de la capçalera ***
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

                Case 2 'modificar llínies
                    ModoLineas = 0
                    
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        ' *** l'Index de Fields es el que canvie de la PK de llínies ***
                        V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
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
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " codusu=" & Val(txtAux(0).Text)
    ' *******************************************************
    
    ObtenerWhereCab = vWhere
End Function

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
Dim cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    cad = "(codusu=" & vUsu.codigo & " and codtrans = " & DBSet(Text1(3).Text, "N")
    cad = cad & " and numfactu = " & DBSet(Text1(0).Text, "T") & " and fecfactu = " & DBSet(Text1(1).Text, "F")
    cad = cad & " and codmacta = " & DBSet(AdoAux(0).Recordset!Codmacta, "T") & ")"
    ' ***************************************
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    If SituarDataMULTI(AdoAux(0), cad, Indicador) Then
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
       
    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
    Select Case Index
        Case 0 'destinos
            txtAux(0).visible = False
            For jj = 1 To 3
                txtAux(jj).visible = b
                txtAux(jj).Top = alto
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
'Inserta registre en les taules de Llínies
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
            If cadWhere <> "" Then b = BloqueaRegistro("albaran", cadWhere)
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
'Modifica registre en les taules de Llínies
Dim nomFrame As String
Dim V As Integer
Dim cad As String
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
            If cadWhere <> "" Then BloqueaRegistro "albaran", cadWhere

            ModoLineas = 0

            If NumTabMto <> 3 Then
                V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
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
'    Cad = Cad & ParaGrid(text1(0), 15, "Cód.")
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
'        frmB.vTitulo = "Clientes" ' ***** repasa açò: títol de BuscaGrid *****
'        frmB.vSelElem = 1
'
'        frmB.Show vbModal
'        Set frmB = Nothing
'        'Si ha posat valors i tenim que es formulari de búsqueda llavors
'        'tindrem que tancar el form llançant l'event
'        If HaDevueltoDatos Then
'            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                cmdRegresar_Click
'        Else   'de ha retornat datos, es a decir NO ha retornat datos
'            PonerFoco text1(kCampo)
'        End If
'    End If
End Sub

Private Sub PonerCampos()
Dim I As Integer
Dim codpobla As String, despobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la capçalera
    
    ' *** si n'hi han llínies en datagrids ***
    'For i = 0 To DataGridAux.Count - 1
'        If i <> 3 Then
            CargaGrid I, True
            If Not AdoAux(I).Recordset.EOF Then _
                PonerCamposForma2 Me, AdoAux(I), 2, "FrameAux" & I
'        End If
    ' *******************************************

    ' *** si n'hi han llínies sense datagrid ***
'    CargaFrame 3, True
    ' ***************************************
    
    ' ************* configurar els camps de les descripcions de la capçalera *************
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
        Sql = "select * from tmpportesv where codusu = " & vUsu.codigo & " and codtrans = " & DBSet(Text1(3).Text, "N")
        Sql = Sql & " and numfactu = " & DBSet(Text1(0).Text, "T") & " and fecfactu = " & DBSet(Text1(1).Text, "F")
        Sql = Sql & " and codmacta = " & DBSet(txtAux(2).Text, "T")
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
'Actives unes Opcions de Menú i Toolbar según el modo en que estem
Dim b As Boolean, bAux As Boolean
Dim I As Byte
    
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
    b = (Modo = 3 Or Modo = 4 Or Modo = 2)
    For I = 0 To ToolAux.Count - 1
        ToolAux(I).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.AdoAux(I).Recordset.RecordCount > 0)
        ToolAux(I).Buttons(2).Enabled = bAux
        ToolAux(I).Buttons(3).Enabled = bAux
    Next I
    ' ****************************************
    
'    ' *** si n'hi han tabs que no tenen grids ***
'    i = 3
'    If AdoAux(i).Recordset.EOF Then
'        ToolAux(i).Buttons(1).Enabled = b
'        ToolAux(i).Buttons(2).Enabled = False
'        ToolAux(i).Buttons(3).Enabled = False
'    Else
'        ToolAux(i).Buttons(1).Enabled = False
'        ToolAux(i).Buttons(2).Enabled = b
'    End If
    ' *******************************************
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
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
    If Not PerderFocoGnralLineas(txtAux(Index), ModoLineas) Then Exit Sub
    
    Select Case Index
        Case 2 'Cta Contable
            If txtAux(Index).Text = "" Then Exit Sub
            Text2(2).Text = PonerNombreCuenta(txtAux(Index), Modo)
            If Text2(2).Text = "" Then
                PonerFoco txtAux(Index)
            End If
            
        Case 3 ' Importe
            If txtAux(Index).Text <> "" Then PonerFormatoDecimal txtAux(Index), 1
            
    End Select
End Sub



Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim I As Byte
    
    Combo1(0).Clear
    
    Combo1(0).AddItem "Transportista"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    
    Combo1(0).AddItem "Comisionista"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    
End Sub


Private Sub RecalcularDatosFactura()
Dim I As Integer
Dim Sql As String
Dim cadAux As String
Dim TotalFactura As Currency
    
Dim ImpBImIVA As Currency
Dim ImpIVA As Currency
Dim ImpIVA1 As Currency
Dim ImpIVA2 As Currency
Dim ImpIVA3 As Currency
    
Dim BaseReten As Currency
Dim PorcReten As Currency
Dim ImpReten As Currency

    On Error GoTo eRecalcularDatosFactura
    
    If Not SeleccionaRegistros Then Exit Sub
    
    If Not BloqueaRegistro("albaran", cadWhere) Then
        ListView1.SelectedItem.Checked = False
    End If
    
    TotalFactura = 0
    If Text1(16).Text <> "" Then
        cadAux = Text1(13).Text
        ImpBImIVA = CCur(ImporteSinFormato(Text1(16).Text))
        If cadAux = "" Then cadAux = "0"
        ImpIVA = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
        
        ImpIVA1 = ImpIVA
        
        
        'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
        'los vamos acumulando
        TotalFactura = TotalFactura + ImpBImIVA + ImpIVA
    End If
    
    If Text1(17).Text <> "" Then
        cadAux = Text1(14).Text
        ImpBImIVA = CCur(ImporteSinFormato(Text1(17).Text))
        If cadAux = "" Then cadAux = "0"
        ImpIVA = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
        
        ImpIVA2 = ImpIVA
        
        
        'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
        'los vamos acumulando
        TotalFactura = TotalFactura + ImpBImIVA + ImpIVA
    End If
    
    
    If Text1(18).Text <> "" Then
        cadAux = Text1(15).Text
        ImpBImIVA = CCur(ImporteSinFormato(Text1(18).Text))
        If cadAux = "" Then cadAux = "0"
        ImpIVA = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
        
        ImpIVA3 = ImpIVA
        
        
        'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
        'los vamos acumulando
        TotalFactura = TotalFactura + ImpBImIVA + ImpIVA
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
        
        For I = 19 To 22
            FormateaCampo Text1(I)
        Next I
        'Quitar ceros de linea IVA 2
        If Val(Text1(14).Text) = 0 And Val(Text1(11).Text) = 0 Then
            For I = 11 To 20 Step 3
                Text1(I).Text = ""
            Next I
        End If
        'Quitar ceros de linea IVA 3
        If Val(Text1(15).Text) = 0 And Val(Text1(12).Text) = 0 Then
            For I = 12 To 21 Step 3
                Text1(I).Text = ""
            Next I
        End If
        Exit Sub
        
   
eRecalcularDatosFactura:
    MuestraError Err.Number, "Recalculando Datos de Factura", Err.Description
End Sub



