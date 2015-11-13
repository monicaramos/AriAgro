VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmVtasLinAlbaranes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Variedades de Albaranes"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   10140
   Icon            =   "frmVtasLinAlbaranes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   10140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameAux0 
      Caption         =   "Calibres"
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
      Height          =   2505
      Left            =   135
      TabIndex        =   42
      Top             =   5145
      Width           =   9750
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   9
         Left            =   9540
         MaxLength       =   6
         TabIndex        =   30
         Tag             =   "Pr.Prov|N|S|||albaran_calibre|preciopro|#0.0000||"
         Text            =   "pr.pro"
         Top             =   1800
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox txtAux2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   9270
         TabIndex        =   72
         Text            =   "kilos/caja"
         Top             =   1800
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   8
         Left            =   6615
         MaxLength       =   6
         TabIndex        =   27
         Tag             =   "Unidades|N|S|||albaran_calibre|unidades|##,##0||"
         Text            =   "unida"
         Top             =   1800
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   7
         Left            =   8460
         MaxLength       =   9
         TabIndex        =   29
         Tag             =   "Peso Neto|N|N|||albaran_calibre|pesoneto|###,##0||"
         Text            =   "peso neto"
         Top             =   1800
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   6
         Left            =   7470
         MaxLength       =   9
         TabIndex        =   28
         Tag             =   "Peos Bruto|N|S|||albaran_calibre|pesobrut|###,##0||"
         Text            =   "peso brut"
         Top             =   1800
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   5
         Left            =   1710
         MaxLength       =   9
         TabIndex        =   33
         Tag             =   "Num.Linea 1|N|N|||albaran_calibre|numline1|00|S|"
         Text            =   "Linea"
         Top             =   1800
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   4
         Left            =   2430
         MaxLength       =   6
         TabIndex        =   24
         Tag             =   "Variedad|N|N|||albaran_calibre|codvarie|000000||"
         Text            =   "varied"
         Top             =   1800
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   3
         Left            =   1035
         MaxLength       =   9
         TabIndex        =   32
         Tag             =   "Num.Linea|N|N|||albaran_calibre|numlinea|00|S|"
         Text            =   "Linea"
         Top             =   1800
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.CommandButton btnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   300
         Index           =   0
         Left            =   3690
         MaskColor       =   &H00000000&
         TabIndex        =   34
         ToolTipText     =   "Buscar calibre"
         Top             =   1800
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   3900
         TabIndex        =   51
         Top             =   1800
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   2
         Left            =   5490
         MaxLength       =   9
         TabIndex        =   26
         Tag             =   "Num.Cajas|N|S|||albaran_calibre|numcajas|#,##0||"
         Text            =   "cajas"
         Top             =   1800
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   1
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   25
         Tag             =   "Calibre|N|N|||albaran_calibre|codcalib|00||"
         Text            =   "calibre"
         Top             =   1800
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   0
         Left            =   225
         MaxLength       =   16
         TabIndex        =   31
         Tag             =   "N�mero Albaran|N|N|||albaran_calibre|numalbar|000000|S|"
         Text            =   "numpedid"
         Top             =   1800
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   0
         Left            =   135
         TabIndex        =   43
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
         Left            =   3735
         Top             =   720
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
         Bindings        =   "frmVtasLinAlbaranes.frx":000C
         Height          =   1695
         Index           =   0
         Left            =   135
         TabIndex        =   44
         Top             =   630
         Width           =   9525
         _ExtentX        =   16801
         _ExtentY        =   2990
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
   Begin VB.Frame Frame2 
      Height          =   4590
      Index           =   0
      Left            =   135
      TabIndex        =   37
      Top             =   540
      Width           =   9765
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   20
         Left            =   2385
         TabIndex        =   74
         Text            =   "12345678901234567890"
         Top             =   2820
         Width           =   6405
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   20
         Left            =   1575
         MaxLength       =   3
         TabIndex        =   9
         Tag             =   "Cod.Comsionista|N|S|0|999|albaran_variedad|codcomis|000||"
         Text            =   "123"
         Top             =   2820
         Width           =   720
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Facturar "
         Height          =   195
         Index           =   1
         Left            =   7860
         TabIndex        =   2
         Tag             =   "Facturar|N|N|||albaran_variedad|sefactura|0||"
         Top             =   300
         Width           =   945
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   8250
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Tag             =   "Tipo Varieadad|N|N|||albaran_variedad|codtipo||N|"
         Top             =   3540
         Width           =   1440
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   19
         Left            =   6750
         MaxLength       =   15
         TabIndex        =   16
         Tag             =   "N�mero Traza|T|S|||albaran_variedad|nrotraza|||"
         Text            =   "123456789012345"
         Top             =   3540
         Width           =   1440
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   1575
         MaxLength       =   3
         TabIndex        =   8
         Tag             =   "Cod.Palet|N|S|0|999|albaran_variedad|codpalet|000||"
         Text            =   "123"
         Top             =   2460
         Width           =   720
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   18
         Left            =   2385
         TabIndex        =   69
         Text            =   "12345678901234567890"
         Top             =   2460
         Width           =   6405
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   5520
         MaxLength       =   12
         TabIndex        =   15
         Tag             =   "Referencia|T|S|||albaran_variedad|referencia|||"
         Top             =   3540
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   6870
         MaxLength       =   6
         TabIndex        =   66
         Tag             =   "Unidades|N|S|0|99999|albaran_variedad|unidades|##,##0||"
         Top             =   4080
         Width           =   945
      End
      Begin VB.TextBox Text1 
         Height          =   555
         Index           =   15
         Left            =   1590
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Tag             =   "Observaciones|T|S|||albaran_variedad|observac|||"
         Top             =   3900
         Width           =   4035
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   14
         Left            =   4410
         MaxLength       =   8
         TabIndex        =   14
         Tag             =   "Imp.Comisi�n|N|S|||albaran_variedad|impcomis|#,##0.00||"
         Top             =   4140
         Width           =   1035
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   2385
         TabIndex        =   62
         Text            =   "12345678901234567890"
         Top             =   2070
         Width           =   6405
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   1575
         MaxLength       =   3
         TabIndex        =   7
         Tag             =   "Incidencia|N|N|0|999|albaran_variedad|codincid|000||"
         Text            =   "123"
         Top             =   2070
         Width           =   720
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   2790
         MaxLength       =   16
         TabIndex        =   12
         Tag             =   "Precio Definitivo|N|S|||albaran_variedad|preciodef|#0.0000||"
         Top             =   3540
         Width           =   1035
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   4035
         MaxLength       =   3
         TabIndex        =   13
         Tag             =   "Total Palets|T|S|||albaran_variedad|totpalet|||"
         Top             =   3540
         Width           =   1035
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   5880
         MaxLength       =   16
         TabIndex        =   19
         Tag             =   "Numero Cajas|N|S|0|999999|albaran_variedad|numcajas|###,##0||"
         Top             =   4080
         Width           =   945
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   1590
         MaxLength       =   16
         TabIndex        =   11
         Tag             =   "Precio Provisional|N|S|||albaran_variedad|preciopro|#0.0000||"
         Top             =   3540
         Width           =   1035
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1575
         MaxLength       =   16
         TabIndex        =   6
         Tag             =   "Forfait|T|N|||albaran_variedad|codforfait|||"
         Text            =   "1234567890123456"
         Top             =   1710
         Width           =   1530
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   3150
         TabIndex        =   56
         Text            =   "12345678901234567890"
         Top             =   1710
         Width           =   5640
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1575
         MaxLength       =   3
         TabIndex        =   5
         Tag             =   "Marca|N|N|0|999|albaran_variedad|codmarca|000||"
         Text            =   "123"
         Top             =   1350
         Width           =   720
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2385
         TabIndex        =   54
         Text            =   "12345678901234567890"
         Top             =   1350
         Width           =   6405
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1575
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "Variedad Comercial|N|N|0|999999|albaran_variedad|codvarco|000000||"
         Text            =   "123456"
         Top             =   990
         Width           =   720
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2385
         TabIndex        =   52
         Text            =   "12345678901234567890"
         Top             =   990
         Width           =   6405
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2385
         TabIndex        =   49
         Text            =   "12345678901234567890"
         Top             =   630
         Width           =   6405
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1575
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "Variedad|N|N|0|999999|albaran_variedad|codvarie|000000||"
         Text            =   "123456"
         Top             =   630
         Width           =   720
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   8805
         MaxLength       =   16
         TabIndex        =   21
         Tag             =   "Peso Neto|N|S|0|999999|albaran_variedad|pesoneto|###,##0||"
         Top             =   4080
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   7860
         MaxLength       =   16
         TabIndex        =   20
         Tag             =   "Peso Bruto|N|S|0|999999|albaran_variedad|pesobrut|###,##0||"
         Top             =   4080
         Width           =   900
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   210
         MaxLength       =   3
         TabIndex        =   10
         Tag             =   "Categoria|T|S|||albaran_variedad|categori|||"
         Top             =   3540
         Width           =   1035
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000013&
         Height          =   285
         Index           =   1
         Left            =   3645
         MaxLength       =   40
         TabIndex        =   1
         Tag             =   "Linea Albaran|N|N|||albaran_variedad|numlinea|00|S|"
         Text            =   "1234657890123456798012345678901234567890"
         Top             =   225
         Width           =   600
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000013&
         Height          =   285
         Index           =   0
         Left            =   1575
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "N�mero albaran|N|N|||albaran_variedad|numalbar|000000|S|"
         Text            =   "1234567"
         Top             =   225
         Width           =   765
      End
      Begin VB.Label Label5 
         Caption         =   "Comisionista"
         Height          =   255
         Left            =   180
         TabIndex        =   75
         Top             =   2820
         Width           =   960
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1305
         ToolTipText     =   "Buscar Comisionista"
         Top             =   2820
         Width           =   240
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   0
         Left            =   8880
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Variedad"
         Height          =   255
         Index           =   11
         Left            =   8250
         TabIndex        =   73
         Top             =   3330
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "N�mero Traza"
         Height          =   255
         Index           =   49
         Left            =   6750
         TabIndex        =   71
         Top             =   3315
         Width           =   1320
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1305
         ToolTipText     =   "Buscar tipo palet"
         Top             =   2460
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo Palet"
         Height          =   255
         Left            =   180
         TabIndex        =   70
         Top             =   2460
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "Referencia"
         Height          =   255
         Index           =   10
         Left            =   5520
         TabIndex        =   68
         Top             =   3315
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Unidades"
         Height          =   255
         Index           =   9
         Left            =   6870
         TabIndex        =   67
         Top             =   3855
         Width           =   960
      End
      Begin VB.Label Label29 
         Caption         =   "Observaciones"
         Height          =   255
         Left            =   210
         TabIndex        =   65
         Top             =   3945
         Width           =   1095
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   1320
         ToolTipText     =   "Zoom descripci�n"
         Top             =   3945
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Imp.Comisi�n"
         Height          =   255
         Index           =   8
         Left            =   4440
         TabIndex        =   64
         Top             =   3900
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Incidencia"
         Height          =   255
         Left            =   180
         TabIndex        =   63
         Top             =   2070
         Width           =   960
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1305
         ToolTipText     =   "Buscar Marca"
         Top             =   2070
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Pr.Definitivo"
         Height          =   255
         Index           =   7
         Left            =   2790
         TabIndex        =   61
         Top             =   3315
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Total Palets"
         Height          =   255
         Index           =   6
         Left            =   4035
         TabIndex        =   60
         Top             =   3315
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Nro.Cajas"
         Height          =   255
         Index           =   5
         Left            =   5880
         TabIndex        =   59
         Top             =   3855
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "Pr.Provisional"
         Height          =   255
         Index           =   4
         Left            =   1605
         TabIndex        =   58
         Top             =   3315
         Width           =   1275
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1305
         ToolTipText     =   "Buscar Forfait"
         Top             =   1710
         Width           =   240
      End
      Begin VB.Label Label13 
         Caption         =   "Forfait"
         Height          =   255
         Left            =   180
         TabIndex        =   57
         Top             =   1710
         Width           =   690
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1305
         ToolTipText     =   "Buscar Marca"
         Top             =   1350
         Width           =   240
      End
      Begin VB.Label Label12 
         Caption         =   "Marca"
         Height          =   255
         Left            =   180
         TabIndex        =   55
         Top             =   1350
         Width           =   690
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1305
         ToolTipText     =   "Buscar Variedad Comercial"
         Top             =   990
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Var.Comercial"
         Height          =   255
         Left            =   180
         TabIndex        =   53
         Top             =   990
         Width           =   1050
      End
      Begin VB.Label Label10 
         Caption         =   "Variedad"
         Height          =   255
         Left            =   180
         TabIndex        =   50
         Top             =   630
         Width           =   690
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1305
         ToolTipText     =   "Buscar Variedad"
         Top             =   630
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Peso Neto"
         Height          =   255
         Index           =   3
         Left            =   8805
         TabIndex        =   48
         Top             =   3855
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "Peso Bruto"
         Height          =   255
         Index           =   2
         Left            =   7860
         TabIndex        =   47
         Top             =   3855
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "Categoria"
         Height          =   255
         Index           =   1
         Left            =   210
         TabIndex        =   46
         Top             =   3315
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Linea"
         Height          =   255
         Index           =   1
         Left            =   3105
         TabIndex        =   45
         Top             =   255
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "N�mero Albar�n"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   38
         Top             =   255
         Width           =   1230
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   135
      TabIndex        =   35
      Top             =   7710
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
         TabIndex        =   36
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8835
      TabIndex        =   23
      Top             =   7845
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7740
      TabIndex        =   22
      Top             =   7845
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   1980
      Top             =   6120
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
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   10140
      _ExtentX        =   17886
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
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
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
      Enabled         =   0   'False
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Index           =   0
         Left            =   6525
         TabIndex        =   41
         Top             =   90
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   8820
      TabIndex        =   39
      Top             =   7845
      Visible         =   0   'False
      Width           =   1035
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
      Begin VB.Menu mnExpandirOperaciones 
         Caption         =   "Expandir &Operaciones"
         Shortcut        =   ^O
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
Attribute VB_Name = "frmVtasLinAlbaranes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MONICA                   -+-+
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
Public Albaran As Long
Public Linea As Integer

Public ModoExt As Byte

' *** declarar els formularis als que vaig a cridar ***
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1

Private WithEvents frmMar As frmManMarcas 'marcas
Attribute frmMar.VB_VarHelpID = -1
Private WithEvents frmVar As frmManVariedad 'variedades
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmFor As frmManForfaits 'forfaits
Attribute frmFor.VB_VarHelpID = -1
Private WithEvents frmCali As frmManCalibres 'calibres
Attribute frmCali.VB_VarHelpID = -1
Private WithEvents frmIncid As frmManInciden 'incidencias
Attribute frmIncid.VB_VarHelpID = -1
Private WithEvents frmPal As frmManPaleConf 'Palets de confreccion
Attribute frmPal.VB_VarHelpID = -1
Private WithEvents frmTra1 As frmManAgencias 'Form Mto de Comisionistas
Attribute frmTra1.VB_VarHelpID = -1


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

Dim KilosAnt As Currency
Dim CajasAnt As Currency
Dim ForfaitAnt As String
Dim CodPaletAnt As String
Dim TotPaletAnt As String

Dim VarieAnt As String


Private BuscaChekc As String


Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    '++monica
'    BloqueaRegistro "palets", "numpalet = " & Text1(0).Text
    
    Select Case Index
        Case 0 'calibres
            Set frmCali = New frmManCalibres
            frmCali.DatosADevolverBusqueda = "0|2|3|"
            frmCali.CodigoActual = txtAux(1).Text
            frmCali.ParamVariedad = txtAux(4).Text
            frmCali.Show vbModal
            Set frmCali = Nothing
            PonerFoco txtAux(1)
    End Select
    If Modo = 4 Then BloqueaRegistro "albaran", "numalbar = " & Text1(0).Text
    'BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub


Private Sub cmdAceptar_Click()
Dim b As Boolean
Dim V As Integer
Dim Forfait As String

    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'B�SQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm2(Me, 1) Then
'                    text2(9).Text = PonerNombreCuenta(text1(9), Modo, text1(0).Text)
        
'                    Data1.RecordSource = "Select * from " & NombreTabla & _
'                                        " where numpalet = " & DBSet(text1(0).Text, "N") & _
'                                        " and numlinea = " & DBSet(text1(1).Text, "N") & " " & Ordenacion
'                    PosicionarData

                    TerminaBloquear
                    BloqueaRegistro "albaran", "numalbar = " & Text1(0).Text
                    
                    CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                    PonerCadenaBusqueda
                    'Ponerse en Modo Insertar Lineas
                    BotonAnyadirLinea 0

                End If
            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                Modificar
                TerminaBloquear
                '++monica
                BloqueaRegistro "albaran", "numalbar = " & Text1(0).Text
                
                PosicionarData
            Else
                ModoLineas = 0
            End If
        ' *** si n'hi han ll�nies ***
        Case 5 'LL�NIES
            Select Case ModoLineas
                Case 1 'afegir ll�nia
                    If InsertarLinea Then
                        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                        PonerCadenaBusqueda
                        b = BLOQUEADesdeFormulario2(Me, Data1, 1)
                        CargaGrid 0, True
                        If b Then BotonAnyadirLinea NumTabMto
            
                        
                    End If
                Case 2 'modificar ll�nies
                    If ModificarLinea Then
                        ModoLineas = 0
                        
                        V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el n� de llinia
                        
                        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                        PonerCadenaBusqueda
                        b = BLOQUEADesdeFormulario2(Me, Data1, 1)
                        
                        CargaGrid NumTabMto, True
                        
                        PonerFocoGrid Me.DataGridAux(NumTabMto)
                        AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
                        
                        LLamaLineas NumTabMto, 0
                        
                        TerminaBloquear
                        '++monica
                        BloqueaRegistro "albaran", "numalbar = " & Text1(0).Text
                        PosicionarData
                    Else
                        PonerFoco txtAux(1)
                    End If
            End Select
'--monica: la actualizacion de costes se hace en insertarlinea y modificarlinea
'            ActualizarCostes Data1.Recordset.Fields(0), Data1.Recordset.Fields(1), True

            'nuevo calculamos los totales de lineas
            CalcularTotales
        ' **************************
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PrimeraVez Then
        PrimeraVez = False
    
        PonerCampos
        ModoLineas = 0
           
        CalcularTotales
        
        Modo = ModoExt
        Select Case Modo
            Case 0
                DatosADevolverBusqueda = "ZZ"
                PonerModo Modo
                CargaGrid 0, True
            Case 3
                mnNuevo_Click
            Case 4
                mnModificar_Click
        End Select
        
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim cad As String

    cad = ""
    If Text1(0).Text <> "" And Text1(1).Text <> "" Then
        cad = Text1(0).Text & "|" & Text1(1).Text & "|"
    End If
    RaiseEvent DatoSeleccionado(cad)

    CheckValueGuardar Me.Name, Me.chkVistaPrevia(0).Value
    Screen.MousePointer = vbDefault
    
    TerminaBloquear

End Sub

Private Sub Form_Load()
Dim I As Integer

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
        .Buttons(12).Image = 10  'Imprimir
        .Buttons(13).Image = 11  'Eixir
        'el 13 i el 14 son separadors
        .Buttons(btnPrimero).Image = 6  'Primer
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Seg�ent
        .Buttons(btnPrimero + 3).Image = 9 '�ltim
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
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
    
    For I = 0 To imgAyuda.Count - 1
        imgAyuda(I).Picture = frmPpal.ImageListB.ListImages(10).Picture
    Next I
    
    'carga IMAGES de mail
'    For i = 0 To Me.imgMail.Count - 1
'        Me.imgMail(i).Picture = frmPpal.imgListImages16.ListImages(2).Picture
'    Next i
    
    CargaCombo
    LimpiarCampos   'Neteja els camps TextBox
    ' ******* si n'hi han ll�nies *******
    DataGridAux(0).ClearFields
    
    '*** canviar el nom de la taula i l'ordenaci� de la cap�alera ***
    NombreTabla = "albaran_variedad"
    Ordenacion = " ORDER BY numalbar"
    
    'Mirem com est� guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    Data1.ConnectionString = Conn
    '***** cambiar el nombre de la PK de la cabecera *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where numalbar=" & Albaran & " and numlinea = " & Linea
    Data1.Refresh
    
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
    Combo1.ListIndex = -1

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub LimpiarCamposLin(FrameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, FrameAux  'M�tode general: Neteja els controls TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO s'habiliten, o no, els diversos camps del
'   formulari en funci� del modo en que anem a treballar
Private Sub PonerModo(Kmodo As Byte, Optional indFrame As Integer)
Dim I As Integer, NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo
 
    'Actualisa Iconos Insertar,Modificar,Eliminar
    'ActualizarToolbar Modo, Kmodo
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo, ModoLineas
       
    BuscaChekc = ""
       
    'Modo 2. N'hi han datos i estam visualisant-los
    '=========================================
    'Posem visible, si es formulari de b�squeda, el bot� "Regresar" quan n'hi han datos
'    If DatosADevolverBusqueda <> "" Then
'        cmdRegresar.visible = (Modo = 2)
'    Else
'        cmdRegresar.visible = False
'    End If
    
    Text1(5).Enabled = True
    
    
    '=======================================
    b = (Modo = 2)
    'Posar Fleches de desplasament visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Nom�s es per a saber que n'hi ha + d'1 registre
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    '---------------------------------------------
    
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    cmdRegresar.visible = Not b

    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
    'Si estem en Insertar a m�s neteja els camps Text1
    BloquearText1 Me, Modo
    BloquearCmb Combo1, Not b
    BloquearChk Check1(1), Not b
    '*** si n'hi han combos a la cap�alera ***
    '**************************
    
    ' ***** bloquejar tots els controls visibles de la clau primaria de la cap�alera ***
    If Modo = 4 Or Modo = 3 Then
        BloquearTxt Text1(0), True, True 'si estic en  modificar, bloqueja la clau primaria
        BloquearTxt Text1(1), True, True  'si estic en  modificar, bloqueja la clau primaria
    End If
    ' **********************************************************************************
    
    ' numero de cajas, peso bruto y peso neto siempre bloqueados
    BloquearTxt Text1(7), True
    BloquearTxt Text1(8), True
    BloquearTxt Text1(10), True
    BloquearTxt Text1(16), True
    
    
    ' **** si n'hi han imagens de buscar en la cap�alera *****
    BloquearImgBuscar Me, Modo, ModoLineas
    BloquearImgZoom Me, Modo, ModoLineas
    ' ********************************************************
    
'[Monica]01/10/2012: dejamos que modifiquen la variedad real
'    imgBuscar(0).visible = (Modo = 3)
'    imgBuscar(0).Enabled = (Modo = 3)
        
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    PonerLongCampos

    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 0, False
    End If
    
    b = (Modo = 4) Or (Modo = 2)
    DataGridAux(0).Enabled = b
      
    ' ****** si n'hi han combos a la cap�alera ***********************
    ' ****************************************************************
    
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
Dim I As Byte
    
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
    Toolbar1.Buttons(12).Enabled = True And Not DeConsulta
       
    ' *** si n'hi han ll�nies que tenen grids (en o sense tab) ***
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
Dim sql As String
Dim tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
               
        Case 0 'CALIBRES
            sql = "SELECT albaran_calibre.numalbar, albaran_calibre.numlinea, albaran_calibre.numline1, "
            sql = sql & "albaran_calibre.codvarie, albaran_calibre.codcalib, calibres.nomcalib, albaran_calibre.numcajas, albaran_calibre.unidades,  "
            sql = sql & "albaran_calibre.pesobrut, albaran_calibre.pesoneto, round(albaran_calibre.pesoneto / albaran_calibre.numcajas,2), albaran_calibre.preciopro "
            sql = sql & " FROM albaran_calibre, calibres "
            If enlaza Then
                sql = sql & ObtenerWhereCab(True)
            Else
                sql = sql & " WHERE albaran_calibre.numalbar = '-1'"
            End If
            sql = sql & " and albaran_calibre.codcalib = calibres.codcalib"
            sql = sql & " and albaran_calibre.codvarie = calibres.codvarie"
            sql = sql & " ORDER BY albaran_calibre.codcalib"
               
    End Select
    
    MontaSQLCarga = sql
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
        '   Com la clau principal es �nica, en posar el sql apuntant
        '   al valor retornat sobre la clau ppal es suficient
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        ' **********************************
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmCali_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 2) 'codcalib
    txtAux2(2).Text = RecuperaValor(CadenaSeleccion, 3) 'descripcion
End Sub

Private Sub frmFor_DatoSeleccionado(CadenaSeleccion As String)
'Forfaits
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codforfait
    text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmIncid_DatoSeleccionado(CadenaSeleccion As String)
'Incidencias
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codincid
    text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmMar_DatoSeleccionado(CadenaSeleccion As String)
'Marcas
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codmarca
    text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmPal_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de paelets de confeccion
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod palet
    text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Descripcion
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Variedades
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codvariedad
    text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(indice).Text = vCampo
End Sub

Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        indice = 15
        frmZ.pTitulo = "Observaciones de la Linea de Albar�n"
        frmZ.pValor = Text1(indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco Text1(indice)
    End If

End Sub

Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
            vCadena = "Si est� marcado y no es una l�nea facturada aparecer� en el listado de" & vbCrLf & _
                      "Albaranes Pdtes de Facturar. " & vbCrLf & vbCrLf & _
                      "En caso contrario no aparecer� en dicho listado " & vbCrLf & vbCrLf
                      
    End Select
    MsgBox vCadena, vbInformation, "Descripci�n de Ayuda"
    
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    Screen.MousePointer = vbHourglass
    frmListConfeccion.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnModificar_Click()
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************
'--monica
'    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
            BotonModificar
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
        Case 12 'Imprimir
            mnImprimir_Click
        Case 13    'Eixir
            mnSalir_Click
            
        Case btnPrimero To btnPrimero + 3 'Fleches Despla�ament
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub BotonBuscar()
Dim I As Integer
' ***** Si la clau primaria de la cap�alera no es Text1(0), canviar-ho en <=== *****
'    If Modo <> 1 Then
'        LimpiarCampos
'        PonerModo 1
'        PonerFoco Text1(0) ' <===
'        Text1(0).BackColor = vbYellow ' <===
'        ' *** si n'hi han combos a la cap�alera ***
'    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
'    End If
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
    cad = ""
    cad = cad & ParaGrid(Text1(0), 20, "C�digo")
    cad = cad & ParaGrid(Text1(1), 20, "Confecci�n")
    cad = cad & ParaGrid(Text1(2), 60, "Descripci�n")
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vtabla = NombreTabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        frmB.vDevuelve = "0|1|2|" '*** els camps que volen que torne ***
        frmB.vTitulo = "Forfaits" ' ***** repasa a��: t�tol de BuscaGrid *****
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
Dim I As Integer
Dim J As Integer

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
    
    If chkVistaPrevia(0).Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

Private Sub BotonAnyadir()

    LimpiarCampos 'Huida els TextBox
    
    PonerModo 3
    
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la cap�alera *******
'    text1(0).Text = SugerirCodigoSiguienteStr("forfaits", "codforfait")
'    FormateaCampo text1(0)
    
    Text1(0).Text = Albaran
    Text1(1).Text = SugerirCodigoSiguienteStr("albaran_variedad", "numlinea", "numalbar = " & Text1(0).Text)
    Text1(0).Locked = True
    Text1(1).Locked = True
    
    Combo1.ListIndex = 0
    
    Check1(1).Value = 1
    
    PonerFoco Text1(2) '*** 1r camp visible que siga PK ***
    
End Sub

Private Sub BotonModificar()

    PonerModo 4
    
    Text1(0).Text = Albaran
    Text1(1).Text = Linea
    
    Text1(0).BackColor = &H80000013
    Text1(1).BackColor = &H80000013

    ' *** bloquejar els camps visibles de la clau primaria de la cap�alera ***
    BloquearTxt Text1(0), True
    BloquearTxt Text1(1), True
    
    '[Monica]01/10/2012: dejo modificar la variedad
'    BloquearTxt Text1(2), True
    
    'guardamos los kilos, cajas y forfaits
    KilosAnt = DBLet(Data1.Recordset!PesoNeto, "N")
    CajasAnt = DBLet(Data1.Recordset!NumCajas, "N")
    ForfaitAnt = DBLet(Data1.Recordset!codforfait, "T")
    CodPaletAnt = DBLet(Data1.Recordset!CodPalet, "N")
    TotPaletAnt = DBLet(Data1.Recordset!TotPalet, "N")
    VarieAnt = DBLet(Data1.Recordset!codvarie, "N")
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    '[Monica]01/10/2012: dejamos modificar la variedad real
    PonerFoco Text1(2)
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
    cad = "�Seguro que desea eliminar el Forfait?"
    cad = cad & vbCrLf & "C�digo: " & Data1.Recordset.Fields(0)
    cad = cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(1)
    
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not eliminar Then
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
Dim codPobla As String, desPobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la cap�alera
    
    ' *** si n'hi han ll�nies en datagrids ***
    'For i = 0 To DataGridAux.Count - 1
    For I = 0 To 0
            CargaGrid I, True
            If Not AdoAux(I).Recordset.EOF Then _
                PonerCamposForma2 Me, AdoAux(I), 2, "FrameAux" & I
    Next I

    
    ' ************* configurar els camps de les descripcions de la cap�alera *************
    text2(2).Text = PonerNombreDeCod(Text1(2), "variedades", "nomvarie")
    text2(3).Text = DevuelveDesdeBDNew(cAgro, "variedades", "nomvarie", "codvarie", Text1(3).Text, "N")
    text2(4).Text = PonerNombreDeCod(Text1(4), "marcas", "nommarca")
    text2(5).Text = PonerNombreDeCod(Text1(5), "forfaits", "nomconfe")
    text2(13).Text = PonerNombreDeCod(Text1(13), "inciden", "nomincid")
    text2(18).Text = PonerNombreDeCod(Text1(18), "confpale", "nompalet")
    text2(20).Text = DevuelveDesdeBDNew(cAgro, "agencias", "nomtrans", "codtrans", Text1(20).Text, "N")
    
    ' ********************************************************************************
    
    CalcularTotales
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu
End Sub

Private Sub Check1_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "check1(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "check1(" & Index & ")|"
    End If
End Sub

Private Sub Check1_GotFocus(Index As Integer)
    PonerFocoChk Me.Check1(Index)
End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub cmdCancelar_Click()
Dim I As Integer
Dim V


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
                '++monica
                BloqueaRegistro "albaran", "numalbar= " & Text1(0).Text
                
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
                        V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el n� de llinia
                        AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
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
Dim sql As String
'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm2(Me, 1)
    If Not b Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
    If (Modo = 3) Then 'insertar
        'comprobar si existe ya el cod. del campo clave primaria
        sql = ""
        sql = DevuelveDesdeBDNew(cAgro, "albaran_calibre", "numalbar", "numalbar", Text1(0).Text, "N", , "numlinea", Text1(1).Text, "N")
        If sql <> "" Then
            MsgBox "Ya existe el numero de linea para este albar�n", vbExclamation
            b = False
        End If
    End If
    
    '[Monica]01/10/2012: Modifican la variedad real
    If Modo = 4 Then
        If CLng(VarieAnt) <> CLng(Text1(2).Text) Then
            'comprobamos que no me vaya a fallar la referencial a calibres
            If Not ExistenMismosCalibres Then
                MsgBox "La variedad no tiene los mismos calibres que el albaran. Revise.", vbExclamation
                b = False
            End If
        End If
    End If
    
    ' ************************************************************************************
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function ExistenMismosCalibres() As Boolean
Dim Rs As ADODB.Recordset
Dim sql As String
Dim Sql2 As String
Dim b As Boolean

    On Error GoTo eExistenMismosCalibres


    sql = "select codcalib from albaran_calibre where numalbar = " & DBSet(Albaran, "N") & " and numlinea = " & DBSet(Linea, "N")

    b = True
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF And b
        Sql2 = "select count(*) from calibres where codvarie = " & DBSet(Text1(2).Text, "N")
        Sql2 = Sql2 & " and codcalib = " & DBSet(Rs!codcalib, "N")
        
        If TotalRegistros(Sql2) = 0 Then b = False
    
        Rs.MoveNext
    Wend
    Set Rs = Nothing

    ExistenMismosCalibres = b
    Exit Function
    
eExistenMismosCalibres:
    MuestraError Err.Number, "Existen mismos calibres", Err.Description
End Function



Private Sub PosicionarData()
Dim cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la cap�alera, no llevar els () ***
    cad = "(numalbar=" & DBSet(Text1(0).Text, "N") & ")"
    cad = cad & " and (numlinea = " & DBSet(Text1(1).Text, "N") & ")"
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    If SituarDataMULTI(Data1, cad, Indicador) Then
    'If SituarData(Data1, cad, Indicador) Then
        If ModoLineas <> 1 Then PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
End Sub

Private Function eliminar() As Boolean
Dim vWhere As String

    On Error GoTo FinEliminar

    Conn.BeginTrans
    ' ***** canviar el nom de la PK de la cap�alera, repasar codEmpre *******
    vWhere = " WHERE codforfait=" & DBSet(Data1.Recordset!codforfait, "T")
        
    ' ***** elimina les ll�nies ****
    Conn.Execute "DELETE FROM forfaits_envases " & vWhere
        
    Conn.Execute "DELETE FROM forfaits_costes " & vWhere
        
    'Eliminar la CAP�ALERA
    Conn.Execute "Delete from " & NombreTabla & vWhere
       
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        Conn.RollbackTrans
        eliminar = False
    Else
        Conn.CommitTrans
        eliminar = True
    End If
End Function

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
Dim Variedad As String

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    
    ' ***************** configurar els LostFocus dels camps de la cap�alera *****************
    Select Case Index
        Case 0 'codigo de forfait
            Text1(Index).Text = UCase(Text1(Index).Text)
        
        Case 2, 3 'Variedad
            If PonerFormatoEntero(Text1(Index)) Then
                text2(Index).Text = DevuelveDesdeBDNew(cAgro, "variedades", "nomvarie", "codvarie", Text1(Index).Text, "N")
                If text2(Index).Text = "" Then
                    cadMen = "No existe la Variedad: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "�Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        indice = Index + 2
                        Set frmVar = New frmManVariedad
                        frmVar.DatosADevolverBusqueda = "0|1|"
                        frmVar.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        
                        frmVar.Show vbModal
                        Set frmVar = Nothing
                        '++monica
                        BloqueaRegistro "albaran", "numalbar = " & Text1(0).Text
'                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                text2(Index).Text = ""
            End If
            
        Case 4 'Marca
            If PonerFormatoEntero(Text1(Index)) Then
                text2(Index) = PonerNombreDeCod(Text1(Index), "marcas", "nommarca")
                If text2(Index).Text = "" Then
                    cadMen = "No existe la Marca: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "�Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        indice = Index + 2
                        Set frmMar = New frmManMarcas
                        frmMar.DatosADevolverBusqueda = "0|1|"
                        frmMar.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        
                        frmMar.Show vbModal
                        Set frmMar = Nothing
                        '++monica
                        BloqueaRegistro "albaran", "numalbar = " & Text1(0).Text
'                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                text2(Index).Text = ""
            End If
                
        Case 5 'Forfait
            If Text1(Index).Text <> "" Then
                text2(Index) = PonerNombreDeCod(Text1(Index), "forfaits", "nomconfe")
                If text2(Index).Text = "" Then
                    cadMen = "No existe el Forfait: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "�Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        indice = Index + 2
                        Set frmFor = New frmManForfaits
                        frmFor.DatosADevolverBusqueda = "0|1|"
                        frmFor.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        
                        frmFor.Show vbModal
                        Set frmFor = Nothing
                        '++monica
                        BloqueaRegistro "albaran", "numalbar = " & Text1(0).Text
'                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                '++monica:02/12/2008 control d que el forfait sea de la variedad introducida
                Else
                    Variedad = ""
                    Variedad = DevuelveDesdeBDNew(cAgro, "forfaits", "codvarie", "codforfait", Text1(Index).Text, "T")
                    If Variedad <> "" Then
                        If CInt(Variedad) <> CInt(Text1(2).Text) Then
                            MsgBox "El Forfait no es de la Variedad introducida.", vbExclamation
                        End If
                    End If
                '++
                End If
            Else
                text2(Index).Text = ""
            End If
        
        Case 13 'Incidencias
            If Text1(Index).Text <> "" Then text2(Index) = PonerNombreDeCod(Text1(Index), "inciden", "nomincid", "codincid", "N")
        
        Case 6 ' categoria
            Text1(Index).Text = UCase(Text1(Index).Text)
            
        Case 7, 8 'peso bruto y peso neto
            PonerFormatoEntero Text1(Index)
            
        Case 9, 12
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 7

        Case 14 ' importe de comision
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 3

        Case 18 ' codigo de paletizacion
            If PonerFormatoEntero(Text1(Index)) Then
                text2(Index).Text = PonerNombreDeCod(Text1(Index), "confpale", "nompalet")
                If text2(Index).Text = "" Then
                    cadMen = "No existe el Tipo de Palet: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "�Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmPal = New frmManPaleConf
                        frmPal.DatosADevolverBusqueda = "0|1|"
                        frmPal.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmPal.Show vbModal
                        Set frmPal = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            End If

        Case 15
            cmdAceptar.SetFocus
            
        Case 20 'Comisionistas
            If PonerFormatoEntero(Text1(Index)) Then
                text2(20).Text = DevuelveDesdeBDNew(cAgro, "agencias", "nomtrans", "codtrans", Text1(Index).Text, "N")
                If text2(20).Text = "" Then
                    MsgBox "No existe el comisionista. Revise.", vbExclamation
                    PonerFoco Text1(Index)
                Else
                    ' comprobamos que se trata de un comisionista
                    If EsTransportista(Text1(Index)) Then
                        MsgBox "Este c�digo corresponde a una Agencia de Transporte. " & vbCrLf & "No a un Comisionista. Revise.", vbExclamation
                        PonerFoco Text1(Index)
                    End If
                End If
            End If
            
    End Select
        ' ***************************************************************************
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 2: KEYBusqueda KeyAscii, 0 'VARIEDAD
                Case 3: KEYBusqueda KeyAscii, 1 'VARIEDAD COMERCIAL
                Case 4: KEYBusqueda KeyAscii, 2 'MARCA
                Case 5: KEYBusqueda KeyAscii, 3 'FORFAIT
                Case 13: KEYBusqueda KeyAscii, 4 'INCIDENCIA
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvan�ar/Retrocedir els camps en les fleches de despla�ament del teclat.
    KEYdown KeyCode
End Sub

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then 'ESC
        If (Modo = 0 Or Modo = 2) Then Unload Me
    End If
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub



'************* LLINIES: ****************************
Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
'-- pon el bloqueo aqui
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
    
    'guardamos los kilos, cajas y forfaits
    KilosAnt = DBLet(Data1.Recordset!PesoNeto, "N")
    CajasAnt = DBLet(Data1.Recordset!NumCajas, "N")
    ForfaitAnt = DBLet(Data1.Recordset!codforfait, "T")
    CodPaletAnt = DBLet(Data1.Recordset!CodPalet, "N")
    TotPaletAnt = DBLet(Data1.Recordset!TotPalet, "N")
    
    Select Case Button.Index
        Case 1
            BotonAnyadirLinea Index
        Case 2
            BotonModificarLinea Index
        Case 3
            BotonEliminarLinea Index
        Case Else
    End Select
    End If
End Sub

Private Sub BotonEliminarLinea(Index As Integer)
Dim sql As String
Dim vWhere As String
Dim eliminar As Boolean
Dim bol As Boolean
Dim MenError As String

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
    eliminar = False
   
    vWhere = ObtenerWhereCab(True)
    
    ' ***** independentment de si tenen datagrid o no,
    ' canviar els noms, els formats i el DELETE *****
    Select Case Index
        Case 0 'calibres
            sql = "�Seguro que desea eliminar el Calibre?"
            sql = sql & vbCrLf & "Calibre: " & AdoAux(Index).Recordset!codcalib
            If MsgBox(sql, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
                sql = "DELETE FROM albaran_calibre "
                sql = sql & vWhere & " AND numline1= " & AdoAux(Index).Recordset!numline1
            End If
            
    End Select

    If eliminar Then
        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        BloqueaRegistro "albaran", "numalbar = " & Text1(0).Text
        '++monica
        Conn.BeginTrans
        
        Conn.Execute sql
        
        bol = True
        If bol Then
            MenError = "Actualizar Variedades"
            bol = ActualizarVariedades(Text1(0), Text1(1))
        End If

        If bol Then
            MenError = "Actualizar Costes"
            bol = ActualizarCostes(Text1(0), Text1(1), True, DBLet(Data1.Recordset!codforfait, "T"), DBLet(Data1.Recordset!CodPalet, "N"))
        End If
        
    End If
    
    ModoLineas = 0
    PosicionarData
    
Error2:
    If Err.Number <> 0 Or bol = False Then
        Screen.MousePointer = vbDefault
        Conn.RollbackTrans
        MuestraError Err.Number, "Eliminando linea" & MenError, Err.Description
    Else
        Conn.CommitTrans
        
        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
        PonerCadenaBusqueda
'--monica:02102008
'        ' *** si n'hi han tabs sense datagrid, posar l'If ***
'        CargaGrid Index, True
'        If Not SituarDataTrasEliminar(AdoAux(Index), NumRegElim, True) Then
''            PonerCampos
'
'        End If
'        CalcularTotales
'--monica
        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
'--monica:02102008
'            BotonModificar
'--monica
        End If
        ' *** si n'hi han tabs ***
'        SituarTab (NumTabMto + 1)
    
    End If
End Sub

Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vtabla As String
Dim anc As Single
Dim I As Integer
    
    ModoLineas = 1 'Posem Modo Afegir Ll�nia
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Cap�alera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index
    
    ' *** bloquejar la clau primaria de la cap�alera ***
    BloquearTxt Text1(0), True
    BloquearTxt Text1(1), True
    

    ' *** posar el nom del les distintes taules de ll�nies ***
    Select Case Index
        Case 0: vtabla = "albaran_calibre"
    End Select
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 0, 1 ' *** pose els index dels tabs de ll�nies que tenen datagrid ***
            ' *** canviar la clau primaria de les ll�nies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            
'            If Index = 1 Then NumF = SugerirCodigoSiguienteStr(vTabla, "codcoste", vWhere)

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
                Case 0 'calibres
                    txtAux(0).Text = Text1(0).Text 'numalbar
                    txtAux(3).Text = Text1(1).Text 'numlinea
                    txtAux(5).Text = SugerirCodigoSiguienteStr("albaran_calibre", "numline1", "numalbar = " & Text1(0).Text & " and numlinea =  " & Text1(1).Text) 'numline1
                    txtAux(4).Text = Text1(2).Text
                    
                    BloquearTxt txtAux(1), False
                    
                    txtAux(1).Text = ""
                    txtAux(2).Text = ""
                    txtAux(6).Text = ""
                    txtAux(7).Text = ""
                    txtAux(8).Text = ""
                    txtAux(9).Text = ""
                    
                    txtAux2(2).Text = ""
                    txtAux2(0).Text = ""
                    BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux0"
                    PonerFoco txtAux(1)
            End Select
            
    End Select
End Sub

Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim I As Integer
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
  
    Select Case Index
        Case 0, 1 ' *** pose els index de ll�nies que tenen datagrid (en o sense tab) ***
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
        Case 0 ' calibres
        
            txtAux(0).Text = DataGridAux(Index).Columns(0).Text
            txtAux(3).Text = DataGridAux(Index).Columns(1).Text
            txtAux(5).Text = DataGridAux(Index).Columns(2).Text
            txtAux(4).Text = DataGridAux(Index).Columns(3).Text
            txtAux(1).Text = DataGridAux(Index).Columns(4).Text
            txtAux(2).Text = DataGridAux(Index).Columns(6).Text
            txtAux(6).Text = DataGridAux(Index).Columns(8).Text
            txtAux(7).Text = DataGridAux(Index).Columns(9).Text
            txtAux(8).Text = DataGridAux(Index).Columns(7).Text
            txtAux2(2).Text = DataGridAux(Index).Columns(5).Text
            txtAux2(0).Text = DataGridAux(Index).Columns(10).Text
            
            txtAux(9).Text = DataGridAux(Index).Columns(11).Text
            
            
            For I = 1 To 1
                BloquearTxt txtAux(I), True
            Next I
            BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux0"
            
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 0 'calibres
            PonerFoco txtAux(2)
    End Select
    ' ***************************************************************************************
End Sub

Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim b As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    DeseleccionaGrid DataGridAux(Index)
       
    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Ll�nies
    Select Case Index
        Case 0 'calibres
            txtAux(1).visible = b 'numalbar
            txtAux(1).Top = alto
            txtAux(2).visible = b 'numlinea
            txtAux(2).Top = alto
            txtAux(6).visible = b
            txtAux(6).Top = alto
            txtAux(7).visible = b
            txtAux(7).Top = alto
            txtAux2(2).visible = b
            txtAux2(2).Top = alto
            btnBuscar(0).visible = b
            btnBuscar(0).Top = alto
            txtAux(8).visible = b
            txtAux(8).Top = alto
            txtAux(9).visible = b
            txtAux(9).Top = alto
            txtAux2(0).visible = b
            txtAux2(0).Top = alto
            
    End Select
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
Dim Forfait As String
Dim sql As String
Dim KilosUni As Currency

    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    

    ' ******* configurar el LostFocus dels camps de ll�nies (dins i fora grid) ********
    Select Case Index
        Case 1 ' codigo de calibre
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(2).Text = DevuelveDesdeBDNew(cAgro, "calibres", "nomcalib", "codvarie", txtAux(4).Text, "N", , "codcalib", txtAux(1).Text, "N")
                If txtAux2(2).Text = "" Then
                    cadMen = "No existe el Calibre: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "�Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCali = New frmManCalibres
                        frmCali.DatosADevolverBusqueda = "0|2|3|"
                        frmCali.NuevoCodigo = txtAux(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        '++monica
                        
                        frmCali.Show vbModal
                        Set frmCali = Nothing
'                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                        BloqueaRegistro "albaran", "numalbar = " & Text1(0).Text
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                End If
            Else
                txtAux2(2).Text = ""
            End If
        
        
        Case 2 ' cajas
            If txtAux(Index).Text <> "" Then PonerFormatoEntero txtAux(Index)
            
            '[Monica]27/01/2014: preguntamos si recalculamos pesos solo si es Montifrut
            If vParamAplic.Cooperativa = 12 Then
                If MsgBox("� Desea calcular peso neto ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    CalculoPesoNeto
                End If
            Else
                CalculoPesoNeto
            End If
            
        Case 6 ' peso bruto
            If txtAux(Index).Text <> "" Then PonerFormatoEntero txtAux(Index)
            
        Case 7 'peso neto
            If txtAux(Index).Text <> "" Then
'                If PonerFormatoEntero(txtAux(Index)) Then cmdAceptar.SetFocus
                PonerFormatoEntero txtAux(Index)
            End If
                
        
        Case 8 'unidades
            ' en el caso de que metan unidades el pesoneto = unidades * forfaits.kilosuni
            If txtAux(Index).Text <> "" Then
                PonerFormatoEntero txtAux(Index)
                Forfait = DevuelveDesdeBDNew(cAgro, "albaran_variedad", "codforfait", "numalbar", Data1.Recordset!NumAlbar, "N", , "numlinea", Data1.Recordset!NumLinea, "N")
                sql = DevuelveDesdeBDNew(cAgro, "forfaits", "kilosuni", "codforfait", Forfait, "T")
                If sql <> "" Then
                    txtAux(7).Text = Round2(ImporteSinFormato(sql) * txtAux(Index), 0)
                    PonerFormatoEntero txtAux(7)
                End If
            End If
    
        Case 9 'precio provisional
            If PonerFormatoDecimal(txtAux(Index), 8) Or txtAux(9).Text = "" Then cmdAceptar.SetFocus
    
    
    End Select
    
    If txtAux(7).Text <> "" And txtAux(2).Text <> "" Then
        If ComprobarCero(txtAux(2).Text) <> 0 Then
            txtAux2(0).Text = Round2(ImporteSinFormato(txtAux(7).Text) / ImporteSinFormato(txtAux(2).Text), 2)
            txtAux2(0).Text = Format(txtAux2(0).Text, "###,##0.00")
        End If
    End If
    
    
End Sub

Private Sub CalculoPesoNeto()
Dim sql As String
Dim Kilos1 As Currency
Dim Kilos2 As Currency
Dim Rs As ADODB.Recordset

    sql = "select kiloscaj, kilosuni  from forfaits where codforfait = " & DBSet(Text1(5).Text, "T")
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Kilos1 = 0
    Kilos2 = 0
    If Not Rs.EOF Then
        Kilos1 = DBLet(Rs!kiloscaj, "N")
        Kilos2 = DBLet(Rs!KilosUni, "N")
    End If
    
    'si hay cajas
    If ComprobarCero(txtAux(2).Text) <> 0 Then
        If Kilos1 <> 0 Then
            txtAux(6).Text = Round2(Kilos1 * ImporteSinFormato(txtAux(2).Text), 0)
            PonerFormatoEntero txtAux(6)
            txtAux(7).Text = txtAux(6).Text
        End If
    End If
    'si hay unidades
    If ComprobarCero(txtAux(8).Text) <> 0 Then
        If Kilos2 <> 0 Then
            txtAux(6).Text = Round2(Kilos2 * ImporteSinFormato(txtAux(8).Text), 0)
            PonerFormatoEntero txtAux(6)
            txtAux(7).Text = txtAux(6).Text
        End If
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
Dim sql As String
Dim b As Boolean
Dim Cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte

    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False
        
    b = CompForm2(Me, 2, nomFrame) 'Comprovar formato datos ok
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

Private Sub imgBuscar_Click(Index As Integer)
    TerminaBloquear
    '++monica
'    BloqueaRegistro "palets", "numpalet = " & Text1(0).Text
    
     indice = Index + 2
     Select Case Index
        Case 0, 1 'variedad y variedad comercial
            indice = Index + 2
            Set frmVar = New frmManVariedad
            frmVar.DatosADevolverBusqueda = "0|1|"
            frmVar.CodigoActual = Text1(indice).Text
            frmVar.Show vbModal
            Set frmVar = Nothing
            PonerFoco Text1(indice)
        Case 2 'Marca
            Set frmMar = New frmManMarcas
            frmMar.DatosADevolverBusqueda = "0|1|"
            frmMar.CodigoActual = Text1(4).Text
            frmMar.Show vbModal
            Set frmMar = Nothing
            PonerFoco Text1(4)
        Case 3 'forfait
            Set frmFor = New frmManForfaits
            frmFor.DatosADevolverBusqueda = "0|1|"
            frmFor.CodigoActual = Text1(5).Text
            frmFor.Show vbModal
            Set frmFor = Nothing
            PonerFoco Text1(5)
        Case 4 'incidencia
            indice = 13
            Set frmIncid = New frmManInciden
            frmIncid.DatosADevolverBusqueda = "0|1|"
            frmIncid.CodigoActual = Text1(13).Text
            frmIncid.Show vbModal
            Set frmIncid = Nothing
            PonerFoco Text1(13)
        Case 5 'codigo de palet
            indice = 18
            PonerFoco Text1(indice)
            Set frmPal = New frmManPaleConf
            frmPal.DatosADevolverBusqueda = "0|1|"
            frmPal.Show vbModal
            Set frmPal = Nothing
            PonerFoco Text1(indice)
            
        Case 6 'comisionista
            PonerFoco Text1(20)
            Set frmTra1 = New frmManAgencias
            frmTra1.DatosADevolverBusqueda = "0|1|2|"
            frmTra1.Show vbModal
            Set frmTra1 = Nothing
            PonerFoco Text1(20)
            
    End Select
    
    If Modo = 4 Then BloqueaRegistro "albaran", "numalbar = " & Text1(0).Text
                'BLOQUEADesdeFormulario2 Me, Data1, 1
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

    AdoAux(Index).ConnectionString = Conn
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
Dim I As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)

    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
    
    Select Case Index
        Case 0 'calibres
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;" 'numalbar,numlinea,numline1,codvarie
            tots = tots & "S|txtAux(1)|T|Calibre|1000|;S|btnBuscar(0)|B|||;"
            tots = tots & "S|txtAux2(2)|T|Denominaci�n|2000|;S|txtAux(2)|T|Cajas|1000|;S|txtAux(8)|T|Uds|900|;S|txtAux(6)|T|Peso Bruto|1000|;S|txtAux(7)|T|Peso Neto|1000|;S|txtAux2(0)|T|Kilos/Caja|1100|;S|txtAux(9)|T|Pr.Prov|900|;"
            
            arregla tots, DataGridAux(Index), Me
        
'            DataGridAux(0).Columns(6).NumberFormat = "#,###0"
            DataGridAux(0).Columns(6).Alignment = dbgRight
            DataGridAux(0).Columns(10).Alignment = dbgRight
            DataGridAux(0).Columns(10).NumberFormat = "###,##0.00"
            
        
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

Private Function InsertarLinea() As Boolean
'Inserta registre en les taules de Ll�nies
Dim nomFrame As String
Dim bol As Boolean
Dim MenError As String
Dim PesoNeto As String
Dim NumCajas As String

    On Error GoTo EInsertarLinea

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomFrame = "FrameAux0" 'calibres
    End Select
    
    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        '++monica
        BloqueaRegistro "albaran", "numalbar = " & Text1(0).Text
        
        'Aqui empieza transaccion
        Conn.BeginTrans
        
        bol = InsertarDesdeForm2(Me, 2, nomFrame)
        If bol Then
            MenError = "Modificando variedades"
            bol = ActualizarVariedades(Text1(0), Text1(1))
        End If
        
        If bol Then
            PesoNeto = ""
            PesoNeto = DevuelveDesdeBDNew(cAgro, "albaran_variedad", "pesoneto", "numalbar", DBLet(Data1.Recordset!NumAlbar, "N"), "N", , "numlinea", DBLet(Data1.Recordset!NumLinea, "N"), "N")
            NumCajas = ""
            NumCajas = DevuelveDesdeBDNew(cAgro, "albaran_variedad", "numcajas", "numalbar", DBLet(Data1.Recordset!NumAlbar, "N"), "N", , "numlinea", DBLet(Data1.Recordset!NumLinea, "N"), "N")
            
            If CCur(ComprobarCero(PesoNeto)) <> KilosAnt Or CCur(ComprobarCero(NumCajas)) <> CajasAnt Then
                MenError = "Actualizar Costes"
                bol = ActualizarCostes(Text1(0), Text1(1), True, DBLet(Data1.Recordset!codforfait, "T"), DBLet(Data1.Recordset!CodPalet, "N"))
            End If
        End If
'
'            b = BLOQUEADesdeFormulario2(Me, Data1, 1)
'            Select Case NumTabMto
'                Case 0, 1 ' *** els index de les llinies en grid (en o sense tab) ***
'                    CargaGrid NumTabMto, True
'                    If b Then BotonAnyadirLinea NumTabMto
'            End Select
'
'            SituarTab (NumTabMto + 1)
    Else
        InsertarLinea = False
        Exit Function
    End If

EInsertarLinea:
        If Err.Number <> 0 Then
            MenError = "Insertando Linea." & vbCrLf & "----------------------------" & vbCrLf & MenError
            MuestraError Err.Number, MenError, Err.Description
            bol = False
        End If
        If bol Then
            Conn.CommitTrans
            InsertarLinea = True
        Else
            Conn.RollbackTrans
            InsertarLinea = False
        End If
End Function

Private Function ModificarLinea() As Boolean
'Modifica registre en les taules de Ll�nies
Dim nomFrame As String
Dim V As Integer
Dim bol As Boolean
Dim MenError As String
Dim PesoNeto As String
Dim NumCajas As String
    
    On Error GoTo eModificarLinea

    ' *** posa els noms del frames, tant si son de grid com si no ***
    nomFrame = "FrameAux0" 'calibres
    
    ModificarLinea = False
    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        
        Conn.BeginTrans
        
        bol = ModificaDesdeFormulario2(Me, 2, nomFrame)
        If bol Then
            MenError = "Modificando variedades"
            bol = ActualizarVariedades(Text1(0), Text1(1))
        End If
        
        
        If bol Then
            MenError = "Actualizando Precio Provisional"
            bol = ActualizarPrecioProv(Text1(0), Text1(1))
        End If
        
        If bol Then
            PesoNeto = ""
            PesoNeto = DevuelveDesdeBDNew(cAgro, "albaran_variedad", "pesoneto", "numalbar", DBLet(Data1.Recordset!NumAlbar, "N"), "N", , "numlinea", DBLet(Data1.Recordset!NumLinea, "N"), "N")
            NumCajas = ""
            NumCajas = DevuelveDesdeBDNew(cAgro, "albaran_variedad", "numcajas", "numalbar", DBLet(Data1.Recordset!NumAlbar, "N"), "N", , "numlinea", DBLet(Data1.Recordset!NumLinea, "N"), "N")
            
            If CCur(ComprobarCero(PesoNeto)) <> KilosAnt Or CCur(ComprobarCero(NumCajas)) <> CajasAnt Then
                MenError = "Actualizar Costes"
                bol = ActualizarCostes(Text1(0), Text1(1), True, DBLet(Data1.Recordset!codforfait, "T"), DBLet(Data1.Recordset!CodPalet, "N"))
            End If
        End If
        
'            ModoLineas = 0
'
'            V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el n� de llinia
'
'            CargaGrid NumTabMto, True
'
'            ' *** si n'hi han tabs ***
''            SituarTab (NumTabMto + 1)
'
'            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
'            PonerFocoGrid Me.DataGridAux(NumTabMto)
'            AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
'
'            LLamaLineas NumTabMto, 0
'            ModificarLinea = True
'        End If
        
        '++monica
'        BloqueaRegistro "pedidos", "numpedid = " & Text1(0).Text
        
    Else
        Exit Function
    End If

eModificarLinea:
    If Err.Number <> 0 Then
        MenError = "Modificando Linea." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        bol = False
    End If
    If bol Then
        Conn.CommitTrans
        ModificarLinea = True
    Else
        Conn.RollbackTrans
        ModificarLinea = False
    End If
End Function

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la cap�alera ***
    vWhere = vWhere & " numalbar=" & Me.Data1.Recordset!NumAlbar & " and numlinea = " & Me.Data1.Recordset!NumLinea
    
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
Dim sql As String
Dim TotalEnvases As String
Dim TotalCostes As String
Dim Valor As Currency

    On Error Resume Next

    'total importes de envases para ese forfait
    sql = "select sum(numcajas) "
    sql = sql & " from albaran_calibre where numalbar = " & DBSet(Text1(0).Text, "N")
    sql = sql & " and numlinea = " & DBSet(Text1(1).Text, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    TotalEnvases = 0
    If Not Rs.EOF Then
        If Rs.Fields(0).Value > 0 Then TotalEnvases = Rs.Fields(0).Value
    End If
    Rs.Close
    Set Rs = Nothing
    
'    Text3(0).Text = Format(TotalEnvases, "###,##0")
    If Err.Number <> 0 Then
        Err.Clear
    End If

End Sub

Private Function ObtenerWhereCP(conW As Boolean) As String
Dim sql As String
On Error Resume Next
    
    sql = ""
    If conW Then sql = " WHERE "
    sql = sql & NombreTabla & ".numalbar= " & DBSet(Text1(0).Text, "N")
    sql = sql & " and " & NombreTabla & ".numlinea=" & Val(Text1(1).Text)
    ObtenerWhereCP = sql
End Function



Private Function ActualizarVariedades(Albaran As String, Linea As String) As Boolean
Dim sql As String
Dim Rs As ADODB.Recordset
Dim SQL1 As String
Dim PrecioMedioProv As Currency

    On Error GoTo eActualizarVariedades

    ActualizarVariedades = False

    SQL1 = "select sum(pesobrut), sum(pesoneto), sum(numcajas), sum(unidades) from albaran_calibre where numalbar = " & DBSet(Albaran, "N")
    SQL1 = SQL1 & " and numlinea = " & DBSet(Linea, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL1, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        If DBLet(Rs.Fields(0).Value, "N") = 0 Then
            sql = "update albaran_variedad set pesobrut = null "
            sql = sql & " where numalbar = " & DBSet(Albaran, "N")
            sql = sql & " and numlinea = " & DBSet(Linea, "N")
    
            Conn.Execute sql
        End If
        If DBLet(Rs.Fields(1).Value, "N") = 0 Then
            sql = "update albaran_variedad set pesoneto = null "
            sql = sql & " where numalbar = " & DBSet(Albaran, "N")
            sql = sql & " and numlinea = " & DBSet(Linea, "N")
    
            Conn.Execute sql
        End If
        If DBLet(Rs.Fields(2).Value, "N") = 0 Then
            sql = "update albaran_variedad set numcajas = null "
            sql = sql & " where numalbar = " & DBSet(Albaran, "N")
            sql = sql & " and numlinea = " & DBSet(Linea, "N")
    
            Conn.Execute sql
        End If
        If DBLet(Rs.Fields(3).Value, "N") = 0 Then
            sql = "update albaran_variedad set unidades = null "
            sql = sql & " where numalbar = " & DBSet(Albaran, "N")
            sql = sql & " and numlinea = " & DBSet(Linea, "N")
    
            Conn.Execute sql
        End If
        
        If DBLet(Rs.Fields(0).Value, "N") <> 0 Then
            sql = "update albaran_variedad set pesobrut = " & DBSet(Rs.Fields(0).Value, "N")
            sql = sql & " where numalbar = " & DBSet(Albaran, "N")
            sql = sql & " and numlinea = " & DBSet(Linea, "N")
    
            Conn.Execute sql
        End If
        
        If DBLet(Rs.Fields(1).Value, "N") <> 0 Then
            sql = "update albaran_variedad set pesoneto = " & DBSet(Rs.Fields(1).Value, "N")
            sql = sql & " where numalbar = " & DBSet(Albaran, "N")
            sql = sql & " and numlinea = " & DBSet(Linea, "N")
    
            Conn.Execute sql
        End If
        If DBLet(Rs.Fields(2).Value, "N") <> 0 Then
            sql = "update albaran_variedad set numcajas = " & DBSet(Rs.Fields(2).Value, "N")
            sql = sql & " where numalbar = " & DBSet(Albaran, "N")
            sql = sql & " and numlinea = " & DBSet(Linea, "N")
    
            Conn.Execute sql
        End If
        If DBLet(Rs.Fields(3).Value, "N") <> 0 Then
            sql = "update albaran_variedad set unidades = " & DBSet(Rs.Fields(3).Value, "N")
            sql = sql & " where numalbar = " & DBSet(Albaran, "N")
            sql = sql & " and numlinea = " & DBSet(Linea, "N")
    
            Conn.Execute sql
        End If
    
    End If
    Rs.Close
    Set Rs = Nothing

eActualizarVariedades:
    If Err.Number = 0 Then ActualizarVariedades = True
    
End Function


Private Function ActualizarCalibres(Albaran As String, Linea As String) As Boolean
Dim sql As String
Dim Rs As ADODB.Recordset
Dim SQL1 As String

    On Error GoTo eActualizarCalibres

    ActualizarCalibres = False
    
    SQL1 = "update albaran_calibre set codvarie = " & DBSet(Text1(2).Text, "N") & " where numalbar = " & DBSet(Albaran, "N")
    SQL1 = SQL1 & " and numlinea = " & DBSet(Linea, "N")
    
    Conn.Execute SQL1

eActualizarCalibres:
    If Err.Number = 0 Then ActualizarCalibres = True
End Function







Private Function Modificar() As Boolean
'Modifica registre en les taules de Ll�nies
Dim nomFrame As String
Dim V As Integer
Dim bol As Boolean
Dim MenError As String
Dim Forfait As String
Dim CodPalet As String
Dim TotPalet As String
Dim LOG As cLOG
Dim campo As String

    On Error GoTo EModificar

    TerminaBloquear
    
    Conn.BeginTrans
    
    bol = ModificaDesdeFormulario2(Me, 1)
    
    If bol Then
        MenError = "Modificando variedades"
        bol = ActualizarVariedades(CStr(Albaran), CStr(Linea))
        
        
        '[Monica]01/10/2012: dejo modificar la variedad pero he de cambiarselo a los calibres
        '                    Solo se permite modificar la variedad si la linea no est� facturada
        If bol Then
            If Text1(2).Text <> VarieAnt Then
                '------------------------------------------------------------------------------
                '  LOG de acciones.
                Set LOG = New cLOG
                campo = "Var.Ant: " & VarieAnt & " Nueva:" & CLng(Text1(2).Text)
                LOG.Insertar 10, vUsu, campo
                Set LOG = Nothing
                '-----------------------------------------------------------------------------
                
                MenError = "Modificando calibres"
                bol = ActualizarCalibres(CStr(Albaran), CStr(Linea))
            End If
        End If
        
        If bol Then
            Forfait = ""
            Forfait = DevuelveDesdeBDNew(cAgro, "albaran_variedad", "codforfait", "numalbar", CStr(Albaran), "N", , "numlinea", CStr(Linea), "N")
            CodPalet = ""
            CodPalet = DevuelveDesdeBDNew(cAgro, "albaran_variedad", "codpalet", "numalbar", CStr(Albaran), "N", , "numlinea", CStr(Linea), "N")
            TotPalet = ""
            TotPalet = DevuelveDesdeBDNew(cAgro, "albaran_variedad", "totpalet", "numalbar", CStr(Albaran), "N", , "numlinea", CStr(Linea), "N")
            
            
            If Forfait <> ForfaitAnt Or CodPalet <> CodPaletAnt Or TotPalet <> TotPaletAnt Then
                MenError = "Actualizar Costes"
                bol = ActualizarCostes(Albaran, Linea, True, ForfaitAnt, CodPaletAnt)
            End If
        End If
    End If

EModificar:
    If Err.Number <> 0 Then
        MenError = "Modificando Registro Albar�n Variedad." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        bol = False
    End If
    If bol Then
        Conn.CommitTrans
        Modificar = True
    Else
        Conn.RollbackTrans
        Modificar = False
    End If
End Function

Private Sub CargaCombo()
Dim cad As String
Dim Rs As ADODB.Recordset

    On Error GoTo ErrCarga
    
    Combo1.Clear
    
    cad = "SELECT * FROM tipovarie ORDER BY codtipo"
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    'Combo1.AddItem "" 'pose uno en blanc sinse valor
    While Not Rs.EOF
        Combo1.AddItem Rs!nomtipo
        Combo1.ItemData(Combo1.NewIndex) = Rs!codtipo
        Rs.MoveNext
        '.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    Exit Sub
    
ErrCarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar datos combo.", Err.Description
End Sub



Private Function ActualizarPrecioProv(Albaran As String, Linea As String)
Dim Rs As ADODB.Recordset
Dim sql As String
Dim SQL1 As String
Dim PrecioMedioProv As Currency

    On Error GoTo eActualizarPrecioProv

    ActualizarPrecioProv = False

    SQL1 = "select * from albaran_calibre where numalbar = " & DBSet(Albaran, "N")
    SQL1 = SQL1 & " and numlinea = " & DBSet(Linea, "N")
    SQL1 = SQL1 & " and not preciopro is null and preciopro <> 0"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL1, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        '[Monica]06/06/2013: en el caso de que me hayan metido un precio provisional recalculo
        sql = "select round(sum( if(preciopro is null, 0, preciopro) * if (pesoneto is null, 0, pesoneto) ) / sum(if(pesoneto is null,0,pesoneto)),4) "
        sql = sql & " from albaran_calibre "
        sql = sql & " where numalbar = " & DBSet(Albaran, "N")
        sql = sql & " and numlinea = " & DBSet(Linea, "N")
    
        PrecioMedioProv = DevuelveValor(sql)
    
        sql = "update albaran_variedad set preciopro = " & DBSet(PrecioMedioProv, "N")
        sql = sql & " where numalbar = " & DBSet(Albaran, "N")
        sql = sql & " and numlinea = " & DBSet(Linea, "N")

        Conn.Execute sql
    End If
    Rs.Close
    Set Rs = Nothing

eActualizarPrecioProv:
    If Err.Number = 0 Then ActualizarPrecioProv = True
End Function

