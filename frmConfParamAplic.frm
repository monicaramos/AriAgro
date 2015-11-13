VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConfParamAplic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parámetros de la Aplicación"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   9075
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfParamAplic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5685
      Left            =   180
      TabIndex        =   55
      Top             =   675
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   10028
      _Version        =   393216
      Tabs            =   6
      Tab             =   3
      TabsPerRow      =   6
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Contabilidad"
      TabPicture(0)   =   "frmConfParamAplic.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "imgBuscar(2)"
      Tab(0).Control(1)=   "Label1(2)"
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(3)=   "Frame6"
      Tab(0).Control(4)=   "Text1(17)"
      Tab(0).Control(5)=   "Text2(17)"
      Tab(0).Control(6)=   "Frame3"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Internet"
      TabPicture(1)   =   "frmConfParamAplic.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "Frame7"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Varios"
      TabPicture(2)   =   "frmConfParamAplic.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1(0)"
      Tab(2).Control(1)=   "Label1(1)"
      Tab(2).Control(2)=   "Label1(29)"
      Tab(2).Control(3)=   "Label1(30)"
      Tab(2).Control(4)=   "Label1(39)"
      Tab(2).Control(5)=   "Label1(40)"
      Tab(2).Control(6)=   "Label1(41)"
      Tab(2).Control(7)=   "Label1(42)"
      Tab(2).Control(8)=   "imgBuscar(10)"
      Tab(2).Control(9)=   "Label1(44)"
      Tab(2).Control(10)=   "Label1(45)"
      Tab(2).Control(11)=   "Label1(46)"
      Tab(2).Control(12)=   "Label1(47)"
      Tab(2).Control(13)=   "imgAyuda(0)"
      Tab(2).Control(14)=   "Combo1(0)"
      Tab(2).Control(15)=   "chkInventar"
      Tab(2).Control(16)=   "chkctrstock"
      Tab(2).Control(17)=   "Frame5"
      Tab(2).Control(18)=   "Text1(15)"
      Tab(2).Control(19)=   "Text1(24)"
      Tab(2).Control(20)=   "Text1(25)"
      Tab(2).Control(21)=   "Text1(30)"
      Tab(2).Control(22)=   "Text1(31)"
      Tab(2).Control(23)=   "Combo1(13)"
      Tab(2).Control(24)=   "Text1(32)"
      Tab(2).Control(25)=   "Text2(32)"
      Tab(2).Control(26)=   "Text1(34)"
      Tab(2).Control(27)=   "Text1(35)"
      Tab(2).Control(28)=   "CheckPase"
      Tab(2).Control(29)=   "Text1(36)"
      Tab(2).Control(30)=   "Text1(37)"
      Tab(2).Control(31)=   "CheckReferencia"
      Tab(2).Control(32)=   "CheckCalculo"
      Tab(2).ControlCount=   33
      TabCaption(3)   =   "Aridoc"
      TabPicture(3)   =   "frmConfParamAplic.frx":0060
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label1(6)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "imgBuscar(6)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label1(7)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "imgBuscar(7)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label1(28)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "imgBuscar(8)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Text2(21)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Text1(21)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Text2(22)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Text1(22)"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Frame8"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "Frame9"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "Text1(23)"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "Text2(23)"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).ControlCount=   14
      TabCaption(4)   =   "Edicom"
      TabPicture(4)   =   "frmConfParamAplic.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label1(31)"
      Tab(4).Control(1)=   "Label1(32)"
      Tab(4).Control(2)=   "Label1(33)"
      Tab(4).Control(3)=   "Text1(26)"
      Tab(4).Control(4)=   "Text1(27)"
      Tab(4).Control(5)=   "Text1(28)"
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "Otros"
      TabPicture(5)   =   "frmConfParamAplic.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "imgAyuda(1)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Label34"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Label1(52)"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Frame11"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "Frame12"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "Text1(45)"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "Frame13"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "Text1(47)"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).ControlCount=   8
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   47
         Left            =   -73380
         MaxLength       =   255
         TabIndex        =   152
         Tag             =   "Path FacturaE|T|S|||sparam|pathfacturae|||"
         Top             =   4110
         Width           =   6120
      End
      Begin VB.Frame Frame13 
         Caption         =   "Datos de Costes"
         ForeColor       =   &H00972E0B&
         Height          =   855
         Left            =   -74700
         TabIndex        =   160
         Top             =   4500
         Width           =   8205
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   46
            Left            =   1320
            MaxLength       =   250
            TabIndex        =   153
            Tag             =   "Path Fichadas|T|S|||sparam|pathfichadas|||"
            Top             =   360
            Width           =   6135
         End
         Begin VB.Label Label3 
            Caption         =   "Path Fichadas"
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
            Left            =   240
            TabIndex        =   161
            Top             =   390
            Width           =   1185
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
         Height          =   285
         Index           =   45
         Left            =   -73380
         MaxLength       =   14
         TabIndex        =   151
         Tag             =   "Número GNN|N|S|||rparam|numeroggn|#############0||"
         Text            =   "12345678901234"
         Top             =   3750
         Width           =   1380
      End
      Begin VB.CheckBox CheckCalculo 
         Caption         =   "Cálculo comisiones sobre kilos reales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70440
         TabIndex        =   20
         Tag             =   "Tipo Cálculo Dtos y Comisiones|N|N|0|1|sparam|calculocomision|||"
         Top             =   810
         Width           =   3015
      End
      Begin VB.CheckBox CheckReferencia 
         Caption         =   "Pase Referencia linea albaran a Arimoney "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70440
         TabIndex        =   22
         Tag             =   "Pase Rel.Linea Arimoney|N|N|0|1|sparam|pasereflineaalb|||"
         Top             =   1620
         Width           =   3585
      End
      Begin VB.Frame Frame12 
         Caption         =   "Ventas a Socios"
         ForeColor       =   &H00972E0B&
         Height          =   705
         Left            =   -74700
         TabIndex        =   155
         Top             =   2970
         Width           =   8205
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
            Height          =   320
            Index           =   43
            Left            =   1305
            MaxLength       =   6
            TabIndex        =   150
            Tag             =   "Cliente Ventas|N|S|||sparam|codclien|000000||"
            Text            =   "Text1"
            Top             =   300
            Width           =   840
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   43
            Left            =   2205
            TabIndex        =   156
            Top             =   300
            Width           =   5235
         End
         Begin VB.Label Label1 
            Caption         =   "Cliente "
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
            Index           =   51
            Left            =   240
            TabIndex        =   157
            Top             =   345
            Width           =   720
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   15
            Left            =   990
            ToolTipText     =   "Buscar Cliente"
            Top             =   345
            Width           =   240
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "CMR"
         ForeColor       =   &H00972E0B&
         Height          =   2415
         Left            =   -74700
         TabIndex        =   146
         Top             =   480
         Width           =   8175
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Index           =   41
            Left            =   270
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   147
            Tag             =   "Texto 1 CMR 13|T|S|||sparam|text1cmr13|||"
            Top             =   570
            Width           =   7545
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Index           =   42
            Left            =   270
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   149
            Tag             =   "Texto 2 CMR 13|T|S|||sparam|text2cmr13|||"
            Top             =   1620
            Width           =   7545
         End
         Begin VB.Label Label33 
            Caption         =   "Texto 1 CMR Seccion 13"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   270
            TabIndex        =   154
            Top             =   270
            Width           =   1845
         End
         Begin VB.Image imgZoom 
            Height          =   240
            Index           =   3
            Left            =   2160
            ToolTipText     =   "Zoom descripción"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Image imgZoom 
            Height          =   240
            Index           =   2
            Left            =   2160
            ToolTipText     =   "Zoom descripción"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label32 
            Caption         =   "Texto 2 CMR Seccion 13"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   270
            TabIndex        =   148
            Top             =   1320
            Width           =   1875
         End
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   37
         Left            =   -67950
         MaxLength       =   3
         TabIndex        =   37
         Tag             =   "Tipo Mov Albaranes|T|N|||sparam|codtipomalb|||"
         Text            =   "Tex"
         Top             =   5010
         Width           =   570
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   36
         Left            =   -72870
         MaxLength       =   15
         TabIndex        =   32
         Tag             =   "Nro Poliza Expediente|T|S|||sparam|nropolizaexp|||"
         Text            =   "Text"
         Top             =   5040
         Width           =   1740
      End
      Begin VB.CheckBox CheckPase 
         Caption         =   "Pase Pedido a Albarán desde Palets. Línea de Variedad por Calibre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70440
         TabIndex        =   21
         Tag             =   "Pase Agrup.Calibre|N|N|0|1|sparam|pasealbaragrupcalib|||"
         Top             =   1230
         Width           =   3015
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
         Height          =   320
         Index           =   35
         Left            =   -67950
         MaxLength       =   4
         TabIndex        =   36
         Tag             =   "Registro OPA|N|S|||sparam|regisopa|###0||"
         Text            =   "Text1"
         Top             =   4170
         Width           =   840
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
         Height          =   320
         Index           =   34
         Left            =   -67950
         MaxLength       =   4
         TabIndex        =   35
         Tag             =   "Registro Coop Sat|T|S|||sparam|regcopsa|||"
         Text            =   "Text1"
         Top             =   3780
         Width           =   840
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   32
         Left            =   -71970
         TabIndex        =   116
         Top             =   4620
         Width           =   5235
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
         Height          =   320
         Index           =   32
         Left            =   -72870
         MaxLength       =   4
         TabIndex        =   31
         Tag             =   "Almacén|N|N|||sparam|codalmac|###0||"
         Text            =   "Text1"
         Top             =   4620
         Width           =   840
      End
      Begin VB.ComboBox Combo1 
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
         Index           =   13
         Left            =   -67965
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Tag             =   "Portes Kilos/Caja|N|N|||sparam|porteskilocaja||N|"
         Top             =   3405
         Width           =   1260
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
         Height          =   320
         Index           =   31
         Left            =   -67965
         MaxLength       =   4
         TabIndex        =   33
         Tag             =   "%Desv|N|N|0|100|sparam|porcdesvia|#0.00||"
         Text            =   "Text1"
         Top             =   3000
         Width           =   840
      End
      Begin VB.Frame Frame3 
         Caption         =   "Facturas Transporte/Comisionistas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1380
         Left            =   -74505
         TabIndex        =   62
         Top             =   2175
         Width           =   7665
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
            Height          =   285
            Index           =   18
            Left            =   2250
            MaxLength       =   10
            TabIndex        =   6
            Tag             =   "Cta.Comis.Retencion|T|S|||sparam|ctacomreten|||"
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   18
            Left            =   3525
            TabIndex        =   120
            Top             =   960
            Width           =   3570
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
            Height          =   285
            Index           =   16
            Left            =   2265
            MaxLength       =   10
            TabIndex        =   5
            Tag             =   "Cta.Trans.Retencion|T|S|||sparam|ctatrareten|||"
            Top             =   630
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   16
            Left            =   3525
            TabIndex        =   78
            Top             =   630
            Width           =   3570
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   3525
            TabIndex        =   63
            Top             =   300
            Width           =   3570
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
            Height          =   285
            Index           =   5
            Left            =   2250
            MaxLength       =   10
            TabIndex        =   4
            Tag             =   "Iva Tranporte|T|S|||sparam|codivatrans|0||"
            Top             =   300
            Width           =   1215
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   1935
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   1005
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Reten.Comis."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   540
            TabIndex        =   121
            Top             =   1005
            Width           =   1200
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   1935
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   660
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Reten.Trans."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   27
            Left            =   540
            TabIndex        =   79
            Top             =   660
            Width           =   1305
         End
         Begin VB.Label Label1 
            Caption         =   "Iva Facturas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   24
            Left            =   540
            TabIndex        =   64
            Top             =   345
            Width           =   1200
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   0
            Left            =   1935
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   300
            Width           =   240
         End
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   17
         Left            =   -71115
         TabIndex        =   111
         Top             =   2925
         Width           =   3570
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
         Height          =   285
         Index           =   17
         Left            =   -72375
         MaxLength       =   10
         TabIndex        =   110
         Tag             =   "Cta.Abono Transporte|T|S|||sparam|ctaabotrans|||"
         Top             =   2925
         Width           =   1215
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
         Height          =   320
         Index           =   30
         Left            =   -72870
         MaxLength       =   4
         TabIndex        =   30
         Tag             =   "Número Fichero|N|N|||sparam|nrofiche|###0||"
         Text            =   "Text1"
         Top             =   4215
         Width           =   840
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
         Height          =   320
         Index           =   25
         Left            =   -72840
         MaxLength       =   10
         TabIndex        =   29
         Tag             =   "Cod.Chep Comunicador|N|N|||sparam|nrocheps|0000000000||"
         Text            =   "Text1"
         Top             =   3780
         Width           =   1335
      End
      Begin VB.Frame Frame6 
         Caption         =   "Venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         Left            =   -74520
         TabIndex        =   80
         Top             =   3570
         Width           =   7665
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
            Height          =   285
            Index           =   40
            Left            =   2250
            MaxLength       =   4
            TabIndex        =   144
            Tag             =   "Cta.Ventas Fras a Cuenta|T|S|||sparam|ccostefraacta|||"
            Top             =   1500
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   40
            Left            =   3525
            TabIndex        =   143
            Top             =   1500
            Width           =   3570
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   39
            Left            =   3525
            TabIndex        =   141
            Top             =   1170
            Width           =   3570
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
            Height          =   285
            Index           =   39
            Left            =   2250
            MaxLength       =   10
            TabIndex        =   140
            Tag             =   "Cta.Ventas Fras a Cuenta|T|S|||sparam|ctaventasfraacta|||"
            Top             =   1170
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   38
            Left            =   3525
            TabIndex        =   138
            Top             =   240
            Width           =   3570
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
            Height          =   285
            Index           =   38
            Left            =   2250
            MaxLength       =   10
            TabIndex        =   7
            Tag             =   "Iva Normal|T|S|||sparam|codivanormal|0||"
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   15
            Left            =   3525
            TabIndex        =   83
            Top             =   855
            Width           =   3570
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
            Height          =   285
            Index           =   20
            Left            =   2250
            MaxLength       =   10
            TabIndex        =   9
            Tag             =   "Iva Rec.Equivalencia|T|S|||sparam|codivarecargo|0||"
            Top             =   855
            Width           =   1215
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
            Height          =   285
            Index           =   19
            Left            =   2250
            MaxLength       =   10
            TabIndex        =   8
            Tag             =   "Iva Exento|T|S|||sparam|codivaexento|0||"
            Top             =   555
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   14
            Left            =   3525
            TabIndex        =   81
            Top             =   555
            Width           =   3570
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   14
            Left            =   1935
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   1545
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Centro de Coste"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   50
            Left            =   540
            TabIndex        =   145
            Top             =   1545
            Width           =   1350
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Vta.Fras a Cta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   49
            Left            =   540
            TabIndex        =   142
            Top             =   1215
            Width           =   1350
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   13
            Left            =   1935
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   1215
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Iva Normal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   48
            Left            =   540
            TabIndex        =   139
            Top             =   285
            Width           =   1200
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   12
            Left            =   1935
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   285
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Iva Recargo Equiv."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   540
            TabIndex        =   84
            Top             =   900
            Width           =   1380
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   1935
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   900
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   1935
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Iva Exento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   540
            TabIndex        =   82
            Top             =   600
            Width           =   1200
         End
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   28
         Left            =   -72930
         MaxLength       =   70
         TabIndex        =   104
         Tag             =   "Path|T|S|||sparam|pathedicom|||"
         Text            =   "1234567890123456789012345678901234567890123456789012345678901234567890"
         Top             =   2115
         Width           =   6435
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   27
         Left            =   -72930
         MaxLength       =   17
         TabIndex        =   102
         Tag             =   "Cód.EDI Vendedor|T|S|||sparam|codigoedi|||"
         Top             =   1080
         Width           =   2610
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   26
         Left            =   -72930
         MaxLength       =   70
         TabIndex        =   103
         Tag             =   "Reg.Mercantil|T|S|||sparam|regmercantil|||"
         Text            =   "1234567890123456789012345678901234567890123456789012345678901234567890"
         Top             =   1440
         Width           =   6435
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
         Height          =   320
         Index           =   24
         Left            =   -72870
         MaxLength       =   4
         TabIndex        =   28
         Tag             =   "Número Lote|N|N|||sparam|nrolote|###0||"
         Text            =   "Text1"
         Top             =   3405
         Width           =   840
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   23
         Left            =   3735
         TabIndex        =   100
         Top             =   1710
         Width           =   4470
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
         Height          =   285
         Index           =   23
         Left            =   2475
         MaxLength       =   10
         TabIndex        =   40
         Tag             =   "Extension|N|N|||sparam|codextension|000||"
         Top             =   1710
         Width           =   1215
      End
      Begin VB.Frame Frame9 
         Caption         =   "Facturas"
         ForeColor       =   &H00972E0B&
         Height          =   1050
         Left            =   360
         TabIndex        =   94
         Top             =   3375
         Width           =   7845
         Begin VB.ComboBox Combo1 
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
            Index           =   8
            Left            =   5805
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Tag             =   "C4 Factura|N|N|||sparam|c4facaridoc||N|"
            Top             =   585
            Width           =   1710
         End
         Begin VB.ComboBox Combo1 
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
            Index           =   7
            Left            =   3915
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Tag             =   "C3 Factura|N|N|||sparam|c3facaridoc||N|"
            Top             =   585
            Width           =   1710
         End
         Begin VB.ComboBox Combo1 
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
            Index           =   6
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Tag             =   "C2 Factura|N|N|||sparam|c2facaridoc||N|"
            Top             =   585
            Width           =   1710
         End
         Begin VB.ComboBox Combo1 
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
            Index           =   5
            Left            =   90
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Tag             =   "C1 Factura|N|N|||sparam|c1facaridoc||N|"
            Top             =   585
            Width           =   1710
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   26
            Left            =   90
            TabIndex        =   98
            Top             =   315
            Width           =   1620
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   25
            Left            =   1980
            TabIndex        =   97
            Top             =   315
            Width           =   1755
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   16
            Left            =   3915
            TabIndex        =   96
            Top             =   315
            Width           =   1620
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   14
            Left            =   5805
            TabIndex        =   95
            Top             =   315
            Width           =   1305
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Albaranes"
         ForeColor       =   &H00972E0B&
         Height          =   1050
         Left            =   360
         TabIndex        =   89
         Top             =   2295
         Width           =   7845
         Begin VB.ComboBox Combo1 
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
            Left            =   5805
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Tag             =   "C4 Albaran|N|N|||sparam|c4albaridoc||N|"
            Top             =   585
            Width           =   1710
         End
         Begin VB.ComboBox Combo1 
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
            Index           =   3
            Left            =   3915
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Tag             =   "C3 Albaran|N|N|||sparam|c3albaridoc||N|"
            Top             =   585
            Width           =   1710
         End
         Begin VB.ComboBox Combo1 
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
            Index           =   1
            Left            =   90
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Tag             =   "C1 Albaran|N|N|||sparam|c1albaridoc||N|"
            Top             =   585
            Width           =   1710
         End
         Begin VB.ComboBox Combo1 
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
            Index           =   2
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Tag             =   "C2 Albaran|N|N|||sparam|c2albaridoc||N|"
            Top             =   585
            Width           =   1710
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   12
            Left            =   5805
            TabIndex        =   93
            Top             =   315
            Width           =   1305
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   10
            Left            =   3915
            TabIndex        =   92
            Top             =   315
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   1980
            TabIndex        =   91
            Top             =   315
            Width           =   1755
         End
         Begin VB.Label Label1 
            Caption         =   "Campo 1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   9
            Left            =   90
            TabIndex        =   90
            Top             =   315
            Width           =   1665
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
         Height          =   285
         Index           =   22
         Left            =   2475
         MaxLength       =   10
         TabIndex        =   39
         Tag             =   "Carpeta Facturas|N|N|||sparam|codcarpetafac|000||"
         Top             =   1305
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   22
         Left            =   3750
         TabIndex        =   87
         Top             =   1305
         Width           =   4470
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
         Height          =   285
         Index           =   21
         Left            =   2475
         MaxLength       =   10
         TabIndex        =   38
         Tag             =   "Carpeta Albaranes|N|N|||sparam|codcarpetaalb|000||"
         Top             =   900
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   21
         Left            =   3750
         TabIndex        =   85
         Top             =   900
         Width           =   4470
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
         Height          =   320
         Index           =   15
         Left            =   -72870
         MaxLength       =   7
         TabIndex        =   27
         Tag             =   "Limite Peso Bruto CMR|N|N|||sparam|limpesobrutcmr|#,###,##0||"
         Text            =   "Text1"
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Frame Frame5 
         Caption         =   "Facturación Compras"
         ForeColor       =   &H00972E0B&
         Height          =   855
         Left            =   -74640
         TabIndex        =   74
         Top             =   1950
         Width           =   7185
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
            Height          =   320
            Index           =   11
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   23
            Tag             =   "Dia 1 de pago compras|N|S|0|31|sparam|diapago1|||"
            Text            =   "Text1"
            Top             =   360
            Width           =   615
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
            Height          =   320
            Index           =   12
            Left            =   2160
            MaxLength       =   2
            TabIndex        =   24
            Tag             =   "Dia 2 de pago compras|N|S|0|31|sparam|diapago2|||"
            Text            =   "Text1"
            Top             =   360
            Width           =   615
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
            Height          =   320
            Index           =   13
            Left            =   2880
            MaxLength       =   2
            TabIndex        =   25
            Tag             =   "Dia 3 de pago compras|N|S|0|31|sparam|diapago3|||"
            Text            =   "Text1"
            Top             =   360
            Width           =   615
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
            Height          =   320
            Index           =   14
            Left            =   6120
            MaxLength       =   2
            TabIndex        =   26
            Tag             =   "Mes a no girar|N|S|0|12|sparam|mesnogir|||"
            Text            =   "Text1"
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Días de pago"
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
            Index           =   11
            Left            =   360
            TabIndex        =   76
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Mes a no girar"
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
            Index           =   13
            Left            =   4920
            TabIndex        =   75
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.CheckBox chkctrstock 
         Caption         =   "Realiza control de Stock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74595
         TabIndex        =   19
         Tag             =   "Control de Stock|N|N|||sparam|ctrstock|||"
         Top             =   1545
         Width           =   2775
      End
      Begin VB.CheckBox chkInventar 
         Caption         =   "Realizar Inventario por Proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74595
         TabIndex        =   18
         Tag             =   "Inventarios por Proveedor|N|N|||sparam|inventar|||"
         Top             =   1215
         Width           =   2775
      End
      Begin VB.ComboBox Combo1 
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
         Index           =   0
         Left            =   -73110
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Tag             =   "Precio Artículo|N|N|||sparam|tipoprecio||N|"
         Top             =   720
         Width           =   2340
      End
      Begin VB.Frame Frame7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -74595
         TabIndex        =   67
         Top             =   600
         Width           =   8010
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   44
            Left            =   2730
            MaxLength       =   30
            TabIndex        =   14
            Tag             =   "LanzaMailOutlook|T|S|||sparam|arigesmail|||"
            Text            =   "3"
            Top             =   2010
            Width           =   1620
         End
         Begin VB.CheckBox chkOutlook 
            Caption         =   "Enviar desde Outlook"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5310
            TabIndex        =   15
            Tag             =   "Outlook|N|N|||sparam|EnvioDesdeOutlook|||"
            Top             =   1980
            Width           =   2175
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   1260
            MaxLength       =   50
            TabIndex        =   10
            Tag             =   "Direccion e-mail|T|S|||sparam|diremail|||"
            Text            =   "3"
            Top             =   450
            Width           =   6210
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   1260
            MaxLength       =   50
            TabIndex        =   11
            Tag             =   "Servidor SMTP|T|S|||sparam|smtpHost|||"
            Text            =   "3"
            Top             =   900
            Width           =   6210
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   8
            Left            =   1260
            MaxLength       =   50
            TabIndex        =   12
            Tag             =   "Usuario SMTP|T|S|||sparam|smtpUser|||"
            Text            =   "3"
            Top             =   1440
            Width           =   3090
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   9
            Left            =   5250
            MaxLength       =   15
            PasswordChar    =   "*"
            TabIndex        =   13
            Tag             =   "Password SMTP|T|S|||sparam|smtpPass|||"
            Text            =   "3"
            Top             =   1440
            Width           =   2220
         End
         Begin VB.Label Label1 
            Caption         =   "Lanza pantalla mail outlook"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   60
            Left            =   120
            TabIndex        =   158
            Top             =   2040
            Width           =   2040
         End
         Begin VB.Label Label1 
            Caption         =   "E-Mail"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   20
            Left            =   120
            TabIndex        =   72
            Top             =   480
            Width           =   1380
         End
         Begin VB.Label Label1 
            Caption         =   "Servidor SMTP"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   21
            Left            =   120
            TabIndex        =   71
            Top             =   960
            Width           =   1380
         End
         Begin VB.Label Label1 
            Caption         =   "Usuario"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   22
            Left            =   120
            TabIndex        =   70
            Top             =   1500
            Width           =   1380
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   23
            Left            =   4440
            TabIndex        =   69
            Top             =   1500
            Width           =   840
         End
         Begin VB.Label Label8 
            Caption         =   "Envio E-Mail"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   68
            Top             =   0
            Width           =   1320
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Soporte"
         ForeColor       =   &H00972E0B&
         Height          =   1035
         Left            =   -74610
         TabIndex        =   65
         Top             =   4080
         Width           =   8025
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   10
            Left            =   1350
            MaxLength       =   100
            TabIndex        =   16
            Tag             =   "Web Soporte|T|S|||sparam|websoporte|||"
            Top             =   360
            Width           =   6135
         End
         Begin VB.Label Label2 
            Caption         =   "Web soporte"
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
            Left            =   180
            TabIndex        =   66
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -74505
         TabIndex        =   56
         Top             =   600
         Width           =   7665
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   2235
            MaxLength       =   20
            TabIndex        =   0
            Tag             =   "Servidor Contabilidad|T|S|||sparam|serconta|||"
            Text            =   "3wwwwwwwwwwwwwwwwwwwwwwwwwwwww"
            Top             =   210
            Width           =   4875
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   4230
            MaxLength       =   15
            TabIndex        =   61
            Tag             =   "Código Parámetros Aplic|N|N|||sparam|codparam||S|"
            Text            =   "1"
            Top             =   240
            Width           =   645
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   3
            Left            =   2235
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   2
            Tag             =   "Password Contabilidad|T|S|||sparam|pasconta|||"
            Text            =   "3"
            Top             =   840
            Width           =   4875
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   2235
            MaxLength       =   20
            TabIndex        =   1
            Tag             =   "Usuario Contabilidad|T|S|||sparam|usuconta|||"
            Text            =   "3wwwwwwwwwwwwwwwwwww"
            Top             =   525
            Width           =   4875
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
            Height          =   285
            Index           =   4
            Left            =   2235
            MaxLength       =   2
            TabIndex        =   3
            Tag             =   "Nº Contabilidad|N|S|||sparam|numconta|||"
            Text            =   "3"
            Top             =   1185
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   15
            Left            =   510
            TabIndex        =   60
            Top             =   900
            Width           =   840
         End
         Begin VB.Label Label1 
            Caption         =   "Usuario"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   17
            Left            =   510
            TabIndex        =   59
            Top             =   600
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "Nº conta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   18
            Left            =   510
            TabIndex        =   58
            Top             =   1230
            Width           =   900
         End
         Begin VB.Label Label1 
            Caption         =   "Servidor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   19
            Left            =   510
            TabIndex        =   57
            Top             =   270
            Width           =   900
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Path FacturaE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   52
         Left            =   -74460
         TabIndex        =   162
         Top             =   4125
         Width           =   2070
      End
      Begin VB.Label Label34 
         Caption         =   "Número GGN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74460
         TabIndex        =   159
         Top             =   3780
         Width           =   1590
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   1
         Left            =   -71940
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   3780
         Width           =   240
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   0
         Left            =   -67380
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   900
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Movimiento Albaranes Salida"
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
         Index           =   47
         Left            =   -70800
         TabIndex        =   137
         Top             =   5070
         Width           =   2760
      End
      Begin VB.Label Label1 
         Caption         =   "Nro.Póliza Expediente"
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
         Index           =   46
         Left            =   -74625
         TabIndex        =   119
         Top             =   5085
         Width           =   1770
      End
      Begin VB.Label Label1 
         Caption         =   "Registro OPA"
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
         Index           =   45
         Left            =   -70770
         TabIndex        =   118
         Top             =   4230
         Width           =   1770
      End
      Begin VB.Label Label1 
         Caption         =   "Registro Coop.Sat"
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
         Index           =   44
         Left            =   -70770
         TabIndex        =   117
         Top             =   3840
         Width           =   1770
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   10
         Left            =   -73185
         ToolTipText     =   "Buscar Almacén"
         Top             =   4665
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Almacén"
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
         Index           =   42
         Left            =   -74625
         TabIndex        =   115
         Top             =   4665
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Reparto de Gastos Portes Albarán"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   41
         Left            =   -70755
         TabIndex        =   114
         Top             =   3450
         Width           =   2610
      End
      Begin VB.Label Label1 
         Caption         =   "%Desviación cálculo Gastos Reales"
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
         Index           =   40
         Left            =   -70755
         TabIndex        =   113
         Top             =   3090
         Width           =   2715
      End
      Begin VB.Label Label1 
         Caption         =   "Cta.Abono"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   -74100
         TabIndex        =   112
         Top             =   2955
         Width           =   1170
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   -72735
         ToolTipText     =   "Buscar Cta.Contable"
         Top             =   2955
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Nro. Fichero"
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
         Index           =   39
         Left            =   -74625
         TabIndex        =   109
         Top             =   4260
         Width           =   1770
      End
      Begin VB.Label Label1 
         Caption         =   "Cod.Chep Comunicador"
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
         Index           =   30
         Left            =   -74625
         TabIndex        =   108
         Top             =   3855
         Width           =   1770
      End
      Begin VB.Label Label1 
         Caption         =   "Path Ficheros generados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   33
         Left            =   -74820
         TabIndex        =   107
         Top             =   2160
         Width           =   1875
      End
      Begin VB.Label Label1 
         Caption         =   "Código Edi Vendedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   32
         Left            =   -74820
         TabIndex        =   106
         Top             =   1125
         Width           =   1650
      End
      Begin VB.Label Label1 
         Caption         =   "Registro Mercantil Emisor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   31
         Left            =   -74820
         TabIndex        =   105
         Top             =   1485
         Width           =   2325
      End
      Begin VB.Label Label1 
         Caption         =   "Último Nro Lote Chep"
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
         Index           =   29
         Left            =   -74625
         TabIndex        =   101
         Top             =   3450
         Width           =   1770
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   2160
         ToolTipText     =   "Buscar Carpeta"
         Top             =   1755
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Extensión"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   28
         Left            =   405
         TabIndex        =   99
         Top             =   1755
         Width           =   1605
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   2160
         ToolTipText     =   "Buscar Carpeta"
         Top             =   1350
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Carpeta Facturas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   405
         TabIndex        =   88
         Top             =   1350
         Width           =   1380
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   2160
         ToolTipText     =   "Buscar Carpeta"
         Top             =   945
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Carpeta Albaranes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   405
         TabIndex        =   86
         Top             =   945
         Width           =   1380
      End
      Begin VB.Label Label1 
         Caption         =   "Límite Peso Bruto CMR"
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
         Left            =   -74625
         TabIndex        =   77
         Top             =   3045
         Width           =   1770
      End
      Begin VB.Label Label1 
         Caption         =   "Precio Envase"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   -74595
         TabIndex        =   73
         Top             =   765
         Width           =   2205
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
      Height          =   285
      Index           =   29
      Left            =   2700
      MaxLength       =   10
      TabIndex        =   134
      Tag             =   "Carpeta Recibos Almacen|N|N|||sparam|codcarpetarecalm|000||"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   29
      Left            =   3960
      TabIndex        =   133
      Top             =   1800
      Width           =   4470
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   33
      Left            =   3960
      TabIndex        =   132
      Top             =   2160
      Width           =   4470
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
      Height          =   285
      Index           =   33
      Left            =   2700
      MaxLength       =   10
      TabIndex        =   131
      Tag             =   "Carpeta Recibos Campo|N|N|||sparam|codcarpetareccamp|000||"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Frame Frame10 
      Caption         =   "Recibos"
      ForeColor       =   &H00972E0B&
      Height          =   1050
      Left            =   540
      TabIndex        =   122
      Top             =   2385
      Width           =   7845
      Begin VB.ComboBox Combo1 
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
         Index           =   12
         Left            =   5805
         Style           =   2  'Dropdown List
         TabIndex        =   126
         Tag             =   "C4 Recibo|N|N|||sparam|c4recaridoc||N|"
         Top             =   585
         Width           =   1710
      End
      Begin VB.ComboBox Combo1 
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
         Index           =   11
         Left            =   3915
         Style           =   2  'Dropdown List
         TabIndex        =   125
         Tag             =   "C3 Recibo|N|N|||sparam|c3recaridoc||N|"
         Top             =   585
         Width           =   1710
      End
      Begin VB.ComboBox Combo1 
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
         Index           =   10
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   124
         Tag             =   "C2 Recibo|N|N|||sparam|c2recaridoc||N|"
         Top             =   585
         Width           =   1710
      End
      Begin VB.ComboBox Combo1 
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
         Index           =   9
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   123
         Tag             =   "C1 Recibo|N|N|||sparam|c1recaridoc||N|"
         Top             =   585
         Width           =   1710
      End
      Begin VB.Label Label1 
         Caption         =   "Campo 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   34
         Left            =   90
         TabIndex        =   130
         Top             =   315
         Width           =   1620
      End
      Begin VB.Label Label1 
         Caption         =   "Campo 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   35
         Left            =   1980
         TabIndex        =   129
         Top             =   315
         Width           =   1755
      End
      Begin VB.Label Label1 
         Caption         =   "Campo 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   36
         Left            =   3915
         TabIndex        =   128
         Top             =   315
         Width           =   1620
      End
      Begin VB.Label Label1 
         Caption         =   "Campo 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   37
         Left            =   5805
         TabIndex        =   127
         Top             =   315
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7845
      TabIndex        =   50
      Top             =   6495
      Width           =   1035
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
      Height          =   540
      Left            =   180
      TabIndex        =   53
      Top             =   6345
      Width           =   3000
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   210
         Width           =   2760
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6570
      TabIndex        =   49
      Top             =   6495
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7830
      TabIndex        =   51
      Top             =   6525
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   52
      Top             =   0
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Añadir"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3630
      Top             =   5250
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   9
      Left            =   2385
      ToolTipText     =   "Buscar Carpeta"
      Top             =   1845
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Carp.Recibos Almacen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   38
      Left            =   630
      TabIndex        =   136
      Top             =   1845
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Carpeta Recibos Campo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   43
      Left            =   630
      TabIndex        =   135
      Top             =   2205
      Width           =   1785
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   11
      Left            =   2385
      ToolTipText     =   "Buscar Carpeta"
      Top             =   2205
      Width           =   240
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnAñadir 
         Caption         =   "&Añadir"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   1
         Shortcut        =   ^M
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmConfParamAplic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ### [Monica] 06/09/2006
' procedimiento nuevo introducido de la gestion

Option Explicit

Private WithEvents frmCtas As frmCtasConta
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmIva As frmTipIVAConta
Attribute frmIva.VB_VarHelpID = -1
Private WithEvents frmDoc As frmCarpetaAridoc
Attribute frmDoc.VB_VarHelpID = -1
Private WithEvents frmExt As frmExtAridoc
Attribute frmExt.VB_VarHelpID = -1
Private WithEvents frmAri As frmCarpAridoc
Attribute frmAri.VB_VarHelpID = -1
Private WithEvents frmAlm As frmManAlmProp
Attribute frmAlm.VB_VarHelpID = -1
Private WithEvents frmCli As frmClientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmCCos As frmCCosConta 'centros de coste
Attribute frmCCos.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1

Private NombreTabla As String  'Nombre de la tabla o de la
Private Ordenacion As String
Private CadenaConsulta As String

Dim indice As Byte
Dim Encontrado As Boolean
Dim Modo As Byte
'0: Inicial
'2: Visualizacion
'3: Añadir
'4: Modificar


Private Sub chkInventar_GotFocus()
    PonerFocoChk chkInventar
End Sub

Private Sub chkInventar_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub chkInventar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkctrstock_GotFocus()
    PonerFocoChk chkctrstock
End Sub

Private Sub chkctrstock_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub checkpase_GotFocus()
    PonerFocoChk CheckPase
End Sub

Private Sub checkpase_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub checkcalculo_GotFocus()
    PonerFocoChk CheckCalculo
End Sub

Private Sub checkcalculo_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub checkreferencia_GotFocus()
    PonerFocoChk CheckReferencia
End Sub

Private Sub checkreferencia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkOutlook_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkOutlook_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim actualiza As Boolean
Dim kms As Currency

    
'    If Modo = 3 Then
'        If DatosOk Then
'            'Cambiamos el path
'            'CambiaPath True
'            If InsertarDesdeForm(Me) Then
'                PonerModo 0
''                ActualizaNombreEmpresa
'                MsgBox "Debe salir de la aplicacion para que los cambios tengan efecto", vbExclamation
'            End If
'
'        End If
'    End If


    If Modo = 4 Then 'MODIFICAR
        If DatosOk Then
            If Not vParamAplic Is Nothing Then
                'Datos contabilidad
                vParamAplic.ServidorConta = Text1(1).Text
                vParamAplic.UsuarioConta = Text1(2).Text
                vParamAplic.PasswordConta = Text1(3).Text
                vParamAplic.NumeroConta = ComprobarCero(Text1(4).Text)
                
                vParamAplic.WebSoporte = Text1(10).Text
                vParamAplic.DireMail = Text1(6).Text
                vParamAplic.Smtphost = Text1(7).Text
                vParamAplic.SmtpUser = Text1(8).Text
                vParamAplic.Smtppass = Text1(9).Text
                vParamAplic.TipoPrecio = Combo1(0).ListIndex
                vParamAplic.InventarioxProv = Me.chkInventar.Value
                vParamAplic.ControlStock = Me.chkctrstock.Value
                vParamAplic.PaseAlbarAgrupCalib = Me.CheckPase.Value
                vParamAplic.PaseRefLineaAlb = Me.CheckReferencia.Value
                vParamAplic.TipoCalculoComision = Me.CheckCalculo.Value
                
                
                vParamAplic.DiaPago1 = ComprobarCero(Text1(11).Text)
                vParamAplic.DiaPago2 = ComprobarCero(Text1(12).Text)
                vParamAplic.DiaPago3 = ComprobarCero(Text1(13).Text)
                vParamAplic.MesNoGirar = ComprobarCero(Text1(14).Text)
                
                vParamAplic.LimPesoCMR = Text1(15).Text
                vParamAplic.CodIvaTrans = Text1(5).Text
                
                vParamAplic.CodIvaNormal = Text1(38).Text
                vParamAplic.CodIvaExento = Text1(19).Text
                vParamAplic.CodIvaRecargo = Text1(20).Text
                
                
                vParamAplic.CtaTraReten = Text1(16).Text
                vParamAplic.CtaComReten = Text1(18).Text
                vParamAplic.CtaVentasFraACta = Text1(39).Text
                vParamAplic.CCosteFraACta = Text1(40).Text
                
                vParamAplic.CtaAboTrans = Text1(17).Text
                vParamAplic.NroLote = ComprobarCero(Text1(24).Text)
                vParamAplic.NroCheps = ComprobarCero(Text1(25).Text)
                vParamAplic.NroFiche = ComprobarCero(Text1(30).Text)
                
                vParamAplic.Almacen = ComprobarCero(Text1(32).Text)
                
                vParamAplic.NroPolizaExp = Text1(36).Text
                
                'aridoc
                vParamAplic.CarpetaAlb = Text1(21)
                vParamAplic.CarpetaFac = Text1(22)
                vParamAplic.CarpetaRecAlmacen = Text1(29)
                vParamAplic.CarpetaRecCampo = Text1(33)
                
                vParamAplic.Extension = Text1(23)
                
                vParamAplic.C1Albaran = Combo1(1).ListIndex
                vParamAplic.C2Albaran = Combo1(2).ListIndex
                vParamAplic.C3Albaran = Combo1(3).ListIndex
                vParamAplic.C4Albaran = Combo1(4).ListIndex
                vParamAplic.C1Factura = Combo1(5).ListIndex
                vParamAplic.C2Factura = Combo1(6).ListIndex
                vParamAplic.C3Factura = Combo1(7).ListIndex
                vParamAplic.C4Factura = Combo1(8).ListIndex
                vParamAplic.C1Recibo = Combo1(9).ListIndex
                vParamAplic.C2Recibo = Combo1(10).ListIndex
                vParamAplic.C3Recibo = Combo1(11).ListIndex
                vParamAplic.C4Recibo = Combo1(12).ListIndex
                vParamAplic.CodigoEdi = Text1(27).Text
                vParamAplic.RegMercantil = Text1(26).Text
'                vParamAplic.PathEdicom = Replace(Text1(28).Text, "\", "\\")
                vParamAplic.PathEdicom = Text1(28).Text
                vParamAplic.PathFacturaE = Text1(47).Text 'Replace(Text1(47).Text, "\", "\\")

                
                vParamAplic.PorcDesvCostes = ComprobarCero(Text1(31).Text)
                vParamAplic.PortesKiloCaja = Combo1(13).ListIndex
                
                vParamAplic.CodTipomAlb = Text1(37).Text
                
                ' solapa: horto
                vParamAplic.Text1CMR13 = Text1(41).Text
                vParamAplic.Text2CMR13 = Text1(42).Text
                
                vParamAplic.ClienteVtas = Text1(43).Text
                                
                ' envio de email por outlook
                vParamAplic.EnvioDesdeOutlook = Me.chkOutlook.Value
            
                ' Para utilizar el arigesmail
                vParamAplic.ExeEnvioMail = Trim(Text1(44).Text)
                
                '[Monica]25/04/2012: nuevo parametro
                vParamAplic.NumeroGGN = ComprobarCero(Text1(45).Text)
                
                '[Monica]25/05/2012: añadimos el path de fichadas para costes
                vParamAplic.PathFichadas = Text1(46).Text
                
                actualiza = vParamAplic.Modificar()
                TerminaBloquear
    
                If actualiza Then  'Inserta o Modifica
                    'Abrir la conexion a la conta q hemos modificado
                    CerrarConexionConta
                    If vParamAplic.NumeroConta <> 0 Then
                        If Not AbrirConexionConta() Then End
                        LeerNivelesEmpresa
                    End If
                    BloqueoMenusSegunContabilidad
                    PonerModo 2
                    If vParamAplic.leer = 0 Then
                        PonerFocoBtn Me.cmdSalir
                    End If
                End If
           End If
        End If
    End If
End Sub

Private Sub cmdCancelar_Click()
    TerminaBloquear
    If Data1.Recordset.EOF Then
        PonerModo 0
    Else
        PonerCampos
        PonerModo 2
    End If
End Sub

Private Sub cmdSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

' *** si n'hi han combos a la capçalera ***
Private Sub Combo1_GotFocus(Index As Integer)
    If Modo = 1 Then Combo1(Index).BackColor = vbYellow
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    If Combo1(Index).BackColor = vbYellow Then Combo1(Index).BackColor = vbWhite
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If Modo = 0 Then PonerCadenaBusqueda
    PonerFoco Text1(0)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
Dim I As Byte
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 3   'Anyadir
        .Buttons(2).Image = 4   'Modificar
        .Buttons(5).Image = 11  'Salir
    End With
    
 
    LimpiarCampos   'Limpia los campos TextBox
   
   'cargar IMAGES de busqueda
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
    'IMAGES para zoom
    For I = 2 To 3
        Me.imgZoom(I).Picture = frmPpal.imgListImages16.ListImages(5).Picture
    Next I

    For I = 0 To imgAyuda.Count - 1
        imgAyuda(I).Picture = frmPpal.ImageListB.ListImages(10).Picture
    Next I

    SSTab1.Tab = 0

    NombreTabla = "sparam"
    Ordenacion = " ORDER BY codparam"
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
    CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    Encontrado = True
    If Data1.Recordset.EOF Then
        'No hay registro de datos de parametros
        'quitar###
        Encontrado = False
    End If
    
    CargaCombo
        
    Me.SSTab1.TabEnabled(3) = (vParamAplic.HayAridoc = 1)
    Me.SSTab1.TabVisible(3) = (vParamAplic.HayAridoc = 1)
    If (vParamAplic.HayAridoc = 1) Then
        Me.SSTab1.TabsPerRow = 6
        AbrirConexionAridoc "root", "aritel"
    Else
        Me.SSTab1.TabsPerRow = 5
    End If
    
    Frame13.Enabled = (vParamAplic.HayCCostes = 1)
    Frame13.visible = (vParamAplic.HayCCostes = 1)
    
    PonerModo 0

End Sub


Private Sub PonerCadenaBusqueda()
On Error GoTo EEPonerBusq
    Screen.MousePointer = vbHourglass

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        LimpiarCampos
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
'        Me.Toolbar1.Buttons(1).Enabled = False 'Modificar
    Else
        Data1.Recordset.MoveFirst
        PonerCampos
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
    Text1(13).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(10)
    Text2(7).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CerrarConexionAridoc
End Sub

Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
'Codigo de almacen por defecto
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codalmac
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nomalmac
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codcliente
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre cliente
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'des macta
End Sub

Private Sub frmDoc_DatoSeleccionado(CadenaSeleccion As String)
'Carpetas de Aridoc
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nombre carpeta
End Sub

Private Sub frmExt_DatoSeleccionado(CadenaSeleccion As String)
'Extension de Aridoc
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmIva_DatoSeleccionado(CadenaSeleccion As String)
'Tipo de iva de la contabilidad
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigiva
    Text2(indice - 5).Text = RecuperaValor(CadenaSeleccion, 3) 'Porceiva
End Sub

Private Sub frmAri_DatoSeleccionado(CadenaSeleccion As String)
Dim cad As String
    cad = RecuperaValor(CadenaSeleccion, 1)
    Text1(indice).Text = Mid(cad, 2, Len(cad))
    Text1(indice).Text = Format(Text1(indice).Text, "000")
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 3)
End Sub

Private Sub frmCCos_DatoSeleccionado(CadenaSeleccion As String)
'Centro de Coste de la contabilidad
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub


Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           ' "____________________________________________________________"
            vCadena = "TIPO DE PRORRATEO DE DESCUENTOS EN FACTURAS DE CLIENTE " & vbCrLf & _
                      "============================================" & vbCrLf & vbCrLf & _
                      "Si está marcado el cálculo del prorrateo de la comisión se realiza " & vbCrLf & _
                      "sobre los Kilos Reales y el prorrateo de los Dtos sobre el Importe " & vbCrLf & _
                      "Bruto. " & vbCrLf & vbCrLf & _
                      "En caso contrario el prorrateo de la comisión se hace sobre el " & vbCrLf & _
                      "Importe Bruto al igual que los Dtos." & vbCrLf & vbCrLf
                      
        Case 1
            vCadena = "Número GGN de Globalgap que aparecerá en el pie de la Factura de" & vbCrLf & _
                      "Venta de clientes y en la Etiqueta de Palet. " & vbCrLf & vbCrLf
                      
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"
    
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim numNivel As Byte

    If vParamAplic.NumeroConta = 0 Then Exit Sub
    
    Select Case Index
        Case 0 'Porcentaje iva de factura de transporte
            If vParamAplic.NumeroConta = 0 Then Exit Sub
            indice = Index + 5
            Set frmIva = New frmTipIVAConta
            frmIva.DatosADevolverBusqueda = "0|1|2|"
            frmIva.CodigoActual = Text1(indice).Text
            frmIva.Show vbModal
            Set frmIva = Nothing
            PonerFoco Text1(indice)
        
        Case 4, 5 'Porcentaje iva de factura de ventas
            If vParamAplic.NumeroConta = 0 Then Exit Sub
            indice = Index + 15
            Set frmIva = New frmTipIVAConta
            frmIva.DatosADevolverBusqueda = "0|1|2|"
            frmIva.CodigoActual = Text1(indice).Text
            frmIva.Show vbModal
            Set frmIva = Nothing
            PonerFoco Text1(indice)
    
        Case 12 'Porcentaje iva de factura de ventas (normal)
            If vParamAplic.NumeroConta = 0 Then Exit Sub
            indice = Index + 26
            Set frmIva = New frmTipIVAConta
            frmIva.DatosADevolverBusqueda = "0|1|2|"
            frmIva.CodigoActual = Text1(indice).Text
            frmIva.Show vbModal
            Set frmIva = Nothing
            PonerFoco Text1(indice)
    
        Case 1, 2, 3  'Cuentas Contables (de contabilidad)
            If vParamAplic.NumeroConta = 0 Then Exit Sub
            indice = Index + 15
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(indice).Text
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(indice)
                
         Case 13 ' cta contacle de ventas de facturas a cuenta
            If vParamAplic.NumeroConta = 0 Then Exit Sub
            indice = 39
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(indice).Text
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(indice)
                
         Case 14 ' centro de coste para facturas a cuenta
            If vParamAplic.NumeroConta = 0 Then Exit Sub
            indice = 40
            Set frmCCos = New frmCCosConta
            frmCCos.DatosADevolverBusqueda = "0|1|"
            frmCCos.CodigoActual = Text1(indice).Text
            frmCCos.Show vbModal
            Set frmCCos = Nothing
            PonerFoco Text1(indice)
         
         Case 6, 7, 9, 11 'carpetas de aridoc
            Select Case Index
                Case 6, 7
                    indice = Index + 15
                Case 9
                    indice = Index + 20
                Case 11
                    indice = Index + 22
            End Select
            
            Set frmAri = New frmCarpAridoc
            frmAri.Opcion = 20
            frmAri.Show vbModal
            Set frmAri = Nothing
            PonerFoco Text1(indice)
        
'            Set frmDoc = New frmCarpetaAridoc
'            frmDoc.DatosADevolverBusqueda = "0|1|"
'            frmDoc.CodigoActual = Text1(indice).Text
'            frmDoc.Show vbModal
'            Set frmDoc = Nothing
'            PonerFoco Text1(indice)
         
         Case 8 'extesion de fichero de aridoc
            indice = Index + 15
            Set frmExt = New frmExtAridoc
            frmExt.DatosADevolverBusqueda = "0|1|"
            frmExt.CodigoActual = Text1(indice).Text
            frmExt.Show vbModal
            Set frmExt = Nothing
            PonerFoco Text1(indice)
            
         Case 10 ' codigo de almacen
            indice = Index + 22
            Set frmAlm = New frmManAlmProp
            frmAlm.DatosADevolverBusqueda = "0|1|"
            frmAlm.CodigoActual = Text1(indice).Text
            frmAlm.Show vbModal
            Set frmAlm = Nothing
            PonerFoco Text1(indice)
         
         Case 15 ' codigo de cliente de ventas
            indice = 43
            PonerFoco Text1(indice)
            Set frmCli = New frmClientes
            frmCli.DatosADevolverBusqueda = "0|1|"
            frmCli.Show vbModal
            Set frmCli = Nothing
            PonerFoco Text1(indice)
         
         
         
            
'        Case 5, 6 'raices de las cuentas contables de socio y cliente
'            If vParamAplic.NumeroConta = 0 Then Exit Sub
'            Indice = Index + 5
'            Set frmCtas = New frmCtasConta
'            numNivel = DevuelveDesdeBDNew(cConta, "empresa", "numnivel", "codempre", Text1(4).Text, "N")
'            frmCtas.NumDigit = DevuelveDesdeBDNew(cConta, "empresa", "numdigi" & numNivel - 1, "codempre", Text1(4).Text, "N")
'            frmCtas.DatosADevolverBusqueda = "0|1|"
'            frmCtas.CodigoActual = Text1(Indice).Text
'            frmCtas.Show vbModal
'            Set frmCtas = Nothing
'            PonerFoco Text1(Indice)
        
    End Select
End Sub
Private Sub frmZ_Actualizar(vCampo As String)
     Text1(indice).Text = vCampo
End Sub


Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    Select Case Index
            
        Case 2
            indice = 41
            frmZ.pTitulo = "Texto 1 para CMR Seccion 13"
            frmZ.pValor = Text1(indice).Text
            frmZ.pModo = Modo
        
            frmZ.Show vbModal
            Set frmZ = Nothing
        
        
        Case 3
            indice = 42
            frmZ.pTitulo = "Texto 2 para CMR Seccion 13"
            frmZ.pValor = Text1(indice).Text
            frmZ.pModo = Modo
        
            frmZ.Show vbModal
            Set frmZ = Nothing
    End Select
    
End Sub

'Private Sub mnAñadir_Click()
'    If BLOQUEADesdeFormulario(Me) Then BotonAnyadir
'End Sub

Private Sub mnModificar_Click()
    If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
'14/02/2007 antes estaba esto
'    KEYpress (KeyAscii)
' ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 5: KEYBusqueda KeyAscii, 0 'tipo de iva transporte
            Case 16: KEYBusqueda KeyAscii, 1 'cuenta de retencion de transportistas
            Case 17: KEYBusqueda KeyAscii, 2 'cuenta de diferencias positivas
            Case 18: KEYBusqueda KeyAscii, 3 'cuenta de retencion de comisionistas
            Case 39: KEYBusqueda KeyAscii, 13 'cuenta ventas de facturas a cuenta
            Case 40: KEYBusqueda KeyAscii, 14 'centro de coste de facturas a cuenta
            Case 43: KEYBusqueda KeyAscii, 15 'codigo de cliente de ventas
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim cad As String

    If Text1(Index).Text = "" Then Exit Sub

    'Quitar espacios en blanco
    Text1(Index).Text = Trim(Text1(Index).Text)
    
    Select Case Index
        Case 4 'numero de contabilidad
            If Not EsNumerico(Text1(Index).Text) Then
                Text1(Index).Text = ""
                PonerFoco Text1(Index)
            Else
                cmdAceptar_Click
            End If
            
        Case 5, 19, 20   ' tipo de iva de contabilidad
            'conConta: BD Contabilidad
            If vParamAplic.NumeroConta <> 0 Then
                If PonerFormatoEntero(Text1(Index)) Then
                    Text2(Index - 5).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(Index), "N")
                Else
                    Text2(Index - 5).Text = ""
                End If
            End If
            
        Case 38 ' tipo de iva de contabilidad (iva normal para facturas a cuenta)
            If vParamAplic.NumeroConta <> 0 Then
                If PonerFormatoEntero(Text1(Index)) Then
                    Text2(Index).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(Index), "N")
                Else
                    Text2(Index).Text = ""
                End If
            End If
            
        Case 40 ' centro de coste de facturas a cuenta
            If Text1(Index).Text <> "" Then
                If vParamAplic.NumeroConta <> 0 Then
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), "cabccost", "nomccost", "codccost", "T", cConta)
                End If
            End If
             
        Case 15, 24, 25, 30, 32
            PonerFormatoEntero Text1(Index)
            
        Case 16, 17, 18, 39
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo)
            
        Case 21, 22, 29, 33
            If Text1(Index).Text = "" Then Exit Sub
'            Text2(Index).Text = PonerNombreDeCod(Text1(Index), "carpetas", "nombre", "codcarpeta", "N", cAridoc)
            Text1(Index).Text = Format(Text1(Index).Text, "000")
            cad = CargaPath(Text1(Index))
            Text2(Index).Text = Mid(cad, 2, Len(cad))
        
        Case 23
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), "extension", "descripcion", "codext", "N", cAridoc)
        
        Case 31 ' porcentaje de desviacion para el calculo de gastos costes
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 4
            
        Case 43
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), "clientes", "nomclien", "codclien", "N")
        
        
    End Select
End Sub


'Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
'    Select Case Index
'        Case 6, 7
'            If Text1(Index).Text <> "" Then
'                If Not EsNumerico(Text1(Index).Text) Then
'                    Cancel = True
'                    ConseguirFoco Text1(Index), Modo
'                End If
'            End If
'    End Select
'
'End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
'        Case 1  'Anyadir
'            BotonAnyadir
        Case 2  'Modificar
            mnModificar_Click
        Case 5 'Salir
            mnSalir_Click
    End Select
End Sub


'Private Sub BotonAnyadir()
'    LimpiarCampos
'    PonerModo 3
'    Text1(0).Text = 1
'    PonerFoco Text1(1)
'End Sub


Private Sub BotonModificar()
    PonerModo 4
    PonerFoco Text1(1)
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
    DatosOk = False
    b = CompForm(Me)
    DatosOk = b
End Function


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerBotonCabecera(b As Boolean)
    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdSalir.visible = b
'    If b Then Me.lblIndicador.Caption = ""
End Sub


Private Sub PonerCampos()
Dim I As Byte
Dim cad As String


On Error GoTo EPonerCampos

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    ' ************* configurar els camps de les descripcions de les comptes *************
    If vParamAplic.NumeroConta <> 0 Then
        Text2(0).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(5), "N")
        Text2(14).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(19), "N")
        Text2(15).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(20), "N")
        Text2(38).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(38), "N")
    End If
    
    If vParamAplic.NumeroConta <> 0 Then
        Text2(16).Text = PonerNombreCuenta(Text1(16), Modo)
        Text2(18).Text = PonerNombreCuenta(Text1(18), Modo)
        Text2(39).Text = PonerNombreCuenta(Text1(39), Modo)
        Text2(40).Text = DevuelveDesdeBDNew(cConta, "cabccost", "nomccost", "codccost", Text1(40), "T")
    End If
    
    If vParamAplic.HayAridoc = 1 Then
         If ComprobarCero(Text1(21).Text) <> 0 Then
            cad = CargaPath(Text1(21))
            Text2(21).Text = Mid(cad, 2, Len(cad))
         End If
         If ComprobarCero(Text1(22).Text) <> 0 Then
            cad = CargaPath(Text1(22))
            Text2(22).Text = Mid(cad, 2, Len(cad))
         End If
'         cad = CargaPath(Text1(29))
'         Text2(29).Text = Mid(cad, 2, Len(cad))
'         cad = CargaPath(Text1(33))
'         Text2(33).Text = Mid(cad, 2, Len(cad))

         Text2(23).Text = DevuelveDesdeBDNew(cAridoc, "extension", "descripcion", "codext", Text1(23).Text, "N")
    End If
    Text2(32).Text = PonerNombreDeCod(Text1(32), "salmpr", "nomalmac", "codalmac", "N")
    Text2(43).Text = PonerNombreDeCod(Text1(43), "clientes", "nomclien", "codclien", "N")
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub LimpiarCampos()
Dim I As Integer

    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    For I = 0 To Combo1.Count - 1
        Combo1(I).ListIndex = -1
    Next I
    '### a mano
End Sub


'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim b As Boolean
Dim I As Byte

    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
      
    '------------------------------------------------------
    'Modo insertar o modificar
    b = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    If cmdCancelar.visible Then
        cmdCancelar.Cancel = True
    Else
        cmdCancelar.Cancel = False
    End If
    PonerBotonCabecera Not b
       
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1 y bloquea la clave primaria
    BloquearText1 Me, Modo
    BloquearCombo Me, Modo
    
    
    'Bloquear imagen de Busqueda
    For I = 0 To imgBuscar.Count - 1
        Me.imgBuscar(I).Enabled = (Modo >= 3)
    Next I
    
    BloquearImgBuscar Me, Modo
    'Bloquear los checkbox
    BloquearChecks Me, Modo
    
    PonerModoOpcionesMenu 'Activar opciones de menu según el Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean
    b = (Modo = 3) Or (Modo = 4)
    Me.Toolbar1.Buttons(1).Enabled = Not Encontrado And Not b  'Añadir
    Me.Toolbar1.Buttons(2).Enabled = Encontrado And Not b 'Modificar
    Me.mnAñadir.Enabled = Not Encontrado And Not b
    Me.mnModificar.Enabled = Encontrado And Not b
'    Me.Toolbar1.Buttons(2).Enabled = (Not b) 'Modificar
End Sub


' ********* si n'hi han combos a la capçalera ************
Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim I As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For I = 0 To Combo1.Count - 1
        Combo1(I).Clear
    Next I
    
    Combo1(0).AddItem "Precio Medio Ponderado"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Ultimo Precio de Compra"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    
    
    'combos de albaranes
    For I = 1 To 4
        Combo1(I).AddItem "Nro.Albaran"
        Combo1(I).ItemData(Combo1(I).NewIndex) = 0
        Combo1(I).AddItem "Cod.Cliente"
        Combo1(I).ItemData(Combo1(I).NewIndex) = 1
        Combo1(I).AddItem "Nom.Cliente"
        Combo1(I).ItemData(Combo1(I).NewIndex) = 2
        Combo1(I).AddItem "Destino"
        Combo1(I).ItemData(Combo1(I).NewIndex) = 3
        Combo1(I).AddItem "Procedencia"
        Combo1(I).ItemData(Combo1(I).NewIndex) = 4
    Next I
    
    'combos de facturas
    For I = 5 To 8
        Combo1(I).AddItem "Nro.Factura"
        Combo1(I).ItemData(Combo1(I).NewIndex) = 0
        Combo1(I).AddItem "Cod.Cliente"
        Combo1(I).ItemData(Combo1(I).NewIndex) = 1
        Combo1(I).AddItem "Nom.Cliente"
        Combo1(I).ItemData(Combo1(I).NewIndex) = 2
        Combo1(I).AddItem "Procedencia"
        Combo1(I).ItemData(Combo1(I).NewIndex) = 3
    Next I
    
    'combos de recibos
    For I = 9 To 12
        Combo1(I).AddItem "Cod.Trabajador"
        Combo1(I).ItemData(Combo1(I).NewIndex) = 0
        Combo1(I).AddItem "Nom.Trabajador"
        Combo1(I).ItemData(Combo1(I).NewIndex) = 1
        Combo1(I).AddItem "Procedencia"
        Combo1(I).ItemData(Combo1(I).NewIndex) = 2
    Next I
    
    'combo de reparto de gastos portes
    Combo1(13).AddItem "Kilos"
    Combo1(13).ItemData(Combo1(13).NewIndex) = 0
    Combo1(13).AddItem "Cajas"
    Combo1(13).ItemData(Combo1(13).NewIndex) = 1
    
End Sub

Private Function CargaPath(Codigo As Integer) As String
Dim Nod As Node
Dim j As Integer
Dim I As Integer
Dim C As String
Dim campo1 As String
Dim padre As String
Dim A As String

    'Primero copiamos la carpeta
    C = "\" & DevuelveDesdeBDNew(cAridoc, "carpetas", "nombre", "codcarpeta", CInt(Codigo), "N")
    campo1 = "nombre"
    padre = DevuelveDesdeBDNew(cAridoc, "carpetas", "padre", "codcarpeta", CStr(Codigo), "N", campo1)
    If CInt(ComprobarCero(padre)) > 0 Then
        C = CargaPath(CInt(padre)) & C
    End If
'
'    If No.Children > 0 Then
'        J = No.Children
'        Set Nod = No.Child
'        For i = 1 To J
'           C = C & CopiaArchivosCarpetaRecursiva(Nod)
'           If i <> J Then Set Nod = Nod.Next
'        Next i
'    End If
    CargaPath = C
End Function

