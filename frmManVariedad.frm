VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmManVariedad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Variedades"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12075
   Icon            =   "frmManVariedad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   12075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   765
      Index           =   0
      Left            =   240
      TabIndex        =   27
      Top             =   480
      Width           =   11655
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1530
         MaxLength       =   6
         TabIndex        =   0
         Tag             =   "C�digo de variedad|N|N|0|999999|variedades|codvarie|000000|S|"
         Top             =   225
         Width           =   855
      End
      Begin VB.TextBox text1 
         Height          =   285
         Index           =   1
         Left            =   3480
         MaxLength       =   20
         TabIndex        =   1
         Tag             =   "Nombre|T|N|||variedades|nomvarie|||"
         Top             =   240
         Width           =   3195
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre "
         Height          =   255
         Left            =   2640
         TabIndex        =   29
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "C�digo Variedad"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   1230
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   285
      TabIndex        =   24
      Top             =   5985
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
         TabIndex        =   25
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10815
      TabIndex        =   18
      Top             =   6075
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9525
      TabIndex        =   17
      Top             =   6060
      Width           =   1035
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4515
      Left            =   240
      TabIndex        =   26
      Top             =   1320
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   7964
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   6
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos b�sicos"
      TabPicture(0)   =   "frmManVariedad.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(26)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label18"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "imgBuscar(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "imgBuscar(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label20"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "imgBuscar(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "imgBuscar(4)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label26"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "imgBuscar(9)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label27"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label28"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "imgBuscar(10)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "imgAyuda(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(1)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(2)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "text2(2)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "text1(2)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "text1(4)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "text1(3)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "text2(3)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "text1(9)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "text2(9)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Frame3"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "text2(26)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "text1(26)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "text1(27)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "text2(27)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Frame4"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Frame6"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Combo1(1)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Combo1(2)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).ControlCount=   32
      TabCaption(1)   =   "Calibres"
      TabPicture(1)   =   "frmManVariedad.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameAux0"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Calidades"
      TabPicture(2)   =   "frmManVariedad.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameAux1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Recolecci�n"
      TabPicture(3)   =   "frmManVariedad.frx":0060
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label7"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label6"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label19"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label8"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label2"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Label3"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Label9"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Label10"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Label11"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Label12"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Label13"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "Label14"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "Label15"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "Label16"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "Label17"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "Label1(19)"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "Label31"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "text1(5)"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "text1(6)"
      Tab(3).Control(18).Enabled=   0   'False
      Tab(3).Control(19)=   "text1(7)"
      Tab(3).Control(19).Enabled=   0   'False
      Tab(3).Control(20)=   "text1(10)"
      Tab(3).Control(20).Enabled=   0   'False
      Tab(3).Control(21)=   "text1(11)"
      Tab(3).Control(21).Enabled=   0   'False
      Tab(3).Control(22)=   "text1(8)"
      Tab(3).Control(22).Enabled=   0   'False
      Tab(3).Control(23)=   "text1(12)"
      Tab(3).Control(23).Enabled=   0   'False
      Tab(3).Control(24)=   "text1(13)"
      Tab(3).Control(24).Enabled=   0   'False
      Tab(3).Control(25)=   "text1(14)"
      Tab(3).Control(25).Enabled=   0   'False
      Tab(3).Control(26)=   "text1(15)"
      Tab(3).Control(26).Enabled=   0   'False
      Tab(3).Control(27)=   "text1(16)"
      Tab(3).Control(27).Enabled=   0   'False
      Tab(3).Control(28)=   "text1(17)"
      Tab(3).Control(28).Enabled=   0   'False
      Tab(3).Control(29)=   "text1(18)"
      Tab(3).Control(29).Enabled=   0   'False
      Tab(3).Control(30)=   "text1(19)"
      Tab(3).Control(30).Enabled=   0   'False
      Tab(3).Control(31)=   "text1(20)"
      Tab(3).Control(31).Enabled=   0   'False
      Tab(3).Control(32)=   "Combo1(0)"
      Tab(3).Control(32).Enabled=   0   'False
      Tab(3).Control(33)=   "text1(30)"
      Tab(3).Control(33).Enabled=   0   'False
      Tab(3).Control(34)=   "Frame5"
      Tab(3).Control(34).Enabled=   0   'False
      Tab(3).ControlCount=   35
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   2
         Left            =   -67440
         Style           =   2  'Dropdown List
         TabIndex        =   134
         Tag             =   "Tipo Variedad|N|N|||variedades|tipovarie2||N|"
         Top             =   900
         Width           =   1440
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   -67440
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Tag             =   "Tipo Mercancia|N|N|||variedades|tipovariedad||N|"
         Top             =   480
         Width           =   1440
      End
      Begin VB.Frame Frame6 
         Caption         =   "Cuenta Contable Comisionistas"
         Height          =   870
         Left            =   -69060
         TabIndex        =   124
         Top             =   3510
         Width           =   5595
         Begin VB.TextBox text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   34
            Left            =   2205
            TabIndex        =   125
            Top             =   360
            Width           =   3210
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   34
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   16
            Tag             =   "Cta Comisionista|T|S|||variedades|ctacomisionista|||"
            Top             =   360
            Width           =   1050
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   16
            Left            =   810
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label35 
            Caption         =   "Cta."
            Height          =   255
            Left            =   180
            TabIndex        =   126
            Top             =   390
            Width           =   510
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Cuentas Contables"
         Height          =   2385
         Left            =   4140
         TabIndex        =   117
         Top             =   1890
         Width           =   7185
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   37
            Left            =   2295
            MaxLength       =   35
            TabIndex        =   114
            Tag             =   "Cta.Acarreo Recolecci�n|T|S|||variedades|ctaacarecol|||"
            Top             =   1860
            Width           =   1410
         End
         Begin VB.TextBox text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   37
            Left            =   3810
            TabIndex        =   131
            Top             =   1860
            Width           =   3210
         End
         Begin VB.TextBox text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   36
            Left            =   3810
            TabIndex        =   129
            Top             =   1530
            Width           =   3210
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   36
            Left            =   2295
            MaxLength       =   35
            TabIndex        =   113
            Tag             =   "Cta.Facturas Transporte|T|S|||variedades|ctatransporte|||"
            Top             =   1530
            Width           =   1410
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   35
            Left            =   2295
            MaxLength       =   35
            TabIndex        =   112
            Tag             =   "Cta.Compras Terceros|T|S|||variedades|ctasiniestros|||"
            Top             =   1215
            Width           =   1410
         End
         Begin VB.TextBox text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   35
            Left            =   3810
            TabIndex        =   127
            Top             =   1215
            Width           =   3210
         End
         Begin VB.TextBox text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   33
            Left            =   3810
            TabIndex        =   123
            Top             =   900
            Width           =   3210
         End
         Begin VB.TextBox text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   32
            Left            =   3810
            TabIndex        =   122
            Top             =   585
            Width           =   3210
         End
         Begin VB.TextBox text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   31
            Left            =   3810
            TabIndex        =   121
            Top             =   240
            Width           =   3180
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   33
            Left            =   2295
            MaxLength       =   35
            TabIndex        =   111
            Tag             =   "Cta.Compras Terceros|T|S|||variedades|ctacomtercero|||"
            Top             =   900
            Width           =   1410
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   32
            Left            =   2295
            MaxLength       =   35
            TabIndex        =   110
            Tag             =   "Cta Liquidaci�n|T|S|||variedades|ctaliquidacion|||"
            Top             =   570
            Width           =   1410
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   31
            Left            =   2295
            MaxLength       =   35
            TabIndex        =   109
            Tag             =   "Cuenta Anticipos|T|S|||variedades|ctaanticipo|||"
            Top             =   240
            Width           =   1410
         End
         Begin VB.Label Label38 
            Caption         =   "Acarreo Recolecci�n"
            Height          =   255
            Left            =   330
            TabIndex        =   132
            Top             =   1890
            Width           =   1500
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   19
            Left            =   1920
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   1890
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   18
            Left            =   1920
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   1560
            Width           =   240
         End
         Begin VB.Label Label37 
            Caption         =   "Facturas Transporte"
            Height          =   255
            Left            =   330
            TabIndex        =   130
            Top             =   1560
            Width           =   1470
         End
         Begin VB.Label Label36 
            Caption         =   "Siniestros"
            Height          =   255
            Left            =   330
            TabIndex        =   128
            Top             =   1245
            Width           =   1470
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   17
            Left            =   1920
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   1245
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   15
            Left            =   1920
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   930
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   14
            Left            =   1920
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   600
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   13
            Left            =   1920
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   270
            Width           =   240
         End
         Begin VB.Label Label34 
            Caption         =   "Compras Terceros"
            Height          =   255
            Left            =   330
            TabIndex        =   120
            Top             =   930
            Width           =   1470
         End
         Begin VB.Label Label33 
            Caption         =   "Liquidaci�n"
            Height          =   255
            Left            =   330
            TabIndex        =   119
            Top             =   600
            Width           =   1440
         End
         Begin VB.Label Label32 
            Caption         =   "Anticipos"
            Height          =   255
            Left            =   330
            TabIndex        =   118
            Top             =   270
            Width           =   1440
         End
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   30
         Left            =   1860
         MaxLength       =   35
         TabIndex        =   100
         Tag             =   "Rdto Maximo|N|S|||variedades|rdtomaximo|###,###,##0||"
         Top             =   3960
         Width           =   1365
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1875
         TabIndex        =   90
         Tag             =   "Clasificaci�n|N|N|0|1|variedades|tipoclasifica|0|N|"
         Text            =   "Combo1"
         Top             =   855
         Width           =   1395
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   20
         Left            =   9930
         MaxLength       =   35
         TabIndex        =   108
         Tag             =   "Euros/kg hanegada|N|S|||variedades|eurhaneg|0.0000||"
         Top             =   825
         Width           =   1410
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   19
         Left            =   9930
         MaxLength       =   35
         TabIndex        =   107
         Tag             =   "Euros/kg tria|N|S|||variedades|eurotria|0.0000||"
         Top             =   510
         Width           =   1410
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   6420
         MaxLength       =   35
         TabIndex        =   106
         Tag             =   "Euros/kg Seg.Social|N|S|||variedades|eursegsoc|0.0000||"
         Top             =   1455
         Width           =   1410
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   6420
         MaxLength       =   35
         TabIndex        =   105
         Tag             =   "Euros/kg mano obra|N|S|||variedades|eurmanob|0.0000||"
         Top             =   1140
         Width           =   1410
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   6420
         MaxLength       =   35
         TabIndex        =   104
         Tag             =   "Euros/kg recolecion|N|S|||variedades|eurecole|0.0000||"
         Top             =   825
         Width           =   1410
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   15
         Left            =   6420
         MaxLength       =   35
         TabIndex        =   103
         Tag             =   "Euros/kg destajo|N|S|||variedades|eurdesta|0.0000||"
         Top             =   510
         Width           =   1410
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   14
         Left            =   1845
         MaxLength       =   35
         TabIndex        =   99
         Tag             =   "Porc.Destrio|N|S|||variedades|porcdest|##0.00||"
         Top             =   3495
         Width           =   1410
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   1845
         MaxLength       =   35
         TabIndex        =   98
         Tag             =   "Porc.Mermas|N|S|||variedades|porcmerm|##0.00||"
         Top             =   3180
         Width           =   1410
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   1845
         MaxLength       =   35
         TabIndex        =   97
         Tag             =   "Porc.Industria|N|S|||variedades|porcindu|##0.00||"
         Top             =   2865
         Width           =   1410
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   1845
         MaxLength       =   35
         TabIndex        =   96
         Tag             =   "Arroba/Jornal|N|S|0|999.99|variedades|arrobjor|##0.00||"
         Top             =   2385
         Width           =   1410
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   9915
         MaxLength       =   35
         TabIndex        =   102
         Tag             =   "Factor Cor.Destrio|N|S|0|999.99|variedades|facorrme|##0.00||"
         Top             =   1455
         Width           =   1410
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   9915
         MaxLength       =   35
         TabIndex        =   101
         Tag             =   "Factor Cor.Destrio|N|S|0|999.99|variedades|facorrde|##0.00||"
         Top             =   1140
         Width           =   1410
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1845
         MaxLength       =   35
         TabIndex        =   95
         Tag             =   "Max Kilos Cajon|N|S|0|999.99|variedades|maxkgcaj|##0.00||"
         Top             =   1980
         Width           =   1410
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1845
         MaxLength       =   35
         TabIndex        =   93
         Tag             =   "Min Kilos Cajon|N|S|0|999.99|variedades|minkgcaj|##0.00||"
         Top             =   1620
         Width           =   1410
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1845
         MaxLength       =   35
         TabIndex        =   92
         Tag             =   "Kilos Cajon|N|S|0|999.99|variedades|kgscajon|##0.00||"
         Top             =   1260
         Width           =   1410
      End
      Begin VB.Frame FrameAux1 
         BorderStyle     =   0  'None
         Height          =   4020
         Left            =   -74865
         TabIndex        =   63
         Top             =   360
         Width           =   11355
         Begin VB.CheckBox chkAux 
            BackColor       =   &H80000005&
            Height          =   255
            Index           =   0
            Left            =   6480
            TabIndex        =   116
            Tag             =   "Hay gastos|N|N|0|1|rcalidad|gastosrec|||"
            Top             =   3690
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtAux1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   9
            Left            =   5880
            MaxLength       =   30
            TabIndex        =   115
            Tag             =   "Nom.Calibrador 2|T|S|||rcalidad|nomcalibrador2|||"
            Text            =   "Cal 2"
            Top             =   3690
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox txtAux1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   8
            Left            =   5265
            MaxLength       =   30
            TabIndex        =   72
            Tag             =   "Nom.Calibrador 1|T|S|||rcalidad|nomcalibrador1|||"
            Text            =   "Cal 1"
            Top             =   3690
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox txtAux1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   7
            Left            =   4635
            MaxLength       =   3
            TabIndex        =   71
            Text            =   "des"
            Top             =   3690
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox txtAux1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   6
            Left            =   4005
            MaxLength       =   3
            TabIndex        =   70
            Tag             =   "Tipo Calidad 1|N|N|||rcalidad|tipcalid1|||"
            Text            =   "tip"
            Top             =   3690
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox txtAux1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   5
            Left            =   3375
            MaxLength       =   3
            TabIndex        =   69
            Text            =   "des"
            Top             =   3690
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox txtAux1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   4
            Left            =   2745
            MaxLength       =   3
            TabIndex        =   68
            Tag             =   "Tipo Calidad|N|N|||rcalidad|tipcalid|||"
            Text            =   "tip"
            Top             =   3690
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox txtAux1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   2
            Left            =   930
            MaxLength       =   12
            TabIndex        =   67
            Tag             =   "Nombre Calidad|T|N|||rcalidad|nomcalid|||"
            Text            =   "nomcali"
            Top             =   3690
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.TextBox txtAux1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   1
            Left            =   570
            MaxLength       =   2
            TabIndex        =   66
            Tag             =   "Codigo Calidad|N|N|1|99|rcalidad|codcalid|00|S|"
            Text            =   "ca"
            Top             =   3690
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox txtAux1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   0
            Left            =   90
            MaxLength       =   6
            TabIndex        =   65
            Tag             =   "C�digo Variedad|N|N|1|999999|rcalidad|codvarie|000000|S|"
            Text            =   "codvar"
            Top             =   3690
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtAux1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   3
            Left            =   2025
            MaxLength       =   3
            TabIndex        =   64
            Tag             =   "Nombre Calidad Abr|T|N|||rcalidad|nomcalab|||"
            Text            =   "cab"
            Top             =   3690
            Visible         =   0   'False
            Width           =   585
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   1
            Left            =   0
            TabIndex        =   73
            Top             =   135
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
         Begin MSAdodcLib.Adodc AdoAux 
            Height          =   375
            Index           =   1
            Left            =   5580
            Top             =   90
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
         Begin MSDataGridLib.DataGrid DataGridAux 
            Bindings        =   "frmManVariedad.frx":007C
            Height          =   3465
            Index           =   1
            Left            =   0
            TabIndex        =   74
            Top             =   135
            Width           =   11310
            _ExtentX        =   19950
            _ExtentY        =   6112
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
      Begin VB.Frame Frame4 
         Caption         =   "Cuentas Contables Transporte"
         Height          =   1095
         Left            =   -69060
         TabIndex        =   58
         Top             =   2295
         Width           =   5595
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   29
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   15
            Tag             =   "Cta Transp.Export.|T|S|||variedades|ctatraexporta|||"
            Top             =   720
            Width           =   1050
         End
         Begin VB.TextBox text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   29
            Left            =   2205
            TabIndex        =   60
            Top             =   720
            Width           =   3210
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   28
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   14
            Tag             =   "Cta Transp.Int.|T|S|||variedades|ctatrainterior|||"
            Top             =   360
            Width           =   1050
         End
         Begin VB.TextBox text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   28
            Left            =   2205
            TabIndex        =   59
            Top             =   375
            Width           =   3210
         End
         Begin VB.Label Label30 
            Caption         =   "Interior"
            Height          =   255
            Left            =   180
            TabIndex        =   62
            Top             =   390
            Width           =   510
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   12
            Left            =   810
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   720
            Width           =   240
         End
         Begin VB.Label Label29 
            Caption         =   "Export."
            Height          =   255
            Left            =   180
            TabIndex        =   61
            Top             =   705
            Width           =   555
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   11
            Left            =   810
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   360
            Width           =   240
         End
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   27
         Left            =   -66855
         TabIndex        =   56
         Top             =   1890
         Width           =   3345
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   27
         Left            =   -67620
         MaxLength       =   4
         TabIndex        =   13
         Tag             =   "Centro Coste|T|S|||variedades|codccost|||"
         Top             =   1890
         Width           =   690
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   26
         Left            =   -73380
         MaxLength       =   2
         TabIndex        =   5
         Tag             =   "Codigo IVA|N|N|0|99|variedades|codigiva|00||"
         Top             =   1530
         Width           =   690
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   26
         Left            =   -72660
         TabIndex        =   54
         Top             =   1545
         Width           =   3435
      End
      Begin VB.Frame Frame3 
         Caption         =   "Cuentas Contables Ventas"
         Height          =   2085
         Left            =   -74820
         TabIndex        =   42
         Top             =   2295
         Width           =   5685
         Begin VB.TextBox text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   25
            Left            =   2205
            TabIndex        =   51
            Top             =   1680
            Width           =   3390
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   25
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   11
            Tag             =   "Cta Vtas Otros|T|S|||variedades|ctavtasotros|||"
            Top             =   1665
            Width           =   1050
         End
         Begin VB.TextBox text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   24
            Left            =   2205
            TabIndex        =   49
            Top             =   1350
            Width           =   3390
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   24
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   10
            Tag             =   "Cta Vtas Retirada|T|S|||variedades|ctavtasretirada|||"
            Top             =   1350
            Width           =   1050
         End
         Begin VB.TextBox text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   23
            Left            =   2205
            TabIndex        =   47
            Top             =   1005
            Width           =   3390
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   23
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   9
            Tag             =   "Cta Vtas Industria|T|S|||variedades|ctavtasindustria|||"
            Top             =   990
            Width           =   1050
         End
         Begin VB.TextBox text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   22
            Left            =   2205
            TabIndex        =   45
            Top             =   645
            Width           =   3390
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   22
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   8
            Tag             =   "Cta Vtas Exportaci�n|T|S|||variedades|ctavtasexportacion|||"
            Top             =   630
            Width           =   1050
         End
         Begin VB.TextBox text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   21
            Left            =   2205
            TabIndex        =   43
            Top             =   285
            Width           =   3390
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   21
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   7
            Tag             =   "Cta Vtas Interior|T|S|||variedades|ctavtasinterior|||"
            Text            =   "1234567890"
            Top             =   270
            Width           =   1050
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   8
            Left            =   810
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   1710
            Width           =   240
         End
         Begin VB.Label Label25 
            Caption         =   "Otros"
            Height          =   255
            Left            =   135
            TabIndex        =   52
            Top             =   1695
            Width           =   465
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   7
            Left            =   810
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   1350
            Width           =   240
         End
         Begin VB.Label Label24 
            Caption         =   "Retirada"
            Height          =   255
            Left            =   135
            TabIndex        =   50
            Top             =   1380
            Width           =   645
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   6
            Left            =   810
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   990
            Width           =   240
         End
         Begin VB.Label Label23 
            Caption         =   "Industria"
            Height          =   255
            Left            =   135
            TabIndex        =   48
            Top             =   1020
            Width           =   1005
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   810
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   630
            Width           =   240
         End
         Begin VB.Label Label22 
            Caption         =   "Export."
            Height          =   255
            Left            =   135
            TabIndex        =   46
            Top             =   660
            Width           =   1005
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   810
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   270
            Width           =   240
         End
         Begin VB.Label Label21 
            Caption         =   "Interior"
            Height          =   255
            Left            =   135
            TabIndex        =   44
            Top             =   300
            Width           =   1005
         End
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   -72660
         TabIndex        =   41
         Top             =   1230
         Width           =   3435
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   -73380
         MaxLength       =   3
         TabIndex        =   4
         Tag             =   "Tipo Unidad|N|S|||variedades|codunida|00||"
         Top             =   1215
         Width           =   690
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -72660
         TabIndex        =   38
         Top             =   870
         Width           =   3435
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   -73380
         MaxLength       =   3
         TabIndex        =   3
         Tag             =   "Clase|N|N|0|999|variedades|codclase|000||"
         Top             =   870
         Width           =   675
      End
      Begin VB.Frame FrameAux0 
         BorderStyle     =   0  'None
         Height          =   3795
         Left            =   -74880
         TabIndex        =   35
         Top             =   495
         Width           =   9055
         Begin VB.TextBox txtAux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   3
            Left            =   1920
            MaxLength       =   3
            TabIndex        =   22
            Tag             =   "Nombre Calibre Abr|T|N|||calibres|nomcalab|||"
            Text            =   "cab"
            Top             =   3600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   0
            Left            =   -120
            MaxLength       =   6
            TabIndex        =   19
            Tag             =   "C�digo Variedad|N|N|1|999999|calibres|codvarie|000000|S|"
            Text            =   "codvar"
            Top             =   3510
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   1
            Left            =   360
            MaxLength       =   2
            TabIndex        =   20
            Tag             =   "Codigo Calibre|N|N|1|99|calibres|codcalib|00|S|"
            Text            =   "ca"
            Top             =   3555
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox txtAux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   2
            Left            =   720
            MaxLength       =   12
            TabIndex        =   21
            Tag             =   "Nombre Calibre|T|N|||calibres|nomcalib|||"
            Text            =   "nomcali"
            Top             =   3555
            Visible         =   0   'False
            Width           =   255
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   0
            Left            =   0
            TabIndex        =   36
            Top             =   0
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
         Begin MSAdodcLib.Adodc AdoAux 
            Height          =   375
            Index           =   0
            Left            =   3720
            Top             =   480
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
            Bindings        =   "frmManVariedad.frx":0094
            Height          =   3465
            Index           =   0
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Width           =   5700
            _ExtentX        =   10054
            _ExtentY        =   6112
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
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   -73380
         MaxLength       =   6
         TabIndex        =   6
         Tag             =   "C.Conselleria|N|N|||variedades|codconse|||"
         Top             =   1890
         Width           =   735
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   -73380
         MaxLength       =   3
         TabIndex        =   2
         Tag             =   "Producto|N|N|0|999|variedades|codprodu|000||"
         Top             =   495
         Width           =   675
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   -72660
         TabIndex        =   23
         Top             =   495
         Width           =   3435
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Variedad"
         Height          =   255
         Index           =   2
         Left            =   -68910
         TabIndex        =   135
         Top             =   930
         Width           =   1380
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Mercancia"
         Height          =   255
         Index           =   1
         Left            =   -68910
         TabIndex        =   133
         Top             =   510
         Width           =   1380
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   0
         Left            =   -65970
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   510
         Width           =   240
      End
      Begin VB.Label Label31 
         Caption         =   "Rdto M�ximo Hda."
         Height          =   255
         Left            =   360
         TabIndex        =   94
         Top             =   3990
         Width           =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "Clasificaci�n"
         Height          =   255
         Index           =   19
         Left            =   360
         TabIndex        =   91
         Top             =   900
         Width           =   1110
      End
      Begin VB.Label Label17 
         Caption         =   "Euros/kg Hanegada"
         Height          =   255
         Left            =   8025
         TabIndex        =   89
         Top             =   855
         Width           =   1770
      End
      Begin VB.Label Label16 
         Caption         =   "Euros/kg Tria"
         Height          =   255
         Left            =   8025
         TabIndex        =   88
         Top             =   540
         Width           =   1770
      End
      Begin VB.Label Label15 
         Caption         =   "Euros/kg Seg.Social"
         Height          =   255
         Left            =   4455
         TabIndex        =   87
         Top             =   1500
         Width           =   2010
      End
      Begin VB.Label Label14 
         Caption         =   "Euros/kg Mano Obra"
         Height          =   255
         Left            =   4455
         TabIndex        =   86
         Top             =   1170
         Width           =   1950
      End
      Begin VB.Label Label13 
         Caption         =   "Euros/kg Recolec."
         Height          =   255
         Left            =   4455
         TabIndex        =   85
         Top             =   825
         Width           =   2040
      End
      Begin VB.Label Label12 
         Caption         =   "Euros/kg Destajo"
         Height          =   255
         Left            =   4455
         TabIndex        =   84
         Top             =   510
         Width           =   1680
      End
      Begin VB.Label Label11 
         Caption         =   "Porcentaje Destrio"
         Height          =   255
         Left            =   360
         TabIndex        =   83
         Top             =   3540
         Width           =   1410
      End
      Begin VB.Label Label10 
         Caption         =   "Porcentaje Mermas"
         Height          =   255
         Left            =   360
         TabIndex        =   82
         Top             =   3225
         Width           =   1410
      End
      Begin VB.Label Label9 
         Caption         =   "Porcentaje Industria"
         Height          =   255
         Left            =   360
         TabIndex        =   81
         Top             =   2940
         Width           =   1680
      End
      Begin VB.Label Label3 
         Caption         =   "Arroba/Jornal"
         Height          =   255
         Left            =   360
         TabIndex        =   80
         Top             =   2430
         Width           =   1410
      End
      Begin VB.Label Label2 
         Caption         =   "Factor Correcci�n Mermas"
         Height          =   255
         Left            =   8025
         TabIndex        =   79
         Top             =   1485
         Width           =   1995
      End
      Begin VB.Label Label8 
         Caption         =   "Factor Correci�n Destrio"
         Height          =   255
         Left            =   8025
         TabIndex        =   78
         Top             =   1170
         Width           =   1860
      End
      Begin VB.Label Label19 
         Caption         =   "Max.Kilos / Cajon"
         Height          =   255
         Left            =   360
         TabIndex        =   77
         Top             =   2025
         Width           =   1410
      End
      Begin VB.Label Label6 
         Caption         =   "Min.Kilos / Cajon"
         Height          =   255
         Left            =   360
         TabIndex        =   76
         Top             =   1665
         Width           =   1230
      End
      Begin VB.Label Label7 
         Caption         =   "Kilos / Cajon"
         Height          =   255
         Left            =   360
         TabIndex        =   75
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   10
         Left            =   -67890
         ToolTipText     =   "Buscar Centro Coste"
         Top             =   1890
         Width           =   240
      End
      Begin VB.Label Label28 
         Caption         =   "Centro Coste"
         Height          =   255
         Left            =   -68925
         TabIndex        =   57
         Top             =   1920
         Width           =   1005
      End
      Begin VB.Label Label27 
         Caption         =   "C�digo IVA"
         Height          =   255
         Left            =   -74685
         TabIndex        =   55
         Top             =   1560
         Width           =   915
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   -73650
         ToolTipText     =   "Buscar Cta.Contable"
         Top             =   1530
         Width           =   240
      End
      Begin VB.Label Label26 
         Caption         =   "C�digo EAN"
         Height          =   255
         Left            =   -65100
         TabIndex        =   53
         Top             =   540
         Width           =   915
      End
      Begin VB.Image imgBuscar 
         Height          =   330
         Index           =   4
         Left            =   -64155
         ToolTipText     =   "C�digos EAN asociados"
         Top             =   495
         Width           =   375
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   -73650
         ToolTipText     =   "Buscar T.Unidad"
         Top             =   1215
         Width           =   240
      End
      Begin VB.Label Label20 
         Caption         =   "Tipo Unidad"
         Height          =   255
         Left            =   -74685
         TabIndex        =   40
         Top             =   1245
         Width           =   915
      End
      Begin VB.Label Label5 
         Caption         =   "Clase"
         Height          =   255
         Left            =   -74685
         TabIndex        =   39
         Top             =   855
         Width           =   735
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   -73650
         ToolTipText     =   "Buscar Clase"
         Top             =   855
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   -73650
         ToolTipText     =   "Buscar Producto"
         Top             =   495
         Width           =   240
      End
      Begin VB.Label Label18 
         Caption         =   "Producto"
         Height          =   255
         Left            =   -74685
         TabIndex        =   31
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "C�d.Conseller�a"
         Height          =   255
         Index           =   26
         Left            =   -74685
         TabIndex        =   30
         Top             =   1920
         Width           =   1320
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   4245
      Top             =   6105
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
      TabIndex        =   33
      Top             =   0
      Width           =   12075
      _ExtentX        =   21299
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
            Object.ToolTipText     =   "Copia Calibres/Calidades"
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
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Index           =   0
         Left            =   8520
         TabIndex        =   34
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10800
      TabIndex        =   32
      Top             =   6030
      Visible         =   0   'False
      Width           =   1035
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
      Begin VB.Menu mnCopiaCalibres 
         Caption         =   "Copia Calibres/Calidades"
         Shortcut        =   ^C
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
Attribute VB_Name = "frmManVariedad"
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
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1

Private WithEvents frmIva As frmTipIVAConta
Attribute frmIva.VB_VarHelpID = -1

Private WithEvents frmPro As frmManProductos  'Productos
Attribute frmPro.VB_VarHelpID = -1
Private WithEvents frmCla As frmManClases  'Clase
Attribute frmCla.VB_VarHelpID = -1
Private WithEvents frmTra As frmTraerCalib   'Traer calibres
Attribute frmTra.VB_VarHelpID = -1
Private WithEvents frmTUn As frmManTipUnid   'Tipos de unidad
Attribute frmTUn.VB_VarHelpID = -1

Private WithEvents frmCtas As frmCtasConta 'cuentas contables
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmCCos As frmCCosConta 'centros de coste
Attribute frmCCos.VB_VarHelpID = -1


Private WithEvents frmCEan As frmCodEAN 'Codigos Ean
Attribute frmCEan.VB_VarHelpID = -1


' *****************************************************


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
Dim cadB As String

Private Sub cmbAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim Produ As Long
Dim vCadena As String

    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'B�SQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm2(Me, 1) Then
                    
                    CargarUnaVariedad CLng(Text1(0).Text), "I"
                    
                    Data1.RecordSource = "Select * from " & NombreTabla & Ordenacion
                    TerminaBloquear
                    PosicionarData
                End If
            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario2(Me, 1) Then
                    '[Monica]18/09/2013: si estamos actualizando variedad en Picassent el claveant en 'PP&VVVV'
                    Produ = DevuelveValor("select codprodu from variedades where codvarie = " & DBSet(Text1(0).Text, "N"))
                    vCadena = CLng(Produ) & "&" & CLng(Text1(0).Text)
                    
                    CargarUnaVariedad CLng(Text1(0).Text), "U", vCadena
                    
                    TerminaBloquear
                    PosicionarData
                End If
            Else
                ModoLineas = 0
            End If
        ' *** si n'hi han ll�nies ***
        Case 5 'LL�NIES
            Select Case ModoLineas
                Case 1 'afegir ll�nia
                    InsertarLinea
                Case 2 'modificar ll�nies
                    If ModificarLinea Then
                        PosicionarData
                    Else
                        PonerFoco txtAux(12)
                    End If
            End Select
        ' **************************
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

' *** si n'hi han combos a la cap�alera ***
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
                Text1(0).BackColor = vbYellow 'codvarie
                ' ****************************************************************************
            End If
        End If
    End If
    SSTab1.TabEnabled(2) = ExisteTabla("rcalidad")
    SSTab1.TabVisible(2) = ExisteTabla("rcalidad")
    SSTab1.TabEnabled(3) = ExisteTabla("rcalidad")
    SSTab1.TabVisible(3) = ExisteTabla("rcalidad")
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
        .Buttons(11).Image = 22   'Copiar Calidades y Calibres
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
    Me.imgBuscar(4).Picture = frmPpal.imgListComun.ListImages(21).Picture
    
    imgAyuda(0).Picture = frmPpal.ImageListB.ListImages(10).Picture

    CargaCombo
    
    ' *** si n'hi han tabs, per a que per defecte sempre es pose al 1r***
    Me.SSTab1.Tab = 0
    ' *******************************************************************
    
    LimpiarCampos   'Neteja els camps TextBox
    ' ******* si n'hi han ll�nies *******
    DataGridAux(0).ClearFields
'    DataGridAux(1).ClearFields
    
    '*** canviar el nom de la taula i l'ordenaci� de la cap�alera ***
    NombreTabla = "variedades"
    Ordenacion = " ORDER BY codvarie"
    
    'Mirem com est� guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    Data1.ConnectionString = conn
    '***** cambiar el nombre de la PK de la cabecera *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where codvarie=-1"
    Data1.Refresh
    
    CargaGrid 0, False
    If ExisteTabla("rcalidad") Then CargaGrid 1, False
    
    ModoLineas = 0
       
    ' Para el chivato
    Set dbAriagro = New BaseDatos
    dbAriagro.abrir_MYSQL vConfig.SERVER, vUsu.CadenaConexion, vConfig.User, vConfig.password
    
    '[Monica]14/11/2014: si estamos en catadau la cuenta de compras de terceros no la usan
    '                    pasa a ser la cuenta de liquidacion de industria
    If vParamAplic.Cooperativa = 0 Then
        Label34.Caption = "Liquidacion Industria"
        Text1(33).Tag = "Cta.Liquidaci�n Industria|T|S|||variedades|ctacomtercero|||"
    End If
    
    '[Monica]15/03/2016: Si es Abn ponemos los gastos molturacion y envasado en la variedad
    If vParamAplic.Cooperativa = 1 Then
        Label14.Caption = "Euros/kg Gtos.Molturaci�n"
        Text1(17).Tag = "Euros/kg Gtos.Molturaci�n|N|S|||variedades|eurmanob|0.0000||"
        Label15.Caption = "Euros/litro Gtos.Envasado"
        Text1(18).Tag = "Euros/kg Gtos.Envasado|N|S|||variedades|eursegsoc|0.0000||"
        Label12.Caption = "Precio Venta"
        Text1(15).Tag = "Precio Venta|N|S|||variedades|eurdesta|0.0000||"
        Label13.Caption = "Precio Excedido"
        Text1(16).Tag = "Precio Excedido|N|S|||variedades|eurecole|0.0000||"
    End If
    
    
    
'    If DatosADevolverBusqueda = "" Then
'        PonerModo 0
'    Else
'        PonerModo 1 'b�squeda
'        ' *** posar de groc els camps visibles de la clau primaria de la cap�alera ***
'        text1(0).BackColor = vbYellow 'codclien
'        ' ****************************************************************************
'    End If
End Sub

Private Sub LimpiarCampos()
    On Error Resume Next
    
    limpiar Me   'M�tode general: Neteja els controls TextBox
    Me.Combo1(0).ListIndex = -1
    Me.Combo1(1).ListIndex = -1
    Me.Combo1(2).ListIndex = -1
    lblIndicador.Caption = ""
    
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
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Nom�s es per a saber que n'hi ha + d'1 registre
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    
    '---------------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
       
    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
    'Si estem en Insertar a m�s neteja els camps Text1
    BloquearText1 Me, Modo
    
    '++monica: si el modo es insertar damos el siguiente pero dejamos modificar
    If Modo = 3 Then Text1(0).Locked = False
    
    ' ***** bloquejar tots els controls visibles de la clau primaria de la cap�alera ***
'    If Modo = 4 Then _
'        BloquearTxt Text1(0), True 'si estic en  modificar, bloqueja la clau primaria
    ' **********************************************************************************
    
    ' **** si n'hi han imagens de buscar en la cap�alera *****
    BloquearImgBuscar Me, Modo, ModoLineas
    ' ********************************************************
    Me.imgBuscar(4).Enabled = (Modo = 2)
    Me.imgBuscar(4).visible = (Modo = 2)
    Me.Label26.visible = (Modo = 2)
    
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    BloquearCombo Me, Modo
    PonerLongCampos

    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 0, False
        If ExisteTabla("rcalidad") Then CargaGrid 1, False
    End If
    
    b = (Modo = 4) Or (Modo = 2)
    DataGridAux(0).Enabled = b
    DataGridAux(1).Enabled = b
      
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
    'copiar calibres y calidades
    Toolbar1.Buttons(11).Enabled = b
    Me.mnCopiaCalibres.Enabled = b
    
    'Imprimir
    'Toolbar1.Buttons(12).Enabled = (b Or Modo = 0)
    Toolbar1.Buttons(12).Enabled = True And Not DeConsulta
       
' no se utiliza el toolaux
'    ' *** si n'hi han ll�nies que tenen grids (en o sense tab) ***
'    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
'    For i = 0 To ToolAux.Count - 1
'        ToolAux(i).Buttons(1).Enabled = b
'        If b Then bAux = (b And Me.Adoaux(i).Recordset.RecordCount > 0)
'        ToolAux(i).Buttons(2).Enabled = bAux
'        ToolAux(i).Buttons(3).Enabled = bAux
'    Next i
    
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
               
        Case 0 'CALIBRES
            Sql = "SELECT codvarie, codcalib, nomcalib, nomcalab"
            Sql = Sql & " FROM calibres "
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE calibres.codvarie = -1"
            End If
            Sql = Sql & " ORDER BY calibres.codcalib"
               
        Case 1 'CALIDADES
            Sql = "SELECT rcalidad.codvarie,codcalid,nomcalid, nomcalab, tipcalid, CASE rcalidad.tipcalid WHEN 0 THEN ""Normal"" WHEN 1 THEN ""Destrio (S�lo una)"" WHEN 2 THEN ""Venta Campo"" END,  "
            Sql = Sql & "rcalidad.tipcalid1, "
            Sql = Sql & " CASE rcalidad.tipcalid1 WHEN 0 THEN ""Comercial"" WHEN 1 THEN ""No Comercial"" WHEN 2 THEN ""Retirada"" END, "
            Sql = Sql & " nomcalibrador1, nomcalibrador2, gastosrec, IF(gastosrec=1,'*','') as dgastorec "
            Sql = Sql & " FROM rcalidad"
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " where codvarie = -1"
            End If
            
            Sql = Sql & " ORDER BY codcalid"
            
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
        cadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        cadB = Aux
        '   Com la clau principal es �nica, en posar el sql apuntant
        '   al valor retornat sobre la clau ppal es suficient
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        ' **********************************
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub


Private Sub frmCCos_DatoSeleccionado(CadenaSeleccion As String)
'Centro de Coste de la contabilidad
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'des macta
End Sub

Private Sub frmIva_DatoSeleccionado(CadenaSeleccion As String)
'Tipo de iva de la contabilidad
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigiva
    FormateaCampo Text1(indice)
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Porceiva
End Sub

Private Sub frmTra_Actualizar(vValor As String)
    On Error GoTo EEPonerBusq
    
    LimpiarCampos
    Text1(0).Text = vValor 'codvarie
    
    FormateaCampo Text1(0)
    
    Screen.MousePointer = vbHourglass
    
    If vValor = "" Then vValor = " codvarie = -1"
    Data1.RecordSource = "select * from variedades where " & vValor
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
    
'        Modo = 1
'        cmdAceptar_Click
End Sub

Private Sub frmTUn_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de tipos de unidad
    Text1(9).Text = RecuperaValor(CadenaSeleccion, 1) 'tipos de unidad
    FormateaCampo Text1(9)
    Text2(9).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre de unidad
End Sub


Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           ' "____________________________________________________________"
            vCadena = "Indica si la variedad es mercancia que es suministrada por el cliente, o si por el contrario es mercancia de la cooperativa." & vbCrLf & vbCrLf & _
                      "Se utiliza para restringir las variedades a mostrar en los informes. " & vbCrLf & _
                      vbCrLf
                      
                      
    End Select
    MsgBox vCadena, vbInformation, "Descripci�n de Ayuda"
    
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnBuscarCalibre_Click()
'    Set frmTra = New frmTraerCalib
'    frmTra.DatosADevolverBusqueda = "0|1|"
'    frmTra.CodigoActual = text1(0).Text
'    frmTra.Show vbModal
'    Set frmTra = Nothing
'    PonerFoco text1(0)
End Sub

Private Sub mnCopiaCalibres_Click()

    If Data1.Recordset.EOF Then Exit Sub
    
    frmCopiaCalibCalid.NumCod = Data1.Recordset!codvarie
    frmCopiaCalibCalid.Show vbModal
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    AbrirListado (12)
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

Private Sub SSTab1_Click(PreviousTab As Integer)
    If PreviousTab = 0 Then
        PonerFocoGrid DataGridAux(0)
    End If
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
        Case 11  'Buscar Tarjeta
            mnCopiaCalibres_Click
'            mnBuscarCalibre_Click
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
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        PonerFoco Text1(0) ' <===
        Text1(0).BackColor = vbYellow ' <===
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

    cadB = ObtenerBusqueda2(Me, 1)
    
    If chkVistaPrevia(0) = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        ' *** foco al 1r camp visible de la cap�alera que siga clau primaria ***
        PonerFoco Text1(0)
        ' **********************************************************************
    End If
End Sub

Private Sub MandaBusquedaPrevia(cadB As String)
    Dim cad As String
        
    'Cridem al form
    ' **************** arreglar-ho per a vore lo que es desije ****************
    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
    cad = ""
    cad = cad & ParaGrid(Text1(0), 15, "C�d.")
    cad = cad & ParaGrid(Text1(1), 60, "Nombre")
    cad = cad & ParaGrid(Text1(2), 25, "Producto")
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vtabla = NombreTabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        frmB.vDevuelve = "0|1|2|" '*** els camps que volen que torne ***
        frmB.vTitulo = "Variedades" ' ***** repasa a��: t�tol de BuscaGrid *****
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
Dim j As Integer

    If Data1.Recordset.EOF Then
        MsgBox "Ning�n registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    cad = ""
    I = 0
    Do
        j = I + 1
        I = InStr(j, DatosADevolverBusqueda, "|")
        If I > 0 Then
            Aux = Mid(DatosADevolverBusqueda, j, I - j)
            j = Val(Aux)
            cad = cad & Text1(j).Text & "|"
        End If
    Loop Until I = 0
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
    cadB = ""
    
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
    
    '******************** canviar taula i camp **************************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        NumF = NuevoCodigo
    Else
        NumF = SugerirCodigoSiguienteStr("variedades", "codvarie")
    End If
    '********************************************************************
    
    
    ' codEmpre i quins camps tenen la PK de la cap�alera *******
    Text1(0).Text = NumF
    FormateaCampo Text1(0)
       
    PonerFoco Text1(0) '*** 1r camp visible que siga PK ***
    
    ' *** si n'hi han camps de descripci� a la cap�alera ***
    'PosarDescripcions

    ' *** si n'hi han tabs, em posicione al 1r ***
    Me.SSTab1.Tab = 0
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
    cad = "�Seguro que desea eliminar la Variedad?"
    cad = cad & vbCrLf & "C�digo: " & Format(Data1.Recordset.Fields(0), FormatoCampo(Text1(0)))
    cad = cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(2)
    
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
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Variedad", Err.Description
End Sub

Private Sub PonerCampos()
Dim I As Integer
Dim codpobla As String, despobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la cap�alera
    
    ' *** si n'hi han ll�nies en datagrids ***
    'For i = 0 To DataGridAux.Count - 1
    For I = 0 To 1
            If Not ExisteTabla("rcalidad") Then Exit For
            
            CargaGrid I, True
            If Not AdoAux(I).Recordset.EOF Then _
                PonerCamposForma2 Me, AdoAux(I), 2, "FrameAux" & I
    Next I

    If vParamAplic.NumeroConta <> 0 Then
        Text2(26).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(26), "N")
    
    ' ************* configurar els camps de les descripcions de la cap�alera *************
'    Text2(21).Text = NombreCuentaCorrecta(Text1(21).Text)
'    Text2(22).Text = NombreCuentaCorrecta(Text1(22).Text)
'    Text2(23).Text = NombreCuentaCorrecta(Text1(23).Text)
'    Text2(24).Text = NombreCuentaCorrecta(Text1(24).Text)
'    Text2(25).Text = NombreCuentaCorrecta(Text1(25).Text)
        Text2(21).Text = PonerNombreCuenta(Text1(21), Modo)
        Text2(22).Text = PonerNombreCuenta(Text1(22), Modo)
        Text2(23).Text = PonerNombreCuenta(Text1(23), Modo)
        Text2(24).Text = PonerNombreCuenta(Text1(24), Modo)
        Text2(25).Text = PonerNombreCuenta(Text1(25), Modo)
        If vParamAplic.ContabilidadNueva Then
            Text2(27).Text = PonerNombreDeCod(Text1(27), "ccoste", "nomccost", "codccost", "T", cConta)
        Else
            Text2(27).Text = PonerNombreDeCod(Text1(27), "cabccost", "nomccost", "codccost", "T", cConta)
        End If
    
        Text2(28).Text = PonerNombreCuenta(Text1(28), Modo)
        Text2(29).Text = PonerNombreCuenta(Text1(29), Modo)
    
        Text2(34).Text = PonerNombreCuenta(Text1(34), Modo)
        
        '[Monica]18/07/2012: las ejecuto solo en el caso de que coindicida la contabilidad de horto con la de comercial
        '                    las vuelvo a poner
            '[Monica]23/03/2012: las cuentas de recoleccion puede que no sean de la conta de parametros
            '                    seran de la seccion que corresponda de recoleccion
            '                    quito las 6 instrucciones siguientes
        If vParamAplic.NumeroConta = DevuelveValor("select empresa_conta from rseccion, rparam where rseccion.codsecci = rparam.seccionhorto") Then
            Text2(31).Text = PonerNombreCuenta(Text1(31), Modo)
            Text2(32).Text = PonerNombreCuenta(Text1(32), Modo)
            Text2(33).Text = PonerNombreCuenta(Text1(33), Modo)
            Text2(35).Text = PonerNombreCuenta(Text1(35), Modo)
            Text2(36).Text = PonerNombreCuenta(Text1(36), Modo)
            Text2(37).Text = PonerNombreCuenta(Text1(37), Modo)
        End If
    Else
        Text2(26).Text = DevuelveDesdeBDNew(cAgro, "tiposiva", "nombriva", "nombriva", Text1(26), "N")
    End If
    Text2(2).Text = PonerNombreDeCod(Text1(2), "productos", "nomprodu", "codprodu", "N")
    Text2(3).Text = PonerNombreDeCod(Text1(3), "clases", "nomclase", "codclase", "N")
    Text2(9).Text = PonerNombreDeCod(Text1(9), "sunida", "nomunida", "codunida", "N")
    ' ********************************************************************************
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu
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
                    
                    ' *** si n'hi han tabs ***
                    SituarTab (NumTabMto + 1)
                    
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        AdoAux(NumTabMto).Recordset.MoveFirst
                    End If

                Case 2 'modificar ll�nies
                    ModoLineas = 0
                    
                    ' *** si n'hi han tabs ***
                    SituarTab (NumTabMto + 1)
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
Dim Sql As String
'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm2(Me, 1)
    If Not b Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
    If (Modo = 3) Then 'insertar
        'comprobar si existe ya el cod. del campo clave primaria
        If ExisteCP(Text1(0)) Then b = False
    End If
    
    If b Then
        If vEmpresa.TieneAnalitica Then 'hay contab. analitica
            If Not vParamAplic.ContabilidadNueva Then
                Sql = DevuelveDesdeBDNew(cConta, "cabccost", "codccost", "codccost", Text1(27), "T")
            Else
                Sql = DevuelveDesdeBDNew(cConta, "ccoste", "codccost", "codccost", Text1(27), "T")
            End If
             If Sql = "" Then
                MsgBox "No existe el Centro de Coste. Reintroduzca.", vbExclamation
                PonerFoco Text1(27)
                b = False
             End If
        End If
    End If

    ' ************************************************************************************
    
    DatosOk = b
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Sub PosicionarData()
Dim cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la cap�alera, no llevar els () ***
    cad = "(codvarie=" & Text1(0).Text & ")"
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(Data1, cad, Indicador) Then
    If SituarData(Data1, cad, Indicador) Then
        If ModoLineas <> 1 Then PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
End Sub

Private Function Eliminar() As Boolean
Dim vWhere As String

    On Error GoTo FinEliminar

    conn.BeginTrans
    ' ***** canviar el nom de la PK de la cap�alera, repasar codEmpre *******
    vWhere = " WHERE codvarie=" & Data1.Recordset!codvarie
        
    ' ***** elimina les ll�nies ****
    conn.Execute "DELETE FROM variane " & vWhere
        
    conn.Execute "DELETE FROM calibres " & vWhere
    
    If ExisteTabla("rcalidad") Then
        conn.Execute "DELETE FROM rcalidad " & vWhere
    End If
        
    CargarUnaVariedad CLng(Data1.Recordset!codvarie), "D"
        
    'Eliminar la CAP�ALERA
    vWhere = " WHERE codvarie=" & Data1.Recordset!codvarie
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
Dim NumDigit As String

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    ' ***************** configurar els LostFocus dels camps de la cap�alera *****************
    Select Case Index
        Case 0 'cod variedad
            PonerFormatoEntero Text1(0)

        Case 1 'NOMBRE
            Text1(Index).Text = UCase(Text1(Index).Text)
        
        Case 2 'Producto
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "productos", "nomprodu")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Producto: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "�Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmPro = New frmManProductos
                        frmPro.DatosADevolverBusqueda = "0|1|"
                        frmPro.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmPro.Show vbModal
                        Set frmPro = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 3 'clase
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index) = PonerNombreDeCod(Text1(Index), "clases", "nomclase", "codclase", "N")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Clase: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "�Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCla = New frmManClases
                        frmCla.DatosADevolverBusqueda = "0|1|"
                        frmCla.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmCla.Show vbModal
                        Set frmCla = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
                
        Case 9 'Tipo de Unidad
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "sunida", "nomunida")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Tipo de Unidad: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "�Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmTUn = New frmManTipUnid
                        frmTUn.DatosADevolverBusqueda = "0|1|"
                        frmTUn.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmTUn.Show vbModal
                        Set frmTUn = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
        Case 5, 6, 7, 8, 10, 11, 12, 13, 14
            If Modo = 1 Then Exit Sub
            cadMen = TransformaPuntosComas(Text1(Index).Text)
            Text1(Index).Text = Format(cadMen, "##0.00")
        Case 15, 16, 17, 18, 19, 20
            If Modo = 1 Then Exit Sub
            cadMen = TransformaPuntosComas(Text1(Index).Text)
            Text1(Index).Text = Format(cadMen, "0.0000")
        
        Case 21, 22, 23, 24, 25, 28, 29, 31, 32, 33, 34, 35, 36, 37 'cta contable de ventas
            If Text1(Index).Text = "" Then
                Text2(Index) = ""
                Exit Sub
            End If
            
'            If Modo <> 1 Then
'                NumDigit = DevuelveDesdeBDNew(cConta, "empresa", "numdigi3", "codempre", vParamAplic.NumeroConta, "N")
'                If Len(Text1(21).Text) <> CCur(NumDigit) Then
'                    MsgBox "La longitud de la cuenta no se corresponde con el nivel 3.", vbExclamation
'                End If
'            End If
            
'            Text2(Index).Text = NombreCuentaCorrecta(Text1(Index).Text)
            Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo)
            
            If Index = 37 Then cmdAceptar.SetFocus
    
    
        Case 26 ' tipo de iva de contabilidad
            'conConta: BD Contabilidad
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "tiposiva", "nombriva", , , cConta)
            Else
                Text2(Index).Text = ""
            End If
            
        Case 27 ' centro de coste
            If Text1(Index).Text <> "" Then
                If vParamAplic.NumeroConta <> 0 Then
                    If vParamAplic.ContabilidadNueva Then
                        Text2(Index).Text = PonerNombreDeCod(Text1(Index), "ccoste", "nomccost", "codccost", "T", cConta)
                    Else
                        Text2(Index).Text = PonerNombreDeCod(Text1(Index), "cabccost", "nomccost", "codccost", "T", cConta)
                    End If
                End If
            End If
        
        Case 30
            If Text1(Index).Text <> "" Then PonerFormatoEntero Text1(Index)
    
            
    End Select
        ' ***************************************************************************
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 2: KEYBusqueda KeyAscii, 0 'producto
                Case 3: KEYBusqueda KeyAscii, 1 'clase
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
        Case 0 'calibres
            Sql = "�Seguro que desea eliminar el Calibre?"
            Sql = Sql & vbCrLf & "Calibre: " & AdoAux(Index).Recordset!codcalib
            Sql = Sql & vbCrLf & "Nombre: " & AdoAux(Index).Recordset!nomcalib
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM calibres"
                Sql = Sql & vWhere & " AND codcalib= " & AdoAux(Index).Recordset!codcalib
            End If
            
        Case 1 'variedades anecoop
            Sql = "�Seguro que desea eliminar la Variedad Anecoop?"
            Sql = Sql & vbCrLf & "C�digo: " & AdoAux(Index).Recordset!numlinea
            Sql = Sql & vbCrLf & "Nombre: " & AdoAux(Index).Recordset!codvaane
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM variane"
                Sql = Sql & vWhere & " AND numlinea= " & AdoAux(Index).Recordset!numlinea
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
        SituarTab (NumTabMto + 1)
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
        Case 0: vtabla = "calibres"
        Case 1: vtabla = "variane"
    End Select
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 0, 1 ' *** pose els index dels tabs de ll�nies que tenen datagrid ***
            ' *** canviar la clau primaria de les ll�nies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            If Index = 0 Then
                NumF = SugerirCodigoSiguienteStr(vtabla, "codcalib", vWhere)
            Else
                NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", vWhere)
            End If

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
                    txtAux(0).Text = Text1(0).Text 'codvarie
'                    txtAux(3).Text = text1(1).Text 'nomcalibre
                    txtAux(1).Text = NumF 'codcalib
                    txtAux(2).Text = ""
                    txtAux(3).Text = ""
                    txtAux(4).Text = ""
                    PonerFoco txtAux(1)
                Case 1 'variedades anecoop
                    txtAux(8).Text = Text1(0).Text 'codvarie
                    txtAux(9).Text = NumF 'numlinea
                    txtAux(10).Text = ""
                    
                    PonerFoco txtAux(9)
                    
            End Select
    End Select
End Sub

Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim I As Integer
    Dim j As Integer
    
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
        Case 0 'calibres
        
            For j = 0 To 3
                txtAux(j).Text = DataGridAux(Index).Columns(j).Text
            Next j
            
            For I = 0 To 1
                BloquearTxt txtAux(I), False
            Next I
            
        Case 1 'variedades anecoop
            For j = 8 To 10
                txtAux(j).Text = DataGridAux(Index).Columns(j - 8).Text
            Next j
            
            For I = 8 To 9
                BloquearTxt txtAux(I), False
            Next I
            
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 0 'cuentas bancarias
            PonerFoco txtAux(2)
        Case 1 'departamentos
            PonerFoco txtAux(10)
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
            For jj = 1 To 4
                txtAux(jj).visible = b
                txtAux(jj).Top = alto
            Next jj
            
        Case 1 'variedades anecoop
            For jj = 9 To 10
                txtAux(jj).visible = b
                txtAux(jj).Top = alto
            Next jj
            
    End Select
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
    
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    

    ' ******* configurar el LostFocus dels camps de ll�nies (dins i fora grid) ********
    Select Case Index
            
        Case 2, 3 ' Nombre de calibre y calibre abreviado
            txtAux(Index).Text = UCase(txtAux(Index).Text)
            
        Case 10 ' variedad anecoop
            txtAux(Index).Text = UCase(txtAux(Index).Text)
            
    End Select
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
   If Not txtAux(Index).MultiLine Then ConseguirFocoLin txtAux(Index)
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not txtAux(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not txtAux(Index).MultiLine Then
        If KeyAscii = teclaBuscar Then
            If Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) Then
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
Dim numNivel As Byte

    TerminaBloquear
     Select Case Index
        Case 0 'productos
            Set frmPro = New frmManProductos
            frmPro.DatosADevolverBusqueda = "0|1|"
            frmPro.CodigoActual = Text1(2).Text
            frmPro.Show vbModal
            Set frmPro = Nothing
            PonerFoco Text1(2)
            
        Case 1 'clases
            Set frmCla = New frmManClases
            frmCla.DatosADevolverBusqueda = "0|1|"
            frmCla.CodigoActual = Text1(3).Text
            frmCla.Show vbModal
            Set frmCla = Nothing
            PonerFoco Text1(3)
        
        Case 2 'tipos de unidad
            Set frmTUn = New frmManTipUnid
            frmTUn.DatosADevolverBusqueda = "0|1|"
            frmTUn.CodigoActual = Text1(3).Text
            frmTUn.Show vbModal
            Set frmTUn = Nothing
            PonerFoco Text1(9)
        
        Case 3, 5, 6, 7, 8 'cuenta contable de venta
            If vParamAplic.NumeroConta = 0 Then Exit Sub
        
            If Index = 3 Then
                indice = Index + 18
            Else
                indice = Index + 17
            End If
            
            Set frmCtas = New frmCtasConta
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(indice).Text
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(indice)
            
       Case 13, 14, 15 ' cuentas contables de recoleccion
            If vParamAplic.NumeroConta = 0 Then Exit Sub
        
            indice = Index + 18
            
            Set frmCtas = New frmCtasConta
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(indice).Text
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(indice)
       
                   
       Case 4 'codigos ean de esa variedad
            Set frmCEan = New frmCodEAN
            frmCEan.Tipo = 2
            frmCEan.CodigoActual = CStr(Me.Data1.Recordset!codvarie)
            frmCEan.Show vbModal
            Set frmCEan = Nothing
             
       Case 9  'Porcentaje iva
            If vParamAplic.NumeroConta = 0 Then Exit Sub
            indice = 26
            Set frmIva = New frmTipIVAConta
            frmIva.DeConsulta = True
            frmIva.DatosADevolverBusqueda = "0|1|2|"
            frmIva.CodigoActual = Text1(indice).Text
            frmIva.Show vbModal
            Set frmIva = Nothing
            PonerFoco Text1(indice)
       
       Case 10 'Centro de Coste
            If vParamAplic.NumeroConta = 0 Then Exit Sub
            indice = 27
            Set frmCCos = New frmCCosConta
            frmCCos.DatosADevolverBusqueda = "0|1|"
            frmCCos.CodigoActual = Text1(indice).Text
            frmCCos.Show vbModal
            Set frmCCos = Nothing
            PonerFoco Text1(indice)
    
        Case 11, 12 'cuentas contables de transporte
            If vParamAplic.NumeroConta = 0 Then Exit Sub
        
            indice = Index + 17
            
            Set frmCtas = New frmCtasConta
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(indice).Text
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(indice)
    
        Case 16, 17, 18, 19
            If vParamAplic.NumeroConta = 0 Then Exit Sub
        
            indice = Index + 18
            
            Set frmCtas = New frmCtasConta
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(indice).Text
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(indice)
    
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub

Private Sub frmCla_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Clases
    Text1(3).Text = RecuperaValor(CadenaSeleccion, 1) 'codclase
    FormateaCampo Text1(3)
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'nomclase
End Sub

Private Sub frmPro_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento Porductos
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1) 'codprodu
    FormateaCampo Text1(2)
    Text2(2).Text = RecuperaValor(CadenaSeleccion, 2) 'nomprodu
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
Private Sub SituarTab(numTab As Integer)
    On Error Resume Next
    
    SSTab1.Tab = numTab
    
    If Err.Number <> 0 Then Err.Clear
End Sub
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
            tots = "N||||0|;S|txtAux(1)|T|Cod|600|;" 'codvarie,codcalib
            tots = tots & "S|txtAux(2)|T|Nombre|3500|;" ' nombre del calibre
            tots = tots & "S|txtAux(3)|T|Abrev.|1000|;" ' nombre de calibre abreviado
            
            arregla tots, DataGridAux(Index), Me
        
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
        Case 1 'Calidades
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;S|txtAux1(1)|T|Cod|600|;" 'codvarie,numlinea
            tots = tots & "S|txtAux1(2)|T|Nombre|2760|;"
            tots = tots & "S|txtAux1(3)|T|Abrev.|1000|;"
            tots = tots & "N||||0|;S|txtAux1(5)|T|Tipo|1500|;"
            tots = tots & "N||||0|;S|txtAux1(7)|T|Tipo|1500|;"
            tots = tots & "S|txtAux1(8)|T|Calibrador 1|1500|;"
            tots = tots & "S|txtAux1(9)|T|Calibrador 2|1500|;"
            tots = tots & "N||||0|;S|chkAux(0)|CB|GR|360|;"
            
            arregla tots, DataGridAux(Index), Me

            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
    End Select
    
    DataGridAux(Index).ScrollBars = dbgAutomatic
      
    ' **** si n'hi han ll�nies en grids i camps fora d'estos ****
    If Not AdoAux(Index).Recordset.EOF Then
        DataGridAux_RowColChange Index, 1, 1
    Else
'        LimpiarCamposFrame Index
    End If
      
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub

Private Sub InsertarLinea()
'Inserta registre en les taules de Ll�nies
Dim nomFrame As String
Dim b As Boolean

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomFrame = "FrameAux0" 'calibres
        Case 1: nomFrame = "FrameAux1" 'variedades anecoop
    End Select
    
    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomFrame) Then
            b = BLOQUEADesdeFormulario2(Me, Data1, 1)
            Select Case NumTabMto
                Case 0, 1 ' *** els index de les llinies en grid (en o sense tab) ***
                     CargaGrid NumTabMto, True
                    If b Then BotonAnyadirLinea NumTabMto
            End Select
           
            SituarTab (NumTabMto + 1)
        End If
    End If
End Sub

Private Function ModificarLinea() As Boolean
'Modifica registre en les taules de Ll�nies
Dim nomFrame As String
Dim V As Integer
    
    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomFrame = "FrameAux0" 'calibres
        Case 1: nomFrame = "FrameAux1" 'variedades anecoop
    End Select
    ModificarLinea = False
    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomFrame) Then
            ModoLineas = 0
            
            V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el n� de llinia
            CargaGrid NumTabMto, True
            
            ' *** si n'hi han tabs ***
            SituarTab (NumTabMto + 1)

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
    ' *** canviar-ho per la clau primaria de la cap�alera ***
    vWhere = vWhere & " codvarie=" & Val(Text1(0).Text)
    
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

' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del rat�n.
Private Sub DataGridAux_GotFocus(Index As Integer)
  WheelHook DataGridAux(Index)
End Sub
Private Sub DataGridAux_LostFocus(Index As Integer)
  WheelUnHook
End Sub

' ********* si n'hi han combos a la cap�alera ************
Private Sub CargaCombo()
Dim I As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For I = 0 To Combo1.Count - 1
        Combo1(I).Clear
    Next I
    
    Combo1(0).AddItem "Campo"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Almac�n"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    

    Combo1(1).AddItem "Cooperativa"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Ajena"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    
    Combo1(2).AddItem "Convencional"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 0
    Combo1(2).AddItem "Biol�gica"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 1

End Sub

