VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCCListados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6615
   Icon            =   "frmCCListados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCobros 
      Height          =   7320
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6435
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
         Index           =   1
         Left            =   4215
         Style           =   2  'Dropdown List
         TabIndex        =   136
         Tag             =   "Tipo Mercancia|N|N|||variedades|tipovariedad||N|"
         Top             =   4140
         Width           =   1845
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   26
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1515
         Width           =   875
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   25
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   0
         Top             =   1110
         Width           =   875
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   25
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   131
         Text            =   "Text5"
         Top             =   1125
         Width           =   3450
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   26
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   130
         Text            =   "Text5"
         Top             =   1515
         Width           =   3450
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Mostrar forfaits"
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
         Left            =   480
         TabIndex        =   129
         Top             =   4950
         Width           =   1905
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   5
         Top             =   3465
         Width           =   1335
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   4
         Top             =   3060
         Width           =   1335
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   1
         Left            =   3180
         Locked          =   -1  'True
         TabIndex        =   125
         Text            =   "Text5"
         Top             =   3465
         Width           =   2940
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   0
         Left            =   3165
         Locked          =   -1  'True
         TabIndex        =   124
         Text            =   "Text5"
         Top             =   3075
         Width           =   2940
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   24
         Left            =   2700
         MaxLength       =   10
         TabIndex        =   9
         Tag             =   "Código Postal|T|S|||clientes|codposta|0000||"
         Top             =   4530
         Width           =   780
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   22
         Left            =   2700
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "Código Postal|T|S|||clientes|codposta|0000||"
         Top             =   4170
         Width           =   780
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   5
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Text5"
         Top             =   2505
         Width           =   3450
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   4
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "Text5"
         Top             =   2115
         Width           =   3450
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   3
         Top             =   2520
         Width           =   875
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   4
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   2
         Top             =   2115
         Width           =   875
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "Código Postal|T|S|||clientes|codposta|00||"
         Top             =   4530
         Width           =   690
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   6
         Tag             =   "Código Postal|T|S|||clientes|codposta|00||"
         Top             =   4170
         Width           =   690
      End
      Begin VB.CommandButton CmdCancel 
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
         Index           =   0
         Left            =   5025
         TabIndex        =   11
         Top             =   6645
         Width           =   975
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
         Left            =   3840
         TabIndex        =   10
         Top             =   6645
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar pb5 
         Height          =   225
         Left            =   510
         TabIndex        =   111
         Top             =   5745
         Visible         =   0   'False
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo de Mercancia"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   225
         Index           =   47
         Left            =   4245
         TabIndex        =   135
         Top             =   3870
         Width           =   2040
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
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
         Index           =   46
         Left            =   780
         TabIndex        =   134
         Top             =   1125
         Width           =   690
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
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
         Index           =   45
         Left            =   780
         TabIndex        =   133
         Top             =   1470
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Clase"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   285
         Index           =   44
         Left            =   480
         TabIndex        =   132
         Top             =   840
         Width           =   750
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   25
         Left            =   1485
         MouseIcon       =   "frmCCListados.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   1125
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   26
         Left            =   1485
         MouseIcon       =   "frmCCListados.frx":015E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   1515
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1470
         MouseIcon       =   "frmCCListados.frx":02B0
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar forfait"
         Top             =   3465
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
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
         Left            =   780
         TabIndex        =   128
         Top             =   3060
         Width           =   690
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
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
         Index           =   4
         Left            =   780
         TabIndex        =   127
         Top             =   3405
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Forfait"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   285
         Index           =   3
         Left            =   510
         TabIndex        =   126
         Top             =   2820
         Width           =   870
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1485
         MouseIcon       =   "frmCCListados.frx":0402
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar forfait"
         Top             =   3060
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Cargando Tabla Temporal"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   43
         Left            =   510
         TabIndex        =   113
         Top             =   6255
         Visible         =   0   'False
         Width           =   5385
      End
      Begin VB.Label Label4 
         Caption         =   "Cargando Tabla Temporal"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   42
         Left            =   510
         TabIndex        =   112
         Top             =   6015
         Visible         =   0   'False
         Width           =   5385
      End
      Begin VB.Label Label4 
         Caption         =   "Cargando Tabla Temporal"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   570
         TabIndex        =   22
         Top             =   5505
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Label Label1 
         Caption         =   "Informe de Costes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   495
         TabIndex        =   21
         Top             =   360
         Width           =   5655
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1485
         MouseIcon       =   "frmCCListados.frx":0554
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   2505
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1485
         MouseIcon       =   "frmCCListados.frx":06A6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   2115
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   285
         Index           =   2
         Left            =   480
         TabIndex        =   20
         Top             =   1830
         Width           =   1080
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
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
         Left            =   780
         TabIndex        =   19
         Top             =   2460
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
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
         Left            =   780
         TabIndex        =   18
         Top             =   2115
         Width           =   690
      End
      Begin VB.Label Label4 
         Caption         =   "Mes / Año"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   285
         Index           =   16
         Left            =   480
         TabIndex        =   15
         Top             =   3840
         Width           =   1110
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
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
         Index           =   15
         Left            =   780
         TabIndex        =   14
         Top             =   4155
         Width           =   690
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
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
         Index           =   14
         Left            =   780
         TabIndex        =   13
         Top             =   4515
         Width           =   645
      End
   End
   Begin VB.Frame FrameListadoCostes 
      Height          =   5340
      Left            =   0
      TabIndex        =   55
      Top             =   0
      Width           =   6435
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Tag             =   "Tipo Variedad|N|N|||variedades|tipovariedad||N|"
         Top             =   4380
         Width           =   1710
      End
      Begin VB.CommandButton CmdAcepInfCostes 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3900
         TabIndex        =   65
         Top             =   4605
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   5085
         TabIndex        =   66
         Top             =   4605
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   19
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   63
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3840
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   62
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3480
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   61
         Top             =   2775
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   60
         Top             =   2415
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   16
         Left            =   2715
         Locked          =   -1  'True
         TabIndex        =   68
         Text            =   "Text5"
         Top             =   2415
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   17
         Left            =   2715
         Locked          =   -1  'True
         TabIndex        =   67
         Text            =   "Text5"
         Top             =   2760
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   15
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   59
         Top             =   1740
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   14
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   58
         Top             =   1365
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   14
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   "Text5"
         Top             =   1380
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   15
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   56
         Text            =   "Text5"
         Top             =   1755
         Width           =   3135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Registro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   11
         Left            =   510
         TabIndex        =   79
         Top             =   4410
         Width           =   1065
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   19
         Left            =   1485
         Picture         =   "frmCCListados.frx":07F8
         ToolTipText     =   "Buscar fecha"
         Top             =   3855
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   18
         Left            =   1485
         Picture         =   "frmCCListados.frx":0883
         ToolTipText     =   "Buscar fecha"
         Top             =   3495
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   32
         Left            =   870
         TabIndex        =   78
         Top             =   3855
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   31
         Left            =   870
         TabIndex        =   77
         Top             =   3495
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   30
         Left            =   510
         TabIndex        =   76
         Top             =   3195
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   29
         Left            =   900
         TabIndex        =   75
         Top             =   2415
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   28
         Left            =   900
         TabIndex        =   74
         Top             =   2760
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   27
         Left            =   540
         TabIndex        =   73
         Top             =   2175
         Width           =   795
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   10
         Left            =   1485
         MouseIcon       =   "frmCCListados.frx":090E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   2775
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   1485
         MouseIcon       =   "frmCCListados.frx":0A60
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   2430
         Width           =   240
      End
      Begin VB.Label Label8 
         Caption         =   "Informe de Costes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   495
         TabIndex        =   72
         Top             =   450
         Width           =   5655
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   26
         Left            =   870
         TabIndex        =   71
         Top             =   1380
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   25
         Left            =   870
         TabIndex        =   70
         Top             =   1755
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Concepto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   24
         Left            =   540
         TabIndex        =   69
         Top             =   1065
         Width           =   690
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   1485
         MouseIcon       =   "frmCCListados.frx":0BB2
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar concepto"
         Top             =   1740
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   1485
         MouseIcon       =   "frmCCListados.frx":0D04
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar concepto"
         Top             =   1380
         Width           =   240
      End
   End
   Begin VB.Frame FrameCambios 
      Height          =   4170
      Left            =   0
      TabIndex        =   96
      Top             =   0
      Width           =   6435
      Begin VB.Frame Frame1 
         Caption         =   "Tipo de Operación"
         ForeColor       =   &H00972E0B&
         Height          =   1275
         Left            =   480
         TabIndex        =   103
         Top             =   930
         Width           =   5505
         Begin VB.OptionButton option1 
            Caption         =   "Introducción de Gerencia"
            Enabled         =   0   'False
            Height          =   255
            Index           =   5
            Left            =   3030
            TabIndex        =   107
            Top             =   930
            Visible         =   0   'False
            Width           =   2235
         End
         Begin VB.OptionButton option1 
            Caption         =   "Cambio Actividad Inexistente"
            Enabled         =   0   'False
            Height          =   255
            Index           =   4
            Left            =   330
            TabIndex        =   106
            Top             =   930
            Visible         =   0   'False
            Width           =   2475
         End
         Begin VB.OptionButton option1 
            Caption         =   "Descontar Almuerzo"
            Height          =   255
            Index           =   3
            Left            =   3030
            TabIndex        =   105
            Top             =   450
            Width           =   1935
         End
         Begin VB.OptionButton option1 
            Caption         =   "Hora Inicio según horario"
            Height          =   255
            Index           =   2
            Left            =   330
            TabIndex        =   104
            Top             =   450
            Width           =   2475
         End
      End
      Begin VB.CommandButton CmdAcepCambios 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3870
         TabIndex        =   98
         Top             =   3420
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   5
         Left            =   4980
         TabIndex        =   97
         Top             =   3420
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar pb4 
         Height          =   225
         Left            =   540
         TabIndex        =   99
         Top             =   2640
         Visible         =   0   'False
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label10 
         Caption         =   "Cambios Masivos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   495
         TabIndex        =   102
         Top             =   450
         Width           =   5655
      End
      Begin VB.Label Label4 
         Caption         =   "Cargando Tabla Temporal"
         Height          =   195
         Index           =   38
         Left            =   540
         TabIndex        =   101
         Top             =   2970
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.Label Label4 
         Caption         =   "Cargando Tabla Temporal"
         Height          =   195
         Index           =   37
         Left            =   540
         TabIndex        =   100
         Top             =   3210
         Visible         =   0   'False
         Width           =   5535
      End
   End
   Begin VB.Frame FrameBorradoMasivo 
      Height          =   3390
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Width           =   6435
      Begin VB.CommandButton CmdBorMasivo 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3900
         TabIndex        =   50
         Top             =   2655
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   5085
         TabIndex        =   51
         Top             =   2655
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   1350
         MaxLength       =   10
         TabIndex        =   49
         Tag             =   "Código Postal|T|S|||clientes|codposta|#0||"
         Top             =   1770
         Width           =   1050
      End
      Begin VB.Label Label4 
         Caption         =   "de 0 a 59"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   17
         Left            =   2730
         TabIndex        =   80
         Top             =   1830
         Width           =   1245
      End
      Begin VB.Label Label7 
         Caption         =   "Proceso que borra los ticajes no procesados, de tiempo inferior a los minutos indicados."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   555
         Left            =   420
         TabIndex        =   54
         Top             =   870
         Width           =   5655
      End
      Begin VB.Label Label4 
         Caption         =   "Minutos"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   13
         Left            =   630
         TabIndex        =   53
         Top             =   1830
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Borrado Masivo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   405
         TabIndex        =   52
         Top             =   450
         Width           =   5655
      End
   End
   Begin VB.Frame FrameModMasiva 
      Height          =   3630
      Left            =   0
      TabIndex        =   39
      Top             =   -60
      Width           =   6435
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   41
         Top             =   1800
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   46
         Text            =   "Text5"
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   40
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1365
         Width           =   1050
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   5085
         TabIndex        =   45
         Top             =   2835
         Width           =   975
      End
      Begin VB.CommandButton CmdModMasiva 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3900
         TabIndex        =   43
         Top             =   2835
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Línea"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   12
         Left            =   540
         TabIndex        =   47
         Top             =   1800
         Width           =   375
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1425
         MouseIcon       =   "frmCCListados.frx":0E56
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar línea coste"
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label5 
         Caption         =   "Modificación Masiva"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   495
         TabIndex        =   44
         Top             =   450
         Width           =   5655
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   9
         Left            =   540
         TabIndex        =   42
         Top             =   1350
         Width           =   615
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1425
         Picture         =   "frmCCListados.frx":0FA8
         ToolTipText     =   "Buscar fecha"
         Top             =   1365
         Width           =   240
      End
   End
   Begin VB.Frame FrameCargaOrdenConf 
      Height          =   3630
      Left            =   30
      TabIndex        =   23
      Top             =   -60
      Width           =   6435
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   21
         Left            =   4620
         MaxLength       =   10
         TabIndex        =   25
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1425
         Width           =   1050
      End
      Begin VB.CommandButton CmdAcepCargaOrd 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3900
         TabIndex        =   26
         Top             =   2835
         Width           =   975
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5085
         TabIndex        =   27
         Top             =   2835
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   24
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1425
         Width           =   1050
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   225
         Left            =   510
         TabIndex        =   31
         Top             =   2070
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Cargando Tabla Temporal"
         Height          =   195
         Index           =   41
         Left            =   570
         TabIndex        =   110
         Top             =   2310
         Visible         =   0   'False
         Width           =   5385
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   36
         Left            =   3720
         TabIndex        =   109
         Top             =   1440
         Width           =   525
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   39
         Left            =   780
         TabIndex        =   108
         Top             =   1440
         Width           =   555
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   21
         Left            =   4305
         Picture         =   "frmCCListados.frx":1033
         ToolTipText     =   "Buscar fecha"
         Top             =   1425
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   6
         Left            =   1395
         Picture         =   "frmCCListados.frx":10BE
         ToolTipText     =   "Buscar fecha"
         Top             =   1425
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   20
         Left            =   540
         TabIndex        =   30
         Top             =   1140
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Carga Automática Ordenes Confección"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   495
         TabIndex        =   29
         Top             =   450
         Width           =   5655
      End
      Begin VB.Label Label4 
         Caption         =   "Cargando Tabla Temporal"
         Height          =   195
         Index           =   10
         Left            =   570
         TabIndex        =   28
         Top             =   2550
         Visible         =   0   'False
         Width           =   5385
      End
   End
   Begin VB.Frame FrameCalculoImpReal 
      Height          =   4740
      Left            =   0
      TabIndex        =   137
      Top             =   0
      Width           =   6435
      Begin VB.CommandButton CmdAcepCalImpRea 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   144
         Top             =   4125
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   6
         Left            =   4905
         TabIndex        =   145
         Top             =   4125
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   31
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   140
         Tag             =   "Código Postal|T|S|||clientes|codposta|00||"
         Top             =   2250
         Width           =   690
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   35
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   142
         Tag             =   "Código Postal|T|S|||clientes|codposta|00||"
         Top             =   2610
         Width           =   690
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   32
         Left            =   2700
         MaxLength       =   10
         TabIndex        =   141
         Tag             =   "Código Postal|T|S|||clientes|codposta|0000||"
         Top             =   2250
         Width           =   780
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   36
         Left            =   2700
         MaxLength       =   10
         TabIndex        =   143
         Tag             =   "Código Postal|T|S|||clientes|codposta|0000||"
         Top             =   2610
         Width           =   780
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   28
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   147
         Text            =   "Text5"
         Top             =   1560
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   27
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   146
         Text            =   "Text5"
         Top             =   1215
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   28
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   139
         Top             =   1560
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   27
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   138
         Top             =   1200
         Width           =   830
      End
      Begin MSComctlLib.ProgressBar pb6 
         Height          =   225
         Left            =   510
         TabIndex        =   148
         Top             =   3210
         Visible         =   0   'False
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   63
         Left            =   870
         TabIndex        =   157
         Top             =   2595
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   62
         Left            =   870
         TabIndex        =   156
         Top             =   2235
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Mes / Año"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   61
         Left            =   480
         TabIndex        =   155
         Top             =   1920
         Width           =   885
      End
      Begin VB.Label Label11 
         Caption         =   "Cálculo de Importes Reales "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   495
         TabIndex        =   154
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label4 
         Caption         =   "Cargando Tabla Temporal"
         Height          =   195
         Index           =   56
         Left            =   540
         TabIndex        =   153
         Top             =   3480
         Visible         =   0   'False
         Width           =   5385
      End
      Begin VB.Label Label4 
         Caption         =   "Cargando Tabla Temporal"
         Height          =   195
         Index           =   55
         Left            =   540
         TabIndex        =   152
         Top             =   3720
         Visible         =   0   'False
         Width           =   5385
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1485
         MouseIcon       =   "frmCCListados.frx":1149
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1485
         MouseIcon       =   "frmCCListados.frx":129B
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   1215
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Concepto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   51
         Left            =   480
         TabIndex        =   151
         Top             =   930
         Width           =   690
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   50
         Left            =   870
         TabIndex        =   150
         Top             =   1560
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   49
         Left            =   870
         TabIndex        =   149
         Top             =   1215
         Width           =   465
      End
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   12
      Left            =   990
      MaxLength       =   4
      TabIndex        =   117
      Top             =   5655
      Width           =   830
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   13
      Left            =   990
      MaxLength       =   4
      TabIndex        =   116
      Top             =   6030
      Width           =   830
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   12
      Left            =   1875
      Locked          =   -1  'True
      TabIndex        =   115
      Text            =   "Text5"
      Top             =   5655
      Width           =   3135
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   13
      Left            =   1875
      Locked          =   -1  'True
      TabIndex        =   114
      Text            =   "Text5"
      Top             =   6030
      Width           =   3135
   End
   Begin VB.Frame FrameCargaFichajes 
      Height          =   3630
      Left            =   30
      TabIndex        =   32
      Top             =   0
      Width           =   6435
      Begin VB.CommandButton CmdLeerFicheros 
         Caption         =   "1. &Leer Ficheros"
         Height          =   585
         Left            =   1260
         TabIndex        =   33
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton CmdProcesarDia 
         Caption         =   "2. &Procesar Dia"
         Height          =   585
         Left            =   3150
         TabIndex        =   34
         Top             =   1305
         Width           =   1395
      End
      Begin MSComctlLib.ProgressBar pb2 
         Height          =   225
         Left            =   510
         TabIndex        =   35
         Top             =   2370
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Cargando Tabla Temporal"
         Height          =   195
         Index           =   7
         Left            =   510
         TabIndex        =   38
         Top             =   2640
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.Label Label4 
         Caption         =   "Cargando Tabla Temporal"
         Height          =   195
         Index           =   8
         Left            =   510
         TabIndex        =   37
         Top             =   2880
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.Label Label3 
         Caption         =   "Proceso de Carga de Fichajes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   495
         TabIndex        =   36
         Top             =   450
         Width           =   5655
      End
   End
   Begin VB.Frame FrameBusqueda 
      Height          =   4590
      Left            =   0
      TabIndex        =   81
      Top             =   0
      Width           =   6435
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   20
         Left            =   1470
         MaxLength       =   2
         TabIndex        =   85
         Top             =   2670
         Width           =   855
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   11
         Left            =   1470
         MaxLength       =   2
         TabIndex        =   84
         Top             =   2190
         Width           =   855
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   9
         Left            =   1470
         MaxLength       =   2
         TabIndex        =   83
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   23
         Left            =   1470
         MaxLength       =   5
         TabIndex        =   82
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   4
         Left            =   4980
         TabIndex        =   89
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepBusqueda 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3870
         TabIndex        =   87
         Top             =   3960
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar pb3 
         Height          =   225
         Left            =   540
         TabIndex        =   92
         Top             =   3180
         Visible         =   0   'False
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   35
         Left            =   570
         TabIndex        =   95
         Top             =   1740
         Width           =   285
      End
      Begin VB.Label Label4 
         Caption         =   "Cargando Tabla Temporal"
         Height          =   195
         Index           =   34
         Left            =   540
         TabIndex        =   94
         Top             =   3750
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.Label Label4 
         Caption         =   "Cargando Tabla Temporal"
         Height          =   195
         Index           =   33
         Left            =   540
         TabIndex        =   93
         Top             =   3510
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hora"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   19
         Left            =   570
         TabIndex        =   91
         Top             =   2685
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   18
         Left            =   570
         TabIndex        =   90
         Top             =   2205
         Width           =   225
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   40
         Left            =   570
         TabIndex        =   88
         Top             =   1215
         Width           =   525
      End
      Begin VB.Label Label9 
         Caption         =   "Búsqueda en Ficheros"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   495
         TabIndex        =   86
         Top             =   450
         Width           =   5655
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   5040
      Top             =   4650
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameStockMaxMin 
      Caption         =   "Agrupado por"
      ForeColor       =   &H00972E0B&
      Height          =   1095
      Left            =   0
      TabIndex        =   121
      Top             =   0
      Width           =   2205
      Begin VB.OptionButton option1 
         Caption         =   "Variedad / Fecha"
         Height          =   255
         Index           =   1
         Left            =   330
         TabIndex        =   123
         Top             =   660
         Width           =   1785
      End
      Begin VB.OptionButton option1 
         Caption         =   "Fecha / Variedad"
         Height          =   255
         Index           =   0
         Left            =   330
         TabIndex        =   122
         Top             =   315
         Width           =   1815
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Desde"
      Height          =   195
      Index           =   23
      Left            =   330
      TabIndex        =   120
      Top             =   315
      Width           =   465
   End
   Begin VB.Label Label4 
      Caption         =   "Hasta"
      Height          =   195
      Index           =   22
      Left            =   330
      TabIndex        =   119
      Top             =   690
      Width           =   420
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Concepto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   195
      Index           =   21
      Left            =   0
      TabIndex        =   118
      Top             =   0
      Width           =   690
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   12
      Left            =   945
      MouseIcon       =   "frmCCListados.frx":13ED
      MousePointer    =   4  'Icon
      ToolTipText     =   "Buscar concepto"
      Top             =   315
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   13
      Left            =   945
      MouseIcon       =   "frmCCListados.frx":153F
      MousePointer    =   4  'Icon
      ToolTipText     =   "Buscar concepto"
      Top             =   705
      Width           =   240
   End
End
Attribute VB_Name = "frmCCListados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MANOLO +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public Opcionlistado As Integer
' 0 = Listado de costes
' 1 = Carga masiva de ordenes de confeccion dada una fecha
' 2 = Proceso de fichajes del reloj

' 3 = Modificacion masiva de cctrabaconf
' 4 = borrado masivo de cctrabaconf (menos de x minutos)

' 5 = informe de ticajes de cctrabconf
' 6 = buscar en ficheros una cadena (solo para root)
' 7 = Procesos Varios de cambios en ticajes
      ' cambio de hora de inicio de tarea si es inferior a 07:00 ponemos 07:00
      ' cambio de ticajes, eliminamos la hora del almuerzo (siempre de 09:00 a 09:30)
      
' 8 = Proceso de carga de importes reales de la contabilidad

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public CadBusqueda As String

Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmFor As frmManForfaits 'forfaits
Attribute frmFor.VB_VarHelpID = -1
Private WithEvents frmVar As frmManVariedad 'Variedad
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmCon As frmCCManConcep 'Conceptos
Attribute frmCon.VB_VarHelpID = -1
Private WithEvents frmInc As frmManInciden 'Incidencias
Attribute frmInc.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmBas  As frmBasico ' Lineas de confeccion
Attribute frmBas.VB_VarHelpID = -1
Private WithEvents frmTra  As frmBasico ' trabajadores
Attribute frmTra.VB_VarHelpID = -1
Private WithEvents frmMensVariedad As frmMensajes 'mensajes
Attribute frmMensVariedad.VB_VarHelpID = -1
Private WithEvents frmMensForfaits As frmMensajes 'mensajes
Attribute frmMensForfaits.VB_VarHelpID = -1
Private WithEvents frmCla As frmManClases  'Clase
Attribute frmCla.VB_VarHelpID = -1


'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadselect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim CadConcepto As String
Dim CadVariedad As String
Dim CadForfait As String


Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim Tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean
Dim Continuar As Boolean

Dim FDesde As Date
Dim FHasta As Date

Dim Variedades As String
Dim Forfaits As String


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub CmdAcepBusqueda_Click()
Dim Directorio As String
Dim SQL As String

    InicializarVbles
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    Directorio = vParamAplic.PathFichadas & "\."
    If Dir(Directorio, vbArchive) <> "" Then
        Directorio = vParamAplic.PathFichadas
        
        If ProcesarDirectorioBusqueda(Directorio & "\", pb3, Label4(33), Label4(34)) Then
            If HayRegParaInforme("tmpinformes", "{tmpinformes.codusu} = " & vUsu.Codigo) Then
                cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
                cadNombreRPT = "rCCInfBusqueda.rpt"
                cadTitulo = "Informe de Busqueda"
                LlamarImprimir
            End If
        End If
    Else
        MsgBox "No existe el directorio de ticajes. Llame a Ariadna.", vbExclamation
        Continuar = False
    End If
    

End Sub

Private Sub CmdAcepCalImpRea_Click()
Dim SQL As String
Dim CADENA As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta

    InicializarVbles

    If Not DatosOk Then Exit Sub

    cadselect = "select count(*) from ccconcostes_mes where date(concat(right(concat('0000',año),4), '-', right(concat('00',mes),2), '-01')) between " & DBSet(CStr(FDesde), "F") & " and " & DBSet(CStr(FHasta), "F")
    
'    If txtCodigo(27).Text <> "" Then Sql = Sql & " and codcoste >= " & DBSet(txtCodigo(27), "N")
'    If txtCodigo(28).Text <> "" Then Sql = Sql & " and codcoste <= " & DBSet(txtCodigo(28), "N")
'
    
    cDesde = Trim(txtcodigo(27).Text)
    cHasta = Trim(txtcodigo(28).Text)
    nDesde = txtNombre(27).Text
    nHasta = txtNombre(28).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "ccconcostes_mes.codcoste"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHProducto= """) Then Exit Sub
    End If
    
    If TotalRegistros(cadselect) = 0 Then
        MsgBox "No hay conceptos entre ese rango de fechas para calcular el importe real.", vbExclamation
    
    Else
        If CalculoImporteReal(cadselect) Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
            cmdCancel_Click (6)
        End If
    End If
End Sub


Private Function CuentasCoste(Coste As String) As String
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim cadResul As String

    On Error Resume Next

    CuentasCoste = ""
    
    SQL = "select ctacontable from ccconcostes_cta where codcoste = " & DBSet(Coste, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cadResul = ""
    While Not Rs.EOF
        cadResul = cadResul & DBSet(Rs.Fields(0).Value, "T") & ","
    
        Rs.MoveNext
    Wend
    Set Rs = Nothing

    ' le quitamos la ultima coma
    If cadResul <> "" Then
        cadResul = Mid(cadResul, 1, Len(cadResul) - 1)
        CuentasCoste = cadResul
    End If
        
End Function


Private Function CalculoImporteReal(vCadena As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Importe As Currency
Dim Rs As ADODB.Recordset
Dim Ctas As String
Dim nRegs As Integer

    On Error GoTo eCalculoImporteReal

    CalculoImporteReal = False
    
    conn.BeginTrans
    
    nRegs = DevuelveValor(vCadena)
    Me.pb6.visible = True
    Label4(56).visible = True
    CargarProgres pb6, nRegs
    DoEvents
    
    'Sql = "select * from ccconcostes_mes where date(concat(right(concat('0000',año),4), " - ", right(concat('00',mes),2), '-01')) between " & DBSet(FDesde, "F") & " and " & DBSet(FHasta, "F")
    'Sql = Sql & " and codcoste >= " & DBSet(txtCodigo(27), "N")
    'Sql = Sql & " and codcoste <= " & DBSet(txtCodigo(28), "N")
    
    vCadena = Replace(vCadena, "count(*)", "*")
    
    Set Rs = New ADODB.Recordset
    Rs.Open vCadena, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        IncrementarProgres Me.pb6, 1
        DoEvents
    
        Ctas = CuentasCoste(DBLet(Rs!codCoste, "N"))
        
        Sql2 = "select sum(if(impmesde is null,0,impmesde)) - sum(if(impmesha is null, 0,impmesha)) as importereal from hsaldos "
        Sql2 = Sql2 & " where codmacta in (" & Ctas & ")"
        Sql2 = Sql2 & " and anopsald = " & DBSet(Rs!año, "N") & " and mespsald = " & DBSet(Rs!mes, "N")
        
        Importe = DevuelveValorConta(Sql2)
        
        If Importe = 0 Then
            'miramos si está en el hco de saldos
            Sql2 = "select sum(if(impmesde is null,0,impmesde)) - sum(if(impmesha is null, 0,impmesha)) as importereal from hsaldos1 "
            Sql2 = Sql2 & " where codmacta in (" & Ctas & ")"
            Sql2 = Sql2 & " and anopsald = " & DBSet(Rs!año, "N") & " and mespsald = " & DBSet(Rs!mes, "N")
                
            Importe = DevuelveValorConta(Sql2)
        End If
    
        Sql2 = "update ccconcostes_mes set importereal = " & DBSet(Importe, "N")
        Sql2 = Sql2 & " where codcoste = " & DBSet(Rs!codCoste, "N")
        Sql2 = Sql2 & " and año = " & DBSet(Rs!año, "N") & " and mes = " & DBSet(Rs!mes, "N")
        
        conn.Execute Sql2
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    conn.CommitTrans
    
    CalculoImporteReal = True
    Me.pb6.visible = False
    Label4(56).visible = False
    
    
    Exit Function
    
eCalculoImporteReal:
    If Err.Number <> 0 Then
        Me.pb6.visible = False
        Label4(56).visible = False
        conn.RollbackTrans
    End If
End Function


Private Sub CmdAcepCambios_Click()
Dim Mens As String
Dim nRegs As Long
Dim SQL As String

    'Modificacion de la hora de inicio segun el horario del trabajador

    'Cambio de hora de inicio, segun horario
    If Me.option1(2).Value Then
        SQL = "select cctrabaconf.* from cctrabaconf, straba, cchorario, cchorario_tramo "
        SQL = SQL & " where procesado = 0 and sinacabar = 0 "
        SQL = SQL & " and cctrabaconf.codtraba = straba.codtraba "
        SQL = SQL & " and straba.codhorario = cchorario.codhorario "
        SQL = SQL & " and cchorario.codhorario = cchorario_tramo.codhorario "
        SQL = SQL & " and time(cctrabaconf.fechaini) between time(cchorario_tramo.fechaini) and time(cchorario_tramo.fechafin) "
        SQL = SQL & " and time(cctrabaconf.fechaini) <> time(cchorario_tramo.fechares) "
        'sql = sql & " where procesado = 0 and sinacabar = 0 and time(fechaini) < '07:00:00'"
        nRegs = TotalRegistrosConsulta(SQL)
        If nRegs <> 0 Then
            Mens = "Se va a iniciar el proceso de ajuste de hora de inicio según horario " & vbCrLf & "en " & nRegs & " fichadas. " & vbCrLf & vbCrLf & "¿ Desesa continuar ? "
            If MsgBox(Mens, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                If CambioHoraInicio(nRegs) Then
                    cmdCancel_Click (5)
                End If
            End If
        Else
            MsgBox "No hay registros a modificar.", vbExclamation
        End If
    End If
    
    'Descontar Hora de Almuerzo
    If Me.option1(3).Value Then
        SQL = "select * from cctrabaconf where procesado = 0 and sinacabar = 0 and time(fechaini) < '09:00:00' and time(fechafin) > '09:30:00'"
        SQL = SQL & " and abs(TIME_TO_SEC(timediff('09:00:00',time(fechaini)))) > 1800 "
        nRegs = TotalRegistrosConsulta(SQL)
        If nRegs <> 0 Then
            Mens = "Se va a iniciar el proceso de ajuste de tiempo de almuerzo de 9:00 a 9:30" & vbCrLf & "horas en " & nRegs & " fichadas." & vbCrLf & vbCrLf & " ¿ Desesa continuar ? "
            If MsgBox(Mens, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
               If CambioTiempoAlmuerzo(nRegs) Then
                   MsgBox "Proceso realizado correctamente.", vbExclamation
                   cmdCancel_Click (5)
               End If
            End If
        Else
            MsgBox "No hay registros a modificar.", vbExclamation
        End If
    End If
    
    'Cambio de actividad inexistente por la de defecto del trabajador
    If Me.option1(4).Value Then
        MsgBox "En proceso de contruccion", vbExclamation
    
    End If
    
    'Introduccion de fichada de gerencia
    If Me.option1(5).Value Then
        MsgBox "En proceso de contruccion", vbExclamation
    
    End If

End Sub

Private Function CambioHoraInicio(nRegs As Long) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim campo As String

    On Error GoTo eCambioHoraInicio


    Screen.MousePointer = vbHourglass

    CambioHoraInicio = False

    '------------------------------------------------------------------------------
    '  LOG de acciones.
    Set LOG = New cLOG
    campo = "Control Costes: Cambio de Hora de Inicio a " & nRegs & " modificados."
    LOG.Insertar 7, vUsu, campo
    Set LOG = Nothing
    '-----------------------------------------------------------------------------
    
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    SQL = "select cctrabaconf.* from cctrabaconf, straba, cchorario, cchorario_tramo "
    SQL = SQL & " where procesado = 0 and sinacabar = 0 "
    SQL = SQL & " and cctrabaconf.codtraba = straba.codtraba "
    SQL = SQL & " and straba.codhorario = cchorario.codhorario "
    SQL = SQL & " and cchorario.codhorario = cchorario_tramo.codhorario "
    SQL = SQL & " and time(cctrabaconf.fechaini) between time(cchorario_tramo.fechaini) and time(cchorario_tramo.fechafin) "
    SQL = SQL & " and time(cctrabaconf.fechaini) <> time(cchorario_tramo.fechares) "
    
    
    Sql2 = "insert into tmpinformes (codusu, codigo1, fecha1, importe1, importe2) "
    Sql2 = Sql2 & " select " & vUsu.Codigo & ", cctrabaconf.codtraba, date(cctrabaconf.fechaini), hour(cctrabaconf.fechaini), minute(cctrabaconf.fechaini) "
    Sql2 = Sql2 & " from cctrabaconf, straba, cchorario, cchorario_tramo "
    Sql2 = Sql2 & " where procesado = 0 and sinacabar = 0 "
    Sql2 = Sql2 & " and cctrabaconf.codtraba = straba.codtraba "
    Sql2 = Sql2 & " and straba.codhorario = cchorario.codhorario "
    Sql2 = Sql2 & " and cchorario.codhorario = cchorario_tramo.codhorario "
    Sql2 = Sql2 & " and time(cctrabaconf.fechaini) between time(cchorario_tramo.fechaini) and time(cchorario_tramo.fechafin) "
    Sql2 = Sql2 & " and time(cctrabaconf.fechaini) <> time(cchorario_tramo.fechares) "
        
    
    conn.Execute Sql2
    
'    Sql2 = "update cctrabaconf set fechaini = concat(date(fechaini),  ' 07:00:00')"
'    Sql2 = Sql2 & " where procesado = 0 and sinacabar = 0 and time(fechaini) < '07:00:00'  "
    
    ' actualizamos con la fecha resultante
    Sql2 = "update cctrabaconf, straba, cchorario, cchorario_tramo set cctrabaconf.fechaini = concat(date(cctrabaconf.fechaini), ' ', time(cchorario_tramo.fechares)) "
    Sql2 = Sql2 & " where procesado = 0 and sinacabar = 0 "
    Sql2 = Sql2 & " and cctrabaconf.codtraba = straba.codtraba "
    Sql2 = Sql2 & " and straba.codhorario = cchorario.codhorario "
    Sql2 = Sql2 & " and cchorario.codhorario = cchorario_tramo.codhorario "
    Sql2 = Sql2 & " and time(cctrabaconf.fechaini) between time(cchorario_tramo.fechaini) and time(cchorario_tramo.fechafin) "
    Sql2 = Sql2 & " and time(cctrabaconf.fechaini) <> time(cchorario_tramo.fechares) "
    
    conn.Execute Sql2
    
    MsgBox "Se han modificado " & nRegs & " registros. ", vbExclamation
    
    CambioHoraInicio = True
    Screen.MousePointer = vbDefault
    
    Exit Function
    
eCambioHoraInicio:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Cambio Hora Inicio", Err.Description
End Function


Private Function CambioTiempoAlmuerzo(nRegs As Long) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim campo As String
Dim Rs As ADODB.Recordset
Dim CadValues As String
Dim FecIni As String
Dim FecFin As String

    On Error GoTo eCambioTiempoAlmuerzo

    Screen.MousePointer = vbHourglass

    CambioTiempoAlmuerzo = False

    conn.BeginTrans

    '------------------------------------------------------------------------------
    '  LOG de acciones.
    Set LOG = New cLOG
    campo = "Control Costes: Cambio tiempo de almuerzo de 9:00 a 9:30, " & nRegs & " modificados."
    LOG.Insertar 8, vUsu, campo
    Set LOG = Nothing
    '-----------------------------------------------------------------------------
    
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    pb4.visible = True
    CargarProgres pb4, CInt(nRegs)
    DoEvents
    
    ' se cambian aquellos registros que entran antes de 1/2h de las 09:00
    SQL = "select * from cctrabaconf where procesado = 0 and sinacabar = 0 and time(fechaini) < '09:00:00' and time(fechafin) > '09:30:00' "
    SQL = SQL & " and abs(TIME_TO_SEC(timediff('09:00:00',time(fechaini)))) > 1800 "
    
    Sql2 = "insert into cctrabaconf (codtraba,fechaini,fechafin,codlinconf,codcoste,procesado,sinacabar) values "
    CadValues = ""
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        IncrementarProgres pb4, 1
    
        ' insertamos el registro nuevo fechaini = 09:30:00 y fechafin la que el cliente tenga en el registro
        FecIni = Format(DBLet(Rs!FechaIni, "F"), "yyyy-mm-dd") & " " & "09:30:00"
        FecFin = Format(DBLet(Rs!FechaIni, "F"), "yyyy-mm-dd") & " " & "09:00:00"
        
        CadValues = CadValues & "(" & DBSet(Rs!codtraba, "N") & "," & DBSet(FecIni, "FH") & "," & DBSet(Rs!FechaFin, "FH") & ","
        CadValues = CadValues & DBSet(Rs!codlinconf, "N") & "," & DBSet(Rs!codCoste, "N") & ",0,0),"
        
        Sql2 = "insert into tmpinformes (codusu, codigo1, fecha1) "
        Sql2 = Sql2 & " select " & vUsu.Codigo & ", codtraba, date(fechaini) from cctrabaconf  "
        Sql2 = Sql2 & " where codtraba = " & DBSet(Rs!codtraba, "N")
        Sql2 = Sql2 & " and fechaini = " & DBSet(Rs!FechaIni, "FH") & " and fechafin = " & DBSet(Rs!FechaFin, "FH")

        conn.Execute Sql2
        
        Sql3 = "update cctrabaconf set fechafin = " & DBSet(FecFin, "FH") & " where codtraba = " & DBSet(Rs!codtraba, "N")
        Sql3 = Sql3 & " and fechaini = " & DBSet(Rs!FechaIni, "FH") & " and fechafin = " & DBSet(Rs!FechaFin, "FH")
        conn.Execute Sql3
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    If CadValues <> "" Then
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
        Sql2 = "insert into cctrabaconf (codtraba,fechaini,fechafin,codlinconf,codcoste,procesado,sinacabar) values "
        conn.Execute Sql2 & CadValues
    End If
    
    conn.CommitTrans
    
    MsgBox "Se han modificado " & nRegs & " registros. ", vbExclamation
    
    CambioTiempoAlmuerzo = True
    Screen.MousePointer = vbDefault
    pb4.visible = False
    DoEvents
    
    Exit Function
    
eCambioTiempoAlmuerzo:
    conn.RollbackTrans
    pb4.visible = False
    DoEvents
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Cambio Hora Inicio", Err.Description
End Function




Private Sub CmdAcepCargaOrd_Click()
Dim SQL As String

    InicializarVbles

'    If txtCodigo(6).Text = "" Then
'        MsgBox "Debe introducir obligatoriamente una Fecha.", vbExclamation
'        PonerFoco txtCodigo(6)
'        Exit Sub
'    End If
    
    SQL = "select count(*) from cctrabaconf where date(fechaini) = " & DBSet(txtcodigo(6).Text, "F")
    
    If TotalRegistros(SQL) = 0 Then
        MsgBox "No hay fichajes en esta fecha para construir ordenes de confección.", vbExclamation
    
    Else
        '[Monica]09/07/2012: comprobamos que todos los registros a procesar se han acabado
        '                    si hay registros con la tarea sin acabar damos un aviso para que los revisen
        SQL = "select count(*) from cctrabaconf where date(fechaini) = " & DBSet(txtcodigo(6).Text, "F") & " and sinacabar = 1"
        If TotalRegistros(SQL) Then
            MsgBox "Hay registros de trabajador que no se han acabado. Revise.", vbExclamation
            PonerFoco txtcodigo(6)
            Exit Sub
        End If
    
        ' Comprobamos las claves referenciales del fichero que cargan que no las tiene en la base de datos para que no dé errores
        ' al cargar la tabla
        If ComprobarFicheroNew(txtcodigo(6).Text) Then
            If TotalRegistrosConsulta("select * from tmpinformes where codusu = " & vUsu.Codigo) <> 0 Then
                cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
                cadTitulo = "Errores en el Traspaso de Costes"
                cadNombreRPT = "rErroresTrasCostes.rpt"
                LlamarImprimir
            Else
                If MsgBox("Atención este proceso borra las tablas de cálculo antes de cargarlas. ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                    If ProcesoCargaOrdenesNew(txtcodigo(6).Text, txtcodigo(21).Text, pb1, Label4(41), Label4(10)) Then
                        MsgBox "Proceso realizado correctamente", vbExclamation
                        cmdCancelar_Click
                    End If
                Else
                    Unload Me
                End If
            End If
        End If
    End If

End Sub

Private Function ComprobarFichero(fecha As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim CadM As String

    On Error GoTo eComprobarFichero

    ComprobarFichero = False

    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL


    SQL = "select * from cctrabaconf where sinacabar = 0 and (date(fechaini) = " & DBSet(fecha, "F") & " or date(fechafin) = " & DBSet(fecha, "F") & ")"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    While Not Rs.EOF
        ' comprobamos que exista el trabajador
        SQL = DevuelveDesdeBDNew(cAgro, "straba", "codtraba", "codtraba", CStr(DBLet(Rs!codtraba, "N")), "N")
        If SQL = "" Then
            CadM = "No existe el trabajador"
                        
            Sql2 = "insert into tmpinformes (codusu,codigo1,nombre1) values (" & vUsu.Codigo & "," & DBSet(Rs!codtraba, "N") & ","
            Sql2 = Sql2 & DBSet(CadM, "T") & ")"
            
            conn.Execute Sql2
        End If
        
        ' comprobamos que exista la linea de coste de confeccion
        SQL = DevuelveDesdeBDNew(cAgro, "cclinconf", "codlinconf", "codlinconf", CStr(DBLet(Rs!codlinconf, "N")), "N")
        If SQL = "" Then
            CadM = "No existe Linea de Coste"

            Sql2 = "insert into tmpinformes (codusu,codigo1,nombre1) values (" & vUsu.Codigo & "," & DBSet(Rs!codlinconf, "N") & ","
            Sql2 = Sql2 & DBSet(CadM, "T") & ")"

            conn.Execute Sql2
        End If
        
        ' comprobamos que exista el concepto de coste
        SQL = DevuelveDesdeBDNew(cAgro, "ccconcostes", "codcoste", "codcoste", CStr(DBLet(Rs!codCoste, "N")), "N")
        If SQL = "" Then
            CadM = "No existe el Concepto"
                        
            Sql2 = "insert into tmpinformes (codusu,codigo1,nombre1) values (" & vUsu.Codigo & "," & DBSet(Rs!codCoste, "N") & ","
            Sql2 = Sql2 & DBSet(CadM, "T") & ")"
            
            conn.Execute Sql2
        End If
        
        '[Monica]28/05/2012: hay fichajes que son de distinto dia de inicio de tarea que de fin
        If Mid(DBLet(Rs!FechaIni), 1, 10) <> Mid(DBLet(Rs!FechaFin), 1, 10) Then
            CadM = "Coste:" & DBLet(Rs!codCoste, "N") & " F.I:" & Mid(DBLet(Rs!FechaIni, "F"), 1, 10) & "-F.F:" & Mid(DBLet(Rs!FechaFin, "F"), 1, 10)
                        
            Sql2 = "insert into tmpinformes (codusu,codigo1,nombre1) values (" & vUsu.Codigo & "," & DBSet(Rs!codtraba, "N") & ","
            Sql2 = Sql2 & DBSet(CadM, "T") & ")"
            
            conn.Execute Sql2
        
        End If
        
        
        Rs.MoveNext
    Wend

    Set Rs = Nothing
    
    ComprobarFichero = True
    Exit Function
    
eComprobarFichero:
    MuestraError Err.Number, "Comprobar Fichero", Err.Description
End Function


Private Function ProcesoCargaOrdenes() As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Sql4 As String

Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim nRegs As Long
Dim CodOrden As Long
Dim Cliente As Long
Dim NumLin As Long

Dim b As Boolean
Dim codTipoM As String
Dim vTipoMov As CTiposMov
Dim devuelve As String
Dim Existe As Boolean
Dim CadValues As String
Dim Mens As String


    On Error GoTo eProcesoCargaOrdenes


    ProcesoCargaOrdenes = False
    
    SQL = "DROP TABLE IF EXISTS tmpcclindia1; "
    conn.Execute SQL
    
    SQL = "CREATE TEMPORARY TABLE `tmpcclindia1` ("
    SQL = SQL & "`fecha` datetime NOT NULL, "
    SQL = SQL & "`codcoste` int(4) unsigned NOT NULL, "
    SQL = SQL & "`numlinea` smallint(4) unsigned NOT NULL,"
    SQL = SQL & "`codtraba` int(6) NOT NULL, "
    SQL = SQL & "`fechaini` datetime NOT NULL, "
    SQL = SQL & "`fechafin` datetime NOT NULL,"
    SQL = SQL & "`horas` decimal(4,2) NOT NULL,"
    SQL = SQL & " PRIMARY KEY  (`fecha`,`codcoste`,`numlinea`),"
    SQL = SQL & " KEY `FK_tmpcclindia1` (`codtraba`)"
    SQL = SQL & ") ENGINE=InnoDB DEFAULT CHARSET=latin1"
    
    conn.Execute SQL
    
    conn.Execute "delete from tmpinformes where " & vUsu.Codigo

    conn.BeginTrans
    
    SQL = "select distinct numpedid from palets where date(fechafin) = " & DBSet(txtcodigo(6).Text, "F") & " and intorden = 0"
    SQL = SQL & " order by 1 "
    
    nRegs = TotalRegistrosConsulta(SQL)
    
    Label4(10).visible = True
    pb1.visible = True
    Label4(10).Caption = "Cargando ordenes de confección"
    DoEvents
    
    If nRegs = 0 Then
        b = False
    Else
        pb1.Max = nRegs
        pb1.Value = 0
    End If
    
    codTipoM = "ORD"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    b = True
    While Not Rs.EOF And b
        IncrementarProgresNew pb1, 1
        
        Label4(10).Caption = "Procesando Palets del Pedido " & DBLet(Rs!numpedid, "N")
        DoEvents
        
        ' Insertamos la cabecera de la orden de confeccion
        Set vTipoMov = New CTiposMov
        If vTipoMov.Leer(codTipoM) Then
            'Comprobar si mientras tanto se incremento el contador de ordenes
            Do
                CodOrden = vTipoMov.ConseguirContador(codTipoM)
                devuelve = DevuelveDesdeBDNew(cAgro, "cccaborden", "codorden", "codorden", CStr(CodOrden), "N")
                If devuelve <> "" Then
                    'Ya existe el contador incrementarlo
                    Existe = True
                    vTipoMov.IncrementarContador (codTipoM)
                    CodOrden = vTipoMov.ConseguirContador(codTipoM)
                Else
                    Existe = False
                End If
            Loop Until Not Existe
                
        Else 'No existe el tipo de Movimiento
            b = False
            Set vTipoMov = Nothing
        End If
        
        ' Insertamos la cabecera CCABORDEN
        If b Then
            Mens = "Insertar Cabecera Orden de Confección"
            b = InsertarCabeceraOrden(Rs, CStr(CodOrden), Mens)
        End If
            
        ' Insertamos lineas CCLINORDEN3
        If b Then
            Mens = "Insertar Líneas Variedad Orden de Confección"
            b = InsertarLineasVariedadOrden(Rs, CStr(CodOrden), Mens)
        End If
            
        ' Insertamos lineas CCLINORDEN4
        If b Then
            Mens = "Insertar Líneas Variedad Resumen Orden de Confección"
            b = InsertarLineasResumenVariedadOrden(Rs, CStr(CodOrden), Mens)
        End If

        ' Insertamos lineas CCLINORDEN5
        If b Then
            Mens = "Insertar palets de Orden de Confección"
            b = InsertarPaletsOrden(Rs, CStr(CodOrden), Mens)
        End If
        
        ' Incrementamos el contador
        If b Then
            vTipoMov.IncrementarContador (codTipoM)
        End If
        Set vTipoMov = Nothing
        
        Rs.MoveNext
    Wend
            
    If b Then
        Mens = "Insertar Trabajadores en las Ordenes de Confeccion"
        b = InsertarTrabajadoresOrden(Mens)
        If b Then
            Mens = "Modificando Categorias"
            b = ModificarCategorias(Mens)
            
            ' en caso de que se haya insertado algo en la temporal tenemos que volcarlo a los costes diarios
            ' cccabdia cclindia1 y cclindia2
            If b Then
                Mens = "Insertando Costes Diarios"
                b = InsertarCostesDiarios(Mens)
            End If
            
        End If
    End If
    
    
    ' actualizamos los palets para que no se vuelvan a procesar
    If b Then
        SQL = "update palets set intorden = 1 where date(horafin) = " & DBSet(txtcodigo(6).Text, "F") & " and intorden = 0"
        conn.Execute SQL
    End If

    Set Rs = Nothing
    
eProcesoCargaOrdenes:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Proceso Carga Ordenes", Err.Description
        b = False
    Else
        If Not b Then
            MsgBox "Proceso Carga Ordenes: " & vbCrLf & vbCrLf & Mens, vbExclamation
        End If
    End If
    
    pb1.visible = False
    Label4(10).visible = False
    
    If b Then
        ProcesoCargaOrdenes = True
        conn.CommitTrans
    Else
        ProcesoCargaOrdenes = False
        conn.RollbackTrans
    End If
End Function

Private Function InsertarCostesDiarios(Mens As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String

Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset

Dim NumF As Long
Dim CadValues As String
Dim Categoria As Long
Dim Horas As Currency



    On Error GoTo eInsertarCostesDiarios
    
    InsertarCostesDiarios = False
    
    SQL = "select count(*) from tmpcclindia1 "
    If TotalRegistros(SQL) = 0 Then
        InsertarCostesDiarios = True
        Exit Function
    End If
    
    SQL = "insert ignore into cccabdia (fecha, codcoste, observac) "
    SQL = SQL & " select fecha, codcoste, 'Carga Ordenes de Confección " & txtcodigo(6).Text & "'"
    SQL = SQL & " from tmpcclindia1 "
    
    conn.Execute SQL
    
    
    ' Necesito otro cursor pq tengo que renumerar el numero de linea por si ya existia la cabecera
    SQL = "select distinct fecha, codcoste from tmpcclindia1 order by 1,2"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        NumF = SugerirCodigoSiguienteStr("cclindia1", "numlinea", "fecha=" & DBSet(Rs!fecha, "F") & " and codcoste = " & DBSet(Rs!codCoste, "N"))
    
        Sql3 = "insert into cclindia1 (fecha,codcoste,numlinea,codtraba,fechaini,fechafin,horas) values "
    
        CadValues = ""
    
        Sql2 = "select * from tmpcclindia1 where fecha = " & DBSet(Rs!fecha, "F") & " and codcoste = " & DBSet(Rs!codCoste, "N")
        Sql2 = Sql2 & " order by codtraba, fechaini, fechafin "
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs2.EOF
            CadValues = CadValues & "(" & DBSet(Rs!fecha, "F") & "," & DBSet(Rs!codCoste, "N") & ","
            CadValues = CadValues & DBSet(NumF, "N") & "," & DBSet(Rs2!codtraba, "N") & ","
            CadValues = CadValues & DBSet(Rs2!FechaIni, "FH") & "," & DBSet(Rs2!FechaFin, "FH") & ","
            CadValues = CadValues & DBSet(Rs2!Horas, "N") & "),"
            
            NumF = NumF + 1
            
            Rs2.MoveNext
        Wend
        Set Rs2 = Nothing
        
        If CadValues <> "" Then
            CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
            conn.Execute Sql3 & CadValues
        End If
        
        'recalculamos la tabla de categorias cclindia2
        Sql2 = "select * from cclindia1 where fecha = " & DBSet(Rs!fecha, "F") & " and codcoste = " & DBSet(Rs!codCoste, "N")
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs2.EOF
            SQL = "select codcateg from straba where codtraba = " & DBSet(Rs2!codtraba, "N")
            Categoria = DevuelveValor(SQL)
        
            SQL = "select count(*) from cclindia2 where fecha = " & DBSet(Rs!fecha, "F")
            SQL = SQL & " and codcoste = " & DBSet(Rs!codCoste, "N")
            SQL = SQL & " and codcateg = " & DBSet(Categoria, "N")
            
            If TotalRegistros(SQL) = 0 Then
                NumF = SugerirCodigoSiguienteStr("cclindia2", "numlinea", "fecha = " & DBSet(Rs!fecha, "F") & " and codcoste = " & DBSet(Rs!codCoste, "N"))
            
                Sql2 = "insert into cclindia2 (fecha,codcoste,numlinea,codcateg,horas) values ("
                Sql2 = Sql2 & DBSet(Rs!fecha, "F") & "," & DBSet(Rs!codCoste, "N") & "," & DBSet(NumF, "N") & ","
                Sql2 = Sql2 & DBSet(Categoria, "N") & "," & DBSet(Rs2!Horas, "N") & ")"
                
                conn.Execute Sql2
            
            Else
                Sql2 = "select sum(horas) from cclindia1 where fecha = " & DBSet(Rs!fecha, "F")
                Sql2 = Sql2 & " and codcoste = " & DBSet(Rs!codCoste, "N")
                Sql2 = Sql2 & " and codtraba in ( select codtraba from straba where codcateg = " & DBSet(Categoria, "N") & ")"
                
                Horas = DevuelveValor(Sql2)
            
                Sql2 = "update cclindia2 set horas = " & DBSet(Horas, "N")
                Sql2 = Sql2 & " where fecha = " & DBSet(Rs!fecha, "F")
                Sql2 = Sql2 & " and codcoste = " & DBSet(Rs!codCoste, "N")
                Sql2 = Sql2 & " and codcateg = " & DBSet(Categoria, "N")
                    
                conn.Execute Sql2
                    
                Sql2 = "delete from cclindia2 where fecha = " & DBSet(Rs!fecha, "F")
                Sql2 = Sql2 & " and codcoste = " & DBSet(Rs!codCoste, "N")
                Sql2 = Sql2 & " and horas = 0"
                
                conn.Execute Sql2
            End If
            
            Rs2.MoveNext
        Wend
        
        Set Rs2 = Nothing
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    
    InsertarCostesDiarios = True
    Exit Function
    
eInsertarCostesDiarios:
    Mens = Mens & vbCrLf & "Insertar Costes Diarios" & vbCrLf & Err.Description
End Function


Private Function InsertarCabeceraOrden(ByRef Rs As ADODB.Recordset, CodOrden As String, Mens As String) As Boolean
Dim Cliente As Long
Dim Sql4 As String
Dim HoraIni As String
Dim HoraFin As String
Dim SQL1 As String
Dim Sql2 As String

    On Error GoTo eInsertarCabeceraOrden
    
    InsertarCabeceraOrden = False

    If DBLet(Rs!numpedid) = "" Then
        Cliente = DevuelveValor("select codclien from pedidos where numpedid is null")
        SQL1 = "select min(time(horaini)) from palets where numpedid is null  and fechafin = " & DBSet(txtcodigo(6).Text, "F") & " and intorden = 0"
        Sql2 = "select max(time(horafin)) from palets where numpedid is null  and fechafin = " & DBSet(txtcodigo(6).Text, "F") & " and intorden = 0"
    Else
        Cliente = DevuelveValor("select codclien from pedidos where numpedid = " & DBSet(Rs!numpedid, "N"))
        SQL1 = "select min(time(horaini)) from palets where numpedid = " & DBSet(Rs!numpedid, "N") & " and fechafin = " & DBSet(txtcodigo(6).Text, "F") & " and intorden = 0"
        Sql2 = "select max(time(horafin)) from palets where numpedid = " & DBSet(Rs!numpedid, "N") & " and fechafin = " & DBSet(txtcodigo(6).Text, "F") & " and intorden = 0"
    End If
    HoraIni = DevuelveValor(SQL1)
    HoraFin = DevuelveValor(Sql2)
    
    Sql4 = "insert into cccaborden (codorden, fechaini, fechafin, codclien, observac) values ("
    Sql4 = Sql4 & DBSet(CodOrden, "N") & "," & DBSet(txtcodigo(6).Text & " " & HoraIni, "FH") & ","
    Sql4 = Sql4 & DBSet(txtcodigo(6).Text & " " & HoraFin, "FH") & "," & DBSet(Cliente, "N") & ","
    Sql4 = Sql4 & ValorNulo & ")"
    conn.Execute Sql4

    ' Insertamos que ordenes de confeccion hemos creado
    Sql4 = "insert into tmpinformes(codusu, importe1) values (" & vUsu.Codigo & "," & DBSet(CodOrden, "N") & ")"
    conn.Execute Sql4



    InsertarCabeceraOrden = True
    Exit Function

eInsertarCabeceraOrden:
    Mens = Mens & vbCrLf & vbCrLf & Err.Description
End Function


Private Function InsertarLineasVariedadOrden(ByRef Rs As ADODB.Recordset, CodOrden As String, Mens As String) As Boolean
Dim Sql2 As String
Dim Sql4 As String
Dim CadValues As String
Dim NumLin As Long
Dim Rs2 As ADODB.Recordset

On Error GoTo eInsertarLineasOrden


    InsertarLineasVariedadOrden = False

    If DBLet(Rs!numpedid) = "" Then
        Sql2 = "select codvarie, codforfait, sum(if(pesoneto is null, 0, pesoneto)) pesoneto, sum(if(numcajas is null, 0, numcajas)) numcajas from palets_variedad, palets where palets.numpedid is null "
    Else
        Sql2 = "select codvarie, codforfait, sum(if(pesoneto is null, 0, pesoneto)) pesoneto, sum(if(numcajas is null, 0, numcajas)) numcajas from palets_variedad, palets where palets.numpedid = " & DBSet(Rs!numpedid, "N")
    End If
    Sql2 = Sql2 & " and palets.intorden = 0 and date(fechafin) = " & DBSet(txtcodigo(6).Text, "F") & " and palets.numpalet = palets_variedad.numpalet "
    Sql2 = Sql2 & " group by 1,2 "
    Sql2 = Sql2 & " order by 1,2 "

    Sql4 = "insert into cclinorden3 (codorden,numlinea,codvarie,codforfait,kilosnet,numcajon) values "
    
    CadValues = ""
    NumLin = 0
    
    Set Rs2 = New ADODB.Recordset
    Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs2.EOF
        NumLin = NumLin + 1
        
        CadValues = CadValues & "(" & DBSet(CodOrden, "N") & "," & DBSet(NumLin, "N") & "," & DBSet(Rs2!codvarie, "N") & ","
        CadValues = CadValues & DBSet(Rs2!codforfait, "T") & "," & DBSet(Rs2!Pesoneto, "N") & "," & DBSet(Rs2!NumCajas, "N") & "),"
        
        Rs2.MoveNext
    Wend
    
    If CadValues <> "" Then
        ' quitamos la ultima coma
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
        conn.Execute Sql4 & CadValues
        
    End If
    
    Set Rs2 = Nothing
    
    InsertarLineasVariedadOrden = True
    Exit Function
        
eInsertarLineasOrden:
    Mens = Mens & vbCrLf & vbCrLf & Err.Description
End Function

Private Function InsertarLineasResumenVariedadOrden(ByRef Rs As ADODB.Recordset, CodOrden As String, Mens As String) As Boolean
Dim Sql2 As String
Dim Sql4 As String
Dim CadValues As String
Dim NumLin As Long
Dim Rs2 As ADODB.Recordset

On Error GoTo eInsertarLineasResumenVariedadOrden


    InsertarLineasResumenVariedadOrden = False

    If DBLet(Rs!numpedid) = "" Then
        Sql2 = "select codforfait, sum(if(pesoneto is null, 0, pesoneto)) pesoneto, sum(if(numcajas is null, 0, numcajas)) numcajas from palets_variedad, palets where palets.numpedid is null "
    Else
        Sql2 = "select codforfait, sum(if(pesoneto is null, 0, pesoneto)) pesoneto, sum(if(numcajas is null, 0, numcajas)) numcajas from palets_variedad, palets where palets.numpedid = " & DBSet(Rs!numpedid, "N")
    End If
    Sql2 = Sql2 & " and palets.intorden = 0 and fechafin = " & DBSet(txtcodigo(6).Text, "F") & " and palets.numpalet = palets_variedad.numpalet "
    Sql2 = Sql2 & " group by 1 "
    Sql2 = Sql2 & " order by 1 "

    Sql4 = "insert into cclinorden4 (codorden,numlinea,codforfait,kilosnet,numcajon) values "
    
    CadValues = ""
    NumLin = 0
    
    Set Rs2 = New ADODB.Recordset
    Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs2.EOF
        NumLin = NumLin + 1
        
        CadValues = CadValues & "(" & DBSet(CodOrden, "N") & "," & DBSet(NumLin, "N") & ","
        CadValues = CadValues & DBSet(Rs2!codforfait, "T") & "," & DBSet(Rs2!Pesoneto, "N") & "," & DBSet(Rs2!NumCajas, "N") & "),"
        
        Rs2.MoveNext
    Wend
    
    If CadValues <> "" Then
        ' quitamos la ultima coma
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
        conn.Execute Sql4 & CadValues
    End If
    
    Set Rs2 = Nothing
    
    InsertarLineasResumenVariedadOrden = True
    Exit Function
        
eInsertarLineasResumenVariedadOrden:
    Mens = Mens & vbCrLf & vbCrLf & Err.Description
End Function




Private Function InsertarPaletsOrden(ByRef Rs As ADODB.Recordset, CodOrden As String, Mens As String) As Boolean
Dim Sql2 As String
Dim Sql4 As String
Dim CadValues As String
Dim NumLin As Long
Dim Rs2 As ADODB.Recordset

On Error GoTo eInsertarPaletsOrden


    InsertarPaletsOrden = False

    If DBLet(Rs!numpedid) = "" Then
        Sql2 = "select " & CodOrden & ", numpalet from palets where palets.numpedid is null "
    Else
        Sql2 = "select " & CodOrden & ", numpalet from palets where palets.numpedid = " & DBSet(Rs!numpedid, "N")
    End If
    Sql2 = Sql2 & " and palets.intorden = 0  and fechafin = " & DBSet(txtcodigo(6).Text, "F")

    Sql4 = "insert into cclinorden5 (codorden,numpalet) "
    
    conn.Execute Sql4 & Sql2
    
    InsertarPaletsOrden = True
    Exit Function
        
eInsertarPaletsOrden:
    Mens = Mens & vbCrLf & Err.Description
End Function


Private Function InsertarTrabajadoresOrden(Mens As String) As Boolean
Dim SQL As String
Dim Sql4 As String
Dim Sql5 As String
Dim CadValues As String
Dim NumLin As Long
Dim Rs2 As ADODB.Recordset
Dim HoraIni As String
Dim HoraFin As String
Dim HoraIni2 As String
Dim HoraFin2 As String
Dim Continuar As String
Dim Diferencia As Currency
Dim Rs As ADODB.Recordset
Dim NumF As Long
Dim CodCost As String
Dim b As Boolean

On Error GoTo eInsertarTrabajadoresOrden

    InsertarTrabajadoresOrden = False
    
    SQL = "select codtraba, time(fechaini), time(fechafin), codlinconf, codcoste from cctrabaconf where date(fechafin) = " & DBSet(txtcodigo(6).Text, "F")
    SQL = SQL & " and procesado = 0 and sinacabar = 0 and not codlinconf is null"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        HoraIni = Mid(Rs.Fields(1).Value, 12, 8)
        HoraFin = Mid(Rs.Fields(2).Value, 12, 8)
        
        Sql4 = "select distinct cclinorden5.codorden, cccaborden.fechaini horaini, cccaborden.fechafin horafin "
        Sql4 = Sql4 & " from tmpinformes, palets, cclinorden5, cccaborden where tmpinformes.codusu = " & vUsu.Codigo
        Sql4 = Sql4 & " and tmpinformes.importe1 = cclinorden5.codorden "
        Sql4 = Sql4 & " and cclinorden5.codorden = cccaborden.codorden "
        Sql4 = Sql4 & " and cclinorden5.numpalet = palets.numpalet "
        Sql4 = Sql4 & " and palets.codlinconf = " & DBSet(Rs!codlinconf, "N")
        Sql4 = Sql4 & " order by 2,3,1 "
        
        HoraIni2 = HoraIni
        HoraFin2 = HoraFin
       
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql4, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        b = True
        
        CodCost = DevuelveDesdeBDNew(cAgro, "ccconcostes", "codcoste", "codcoste", CStr(DBLet(Rs!codCoste, "N")), "N")
        
        While Not Rs2.EOF And HoraIni2 <> HoraFin2 And b
            
            'la orden tiene una fecha de inicio superior a la fecha de inicio del trabajador
            If CDate(Rs2!HoraIni) > CDate(txtcodigo(6).Text & " " & HoraIni2) Then
                ' que el trabajador acabe antes que el inicio de la orden
                If CDate(txtcodigo(6).Text & " " & HoraFin2) < CDate(Rs2!HoraIni) Then
                    
                    ' Insertamos en la tabla auxiliar para procesos diarios todo el tiempo del trabajador
                    b = InsertarCosteDiario(DBLet(Rs!codtraba), txtcodigo(6).Text, CodCost, HoraIni2, HoraFin2, Mens)
                    
                    HoraIni2 = HoraFin2
                    
                    Rs2.MoveNext
                Else
                ' que el trabajador acabe despues de la orden
                    
                    ' Insertamos en la tabla auxiliar para procesos diarios desde inicio del trabajador a inicio de orden
                    b = InsertarCosteDiario(DBLet(Rs!codtraba), txtcodigo(6).Text, CodCost, HoraIni2, Mid(Rs2!HoraIni, 12, 8), Mens)
                    
                    HoraIni2 = Mid(DBLet(Rs2!HoraIni), 12, 8)
                    
                    ' no cambiamos de registro para que procese el resto de tiempo del trabajador
                    'Rs2.MoveNext
                End If
            Else
                ' si el rango de tiempo del trabajador esta dentro del rango del palet, ese trabajador ha estado trabajando todo el tiempo
                ' sobre esa orden
                If CDate(Rs2!HoraIni) <= CDate(txtcodigo(6).Text & " " & HoraIni2) And CDate(txtcodigo(6).Text & " " & HoraFin2) <= CDate(Rs2!HoraFin) Then
                    
                    Diferencia = Round2(DateDiff("n", CDate(txtcodigo(6).Text & " " & Format(HoraIni2, "hh:mm:ss")), CDate(txtcodigo(6).Text & " " & Format(HoraFin2, "hh:mm:ss"))) / 60, 2)
                    
                    NumF = SugerirCodigoSiguienteStr("cclinorden1", "numlinea", "codorden=" & DBLet(Rs2!CodOrden, "N"))
                    
                    Sql5 = "insert into cclinorden1 (codorden,numlinea,codtraba,codcoste,fechaini,fechafin,horas) values ("
                    Sql5 = Sql5 & DBSet(Rs2!CodOrden, "N") & "," & NumF & "," & DBSet(Rs!codtraba, "N") & "," & DBSet(CodCost, "N") & ","
                    Sql5 = Sql5 & DBSet(txtcodigo(6).Text & " " & HoraIni2, "FH") & "," & DBSet(txtcodigo(6).Text & " " & HoraFin2, "FH") & ","
                    Sql5 = Sql5 & DBSet(Diferencia, "N") & ")  "
                        
                    conn.Execute Sql5
                    
                    HoraIni2 = HoraFin2
                    
                    Rs2.MoveNext
                Else
                    ' si el palet se acaba antes de que el trabajador acabe
                    If CDate(Rs2!HoraIni) <= CDate(txtcodigo(6).Text & " " & HoraIni2) And CDate(txtcodigo(6).Text & " " & HoraIni2) < CDate(Rs2!HoraFin) Then
                        Diferencia = Round2(DateDiff("n", CDate(txtcodigo(6).Text & " " & Format(HoraIni2, "hh:mm:ss")), CDate(Rs2!HoraFin)) / 60, 2)
                        
                        NumF = SugerirCodigoSiguienteStr("cclinorden1", "numlinea", "codorden=" & DBLet(Rs2!CodOrden, "N"))
                        
                        Sql5 = "insert into cclinorden1 (codorden,numlinea,codtraba,codcoste,fechaini,fechafin,horas) values ("
                        Sql5 = Sql5 & DBSet(Rs2!CodOrden, "N") & "," & NumF & "," & DBSet(Rs!codtraba, "N") & "," & DBSet(CodCost, "N") & ","
                        Sql5 = Sql5 & DBSet(txtcodigo(6).Text & " " & HoraIni2, "FH") & "," & DBSet(Rs2!HoraFin, "FH") & ","
                        Sql5 = Sql5 & DBSet(Diferencia, "N") & ")  "
                        
                        conn.Execute Sql5
                        
                        'resto de tiempo de trabajo del trabajador
                        HoraIni2 = Mid(DBLet(Rs2!HoraFin, "H"), 12, 8)
                        HoraFin2 = HoraFin
                    
                        Rs2.MoveNext
                    
                    Else
                        'en este caso el trabajador no trabaja en esta orden
                        Rs2.MoveNext
                    End If
                End If
            End If
            
        Wend
        
        
        'el trabajador ha estado haciendo cosas varias (no asociadas a ninguna orden de confeccion en este tramo)
        If HoraIni2 <> HoraFin2 Then
            b = InsertarCosteDiario(CStr(Rs!codtraba), txtcodigo(6).Text, CodCost, HoraIni2, HoraFin2, Mens)
        End If
        
        'marcamos que el registro ha sido procesado
        If b Then
            b = ActualizarRegistro(CStr(Rs!codtraba), txtcodigo(6).Text)
        End If
        
        ' pasamos al siguiente trabajador
        Rs.MoveNext
        
    Wend
    
    Set Rs = Nothing
    
    
    InsertarTrabajadoresOrden = True
    Exit Function
        
eInsertarTrabajadoresOrden:
    Mens = Mens & vbCrLf & vbCrLf & Err.Description
End Function

Private Function ActualizarRegistro(Trabajador As String, fecha As String) As Boolean
Dim SQL As String

    On Error GoTo eActualizarRegistro

    ActualizarRegistro = False
    
    SQL = "update cctrabaconf set procesado = 1 where codtraba = " & DBSet(Trabajador, "N")
    SQL = SQL & " and date(fechaini) = " & DBSet(fecha, "F")
    
    conn.Execute SQL
    
    ActualizarRegistro = True
    Exit Function
    
eActualizarRegistro:
    MuestraError Err.Number, "Actualizar Registro", Err.Description
End Function


Private Function InsertarCosteDiario(Trabajador As String, fecha As String, CodCost As String, HoraIni As String, HoraFin As String, Mens As String) As Boolean
Dim Diferencia As Currency
Dim NumF As Long
Dim SQL As String
Dim FechaIni As String
Dim FechaFin As String

    On Error GoTo eInsertarCosteDiario
    
    InsertarCosteDiario = False
    
    Diferencia = Round2(DateDiff("n", CDate(fecha & " " & Format(HoraIni, "hh:mm:ss")), CDate(fecha & " " & Format(HoraFin, "hh:mm:ss"))) / 60, 2)
                    
    NumF = SugerirCodigoSiguienteStr("tmpcclindia1", "numlinea", "fecha = " & DBSet(fecha, "F") & " and codcoste = " & DBSet(CodCost, "N"))

    FechaIni = fecha & " " & HoraIni
    FechaFin = fecha & " " & HoraFin

    SQL = "insert into tmpcclindia1 (fecha,codcoste,numlinea,codtraba,fechaini,fechafin,horas) values ("
    SQL = SQL & DBSet(fecha, "F") & "," & DBSet(CodCost, "N") & "," & DBSet(NumF, "N") & ","
    SQL = SQL & DBSet(Trabajador, "N") & "," & DBSet(FechaIni, "FH") & "," & DBSet(FechaFin, "FH") & ","
    SQL = SQL & DBSet(Diferencia, "N") & ")"
    
    conn.Execute SQL
    
    InsertarCosteDiario = True
    Exit Function
    
eInsertarCosteDiario:
    Mens = Mens & vbCrLf & "Insertar Coste Diario:" & vbCrLf & Err.Description
End Function


Private Sub CmdAcepInfCostes_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String

Dim SQL As String

    InicializarVbles
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H concepto
    cDesde = Trim(txtcodigo(14).Text)
    cHasta = Trim(txtcodigo(15).Text)
    nDesde = txtNombre(14).Text
    nHasta = txtNombre(15).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{cctrabaconf.codcoste}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHConcepto= """) Then Exit Sub
    End If
    
    'D/H Trabajador
    cDesde = Trim(txtcodigo(16).Text)
    cHasta = Trim(txtcodigo(17).Text)
    nDesde = txtNombre(16).Text
    nHasta = txtNombre(17).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{cctrabaconf.codtraba}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHTrabajador= """) Then Exit Sub
    End If
    
    'D/H Fecha
    cDesde = Trim(txtcodigo(18).Text)
    cHasta = Trim(txtcodigo(19).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "date({cctrabaconf.fechaini})"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    Select Case Combo1(0).ListIndex
        Case 0 ' no procesados
            If Not AnyadirAFormula(cadselect, "{cctrabaconf.procesado} = 0") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "{cctrabaconf.procesado} = 0") Then Exit Sub
        Case 1 ' procesados
            If Not AnyadirAFormula(cadselect, "{cctrabaconf.procesado} = 1") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "{cctrabaconf.procesado} = 1") Then Exit Sub
        Case 2 ' ambos
        
    End Select
    
    
    cadTABLA = " cctrabaconf INNER JOIN straba ON cctrabaconf.codtraba = straba.codtraba "
    
    If HayRegParaInforme(cadTABLA, cadselect) Then
    
        cadTitulo = "Informe de Costes por Fecha"
        cadNombreRPT = "rCCInfCostes.rpt"
        
        LlamarImprimir
    End If

End Sub

Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim SQL As String
Dim vsqlVariedad As String

    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    SQL = "select count(*) from cctrabaconf where date(fechaini) >= " & DBSet(FDesde, "F") & " and date(fechafin) <= " & DBSet(FHasta, "F")
    
    If TotalRegistros(SQL) = 0 Then
        MsgBox "No hay fichajes en esta fecha para construir ordenes de confección.", vbExclamation
    
    Else
    
        Variedades = ""
        Forfaits = ""
    
    
        vsqlVariedad = ""
        If txtcodigo(4).Text <> "" Then vsqlVariedad = vsqlVariedad & " and variedades.codvarie >= " & DBSet(txtcodigo(4).Text, "N")
        If txtcodigo(5).Text <> "" Then vsqlVariedad = vsqlVariedad & " and variedades.codvarie <= " & DBSet(txtcodigo(5).Text, "N")
        
        '[Monica]21/11/2013: antes tipovarie2
        vsqlVariedad = vsqlVariedad & " and tipovariedad = " & Combo1(1).ListIndex
    
        If txtcodigo(25).Text <> "" Then vsqlVariedad = vsqlVariedad & " and variedades.codclase >= " & DBSet(txtcodigo(25).Text, "N")
        If txtcodigo(26).Text <> "" Then vsqlVariedad = vsqlVariedad & " and variedades.codclase <= " & DBSet(txtcodigo(26).Text, "N")
        
    
        Set frmMensVariedad = New frmMensajes
    
        frmMensVariedad.OpcionMensaje = 21
        frmMensVariedad.Label5 = "Variedades"
        frmMensVariedad.cadwhere = vsqlVariedad
        frmMensVariedad.Show vbModal
    
        Set frmMensVariedad = Nothing
        
        
        vsqlVariedad = ""
        If txtcodigo(0).Text <> "" Then vsqlVariedad = vsqlVariedad & " and forfaits.codforfait >= " & DBSet(txtcodigo(0).Text, "T")
        If txtcodigo(1).Text <> "" Then vsqlVariedad = vsqlVariedad & " and forfaits.codforfait <= " & DBSet(txtcodigo(1).Text, "T")
        
        
        Set frmMensForfaits = New frmMensajes
    
        frmMensForfaits.OpcionMensaje = 21
        frmMensForfaits.Label5 = "Forfaits"
        frmMensForfaits.cadwhere = vsqlVariedad
        frmMensForfaits.Show vbModal
    
        Set frmMensForfaits = Nothing
        
        If Variedades = "" Then
            MsgBox "No hay datos para mostrar en el Informe.", vbExclamation
            Exit Sub
        End If
        
        '[Monica]09/07/2012: comprobamos que todos los registros a procesar se han acabado
        '                    si hay registros con la tarea sin acabar damos un aviso para que los revisen
        SQL = "select count(*) from cctrabaconf where date(fechaini) >= " & DBSet(FDesde, "F") & " and date(fechafin) <= " & DBSet(FHasta, "F") & " and sinacabar = 1"
        If TotalRegistros(SQL) Then
            MsgBox "Hay registros de trabajador que no se han acabado. Revise.", vbExclamation
            PonerFoco txtcodigo(6)
            Exit Sub
        End If
    
        ' Comprobamos las claves referenciales del fichero que cargan que no las tiene en la base de datos para que no dé errores
        ' al cargar la tabla
        If ComprobarFicheroNew(CStr(FDesde)) Then
            If TotalRegistrosConsulta("select * from tmpinformes where codusu = " & vUsu.Codigo) <> 0 Then
                cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
                cadTitulo = "Errores en el Traspaso de Costes"
                cadNombreRPT = "rErroresTrasCostes.rpt"
                LlamarImprimir
            Else
                If MsgBox("Atención este proceso borra las tablas de cálculo antes de cargarlas. ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                    If ProcesoCargaOrdenesNew(CStr(FDesde), CStr(FHasta), pb5, Label4(42), Label4(43)) Then
                    
                        ' insertamos un report de aquellos importes no asignados a ninguna variedad
                        SQL = "select * from cccabecera where codcoste in (select codcoste from ccconcostes where tipocoste = 0 and tipokilos <> 0) and codvarie = 0"
                        If TotalRegistrosConsulta(SQL) <> 0 Then
                            If MsgBox("Se ha encontrado fichajes no asignados a variedad. ¿Desea imprimirlos?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                                cadTitulo = "Revision de Fichadas"
                                cadNombreRPT = "rCCRevisionFichadas.rpt"
                                
                                cadFormula = "{ccconcostes.tipocoste} = 0 and {ccconcostes.tipokilos} <> 0 and {cccabecera.codvarie} = 0"
                                
                                LlamarImprimir
                            End If
                            
                            If MsgBox("¿ Quiere seguir con el proceso de cálculo ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                                Exit Sub
                            End If
                        End If
                        
                    
                        ' cargamos las tablas intermedias para el informe
                        If ProcesoCargaCampo(Str(FDesde), Str(FHasta), pb5, Label4(42), Label4(43)) Then
                            If HayRegParaInforme("cccabecera", "") Then 'HayRegParaInforme("tmpinformes", "codusu = " & vUsu.Codigo)  Then
                                'D/H Variedad
                                cDesde = Trim(txtcodigo(4).Text)
                                cHasta = Trim(txtcodigo(5).Text)
                                nDesde = txtNombre(4).Text
                                nHasta = txtNombre(5).Text
                                If Not (cDesde = "" And cHasta = "") Then
                                    'Cadena para seleccion Desde y Hasta
                                    Codigo = "{" & Tabla & ".codvarie}"
                                    TipCod = "N"
                                    If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad= """) Then Exit Sub
                                End If
                                
                                'D/H Forfait
                                cDesde = Trim(txtcodigo(0).Text)
                                cHasta = Trim(txtcodigo(1).Text)
                                nDesde = txtNombre(0).Text
                                nHasta = txtNombre(1).Text
                                If Not (cDesde = "" And cHasta = "") Then
                                    'Cadena para seleccion Desde y Hasta
                                    Codigo = "{" & Tabla & ".codforfait}"
                                    TipCod = "N"
                                    If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHForfait= """) Then Exit Sub
                                End If
                                
                                cadFormula = ""
                                
                                cadParam = cadParam & "pUsu=" & vUsu.Codigo & "|"
                                numParam = numParam + 1
                                
                                cadParam = cadParam & "pForfaits=" & Me.Check1.Value & "|"
                                numParam = numParam + 1
                
                                cadTitulo = "Informe de Costes"
                                cadNombreRPT = "rCCInformeCostes.rpt"
                                
                                LlamarImprimir
                            End If
                        End If
                        cmdCancelar_Click
                    End If
                Else
                    Unload Me
                End If
            End If
        End If
    End If
            

End Sub

 
Private Function DevuelveKilos(Coste As String, Varie As String) As Long
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Tipo As Byte

    SQL = "select tipokilos from ccconcostes where codcoste = " & DBSet(Coste, "N")
    Tipo = DevuelveValor(SQL)
    
    Select Case Tipo
        Case 0
            SQL = "select kentrados from cclineas6 where codvarie = " & DBSet(Varie, "N")
        Case 1
            SQL = "select kvolcados from cclineas6 where codvarie = " & DBSet(Varie, "N")
        Case 2
            SQL = "select kconfeccionados from cclineas6 where codvarie = " & DBSet(Varie, "N")
    End Select
    
    DevuelveKilos = DevuelveValor(SQL)

End Function



Private Function ProcesoCargaCampo(FecDesde As String, FecHasta As String, ByRef pb1 As ProgressBar, Label4 As label, ByRef Label5 As label) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset

Dim Importe As Currency
Dim Kilos As Long

Dim C1_1_1 As Currency
Dim C1_2_1 As Currency
Dim C1_3_1 As Currency

Dim CadValues As String
Dim Coste As Currency
Dim KilosEntrados As Long

Dim SQLinsert As String
Dim ImporteTrans As Currency
Dim KilosTotal As Long
Dim ImporteTotal As Currency

Dim nRegs As Long
Dim CosteKilo As Currency
Dim KilosTot As Long

Dim Precio As Currency

On Error GoTo eProcesoCargaCampo
    
    ProcesoCargaCampo = False

    SQL = "delete from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
    conn.Execute SQL
    
    SQL = "delete from cclineas3 "
    conn.Execute SQL
    
    'para imprimir costes de envases
    SQL = "delete from cclineas5 "
    conn.Execute SQL
    
    ' para imprimir costes de paletizacion
    SQL = "delete from cclineas4 "
    conn.Execute SQL
    
    
    ' Insertamos en la temporal                      p1,      p2,     p3,     importe,  coste,   kilos, codvarie, codcoste
    SQLinsert = "insert into tmpinformes (codusu, codigo1, campo1, campo2, importe1, precio1, importe2, importe3, importe4) values "
    
    Label5.visible = True
    pb1.visible = True
    Label5.Caption = "Cargando Tablas intermedias"
        
'1.-Mano de obra de recoleccion
    Label4.visible = True
    Label4.Caption = "Procesando mano de obra de recoleccion "
    DoEvents
    
    SQL = "select codvarie, sum(kilosrec), sum(importe) from ((rpartes_trabajador inner join rpartes on rpartes_trabajador.nroparte = rpartes.nroparte) "
    SQL = SQL & " inner join rcuadrilla on rpartes.codcuadrilla = rcuadrilla.codcuadrilla) "
    SQL = SQL & " inner join rcapataz on rcuadrilla.codcapat = rcapataz.codcapat "
    SQL = SQL & " where rpartes.fecentrada between " & DBSet(FecDesde, "F") & " and " & DBSet(FecHasta, "F")
    SQL = SQL & " and rpartes_trabajador.codtraba <> if (rcapataz.codtraba  is null,0,rcapataz.codtraba) "
    
'    If txtCodigo(4).Text <> "" Then SQL = SQL & " and rpartes_trabajador.codvarie >= " & DBSet(txtCodigo(4).Text, "N")
'    If txtCodigo(5).Text <> "" Then SQL = SQL & " and rpartes_trabajador.codvarie <= " & DBSet(txtCodigo(5).Text, "N")
    SQL = SQL & " and rpartes_trabajador.codvarie in " & Variedades
    
    SQL = SQL & " group by codvarie order by codvarie "
     
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
     
    CadValues = ""
    
    While Not Rs.EOF
        C1_1_1 = 0
        
        Importe = ComprobarCero(DBLet(Rs.Fields(2).Value))
        Kilos = ComprobarCero(DBLet(Rs.Fields(1).Value))
        If Kilos <> 0 Then C1_1_1 = Round2(Importe / Kilos, 4)
    
        CadValues = CadValues & "(" & vUsu.Codigo & ",1,1,1," & DBSet(Importe, "N") & "," & DBSet(C1_1_1, "N") & ","
        CadValues = CadValues & DBSet(Kilos, "N") & "," & DBSet(Rs.Fields(0).Value, "N") & ",0),"
    
        Rs.MoveNext
    Wend
    Set Rs = Nothing
        
    If CadValues <> "" Then
        ' quitamos la ultima coma
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
        conn.Execute SQLinsert & CadValues
    End If
    
'1.2.1 Jefe de campo (Periodico)


'1.2.2 Seguridad Social (Periodico)
    
    
    
    
'2.-Mano de obra de jefes de campo
'    Label4.visible = True
'    Label4.Caption = "Procesando mano de obra de jefes de campo "
'    DoEvents
'
'    Sql = "select codvarie, sum(kilosrec), sum(importe) from ((rpartes_trabajador inner join rpartes on rpartes_trabajador.nroparte = rpartes.nroparte) "
'    Sql = Sql & " inner join rcuadrilla on rpartes.codcuadrilla = rcuadrilla.codcuadrilla) "
'    Sql = Sql & " inner join rcapataz on rcuadrilla.codcapat = rcapataz.codcapat "
'    Sql = Sql & " where rpartes.fecentrada between " & DBSet(FecDesde, "F") & " and " & DBSet(FecHasta, "F")
'    Sql = Sql & " and rpartes_trabajador.codtraba = if (rcapataz.codtraba  is null,0,rcapataz.codtraba) "
'
'    If txtCodigo(4).Text <> "" Then Sql = Sql & " and rpartes_trabajador.codvarie >= " & DBSet(txtCodigo(4).Text, "N")
'    If txtCodigo(5).Text <> "" Then Sql = Sql & " and rpartes_trabajador.codvarie <= " & DBSet(txtCodigo(5).Text, "N")
'
'    Sql = Sql & " group by codvarie order by codvarie "
'
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    CadValues = ""
'
'    While Not Rs.EOF
'        C1_2_1 = 0
'        Importe = ComprobarCero(DBLet(Rs.Fields(2).Value))
'        Kilos = ComprobarCero(DBLet(Rs.Fields(1).Value))
'        If Kilos <> 0 Then C1_2_1 = Round2(Importe / Kilos, 4)
'
'        CadValues = CadValues & "(" & vUsu.Codigo & ",1,2,1," & DBSet(Importe, "N") & "," & DBSet(C1_2_1, "N") & ","
'        CadValues = CadValues & DBSet(Kilos, "N") & "," & DBSet(Rs.Fields(0).Value, "N") & ",0),"
'
'        Rs.MoveNext
'    Wend
'    Set Rs = Nothing
'    If CadValues <> "" Then
'        ' quitamos la ultima coma
'        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
'        conn.Execute SQLinsert & CadValues
'    End If
    
'3.-transporte
    Label4.visible = True
    Label4.Caption = "Procesando transporte "
    DoEvents
    
'    Sql = "select sum(importe) from rpartes_gastos inner join rpartes on rpartes_gastos.nroparte = rpartes.nroparte "
'    Sql = Sql & " where rpartes.fecentrada between " & DBSet(FecDesde, "F") & " and " & DBSet(FecHasta, "F")
'
'    ImporteTrans = DevuelveValor(Sql)
    
    SQL = "select codvarie, sum(kilosrec) kilos from rpartes_variedad inner join rpartes on rpartes_variedad.nroparte = rpartes.nroparte "
    SQL = SQL & " where rpartes.fecentrada between " & DBSet(FecDesde, "F") & " and " & DBSet(FecHasta, "F")
    
    SQL = SQL & " and rpartes_variedad.codvarie in " & Variedades
    
    SQL = SQL & " group by codvarie order by codvarie "
    
    KilosTotal = DevuelveValor("select sum(kilos)  from (" & SQL & ") aaaaa")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
     
    CadValues = ""
    
    While Not Rs.EOF
        C1_3_1 = 0
        Importe = 0
        
        Precio = DevuelveValor("select eurmanob from variedades where codvarie = " & DBSet(Rs!codvarie, "N"))
        
'        If KilosTotal <> 0 Then
'            Importe = Round2(ImporteTrans * DBLet(Rs.Fields(1).Value, "N") / KilosTotal, 2)
'        End If
        Kilos = ComprobarCero(DBLet(Rs.Fields(1).Value))
        
        Importe = Round2(Kilos * Precio, 2)
        If Kilos <> 0 Then C1_3_1 = Round2(Importe / Kilos, 4)
    
        CadValues = CadValues & "(" & vUsu.Codigo & ",1,3,1," & DBSet(Importe, "N") & "," & DBSet(C1_3_1, "N") & ","
        CadValues = CadValues & DBSet(Kilos, "N") & "," & DBSet(Rs.Fields(0).Value, "N") & ",0),"
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing

    If CadValues <> "" Then
        ' quitamos la ultima coma
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
        conn.Execute SQLinsert & CadValues
    End If

'4.- cursor por puntos de menu para calculo de importes y de kilos segun las tareas de costes directos
    Label4.visible = True
    Label4.Caption = "Procesando tareas de costes directos "
    DoEvents
    
    SQL = "select p1,p2,p3, cccabecera.codvarie, sum(cccabecera.importe) importe, sum(cccabecera.kilos) kilos, cccabecera.codcoste "
    SQL = SQL & " from cccabecera inner join ccconcostes on cccabecera.codcoste = ccconcostes.codcoste "
    SQL = SQL & " WHERE ccconcostes.tipocoste = 0 " ' solo costes directos
    SQL = SQL & " and tipokilos <> 2"
    
'    If txtCodigo(4).Text <> "" Then SQL = SQL & " and cccabecera.codvarie >= " & DBSet(txtCodigo(4).Text, "N")
'    If txtCodigo(5).Text <> "" Then SQL = SQL & " and cccabecera.codvarie <= " & DBSet(txtCodigo(5).Text, "N")
    SQL = SQL & " and cccabecera.codvarie in " & Variedades
    
    SQL = SQL & " group by 1,2,3,4, 7 order by 1,2,3,4, 7 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CadValues = ""
    
    While Not Rs.EOF
        
        Kilos = DBLet(Rs!Kilos, "N") 'DevuelveKilos(CStr(Rs!codCoste), CStr(Rs!codvarie))
        
        Coste = 0
        If ComprobarCero(CStr(Kilos)) <> 0 Then
            Coste = Round2(DBLet(Rs.Fields(4).Value, "N") / Kilos, 4)
        End If
    
        CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(Rs.Fields(1).Value, "N") & ","
        CadValues = CadValues & DBSet(Rs.Fields(2).Value, "N") & "," & DBSet(Rs.Fields(4).Value, "N") & "," & DBSet(Coste, "N") & ","
        CadValues = CadValues & DBSet(Kilos, "N") & "," & DBSet(Rs.Fields(3).Value, "N") & ","
        CadValues = CadValues & DBSet(Rs.Fields(6).Value, "N") & "),"
    
        Rs.MoveNext
    Wend
    
    If CadValues <> "" Then
        ' quitamos la ultima coma
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
        conn.Execute SQLinsert & CadValues
    End If
    
    Set Rs = Nothing
    
'5.- cursor de los costes indirectos de ese periodo (sobre los kilos entrados)
    Label4.visible = True
    Label4.Caption = "Procesando tareas de costes periodicos "
    DoEvents
    
    SQL = "select p1,p2,p3, sum(if(importereal is null or importereal = 0, if(importeprev is null, 0, importeprev),0)) importe, ccconcostes_mes.codcoste "
    SQL = SQL & " from ccconcostes_mes inner join ccconcostes on ccconcostes_mes.codcoste = ccconcostes.codcoste "
    SQL = SQL & " WHERE ccconcostes.tipocoste = 1 " ' solo costes indirectos
    SQL = SQL & " AND date(concat(año,""-"",MES,""-01"")) between " & DBSet(FecDesde, "F") & " and " & DBSet(FecHasta, "F")
    SQL = SQL & " group by 1,2,3,5 order by 1,2,3,5 "
    
    nRegs = TotalRegistrosConsulta(SQL)
    
    pb1.visible = True
    nRegs = TotalRegistrosConsulta(SQL)
    CargarProgresNew pb1, CInt(nRegs)
    pb1.Value = 0
    DoEvents
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CadValues = ""
    
    While Not Rs.EOF
        
        IncrementarProgresNew pb1, 1
        
        ImporteTotal = DBLet(Rs.Fields(3).Value, "N")
        
        
        If DBLet(Rs!p1) = 1 And DBLet(Rs!p2) = 2 And DBLet(Rs!p3) = 1 Then
            '[Monica] Si es Jefes de campo solo sobre el total de las variedades convencionales
            ' sobre los kilos de las variedades que no sean bio                                                               '[Monica]21/11/2013: antes tipovarie2
            Sql2 = "select codvarie, kentrados kilos from cclineas6 where codvarie in (select codvarie from variedades where tipovariedad = 0) "
            Sql2 = Sql2 & " order by 1 "
        Else
        
            Sql2 = "select codvarie, kentrados kilos from cclineas6 where (1=1) "
            Sql2 = Sql2 & " order by 1 "
        End If
        KilosTotal = DevuelveValor("select sum(kilos) from (" & Sql2 & ") aaaaaaa ")
        
        Coste = 0
        If KilosTotal <> 0 Then
            Coste = Round2(ImporteTotal / KilosTotal, 4)
        End If

        CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(Rs.Fields(1).Value, "N") & ","
        CadValues = CadValues & DBSet(Rs.Fields(2).Value, "N") & "," & DBSet(ImporteTotal, "N") & "," & DBSet(Coste, "N") & ","
        CadValues = CadValues & DBSet(KilosTotal, "N") & "," & DBSet(0, "N") & ","   ' es para todas las variedades
        CadValues = CadValues & DBSet(Rs!codCoste, "N") & "),"

        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    If CadValues <> "" Then
        ' quitamos la ultima coma
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
        conn.Execute SQLinsert & CadValues
    End If
    
'6.- calculamos los costes asociados por confeccion
    Label4.visible = True
    Label4.Caption = "Procesando costes de postconfeccion "
    DoEvents
        
    SQLinsert = "insert into cclineas3 (idcontador, codvarie, codforfait, importe, kilos, codcoste) values "
    
    SQL = "select cccabecera.codvarie, cccabecera.codcoste, sum(cccabecera.importe) importe "
    SQL = SQL & " from cccabecera inner join ccconcostes on cccabecera.codcoste = ccconcostes.codcoste"
    SQL = SQL & " where p1 = 2 And p2 = 1 And p3 = 2"
    If Variedades <> "" Then SQL = SQL & " and codvarie in " & Variedades
    SQL = SQL & " group by 1,2 order by 1,2 "
    
    
    pb1.visible = True
    nRegs = TotalRegistrosConsulta(SQL)
    CargarProgresNew pb1, CInt(nRegs)
    pb1.Value = 0
    DoEvents

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CadValues = ""
    
    While Not Rs.EOF
        IncrementarProgresNew pb1, 1
        
        Sql2 = "select codforfait, sum(pesoneto) kilos "
        Sql2 = Sql2 & " from palets_variedad"
        Sql2 = Sql2 & " where numpalet IN (select distinct cclineas2.numpalet from cclineas2 where idcontador in (select idcontador from cccabecera where codcoste = " & DBSet(Rs!codCoste, "N") & " and codvarie = " & DBSet(Rs!codvarie, "N") & "))"

        If Forfaits <> "" Then Sql2 = Sql2 & " and codforfait in " & Forfaits
        Sql2 = Sql2 & " group by 1 order by 1"

        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

        Kilos = DevuelveKilos(CStr(Rs!codCoste), CStr(Rs!codvarie))
        
        While Not Rs2.EOF
            
            Importe = Round2(Rs!Importe * Rs2!Kilos / Kilos, 2)

            CadValues = CadValues & "(" & DBSet(0, "N") & "," & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs2!codforfait, "T") & "," & DBSet(Importe, "N") & ","
            CadValues = CadValues & DBSet(Rs2!Kilos, "N") & "," & DBSet(Rs!codCoste, "N") & "),"

            Rs2.MoveNext
        Wend
        
        Set Rs2 = Nothing
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    If CadValues <> "" Then
        ' quitamos la ultima coma
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
        conn.Execute SQLinsert & CadValues
    End If
    
'    Sql = "select idcontador, codvarie, cccabecera.importe, kilos, cccabecera.codcoste  "
'    Sql = Sql & " from cccabecera inner join ccconcostes on cccabecera.codcoste = ccconcostes.codcoste"
'    Sql = Sql & " where p1 = 2 And p2 = 1 And p3 = 2"
'
''    If txtCodigo(4).Text <> "" Then SQL = SQL & " and cccabecera.codvarie >= " & DBSet(txtCodigo(4).Text, "N")
''    If txtCodigo(5).Text <> "" Then SQL = SQL & " and cccabecera.codvarie <= " & DBSet(txtCodigo(5).Text, "N")
'    If Variedades <> "" Then Sql = Sql & " and cccabecera.codvarie in " & Variedades
'
'    Pb1.visible = True
'    nRegs = TotalRegistrosConsulta(Sql)
'    CargarProgresNew Pb1, CInt(nRegs)
'    Pb1.Value = 0
'    DoEvents
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    CadValues = ""
'
'    While Not Rs.EOF
'        IncrementarProgresNew Pb1, 1
'
'        Sql2 = "select codforfait, sum(pesoneto) kilos "
'        Sql2 = Sql2 & " from palets_variedad"
'        Sql2 = Sql2 & " where numpalet IN (select cclineas2.numpalet from cclineas2 where idcontador = " & Rs.Fields(0).Value & ")"
'
'        If Forfaits <> "" Then Sql2 = Sql2 & " and codforfait in " & Forfaits
'
'        Sql2 = Sql2 & " group by 1 order by 1"
'
'        Set Rs2 = New ADODB.Recordset
'        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'        While Not Rs2.EOF
'            Importe = Round2(Rs!Importe * Rs2!Kilos / Rs!Kilos, 2)
'
'            CadValues = CadValues & "(" & DBSet(Rs!idcontador, "N") & "," & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs2!codforfait, "T") & "," & DBSet(Importe, "N") & ","
'            CadValues = CadValues & DBSet(Rs2!Kilos, "N") & "," & DBSet(Rs!codCoste, "N") & "),"
'
'            Rs2.MoveNext
'        Wend
'
'        Set Rs2 = Nothing
'
'        Rs.MoveNext
'    Wend
'
'    Set Rs = Nothing
'    If CadValues <> "" Then
'        ' quitamos la ultima coma
'        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
'        conn.Execute SQLinsert & CadValues
'
'    End If
    
'7.-Calcular costes de envases
    Label4.visible = True
    Label4.Caption = "Procesando costes de envases "
    DoEvents
    
    SQL = "select codforfait, codvarie, sum(kilos) kilos from cclineas3 where (1=1) "
    
'    If txtCodigo(4).Text <> "" Then SQL = SQL & " and cclineas3.codvarie >= " & DBSet(txtCodigo(4).Text, "N")
'    If txtCodigo(5).Text <> "" Then SQL = SQL & " and cclineas3.codvarie <= " & DBSet(txtCodigo(5).Text, "N")
    If Variedades <> "" Then SQL = SQL & " and cclineas3.codvarie in " & Variedades
    
    SQL = SQL & " group by 1, 2 order by 1"
    
    pb1.visible = True
    nRegs = TotalRegistrosConsulta(SQL)
    CargarProgresNew pb1, CInt(nRegs)
    pb1.Value = 0
    DoEvents
    
    SQLinsert = "insert into cclineas5 (codforfait, codvarie, kilos, importe) values "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CadValues = ""
    
    While Not Rs.EOF
        IncrementarProgresNew pb1, 1

        CosteKilo = 0

        CalcularCosteForfait Rs!codforfait, Rs!Kilos, CosteKilo
        
        Importe = CosteKilo
        
        CadValues = CadValues & "(" & DBSet(Rs!codforfait, "T") & "," & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!Kilos, "N") & "," & DBSet(Importe, "N") & "),"
        
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    If CadValues <> "" Then
        ' quitamos la ultima coma
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
        conn.Execute SQLinsert & CadValues
    End If
    
    
'8.-Calcular costes de paletizacion
    Label4.visible = True
    Label4.Caption = "Procesando costes de paletización "
    DoEvents
    
    SQLinsert = "insert into cclineas4 (codvarie, codforfait, importe, kilos) values "
    
    'para cada palet confeccionado cogemos el coste por palet para prorratearlo entre los kilos de cada variedad/forfait
    SQL = "select distinct numpalet from cclineas2"
    
    pb1.visible = True
    nRegs = TotalRegistrosConsulta(SQL)
    CargarProgresNew pb1, CInt(nRegs)
    pb1.Value = 0
    DoEvents
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CadValues = ""
    
    While Not Rs.EOF
        IncrementarProgresNew pb1, 1
        
        Coste = CalcularCostePalet(CStr(DBLet(Rs!numpalet)))
        
        Sql2 = "select codvarie, codforfait, sum(pesoneto) as kilos from palets_variedad where numpalet = " & DBSet(Rs!numpalet, "N")
        If Variedades <> "" Then Sql2 = Sql2 & " and palets_variedad.codvarie in " & Variedades
        If Forfaits <> "" Then Sql2 = Sql2 & " and palets_variedad.codforfait in " & Forfaits
        Sql2 = Sql2 & " group by 1,2 order by 1,2 "
        
        'kilos totales del palet
        KilosTot = DevuelveValor("select sum(kilos) from (" & Sql2 & ") aaaaaaaa")
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs2.EOF
            Importe = 0
            If KilosTot <> 0 Then
                Importe = Round2(Coste * DBLet(Rs2!Kilos) / KilosTot, 2)
            End If
            CadValues = CadValues & "(" & DBSet(Rs2!codvarie, "N") & "," & DBSet(Rs2!codforfait, "T") & "," & DBSet(Importe, "N") & ","
            CadValues = CadValues & DBSet(Rs2!Kilos, "N") & "),"
            
            Rs2.MoveNext
        Wend
        
        Set Rs2 = Nothing
        
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    If CadValues <> "" Then
        ' quitamos la ultima coma
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
        conn.Execute SQLinsert & CadValues
        
'        SQL = ""
'        If txtCodigo(4).Text <> "" Then SQL = SQL & " and codvarie >= " & DBSet(txtCodigo(4).Text, "N")
'        If txtCodigo(5).Text <> "" Then SQL = SQL & " and codvarie <= " & DBSet(txtCodigo(5).Text, "N")
'        If SQL <> "" Then
'            Sql2 = "select codvarie from variedades where (1=1) " & SQL
            
            If Variedades <> "" Then
                conn.Execute "delete from cclineas4 where not codvarie in " & Variedades 'Sql2 & ")"
            End If
'        End If
        
    End If
    
    ProcesoCargaCampo = True
    
    Label4.visible = False
    Label5.visible = False
    pb1.visible = False
    
    Exit Function

eProcesoCargaCampo:
    Label4.visible = False
    Label5.visible = False
    pb1.visible = False


    MuestraError Err.Number, "Proceso Carga Campo", Err.Description
End Function

Private Sub CalcularCosteForfait(Forfait As String, Kilos As Long, ByRef CosteKilo As Currency)
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim TotalEnvases As String
Dim TotalCostes As String
Dim Valor As Currency
Dim CostesTotalEnvases As Currency
Dim CostesTotalCostes As Currency
Dim KilosCaja As Integer
Dim Cajas As Long
Dim CajaKilo As Integer

    On Error Resume Next

    CostesTotalEnvases = 0
    CostesTotalCostes = 0

    'total importes de envases para ese forfait
    SQL = "select sum(round(cantidad * "
    If vParamAplic.TipoPrecio = 0 Then 'precio medio ponderado
        SQL = SQL & " preciomp,4))"
    Else 'precio ultima compra
        SQL = SQL & " preciouc,4))"
    End If
    
    SQL = SQL & " from forfaits_envases, sartic where codforfait = " & DBSet(Forfait, "T")
    SQL = SQL & " and forfaits_envases.codartic = sartic.codartic"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    TotalEnvases = 0
    If Not Rs.EOF Then
        If Rs.Fields(0).Value > 0 Then TotalEnvases = Rs.Fields(0).Value
    End If
    Rs.Close
    Set Rs = Nothing
    
    KilosCaja = DevuelveValor("select kiloscaj from forfaits where codforfait = " & DBSet(Forfait, "T"))
    
    Cajas = 0
    If KilosCaja <> 0 Then
        Cajas = Round2(Kilos / KilosCaja, 0)
    End If
    
    'siempre por cajas
    CostesTotalEnvases = Round2(Cajas * TotalEnvases, 2)
    
'    'total costes para ese forfait
'    Sql = "select sum(importes) "
'    Sql = Sql & " from forfaits_costes where codforfait = " & DBSet(Forfait, "T")
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    TotalCostes = 0
'    If Not Rs.EOF Then
'        If Rs.Fields(0).Value > 0 Then TotalCostes = Rs.Fields(0).Value
'    End If
'    Rs.Close
'    Set Rs = Nothing
'
'    Cajakilo = DevuelveValor("select cajakilo from forfaits where codforfait = " & DBSet(Forfait, "T"))
'    If Cajakilo = 0 Then 'cajas
'        CostesTotalCostes = Round2(Cajas * TotalCostes, 2)
'    Else 'kilos
'        CostesTotalCostes = Round2(Kilos * TotalCostes, 2)
'    End If
    
    CosteKilo = Round2(CCur(CostesTotalEnvases), 4) '  + CCur(CostesTotalCostes), 4)
    
    
    If Err.Number <> 0 Then
        Err.Clear
    End If

End Sub






Private Sub cmdAceptarOld_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String


InicializarVbles
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
     '======== FORMULA  ====================================
    'Seleccionar registros de la empresa conectada
'    Codigo = "{" & tabla & ".codempre}=" & vEmpresa.codEmpre
'    If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
'    If Not AnyadirAFormula(cadSelect, Codigo) Then Exit Sub
        
    'D/H concepto
    cDesde = Trim(txtcodigo(12).Text)
    cHasta = Trim(txtcodigo(13).Text)
    nDesde = txtNombre(12).Text
    nHasta = txtNombre(13).Text
'    If Not (cDesde = "" And cHasta = "") Then
'        'Cadena para seleccion Desde y Hasta
'        Codigo = "{cccabdia.codcoste}"
'        TipCod = "N"
'        If Not PonerDesdeHasta1(cDesde, cHasta, nDesde, nHasta, "pDHConcepto= """) Then Exit Sub
'    End If
    CadConcepto = ""
    If cDesde <> "" Then CadConcepto = CadConcepto & " cccabdia.codcoste >= " & cDesde
    If cHasta <> "" Then
        If CadConcepto = "" Then
            CadConcepto = " cccabdia.codcoste <= " & cHasta
        Else
            CadConcepto = CadConcepto & " and cccabdia.codcoste<= " & cHasta
        End If
    End If
    If CadConcepto <> "" Then
        cadParam = cadParam & AnyadirParametroDH("pDHConcepto=""", cDesde, cHasta, nDesde, nHasta)
        numParam = numParam + 1
    End If

    'D/H forfait
    cDesde = Trim(txtcodigo(0).Text)
    cHasta = Trim(txtcodigo(1).Text)
    nDesde = txtNombre(0).Text
    nHasta = txtNombre(1).Text
'    If Not (cDesde = "" And cHasta = "") Then
'        'Cadena para seleccion Desde y Hasta
'        Codigo = "{cccaborden.codforfait}"
'        TipCod = "N"
'        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHForfait= """) Then Exit Sub
'    End If
    CadForfait = ""
    If cDesde <> "" Then CadForfait = CadForfait & " cclinorden3.codforfait >= '" & Trim(cDesde) & "'"
    If cHasta <> "" Then
        If CadForfait = "" Then
            CadForfait = " cclinorden3.codforfait <= '" & Trim(cHasta) & "'"
        Else
            CadForfait = CadForfait & " and cclinorden3.codforfait <= '" & Trim(cHasta) & "'"
        End If
    End If
    If CadForfait <> "" Then
        cadParam = cadParam & AnyadirParametroDH("pDHForfait=""", cDesde, cHasta, nDesde, nHasta)
        numParam = numParam + 1
    End If
    
    'D/H Variedad
    cDesde = Trim(txtcodigo(4).Text)
    cHasta = Trim(txtcodigo(5).Text)
    nDesde = txtNombre(4).Text
    nHasta = txtNombre(5).Text
'    If Not (cDesde = "" And cHasta = "") Then
'        'Cadena para seleccion Desde y Hasta
'        Codigo = "{cccaborden.codvarie}"
'        TipCod = "N"
'        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad= """) Then Exit Sub
'    End If
    CadVariedad = ""
    If cDesde <> "" Then CadVariedad = CadVariedad & " cclinorden3.codvarie >= " & cDesde
    If cHasta <> "" Then
        If CadVariedad = "" Then
            CadVariedad = " cclinorden3.codvarie <= " & cHasta
        Else
            CadVariedad = CadVariedad & " and cclinorden3.codvarie <= " & cHasta
        End If
    End If
    If CadVariedad <> "" Then
        cadParam = cadParam & AnyadirParametroDH("pDHVariedad=""", cDesde, cHasta, nDesde, nHasta)
        numParam = numParam + 1
    End If
    
    'D/H Fecha
    cDesde = Trim(txtcodigo(2).Text)
    cHasta = Trim(txtcodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{date(cccaborden.fechaini)}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    cadTABLA = " cccaborden INNER JOIN cclinorden3 ON cccaborden.codorden = cclinorden3.codorden "
    
    If HayRegParaInforme(cadTABLA, cadselect) Then
' Hemos cambiado de informe, me guardo como se grababa para el anterior informe
'        If CargarTemporal(cadTABLA, cadselect, CadConcepto, CadForfait, CadVariedad) Then
        If CargarTemporal2(cadTABLA, cadselect, CadConcepto, CadForfait, CadVariedad) Then
            cadFormula = "{tmpinfventas.codusu} = " & vUsu.Codigo
        
        
            CargarConceptos
        
            If option1(0).Value Then
                cadTitulo = "Informe de Costes por Fecha / Variedad"
                cadNombreRPT = "rCCInfCostesFec.rpt"
            Else
                cadTitulo = "Informe de Costes por Variedad / Fecha"
                cadNombreRPT = "rCCInfCostesVar.rpt"
            End If
        Else
            Exit Sub
        End If
        
        LlamarImprimir
    End If
    
End Sub

Private Sub CargarConceptos()
Dim SQL As String
Dim Rs As ADODB.Recordset

    On Error Resume Next

    SQL = "select codtipco, nomtipco from conftipo order by 1"

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        cadParam = cadParam & "pNombre" & Format(Rs!Codtipco, "0") & "=""" & DBLet(Rs!nomtipco) & """|"
        numParam = numParam + 1
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing


End Sub


Private Function CargarTemporal(cadTABLA As String, cadselect As String, CadConcepto As String, CadForfait As String, CadVariedad As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim CadValues As String
Dim ImpFactu As Currency
Dim PrecioFact As Currency
Dim Diferencia As Currency
Dim Importe As Currency
Dim Horas As Currency

Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset

Dim TotalKilos As Long

    On Error GoTo eCargarTemporal
    
    CargarTemporal = False
    Screen.MousePointer = vbHourglass
    Label4(6).visible = True
    DoEvents
    
    
'1 - CARGAMOS UNA TABLA INTERMEDIA PARA LOS GASTOS DIARIOS QUE HAY QUE PRORRATEAR ENTRE VARIEDADES
    BorrarTMPCostesDia
    CrearTMPCostesDia ""
  
    SQL = "insert into tmpCostesDia (fecha, codcoste, horas, impcoste) "
    SQL = SQL & " select cclindia2.fecha, cclindia2.codcoste, sum(cclindia2.horas) horas, round(sum(cclindia2.horas) * salarios.impsalar,2) as impcoste  "
    SQL = SQL & " from cclindia2 inner join salarios on cclindia2.codcateg = salarios.codcateg "
    SQL = SQL & " where date(fecha) in (select date(fecha) from cccabdia "
    If cadselect <> "" Or CadConcepto <> "" Then SQL = SQL & " where (1=1) "
    
    If cadselect <> "" Then SQL = SQL & " and " & Replace(cadselect, "cccaborden.fechaini", "cccabdia.fecha")
    If CadConcepto <> "" Then SQL = SQL & " and " & CadConcepto
    
    SQL = SQL & ") "
    SQL = SQL & " group by 1,2 "
    SQL = SQL & " order by 1,2 "
    
    conn.Execute SQL
    
'2 - CARGAMOS LOS KILOS DIARIOS POR VARIEDAD DE CADA TIPO (DESTRIO=1, PEQUEÑO=4, PODRIDO=3, COMERCIAL=0 Y 2)
    BorrarTMPKilos
    CrearTMPKilos ""
    
    ' introducimos lo que hay en entradas clasificadas
    SQL = "insert into tmpKilos (fecha, codvarie, destrio, pequeno, podrido, comercial) "
    SQL = SQL & _
        " select rclasifica.codvarie,rclasifica.fechaent, " & _
            " sum(if(codcalid in (select codcalid from rcalidad where tipcalid in (0,2)),rclasifica_clasif.kilosnet,0)) comercial, " & _
            " sum(if(codcalid in (select codcalid from rcalidad where tipcalid = 1),rclasifica_clasif.kilosnet,0)) destrio, " & _
            " sum(if(codcalid in (select codcalid from rcalidad where tipcalid = 3),rclasifica_clasif.kilosnet,0)) podrido, " & _
            " sum(if(codcalid in (select codcalid from rcalidad where tipcalid = 4),rclasifica_clasif.kilosnet,0)) pequeno " & _
          "  from rclasifica_clasif inner join rclasifica on rclasifica.numnotac = rclasifica_clasif.numnotac "
          
    If cadselect <> "" Or CadVariedad <> "" Then SQL = SQL & " where (1=1) "
    If cadselect <> "" Then SQL = SQL & " and " & Replace(cadselect, "cccaborden.fechaini", "rclasifica.fechaent")
    If CadVariedad <> "" Then SQL = SQL & " and " & Replace(CadVariedad, "cccaborden.codvarie", "rclasifica.codvarie")

    SQL = SQL & "  group by 1, 2 " & _
                "  order by 1, 2 "
    
    conn.Execute SQL
    
    ' introducimos lo que hay en el histórico de fruta
    SQL = "insert into tmpKilos (fecha, codvarie, destrio, pequeno, podrido, comercial) values "
    
    Sql2 = " select rhisfruta.codvarie,rhisfruta.fecalbar, " & _
                " sum(if(codcalid in (select codcalid from rcalidad where tipcalid in (0,2)),rhisfruta_clasif.kilosnet,0)) comercial, " & _
                " sum(if(codcalid in (select codcalid from rcalidad where tipcalid = 1),rhisfruta_clasif.kilosnet,0)) destrio, " & _
                " sum(if(codcalid in (select codcalid from rcalidad where tipcalid = 3),rhisfruta_clasif.kilosnet,0)) podrido, " & _
                " sum(if(codcalid in (select codcalid from rcalidad where tipcalid = 4),rhisfruta_clasif.kilosnet,0)) pequeno " & _
           "  from rhisfruta_clasif inner join rhisfruta on rhisfruta.numalbar = rhisfruta_clasif.numalbar "
          
    If cadselect <> "" Or CadVariedad <> "" Then Sql2 = Sql2 & " where (1=1) "
    If cadselect <> "" Then Sql2 = Sql2 & " and " & Replace(cadselect, "cccaborden.fechaini", "rhisfruta.fecalbar")
    If CadVariedad <> "" Then Sql2 = Sql2 & " and " & Replace(CadVariedad, "cccaborden.codvarie", "rhisfruta.codvarie")

    Sql2 = Sql2 & "  group by 1, 2 " & _
                  "  order by 1, 2 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CadValues = ""
    While Not Rs.EOF
        Sql3 = "select count(*) from tmpkilos where codvarie = " & DBSet(Rs!codvarie, "N") & " and fecha = " & DBSet(Rs!FecAlbar, "F")
        If TotalRegistros(Sql3) = 0 Then
            CadValues = CadValues & " (" & DBSet(Rs!FecAlbar, "F") & "," & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!Destrio, "N") & ","
            CadValues = CadValues & DBSet(Rs!pequeno, "N") & "," & DBSet(Rs!podrido, "N") & "," & DBSet(Rs!comercial, "N") & "),"
        Else
            Sql3 = "update tmpkilos set destrio = destrio + " & DBSet(Rs!Destrio, "N") & "," & _
                           " pequeno = pequeno + " & DBSet(Rs!pequeno, "N") & "," & _
                           " podrido = podrido + " & DBSet(Rs!podrido, "N") & "," & _
                           " comercial = comercial + " & DBSet(Rs!comercial, "N") & _
                       " where codvarie = " & DBSet(Rs!codvarie, "N") & " and fecha = " & DBSet(Rs!FecAlbar, "F")
            conn.Execute Sql3
        End If
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    If CadValues <> "" Then
        ' ejecutamos los insert quitando la ultima coma
        conn.Execute SQL & Mid(CadValues, 1, Len(CadValues) - 1)
    End If
    
'3 - REPARTIMOS POR VARIEDAD/FECHA LOS COSTES
    BorrarTMPResul
    CrearTMPResul ""
    
    CadValues = ""
    
    SQL = "select fecha, codvarie, (destrio + pequeno + podrido + comercial) kilos from tmpKilos "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        TotalKilos = DevuelveValor("select sum(destrio + pequeno + podrido + comercial) from tmpKilos where fecha = " & DBSet(Rs!fecha, "F"))
        
        Sql2 = "select codcoste, horas, impcoste from tmpCostesDia where fecha = " & DBSet(Rs!fecha, "F")
        Sql2 = Sql2 & " order by 1 "
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs2.EOF
            Importe = Round2(DBLet(Rs!Kilos, "N") * DBLet(Rs2!ImpCoste, "N") / TotalKilos, 2)
            Horas = Round2(DBLet(Rs!Kilos, "N") * DBLet(Rs2!Horas, "N") / TotalKilos, 2)
            
            CadValues = CadValues & "(" & DBSet(Rs!fecha, "F") & "," & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs2!codCoste, "N") & ","
            CadValues = CadValues & DBSet(Horas, "N") & "," & DBSet(Importe, "N") & "),"
            
            Rs2.MoveNext
        Wend
        
        Set Rs2 = Nothing
        
        Rs.MoveNext
    Wend
    Set Rs = Nothing
        
    If CadValues <> "" Then
        Sql3 = "insert into tmpResul (fecha, codvarie, codcoste, horas, impcoste) values "
        conn.Execute Sql3 & Mid(CadValues, 1, Len(CadValues) - 1)
    End If
    
'4 - INSERTAMOS EN LAS TABLAS TEMPORALES PARA EL REPORT
'    dejamos en tmpinformes y en tmpinfventas los datos prorrateados para el informe
    
    conn.Execute "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute "delete from tmpinfventas where codusu = " & vUsu.Codigo
                                            ' fecha,    codvarie, destrio, pequeno, podrido, comercial
    SQL = "insert into tmpinfventas  (codusu, fecalbar, numalbar, neto1,   neto2,   neto3,   neto4) "
    SQL = SQL & " select " & vUsu.Codigo & ", fecha,    codvarie, destrio, pequeno, podrido, comercial "
    SQL = SQL & " from tmpKilos "
    
    conn.Execute SQL
    
    SQL = "insert into tmpinformes (codusu, fecha1, codigo1, campo1, importe1, importe2) "
    SQL = SQL & " select " & vUsu.Codigo & ", fecha, codvarie, codcoste, horas, impcoste "
    SQL = SQL & " from tmpResul "

    conn.Execute SQL
    
    Screen.MousePointer = vbDefault
    Label4(6).visible = False
    DoEvents
    Set Rs = Nothing
    CargarTemporal = True
    Exit Function
    
eCargarTemporal:
    Screen.MousePointer = vbDefault
    Label4(6).visible = False
    MuestraError Err.Number, "Cargar Temporal", Err.Description
End Function


Private Function CargarTemporal2(cadTABLA As String, cadselect As String, CadConcepto As String, CadForfait As String, CadVariedad As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim Sql4 As String
Dim SQLinsert As String
Dim CadValues As String
Dim ImpFactu As Currency
Dim PrecioFact As Currency
Dim Diferencia As Currency
Dim Importe As Currency
Dim Horas As Currency
Dim TotalKilos As Long
Dim nRegs As Long
Dim NRegs2 As Long
Dim NConceptos As Long
Dim Kilos As Long
Dim ImporteVar As Currency

Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset


    On Error GoTo eCargarTemporal
    
    CargarTemporal2 = False
    Screen.MousePointer = vbHourglass
    Label4(6).visible = True
    DoEvents
    
    
'1 - CARGAMOS UNA TABLA INTERMEDIA PARA LOS GASTOS DIARIOS QUE HAY QUE PRORRATEAR ENTRE VARIEDADES
    BorrarTMPCostesDia
    CrearTMPCostesDia ""
  
    SQL = "insert into tmpCostesDia (fecha, codcoste, horas, impcoste) "
    SQL = SQL & " select fecha, codcoste, sum(horas), sum(impcoste) from ("
    SQL = SQL & " select cclindia1.fecha, cclindia1.codtraba, cclindia1.codcoste, sum(cclindia1.horas) horas, round(sum(cclindia1.horas) * straba.prhoracoste,2) as impcoste  "
    SQL = SQL & " from cclindia1 inner join straba on cclindia1.codtraba = straba.codtraba "
    SQL = SQL & " where date(fecha) in (select date(fecha) from cccabdia "
    SQL = SQL & " where (1=1) "
    
    If cadselect <> "" Then SQL = SQL & " and " & Replace(cadselect, "cccaborden.fechaini", "cccabdia.fecha")
    If CadConcepto <> "" Then SQL = SQL & " and " & CadConcepto
    
    SQL = SQL & ") "
    SQL = SQL & " group by 1,2,3  "
    SQL = SQL & " order by 1,2,3 )  aaaaa group by 1,2 order by 1,2 "
    
    conn.Execute SQL
    
'2 - CARGAMOS LA TABLA DE RESULTADOS
    BorrarTMPKilosRes
    CrearTMPKilosRes ""
    
    ' introducimos lo que hay en las ordenes de confeccion
    SQLinsert = "insert into tmpKilosRes (codorden, fecha, codvarie, codtipco, codcoste, kilos, importe) values "
    
    SQL = " select cccaborden.codorden, cccaborden.fechaini, cclinorden3.codvarie, forfaits.codtipco, sum(cclinorden3.kilosnet) kilos " & _
          "  from (cccaborden inner join cclinorden3 on cccaborden.codorden = cclinorden3.codorden) " & _
          "  inner join forfaits on cclinorden3.codforfait = forfaits.codforfait "
    SQL = SQL & " where (1=1) "
    
    If cadselect <> "" Then SQL = SQL & " and " & cadselect
    If CadVariedad <> "" Then SQL = SQL & " and " & CadVariedad
    If CadForfait <> "" Then SQL = SQL & " and " & CadForfait

    ' necesito el sql4 para añadirle la condicion del codigo de orden
    Sql4 = SQL

    SQL = SQL & "  group by 1, 2, 3, 4 " & _
                "  order by 1, 2, 3, 4 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Sql2 = "select codcoste, sum(if(importe is null, 0,importe)) importe from ("
        Sql2 = Sql2 & "select cclinorden1.codcoste, cclinorden1.codtraba, round(sum(cclinorden1.horas) * straba.prhoracoste,2) importe "
        Sql2 = Sql2 & " from cclinorden1 inner join straba on cclinorden1.codtraba = straba.codtraba "
        Sql2 = Sql2 & " where cclinorden1.codorden = " & DBSet(Rs!CodOrden, "N")
        Sql2 = Sql2 & " group by 1,2 "
        Sql2 = Sql2 & " order by 1,2 ) vconsulta "
        Sql2 = Sql2 & " group by 1 "
        Sql2 = Sql2 & " order by 1 "
        
        TotalKilos = DBLet(Rs!Kilos)
        nRegs = TotalRegistrosConsulta("select * from (" & Sql4 & " and cccaborden.codorden = " & DBSet(Rs!CodOrden, "N") & "  group by 1, 2, 3, 4  order by 1, 2, 3, 4 ) aaaconsulta ")
        NConceptos = TotalRegistrosConsulta(Sql2)
        
        ' si esta orden no ha tenido a nadie trabajando ( no se ha cargado nada desde el reloj )
        ' la cargo igualmente con concepto -1 ---> ESTO NO DEBERIA OCURRIR
        If NConceptos = 0 Then
            
            CadValues = "(" & DBSet(Rs!CodOrden, "N") & "," & DBSet(Rs!FechaIni, "F") & ","
            CadValues = CadValues & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!Codtipco, "N") & ","
            CadValues = CadValues & "0," ' sin concepto
            CadValues = CadValues & DBSet(TotalKilos, "N") & ","
            CadValues = CadValues & DBSet(0, "N") & "),"
        
            CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
            conn.Execute SQLinsert & CadValues
        
        Else
        
            CadValues = ""
            
            ' repartimos
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            While Not Rs2.EOF
                Kilos = Round2(TotalKilos / NConceptos, 0)
                Importe = Round2(DBLet(Rs2!Importe) / nRegs, 2)
            
                CadValues = CadValues & "(" & DBSet(Rs!CodOrden, "N") & "," & DBSet(Rs!FechaIni, "F") & ","
                CadValues = CadValues & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!Codtipco, "N") & ","
                CadValues = CadValues & DBSet(Rs2!codCoste, "N") & ","
                CadValues = CadValues & DBSet(Kilos, "N") & ","
                CadValues = CadValues & DBSet(Importe, "N") & "),"
                
                Rs2.MoveNext
            Wend
            
            Set Rs2 = Nothing
        
            If CadValues <> "" Then
                CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
                conn.Execute SQLinsert & CadValues
            End If
    
        End If
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    
'3 - INSERTAMOS EN LAS TABLAS TEMPORALES PARA EL REPORT
'    dejamos en tmpinformes y en tmpliquidacion los datos prorrateados para el informe
    
    conn.Execute "delete from tmpinformes where codusu = " & vUsu.Codigo
    
    
    SQL = "insert into tmpinformes (codusu, fecha1, codigo1, campo1, campo2, importe2, importe3) "
    SQL = SQL & " select " & vUsu.Codigo & ", fecha, codvarie, codtipco, codcoste, sum(kilos), sum(importe) "
    SQL = SQL & " from tmpkilosres "
    SQL = SQL & " group by 1, 2, 3, 4, 5"

    conn.Execute SQL
    
'4 - REPARTIMOS LOS COSTES GENERALES DIARIOS SEGUN LOS KILOS
'    dejamos en tmpinformes y en tmpliquidacion los datos prorrateados para el informe
    conn.Execute "delete from tmpliquidacion where codusu = " & vUsu.Codigo
    
                                                    'fecha,    variedad, codcoste, importe
    SQLinsert = "insert into tmpliquidacion (codusu, fechaini, codvarie, codsocio, importe) values "
    
    CadValues = ""
    
    SQL = "select fecha, codcoste, sum(impcoste) importe from tmpcostesdia group by 1,2 order by 1,2"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Sql2 = "select sum(importe2) as totalkilos from tmpinformes where codusu = " & vUsu.Codigo & " and fecha1 = " & DBSet(Rs!fecha, "F")
        TotalKilos = DevuelveValor(Sql2)
    
        Sql2 = "select codigo1, sum(importe2) as kilosvar from tmpinformes where codusu = " & vUsu.Codigo & " and fecha1 = " & DBSet(Rs!fecha, "F")
        Sql2 = Sql2 & " group by 1 "
        Sql2 = Sql2 & " order by 1 "
    
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs2.EOF
            ImporteVar = 0
            
            If TotalKilos <> 0 Then
                ImporteVar = Round2(Rs2!KilosVar * Rs!Importe / TotalKilos, 2)
            End If
            
            CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(Rs!fecha, "F") & "," & DBSet(Rs2!Codigo1, "N") & ","
            CadValues = CadValues & DBSet(Rs!codCoste, "N") & "," & DBSet(ImporteVar, "N") & "),"
            
            Rs2.MoveNext
        Wend
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    If CadValues <> "" Then
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
    
        conn.Execute SQLinsert & CadValues
    End If
    

'5 - CARGAMOS LA TABLA TMPINFVENTAS PARA PODER IMPRIMIR LOS CONCEPTOS DE CONFECCION ENCOLUMNADOS
'
    conn.Execute "delete from tmpinfventas where codusu = " & vUsu.Codigo
    
                                                    'fecha,  variedad, codcoste, kilos1,  importe1....
    SQL = "insert into tmpinfventas (codusu, fecalbar, numalbar, numcajas, codigo1, gastos1, codigo2, gastos2, codigo3, gastos3, codigo4, gastos4, codigo5, gastos5) "
    '                                       fecha,  variedad, codcoste,kilos, importe,.....
    SQL = SQL & "select " & vUsu.Codigo & ",fecha1, codigo1,  campo2,  0,0,0,0,0,0,0,0,0,0 from tmpinformes where codusu = " & vUsu.Codigo
    SQL = SQL & " group by 1,2,3,4 "
    
    conn.Execute SQL
    
    SQL = "select fecha1, codigo1, campo2, campo1, importe2, importe3 from tmpinformes where codusu = " & vUsu.Codigo
    SQL = SQL & " order by 1,2,3,4 "
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(Rs!fecha1, "F") & "," & DBSet(Rs!Codigo1, "N") & ","
        CadValues = CadValues & DBSet(Rs!campo1, "N") & ","
    
        SQL = "update tmpinfventas set codigo" & Format(DBLet(Rs!campo1), "0") & " = codigo" & Format(DBLet(Rs!campo1), "0") & " + " & DBSet(Rs!importe2, "N") ' kilos
        SQL = SQL & ", gastos" & Format(DBLet(Rs!campo1), "0") & " = gastos" & Format(DBLet(Rs!campo1), "0") & " + " & DBSet(Rs!importe3, "N") ' importe
        SQL = SQL & " where codusu = " & vUsu.Codigo
        SQL = SQL & " and fecalbar = " & DBSet(Rs!fecha1, "F") ' fecha
        SQL = SQL & " and numalbar = " & DBSet(Rs!Codigo1, "N") ' variedad
        SQL = SQL & " and numcajas = " & DBSet(Rs!campo2, "N") ' codcoste
                
        conn.Execute SQL
                
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    
'[Monica]22/08/2012: añadimos los costes de recoleccion de los partes de recoleccion
'6 - CARGAMOS LA TABLA TMPINFORMES PARA PODER IMPRIMIR LOS COSTES DE RECOLECCION
'
    conn.Execute "delete from tmpinformes where codusu = " & vUsu.Codigo
    
                                           'fecha,  variedad, kilos ,  importe)
    SQL = "insert into tmpinformes (codusu, fecha1, codigo1, importe1, importe2) "
    SQL = SQL & "select " & vUsu.Codigo & ", rpartes.fechapar, rpartes_trabajador.codvarie, sum(rpartes_trabajador.kilosrec), sum(rpartes_trabajador.importe) "
    SQL = SQL & " from rpartes, rpartes_trabajador "
    SQL = SQL & " where rpartes.nroparte = rpartes_trabajador.nroparte "
    
    If cadselect <> "" Then SQL = SQL & " and " & Replace(cadselect, "cccaborden.fechaini", "rpartes.fechapar")
    If CadVariedad <> "" Then SQL = SQL & " and " & Replace(CadVariedad, "cclinorden3", "rpartes_trabajador")
    
    SQL = SQL & " group by 1,2,3"
    SQL = SQL & " order by 1,2,3"
    
    
    conn.Execute SQL
    
    
    
    Screen.MousePointer = vbDefault
    Label4(6).visible = False
    DoEvents
    Set Rs = Nothing
    CargarTemporal2 = True
    Exit Function
    
eCargarTemporal:
    Screen.MousePointer = vbDefault
    Label4(6).visible = False
    MuestraError Err.Number, "Cargar Temporal", Err.Description
End Function





Private Sub BorrarTMPCostesDia()
On Error Resume Next
    conn.Execute " DROP TABLE IF EXISTS tmpCostesDia;"
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Function CrearTMPCostesDia(cadTABLA As String) As Boolean
'Crea una temporal donde insertara la clave primaria de las
'facturas erroneas al facturar
Dim SQL As String
    
    On Error GoTo ECrear
    
    CrearTMPCostesDia = False
    
    SQL = "CREATE /*TEMPORARY*/ TABLE tmpCostesDia ( "
    SQL = SQL & "fecha date NOT NULL default '0000-00-00', "
    SQL = SQL & "codcoste int(4) NOT NULL default 0,"
    SQL = SQL & "horas decimal(10,2) unsigned NOT NULL default '0.00',"
    SQL = SQL & "impcoste decimal(10,2) unsigned NOT NULL default '0.00') "
    conn.Execute SQL
     
     CrearTMPCostesDia = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPCostesDia = False
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpCostesDia;"
        conn.Execute SQL
    End If
End Function


Private Sub BorrarTMPKilos()
On Error Resume Next
    conn.Execute " DROP TABLE IF EXISTS tmpKilos;"
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Function CrearTMPKilos(cadTABLA As String) As Boolean
'Crea una temporal donde insertara la clave primaria de las
'facturas erroneas al facturar
Dim SQL As String
    
    On Error GoTo ECrear
    
    CrearTMPKilos = False
    
    SQL = "CREATE TEMPORARY TABLE tmpKilos ( "
    SQL = SQL & "fecha date NOT NULL default '0000-00-00', "
    SQL = SQL & "codvarie int(6) NOT NULL default 0,"
    SQL = SQL & "destrio int(6) NOT NULL default '0',"
    SQL = SQL & "pequeno int(6) NOT NULL default '0',"
    SQL = SQL & "podrido int(6) NOT NULL default '0',"
    SQL = SQL & "comercial int(6) NOT NULL default '0') "
    conn.Execute SQL
     
    CrearTMPKilos = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPKilos = False
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpKilos;"
        conn.Execute SQL
    End If
End Function


Private Sub BorrarTMPKilos2()
On Error Resume Next
    conn.Execute " DROP TABLE IF EXISTS tmpKilos;"
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Function CrearTMPKilos2(cadTABLA As String) As Boolean
'Crea una temporal donde insertara la clave primaria de las
'facturas erroneas al facturar
Dim SQL As String
    
    On Error GoTo ECrear
    
    CrearTMPKilos2 = False
    
    SQL = "CREATE TEMPORARY TABLE tmpKilos ( "
    SQL = SQL & "fecha date NOT NULL default '0000-00-00', "
    SQL = SQL & "codvarie int(6) NOT NULL default 0,"
    SQL = SQL & "codtipco int(6) NOT NULL default '0',"
    SQL = SQL & "kilos int(6) NOT NULL default '0')"
    conn.Execute SQL
     
    CrearTMPKilos2 = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPKilos2 = False
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpKilos;"
        conn.Execute SQL
    End If
End Function


Private Sub BorrarTMPKilosRes()
On Error Resume Next
    conn.Execute " DROP TABLE IF EXISTS tmpKilosRes;"
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Function CrearTMPKilosRes(cadTABLA As String) As Boolean
'Crea una temporal donde insertara la clave primaria de las
'facturas erroneas al facturar
Dim SQL As String
    
    On Error GoTo ECrear
    
    CrearTMPKilosRes = False
    
    SQL = "CREATE /*TEMPORARY*/ TABLE tmpKilosRes ( "
    SQL = SQL & "codorden int(7) unsigned NOT NULL,"
    SQL = SQL & "fecha date NOT NULL default '0000-00-00', "
    SQL = SQL & "codvarie int(6) NOT NULL default 0,"
    SQL = SQL & "codtipco int(6) NOT NULL default '0',"
    SQL = SQL & "codcoste int(6) NOT NULL default 0,"
    SQL = SQL & "kilos int(6) NOT NULL default '0',"
    SQL = SQL & "importe decimal(10,2) NOT NULL default '0')"
    conn.Execute SQL
     
    CrearTMPKilosRes = True
    
ECrear:
    If Err.Number <> 0 Then
        CrearTMPKilosRes = False
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpKilosRes;"
        conn.Execute SQL
    End If
End Function



Private Sub BorrarTMPResul()
On Error Resume Next
    conn.Execute " DROP TABLE IF EXISTS tmpResul;"
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Function CrearTMPResul(cadTABLA As String) As Boolean
'Crea una temporal donde insertara la clave primaria de las
'facturas erroneas al facturar
Dim SQL As String
    
    On Error GoTo ECrear
    
    CrearTMPResul = False
    
    SQL = "CREATE TEMPORARY TABLE tmpResul ( "
    SQL = SQL & "fecha date NOT NULL default '0000-00-00', "
    SQL = SQL & "codvarie int(6) NOT NULL default 0,"
    SQL = SQL & "codcoste int(4) NOT NULL default '0',"
    SQL = SQL & "horas dec(8,2) NOT NULL default '0.0',"
    SQL = SQL & "impcoste dec(8,2) NOT NULL default '0.0') "
    conn.Execute SQL
     
    CrearTMPResul = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPResul = False
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpResul;"
        conn.Execute SQL
    End If
End Function






Private Sub CmdBorMasivo_Click()
    If BorradoMasivo(CadBusqueda) Then
        MsgBox "Proceso realizado correctamente.", vbExclamation
        cmdCancel_Click (1)
    End If
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub


Private Sub CmdLeerFicheros_Click()
    
Dim Directorio As String

'    Directorio = "\\" & vConfig.SERVER & "\reloj\."
    Directorio = vParamAplic.PathFichadas & "\."
    If Dir(Directorio, vbArchive) <> "" Then
'        Directorio = "\\" & vConfig.SERVER & "\reloj"
        Directorio = vParamAplic.PathFichadas
        
        Continuar = ProcesarDirectorio(Directorio & "\", pb2, Label4(7), Label4(8))
        
    Else
        MsgBox "No existe el directorio de ticajes. Llame a Ariadna.", vbExclamation
        Continuar = False
    End If
    
End Sub

Private Sub CmdModMasiva_Click()
    If ModificacionMasiva(CadBusqueda) Then
        MsgBox "Proceso realizado correctamente.", vbExclamation
        cmdCancel_Click (1)
    End If
End Sub

Private Sub CmdProcesarDia_Click()

    Continuar = TotalRegistros("select count(*) from ccticajes")

    If Continuar Then
        If CargarTrabajadores Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
            cmdCancel_Click (1)
        End If
    Else
        MsgBox "Debe de leer ficheros previamente.", vbExclamation
    End If

End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case Opcionlistado
            Case 0:
                PonerFoco txtcodigo(4)
                
            Case 1:
                PonerFoco txtcodigo(6)
                
            Case 5:
                PonerFoco txtcodigo(14)
                
            Case 7:
                option1(2).Value = True
                
            Case 8: 'calculo de importe real de la conta
                PonerFoco txtcodigo(27)
                
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer, i As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me
    
   'IMAGES para busqueda
    For i = 0 To 10
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i

    Me.imgBuscar(12).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Me.imgBuscar(13).Picture = frmPpal.imgListImages16.ListImages(1).Picture

    For i = 25 To 26
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    
    '###Descomentar
'    CommitConexion
         
    Me.FrameCobros.visible = False
    Me.FrameCargaOrdenConf.visible = False
    Me.FrameCargaFichajes.visible = False
    Me.FrameModMasiva.visible = False
    Me.FrameBorradoMasivo.visible = False
    Me.FrameListadoCostes.visible = False
    Me.FrameBusqueda.visible = False
    Me.FrameCambios.visible = False
    Me.FrameCalculoImpReal.visible = False
    
    
    CargaCombo
    
    Select Case Opcionlistado
        Case 0
            ' informe de costes
            Tabla = "cccaborden"
            option1(0).Value = True
            FrameCobrosVisible True, H, W
        
            txtcodigo(2).Text = Month(Now)
            txtcodigo(3).Text = Month(Now)
            txtcodigo(22).Text = Year(Now)
            txtcodigo(24).Text = Year(Now)
        
            Combo1(1).ListIndex = 0
        
        Case 1
            ' carga automatica de costes
            Tabla = "cccaborden"
            Me.pb1.visible = False
            Me.Label4(10).visible = False
            FrameCargaOrdenConfVisible True, H, W
            
        Case 2
            ' Proceso de fichajes
            Tabla = "cccaborden"
            Me.pb2.visible = False
            Me.Label4(8).visible = False
            FrameCargaFichajesVisible True, H, W
            
        Case 3
            ' Modificacion masiva
            Tabla = "cctrabaconf"
            FrameModMasivaVisible True, H, W
            
        Case 4
            ' Borrado masivo
            Tabla = "cctrabaconf"
            txtcodigo(10).Text = 3
            FrameBorMasivoVisible True, H, W
            
        Case 5
            ' informe de cctrabaconf
            Tabla = "cctrabaconf"
            FrameListadoCostesVisible True, H, W
            
            Combo1(0).ListIndex = 0
            
        Case 6
            ' busqueda de una cadena
            FrameBusquedaVisible True, H, W
            
        Case 7 ' procesos varios inicio segun horario y eliminamos hora de almuerzo
            FrameCambiosVisible True, H, W
        
        Case 8 ' proceso de calculo de importes reales
            FrameCalculoImpRealVisible True, H, W
        
    End Select
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
'    Me.CmdCancel.Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CerrarConexionMultibase
End Sub





Private Sub frmBas_DatoSeleccionado(CadenaSeleccion As String)
'Form de Lineas de conceptos
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtcodigo(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub


Private Sub frmCla_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clases
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCon_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Conceptos
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFor_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de forfaits
    txtcodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmInc_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Incidencias
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMensFofaits_DatoSeleccionado(CadenaSeleccion As String)
Dim SQL As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        Forfaits = "T1.codvarie in (" & CadenaSeleccion & ")"
    Else
        Forfaits = "T1.codvarie = -1 "
    End If


End Sub

Private Sub frmMensForfaits_DatoSeleccionado(CadenaSeleccion As String)
Dim SQL As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        Forfaits = "(" & CadenaSeleccion & ")"
    Else
        Forfaits = ""
    End If
End Sub

Private Sub frmMensVariedad_DatoSeleccionado(CadenaSeleccion As String)
Dim SQL As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        Variedades = "(" & CadenaSeleccion & ")"
    Else
        Variedades = ""
    End If
End Sub

Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de trabajadores
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Variedades
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgFec_Click(Index As Integer)
'FEchas
    Dim esq, dalt As Long
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
       
    ' es desplega dalt i cap a la esquerra
    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + 420 + 30

    ' ***canviar l'index de imgFec pel 1r index de les imagens de buscar data***
    imgFec(2).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtcodigo(Index).Text <> "" Then frmC.NovaData = txtcodigo(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtcodigo(CByte(imgFec(2).Tag))
    ' ***************************
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1 'FORFAIT
            AbrirFrmForfaits (Index)
            
        Case 2 ' Linea de coste
            indCodigo = 8
            
            Set frmBas = New frmBasico
            frmBas.DatosADevolverBusqueda = "0|1|"
            frmBas.DeConsulta = True
            frmBas.CodigoActual = txtcodigo(indCodigo).Text
            frmBas.CadenaTots = "S|txtAux(0)|T|Código|800|;S|txtAux(1)|T|Descripción|3930|;"
            frmBas.CadenaConsulta = "SELECT cclinconf.codlinconf, cclinconf.nomlinconf "
            frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM cclinconf "
            frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
            frmBas.Tag1 = "Código|N|N|0|9999|cclinconf|codlinconf|00|S|"
            frmBas.Tag2 = "Descripción|T|N|||cclinconf|nomlinconf|||"
            frmBas.Maxlen1 = 2
            frmBas.Maxlen2 = 40
            frmBas.Tabla = "cclinconf"
            frmBas.CampoCP = "codlinconf"
            frmBas.Report = "rManCCLineasConf.rpt"
            frmBas.Caption = "Lineas de Confección"
            frmBas.Show vbModal
            Set frmBas = Nothing

        Case 4, 5 'VARIEDADES
            AbrirFrmVariedades (Index)
        
        Case 12, 13 ' CONCEPTO
            AbrirFrmConceptos (Index)
        
        Case 7, 8 ' CONCEPTO
            AbrirFrmConceptos (Index + 7)
        
        Case 9, 10 'TRABAJADORES
            indCodigo = Index + 7
            Set frmTra = New frmBasico
            AyudaTrabajadores frmTra, txtcodigo(indCodigo)
            Set frmTra = Nothing
            PonerFoco txtcodigo(indCodigo)
        
        Case 25, 26 ' CLASES
            AbrirFrmClases Index
        
        Case 3 'CONCEPTO
            AbrirFrmConceptos (Index + 24)
        Case 6 'CONCEPTO
            AbrirFrmConceptos (Index + 22)
        
        
        
    End Select
    PonerFoco txtcodigo(indCodigo)
End Sub




Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtcodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'14/02/2007 antes
'    KEYpress KeyAscii
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 4: KEYBusqueda KeyAscii, 4 'variedad desde
            Case 5: KEYBusqueda KeyAscii, 5 'variedad hasta
            Case 2: KEYFecha KeyAscii, 2 'fecha desde
            Case 3: KEYFecha KeyAscii, 3 'fecha hasta
            Case 12: KEYBusqueda KeyAscii, 12 'incidencia desde
            Case 13: KEYBusqueda KeyAscii, 12 'incidencia hasta
            Case 0: KEYBusqueda KeyAscii, 0 'cliente desde
            Case 1: KEYBusqueda KeyAscii, 1 'cliente hasta
            Case 6: KEYFecha KeyAscii, 6 'fecha de carga de ordenes de confeccion
            Case 21: KEYFecha KeyAscii, 21 'fecha hasta de carga de ordenes de confeccion
            Case 7: KEYFecha KeyAscii, 2 'fecha de cambio de cctrabaconf
            Case 8: KEYBusqueda KeyAscii, 2 'linea de coste
            
            ' Informe de costes
            Case 14: KEYBusqueda KeyAscii, 7 'concepto desde
            Case 15: KEYBusqueda KeyAscii, 8 'concepto hasta
            Case 16: KEYBusqueda KeyAscii, 9 'trabajador desde
            Case 17: KEYBusqueda KeyAscii, 10 'trabajador hasta
            Case 18: KEYFecha KeyAscii, 1 'fecha desde
            Case 19: KEYFecha KeyAscii, 4 'fecha hasta
            
            Case 25: KEYBusqueda KeyAscii, 25 'clase desde
            Case 26: KEYBusqueda KeyAscii, 26 'clase hasta
        
        
            Case 27: KEYBusqueda KeyAscii, 3 'concepto desde
            Case 28: KEYBusqueda KeyAscii, 6 'concepto hasta
        
        
        
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente
Dim SQL As String
    'Quitar espacios en blanco por los lados
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
        Case 6, 7, 18, 19, 21  'FECHAS antes tb 2,3
            If txtcodigo(Index).Text <> "" Then PonerFormatoFecha txtcodigo(Index)
            
        Case 4, 5 'VARIEDAD
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000000")
            
        Case 0, 1 'FORFAITS
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "forfaits", "nomconfe", "codforfait", "T")
                        
        Case 12, 13, 14, 15, 27, 28 'CONCEPTOS
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "ccconcostes", "nomcoste", "codcoste", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "0000")
                        
        Case 10 ' minutos de duracion de la tarea
            PonerFormatoEntero txtcodigo(10)
                        
        Case 8 ' linea de confeccion
            If txtcodigo(Index).Text <> "" Then
                SQL = DevuelveDesdeBDNew(cAgro, "cclinconf", "nomlinconf", "codlinconf", txtcodigo(Index).Text, "N")
                If SQL = "" Then
                    MsgBox "No existe la línea de confección. Revise.", vbExclamation
                    PonerFoco txtcodigo(Index)
                Else
                    txtNombre(Index).Text = SQL
                End If
            End If
            
        Case 16, 17 ' trabajadores
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "straba", "nomtraba", "codtraba", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000000")
        
        Case 9, 11, 20, 23 ' ponemos formato entero
            PonerFormatoEntero txtcodigo(Index)
        
        Case 2, 22, 3, 24
            PonerFormatoEntero txtcodigo(Index)
        
        Case 25, 26
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "clases", "nomclase", "codclase", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000")
        
        
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 7320 '6480
        Me.FrameCobros.Width = 6435
        W = Me.FrameCobros.Width
        H = Me.FrameCobros.Height
        
    End If
End Sub

Private Sub FrameCargaOrdenConfVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCargaOrdenConf.visible = visible
    If visible = True Then
        Me.FrameCargaOrdenConf.Top = -90
        Me.FrameCargaOrdenConf.Left = 0
        Me.FrameCargaOrdenConf.Height = 3630
        Me.FrameCargaOrdenConf.Width = 6435
        W = Me.FrameCargaOrdenConf.Width
        H = Me.FrameCargaOrdenConf.Height
        
    End If
End Sub


Private Sub FrameCargaFichajesVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCargaFichajes.visible = visible
    If visible = True Then
        Me.FrameCargaFichajes.Top = -90
        Me.FrameCargaFichajes.Left = 0
        Me.FrameCargaFichajes.Height = 3630
        Me.FrameCargaFichajes.Width = 6435
        W = Me.FrameCargaFichajes.Width
        H = Me.FrameCargaFichajes.Height
        
    End If
End Sub


Private Sub FrameModMasivaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameModMasiva.visible = visible
    If visible = True Then
        Me.FrameModMasiva.Top = -90
        Me.FrameModMasiva.Left = 0
        Me.FrameModMasiva.Height = 3630
        Me.FrameModMasiva.Width = 6435
        W = Me.FrameModMasiva.Width
        H = Me.FrameModMasiva.Height
        
    End If
End Sub


Private Sub FrameBorMasivoVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameBorradoMasivo.visible = visible
    If visible = True Then
        Me.FrameBorradoMasivo.Top = -90
        Me.FrameBorradoMasivo.Left = 0
        Me.FrameBorradoMasivo.Height = 3630
        Me.FrameBorradoMasivo.Width = 6435
        W = Me.FrameBorradoMasivo.Width
        H = Me.FrameBorradoMasivo.Height
    End If
End Sub


Private Sub FrameListadoCostesVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameListadoCostes.visible = visible
    If visible = True Then
        Me.FrameListadoCostes.Top = -90
        Me.FrameListadoCostes.Left = 0
        Me.FrameListadoCostes.Height = 5340
        Me.FrameListadoCostes.Width = 6435
        W = Me.FrameListadoCostes.Width
        H = Me.FrameListadoCostes.Height
    End If
End Sub

Private Sub FrameBusquedaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameBusqueda.visible = visible
    If visible = True Then
        Me.FrameBusqueda.Top = -90
        Me.FrameBusqueda.Left = 0
        Me.FrameBusqueda.Height = 4590
        Me.FrameBusqueda.Width = 6435
        W = Me.FrameBusqueda.Width
        H = Me.FrameBusqueda.Height
    End If
End Sub

Private Sub FrameCambiosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCambios.visible = visible
    If visible = True Then
        Me.FrameBusqueda.Top = -90
        Me.FrameBusqueda.Left = 0
        Me.FrameBusqueda.Height = 4170
        Me.FrameBusqueda.Width = 6435
        W = Me.FrameBusqueda.Width
        H = Me.FrameBusqueda.Height
    End If
End Sub


Private Sub FrameCalculoImpRealVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCalculoImpReal.visible = visible
    If visible = True Then
        Me.FrameCalculoImpReal.Top = -90
        Me.FrameCalculoImpReal.Left = 0
        Me.FrameCalculoImpReal.Height = 4740
        Me.FrameCalculoImpReal.Width = 6435
        W = Me.FrameCalculoImpReal.Width
        H = Me.FrameCalculoImpReal.Height
    End If
End Sub





Private Sub InicializarVbles()
    cadFormula = ""
    cadselect = ""
    
    CadVariedad = ""
    CadConcepto = ""
    CadForfait = ""

    cadParam = ""
    numParam = 0
End Sub

Private Function PonerDesdeHasta(codD As String, codH As String, nomD As String, nomH As String, param As String) As Boolean
'IN: codD,codH --> codigo Desde/Hasta
'    nomD,nomH --> Descripcion Desde/Hasta
'Añade a cadFormula y cadSelect la cadena de seleccion:
'       "(codigo>=codD AND codigo<=codH)"
' y añade a cadParam la cadena para mostrar en la cabecera informe:
'       "codigo: Desde codD-nomd Hasta: codH-nomH"

Dim devuelve As String
Dim devuelve2 As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(codD, codH, Codigo, TipCod)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    If TipCod <> "F" Then 'Fecha
        If Not AnyadirAFormula(cadselect, devuelve) Then Exit Function
    Else
        devuelve2 = CadenaDesdeHastaBD(codD, codH, Codigo, TipCod)
        If devuelve2 = "Error" Then Exit Function
        If Not AnyadirAFormula(cadselect, devuelve2) Then Exit Function
    End If
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            cadParam = cadParam & AnyadirParametroDH(param, codD, codH, nomD, nomH)
            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function

Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .ConSubInforme = True
        .Opcion = 0
        .Show vbModal
    End With
End Sub


Private Sub AbrirFrmForfaits(indice As Integer)
    indCodigo = indice
    Set frmFor = New frmManForfaits
    frmFor.DatosADevolverBusqueda = "0|1|"
    frmFor.DeConsulta = True
    frmFor.CodigoActual = txtcodigo(indCodigo)
    frmFor.Show vbModal
    Set frmVar = Nothing
End Sub


Private Sub AbrirFrmVariedades(indice As Integer)
    indCodigo = indice
    Set frmVar = New frmManVariedad
    frmVar.DatosADevolverBusqueda = "0|1|"
    frmVar.DeConsulta = True
    frmVar.CodigoActual = txtcodigo(indCodigo)
    frmVar.Show vbModal
    Set frmVar = Nothing
End Sub
 
Private Sub AbrirFrmConceptos(indice As Integer)
    indCodigo = indice
    Set frmCon = New frmCCManConcep
    frmCon.DatosADevolverBusqueda = "0|1|"
    frmCon.DeConsulta = True
    frmCon.CodigoActual = txtcodigo(indCodigo)
    frmCon.Show vbModal
    Set frmCon = Nothing
End Sub
 

Private Sub AbrirFrmIncidencias(indice As Integer)
    indCodigo = indice
    Set frmInc = New frmManInciden
    frmInc.DatosADevolverBusqueda = "0|1|"
    frmInc.DeConsulta = True
    frmInc.CodigoActual = txtcodigo(indCodigo)
    frmInc.Show vbModal
    Set frmInc = Nothing
End Sub
 
 
Private Sub AbrirFrmClases(indice As Integer)
    indCodigo = indice
    Set frmCla = New frmManClases
    frmCla.DatosADevolverBusqueda = "0|1|"
    frmCla.DeConsulta = True
    frmCla.CodigoActual = txtcodigo(indCodigo)
    frmCla.Show vbModal
    Set frmCla = Nothing
End Sub
 

Private Sub AbrirVisReport()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = cadFormula
        '.SoloImprimir = (Me.OptVisualizar(indFrame).Value = 1)
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        '##descomen
'        .MostrarTree = MostrarTree
'        .Informe = MIPATH & Nombre
'        .InfConta = InfConta
        '##
        
'        If NombreSubRptConta <> "" Then
'            .SubInformeConta = NombreSubRptConta
'        Else
'            .SubInformeConta = ""
'        End If
        '##descomen
'        .ConSubInforme = ConSubInforme
        '##
        .Opcion = ""
'        .ExportarPDF = (chkEMAIL.Value = 1)
        .Show vbModal
    End With
    
'    If Me.chkEMAIL.Value = 1 Then
'    '####Descomentar
'        If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
'    End If
    Unload Me
End Sub

Private Sub AbrirEMail()
    If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
End Sub


Private Function ModificarCategorias(Mens As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Categoria As Integer
Dim NumF As String
Dim Horas As String
Dim Rs As ADODB.Recordset

    On Error GoTo eModificarCategorias


    ModificarCategorias = False

    SQL = "delete from cclinorden2 where codorden in (select importe1 from tmpinformes where codusu = " & vUsu.Codigo & ")"
    conn.Execute SQL

    SQL = "select * from cclinorden1 where codorden in (select importe1 from tmpinformes where codusu = " & vUsu.Codigo & ")"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    While Not Rs.EOF
    
        SQL = "select codcateg from straba where codtraba = " & DBSet(Rs!codtraba, "N")
        Categoria = DevuelveValor(SQL)
    
        SQL = "select count(*) from cclinorden2 where codorden = " & DBSet(Rs!CodOrden, "N")
        SQL = SQL & " and codcoste = " & DBSet(Rs!codCoste, "N")
        SQL = SQL & " and codcateg = " & DBSet(Categoria, "N")
        
        If TotalRegistros(SQL) = 0 Then
            NumF = SugerirCodigoSiguienteStr("cclinorden2", "numlinea", "codorden = " & DBSet(Rs!CodOrden, "N"))
        
            Sql2 = "insert into cclinorden2 (codorden,numlinea,codcoste,codcateg,horas) values ("
            Sql2 = Sql2 & DBSet(Rs!CodOrden, "N") & "," & DBSet(NumF, "N") & "," & DBSet(Rs!codCoste, "N") & ","
            Sql2 = Sql2 & DBSet(Categoria, "N") & "," & DBSet(Rs!Horas, "N") & ")"
            
            conn.Execute Sql2
        
        Else
            Sql2 = "update cclinorden2 set horas = horas + " & DBSet(Rs!Horas, "N")
            Sql2 = Sql2 & " where codorden = " & DBSet(Rs!CodOrden, "N")
            Sql2 = Sql2 & " and codcoste = " & DBSet(Rs!codCoste, "N")
            Sql2 = Sql2 & " and codcateg = " & DBSet(Categoria, "N")
                
            conn.Execute Sql2
        End If
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    ModificarCategorias = True
    Exit Function
    
eModificarCategorias:
    Mens = Mens & vbCrLf & Err.Description
End Function



Public Function ProcesarDirectorio(nomDir As String, ByRef pb1 As ProgressBar, ByRef Label1 As label, ByRef Label2 As label) As Boolean
Dim NF As Long
Dim Cad As String
Dim i As Integer
Dim Longitud As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim Numreg As Long
Dim SQL As String
Dim SQL1 As String
Dim Total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean
Dim NomFic As String
Dim Nota As String
Dim Linea As Integer


    On Error GoTo eProcesarDirectorio

    ProcesarDirectorio = False
    b = True
    
    ' Muestra los nombres en C:\ que representan directorios.
    NomFic = Dir(nomDir & "HU*.txt", vbArchive) ' Recupera la primera entrada.
    
    If NomFic = "" Then
        MsgBox "No hay ficheros a procesar.", vbExclamation
        ProcesarDirectorio = True
        Exit Function
    End If
    
    Label1.visible = True
    Label2.visible = True
    pb1.visible = True
    DoEvents
    Do While NomFic <> "" And b   ' Inicia el bucle.
       ' Ignora el directorio actual y el que lo abarca.
       If NomFic <> "." And NomFic <> ".." Then
          ' Realiza una comparación a nivel de bit para asegurarse de que MiNombre es un directorio.
'[Monica]13/02/2013: no hay qu comprobar que es un fichero
'          If (GetAttr(nomDir & NomFic) And vbArchive) = vbArchive Then
            
            ' una vez pocesado lo renombramos
            Name nomDir & NomFic As nomDir & Replace(NomFic, ".txt", ".dat")
            
            NF = FreeFile
            
            Open nomDir & Replace(NomFic, ".txt", ".dat") For Input As #NF
            
            Line Input #NF, Cad
            
            Label1.Caption = "Procesando Fichero: " & NomFic
            Longitud = FileLen(nomDir & Replace(NomFic, ".txt", ".dat"))
            
            CargarProgres pb1, CInt(Longitud)
            
            Linea = 1
            If Cad <> "" Then
                b = ProcesarFichero(NF, Cad, pb1, Label1, Label2)
            End If
            
            Close #NF
          
'          End If   ' solamente si representa un directorio.
       End If
       NomFic = Dir   ' Obtiene siguiente entrada.
    Loop
    
    ProcesarDirectorio = b
    
    pb1.visible = False
    Label1.Caption = ""
    Label2.Caption = ""
    Exit Function
    
eProcesarDirectorio:
    MuestraError Err.Number, "Procesar Directorio", Err.Description
End Function


Private Function ProcesarFichero(NF As Long, Cad As String, ByRef pb1 As ProgressBar, ByRef Label1 As label, ByRef Label2 As label) As Boolean
'Dim b As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Mens As String
Dim NumLinea As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim RSaux As ADODB.Recordset
Dim Aux As String

Dim i As Long
Dim vValues As String

    On Error GoTo eProcesarFichero

    ProcesarFichero = False
    
    i = 0
    vValues = ""
    
    While Not EOF(NF)
       i = i + 1
        
       IncrementarProgres pb1, Len(Cad)
       Label2.Caption = "Linea " & i
        'Me.Refresh
       DoEvents
        
       If Cad <> "" Then
        
            '[Monica]22/10/2012: si fichan con un codigo de trabajador que no existe me pone actividad = 0, tambien tengo que descartarlas
            If Mid(Cad, 1, 1) <> "F" And Mid(Cad, 25, 2) <> "00" Then
       
                Aux = Mid(Cad, 6, 5)  'trab
                Aux = Aux & "," & Mid(Cad, 13, 2)  'mes
                Aux = Aux & "," & Mid(Cad, 15, 2)  'dia
                Aux = Aux & "," & Mid(Cad, 17, 2)  'hora
                Aux = Aux & "," & Mid(Cad, 19, 2)  'min
                Aux = Aux & "," & CDec("&H" & Mid(Cad, 25, 2))  ' actividad me viene en hexadecimal
                
                '[Monica]07/01/2013: el año lo tengo q guardar también para procesar
                Aux = Aux & "," & Mid(Cad, 11, 2)  'año
                
                vValues = vValues & "(" & Aux & "),"
       
            End If
            
       End If
        
       Line Input #NF, Cad
    Wend
    
    If Cad <> "" Then
        If Mid(Cad, 1, 1) <> "F" Then
        ' ultimo registro
            Aux = Mid(Cad, 6, 5)  'trab
            Aux = Aux & "," & Mid(Cad, 13, 2)  'mes
            Aux = Aux & "," & Mid(Cad, 15, 2)  'dia
            Aux = Aux & "," & Mid(Cad, 17, 2)  'hora
            Aux = Aux & "," & Mid(Cad, 19, 2)  'min
            Aux = Aux & "," & CDec("&H" & Mid(Cad, 25, 2)) ' actividad me viene en hexadecimal
            
            '[Monica]07/01/2013: el año lo tengo q guardar también para procesar
            Aux = Aux & "," & Mid(Cad, 11, 2)  'año
            
            vValues = vValues & "(" & Aux & "),"
            
        End If
    End If
    
    Close #NF
        
    If vValues <> "" Then
        SQL = "insert into ccticajes (codtraba, mes, dia, hora, minuto, actividad, anyo) values "
        SQL = SQL & Mid(vValues, 1, Len(vValues) - 1)
        
        conn.Execute SQL
    End If

    ProcesarFichero = True
    Exit Function

eProcesarFichero:
    If Err.Number <> 0 Then
        ProcesarFichero = False
        MsgBox "Error en Procesar Linea " & Err.Description, vbExclamation
    End If
End Function


Public Function ProcesarDirectorioBusqueda(nomDir As String, ByRef pb1 As ProgressBar, ByRef Label1 As label, ByRef Label2 As label) As Boolean
Dim NF As Long
Dim Cad As String
Dim i As Integer
Dim Longitud As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim Numreg As Long
Dim SQL As String
Dim SQL1 As String
Dim Total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean
Dim NomFic As String
Dim Nota As String
Dim Linea As Integer


    On Error GoTo eProcesarDirectorio

    ProcesarDirectorioBusqueda = False
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    b = True
    
    ' Muestra los nombres en C:\ que representan directorios.
    NomFic = Dir(nomDir & "HU*.*")  ' Recupera la primera entrada.
    
    If NomFic = "" Then
        MsgBox "No hay ficheros a procesar.", vbExclamation
        ProcesarDirectorioBusqueda = True
        Exit Function
    End If
    
    Label1.visible = True
    Label2.visible = True
    pb1.visible = True
    DoEvents
    Do While NomFic <> "" And b   ' Inicia el bucle.
       ' Ignora el directorio actual y el que lo abarca.
       If NomFic <> "." And NomFic <> ".." Then
          ' Realiza una comparación a nivel de bit para asegurarse de que MiNombre es un directorio.
          If (GetAttr(nomDir & NomFic) And vbArchive) = vbArchive Then
          'If Dir(nomDir & NomFic, vbArchive) <> "" Then
            
            NF = FreeFile
            
            Open nomDir & NomFic For Input As #NF
            
            Line Input #NF, Cad
            
            Label1.Caption = "Procesando Fichero Busqueda: " & NomFic
            Longitud = FileLen(nomDir & NomFic)
            
            CargarProgres pb3, CInt(Longitud)
            
            Linea = 1
            If Cad <> "" Then
                b = ProcesarFicheroBusqueda(NF, Cad, pb3, Label1, Label2)
            End If
            
            Close #NF
          
          End If   ' solamente si representa un directorio.
       End If
       NomFic = Dir   ' Obtiene siguiente entrada.
    Loop
    
    ProcesarDirectorioBusqueda = b
    
    pb1.visible = False
    Label1.Caption = ""
    Label2.Caption = ""
    DoEvents
    Exit Function
    
eProcesarDirectorio:
    MuestraError Err.Number, "Procesar Directorio Busqueda", Err.Description
End Function

Private Function ProcesarFicheroBusqueda(NF As Long, Cad As String, ByRef pb1 As ProgressBar, ByRef Label1 As label, ByRef Label2 As label) As Boolean
'Dim b As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Mens As String
Dim NumLinea As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim RSaux As ADODB.Recordset
Dim Aux As String

Dim i As Long
Dim vValues As String

    On Error GoTo eProcesarFicheroBusqueda

    ProcesarFicheroBusqueda = False
    
    i = 0
    vValues = ""
    
    While Not EOF(NF)
       i = i + 1
        
       IncrementarProgres pb1, Len(Cad)
       Label2.Caption = "Linea " & i
        'Me.Refresh
       DoEvents
        
       Aux = ""
        
       If Cad <> "" Then
            If Mid(Cad, 1, 1) <> "F" Then
                If (CLng(Mid(Cad, 6, 5)) = CLng(ComprobarCero(txtcodigo(23).Text)) And txtcodigo(23).Text <> "") And _
                   ((CInt(Mid(Cad, 13, 2)) = CInt(ComprobarCero(txtcodigo(9).Text)) And txtcodigo(9).Text <> "") Or txtcodigo(9).Text = "") And _
                   ((CInt(Mid(Cad, 15, 2)) = CInt(ComprobarCero(txtcodigo(11).Text)) And txtcodigo(11).Text <> "") Or txtcodigo(11).Text = "") And _
                   ((CInt(Mid(Cad, 17, 2)) = CInt(ComprobarCero(txtcodigo(20).Text)) And txtcodigo(20).Text <> "") Or txtcodigo(20).Text = "") Then   ' trabajador
                    Aux = Mid(Cad, 6, 5)  'trab
                    Aux = Aux & "," & Mid(Cad, 13, 2)  'mes
                    Aux = Aux & "," & Mid(Cad, 15, 2)  'dia
                    Aux = Aux & "," & Mid(Cad, 17, 2)  'hora
                    Aux = Aux & "," & Mid(Cad, 19, 2)  'min
                    Aux = Aux & "," & CDec("&H" & Mid(Cad, 25, 2))  ' actividad me viene en hexadecimal
                End If
           End If
       End If
       
       If Aux <> "" Then
            vValues = vValues & "(" & vUsu.Codigo & "," & Trim(Aux) & "),"
       End If
        
       Line Input #NF, Cad
    Wend
    
    If Cad <> "" Then
        If Mid(Cad, 1, 1) <> "F" Then
            If (CLng(Mid(Cad, 6, 5)) = CLng(ComprobarCero(txtcodigo(23).Text)) And txtcodigo(23).Text <> "") And _
               ((CInt(Mid(Cad, 13, 2)) = CInt(ComprobarCero(txtcodigo(9).Text)) And txtcodigo(9).Text <> "") Or txtcodigo(9).Text = "") And _
               ((CInt(Mid(Cad, 15, 2)) = CInt(ComprobarCero(txtcodigo(11).Text)) And txtcodigo(11).Text <> "") Or txtcodigo(11).Text = "") And _
               ((CInt(Mid(Cad, 17, 2)) = CInt(ComprobarCero(txtcodigo(20).Text)) And txtcodigo(20).Text <> "") Or txtcodigo(20).Text = "") Then   ' trabajador
             
                 Aux = Mid(Cad, 6, 5)  'trab
                 Aux = Aux & "," & Mid(Cad, 13, 2)  'mes
                 Aux = Aux & "," & Mid(Cad, 15, 2)  'dia
                 Aux = Aux & "," & Mid(Cad, 17, 2)  'hora
                 Aux = Aux & "," & Mid(Cad, 19, 2)  'min
                 Aux = Aux & "," & CDec("&H" & Mid(Cad, 25, 2))  ' actividad me viene en hexadecimal
             End If
         End If
        
         If Aux <> "" Then
             vValues = vValues & "(" & vUsu.Codigo & "," & Aux & "),"
         End If
    End If
    
    Close #NF
        
    If vValues <> "" Then                      'codtraba, mes,      dia,      hora,     minuto,   actividad
        SQL = "insert into tmpinformes (codusu, codigo1,  importe1, importe2, importe3, importe4, importe5) values "
        SQL = SQL & Mid(vValues, 1, Len(vValues) - 1)
        
        conn.Execute SQL
    End If

    ProcesarFicheroBusqueda = True
    Exit Function

eProcesarFicheroBusqueda:
    If Err.Number <> 0 Then
        ProcesarFicheroBusqueda = False
        MsgBox "Error en Procesar Linea Busqueda " & Err.Description, vbExclamation
    End If
End Function




Private Function CargarTrabajadores() As Boolean
Dim SQL As String
Dim TrabaAnt As String
Dim FecIni As String
Dim FecFin As String
Dim Sql3 As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim vValues As String
Dim Linea As Integer
Dim Actividad As Integer
Dim nRegs As Long
Dim Trabajador As String


    On Error GoTo eCargarTrabajadores

    If HayTarjetasSinTrabajador Then
        CargarTrabajadores = False
        Exit Function
    End If
    
    
    Sql3 = "insert ignore into cctrabaconf (codtraba,fechaini,fechafin,codlinconf,codcoste,procesado,sinacabar) values "


    '[Monica]24/05/2012: realmente lo que tenemos en ccticajes.codtraba es el idtarjeta del fichero
    SQL = "select distinct codtraba from ccticajes group by codtraba order by codtraba"
        
    nRegs = TotalRegistrosConsulta(SQL)
    
    If nRegs <> 0 Then
        pb2.visible = True
        Label4(7).visible = True
        Label4(8).visible = True
        DoEvents
        CargarProgres pb2, CInt(nRegs)
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        ' para cada trabajador vemos todos los registros que tiene ordenados por fecha
        Sql2 = "select * from ccticajes where codtraba = " & DBSet(Rs!codtraba, "N")
        '[Monica]23/05/2012: si la actividad es entrada no hacemos nada
        Sql2 = Sql2 & " and actividad <> 1 "
        Sql2 = Sql2 & " order by mes, dia, hora, minuto "
        
        '[Monica]24/05/2012: sacamos el trabajador asociado a la tarjeta
        Trabajador = DevuelveValor("select codtraba from straba where idtarjeta = " & DBSet(Rs!codtraba, "N"))
        
        IncrementarProgres pb2, 1
        Label4(7).Caption = "Procesando trabajadores"
        Label4(8).Caption = "Trabajador " & Trabajador 'RS!codtraba
        DoEvents
        vValues = ""
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs2.EOF                                                       '[Monica]07/01/2013:antes Year(Now)
            FecIni = Format(Rs2!dia, "00") & "/" & Format(Rs2!mes, "00") & "/" & Format(Rs2!Anyo, "0000") & " " & Format(Rs2!hora, "00") & ":" & Format(Rs2!minuto, "00") & ":00"
            FecFin = FecIni
            
            If Rs2!Actividad <> 2 Then
                If HaySinAcabar(Trabajador) Then
                    SQL = "update cctrabaconf set sinacabar = 0, fechafin = " & DBSet(FecFin, "FH")
                    SQL = SQL & " where codtraba = " & DBSet(Trabajador, "N") & " and sinacabar = 1 "
                    conn.Execute SQL
                End If
'23/05/2012
'               Linea = Mid(Format(Rs2!Actividad, "00"), 1, 1)
'               Actividad = Mid(Format(Rs2!Actividad, "00"), 2, 1)
                
                Linea = 0
                Actividad = Rs2!Actividad
                
                ' insertamos nuestro codigo de coste
                'Actividad = DevuelveValor("select codcoste from ccconcostes where codincid = " & DBSet(Actividad, "N"))
                '[Monica]23/05/2012
                Linea = DevuelveValor("select codincid from ccconcostes where codcoste = " & DBSet(Actividad, "N"))
                
                vValues = Replace(vValues, "xxxxxxxxxx", DBSet(FecFin, "FH"))
                vValues = vValues & "(" & DBSet(Trabajador, "N") & "," & DBSet(FecIni, "FH") & ",xxxxxxxxxx," & DBSet(Linea, "N") & "," & DBSet(Actividad, "N") & ",0,0),"
                    
            Else
                If HaySinAcabar(Trabajador) Then
                    SQL = "update cctrabaconf set sinacabar = 0, fechafin = " & DBSet(FecFin, "FH")
                    SQL = SQL & " where codtraba = " & DBSet(Trabajador, "N") & " and sinacabar = 1 "
                    conn.Execute SQL
                Else
                    vValues = Replace(vValues, "xxxxxxxxxx", DBSet(FecFin, "FH"))
                End If
            End If
            
            Rs2.MoveNext
        Wend
    
        If vValues <> "" Then
            vValues = Replace(vValues, "xxxxxxxxxx", DBSet(FecIni, "FH"))
            vValues = Mid(vValues, 1, Len(vValues) - 1)
            conn.Execute Sql3 & vValues
            
            '
            Sql2 = "delete from ccticajes where codtraba = " & DBSet(Rs!codtraba, "N")
            conn.Execute Sql2
'de momento
'            '[Monica]24/05/2012: borramos las fichadas en que desde y hasta son iguales
'            Sql2 = "delete from cctrabaconf where codtraba = " & DBSet(Trabajador, "N")
'            Sql2 = Sql2 & " and fechaini = fechafin"
'            conn.Execute Sql2
            
            ' Puede que venga en el fichero siguiente, lo marcamos
            Sql2 = "update cctrabaconf set sinacabar = 1 where codtraba = " & DBSet(Trabajador, "N")
            Sql2 = Sql2 & " and fechaini = fechafin and fechaini = " & DBSet(FecIni, "FH")
            conn.Execute Sql2
            
        Else
            Sql2 = "delete from ccticajes where codtraba = " & DBSet(Rs!codtraba, "N")
            conn.Execute Sql2
        
        End If
        
        Rs.MoveNext
    Wend

    CargarTrabajadores = True
    
    pb2.visible = False
    Label4(7).visible = False
    Label4(8).visible = False
    
    Exit Function

eCargarTrabajadores:
    pb2.visible = False
    Label4(7).visible = False
    Label4(8).visible = False
    MuestraError Err.Number, "Cargar Trabajadores", Err.Description
End Function

Private Function HaySinAcabar(Trabajador As String) As Boolean
Dim SQL As String

    SQL = "select count(*) from cctrabaconf where codtraba = " & DBSet(Trabajador, "N")
    SQL = SQL & " and sinacabar = 1"
    HaySinAcabar = (TotalRegistros(SQL) <> 0)

End Function



Private Function HayTarjetasSinTrabajador() As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim cadResul As String

    On Error GoTo eHayTarjetasSinTrabajador
    
    HayTarjetasSinTrabajador = True
    
    SQL = "select distinct codtraba from ccticajes where not codtraba in (select idtarjeta from straba) "
    If DevuelveValor(SQL) <> 0 Then
    
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        cadResul = ""
        While Not Rs.EOF
            cadResul = cadResul & DBLet(Rs!codtraba, "N") & ","
        
            Rs.MoveNext
        Wend
        
        If cadResul <> "" Then
            cadResul = "Las siguientes tarjetas no tienen trabajador asociado en Ariagro: " & vbCrLf & vbCrLf & Mid(cadResul, 1, Len(cadResul) - 1)
            MsgBox cadResul, vbExclamation
            HayTarjetasSinTrabajador = True
        Else
            HayTarjetasSinTrabajador = False
        End If
        
        Set Rs = Nothing
    Else
        HayTarjetasSinTrabajador = False
    End If
    Exit Function
    
eHayTarjetasSinTrabajador:
    MuestraError Err.Number, "Tarjetas sin Trabajador", Err.Description
End Function


Private Function ModificacionMasiva(cSelect As String) As Boolean
Dim SQL As String
Dim Sql2 As String

    On Error GoTo eModificacionMasiva
    
    If Not DatosOk Then Exit Function
    
    
    conn.BeginTrans
    
    Sql2 = "update cctrabaconf "
    
    If txtcodigo(7).Text <> "" Then
        SQL = Sql2 & " set fechaini = concat(concat(" & DBSet(txtcodigo(7).Text, "F") & ", ' '),time(fechaini)) "
        SQL = SQL & ", fechafin = concat(concat(" & DBSet(txtcodigo(7).Text, "F") & ", ' '),time(fechafin)) "
        SQL = SQL & " where " & cSelect
        SQL = SQL & " and procesado = 0 "
        conn.Execute SQL
    End If
    
    If txtcodigo(8).Text <> "" Then
        SQL = Sql2 & " set codlinconf = " & DBSet(txtcodigo(8).Text, "N")
        SQL = SQL & " where " & cSelect
        SQL = SQL & " and procesado = 0 "
        
        conn.Execute SQL
    End If
    
    conn.CommitTrans
    ModificacionMasiva = True
    Exit Function
    
eModificacionMasiva:
    conn.RollbackTrans
    ModificacionMasiva = False
End Function


Private Function BorradoMasivo(cSelect As String) As Boolean
Dim SQL As String
Dim Sql2 As String

    On Error GoTo eBorradoMasivo
    
    If Not DatosOk Then Exit Function
    
    Sql2 = "delete from cctrabaconf "
    Sql2 = Sql2 & " where hour(timediff(fechafin, fechaini)) = 0 and"
    Sql2 = Sql2 & " Minute (timediff(FechaFin, FechaIni)) <= " & DBSet(txtcodigo(10).Text, "N")
    Sql2 = Sql2 & " and procesado = 0 and sinacabar = 0 "
    
    conn.Execute Sql2
    
    BorradoMasivo = True
    Exit Function
    
eBorradoMasivo:
    BorradoMasivo = False
End Function





Private Function DatosOk() As Boolean
Dim b As Boolean
Dim HMes As Integer
Dim HAnyo As Integer

    On Error GoTo EDatosOK

    DatosOk = False
    b = True
    
    Select Case Opcionlistado
        Case 0 ' Informe de costes
            If txtcodigo(2).Text = "" Or txtcodigo(22).Text = "" Or txtcodigo(3).Text = "" Or txtcodigo(24).Text = "" Then
                MsgBox "El periodo de fechas debe de estar definido. Revise."
                PonerFoco txtcodigo(2)
                b = False
            Else
                ' comprobamos q la fecha desde es inferior o igual a fecha hasta
                FDesde = CDate("01" & "/" & Format(txtcodigo(2).Text, "00") & "/" & Format(txtcodigo(22).Text, "0000"))
                HMes = txtcodigo(3).Text
                HAnyo = txtcodigo(24).Text
                If HMes = 12 Then
                    HMes = 1
                    HAnyo = HAnyo + 1
                Else
                    HMes = HMes + 1
                End If
                FHasta = CDate("01" & "/" & Format(HMes, "00") & "/" & Format(HAnyo, "0000"))
                FHasta = DateAdd("d", (-1), FHasta)
            End If
    
    
        Case 3 ' modificacion masiva
            If txtcodigo(7).Text = "" And txtcodigo(8).Text = "" Then
                MsgBox "Debe introducir algun valor en los campos. Revise.", vbExclamation
                PonerFoco txtcodigo(7)
                b = False
            End If
            
        Case 4 ' borrado masivo
            If txtcodigo(10).Text = "" Then
                MsgBox "Debe introducir un valor en el campo minutos. Revise.", vbExclamation
                PonerFoco txtcodigo(10)
                b = False
            Else
                If CInt(txtcodigo(10).Text) > 59 Then
                    MsgBox "El tiempo no puede ser superior a una hora. Revise.", vbExclamation
                    PonerFoco txtcodigo(10)
                    b = False
                End If
            End If
    
        Case 8 ' cálculo de importes de reales
            If txtcodigo(31).Text = "" Or txtcodigo(32).Text = "" Or txtcodigo(35).Text = "" Or txtcodigo(36).Text = "" Then
                MsgBox "El periodo de fechas debe de estar definido. Revise."
                PonerFoco txtcodigo(31)
                b = False
            Else
                ' comprobamos q la fecha desde es inferior o igual a fecha hasta
                FDesde = CDate("01" & "/" & Format(txtcodigo(31).Text, "00") & "/" & Format(txtcodigo(32).Text, "0000"))
                HMes = txtcodigo(35).Text
                HAnyo = txtcodigo(36).Text
                If HMes = 12 Then
                    HMes = 1
                    HAnyo = HAnyo + 1
                Else
                    HMes = HMes + 1
                End If
                FHasta = CDate("01" & "/" & Format(HMes, "00") & "/" & Format(HAnyo, "0000"))
                FHasta = DateAdd("d", (-1), FHasta)
            End If
    
    End Select
    
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub CargaCombo()
Dim i As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For i = 0 To Combo1.Count - 1
        Combo1(i).Clear
    Next i
    
    Combo1(0).AddItem "No procesados"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Procesados"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "Ambos"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2


    Combo1(1).AddItem "Cooperativa"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Ajena"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1


End Sub

' antes de cambiar el calculo a horas por precio por trabajador en lugar de por categoria
'
'Private Function CargarTemporal2(cadTABLA As String, cadselect As String, CadConcepto As String, CadForfait As String, CadVariedad As String) As Boolean
'Dim Sql As String
'Dim Sql2 As String
'Dim Sql3 As String
'Dim Sql4 As String
'Dim SqlInsert As String
'Dim CadValues As String
'Dim ImpFactu As Currency
'Dim PrecioFact As Currency
'Dim Diferencia As Currency
'Dim Importe As Currency
'Dim Horas As Currency
'Dim TotalKilos As Long
'Dim Nregs As Long
'Dim NRegs2 As Long
'Dim NConceptos As Long
'Dim Kilos As Long
'Dim ImporteVar As Currency
'
'Dim Rs As ADODB.Recordset
'Dim Rs2 As ADODB.Recordset
'
'
'    On Error GoTo eCargarTemporal
'
'    CargarTemporal2 = False
'    Screen.MousePointer = vbHourglass
'    Label4(6).visible = True
'    DoEvents
'
'
''1 - CARGAMOS UNA TABLA INTERMEDIA PARA LOS GASTOS DIARIOS QUE HAY QUE PRORRATEAR ENTRE VARIEDADES
'    BorrarTMPCostesDia
'    CrearTMPCostesDia ""
'
'    Sql = "insert into tmpCostesDia (fecha, codcoste, horas, impcoste) "
'    Sql = Sql & " select cclindia2.fecha, cclindia2.codcoste, sum(cclindia2.horas) horas, round(sum(cclindia2.horas) * salarios.impsalar,2) as impcoste  "
'    Sql = Sql & " from cclindia2 inner join salarios on cclindia2.codcateg = salarios.codcateg "
'    Sql = Sql & " where date(fecha) in (select date(fecha) from cccabdia "
'    Sql = Sql & " where (1=1) "
'
'    If cadselect <> "" Then Sql = Sql & " and " & Replace(cadselect, "cccaborden.fechaini", "cccabdia.fecha")
'    If CadConcepto <> "" Then Sql = Sql & " and " & CadConcepto
'
'    Sql = Sql & ") "
'    Sql = Sql & " group by 1,2 "
'    Sql = Sql & " order by 1,2 "
'
'    conn.Execute Sql
'
''2 - CARGAMOS LA TABLA DE RESULTADOS
'    BorrarTMPKilosRes
'    CrearTMPKilosRes ""
'
'    ' introducimos lo que hay en las ordenes de confeccion
'    SqlInsert = "insert into tmpKilosRes (codorden, fecha, codvarie, codtipco, codcoste, kilos, importe) values "
'
'    Sql = " select cccaborden.codorden, cccaborden.fechaini, cclinorden3.codvarie, forfaits.codtipco, sum(cclinorden3.kilosnet) kilos " & _
'          "  from (cccaborden inner join cclinorden3 on cccaborden.codorden = cclinorden3.codorden) " & _
'          "  inner join forfaits on cclinorden3.codforfait = forfaits.codforfait "
'    Sql = Sql & " where (1=1) "
'
'    If cadselect <> "" Then Sql = Sql & " and " & cadselect
'    If CadVariedad <> "" Then Sql = Sql & " and " & CadVariedad
'    If CadForfait <> "" Then Sql = Sql & " and " & CadForfait
'
'    ' necesito el sql4 para añadirle la condicion del codigo de orden
'    Sql4 = Sql
'
'    Sql = Sql & "  group by 1, 2, 3, 4 " & _
'                "  order by 1, 2, 3, 4 "
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    While Not Rs.EOF
'        Sql2 = "select codcoste, sum(if(importe is null, 0,importe)) importe from ("
'        Sql2 = Sql2 & "select cclinorden2.codcoste, cclinorden2.codcateg, round(sum(cclinorden2.horas) * salarios.impsalar,2) importe "
'        Sql2 = Sql2 & " from cclinorden2 inner join salarios on cclinorden2.codcateg = salarios.codcateg "
'        Sql2 = Sql2 & " where cclinorden2.codorden = " & DBSet(Rs!CodOrden, "N")
'        Sql2 = Sql2 & " group by 1,2 "
'        Sql2 = Sql2 & " order by 1,2 ) vconsulta "
'        Sql2 = Sql2 & " group by 1 "
'        Sql2 = Sql2 & " order by 1 "
'
'        TotalKilos = DBLet(Rs!Kilos)
'        Nregs = TotalRegistrosConsulta("select * from (" & Sql4 & " and cccaborden.codorden = " & DBSet(Rs!CodOrden, "N") & "  group by 1, 2, 3, 4  order by 1, 2, 3, 4 ) aaaconsulta ")
'        NConceptos = TotalRegistrosConsulta(Sql2)
'
'        ' si esta orden no ha tenido a nadie trabajando ( no se ha cargado nada desde el reloj )
'        ' la cargo igualmente con concepto -1 ---> ESTO NO DEBERIA OCURRIR
'        If NConceptos = 0 Then
'
'            CadValues = "(" & DBSet(Rs!CodOrden, "N") & "," & DBSet(Rs!FechaIni, "F") & ","
'            CadValues = CadValues & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!codtipco, "N") & ","
'            CadValues = CadValues & "0," ' sin concepto
'            CadValues = CadValues & DBSet(TotalKilos, "N") & ","
'            CadValues = CadValues & DBSet(0, "N") & "),"
'
'            CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
'            conn.Execute SqlInsert & CadValues
'
'        Else
'
'            CadValues = ""
'
'            ' repartimos
'            Set Rs2 = New ADODB.Recordset
'            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'            While Not Rs2.EOF
'                Kilos = Round2(TotalKilos / NConceptos, 0)
'                Importe = Round2(DBLet(Rs2!Importe) / Nregs, 2)
'
'                CadValues = CadValues & "(" & DBSet(Rs!CodOrden, "N") & "," & DBSet(Rs!FechaIni, "F") & ","
'                CadValues = CadValues & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!codtipco, "N") & ","
'                CadValues = CadValues & DBSet(Rs2!codcoste, "N") & ","
'                CadValues = CadValues & DBSet(Kilos, "N") & ","
'                CadValues = CadValues & DBSet(Importe, "N") & "),"
'
'                Rs2.MoveNext
'            Wend
'
'            Set Rs2 = Nothing
'
'            If CadValues <> "" Then
'                CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
'                conn.Execute SqlInsert & CadValues
'            End If
'
'        End If
'
'        Rs.MoveNext
'    Wend
'
'    Set Rs = Nothing
'
'
''3 - INSERTAMOS EN LAS TABLAS TEMPORALES PARA EL REPORT
''    dejamos en tmpinformes y en tmpliquidacion los datos prorrateados para el informe
'
'    conn.Execute "delete from tmpinformes where codusu = " & vUsu.Codigo
'
'
'    Sql = "insert into tmpinformes (codusu, fecha1, codigo1, campo1, campo2, importe2, importe3) "
'    Sql = Sql & " select " & vUsu.Codigo & ", fecha, codvarie, codtipco, codcoste, sum(kilos), sum(importe) "
'    Sql = Sql & " from tmpkilosres "
'    Sql = Sql & " group by 1, 2, 3, 4, 5"
'
'    conn.Execute Sql
'
''4 - REPARTIMOS LOS COSTES GENERALES DIARIOS SEGUN LOS KILOS
''    dejamos en tmpinformes y en tmpliquidacion los datos prorrateados para el informe
'    conn.Execute "delete from tmpliquidacion where codusu = " & vUsu.Codigo
'
'                                                    'fecha,    variedad, codcoste, importe
'    SqlInsert = "insert into tmpliquidacion (codusu, fechaini, codvarie, codsocio, importe) values "
'
'    CadValues = ""
'
'    Sql = "select fecha, codcoste, sum(impcoste) importe from tmpcostesdia group by 1,2 order by 1,2"
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    While Not Rs.EOF
'        Sql2 = "select sum(importe2) as totalkilos from tmpinformes where codusu = " & vUsu.Codigo & " and fecha1 = " & DBSet(Rs!Fecha, "F")
'        TotalKilos = DevuelveValor(Sql2)
'
'        Sql2 = "select codigo1, sum(importe2) as kilosvar from tmpinformes where codusu = " & vUsu.Codigo & " and fecha1 = " & DBSet(Rs!Fecha, "F")
'        Sql2 = Sql2 & " group by 1 "
'        Sql2 = Sql2 & " order by 1 "
'
'        Set Rs2 = New ADODB.Recordset
'        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        While Not Rs2.EOF
'            ImporteVar = 0
'
'            If TotalKilos <> 0 Then
'                ImporteVar = Round2(Rs2!kilosvar * Rs!Importe / TotalKilos, 2)
'            End If
'
'            CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(Rs!Fecha, "F") & "," & DBSet(Rs2!Codigo1, "N") & ","
'            CadValues = CadValues & DBSet(Rs!codcoste, "N") & "," & DBSet(ImporteVar, "N") & "),"
'
'            Rs2.MoveNext
'        Wend
'
'        Rs.MoveNext
'    Wend
'
'    Set Rs = Nothing
'
'    If CadValues <> "" Then
'        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
'
'        conn.Execute SqlInsert & CadValues
'    End If
'
'
''5 - CARGAMOS LA TABLA TMPINFVENTAS PARA PODER IMPRIMIR LOS CONCEPTOS DE CONFECCION ENCOLUMNADOS
''
'    conn.Execute "delete from tmpinfventas where codusu = " & vUsu.Codigo
'
'                                                    'fecha,  variedad, codcoste, kilos1,  importe1....
'    Sql = "insert into tmpinfventas (codusu, fecalbar, numalbar, numcajas, codigo1, gastos1, codigo2, gastos2, codigo3, gastos3, codigo4, gastos4, codigo5, gastos5) "
'    '                                       fecha,  variedad, codcoste,kilos, importe,.....
'    Sql = Sql & "select " & vUsu.Codigo & ",fecha1, codigo1,  campo2,  0,0,0,0,0,0,0,0,0,0 from tmpinformes where codusu = " & vUsu.Codigo
'    Sql = Sql & " group by 1,2,3,4 "
'
'    conn.Execute Sql
'
'    Sql = "select fecha1, codigo1, campo2, campo1, importe2, importe3 from tmpinformes where codusu = " & vUsu.Codigo
'    Sql = Sql & " order by 1,2,3,4 "
'
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    While Not Rs.EOF
'        CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(Rs!fecha1, "F") & "," & DBSet(Rs!Codigo1, "N") & ","
'        CadValues = CadValues & DBSet(Rs!campo1, "N") & ","
'
'        Sql = "update tmpinfventas set codigo" & Format(DBLet(Rs!campo1), "0") & " = codigo" & Format(DBLet(Rs!campo1), "0") & " + " & DBSet(Rs!importe2, "N") ' kilos
'        Sql = Sql & ", gastos" & Format(DBLet(Rs!campo1), "0") & " = gastos" & Format(DBLet(Rs!campo1), "0") & " + " & DBSet(Rs!importe3, "N") ' importe
'        Sql = Sql & " where codusu = " & vUsu.Codigo
'        Sql = Sql & " and fecalbar = " & DBSet(Rs!fecha1, "F") ' fecha
'        Sql = Sql & " and numalbar = " & DBSet(Rs!Codigo1, "N") ' variedad
'        Sql = Sql & " and numcajas = " & DBSet(Rs!campo2, "N") ' codcoste
'
'        conn.Execute Sql
'
'        Rs.MoveNext
'    Wend
'
'    Set Rs = Nothing
'
'
'    Screen.MousePointer = vbDefault
'    Label4(6).visible = False
'    DoEvents
'    Set Rs = Nothing
'    CargarTemporal2 = True
'    Exit Function
'
'eCargarTemporal:
'    Screen.MousePointer = vbDefault
'    Label4(6).visible = False
'    MuestraError Err.Number, "Cargar Temporal", Err.Description
'End Function
'
'

'
' Antigua Carga de ordenes de confeccion
'
Private Sub CmdAcepCargaOrdAntiguo_Click()
Dim SQL As String

    InicializarVbles

    If txtcodigo(6).Text = "" Then
        MsgBox "Debe introducir obligatoriamente una Fecha.", vbExclamation
        PonerFoco txtcodigo(6)
        Exit Sub
    End If
    
    SQL = "select count(*) from palets where date(horafconf) = " & DBSet(txtcodigo(6).Text, "F") & " and intorden = 0 "
    
    If TotalRegistros(SQL) = 0 Then
        MsgBox "No hay palets en esta fecha para construir ordenes de confección.", vbExclamation
    
    Else
        '[Monica]09/07/2012: comprobamos que todos los registros a procesar se han acabado
        '                    si hay registros con la tarea sin acabar damos un aviso para que los revisen
        SQL = "select count(*) from cctrabaconf where date(fechaini) = " & DBSet(txtcodigo(6).Text, "F") & " and procesado = 0 and sinacabar = 1"
        If TotalRegistros(SQL) Then
            MsgBox "Hay registros de trabajador que no se han acabado. Revise.", vbExclamation
            PonerFoco txtcodigo(6)
            Exit Sub
        End If
    
        ' Comprobamos las claves referenciales del fichero que cargan que no las tiene en la base de datos para que no dé errores
        ' al cargar la tabla
        If ComprobarFichero(txtcodigo(6).Text) Then
            If TotalRegistrosConsulta("select * from tmpinformes where codusu = " & vUsu.Codigo) <> 0 Then
                cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
                cadTitulo = "Errores en el Traspaso de Costes"
                cadNombreRPT = "rErroresTrasCostes.rpt"
                LlamarImprimir
            Else
                If ProcesoCargaOrdenes Then
                    MsgBox "Proceso realizado correctamente", vbExclamation
                    cmdCancelar_Click
                End If
            End If
        End If
    End If

End Sub


'********************************************
' NUEVO PUNTO DE PROCESO DE CARGA DE ORDENES
'********************************************

Private Function ProcesoCargaOrdenesNew(FIni As String, FFin As String, ByRef pb1 As ProgressBar, Label4 As label, ByRef Label5 As label) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Sql4 As String
Dim SqlFec As String

Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim RsFec As ADODB.Recordset

Dim nRegs As Long
Dim CodOrden As Long
Dim Cliente As Long
Dim NumLin As Long

Dim b As Boolean
Dim codTipoM As String
Dim vTipoMov As CTiposMov
Dim devuelve As String
Dim Existe As Boolean
Dim CadValues As String
Dim Mens As String


    On Error GoTo eProcesoCargaOrdenes


    ProcesoCargaOrdenesNew = False
    
    conn.Execute "delete from cccabecera "
    conn.Execute "delete from cclineas1 "
    conn.Execute "delete from cclineas2 "

'    Conn.BeginTrans
    
    ' primer cursor por fecha
    SqlFec = "select distinct date(fechafin) fechafin from cctrabaconf where date(fechafin) between " & DBSet(FIni, "F") & " and " & DBSet(FFin, "F")
    
    Set RsFec = New ADODB.Recordset
    RsFec.Open SqlFec, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    CodOrden = 0
    
    
    While Not RsFec.EOF
        Label4.visible = True
        Label4.Caption = "Procesando fecha " & Format(DBLet(RsFec!FechaFin, "F"), "dd/mm/yyyy")
        DoEvents
    
        SQL = "select ccconcostes.tipokilos, cctrabaconf.* from cctrabaconf inner join ccconcostes on cctrabaconf.codcoste = ccconcostes.codcoste where date(fechafin) = " & DBSet(RsFec!FechaFin, "F")
        SQL = SQL & " and ccconcostes.tipokilos = 0 and ccconcostes.tipocoste = 0 "
        SQL = SQL & " union "
        SQL = SQL & " select ccconcostes.tipokilos, cctrabaconf.* from cctrabaconf inner join ccconcostes on cctrabaconf.codcoste = ccconcostes.codcoste where date(fechafin) = " & DBSet(RsFec!FechaFin, "F")
        SQL = SQL & " and ccconcostes.tipokilos = 1 and ccconcostes.tipocoste = 0 "
        SQL = SQL & " union "
        SQL = SQL & " select ccconcostes.tipokilos, cctrabaconf.* from cctrabaconf inner join ccconcostes on cctrabaconf.codcoste = ccconcostes.codcoste where date(fechafin) = " & DBSet(RsFec!FechaFin, "F")
        SQL = SQL & " and ccconcostes.tipokilos = 2 and ccconcostes.tipocoste = 0 "
'        Sql = Sql & " union "
'        Sql = Sql & " select ccconcostes.tipokilos, cctrabaconf.* from cctrabaconf inner join ccconcostes on cctrabaconf.codcoste = ccconcostes.codcoste where date(fechafin) = " & DBSet(RsFec!FechaFin, "F")
'        Sql = Sql & " and ccconcostes.tipokilos = 3 "
        SQL = SQL & " order by 1, 2, 3 "
        
        nRegs = TotalRegistrosConsulta(SQL)
        
        Label5.visible = True
        pb1.visible = True
        Label5.Caption = "Cargando Tablas intermedias"
        DoEvents
        
        If nRegs = 0 Then
            b = False
        Else
            pb1.Max = nRegs
            pb1.Value = 0
        End If
        
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        b = True
        While Not Rs.EOF And b
            IncrementarProgresNew pb1, 1
            
            Label5.Caption = "Procesando Registro del Trabajador " & DBLet(Rs!codtraba, "N")
            DoEvents
            
            Select Case DBLet(Rs!tipokilos, "N")
                Case 0 ' sobre kilos entrados
                    If b Then
                        Mens = "Insertar sobre Kilos Entrados"
                        b = InsertarKilosEntrados(Rs, CodOrden, Mens)
                    End If
                
                Case 1 ' sobre kilos volcados
                    If b Then
                        Mens = "Insertar sobre Kilos Volcados"
                        b = InsertarKilosVolcados(Rs, CodOrden, Mens)
                    End If
                
                Case 2  ' sobre kilos confeccionados
                    If b Then
                        Mens = "Insertar sobre Kilos Confeccionados"
                        b = InsertarKilosConfeccionados(Rs, CodOrden, Mens)
                    End If
                
                Case 3 ' Tipo de entrada de no procesa
                    If b Then
'                        Mens = "Insertar sobre Kilos Totales"
'                        b = InsertarKilosTotales(RS, CodOrden, Mens)
                    End If
            End Select
            
    
            Rs.MoveNext
        Wend
                
        Set Rs = Nothing
    
        RsFec.MoveNext
    
    Wend
    
    Set RsFec = Nothing
    
    
    'kilos entrados, kilos volcados, kilos confeccionados por variedad
    ' en la cclineas6
    Mens = "Calcular Kilos Entrados"
    b = ActualizarKilos2(Mens)
    
    
eProcesoCargaOrdenes:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Proceso Carga Temporales", Err.Description
        b = False
    Else
        If Not b Then
            MsgBox "Proceso Carga Temporales: " & vbCrLf & vbCrLf & Mens, vbExclamation
        End If
    End If
    
    pb1.visible = False
    Label5.visible = False
    Label4.visible = False
    
    If b Then
        ProcesoCargaOrdenesNew = True
'        Conn.CommitTrans
    Else
        ProcesoCargaOrdenesNew = False
'        Conn.RollbackTrans
    End If
End Function




Private Function InsertarKilosEntrados(ByRef Rs As ADODB.Recordset, CodOrden As Long, Mens As String) As Boolean
Dim Cliente As Long
Dim Sql4 As String
Dim HoraIni As String
Dim HoraFin As String
Dim SQL1 As String
Dim Sql2 As String

Dim SQLinsert As String
Dim SqlValues As String

Dim Diferencia As Currency
Dim Importe As Currency
Dim Precio As Currency
Dim CosteKilo As Single
Dim KilosTotales As Long
Dim ImporteLinea As Currency

Dim Rs1 As ADODB.Recordset
Dim SqlTrab As String
Dim NTrabajadores As Long


    On Error GoTo eInsertarKilosEntrados
    
    InsertarKilosEntrados = False

    SQLinsert = "insert into cccabecera (idcontador,fecha,codtraba,codcoste,fechaini,fechafin,codvarie,kilos,importe,costekilo,importetrab) values "

    Diferencia = Round2(DateDiff("n", CDate(Rs!FechaIni), CDate(Rs!FechaFin)) / 60, 2)
    Precio = DevuelveValor("select prhoracoste from straba where codtraba = " & DBSet(Rs!codtraba, "N"))
    Importe = Round2(Diferencia * Precio, 2)

    '[Monica]17/12/2012: los kilos los tengo que dividir entre el nro de trabajadores que estan trabajando en ese momento
    SqlTrab = "select count(distinct codtraba) from cctrabaconf where codcoste = " & DBSet(Rs!codCoste, "N")
    SqlTrab = SqlTrab & " and (fechaini between " & DBSet(Rs!FechaIni, "FH") & " and " & DBSet(Rs!FechaFin, "FH") & " or "
    SqlTrab = SqlTrab & " fechafin between " & DBSet(Rs!FechaIni, "FH") & " and " & DBSet(Rs!FechaFin, "FH") & ")"
    
    NTrabajadores = TotalRegistros(SqlTrab)


    'Vamos a trzpalets del periodo agrupando kilosnetos por variedad
    '[Monica]17/12/2012: los kilos los tengo que dividir entre el nro de trabajadores que estan trabajando en ese momento
    SQL1 = "select codvarie, round(sum(numkilos) / " & NTrabajadores & ",0)  kilos from trzpalets where hora between " & DBSet(Rs!FechaIni, "FH") & " and " & DBSet(Rs!FechaFin, "FH")
'    If Variedades <> "" Then Sql1 = Sql1 & " and codvarie in " & Variedades
    SQL1 = SQL1 & " group by 1 "
    SQL1 = SQL1 & " order by 1 "
    
    KilosTotales = DevuelveValor("select sum(kilos) from (" & SQL1 & ") kilostot")
    
    
    ' Si ese dia no han entrado kilos inserto un registro sin variedad, kilos unicamente con el importe y sin coste por kilo
    If KilosTotales = 0 Then
        CodOrden = CodOrden + 1
    
        SqlValues = "(" & DBSet(CodOrden, "N") & "," & DBSet(Rs!FechaFin, "F") & "," & DBSet(Rs!codtraba, "N") & "," & DBSet(Rs!codCoste, "N")
        SqlValues = SqlValues & "," & DBSet(Rs!FechaIni, "FH") & "," & DBSet(Rs!FechaFin, "FH") & ","
        
        CosteKilo = 0
    
        SqlValues = SqlValues & DBSet(0, "N") & "," & DBSet(0, "N") & "," & DBSet(Importe, "N") & "," & TransformaComasPuntos(ImporteSinFormato(CStr(CosteKilo))) & "," & DBSet(Importe, "N") & "),"
    
    Else
    
        Set Rs1 = New ADODB.Recordset
        Rs1.Open SQL1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        SqlValues = ""
        
        While Not Rs1.EOF
        
            CodOrden = CodOrden + 1
        
            ImporteLinea = Round2(Importe * Rs1!Kilos / KilosTotales, 2)
        
            SqlValues = SqlValues & "(" & DBSet(CodOrden, "N") & "," & DBSet(Rs!FechaFin, "F") & "," & DBSet(Rs!codtraba, "N") & "," & DBSet(Rs!codCoste, "N")
            SqlValues = SqlValues & "," & DBSet(Rs!FechaIni, "FH") & "," & DBSet(Rs!FechaFin, "FH") & ","
            
            CosteKilo = ImporteLinea / DBLet(Rs1!Kilos, "N")
        
            SqlValues = SqlValues & DBSet(Rs1!codvarie, "N") & "," & DBSet(Rs1!Kilos, "N") & "," & DBSet(ImporteLinea, "N") & "," & TransformaComasPuntos(ImporteSinFormato(CStr(CosteKilo))) & "," & DBSet(Importe, "N") & "),"
            
            'Insertamos que palets de traza ha intervenido de cada variedad
            Sql2 = "insert into cclineas1 (idcontador,idpalet) "
            Sql2 = Sql2 & " select " & DBSet(CodOrden, "N") & ", idpalet from trzpalets where hora between " & DBSet(Rs!FechaIni, "FH") & " and " & DBSet(Rs!FechaFin, "FH")
            Sql2 = Sql2 & " and codvarie = " & DBSet(Rs1!codvarie, "N")
            
            conn.Execute Sql2
            
            
'            'Insertamos/actualizamos kilos entrados por variedad
'            ActualizarKilos CStr(RS1!codvarie), 0, CStr(RS1!Kilos)
            
            
            Rs1.MoveNext
        Wend
        
        Set Rs1 = Nothing

    End If

    If SqlValues <> "" Then
        SqlValues = Mid(SqlValues, 1, Len(SqlValues) - 1)
        conn.Execute SQLinsert & SqlValues
    End If

    InsertarKilosEntrados = True
    Exit Function

eInsertarKilosEntrados:
    Mens = Mens & vbCrLf & vbCrLf & Err.Description
End Function


Private Sub ActualizarKilos(Variedad As String, Tipo As Byte, Kilos As String)
'Tipo: 0=Entrados
'      1=Volcados
'      2=Confeccionados
Dim SQL As String
Dim Sql2 As String

    On Error Resume Next

    SQL = "select count(*) from cclineas6 where codvarie = " & DBSet(Variedad, "N")
    
    Sql2 = ""
    
    If TotalRegistros(SQL) = 0 Then
        Select Case Tipo
            Case 0
                Sql2 = "insert into cclineas6(codvarie, kentrados, kvolcados, kconfeccionados) values (" & DBSet(Variedad, "N") & ","
                Sql2 = Sql2 & DBSet(Kilos, "N") & ",0,0)"
            Case 1
                Sql2 = "insert into cclineas6(codvarie, kentrados, kvolcados, kconfeccionados) values (" & DBSet(Variedad, "N") & ","
                Sql2 = Sql2 & "0," & DBSet(Kilos, "N") & ",0)"
            Case 2
                Sql2 = "insert into cclineas6(codvarie, kentrados, kvolcados, kconfeccionados) values (" & DBSet(Variedad, "N") & ","
                Sql2 = Sql2 & "0,0," & DBSet(Kilos, "N") & ")"
        End Select
    Else
        Select Case Tipo
            Case 0
                Sql2 = "update cclineas6 set kentrados = kentrados + " & DBSet(Kilos, "N") & " where codvarie = " & DBSet(Variedad, "N")
            Case 1
                Sql2 = "update cclineas6 set kvolcados = kvolcados + " & DBSet(Kilos, "N") & " where codvarie = " & DBSet(Variedad, "N")
            Case 2
                Sql2 = "update cclineas6 set kconfeccionados = kconfeccionados + " & DBSet(Kilos, "N") & " where codvarie = " & DBSet(Variedad, "N")
        End Select
    End If
        
    If Sql2 <> "" Then
        conn.Execute Sql2
    End If

End Sub


Private Function ActualizarKilos2(Mens As String) As Boolean
'Tipo: 0=Entrados
'      1=Volcados
'      2=Confeccionados
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim Sql4 As String

Dim Rs As ADODB.Recordset

    On Error GoTo eActualizarKilos2

    ActualizarKilos2 = False

    SQL = "delete from cclineas6 "
    conn.Execute SQL
    
    ' entrados
    Label4(43).Caption = "Cargando Kilos Entrados"
    DoEvents
    
    SQL = "select distinct idpalet from cclineas1 where idcontador in (select idcontador from cccabecera where codcoste in (select codcoste from ccconcostes where tipokilos = 0)) "
    
    Sql2 = "select codvarie, sum(numkilos) kilos from trzpalets where idpalet in (" & SQL & ") group by 1 order by 1"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Sql3 = "select count(*) from cclineas6 where codvarie = " & DBSet(Rs!codvarie, "N")
        If TotalRegistros(Sql3) = 0 Then
            Sql4 = "insert into cclineas6(codvarie, kentrados, kvolcados, kconfeccionados) values (" & DBSet(Rs!codvarie, "N") & ","
            Sql4 = Sql4 & DBSet(Rs!Kilos, "N") & ",0,0)"
        Else
            Sql4 = "update cclineas6 set kentrados = kentrados + " & DBSet(Rs!Kilos, "N") & " where codvarie = " & DBSet(Rs!codvarie, "N")
        End If
        
        conn.Execute Sql4
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    
    ' VOLCADOS
    Label4(43).Caption = "Cargando Kilos Volcados"
    DoEvents
    
    
    SQL = "select distinct idpalet from cclineas1 where idcontador in (select idcontador from cccabecera where codcoste in (select codcoste from ccconcostes where tipokilos = 1)) "
    
    Sql2 = "select codvarie, sum(numkilos) kilos from trzpalets where idpalet in (" & SQL & ") group by 1 order by 1"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Sql3 = "select * from cclineas6 where codvarie = " & DBSet(Rs!codvarie, "N")
        If TotalRegistrosConsulta(Sql3) = 0 Then
            Sql4 = "insert into cclineas6(codvarie, kentrados, kvolcados, kconfeccionados) values (" & DBSet(Rs!codvarie, "N") & ","
            Sql4 = Sql4 & "0," & DBSet(Rs!Kilos, "N") & ",0)"
        Else
            Sql4 = "update cclineas6 set kvolcados = kvolcados + " & DBSet(Rs!Kilos, "N") & " where codvarie = " & DBSet(Rs!codvarie, "N")
        End If
        
        conn.Execute Sql4
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    
    ' CONFECCIONADOS
    Label4(43).Caption = "Cargando Kilos Confeccionados"
    DoEvents
    
    
    SQL = "select distinct numpalet from cclineas2 where idcontador in (select idcontador from cccabecera where codcoste in (select codcoste from ccconcostes where tipokilos = 2)) "
    
    Sql2 = "select codvarie, sum(pesoneto) kilos from palets_variedad where numpalet in (" & SQL & ") group by 1 order by 1"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Sql3 = "select count(*) from cclineas6 where codvarie = " & DBSet(Rs!codvarie, "N")
        If TotalRegistros(Sql3) = 0 Then
            Sql4 = "insert into cclineas6(codvarie, kentrados, kvolcados, kconfeccionados) values (" & DBSet(Rs!codvarie, "N") & ","
            Sql4 = Sql4 & "0,0," & DBSet(Rs!Kilos, "N") & ")"
        Else
            Sql4 = "update cclineas6 set kconfeccionados = kconfeccionados + " & DBSet(Rs!Kilos, "N") & " where codvarie = " & DBSet(Rs!codvarie, "N")
        End If
        
        conn.Execute Sql4
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    ActualizarKilos2 = True
    Exit Function
    
eActualizarKilos2:
    Mens = Mens & Err.Description
End Function



Private Function InsertarKilosVolcados(ByRef Rs As ADODB.Recordset, CodOrden As Long, Mens As String) As Boolean
Dim Cliente As Long
Dim Sql4 As String
Dim HoraIni As String
Dim HoraFin As String
Dim SQL1 As String
Dim Sql2 As String
Dim Rs1 As ADODB.Recordset
Dim Rs2 As ADODB.Recordset

Dim SQLinsert As String
Dim Diferencia As Currency
Dim Precio As Currency
Dim Importe As Currency
Dim CosteKilo As Single

Dim SqlValues As String

Dim Lineas As String

Dim KilosTotales As Long
Dim ImporteLinea As Currency
Dim TKilos As Long
Dim LineasTraza As String
Dim SqlTrab As String
Dim NTrabajadores As Long

    On Error GoTo eInsertarKilosVolcados
    
    InsertarKilosVolcados = False

    SQLinsert = "insert into cccabecera (idcontador,fecha,codtraba,codcoste,fechaini,fechafin,codvarie,kilos,importe,costekilo,importetrab)  values "

    Diferencia = Round2(DateDiff("n", CDate(Rs!FechaIni), CDate(Rs!FechaFin)) / 60, 2)
    Precio = DevuelveValor("select prhoracoste from straba where codtraba = " & DBSet(Rs!codtraba, "N"))
    Importe = Round2(Diferencia * Precio, 2)

    
    Lineas = LineasActividad(CStr(Rs!codCoste))
    
    TKilos = 0

    'Primero para cada linea de palet confeccionado en ese rango y en esa linea/lineas de entrada (palets)
    ' cojo la linea de confeccion y me voy a trzlineas_cargas
                                            '30/01/2013:horaini
    SQL1 = "select distinct linconfe from palets where (horaiconf between " & DBSet(Rs!FechaIni, "FH") & " and " & DBSet(Rs!FechaFin, "FH")
    SQL1 = SQL1 & " or horafin between " & DBSet(Rs!FechaIni, "FH") & " and " & DBSet(Rs!FechaFin, "FH") & ")"
    If Lineas <> "" Then
        SQL1 = SQL1 & " and linentrada in (" & Lineas & ")"
    End If
    SQL1 = SQL1 & " order by 1 "

    Set Rs1 = New ADODB.Recordset
    Rs1.Open SQL1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    'Las lineas de traza no puedes ser un campo vacio
    LineasTraza = ""
    While Not Rs1.EOF
        LineasTraza = LineasTraza & DBLet(Rs1!linconfe) & ","
        Rs1.MoveNext
    Wend
    Set Rs1 = Nothing
    If Len(LineasTraza) > 0 Then LineasTraza = Mid(LineasTraza, 1, Len(LineasTraza) - 1)
    
    If LineasTraza = "" Then
        CodOrden = CodOrden + 1
    
        TKilos = 0
    
    Else
        
        '[Monica]17/12/2012: los kilos los tengo que dividir entre el nro de trabajadores que estan trabajando en ese momento
        SqlTrab = "select count(distinct codtraba) from cctrabaconf where codcoste = " & DBSet(Rs!codCoste, "N")
        SqlTrab = SqlTrab & " and (fechaini between " & DBSet(Rs!FechaIni, "FH") & " and " & DBSet(Rs!FechaFin, "FH") & " or "
        SqlTrab = SqlTrab & " fechafin between " & DBSet(Rs!FechaIni, "FH") & " and " & DBSet(Rs!FechaFin, "FH") & ")"
        
        NTrabajadores = TotalRegistros(SqlTrab)
        
        Sql2 = "select codvarie, round(sum(numkilos) / " & NTrabajadores & ",0) kilos from trzpalets where idpalet in "
        Sql2 = Sql2 & "(select idpalet from trzlineas_cargas where linea in (" & LineasTraza & ") " 'DBSet(RS1!linconfe, "N")
        Sql2 = Sql2 & " and fechahora between " & DBSet(Rs!FechaIni, "FH") & " and " & DBSet(Rs!FechaFin, "FH") & ")"
'        If Variedades <> "" Then Sql2 = Sql2 & " and codvarie in " & Variedades
        Sql2 = Sql2 & " group by 1 "
        Sql2 = Sql2 & " order by 1 "
    
        KilosTotales = DevuelveValor("select sum(kilos) from (" & Sql2 & ") aaaaa ")
        
        If KilosTotales <> 0 Then
            TKilos = TKilos + KilosTotales
    
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            SqlValues = ""
        
            While Not Rs2.EOF
        
                CodOrden = CodOrden + 1
        
                ImporteLinea = Round2(Importe * Rs2!Kilos / KilosTotales, 2)
        
                SqlValues = SqlValues & "(" & DBSet(CodOrden, "N") & "," & DBSet(Rs!FechaFin, "F") & "," & DBSet(Rs!codtraba, "N") & "," & TransformaComasPuntos(ImporteSinFormato(Rs!codCoste))
                SqlValues = SqlValues & "," & DBSet(Rs!FechaIni, "FH") & "," & DBSet(Rs!FechaFin, "FH") & ","
                
                CosteKilo = ImporteLinea / DBLet(Rs2!Kilos, "N")
            
                SqlValues = SqlValues & DBSet(Rs2!codvarie, "N") & "," & DBSet(Rs2!Kilos, "N") & "," & DBSet(ImporteLinea, "N") & "," & TransformaComasPuntos(ImporteSinFormato(CStr(CosteKilo))) & "," & DBSet(Importe, "N") & "),"
                
                'Insertamos que palets de traza han intervenido de cada variedad
                Sql2 = "insert into cclineas1 (idcontador,idpalet) "
                Sql2 = Sql2 & " select " & DBSet(CodOrden, "N") & ", trzlineas_cargas.idpalet from trzlineas_cargas inner join trzpalets on trzlineas_cargas.idpalet = trzpalets.idpalet where linea in (" & LineasTraza & ") " ' & DBSet(RS1!linconfe, "N")
                Sql2 = Sql2 & " and trzlineas_cargas.fechahora between " & DBSet(Rs!FechaIni, "FH") & " and " & DBSet(Rs!FechaFin, "FH")
                Sql2 = Sql2 & " and trzpalets.codvarie = " & DBSet(Rs2!codvarie, "N")
                    
                conn.Execute Sql2
                
                'Insertamos que palets confeccionados han intervenido en cada variedad
                Sql2 = "insert into cclineas2 (idcontador,numpalet) select distinct " & DBSet(CodOrden, "N") & ", palets.numpalet  "
                                                                                                                '30/01/2013:horaini
                Sql2 = Sql2 & " from palets inner join palets_variedad on palets.numpalet = palets_variedad.numpalet where (horaiconf between " & DBSet(Rs!FechaIni, "FH") & " and " & DBSet(Rs!FechaFin, "FH")
                Sql2 = Sql2 & " or horafin between " & DBSet(Rs!FechaIni, "FH") & " and " & DBSet(Rs!FechaFin, "FH") & ")"
                Sql2 = Sql2 & " and palets_variedad.codvarie = " & DBSet(Rs2!codvarie, "N")
                If Lineas <> "" Then
                    Sql2 = Sql2 & " and linentrada in (" & Lineas & ")"
                End If
                conn.Execute Sql2
                    
'                'Insertamos/actualizamos kilos volcados por variedad
'                ActualizarKilos CStr(Rs2!codvarie), 1, CStr(Rs2!Kilos)
                    
                Rs2.MoveNext
            Wend
            Set Rs2 = Nothing
            
        Else
            CodOrden = CodOrden + 1
            
        End If
    End If
    '        RS1.MoveNext
    '    Wend
    '    Set RS1 = Nothing
    
    If TKilos = 0 Then
        SqlValues = "(" & DBSet(CodOrden, "N") & "," & DBSet(Rs!FechaFin, "F") & "," & DBSet(Rs!codtraba, "N") & "," & DBSet(Rs!codCoste, "N")
        SqlValues = SqlValues & "," & DBSet(Rs!FechaIni, "FH") & "," & DBSet(Rs!FechaFin, "FH") & ","
        
        CosteKilo = 0
    
        SqlValues = SqlValues & DBSet(0, "N") & "," & DBSet(0, "N") & "," & DBSet(Importe, "N") & "," & TransformaComasPuntos(ImporteSinFormato(CStr(CosteKilo))) & "," & DBSet(Importe, "N") & "),"
    
    End If

    If SqlValues <> "" Then
        SqlValues = Mid(SqlValues, 1, Len(SqlValues) - 1)
        conn.Execute SQLinsert & SqlValues
    End If
    


    InsertarKilosVolcados = True
    Exit Function

eInsertarKilosVolcados:
    Mens = Mens & vbCrLf & vbCrLf & Err.Description
End Function


Private Function InsertarKilosConfeccionados(ByRef Rs As ADODB.Recordset, CodOrden As Long, Mens As String) As Boolean
Dim Cliente As Long
Dim Sql4 As String
Dim Sql3 As String
Dim Rs1 As ADODB.Recordset
Dim RS3 As ADODB.Recordset
Dim HoraIni As String
Dim HoraFin As String
Dim SQL1 As String
Dim Sql2 As String

Dim SQLinsert As String
Dim Diferencia As Currency
Dim Precio As Currency
Dim Importe As Currency
Dim CosteKilo As Single

Dim SqlValues As String

Dim Lineas As String

Dim KilosTotales As Long
Dim ImporteLinea As Currency

Dim SqlTrab As String
Dim NTrabajadores As Long

    On Error GoTo eInsertarKilosConfeccionados
    
    InsertarKilosConfeccionados = False

    SQLinsert = "insert into cccabecera (idcontador,fecha,codtraba,codcoste,fechaini,fechafin,codvarie,kilos,importe,costekilo,importetrab) values "

    Diferencia = Round2(DateDiff("n", CDate(Rs!FechaIni), CDate(Rs!FechaFin)) / 60, 2)
    Precio = DevuelveValor("select prhoracoste from straba where codtraba = " & DBSet(Rs!codtraba, "N"))
    Importe = Round2(Diferencia * Precio, 2)

    
    Lineas = LineasActividad(CStr(Rs!codCoste))

    '[Monica]17/12/2012: los kilos los tengo que dividir entre el nro de trabajadores que estan trabajando en ese momento
    SqlTrab = "select count(distinct codtraba) from cctrabaconf where codcoste = " & DBSet(Rs!codCoste, "N")
    SqlTrab = SqlTrab & " and (fechaini between " & DBSet(Rs!FechaIni, "FH") & " and " & DBSet(Rs!FechaFin, "FH") & " or "
    SqlTrab = SqlTrab & " fechafin between " & DBSet(Rs!FechaIni, "FH") & " and " & DBSet(Rs!FechaFin, "FH") & ")"
    
    NTrabajadores = TotalRegistros(SqlTrab)
    
    
    
    'Vamos a palets del periodo agrupando kilosnetos por variedad
                                                                                                                                                                  '30/01/2013:horaini
    SQL1 = "select codvarie, round(sum(pesoneto) / " & NTrabajadores & ",0) kilos from palets inner join palets_variedad on palets.numpalet = palets_variedad.numpalet where (horaiconf between " & DBSet(Rs!FechaIni, "FH") & " and " & DBSet(Rs!FechaFin, "FH")
    SQL1 = SQL1 & " or horafin between " & DBSet(Rs!FechaIni, "FH") & " and " & DBSet(Rs!FechaFin, "FH") & ")"
    If Lineas <> "" Then
        SQL1 = SQL1 & " and linsalida in (" & Lineas & ")"
    End If
'    If Variedades <> "" Then Sql1 = Sql1 & " and codvarie in " & Variedades
    SQL1 = SQL1 & " group by 1 "
    SQL1 = SQL1 & " order by 1 "
    
    KilosTotales = DevuelveValor("select sum(kilos) from (" & SQL1 & ") aaaaa ")
    
    ' Si ese dia no han entrado kilos inserto un registro sin variedad, kilos unicamente con el importe y sin coste por kilo
    If KilosTotales = 0 Then
        CodOrden = CodOrden + 1
    
        SqlValues = "(" & DBSet(CodOrden, "N") & "," & DBSet(Rs!FechaFin, "F") & "," & DBSet(Rs!codtraba, "N") & "," & DBSet(Rs!codCoste, "N")
        SqlValues = SqlValues & "," & DBSet(Rs!FechaIni, "FH") & "," & DBSet(Rs!FechaFin, "FH") & ","
        
        CosteKilo = 0
    
        SqlValues = SqlValues & DBSet(0, "N") & "," & DBSet(0, "N") & "," & DBSet(Importe, "N") & "," & TransformaComasPuntos(ImporteSinFormato(CStr(CosteKilo))) & "," & DBSet(Importe, "N") & "),"
    
    Else
    
        Set Rs1 = New ADODB.Recordset
        Rs1.Open SQL1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        SqlValues = ""
        
        While Not Rs1.EOF
            CodOrden = CodOrden + 1
        
            ImporteLinea = Round2(Importe * Rs1!Kilos / KilosTotales, 2)
        
        
            SqlValues = SqlValues & "(" & DBSet(CodOrden, "N") & "," & DBSet(Rs!FechaFin, "F") & "," & DBSet(Rs!codtraba, "N") & "," & DBSet(Rs!codCoste, "N")
            SqlValues = SqlValues & "," & DBSet(Rs!FechaIni, "FH") & "," & DBSet(Rs!FechaFin, "FH") & ","
            
            CosteKilo = ImporteLinea / DBLet(Rs1!Kilos, "N")
        
            SqlValues = SqlValues & DBSet(Rs1!codvarie, "N") & "," & DBSet(Rs1!Kilos, "N") & "," & DBSet(ImporteLinea, "N") & "," & TransformaComasPuntos(ImporteSinFormato(CStr(CosteKilo))) & "," & DBSet(Importe, "N") & "),"
            
            'Insertamos que palets de traza ha intervenido de cada variedad
            Sql2 = "insert into cclineas2 (idcontador,numpalet) "
                                                                                                                                                                 '30/01/2013:horaini
            Sql2 = Sql2 & " select " & DBSet(CodOrden, "N") & ", palets.numpalet from palets inner join palets_variedad on palets.numpalet = palets_variedad.numpalet where (horaiconf between " & DBSet(Rs!FechaIni, "FH") & " and " & DBSet(Rs!FechaFin, "FH")
            Sql2 = Sql2 & " or horafin between " & DBSet(Rs!FechaIni, "FH") & " and " & DBSet(Rs!FechaFin, "FH") & ")"
            If Lineas <> "" Then Sql2 = Sql2 & " and linsalida in (" & Lineas & ") "
            Sql2 = Sql2 & " and codvarie = " & DBSet(Rs1!codvarie, "N")
            
            conn.Execute Sql2
            
'            'Insertamos/actualizamos kilos confeccionados por variedad
'            ActualizarKilos CStr(RS1!codvarie), 2, CStr(RS1!Kilos)
            
            Rs1.MoveNext
        Wend
        
        Set Rs1 = Nothing
    
    End If
    
    If SqlValues <> "" Then
        SqlValues = Mid(SqlValues, 1, Len(SqlValues) - 1)
        conn.Execute SQLinsert & SqlValues
    End If

    InsertarKilosConfeccionados = True
    Exit Function

eInsertarKilosConfeccionados:
    Mens = Mens & vbCrLf & vbCrLf & Err.Description
End Function


Private Function InsertarKilosTotales(ByRef Rs As ADODB.Recordset, CodOrden As Long, Mens As String) As Boolean
Dim Cliente As Long
Dim Sql4 As String
Dim HoraIni As String
Dim HoraFin As String
Dim SQL1 As String
Dim Sql2 As String
Dim SQLinsert As String
Dim SqlValues As String
Dim Diferencia As Currency
Dim Precio As Currency
Dim Importe As Currency
Dim CosteKilo As Single
Dim Lineas As String

Dim Rs1 As ADODB.Recordset

    On Error GoTo eInsertarKilosTotales
    
    InsertarKilosTotales = False

    SQLinsert = "insert into cccabecera (idcontador,fecha,codtraba,codcoste,fechaini,fechafin,codvarie,kilos,importe,costekilo,importetrab) values "

    Diferencia = Round2(DateDiff("n", CDate(Rs!FechaIni), CDate(Rs!FechaFin)) / 60, 2)
    Precio = DevuelveValor("select prhoracoste from straba where codtraba = " & DBSet(Rs!codtraba, "N"))
    Importe = Round2(Diferencia * Precio, 2)

    Lineas = LineasActividad(CStr(Rs!codCoste))

    CodOrden = CodOrden + 1


    SqlValues = "(" & DBSet(CodOrden, "N") & "," & DBSet(Rs!FechaFin, "F") & "," & DBSet(Rs!codtraba, "N") & "," & DBSet(Rs!codCoste, "N")
    SqlValues = SqlValues & "," & DBSet(Rs!FechaIni, "FH") & "," & DBSet(Rs!FechaFin, "FH") & ","
    
    CosteKilo = 0

    SqlValues = SqlValues & DBSet(0, "N") & "," & DBSet(0, "N") & "," & DBSet(Importe, "N") & "," & TransformaComasPuntos(ImporteSinFormato(CStr(CosteKilo))) & "," & DBSet(Importe, "N") & "),"


    If SqlValues <> "" Then
        SqlValues = Mid(SqlValues, 1, Len(SqlValues) - 1)
        conn.Execute SQLinsert & SqlValues
    End If


    InsertarKilosTotales = True
    Exit Function

eInsertarKilosTotales:
    Mens = Mens & vbCrLf & vbCrLf & Err.Description
End Function


Private Function LineasActividad(codCoste As String) As String
Dim Sql3 As String
Dim RS3 As ADODB.Recordset
Dim Lineas As String

    On Error Resume Next

    LineasActividad = ""

    'lineas de salida de la actividad
    Sql3 = "select codlinea from ccconcostes_lin where codcoste = " & DBSet(codCoste, "N")
    Set RS3 = New ADODB.Recordset
    RS3.Open Sql3, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Lineas = ""
    
    While Not RS3.EOF
        Lineas = Lineas & DBSet(RS3!codlinea, "N") & ","
        
        RS3.MoveNext
    Wend
    
    Set RS3 = Nothing
    If Lineas <> "" Then Lineas = Mid(Lineas, 1, Len(Lineas) - 1)
    
    LineasActividad = Lineas

    
End Function



Private Function ComprobarFicheroNew(fecha As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim CadM As String

    On Error GoTo eComprobarFichero

    ComprobarFicheroNew = False

    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL


    SQL = "select * from cctrabaconf where sinacabar = 0 and (date(fechaini) = " & DBSet(fecha, "F") & " or date(fechafin) = " & DBSet(fecha, "F") & ")"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    While Not Rs.EOF
        ' comprobamos que exista el trabajador
        SQL = DevuelveDesdeBDNew(cAgro, "straba", "codtraba", "codtraba", CStr(DBLet(Rs!codtraba, "N")), "N")
        If SQL = "" Then
            CadM = "No existe el trabajador"
                        
            Sql2 = "insert into tmpinformes (codusu,codigo1,nombre1) values (" & vUsu.Codigo & "," & DBSet(Rs!codtraba, "N") & ","
            Sql2 = Sql2 & DBSet(CadM, "T") & ")"
            
            conn.Execute Sql2
        End If
        
        ' comprobamos que exista el concepto de coste
        SQL = DevuelveDesdeBDNew(cAgro, "ccconcostes", "codcoste", "codcoste", CStr(DBLet(Rs!codCoste, "N")), "N")
        If SQL = "" Then
            CadM = "No existe el Concepto"
                        
            Sql2 = "insert into tmpinformes (codusu,codigo1,nombre1) values (" & vUsu.Codigo & "," & DBSet(Rs!codCoste, "N") & ","
            Sql2 = Sql2 & DBSet(CadM, "T") & ")"
            
            conn.Execute Sql2
        End If
        
        '[Monica]28/05/2012: hay fichajes que son de distinto dia de inicio de tarea que de fin
        If Mid(DBLet(Rs!FechaIni), 1, 10) <> Mid(DBLet(Rs!FechaFin), 1, 10) Then
            CadM = "Coste:" & DBLet(Rs!codCoste, "N") & " F.I:" & Mid(DBLet(Rs!FechaIni, "F"), 1, 10) & "-F.F:" & Mid(DBLet(Rs!FechaFin, "F"), 1, 10)
                        
            Sql2 = "insert into tmpinformes (codusu,codigo1,nombre1) values (" & vUsu.Codigo & "," & DBSet(Rs!codtraba, "N") & ","
            Sql2 = Sql2 & DBSet(CadM, "T") & ")"
            
            conn.Execute Sql2
        
        End If
        
        
        Rs.MoveNext
    Wend

    Set Rs = Nothing
    
    ComprobarFicheroNew = True
    Exit Function
    
eComprobarFichero:
    MuestraError Err.Number, "Comprobar Fichero", Err.Description
End Function


Private Function CalcularCostePalet(Palet As String) As Currency
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim TotalEnvases As String
Dim TotalCostes As String
Dim Valor As Currency
Dim NPalet As Long

    On Error Resume Next


    CalcularCostePalet = 0

    NPalet = DevuelveValor("select codpalet from palets where numpalet = " & DBSet(Palet, "N"))


    'total importes de envases para ese palet
    SQL = "select sum(round(cantidad * "
    If vParamAplic.TipoPrecio = 0 Then 'precio medio ponderado
        SQL = SQL & " preciomp,4))"
    Else 'precio ultima compra
        SQL = SQL & " preciouc,4))"
    End If
    
    SQL = SQL & " from confpale_envases, sartic where codpalet = " & DBSet(NPalet, "N")
    SQL = SQL & " and confpale_envases.codartic = sartic.codartic"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    TotalEnvases = 0
    If Not Rs.EOF Then
        If Rs.Fields(0).Value > 0 Then TotalEnvases = Rs.Fields(0).Value
    End If
    Rs.Close
    Set Rs = Nothing
    
    CalcularCostePalet = TotalEnvases
    
    If Err.Number <> 0 Then
        Err.Clear
    End If

End Function


