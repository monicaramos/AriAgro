VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListNominas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   9000
   Icon            =   "frmListNominas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCalculoHorasProductivas 
      Height          =   3525
      Left            =   0
      TabIndex        =   125
      Top             =   0
      Width           =   5835
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   24
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   135
         Text            =   "Text5"
         Top             =   2190
         Width           =   2955
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   24
         Left            =   1620
         MaxLength       =   2
         TabIndex        =   128
         Top             =   2190
         Width           =   960
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   27
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   126
         Top             =   1290
         Width           =   1005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   25
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   127
         Top             =   1740
         Width           =   990
      End
      Begin VB.CommandButton CmdAcepCalHProd 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2790
         TabIndex        =   129
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelCalHProd 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   3900
         TabIndex        =   130
         Top             =   2760
         Width           =   975
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   20
         Left            =   1290
         MouseIcon       =   "frmListNominas.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar almac�n"
         Top             =   2220
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Almac�n"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   36
         Left            =   570
         TabIndex        =   136
         Top             =   2250
         Width           =   615
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   13
         Left            =   1320
         Picture         =   "frmListNominas.frx":015E
         Top             =   1290
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   41
         Left            =   570
         TabIndex        =   133
         Top             =   1290
         Width           =   450
      End
      Begin VB.Label Label8 
         Caption         =   "C�lculo de Horas Productivas"
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
         Left            =   480
         TabIndex        =   132
         Top             =   480
         Width           =   4725
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Porcentaje"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   38
         Left            =   570
         TabIndex        =   131
         Top             =   1800
         Width           =   765
      End
   End
   Begin VB.Frame FrameHorasTrabajadas 
      Height          =   4455
      Left            =   45
      TabIndex        =   96
      Top             =   90
      Width           =   7425
      Begin VB.CheckBox Check3 
         Caption         =   "Sobre Horas Productivas"
         Height          =   195
         Left            =   600
         TabIndex        =   134
         Top             =   3360
         Width           =   2220
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   5
         Left            =   4560
         TabIndex        =   106
         Top             =   3735
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   3
         Left            =   3480
         TabIndex        =   104
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   19
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   100
         Top             =   1665
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   99
         Top             =   1305
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   18
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   98
         Text            =   "Text5"
         Top             =   1305
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   19
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   97
         Text            =   "Text5"
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   102
         Top             =   2745
         Width           =   1005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   101
         Top             =   2340
         Width           =   1005
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1215
         Index           =   4
         Left            =   5355
         TabIndex        =   112
         Top             =   2250
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2143
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   29
         Left            =   960
         TabIndex        =   111
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   28
         Left            =   960
         TabIndex        =   110
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   27
         Left            =   600
         TabIndex        =   109
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label Label7 
         Caption         =   "Informe de Horas Trabajadas"
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
         TabIndex        =   108
         Top             =   405
         Width           =   5925
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   26
         Left            =   960
         TabIndex        =   107
         Top             =   2400
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   25
         Left            =   960
         TabIndex        =   105
         Top             =   2715
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   24
         Left            =   600
         TabIndex        =   103
         Top             =   2160
         Width           =   450
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   15
         Left            =   1620
         MouseIcon       =   "frmListNominas.frx":01E9
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   1665
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   14
         Left            =   1620
         MouseIcon       =   "frmListNominas.frx":033B
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   1305
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   1575
         Picture         =   "frmListNominas.frx":048D
         Top             =   2745
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   1575
         Picture         =   "frmListNominas.frx":0518
         Top             =   2340
         Width           =   240
      End
   End
   Begin VB.Frame FrameCalibres 
      Height          =   4455
      Left            =   90
      TabIndex        =   52
      Top             =   90
      Width           =   7020
      Begin VB.Frame FrameStockMaxMin 
         Caption         =   "Ordenar por"
         ForeColor       =   &H00972E0B&
         Height          =   975
         Left            =   495
         TabIndex        =   73
         Top             =   3195
         Width           =   2190
         Begin VB.OptionButton Opcion 
            Caption         =   "Calibre"
            Height          =   255
            Index           =   1
            Left            =   495
            TabIndex        =   75
            Top             =   585
            Width           =   975
         End
         Begin VB.OptionButton Opcion 
            Caption         =   "Variedad "
            Height          =   345
            Index           =   0
            Left            =   495
            TabIndex        =   74
            Top             =   225
            Width           =   1290
         End
      End
      Begin VB.CommandButton Command6 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListNominas.frx":05A3
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command5 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListNominas.frx":08AD
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   11
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   62
         Text            =   "Text5"
         Top             =   2760
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   10
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   61
         Text            =   "Text5"
         Top             =   2400
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   60
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   59
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   9
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   58
         Text            =   "Text5"
         Top             =   1635
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   8
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   "Text5"
         Top             =   1275
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   56
         Top             =   1635
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   55
         Top             =   1275
         Width           =   750
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   1
         Left            =   3480
         TabIndex        =   54
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   4560
         TabIndex        =   53
         Top             =   3735
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1215
         Index           =   3
         Left            =   6120
         TabIndex        =   65
         Top             =   1440
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2143
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   11
         Left            =   1560
         MouseIcon       =   "frmListNominas.frx":0BB7
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar art�culo"
         Top             =   2790
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   10
         Left            =   1560
         MouseIcon       =   "frmListNominas.frx":0D09
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar articulo"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   1560
         MouseIcon       =   "frmListNominas.frx":0E5B
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar familia"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   1560
         MouseIcon       =   "frmListNominas.frx":0FAD
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar familia"
         Top             =   1275
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Calibre"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   17
         Left            =   600
         TabIndex        =   72
         Top             =   2160
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   16
         Left            =   960
         TabIndex        =   71
         Top             =   2790
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   960
         TabIndex        =   70
         Top             =   2400
         Width           =   465
      End
      Begin VB.Label Label5 
         Caption         =   "Informe de Calibres"
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
         Left            =   600
         TabIndex        =   69
         Top             =   360
         Width           =   6735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   14
         Left            =   600
         TabIndex        =   68
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   13
         Left            =   960
         TabIndex        =   67
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   12
         Left            =   960
         TabIndex        =   66
         Top             =   1320
         Width           =   465
      End
   End
   Begin VB.Frame FrameMovimientoEnvases 
      Height          =   5985
      Left            =   45
      TabIndex        =   76
      Top             =   0
      Width           =   6840
      Begin VB.CheckBox Check2 
         Caption         =   "Sacar Compras"
         Height          =   195
         Left            =   900
         TabIndex        =   124
         Top             =   5490
         Width           =   2220
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Ordenado por Cliente"
         Height          =   195
         Left            =   900
         TabIndex        =   123
         Top             =   5085
         Width           =   2220
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   23
         Left            =   1830
         MaxLength       =   6
         TabIndex        =   82
         Top             =   3615
         Width           =   885
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   22
         Left            =   1830
         MaxLength       =   6
         TabIndex        =   81
         Top             =   3210
         Width           =   885
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   23
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   119
         Text            =   "Text5"
         Top             =   3615
         Width           =   3195
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   22
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   118
         Text            =   "Text5"
         Top             =   3210
         Width           =   3195
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   21
         Left            =   1845
         MaxLength       =   16
         TabIndex        =   80
         Top             =   2655
         Width           =   1335
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   20
         Left            =   1845
         MaxLength       =   16
         TabIndex        =   79
         Top             =   2265
         Width           =   1335
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   21
         Left            =   3330
         Locked          =   -1  'True
         TabIndex        =   114
         Text            =   "Text5"
         Top             =   2655
         Width           =   2700
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   20
         Left            =   3330
         Locked          =   -1  'True
         TabIndex        =   113
         Text            =   "Text5"
         Top             =   2265
         Width           =   2700
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   15
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   84
         Top             =   4590
         Width           =   1005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   14
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   83
         Top             =   4185
         Width           =   1005
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   13
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   88
         Text            =   "Text5"
         Top             =   1680
         Width           =   3465
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   12
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   87
         Text            =   "Text5"
         Top             =   1320
         Width           =   3465
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   1845
         MaxLength       =   2
         TabIndex        =   78
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   1845
         MaxLength       =   2
         TabIndex        =   77
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   2
         Left            =   4065
         TabIndex        =   85
         Top             =   5340
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   4
         Left            =   5145
         TabIndex        =   86
         Top             =   5355
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   35
         Left            =   900
         TabIndex        =   122
         Top             =   3255
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   34
         Left            =   900
         TabIndex        =   121
         Top             =   3615
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   33
         Left            =   540
         TabIndex        =   120
         Top             =   3015
         Width           =   480
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   19
         Left            =   1470
         MouseIcon       =   "frmListNominas.frx":10FF
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   3615
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   18
         Left            =   1470
         MouseIcon       =   "frmListNominas.frx":1251
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   3255
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   32
         Left            =   915
         TabIndex        =   117
         Top             =   2310
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   31
         Left            =   915
         TabIndex        =   116
         Top             =   2670
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Art�culo"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   30
         Left            =   555
         TabIndex        =   115
         Top             =   2070
         Width           =   555
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   17
         Left            =   1485
         MouseIcon       =   "frmListNominas.frx":13A3
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar art�culo"
         Top             =   2670
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   16
         Left            =   1485
         MouseIcon       =   "frmListNominas.frx":14F5
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar art�culo"
         Top             =   2310
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   1485
         Picture         =   "frmListNominas.frx":1647
         Top             =   4590
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1485
         Picture         =   "frmListNominas.frx":16D2
         Top             =   4185
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   13
         Left            =   1485
         MouseIcon       =   "frmListNominas.frx":175D
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar tipo envase"
         Top             =   1725
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   12
         Left            =   1485
         MouseIcon       =   "frmListNominas.frx":18AF
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar tipo envase"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   23
         Left            =   555
         TabIndex        =   95
         Top             =   4005
         Width           =   450
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   22
         Left            =   915
         TabIndex        =   94
         Top             =   4560
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   21
         Left            =   915
         TabIndex        =   93
         Top             =   4245
         Width           =   465
      End
      Begin VB.Label Label6 
         Caption         =   "Informe de Movimientos de Envases"
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
         Left            =   585
         TabIndex        =   92
         Top             =   360
         Width           =   5430
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Envase"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   20
         Left            =   555
         TabIndex        =   91
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   19
         Left            =   915
         TabIndex        =   90
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   18
         Left            =   915
         TabIndex        =   89
         Top             =   1320
         Width           =   465
      End
   End
   Begin VB.Frame FrameVariedades 
      Height          =   4455
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   8595
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   4560
         TabIndex        =   27
         Top             =   3735
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   3480
         TabIndex        =   26
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   25
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   24
         Top             =   1275
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   7
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Text5"
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   6
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Text5"
         Top             =   1275
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   21
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   20
         Top             =   2355
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "Text5"
         Top             =   2760
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "Text5"
         Top             =   2355
         Width           =   3015
      End
      Begin VB.CommandButton Command2 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListNominas.frx":1A01
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command1 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListNominas.frx":1D0B
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1215
         Index           =   1
         Left            =   6360
         TabIndex        =   28
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2143
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   11
         Left            =   960
         TabIndex        =   36
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   10
         Left            =   960
         TabIndex        =   35
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   9
         Left            =   600
         TabIndex        =   34
         Top             =   1080
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "Informe de Variedades"
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
         Left            =   600
         TabIndex        =   33
         Top             =   360
         Width           =   6735
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   8
         Left            =   960
         TabIndex        =   32
         Top             =   2400
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   7
         Left            =   960
         TabIndex        =   31
         Top             =   2715
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   6
         Left            =   600
         TabIndex        =   30
         Top             =   2160
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Orden del Informe"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   0
         Left            =   6360
         TabIndex        =   29
         Top             =   1200
         Width           =   1260
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   1560
         MouseIcon       =   "frmListNominas.frx":2015
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar familia"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1560
         MouseIcon       =   "frmListNominas.frx":2167
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar familia"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1560
         MouseIcon       =   "frmListNominas.frx":22B9
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar articulo"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1560
         MouseIcon       =   "frmListNominas.frx":240B
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar art�culo"
         Top             =   2400
         Width           =   240
      End
   End
   Begin VB.Frame FrameProveedores 
      Height          =   3420
      Left            =   45
      TabIndex        =   37
      Top             =   90
      Width           =   8595
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   4560
         TabIndex        =   45
         Top             =   2700
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptar3 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   44
         Top             =   2685
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   43
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   42
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   41
         Text            =   "Text5"
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "Text5"
         Top             =   1320
         Width           =   3015
      End
      Begin VB.CommandButton Command4 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListNominas.frx":255D
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command3 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListNominas.frx":2867
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1215
         Index           =   2
         Left            =   6360
         TabIndex        =   46
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2143
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Informe de Proveedores"
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
         Left            =   600
         TabIndex        =   51
         Top             =   360
         Width           =   6735
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   2
         Left            =   960
         TabIndex        =   50
         Top             =   1365
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   960
         TabIndex        =   49
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   48
         Top             =   1125
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Orden del Informe"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   1
         Left            =   6360
         TabIndex        =   47
         Top             =   1200
         Width           =   1260
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1560
         MouseIcon       =   "frmListNominas.frx":2B71
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1560
         MouseIcon       =   "frmListNominas.frx":2CC3
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   1365
         Width           =   240
      End
   End
   Begin VB.Frame FrameClientes 
      Height          =   3420
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   8595
      Begin VB.CommandButton cmdSubir 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListNominas.frx":2E15
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton cmdBajar 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListNominas.frx":311F
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "Text5"
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "Text5"
         Top             =   1365
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   2
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1365
         Width           =   735
      End
      Begin VB.CommandButton cmdAceptar2 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   3
         Top             =   2685
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   4560
         TabIndex        =   4
         Top             =   2700
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1215
         Index           =   0
         Left            =   6360
         TabIndex        =   9
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2143
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1560
         MouseIcon       =   "frmListNominas.frx":3429
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   1725
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1560
         MouseIcon       =   "frmListNominas.frx":357B
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   1365
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Orden del Informe"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   46
         Left            =   6360
         TabIndex        =   14
         Top             =   1200
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   3
         Left            =   600
         TabIndex        =   13
         Top             =   1125
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   5
         Left            =   960
         TabIndex        =   12
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   4
         Left            =   960
         TabIndex        =   11
         Top             =   1365
         Width           =   465
      End
      Begin VB.Label lbltitulo2 
         Caption         =   "Informe de Clientes"
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
         Left            =   600
         TabIndex        =   10
         Top             =   360
         Width           =   6735
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmListNominas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: LAURA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public OpcionListado As Byte
    '==== Listados BASICOS ====
    '=============================
    ' 10 .- Listado de Clientes
    ' 11 .- Listado de Proveedores
    ' 12 .- Listado de Variedades
    ' 13 .- Listado de Calibres
    ' 15 .- Listado de Horas trababajadas
    ' 16 .- Calculo de Horas productivas
    
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar n� oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public Event RectificarFactura(cliente As String, observaciones As String)

Private Conexion As Byte
'1.- Conexi�n a BD Ariges  2.- Conexi�n a BD Conta

Private HaDevueltoDatos As Boolean


Private WithEvents frmPro As frmManProve 'Proveedores
Attribute frmPro.VB_VarHelpID = -1
Private WithEvents frmCli As frmClientes 'Clientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmProd As frmManProductos 'Productos
Attribute frmProd.VB_VarHelpID = -1
Private WithEvents frmVar As frmManVariedad 'Variedades
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmCal As frmManCalibres 'Calibres
Attribute frmCal.VB_VarHelpID = -1
Private WithEvents frmTArt As frmManTipArtic 'tipos de articulos
Attribute frmTArt.VB_VarHelpID = -1
Private WithEvents frmTra As frmManTraba 'mantenimiento de trabajadores
Attribute frmTra.VB_VarHelpID = -1
Private WithEvents frmArt As frmManArtic 'mantenimiento de articulos
Attribute frmArt.VB_VarHelpID = -1
Private WithEvents frmAlm As frmManAlmProp 'mantenimiento de almacenes propios
Attribute frmAlm.VB_VarHelpID = -1

Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadselect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe


Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'n� de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'C�digo para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim tipo As String


Dim PrimeraVez As Boolean
Dim Contabilizada As Byte
Dim ConSubinforme As Boolean


Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub CmdAcepCalHProd_Click()
Dim sql As String

    If txtCodigo(27).Text = "" Then
        MsgBox "Debe introducir obligatoriamente un valor en el campo Fecha para realizar el c�lculo.", vbExclamation
        Exit Sub
    End If
    
    If txtCodigo(25).Text = "" Then
        MsgBox "Debe introducir un porcentaje para realizar el c�lculo.", vbExclamation
        Exit Sub
    End If

    If txtCodigo(24).Text = "" Then
        MsgBox "Debe introducir el almac�n para realizar el c�lculo.", vbExclamation
        Exit Sub
    End If
    
    sql = "select * from horas where fechahora = " & DBSet(txtCodigo(27).Text, "F")
    sql = sql & " and codalmac = " & DBSet(txtCodigo(24), "N")
    sql = sql & " and codtraba in (select codtraba from straba where codsecci = 1)"

    If TotalRegistros(sql) = 0 Then
        MsgBox "No existen registros para esa fecha en el almac�n introducido. Revise.", vbExclamation
        PonerFoco txtCodigo(27)
    Else
        If CalculoHorasProductivas Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
           
            cmdCancelCalHProd_Click
        End If
    End If
End Sub

Private Sub cmdAceptar_Click(Index As Integer)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
    
    
    InicializarVbles
    
    'A�adir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
    numParam = numParam + 1
    
    Select Case Index
       Case 0 'Frame Informe de articulos
            '======== FORMULA  ====================================
            'D/H Producto
            cDesde = Trim(txtCodigo(6).Text)
            cHasta = Trim(txtCodigo(7).Text)
            nDesde = txtNombre(6).Text
            nHasta = txtNombre(7).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{" & tabla & ".codprodu}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHProducto= """) Then Exit Sub
            End If
            
            'D/H Variedad
            cDesde = Trim(txtCodigo(0).Text)
            cHasta = Trim(txtCodigo(1).Text)
            nDesde = txtNombre(0).Text
            nHasta = txtNombre(1).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{" & tabla & ".codvarie}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad= """) Then Exit Sub
            End If
            
            'Obtener el parametro con el ORDEN del Informe
            '---------------------------------------------
        '    numOp = PonerGrupo(1, ListView1.ListItems(1).Text)
        '    numOp = PonerGrupo(2, ListView1.ListItems(2).Text)
        ' ### [Monica] 10/11/2006    he sustituido las dos anteriores instrucciones por la siguiente
            If ListView1(1).SelectedItem = "Producto" Then
                numOp = PonerGrupo(1, ListView1(1).ListItems(1).Text)
                numOp = PonerGrupo(2, ListView1(1).ListItems(2).Text)
            Else
                numOp = PonerGrupo(1, ListView1(1).ListItems(2).Text)
                numOp = PonerGrupo(2, ListView1(1).ListItems(1).Text)
            End If
'            Debug.Print cadParam
            
            cadNombreRPT = "rManVarie.rpt"
            cadTitulo = "Listado de Variedades"
            ConSubinforme = False
            
       Case 1 'Frame Informe de calibres
            '======== FORMULA  ====================================
            'D/H Variedad
            cDesde = Trim(txtCodigo(8).Text)
            cHasta = Trim(txtCodigo(9).Text)
            nDesde = txtNombre(8).Text
            nHasta = txtNombre(9).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{" & tabla & ".codvarie}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad= """) Then Exit Sub
            End If
            
            'D/H Calibre
            cDesde = Trim(txtCodigo(10).Text)
            cHasta = Trim(txtCodigo(11).Text)
            nDesde = txtNombre(10).Text
            nHasta = txtNombre(11).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{" & tabla & ".codcalib}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCalibre= """) Then Exit Sub
            End If
            
            'Obtener el parametro con el ORDEN del Informe
            '---------------------------------------------
        '    numOp = PonerGrupo(1, ListView1.ListItems(1).Text)
        '    numOp = PonerGrupo(2, ListView1.ListItems(2).Text)
        ' ### [Monica] 10/11/2006    he sustituido las dos anteriores instrucciones por la siguiente
            If Opcion(0).Value Then numOp = PonerGrupo(1, ListView1(3).ListItems(2).Text)
            If Opcion(1).Value Then numOp = PonerGrupo(1, ListView1(3).ListItems(1).Text)
            
            cadNombreRPT = "rManCalibres.rpt"
            cadTitulo = "Listado de Calibres"
            ConSubinforme = False
        
        Case 2 ' informe de movimiento de envases
            '******************************************************
            ' SOLO SACAMOS LOS REGISTROS DE LA TABLA ALBARAN_ENVASE
            '******************************************************
            If Me.Check2.Value = 0 Then
                '======== FORMULA  ====================================
                'D/H TIPO DE ENVASE
                cDesde = Trim(txtCodigo(12).Text)
                cHasta = Trim(txtCodigo(13).Text)
                nDesde = txtNombre(12).Text
                nHasta = txtNombre(13).Text
                If Not (cDesde = "" And cHasta = "") Then
                    'Cadena para seleccion Desde y Hasta
                    Codigo = "{sartic.codtipar}"
                    TipCod = "T"
                    If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHTipo=""") Then Exit Sub
                End If
                
                'D/H ARTICULO
                cDesde = Trim(txtCodigo(20).Text)
                cHasta = Trim(txtCodigo(21).Text)
                nDesde = txtNombre(20).Text
                nHasta = txtNombre(21).Text
                If Not (cDesde = "" And cHasta = "") Then
                    'Cadena para seleccion Desde y Hasta
                    Codigo = "{albaran_envase.codartic}"
                    TipCod = "T"
                    If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHArticulo=""") Then Exit Sub
                End If
                
                'D/H CLIENTE
                cDesde = Trim(txtCodigo(22).Text)
                cHasta = Trim(txtCodigo(23).Text)
                nDesde = txtNombre(22).Text
                nHasta = txtNombre(23).Text
                If Not (cDesde = "" And cHasta = "") Then
                    'Cadena para seleccion Desde y Hasta
                    Codigo = "{albaran_envase.codclien}"
                    TipCod = "N"
                    If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCliente=""") Then Exit Sub
                End If
                
                'D/H fecha movimiento
                cDesde = Trim(txtCodigo(14).Text)
                cHasta = Trim(txtCodigo(15).Text)
                If Not (cDesde = "" And cHasta = "") Then
                    'Cadena para seleccion Desde y Hasta
                    Codigo = "{albaran_envase.fechamov}"
                    TipCod = "F"
                    If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
                End If
    
                cadParam = cadParam & "pOrden=" & Me.Check1.Value & "|"
                numParam = numParam + 1
                
                If Me.Check1.Value = 1 Then
                    cadParam = cadParam & "pGroup1={albaran_envase.codclien}|"
                    cadParam = cadParam & "pGroupName1={clientes.nomclien}|"
                Else
                    cadParam = cadParam & "pGroup1={albaran_envase.codartic}|"
                    cadParam = cadParam & "pGroupName1={sartic.nomartic}|"
                End If
                numParam = numParam + 2
    
                cadNombreRPT = "rMovEnvasesRet.rpt"
                cadTitulo = "Informe de Movimientos de Envases"
                ConSubinforme = True
                
                tabla = "(albaran_envase INNER JOIN sartic on albaran_envase.codartic = sartic.codartic)"
                tabla = tabla & " INNER JOIN stipar on sartic.codtipar = stipar.codtipar "
            Else
            '******************************************************
            ' SACAMOS LOS REGISTROS DE LAS TABLAS: ALBARAN_ENVASE Y SMOVAL
            '******************************************************
                InicializarVbles
                cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
                numParam = numParam + 1
                 
                 'D/H TIPO DE ENVASE
                cDesde = Trim(txtCodigo(12).Text)
                cHasta = Trim(txtCodigo(13).Text)
                nDesde = txtNombre(12).Text
                nHasta = txtNombre(13).Text
                If Not (cDesde = "" And cHasta = "") Then
                    'Cadena para seleccion Desde y Hasta
                    Codigo = "{sartic.codtipar}"
                    TipCod = "T"
                    If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHTipo=""") Then Exit Sub
                End If
                
                'D/H ARTICULO
                cDesde = Trim(txtCodigo(20).Text)
                cHasta = Trim(txtCodigo(21).Text)
                nDesde = txtNombre(20).Text
                nHasta = txtNombre(21).Text
                If Not (cDesde = "" And cHasta = "") Then
                    'Cadena para seleccion Desde y Hasta
                    Codigo = "{albaran_envase.codartic}"
                    TipCod = "T"
                    If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHArticulo=""") Then Exit Sub
                End If
                
                'D/H CLIENTE
                cDesde = Trim(txtCodigo(22).Text)
                cHasta = Trim(txtCodigo(23).Text)
                nDesde = txtNombre(22).Text
                nHasta = txtNombre(23).Text
                If Not (cDesde = "" And cHasta = "") Then
                    'Cadena para seleccion Desde y Hasta
                    Codigo = "{albaran_envase.codclien}"
                    TipCod = "N"
                    If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCliente=""") Then Exit Sub
                End If
                
                'D/H fecha movimiento
                cDesde = Trim(txtCodigo(14).Text)
                cHasta = Trim(txtCodigo(15).Text)
                If Not (cDesde = "" And cHasta = "") Then
                    'Cadena para seleccion Desde y Hasta
                    Codigo = "{albaran_envase.fechamov}"
                    TipCod = "F"
                    If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
                End If
                
                'A�adir el parametro de Empresa

                If CargarTablaTemporal Then
                    cadNombreRPT = "rMovEnvasesRetCompras.rpt"
                    cadTitulo = "Informe de Movimientos de Envases"
                    
                    If HayRegParaInforme("tmpinformes", "codusu= " & vUsu.Codigo) Then
                        cadFormula = "{tmpenvasesret.codusu} = " & vUsu.Codigo
                        LlamarImprimir
                    End If
                End If
                Exit Sub
            End If
        Case 3 ' informe de horas trabajadas
            '======== FORMULA  ====================================
            'D/H TRABAJADOR
            cDesde = Trim(txtCodigo(18).Text)
            cHasta = Trim(txtCodigo(19).Text)
            nDesde = txtNombre(18).Text
            nHasta = txtNombre(19).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{horas.codtraba}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHTrabajador=""") Then Exit Sub
            End If
            
            'D/H fecha
            cDesde = Trim(txtCodigo(16).Text)
            cHasta = Trim(txtCodigo(17).Text)
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{horas.fechahora}"
                TipCod = "F"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
            End If
'
'        ' ### [Monica] 10/11/2006    he sustituido las dos anteriores instrucciones por la siguiente
'            If Opcion(2).Value Then numOp = PonerGrupo(1, ListView1(4).ListItems(1).Text)
'            If Opcion(3).Value Then numOp = PonerGrupo(1, ListView1(4).ListItems(2).Text)

            cadParam = cadParam & "pHProductivas=" & Me.Check3.Value & "|"
            numParam = numParam + 1
            
            ConSubinforme = False
            cadNombreRPT = "rManHorasTrab.rpt"
            cadTitulo = "Informe de Horas Trabajadas"
    
    End Select
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(tabla, cadselect) Then
        LlamarImprimir
    End If

End Sub

'Frame Informe Clientes

Private Sub cmdAceptar2_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

    InicializarVbles
    
    'A�adir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    
    'D/H Cliente
    cDesde = Trim(txtCodigo(4).Text)
    cHasta = Trim(txtCodigo(5).Text)
    nDesde = txtNombre(4).Text
    nHasta = txtNombre(5).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        If OpcionListado = 10 Then
            Codigo = "{" & tabla & ".codclien}"
        ElseIf OpcionListado = 14 Then
            Codigo = "{" & tabla & ".gruprove}"
        End If
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCliente= """) Then Exit Sub
    End If
    
    'Obtener el parametro con el ORDEN del Informe
    '---------------------------------------------
'    numOp = PonerGrupo(1, ListView1.ListItems(1).Text)
'    numOp = PonerGrupo(2, ListView1.ListItems(2).Text)
' ### [Monica] 10/11/2006    he sustituido las dos anteriores instrucciones por la siguiente
    numOp = PonerOrden(ListView1(0).SelectedItem.Text)

    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(tabla, cadselect) Then
        cadNombreRPT = "rManClien.rpt"
        cadTitulo = "Listado de Clientes " & tipo
        cadParam = cadParam & "pTipo= """ & tipo & """|"
        numParam = numParam + 1
        ConSubinforme = False
        LlamarImprimir
    End If
End Sub

'Private Sub cmdAceptarFacRect_Click()
'    If DatosOk Then
''        RaiseEvent RectificarFactura(txtCodigo(8).Text, txtCodigo(87))
''        cmdCancel_Click (1)
'        If CrearFacturaRectificativa(txtCodigo(9).Text, txtCodigo(10).Text, txtCodigo(11).Text, txtCodigo(87).Text, txtCodigo(8).Text, txtCodigo(13).Text) = 0 Then
'            MsgBox "Proceso realizado correctamente", vbExclamation
'            cmdCancel_Click (1)
'        End If
'    End If
'End Sub

Private Sub CmdAceptar3_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

    InicializarVbles
    
    'A�adir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    
    'D/H Cliente
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    nDesde = txtNombre(2).Text
    nHasta = txtNombre(3).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".codprove}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHProveedor= """) Then Exit Sub
    End If
    
    'Obtener el parametro con el ORDEN del Informe
    '---------------------------------------------
'    numOp = PonerGrupo(1, ListView1.ListItems(1).Text)
'    numOp = PonerGrupo(2, ListView1.ListItems(2).Text)
' ### [Monica] 10/11/2006    he sustituido las dos anteriores instrucciones por la siguiente
    numOp = PonerOrden(ListView1(2).SelectedItem.Text)

    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(tabla, cadselect) Then
        cadNombreRPT = "rManProve.rpt"
        cadTitulo = "Listado de Proveedores " & tipo
        cadParam = cadParam & "pTipo= """ & tipo & """|"
        numParam = numParam + 1
        ConSubinforme = False
        LlamarImprimir
    End If

End Sub

Private Sub cmdBajar_Click()
'Bajar el item seleccionado del listview2
    BajarItemList Me.ListView1
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdCancelCalHProd_Click()
    Unload Me
End Sub

Private Sub cmdSubir_Click()
    SubirItemList Me.ListView1
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 10 ' Listado de Clientes
                PonerFoco txtCodigo(4)
                
            Case 11 ' Listado de Proveedores
                PonerFoco txtCodigo(2)
            
            Case 12 ' Listado de Variedades
                PonerFoco txtCodigo(6)
        
            Case 13 ' Listado de Calibres
                PonerFoco txtCodigo(8)
                
            Case 14 ' Imforme de Movimientos de calibres
                PonerFoco txtCodigo(12)
            
            Case 15 ' Informe de Horas Trabajadas
                PonerFoco txtCodigo(18)
                
            Case 16 ' calculo de horas productivas
                PonerFoco txtCodigo(27)
            
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
    Set List = New Collection
    For H = 24 To 27
        List.Add H
    Next H
    For H = 1 To 10
        List.Add H
    Next H
    List.Add 12
    List.Add 13
    List.Add 14
    List.Add 15
    List.Add 18
    List.Add 19
    
    
'    For h = 1 To List.Count
'        Me.imgBuscar(List.item(h)).Picture = frmPpal.imgListImages16.ListImages(1).Picture
'    Next h
' ### [Monica] 09/11/2006    he sustituido el anterior
    For H = 0 To imgBuscar.Count - 1
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
     
    
    Set List = Nothing

    'Ocultar todos los Frames de Formulario
    FrameClientes.visible = False
    FrameVariedades.visible = False
    FrameProveedores.visible = False
    FrameCalibres.visible = False
    Me.FrameMovimientoEnvases.visible = False
    Me.FrameHorasTrabajadas.visible = False
    Me.FrameCalculoHorasProductivas.visible = False
    
    '###Descomentar
'    CommitConexion
    
    Select Case OpcionListado
    
    'LISTADOS DE MANTENIMIENTOS BASICOS
    '---------------------
    Case 10 '10: Listado de Clientes
        FrameClienteVisible True, H, W
        CargarListViewOrden (0)
        Me.lbltitulo2.Caption = "Informe de Clientes"
        Me.Label2(3).Caption = "Cliente"
        indFrame = 2
        tabla = "clientes"
    
    Case 11 ' Listado de Proveedores
        FrameProveedoresVisible True, H, W
        CargarListViewOrden (2)
        Me.lbltitulo2.Caption = "Informe de Provedores"
        Me.Label2(3).Caption = "Proveedores"
        indFrame = 0
        tabla = "proveedor"
    
    Case 12 ' Listado de Variedades
        FrameVariedadesVisible True, H, W
        CargarListViewOrden (1)
        Me.lbltitulo2.Caption = "Informe de Variedades"
        Me.Label2(3).Caption = "Variedades"
        indFrame = 0
        tabla = "variedades"
    
    Case 13 ' Listado de Calibres
        FrameCalibresVisible True, H, W
        Opcion(0).Value = True
        CargarListViewOrden (3)
        Me.lbltitulo2.Caption = "Informe de Calibres"
        Me.Label2(3).Caption = "Calibres"
        indFrame = 0
        tabla = "calibres"
        
    Case 14 ' Informe de Movimientos de envases
        FrameMovimientosVisible True, H, W
        indFrame = 0
        tabla = "albaran_envase"
        
    Case 15 ' Informe de Horas Trabajadas
        FrameHorasTrabajadasVisible True, H, W
        CargarListViewOrden (4)
        indFrame = 0
        tabla = "horas"
        
    Case 16 ' Proceso de Calculo de Horas Productivas
        FrameCalculoHorasProductivasVisible True, H, W
        CargarListViewOrden (4)
        indFrame = 0
        tabla = "horas"
        
    End Select
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub


Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtCodigo(CByte(imgFecha(0).Tag) + 14).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmCal_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCol_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFam_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmPro_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmTArt_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1 'Para listados b�sicos
            ' productos
             AbrirFrmProductos (Index)
            
        Case 2, 3 'PROVEEDORES
            AbrirFrmProveedores (Index)
            
        Case 4, 5 'CLIENTES
            AbrirFrmClientes (Index)
            
        Case 6, 7 'Clientes / Proveedores
'            AbrirFrmFamilias (Index)

        Case 8, 9 'VARIEDADES
            AbrirFrmVariedades (Index)
        
        Case 10, 11 'CALIBRES
            AbrirFrmCalibres (Index)
        
        Case 12, 13 'TIPOS DE ENVASES
            AbrirFrmTipEnvases (Index)
    
        Case 14, 15 'Horas trabajadas
            AbrirFrmManTraba (Index)
    
        Case 16, 17 'Articulos
            AbrirFrmManArtic (Index)
        
        Case 18, 19 'Clientes
            AbrirFrmManClien (Index)
    
        Case 20
            AbrirFrmManAlmac (Index)
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub


Private Sub imgFecha_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim obj As Object

    Set frmC = New frmCal
    
    esq = imgFecha(Index).Left
    dalt = imgFecha(Index).Top
        
    Set obj = imgFecha(Index).Container
      
      While imgFecha(Index).Parent.Name <> obj.Name
            esq = esq + obj.Left
            dalt = dalt + obj.Top
            Set obj = obj.Container
      Wend
    
    menu = Me.Height - Me.ScaleHeight 'ac� tinc el heigth del men� i de la toolbar

    frmC.Left = esq + imgFecha(Index).Parent.Left + 30
    frmC.Top = dalt + imgFecha(Index).Parent.Top + imgFecha(Index).Height + menu - 40

    imgFecha(0).Tag = Index '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtCodigo(Index + 14).Text <> "" Then frmC.NovaData = txtCodigo(Index + 14).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtCodigo(CByte(imgFecha(0).Tag) + 14) '<===
    ' ********************************************
End Sub



Private Sub ListView1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'15/02/2007
'    KEYpress KeyAscii
'ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 2: KEYBusqueda KeyAscii, 2 'proveedor desde
            Case 3: KEYBusqueda KeyAscii, 3 'proveedor hasta
            Case 4: KEYBusqueda KeyAscii, 4 'cliente desde
            Case 5: KEYBusqueda KeyAscii, 5 'cliente hasta
            Case 6: KEYBusqueda KeyAscii, 6 'producto desde
            Case 7: KEYBusqueda KeyAscii, 7 'producto hasta
            Case 0: KEYBusqueda KeyAscii, 0 'variedad desde
            Case 1: KEYBusqueda KeyAscii, 1 'variedad hasta
            Case 8: KEYBusqueda KeyAscii, 8 'cliente de la factura rectificativa
            Case 18: KEYBusqueda KeyAscii, 14 'trabajador desde
            Case 19: KEYBusqueda KeyAscii, 15 'trabajador hasta
            Case 20: KEYBusqueda KeyAscii, 16 'articulo desde
            Case 21: KEYBusqueda KeyAscii, 17 'articulo hasta
            Case 24: KEYBusqueda KeyAscii, 20 'almacen para el calculo de horas productivas
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub


Private Sub txtCodigo_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 1 'VARIEDADES
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
        
        Case 2, 3 'PROVEEDORES
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "proveedor", "nomprove", "codprove", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
        
        Case 4, 5 'CLIENTE
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "clientes", "nomclien", "codclien", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
            
            If Index = 8 Then ' en la factura rectificativa el nuevo cliente ha de existir
                If txtCodigo(8).Text <> "" And txtNombre(8).Text = "" Then
                    MsgBox "El cliente introducido no existe. Si introduce n�mero de cliente �ste debe existir.", vbExclamation
                    PonerFoco txtCodigo(8)
                End If
            End If
        
        Case 6, 7 'PRODUCTOS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "productos", "nomprodu", "codprodu", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
        
        Case 8, 9 'VARIEDADES
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
            
        Case 10, 11 'CALIBRES
            If txtCodigo(8).Text = txtCodigo(9).Text And txtCodigo(8).Text <> "" Then
                txtNombre(Index) = DevuelveDesdeBDNew(cAgro, "calibres", "codcalib", "codvarie", txtCodigo(8).Text, "N", , "codcalib", txtCodigo(Index).Text, "N")
            End If
            
        Case 12, 13 'TIPOS DE ENVASES
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "stipar", "nomtipar", "codtipar", "N")
            
        Case 14, 15, 16, 17, 27 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 18, 19 'TRABAJADORES
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "straba", "nomtraba", "codtraba", "N")
            
        Case 20, 21 'ARTICULOS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "sartic", "nomartic", "codartic", "T")
            
        Case 25 ' porcentaje
            If txtCodigo(Index).Text <> "" Then
                 PonerFormatoDecimal txtCodigo(Index), 9
            End If

        Case 24 'ALMACEN
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "salmpr", "nomalmac", "codalmac", "N")

'
'        Case 9 'SERIE de la factura a rectificar
'            txtCodigo(Index).Text = UCase(txtCodigo(Index).Text)
'            If txtCodigo(9).Text <> "" And txtCodigo(10).Text <> "" And txtCodigo(11).Text <> "" Then
'                PonerValoresFactura
'            End If
'
'        Case 10 'FACTURAS de la factura a rectificar
'            txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
'            If txtCodigo(9).Text <> "" And txtCodigo(10).Text <> "" And txtCodigo(11).Text <> "" Then
'                PonerValoresFactura
'            End If
'
'        Case 13 'FECHAS de la factura rectificativa
'            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
    End Select
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim tabla As String
Dim Titulo As String

    'Llamamos a al form
    cad = ""
    Conexion = cAgro    'Conexi�n a BD: Ariges
'    Select Case OpcionListado
'        Case 7 'Traspaso de Almacenes
'            cad = cad & "N� Trasp|scatra|codtrasp|N|0000000|40�Almacen Origen|scatra|almaorig|N|000|20�Almacen Destino|scatra|almadest|N|000|20�Fecha|scatra|fechatra|F||20�"
'            Tabla = "scatra"
'            titulo = "Traspaso Almacenes"
'        Case 8 'Movimientos de Almacen
'            cad = cad & "N� Movim.|scamov|codmovim|N|0000000|40�Almacen|scamov|codalmac|N|000|30�Fecha|scamov|fecmovim|F||30�"
'            Tabla = "scamov"
'            titulo = "Movimientos Almacen"
'        Case 9, 12, 13, 14, 15, 16, 17 '9: Movimientos Articulos
'                   '12: Inventario Articulos
'                   '14:Actualizar Diferencias de Stock Inventariado
'                   '16: Listado Valoracion stock inventariado
'            cad = cad & "C�digo|sartic|codartic|T||30�Denominacion|sartic|nomartic|T||70�"
'            Tabla = "sartic"
'            titulo = "Articulos"
'    End Select
          
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vtabla = tabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        '###A mano
        'frmB.vDevuelve = "0|1|"
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = Titulo
        frmB.vSelElem = 0
'        frmB.vConexionGrid = Conexion
'        frmB.vBuscaPrevia = 1
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

Private Sub FrameClienteVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de clientes
    Me.FrameClientes.visible = visible
    If visible = True Then
        Me.FrameClientes.Top = -90
        Me.FrameClientes.Left = 0
        Me.FrameClientes.Height = 3420
        Me.FrameClientes.Width = 8600
        W = Me.FrameClientes.Width
        H = Me.FrameClientes.Height
    End If
End Sub

Private Sub FrameProveedoresVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de clientes
    Me.FrameProveedores.visible = visible
    If visible = True Then
        Me.FrameProveedores.Top = -90
        Me.FrameProveedores.Left = 0
        Me.FrameProveedores.Height = 3420
        Me.FrameProveedores.Width = 8600
        W = Me.FrameProveedores.Width
        H = Me.FrameProveedores.Height
    End If
End Sub

Private Sub FrameVariedadesVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de clientes
    Me.FrameVariedades.visible = visible
    If visible = True Then
        Me.FrameVariedades.Top = -90
        Me.FrameVariedades.Left = 0
        Me.FrameVariedades.Height = 4820
        Me.FrameVariedades.Width = 8600
        W = Me.FrameVariedades.Width
        H = Me.FrameVariedades.Height
    End If
End Sub

Private Sub FrameCalibresVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de clientes
    Me.FrameCalibres.visible = visible
    If visible = True Then
        Me.FrameCalibres.Top = -90
        Me.FrameCalibres.Left = 0
        Me.FrameCalibres.Height = 4820
        Me.FrameCalibres.Width = 6600
        W = Me.FrameCalibres.Width
        H = Me.FrameCalibres.Height
    End If
End Sub

Private Sub FrameMovimientosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de clientes
    Me.FrameMovimientoEnvases.visible = visible
    If visible = True Then
        Me.FrameMovimientoEnvases.Top = -90
        Me.FrameMovimientoEnvases.Left = 0
        Me.FrameMovimientoEnvases.Height = 5985
        Me.FrameMovimientoEnvases.Width = 6840
        W = Me.FrameMovimientoEnvases.Width
        H = Me.FrameMovimientoEnvases.Height
    End If
End Sub

Private Sub FrameHorasTrabajadasVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de horas trabajadas
    Me.FrameHorasTrabajadas.visible = visible
    If visible = True Then
        Me.FrameHorasTrabajadas.Top = -90
        Me.FrameHorasTrabajadas.Left = 0
        Me.FrameHorasTrabajadas.Height = 4455
        Me.FrameHorasTrabajadas.Width = 7425
        W = Me.FrameHorasTrabajadas.Width
        H = Me.FrameHorasTrabajadas.Height
    End If
End Sub


Private Sub FrameCalculoHorasProductivasVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el calculo de horas productivas
    Me.FrameCalculoHorasProductivas.visible = visible
    If visible = True Then
        Me.FrameCalculoHorasProductivas.Top = -90
        Me.FrameCalculoHorasProductivas.Left = 0
        Me.FrameCalculoHorasProductivas.Height = 3525
        Me.FrameCalculoHorasProductivas.Width = 5835
        W = Me.FrameCalculoHorasProductivas.Width
        H = Me.FrameCalculoHorasProductivas.Height
    End If
End Sub




Private Sub CargarListViewOrden(Index As Integer)
Dim ItmX As ListItem

    'Los encabezados
    ListView1(Index).ColumnHeaders.Clear
    ListView1(Index).ColumnHeaders.Add , , "Campo", 1390

    Select Case Index
        Case 0
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Codigo"
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Alfab�tico"
        Case 1
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Clase"
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Producto"
        Case 2
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Codigo"
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Alfab�tico"
        Case 3
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Calibre"
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Variedad"
        Case 4
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Trabajador"
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Fecha"
    End Select
        
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadselect = ""
    cadParam = ""
    numParam = 0
End Sub

Private Function PonerDesdeHasta(codD As String, codH As String, nomD As String, nomH As String, param As String) As Boolean
'IN: codD,codH --> codigo Desde/Hasta
'    nomD,nomH --> Descripcion Desde/Hasta
'A�ade a cadFormula y cadSelect la cadena de seleccion:
'       "(codigo>=codD AND codigo<=codH)"
' y a�ade a cadParam la cadena para mostrar en la cabecera informe:
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
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .EnvioEMail = False
        .ConSubinforme = ConSubinforme
        .Opcion = OpcionListado
        .Show vbModal
    End With
End Sub

Private Function PonerGrupo(numGrupo As Byte, cadgrupo As String) As Byte
Dim campo As String
Dim nomcampo As String

    campo = "pGroup" & numGrupo & "="
    nomcampo = "pGroup" & numGrupo & "Name="
    PonerGrupo = 0

    Select Case cadgrupo
'        Case "Codigo"
'            cadParam = cadParam & campo & "{" & Tabla & ".codclien}" & "|"
'            cadParam = cadParam & nomcampo & " {" & "scoope" & ".nomcoope}" & "|"
'            cadParam = cadParam & "pTitulo1" & "=""C�digo""" & "|"
'            numParam = numParam + 3
'
'        Case "Alfabetico"
'            cadParam = cadParam & campo & "{" & Tabla & ".tipsocio}" & "|"
'            cadParam = cadParam & nomcampo & " {" & "tiposoci" & ".nomtipso}" & "|"
'            cadParam = cadParam & "pTitulo1" & "=""Colectivo""" & "|"
'            numParam = numParam + 3
'
        
        'Informe de variedades
        Case "Clase"
            cadParam = cadParam & campo & "{" & tabla & ".codclase}" & "|"
            cadParam = cadParam & nomcampo & " {" & "clases" & ".nomclase}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Producto""" & "|"
            numParam = numParam + 3
            
        Case "Producto"
            cadParam = cadParam & campo & "{" & tabla & ".codprodu}" & "|"
            cadParam = cadParam & nomcampo & " {" & "productos" & ".nomprodu}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Clase""" & "|"
            numParam = numParam + 3

        'Informe de calibres
        Case "Variedad"
            cadParam = cadParam & campo & "{" & tabla & ".codvarie}" & "|"
            cadParam = cadParam & nomcampo & " {" & "variedades" & ".nomvarie}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Variedad""" & "|"
            numParam = numParam + 3
            
        Case "Calibre"
            cadParam = cadParam & campo & "{" & tabla & ".codcalib}" & "|"
            cadParam = cadParam & nomcampo & " {" & "calibres" & ".nomcalib}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Calibre""" & "|"
            numParam = numParam + 3
            
'        'Informe de Horas Trabajadas
'        Case "Trabajador"
'            cadParam = cadParam & campo & "{" & Tabla & ".codtraba}" & "|"
'            cadParam = cadParam & nomcampo & " {" & "straba" & ".nomtraba}" & "|"
'            cadParam = cadParam & "pTitulo1" & "=""Fecha""" & "|"
'            numParam = numParam + 3
'
'        Case "Fecha"
'            cadParam = cadParam & "pGroup1=" & "{" & Tabla & ".fechahora}" & "|"
'            cadParam = cadParam & "pGroup1Name=" & " {" & "horas" & ".fechahora}" & "|"
'            cadParam = cadParam & "pTitulo1" & "=""Trabajadores""" & "|"
'            numParam = numParam + 3
        

End Select

End Function

Private Function PonerOrden(cadgrupo As String) As Byte
Dim campo As String
Dim nomcampo As String

    PonerOrden = 0

    Select Case cadgrupo
        Case "Codigo"
            cadParam = cadParam & "Orden" & "= {" & tabla
            Select Case OpcionListado
                Case 10
                    cadParam = cadParam & ".codclien}|"
                Case 11
                    cadParam = cadParam & ".codprove}|"
            End Select
            tipo = "C�digo"
        Case "Alfab�tico"
            cadParam = cadParam & "Orden" & "= {" & tabla
            Select Case OpcionListado
                Case 10
                    cadParam = cadParam & ".nomclien}|"
                Case 11
                    cadParam = cadParam & ".nomprove}|"
            End Select
            tipo = "Alfab�tico"
    End Select
    
    numParam = numParam + 1

End Function

Private Sub AbrirFrmProveedores(indice As Integer)
    indCodigo = indice
    Set frmPro = New frmManProve
    frmPro.DatosADevolverBusqueda = "0|1|"
    frmPro.DeConsulta = True
    frmPro.CodigoActual = txtCodigo(indCodigo)
    frmPro.Show vbModal
    Set frmPro = Nothing
End Sub

Private Sub AbrirFrmProductos(indice As Integer)
    indCodigo = indice
    Set frmProd = New frmManProductos
    frmProd.DatosADevolverBusqueda = "0|1|"
    frmProd.DeConsulta = True
    frmProd.CodigoActual = txtCodigo(indCodigo)
    frmProd.Show vbModal
    Set frmProd = Nothing
End Sub




Private Sub AbrirFrmClientes(indice As Integer)
    indCodigo = indice
    Set frmCli = New frmClientes
    frmCli.DatosADevolverBusqueda = "0|2|"
'    frmCli.DeConsulta = True
'    frmCli.CodigoActual = txtCodigo(indCodigo)
    frmCli.Show vbModal
    Set frmCli = Nothing
End Sub

Private Sub AbrirFrmVariedades(indice As Integer)
    indCodigo = indice
    Set frmVar = New frmManVariedad
    frmVar.DatosADevolverBusqueda = "0|1|"
'    frmCli.DeConsulta = True
'    frmCli.CodigoActual = txtCodigo(indCodigo)
    frmVar.Show vbModal
    Set frmVar = Nothing
End Sub

Private Sub AbrirFrmCalibres(indice As Integer)
    indCodigo = indice
    Set frmCal = New frmManCalibres
    frmCal.DatosADevolverBusqueda = "2|3|"
'    frmCli.DeConsulta = True
'    frmCli.CodigoActual = txtCodigo(indCodigo)
    frmCal.Show vbModal
    Set frmCal = Nothing
End Sub

Private Sub AbrirFrmTipEnvases(indice As Integer)
    indCodigo = indice
    Set frmTArt = New frmManTipArtic
    frmTArt.DatosADevolverBusqueda = "0|1|"
'    frmCli.DeConsulta = True
'    frmCli.CodigoActual = txtCodigo(indCodigo)
    frmTArt.Show vbModal
    Set frmTArt = Nothing
End Sub

Private Sub AbrirFrmManTraba(indice As Integer)
    indCodigo = indice + 4
    Set frmTra = New frmManTraba
    frmTra.DatosADevolverBusqueda = "0|2|"
'    frmCli.DeConsulta = True
'    frmCli.CodigoActual = txtCodigo(indCodigo)
    frmTra.Show vbModal
    Set frmTra = Nothing
End Sub

Private Sub AbrirFrmManArtic(indice As Integer)
    indCodigo = indice + 4
    Set frmArt = New frmManArtic
    frmArt.DatosADevolverBusqueda = "0|1|"
'    frmCli.DeConsulta = True
'    frmCli.CodigoActual = txtCodigo(indCodigo)
    frmArt.Show vbModal
    Set frmArt = Nothing
End Sub

Private Sub AbrirFrmManClien(indice As Integer)
    indCodigo = indice + 4
    Set frmCli = New frmClientes
    frmCli.DatosADevolverBusqueda = "0|2|"
'    frmCli.DeConsulta = True
'    frmCli.CodigoActual = txtCodigo(indCodigo)
    frmCli.Show vbModal
    Set frmCli = Nothing
End Sub

Private Sub AbrirFrmManAlmac(indice As Integer)
    indCodigo = indice + 4
    Set frmAlm = New frmManAlmProp
    frmAlm.DatosADevolverBusqueda = "0|1|"
'    frmCli.DeConsulta = True
'    frmCli.CodigoActual = txtCodigo(indCodigo)
    frmAlm.Show vbModal
    Set frmAlm = Nothing
End Sub

Private Sub AbrirVisReport()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = cadFormula
'        .SoloImprimir = (Me.OptVisualizar(indFrame).Value = 1)
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
        .Opcion = OpcionListado
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

Private Sub PonerValoresFactura()
Dim intconta As String
Dim cad As String
'    txtCodigo(9).Text = RecuperaValor(CadTag, 1)
'    txtCodigo(10).Text = RecuperaValor(CadTag, 2)
'    txtCodigo(11).Text = RecuperaValor(CadTag, 3)
'    txtCodigo(12).Text = RecuperaValor(CadTag, 4)
'    txtNombre(9).Text = RecuperaValor(CadTag, 5)
'    Contabilizada = RecuperaValor(CadTag, 6)
     intconta = "intconta"
     txtCodigo(12).Text = ""
     txtCodigo(12).Text = DevuelveDesdeBDNew(cAgro, "schfac", "codsocio", "letraser", txtCodigo(9).Text, "T", intconta, "numfactu", txtCodigo(10).Text, "N", "fecfactu", txtCodigo(11).Text, "F")
     If txtCodigo(12).Text <> "" Then
        txtNombre(9).Text = PonerNombreDeCod(txtCodigo(12), "ssocio", "nomsocio", "codsocio", "N")
        Contabilizada = CInt(intconta)
     Else
        cad = "No existe la factura. Reintroduzca. " & vbCrLf & vbCrLf
        cad = cad & "   Serie: " & txtCodigo(9).Text & vbCrLf
        cad = cad & "   Factura: " & txtCodigo(10).Text & vbCrLf
        cad = cad & "   Fecha: " & txtCodigo(11).Text & vbCrLf
        cad = cad & vbCrLf
        MsgBox cad, vbExclamation
        txtCodigo(9).Text = ""
        txtCodigo(10).Text = ""
        txtCodigo(11).Text = ""
        PonerFoco txtCodigo(9)
     End If
End Sub


'Private Function DatosOk() As Boolean
'Dim b As Boolean
'Dim Sql As String
'Dim Sql2 As String
'Dim vClien As CSocio
'' a�adido
'Dim Mens As String
'Dim numfactu As String
'Dim numser As String
'Dim Fecha As Date
'
'    b = True
'    If txtCodigo(9).Text = "" Or txtCodigo(10).Text = "" Or txtCodigo(11).Text = "" Then
'        MsgBox "Debe introducir la letra de serie, el n�mero de factura y la fecha de factura para localizar la factura a rectificar", vbExclamation
'        b = False
'    End If
'    If b And vParamAplic.Cooperativa = 2 Then
'        If txtCodigo(8).Text = "" Then
'            MsgBox "Debe introducir el cliente. Reintroduzca.", vbExclamation
'            b = False
'        Else
'            ' obtenemos la cooperativa del anterior cliente y del nuevo pq tienen que coincidir
'            ' anterior cliente
'            Sql = ""
'            Sql = DevuelveDesdeBDNew(cAgro, "ssocio", "codcoope", "codsocio", txtCodigo(12).Text, "N")
'            ' nuevo cliente
'            Sql2 = ""
'            Sql2 = DevuelveDesdeBDNew(cAgro, "ssocio", "codcoope", "codsocio", txtCodigo(8).Text, "N")
'            If Sql <> Sql2 Then
'                MsgBox "El nuevo cliente debe pertenecer al mismo colectivo que el cliente de la factura a rectificar. Reintroduzca.", vbExclamation
'                b = False
'            End If
'        End If
'    End If
'
''    If b And Contabilizada = 1 And vParamAplic.NumeroConta <> 0 And txtCodigo(8).Text <> "" Then 'comprobamos que la cuenta contable del nuevo cliente existe
''        Set vClien = New CSocio
''        If vClien.LeerDatos(txtCodigo(8).Text) Then
''            sql = ""
''            sql = DevuelveDesdeBDNew(cConta, "cuentas", "codmacta", "codmacta", vClien.CuentaConta, "T")
''            If sql = "" Then
''                MsgBox "La cuenta contable del nuevo cliente no existe. Revise", vbExclamation
''                b = False
''            End If
''        End If
''    End If
'
'' a�adido
''    b = True
'
'    If ConTarjetaProfesional(txtCodigo(9).Text, txtCodigo(10).Text, txtCodigo(11).Text) Then
'        MsgBox "Este Factura tiene alguna tarjeta profesional, no se permite hacer la factura rectificativa", vbExclamation
'        b = False
'    Else
'        If txtCodigo(13).Text = "" Then
'            MsgBox "Debe introducir obligatoriamente una Fecha de Facturaci�n.", vbExclamation
'            b = False
'            PonerFoco txtCodigo(13)
'        Else
'                If Not FechaDentroPeriodoContable(CDate(txtCodigo(13).Text)) Then
'                    Mens = "La Fecha de Facturaci�n no es del ejercicio actual ni siguiente. Reintroduzca."
'                    MsgBox Mens, vbExclamation
'                    b = False
'                    PonerFoco txtCodigo(13)
'                Else
'                    'VRS:2.0.1(0)
'                    If Not FechaSuperiorUltimaLiquidacion(CDate(txtCodigo(13).Text)) Then
'                        Mens = "  La Fecha de Facturaci�n es inferior a la �ltima liquidaci�n de Iva. " & vbCrLf & vbCrLf
'                        ' unicamente si el usuario es root el proceso continuar�
'                        If vSesion.Nivel > 0 Then
'                            Mens = Mens & "  El proceso no continuar�."
'                            MsgBox Mens, vbExclamation
'                            b = False
'                            PonerFoco txtCodigo(13)
'                        Else
'                            Mens = Mens & "                        � Desea continuar ?    " & vbCrLf
'                            If MsgBox(Mens, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'                                b = False
'                                PonerFoco txtCodigo(13)
'                            End If
'                        End If
'                    End If
'                    ' la fecha de factura no debe ser inferior a la ultima factura de la serie
'                    numser = "letraser"
'                    numfactu = ""
'                    numfactu = DevuelveDesdeBDNew(cAgro, "stipom", "contador", "codtipom", "FAG", "T", numser)
'                    If numfactu <> "" Then
'                        If FechaFacturaInferiorUltimaFacturaSerieHco(CDate(txtCodigo(13).Text), CLng(numfactu), numser, 0) Then
'                            Mens = "La Fecha de Factura es inferior a la �ltima factura de la serie." & vbCrLf & vbCrLf
'                            Mens = Mens & "                        � Desea continuar ?    " & vbCrLf
'                            If MsgBox(Mens, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'                                b = False
'                                PonerFoco txtCodigo(13)
'                            End If
'                        End If
'                    End If
'                End If
'        End If
'    End If
'
'    DatosOk = b
'
'
'' end a�adido
'    If b And txtCodigo(87).Text = "" Then
'        MsgBox "Para rectificar una factura ha de introducir obligatoriamente un motivo. Reintroduzca", vbExclamation
'        b = False
'    End If
'    DatosOk = b
'
'End Function
'

Private Function ConTarjetaProfesional(letraser As String, numfactu As String, fecfactu As String) As Boolean
Dim sql As String
Dim Rs As ADODB.Recordset

    sql = "select count(*) from slhfac, starje where letraser = " & DBSet(letraser, "T") & " and numfactu = " & DBSet(numfactu, "N")
    sql = sql & " and fecfactu = " & DBSet(fecfactu, "F") & " and starje.tiptarje = 2 and slhfac.numtarje = starje.numtarje "
    
    ConTarjetaProfesional = (TotalRegistros(sql) <> 0)

End Function



Private Function CargarTablaTemporal() As Boolean
Dim sql As String
Dim Sql1 As String
Dim Rs As ADODB.Recordset

    On Error GoTo eCargarTablaTemporal
    
    CargarTablaTemporal = False

    sql = "delete from tmpenvasesret where codusu = " & DBSet(vUsu.Codigo, "N")
    Conn.Execute sql

'select albaran_envase.codartic, albaran_envase.fechamov
'from (albaran_envase inner join sartic on albaran_envase.codartic = sartic.codartic) inner join stipar on sartic.codtipar = stipar.codtipar
'Where stipar.esretornable = 1
'Union
'select smoval.codartic, smoval.fechamov
'from (smoval inner join  sartic on smoval.codartic = sartic.codartic) inner join stipar on sartic.codtipar = stipar.codtipar
'Where stipar.esretornable = 1

    sql = "select " & vUsu.Codigo & ", albaran_envase.codartic, albaran_envase.fechamov, albaran_envase.cantidad, albaran_envase.tipomovi, albaran_envase.numalbar, "
    sql = sql & " albaran_envase.codclien, clientes.nomclien, " & DBSet("ALV", "T")
    sql = sql & " from ((albaran_envase inner join sartic on albaran_envase.codartic = sartic.codartic) inner join stipar on sartic.codtipar = stipar.codtipar) "
    sql = sql & " inner join clientes on albaran_envase.codclien = clientes.codclien "
    sql = sql & " where stipar.esretornable = 1 "
    
    If txtCodigo(12).Text <> "" Then sql = sql & " and stipar.codtipar >= " & DBSet(txtCodigo(12).Text, "N")
    If txtCodigo(13).Text <> "" Then sql = sql & " and stipar.codtipar <= " & DBSet(txtCodigo(13).Text, "N")
    
    If txtCodigo(20).Text <> "" Then sql = sql & " and albaran_envase.codartic >= " & DBSet(txtCodigo(20).Text, "T")
    If txtCodigo(21).Text <> "" Then sql = sql & " and albaran_envase.codartic <= " & DBSet(txtCodigo(21).Text, "T")
    
    If txtCodigo(22).Text <> "" Then sql = sql & " and albaran_envase.codclien >= " & DBSet(txtCodigo(22).Text, "N")
    If txtCodigo(23).Text <> "" Then sql = sql & " and albaran_envase.codclien <= " & DBSet(txtCodigo(23).Text, "N")
    
    If txtCodigo(14).Text <> "" Then sql = sql & " and albaran_envase.fechamov >= " & DBSet(txtCodigo(14).Text, "F")
    If txtCodigo(15).Text <> "" Then sql = sql & " and albaran_envase.fechamov <= " & DBSet(txtCodigo(15).Text, "F")
    
    sql = sql & " union "
    
    sql = sql & "select " & vUsu.Codigo & ", smoval.codartic, smoval.fechamov, smoval.cantidad, smoval.tipomovi, smoval.document, "
    sql = sql & " smoval.codigope, proveedor.nomprove, " & DBSet("ALC", "T")
    sql = sql & " from ((smoval inner join sartic on smoval.codartic = sartic.codartic) inner join stipar on sartic.codtipar = stipar.codtipar) "
    sql = sql & " inner join proveedor on smoval.codigope = proveedor.codprove "
    sql = sql & " where stipar.esretornable = 1 "
    
    If txtCodigo(12).Text <> "" Then sql = sql & " and stipar.codtipar >= " & DBSet(txtCodigo(12).Text, "N")
    If txtCodigo(13).Text <> "" Then sql = sql & " and stipar.codtipar <= " & DBSet(txtCodigo(13).Text, "N")
    
    If txtCodigo(20).Text <> "" Then sql = sql & " and smoval.codartic >= " & DBSet(txtCodigo(20).Text, "T")
    If txtCodigo(21).Text <> "" Then sql = sql & " and smoval.codartic <= " & DBSet(txtCodigo(21).Text, "T")
    
    If txtCodigo(14).Text <> "" Then sql = sql & " and smoval.fechamov >= " & DBSet(txtCodigo(14).Text, "F")
    If txtCodigo(15).Text <> "" Then sql = sql & " and smoval.fechamov <= " & DBSet(txtCodigo(15).Text, "F")

    Sql1 = "insert into tmpenvasesret " & sql
    Conn.Execute Sql1
    
    CargarTablaTemporal = True
    Exit Function
    
eCargarTablaTemporal:
    MuestraError Err.Number, "Carga Tabla Temporal"
End Function

Private Function CalculoHorasProductivas() As Boolean
Dim Rs As ADODB.Recordset
Dim sql As String
Dim Sql1 As String

    On Error GoTo eCalculoHorasProductivas

    CalculoHorasProductivas = False

    sql = "fechahora = " & DBSet(txtCodigo(27).Text, "F") & " and codalmac = " & DBSet(txtCodigo(24), "N")
    sql = sql & " and codtraba in (select codtraba from straba where codsecci = 1)"


    If BloqueaRegistro("horas", sql) Then
        Sql1 = "update horas set horasproduc = round(horasdia * (1 + (" & DBSet(txtCodigo(25), "N") & "/ 100)),2) "
        Sql1 = Sql1 & " where fechahora = " & DBSet(txtCodigo(27).Text, "F")
        Sql1 = Sql1 & " and codalmac = " & DBSet(txtCodigo(24), "N")
        Sql1 = Sql1 & " and codtraba in (select codtraba from straba where codsecci = 1) "
        
        Conn.Execute Sql1
    
        CalculoHorasProductivas = True
    End If

    TerminaBloquear
    Exit Function

eCalculoHorasProductivas:
    MuestraError Err.Number, "Calculo Horas Productivas", Err.Description
    TerminaBloquear
End Function

