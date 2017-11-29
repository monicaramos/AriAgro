VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListadoOfer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   10215
   Icon            =   "frmListadoOfer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameEnvioFacMail 
      Height          =   6015
      Left            =   0
      TabIndex        =   125
      Top             =   0
      Width           =   10215
      Begin VB.CheckBox Check2 
         Caption         =   "Incluir los ya traspasados"
         Enabled         =   0   'False
         Height          =   255
         Left            =   420
         TabIndex        =   159
         Top             =   5310
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.ListBox ListTipoMov 
         Height          =   960
         Index           =   1000
         Left            =   1320
         Style           =   1  'Checkbox
         TabIndex        =   155
         Top             =   4170
         Width           =   4095
      End
      Begin VB.OptionButton optEnvioMail 
         Caption         =   "comercial"
         Height          =   255
         Index           =   0
         Left            =   6570
         TabIndex        =   154
         Top             =   1230
         Width           =   1095
      End
      Begin VB.OptionButton optEnvioMail 
         Caption         =   "administración"
         Height          =   255
         Index           =   1
         Left            =   8160
         TabIndex        =   153
         Top             =   1230
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton cmdEnvioMail 
         Caption         =   "Enviar"
         Height          =   375
         Left            =   7920
         TabIndex        =   143
         Top             =   5370
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   1665
         Index           =   1
         Left            =   5640
         MultiLine       =   -1  'True
         TabIndex        =   141
         Text            =   "frmListadoOfer.frx":000C
         Top             =   3480
         Width           =   4335
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   111
         Left            =   1380
         MaxLength       =   6
         TabIndex        =   128
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   111
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   129
         Text            =   "Text5"
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   110
         Left            =   1380
         MaxLength       =   6
         TabIndex        =   127
         Top             =   1185
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   110
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   126
         Text            =   "Text5"
         Top             =   1185
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   320
         Index           =   0
         Left            =   5640
         TabIndex        =   139
         Text            =   "Text1"
         Top             =   2760
         Width           =   4335
      End
      Begin VB.CheckBox chkMail 
         Caption         =   "Copia remitente"
         Height          =   255
         Left            =   5610
         TabIndex        =   137
         Top             =   1830
         Width           =   1695
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   18
         Left            =   9000
         TabIndex        =   145
         Top             =   5370
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   109
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   131
         Top             =   2778
         Width           =   1080
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   108
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   130
         Top             =   2778
         Width           =   1080
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   106
         Left            =   1320
         MaxLength       =   7
         TabIndex        =   132
         Text            =   "wwwwwww"
         Top             =   3660
         Width           =   1365
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   107
         Left            =   3840
         MaxLength       =   7
         TabIndex        =   133
         Top             =   3660
         Width           =   1365
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Mensaje"
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
         Left            =   5640
         TabIndex        =   152
         Top             =   3180
         Width           =   600
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Factura"
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
         Index           =   20
         Left            =   240
         TabIndex        =   151
         Top             =   4110
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   34
         Left            =   600
         TabIndex        =   150
         Top             =   1800
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   57
         Left            =   1080
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   33
         Left            =   600
         TabIndex        =   149
         Top             =   1185
         Width           =   450
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Index           =   32
         Left            =   240
         TabIndex        =   148
         Top             =   840
         Width           =   495
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   56
         Left            =   1080
         Top             =   1185
         Width           =   240
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Asunto"
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
         Left            =   5640
         TabIndex        =   147
         Top             =   2430
         Width           =   510
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "E-mail"
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
         Index           =   2
         Left            =   5640
         TabIndex        =   146
         Top             =   1230
         Width           =   420
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   18
         Left            =   3120
         TabIndex        =   144
         Top             =   2823
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   34
         Left            =   3600
         Picture         =   "frmListadoOfer.frx":0012
         Top             =   2800
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   33
         Left            =   1080
         Picture         =   "frmListadoOfer.frx":009D
         Top             =   2800
         Width           =   240
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Fact."
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
         Index           =   1
         Left            =   240
         TabIndex        =   142
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   600
         TabIndex        =   140
         Top             =   2823
         Width           =   450
      End
      Begin VB.Label Label14 
         Caption         =   "Envio facturas por mail"
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
         Index           =   16
         Left            =   240
         TabIndex        =   138
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Nº Factura"
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
         Index           =   15
         Left            =   240
         TabIndex        =   136
         Top             =   3360
         Width           =   780
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   600
         TabIndex        =   135
         Top             =   3645
         Width           =   450
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   3360
         TabIndex        =   134
         Top             =   3645
         Width           =   420
      End
   End
   Begin VB.Frame FramePasarHco 
      Height          =   4575
      Left            =   45
      TabIndex        =   108
      Top             =   45
      Width           =   6915
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   50
         Left            =   2340
         MaxLength       =   10
         TabIndex        =   110
         Top             =   2010
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptarHco 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4320
         TabIndex        =   114
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   8
         Left            =   5400
         TabIndex        =   116
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   52
         Left            =   1845
         MaxLength       =   4
         TabIndex        =   112
         Top             =   2760
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   52
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   109
         Text            =   "Text5"
         Top             =   2760
         Width           =   3735
      End
      Begin VB.Label Label3 
         Caption         =   "Pasar Albaran al histórico"
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
         Index           =   4
         Left            =   600
         TabIndex        =   117
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Eliminación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   62
         Left            =   720
         TabIndex        =   115
         Top             =   2010
         Width           =   1245
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   14
         Left            =   2040
         Picture         =   "frmListadoOfer.frx":0128
         Top             =   2010
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Introduzca los siguiente campos para el histórico: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   63
         Left            =   600
         TabIndex        =   113
         Top             =   1200
         Width           =   4245
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Incidencia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   65
         Left            =   720
         TabIndex        =   111
         Top             =   2760
         Width           =   720
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   29
         Left            =   1530
         Top             =   2790
         Width           =   240
      End
   End
   Begin VB.Frame FrameEnvioMail 
      Height          =   1215
      Left            =   0
      TabIndex        =   105
      Top             =   0
      Visible         =   0   'False
      Width           =   6615
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   360
         TabIndex        =   106
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Preparando datos envio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   22
         Left            =   360
         TabIndex        =   107
         Top             =   840
         Width           =   5805
      End
   End
   Begin VB.Frame FrameCambioProveedor 
      Height          =   3225
      Left            =   0
      TabIndex        =   118
      Top             =   0
      Width           =   7035
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   5
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   122
         Text            =   "Text5"
         Top             =   1500
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   121
         Top             =   1500
         Width           =   855
      End
      Begin VB.CommandButton CmdCambioProv 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4350
         TabIndex        =   120
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   5430
         TabIndex        =   119
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Cambio de Proveedor"
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
         Index           =   31
         Left            =   600
         TabIndex        =   124
         Top             =   360
         Width           =   3855
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   4
         Left            =   1440
         Top             =   1500
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
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
         Index           =   30
         Left            =   600
         TabIndex        =   123
         Top             =   1500
         Width           =   750
      End
   End
   Begin VB.Frame FrameEtiqProv 
      Height          =   5325
      Left            =   45
      TabIndex        =   18
      Top             =   45
      Width           =   7035
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   62
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   10
         Top             =   3360
         Width           =   4335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   9
         Left            =   5400
         TabIndex        =   14
         Top             =   4560
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarEtiqProv 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4320
         TabIndex        =   13
         Top             =   4560
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   360
         TabIndex        =   31
         Top             =   3645
         Width           =   6255
         Begin VB.Frame Frame3 
            Caption         =   "e-Mail"
            Height          =   780
            Left            =   1960
            TabIndex        =   34
            Top             =   560
            Width           =   1575
            Begin VB.OptionButton OptMailAdm 
               Caption         =   "Administración"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   36
               Top             =   210
               Value           =   -1  'True
               Width           =   1335
            End
            Begin VB.OptionButton OptMailCom 
               Caption         =   "Compras"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   35
               Top             =   460
               Width           =   1335
            End
         End
         Begin VB.CheckBox chkEmail 
            Caption         =   "Enviar por e-mail"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   12
            Top             =   560
            Width           =   1575
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   63
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   32
            Text            =   "Text5"
            Top             =   105
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   63
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   11
            Top             =   105
            Width           =   855
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   39
            Left            =   1080
            Top             =   120
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Carta"
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
            Index           =   9
            Left            =   240
            TabIndex        =   33
            Top             =   120
            Width           =   405
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   60
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "Text5"
         Top             =   2520
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   60
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   8
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   61
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "Text5"
         Top             =   2865
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   61
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   9
         Top             =   2865
         Width           =   855
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   59
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   7
         Top             =   1845
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   59
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "Text5"
         Top             =   1845
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   58
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   6
         Top             =   1500
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   58
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "Text5"
         Top             =   1500
         Width           =   3735
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "A la atención de:"
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
         Index           =   8
         Left            =   600
         TabIndex        =   25
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "CPostal"
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
         Index           =   5
         Left            =   600
         TabIndex        =   30
         Top             =   2280
         Width           =   540
      End
      Begin VB.Image imgBuscarOfer 
         Enabled         =   0   'False
         Height          =   240
         Index           =   37
         Left            =   1440
         Top             =   2520
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   960
         TabIndex        =   29
         Top             =   2520
         Width           =   450
      End
      Begin VB.Image imgBuscarOfer 
         Enabled         =   0   'False
         Height          =   240
         Index           =   38
         Left            =   1440
         Top             =   2865
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   960
         TabIndex        =   28
         Top             =   2865
         Width           =   420
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   960
         TabIndex        =   24
         Top             =   1845
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   36
         Left            =   1440
         Top             =   1845
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   960
         TabIndex        =   23
         Top             =   1500
         Width           =   450
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
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
         Index           =   2
         Left            =   600
         TabIndex        =   22
         Top             =   1155
         Width           =   750
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   35
         Left            =   1440
         Top             =   1500
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Etiquetas Proveedores"
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
         Index           =   1
         Left            =   600
         TabIndex        =   21
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame FramePedidos 
      Height          =   4455
      Left            =   600
      TabIndex        =   64
      Top             =   240
      Width           =   6075
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   76
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   66
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   75
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   68
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   12
         Left            =   4080
         TabIndex        =   70
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarPedCom 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3000
         TabIndex        =   69
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   74
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   67
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   73
         Left            =   1680
         MaxLength       =   7
         TabIndex        =   65
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Ped."
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
         Index           =   2
         Left            =   600
         TabIndex        =   77
         Top             =   1680
         Width           =   810
      End
      Begin VB.Label Label12 
         Caption         =   "Imprimir otros Pedidos del Proveedor:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   76
         Top             =   2160
         Width           =   3375
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   19
         Left            =   3480
         Picture         =   "frmListadoOfer.frx":01B3
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   840
         TabIndex        =   75
         Top             =   2880
         Width           =   450
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   4
         Left            =   600
         TabIndex        =   74
         Top             =   2520
         Width           =   435
      End
      Begin VB.Label Label12 
         Caption         =   "Informe de Pedido Compras"
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
         Index           =   0
         Left            =   600
         TabIndex        =   73
         Top             =   360
         Width           =   4335
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   18
         Left            =   1320
         Picture         =   "frmListadoOfer.frx":023E
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   3000
         TabIndex        =   72
         Top             =   2880
         Width           =   420
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Nº Pedido"
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
         Index           =   1
         Left            =   600
         TabIndex        =   71
         Top             =   1320
         Width           =   705
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6435
      Top             =   5985
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameCompras 
      Height          =   5205
      Left            =   180
      TabIndex        =   78
      Top             =   90
      Width           =   7035
      Begin VB.Frame Frame9 
         Caption         =   "Agrupar por"
         ForeColor       =   &H00972E0B&
         Height          =   945
         Left            =   480
         TabIndex        =   102
         Top             =   3880
         Width           =   2175
         Begin VB.OptionButton OptCompras 
            Caption         =   "Familia"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   104
            Top             =   225
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton OptCompras 
            Caption         =   "Familia, Artículo"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   103
            Top             =   550
            Width           =   1575
         End
      End
      Begin VB.Frame Frame8 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   360
         TabIndex        =   96
         Top             =   2640
         Width           =   6495
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   94
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   98
            Text            =   "Text5"
            Top             =   360
            Width           =   3855
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   94
            Left            =   1380
            MaxLength       =   16
            TabIndex        =   83
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   95
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   97
            Text            =   "Text5"
            Top             =   705
            Width           =   3855
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   95
            Left            =   1380
            MaxLength       =   16
            TabIndex        =   84
            Top             =   705
            Width           =   735
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Familia"
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
            Index           =   20
            Left            =   240
            TabIndex        =   101
            Top             =   120
            Width           =   480
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   50
            Left            =   1080
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   12
            Left            =   600
            TabIndex        =   100
            Top             =   360
            Width           =   450
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   51
            Left            =   1080
            Top             =   705
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   11
            Left            =   600
            TabIndex        =   99
            Top             =   705
            Width           =   420
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   16
         Left            =   5400
         TabIndex        =   86
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarCompras 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4320
         TabIndex        =   85
         Top             =   4200
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   91
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   80
         Top             =   1605
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   91
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   88
         Text            =   "Text5"
         Top             =   1605
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   90
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   79
         Top             =   1260
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   90
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   87
         Text            =   "Text5"
         Top             =   1260
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   92
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   81
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   93
         Left            =   4140
         MaxLength       =   10
         TabIndex        =   82
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   24
         Left            =   960
         TabIndex        =   95
         Top             =   1605
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   49
         Left            =   1440
         Top             =   1605
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   23
         Left            =   960
         TabIndex        =   94
         Top             =   1260
         Width           =   450
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
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
         Index           =   22
         Left            =   600
         TabIndex        =   93
         Top             =   1035
         Width           =   750
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   48
         Left            =   1440
         Top             =   1260
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Compras por Familia/Artículo"
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
         Index           =   21
         Left            =   600
         TabIndex        =   92
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   88
         Left            =   3360
         TabIndex        =   91
         Top             =   2280
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   27
         Left            =   1455
         Picture         =   "frmListadoOfer.frx":02C9
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   87
         Left            =   600
         TabIndex        =   90
         Top             =   2040
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   83
         Left            =   960
         TabIndex        =   89
         Top             =   2280
         Width           =   450
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   28
         Left            =   3840
         Picture         =   "frmListadoOfer.frx":0354
         Top             =   2280
         Width           =   240
      End
   End
   Begin VB.Frame FramePteRecibir 
      Height          =   5205
      Left            =   45
      TabIndex        =   45
      Top             =   45
      Width           =   7035
      Begin VB.Frame Frame7 
         Caption         =   "Ordenar por"
         ForeColor       =   &H00972E0B&
         Height          =   940
         Left            =   600
         TabIndex        =   61
         Top             =   3960
         Width           =   2055
         Begin VB.OptionButton OptOrdenPed 
            Caption         =   "Nº Pedido"
            Height          =   255
            Left            =   240
            TabIndex        =   63
            Top             =   550
            Width           =   1215
         End
         Begin VB.OptionButton OptOrdenArt 
            Caption         =   "Artículo"
            Height          =   255
            Left            =   240
            TabIndex        =   62
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   360
         TabIndex        =   55
         Top             =   2760
         Width           =   6495
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   68
            Left            =   1380
            MaxLength       =   16
            TabIndex        =   42
            Top             =   705
            Width           =   1095
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   68
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   57
            Text            =   "Text5"
            Top             =   705
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   67
            Left            =   1380
            MaxLength       =   16
            TabIndex        =   41
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   67
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   56
            Text            =   "Text5"
            Top             =   360
            Width           =   3735
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   15
            Left            =   600
            TabIndex        =   60
            Top             =   705
            Width           =   420
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   44
            Left            =   1080
            Top             =   705
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   14
            Left            =   600
            TabIndex        =   59
            Top             =   360
            Width           =   450
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   43
            Left            =   1080
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Artículo"
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
            Index           =   13
            Left            =   240
            TabIndex        =   58
            Top             =   120
            Width           =   540
         End
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   70
         Left            =   4140
         MaxLength       =   10
         TabIndex        =   40
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   69
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   39
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   65
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   47
         Text            =   "Text5"
         Top             =   1380
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   65
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   37
         Top             =   1380
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   66
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   46
         Text            =   "Text5"
         Top             =   1725
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   66
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   38
         Top             =   1725
         Width           =   855
      End
      Begin VB.CommandButton cmdAceptarPte 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4200
         TabIndex        =   43
         Top             =   4440
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   10
         Left            =   5280
         TabIndex        =   44
         Top             =   4440
         Width           =   975
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   16
         Left            =   3840
         Picture         =   "frmListadoOfer.frx":03DF
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   75
         Left            =   960
         TabIndex        =   54
         Top             =   2400
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   74
         Left            =   600
         TabIndex        =   53
         Top             =   2160
         Width           =   435
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   15
         Left            =   1455
         Picture         =   "frmListadoOfer.frx":046A
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   72
         Left            =   3360
         TabIndex        =   52
         Top             =   2400
         Width           =   420
      End
      Begin VB.Label Label9 
         Caption         =   "Material pendiente de recibir"
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
         Index           =   19
         Left            =   600
         TabIndex        =   51
         Top             =   360
         Width           =   4455
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   41
         Left            =   1440
         Top             =   1380
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
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
         Left            =   600
         TabIndex        =   50
         Top             =   1035
         Width           =   750
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   960
         TabIndex        =   49
         Top             =   1380
         Width           =   450
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   42
         Left            =   1440
         Top             =   1725
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   960
         TabIndex        =   48
         Top             =   1725
         Width           =   420
      End
   End
   Begin VB.Frame FrameGenAlbCom 
      Height          =   4455
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6315
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   48
         Left            =   1980
         MaxLength       =   10
         TabIndex        =   1
         Top             =   2115
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   49
         Left            =   1980
         MaxLength       =   10
         TabIndex        =   2
         Top             =   2595
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptarAlbCom 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   3435
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   7
         Left            =   4440
         TabIndex        =   4
         Top             =   3435
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nº Albaran"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   61
         Left            =   840
         TabIndex        =   17
         Top             =   2115
         Width           =   780
      End
      Begin VB.Label Label3 
         Caption         =   "Pasar Pedido a Albaran"
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
         Index           =   3
         Left            =   600
         TabIndex        =   16
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Alb."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   60
         Left            =   840
         TabIndex        =   15
         Top             =   2595
         Width           =   765
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   13
         Left            =   1665
         Picture         =   "frmListadoOfer.frx":04F5
         Top             =   2595
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Introduzca los siguiente campos para el Albaran de compra: "
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
         Index           =   59
         Left            =   600
         TabIndex        =   5
         Top             =   1200
         Width           =   4335
      End
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   51
      Left            =   1125
      MaxLength       =   6
      TabIndex        =   156
      Top             =   0
      Width           =   750
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   51
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   157
      Text            =   "Text5"
      Top             =   0
      Width           =   3735
   End
   Begin VB.Image imgBuscarOfer 
      Height          =   240
      Index           =   28
      Left            =   810
      Top             =   15
      Width           =   240
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Operador"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   64
      Left            =   0
      TabIndex        =   158
      Top             =   0
      Width           =   690
   End
End
Attribute VB_Name = "frmListadoOfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DatoSeleccionado(CadenaSeleccion As String)

Public Opcionlistado As Integer
    '(ver opciones en frmListado)
        
        
        
    '315:  Envio por mail de las facturas
        
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta/pedido a imprimir

Public CodClien As String 'Para seleccionar inicialmente las ofertas del Cliente
                          'en el listado de Recordatorio de Ofertas y de Valoracion de Ofertas

Public FecEntre As String 'Para pasar inicialmente la fecha de entrega de la Oferta que se va a pasar a pedido
                          'como la fecha de entega del PEdido
                          
Private NomTabla As String
Private NomTablaLin As String

Private WithEvents frmMtoCartasOfe As frmFacCartasOferta
Attribute frmMtoCartasOfe.VB_VarHelpID = -1
Private WithEvents frmMtoCliente As frmClientes
Attribute frmMtoCliente.VB_VarHelpID = -1
Private WithEvents frmMtoProve As frmManProve
Attribute frmMtoProve.VB_VarHelpID = -1
Private WithEvents frmMtoSitua As frmManSitua
Attribute frmMtoSitua.VB_VarHelpID = -1
Private WithEvents frmMtoIncid As frmManInciden
Attribute frmMtoIncid.VB_VarHelpID = -1
Private WithEvents frmMtoArtic As frmManArtic
Attribute frmMtoArtic.VB_VarHelpID = -1
Private WithEvents frmMtoFamilia As frmManFamilias
Attribute frmMtoFamilia.VB_VarHelpID = -1
Private WithEvents frmTra As frmManTraba
Attribute frmTra.VB_VarHelpID = -1

'Private WithEvents frmB As frmBuscaGrid  'Busquedas
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
'Private WithEvents frmCP As frmCPostal 'codigo postal
Private WithEvents frmMen As frmMensajes  'Form Mensajes para mostrar las etiquetas a imprimir
Attribute frmMen.VB_VarHelpID = -1



'----- Variables para el INforme ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'cadena con los parametros q se pasan a Crystal R.
Private numParam As Byte
Private cadselect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private Titulo As String 'Titulo informe que se pasa a frmImprimir
Private nomRPT As String 'nombre del fichero .rpt a imprimir
Private conSubRPT As Boolean 'si tiene subinformes para enlazarlos a las tablas correctas
'-------------------------------------

Dim indCodigo As Byte 'indice para txtCodigo
Dim Codigo As String 'Código para FormulaSelection de Crystal Report

Dim PrimeraVez As Boolean


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub chkEmail_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpress KeyAscii
End Sub


Private Sub cmdAceptarAlbCom_Click()
'Solicitar datos para Generar Albaran  a partir de Pedido de Compras
Dim cad As String

    cad = "" 'txtCodigo(47).Text & "|"
    cad = cad & txtCodigo(48).Text & "|"
    cad = cad & txtCodigo(49).Text & "|"
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub



Private Sub cmdAceptarCompras_Click()
'Listados de Compras
Dim campo As String
Dim cad As String
Dim tabla As String

    InicializarVbles
    
    'Pasar nombre de la Empresa como parametro
    cadParam = "|pNomEmpre=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion D/H PROVEEDOR
    '--------------------------------------------
     If txtCodigo(90).Text <> "" Or txtCodigo(91).Text <> "" Then
        campo = "{scafpc.codprove}"
        'Parametro Desde/Hasta Proveedor
        cad = "pDHProve=""Proveedor: "
        If Not PonerDesdeHasta(campo, "N", 90, 91, cad) Then Exit Sub
    End If
   
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    If txtCodigo(92).Text <> "" Or txtCodigo(93).Text <> "" Then
        campo = "{scafpc.fecfactu}"
        cad = "pDHFecha=""Fecha Fact.: "
        If Not PonerDesdeHasta(campo, "F", 92, 93, cad) Then Exit Sub
    End If
    
    tabla = "scafpc"
    If Opcionlistado = 311 Then
        'Cadena para seleccion D/H FAMILIA
        '--------------------------------------------
         If txtCodigo(94).Text <> "" Or txtCodigo(95).Text <> "" Then
            campo = "{sartic.codfamia}"
            'Parametro Desde/Hasta Familia
            cad = "pDHFamilia=""Familia: "
            If Not PonerDesdeHasta(campo, "N", 94, 95, cad) Then Exit Sub
            
            tabla = "( scafpc INNER JOIN slifpc ON scafpc.codprove=slifpc.codprove AND scafpc.numfactu=slifpc.numfactu "
            tabla = tabla & " AND scafpc.fecfactu=slifpc.fecfactu )"
            tabla = tabla & " INNER JOIN sartic ON slifpc.codartic=sartic.codartic "
        End If
    End If
        
    
    'Comprobar si hay registros para mostrar en el informe
    '========================================================
    
    If Not HayRegParaInforme(tabla, cadselect) Then Exit Sub
    
    
    'Abrir el listado
    '=======================================
    'Nombre fichero .rpt a Imprimir
    If Opcionlistado = 311 Then
        If Me.OptCompras(0).Value = True Then
            nomRPT = "rComEstProFam.rpt"
            Titulo = "Listado Compras por Familia"
            conSubRPT = True
        Else
            nomRPT = "rComEstProArt.rpt"
            Titulo = "Listado Compras por Artículo"
            conSubRPT = False
        End If
    Else
        nomRPT = "rComEstProImp.rpt"
        Titulo = "Listado Compras por Proveedor"
        conSubRPT = False

    End If
'    conSubRPT = False
    
    LlamarImprimir
    
End Sub



Private Sub cmdAceptarEtiqProv_Click()
'305: Listado para etiquetas de proveedor
'306: Listado para cartas a proveedor
Dim campo As String

    InicializarVbles
    
    'si es listado de CARTAS/eMAIL a proveedores comprobar que se ha seleccionado
    'una carta para imprimir
    If Opcionlistado = 306 Then
        If txtCodigo(63).Text = "" Then
            MsgBox "Debe seleccionar una carta para imprimir.", vbInformation
            Exit Sub
        End If
        
        'Parametro cod. carta
        cadParam = "|pCodCarta= " & txtCodigo(63).Text & "|"
        numParam = numParam + 1
        
        'Nombre fichero .rpt a Imprimir
        nomRPT = "rComProveCarta.rpt"
        Titulo = "Cartas a Proveedores"
        conSubRPT = True
        
    Else 'ETIQUETAS
        cadParam = "|"
    
        'Nombre fichero .rpt a Imprimir
        nomRPT = "rComProveEtiq.rpt"
        Titulo = "Etiquetas de Proveedores"
        conSubRPT = False
    End If
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion D/H PROVEEDOR
    '--------------------------------------------
     If txtCodigo(58).Text <> "" Or txtCodigo(59).Text <> "" Then
        campo = "{proveedor.codprove}"
        'Parametro Desde/Hasta Proveedor
        If Not PonerDesdeHasta(campo, "N", 58, 59, "") Then Exit Sub
    End If
    
    'Cadena para seleccion D/H COD. POSTAL
    '--------------------------------------------
     If txtCodigo(60).Text <> "" Or txtCodigo(61).Text <> "" Then
        campo = "{proveedor.codpobla}"
        'Parametro Desde/Hasta cod. Postal
        If Not PonerDesdeHasta(campo, "T", 60, 61, "") Then Exit Sub
    End If
    
    '====================================================
        
        
    'Parametro a la Atencion de
    If txtCodigo(62).Text <> "" Then
        cadParam = cadParam & "pAtencion=""Att. " & txtCodigo(62).Text & """|"
        numParam = numParam + 1
    Else
        cadParam = cadParam & "pAtencion=""""|"
        numParam = numParam + 1
    End If
    
    'ver si hay registros seleccionados para mostrar en el informe
    If Not HayRegParaInforme("proveedor", cadselect) Then Exit Sub
    
    Set frmMen = New frmMensajes
    frmMen.cadWHERE = cadselect
    frmMen.OpcionMensaje = 9 'Etiquetas proveedores
    frmMen.Show vbModal
    Set frmMen = Nothing
    If cadselect = "" Then Exit Sub
    
    If Opcionlistado = 306 And Me.chkEMAIL(0).Value = 1 Then
        'Enviarlo por e-mail
        EnviarEMailMulti cadselect, Titulo, "rComProveCarta.rpt", "proveedor" 'email para proveedores
    Else
        LlamarImprimir
    End If
    
End Sub



Private Sub cmdAceptarHco_Click()
'pedir datos para Pasar de Albaranes a historico
Dim cad As String

    'comprobar que todos los camos tienen valor
    If txtCodigo(50).Text = "" Or txtCodigo(52).Text = "" Then  'Or txtCodigo(51).Text = ""
        MsgBox "Debe rellenar todos los campos para pasar al histórico.", vbInformation
        Exit Sub
    End If

    'datos a devolver
    cad = txtCodigo(50).Text & "|"
'    cad = cad & txtCodigo(51).Text & "|"
    cad = cad & txtCodigo(52).Text & "|"
    RaiseEvent DatoSeleccionado(cad)
    Unload Me

End Sub

Private Sub cmdAceptarPedCom_Click()
'55: Informe Pedido de Compras (a Proveedor)
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim devuelve As String, campo As String
Dim CodPed As String
Dim campo1 As String, campo2 As String, campo3 As String
    
    If txtCodigo(73).Text = "" Then 'Nº del Pedido
        MsgBox "Debe seleccionar un Pedido para Imprimir.", vbInformation
        PonerFoco txtCodigo(73)
        Exit Sub
    Else
        NumCod = txtCodigo(73).Text
    End If
    
    If (Opcionlistado = 239) And txtCodigo(76).Text = "" Then
        MsgBox "Debe seleccionar un Pedido y Fecha para Imprimir.", vbInformation
        PonerFoco txtCodigo(76)
        Exit Sub
    End If
    
    
    InicializarVbles
    conSubRPT = True
    
    '===================================================
    '============ PARAMETROS ===========================
    Select Case Opcionlistado
        Case 38
            indRPT = 7 '7: Pedidos de Clientes
            Titulo = "Pedido de Ventas"
        Case 239
            indRPT = 8 '8: Pedidos de Clientes (Historico)
            Titulo = "Hist. Pedido de Venta"
        Case 55
            indRPT = 14 '14: Pedidos a Proveedores
            Titulo = "Pedidos de Compras"
        Case 56
            indRPT = 15
            Titulo = "Hist. Pedidos de Compras"
    End Select
    
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomRPT) Then Exit Sub
     
    If Opcionlistado = 38 Or Opcionlistado = 239 Then
        campo1 = "numpedcl"
        campo2 = "fecpedcl"
        campo3 = "codclien"
    Else
        campo1 = "numpedpr"
        campo2 = "fecpedpr"
        campo3 = "codprove"
    End If
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de PEDIDO
    '--------------------------------------------
    If NumCod <> "" Then
        devuelve = "{" & NomTabla & "." & campo1 & "}=" & Val(NumCod)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        cadselect = cadFormula
        
        If Opcionlistado = 239 Then 'historico ( hay fecha)
            devuelve = "{" & NomTabla & "." & campo2 & "}= Date(" & Year(txtCodigo(76).Text) & "," & Month(txtCodigo(76).Text) & "," & Day(txtCodigo(76).Text) & ")"
            If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
            devuelve = NomTabla & "." & campo2 & "='" & Format(txtCodigo(76).Text, FormatoFecha) & "'"
            If Not AnyadirAFormula(cadselect, devuelve) Then Exit Sub
        End If
        
        'Seleccionar otros PEdidos entre esas FEchas
        If Not (txtCodigo(74).Text = "" And txtCodigo(75).Text = "") Then
            campo = "{" & NomTabla & "." & campo2 & "}"
            devuelve = CadenaDesdeHasta(txtCodigo(74).Text, txtCodigo(75).Text, campo, "F")
            If devuelve = "Error" Then Exit Sub
            If cadFormula <> "" Then
                cadFormula = "(" & cadFormula & " OR " & devuelve & ")"
                cadselect = "((" & cadselect & ") OR " & CadenaDesdeHastaBD(txtCodigo(74).Text, txtCodigo(75).Text, campo, "F") & ")"
            Else
                cadFormula = devuelve
                cadselect = CadenaDesdeHastaBD(txtCodigo(74).Text, txtCodigo(75).Text, campo, "F")
            End If
        
            'Filtrar solo los Pedidos del CLIENTE/PROVEEDOR que las solicita
            If CodClien <> "" Then
                campo = "{" & NomTabla & "." & campo3 & "}=" & CodClien
                If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
                If Not AnyadirAFormula(cadselect, campo) Then Exit Sub
            End If
        End If
    Else
'        'Comprobar si se imprimen varios Pedidos
'        If txtCodigo(3).Text <> "" Or txtCodigo(4).Text <> "" Then
'         'Cadena para seleccion Desde y Hasta FECHA
'         '--------------------------------------------
'            campo = "{" & NomTabla & ".fecpedcl}"
'            devuelve = CadenaDesdeHasta(txtCodigo(3).Text, txtCodigo(4).Text, campo, "F")
'            If devuelve = "Error" Then Exit Sub
'            If Not AnyadirAFormula(cadFormula, devuelve) Then
'                Exit Sub
'            Else
'                devuelve = CadenaDesdeHastaBD(txtCodigo(3).Text, txtCodigo(4).Text, campo, "F")
'                If devuelve = "Error" Then Exit Sub
'                If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
'            End If
'        End If
    End If
    
    If Opcionlistado = 38 Or Opcionlistado = 239 Then
        'Añadimos a los parametros el tipo de IVA que se aplica a ese cliente (para saber si esta exento o no de IVA)
        devuelve = DevuelveDesdeBDNew(cAgro, "clientes", "tipoiva", "codclien", CodClien, "N")
        If devuelve <> "" Then
            cadParam = cadParam & "pTipoIVA=" & devuelve & "|"
            numParam = numParam + 1
        End If
    End If

    'comprobar que hay datos para mostrar en el Informe
    If Not HayRegParaInforme(NomTabla, cadselect) Then Exit Sub
    
    LlamarImprimir
End Sub

Private Sub cmdAceptarPte_Click()
'LIstado Material Pendiente de recibir
Dim Codigo As String
Dim cad As String

    InicializarVbles
    
    'Pasar nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    'Pasar el ORDEN del informe como parametro
    If Opcionlistado = 307 Then
        If Me.OptOrdenArt Then
            cad = "{slippr.codartic}"
        Else
            cad = "{scappr.numpedpr}"
        End If
        cadParam = cadParam & "pOrden=" & cad & "|"
        numParam = numParam + 1
    End If
    
    
    '===================================================
    '================= FORMULA =========================
    'será la cadena WHERE para el Informe
    
    'Cadena para seleccion D/H PROVEEDOR
    '--------------------------------------------
    If txtCodigo(65).Text <> "" Or txtCodigo(66).Text <> "" Then
        Codigo = "{scappr.codprove}"
        If Opcionlistado = 308 Then Codigo = "{scaalp.codprove}"
        cad = "pDHProveedor=""Proveedor: "
        If Not PonerDesdeHasta(Codigo, "N", 65, 66, cad) Then Exit Sub
    End If
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    If txtCodigo(69).Text <> "" Or txtCodigo(70).Text <> "" Then
        Codigo = "{scappr.fecpedpr}"
        If Opcionlistado = 308 Then Codigo = "{scaalp.fechaalb}"
        cad = "pDHFecha=""Fecha Ped.: "
        If Opcionlistado = 308 Then cad = "pDHFecha=""Fecha Alb.: "
        If Not PonerDesdeHasta(Codigo, "F", 69, 70, cad) Then Exit Sub
    End If
    
    If Opcionlistado = 307 Then '307: List. Materia pendiente de recibir
        'Cadena para seleccion D/H ARTICULO
        '--------------------------------------------
        If txtCodigo(67).Text <> "" Or txtCodigo(68).Text <> "" Then
            Codigo = "{slippr.codartic}"
            cad = "pDHArticulo=""Artículo: "
            If Not PonerDesdeHasta(Codigo, "T", 67, 68, cad) Then Exit Sub
        End If
    End If
    
    'Comprobar que hay datos que mostrar antes de Abrir el Informe
    If Opcionlistado = 307 Then
        cad = "scappr INNER JOIN slippr ON scappr.numpedpr=slippr.numpedpr "
        Titulo = "Material Pendiente de recibir"
        nomRPT = "rComPteRecibir.rpt"
    Else
        cad = "scaalp INNER JOIN slialp ON scaalp.numalbar=slialp.numalbar AND scaalp.fechaalb=slialp.fechaalb AND scaalp.codprove=slialp.codprove "
        Titulo = "Pendiente de Factura"
        nomRPT = "rComPteFactura.rpt"
    End If
    
    If Not HayRegParaInforme(cad, cadselect) Then Exit Sub

    'Mostrar el Informe
    conSubRPT = False
    LlamarImprimir
End Sub



Private Sub CmdCambioProv_Click()
    If txtCodigo(5).Text = "" Or Me.txtNombre(5).Text = "" Then
        MsgBox "Seleccione el proveedor", vbExclamation
        Exit Sub
    End If
    
'     'Compruebo si esta bloqueado el proveedor
'    miSQL = DevuelveDesdeBDNew(cAgro, "proveedor", "codsitua", "codprove", txtCodigo(5).Text, "N")
'
'    If Val(miSQL) > 0 Then
'            devuelve = "tipositu"
'            miSQL = DevuelveDesdeBDNew(cAgro, "ssitua", "nomsitua", "codsitua", miSQL, "N", devuelve)
'
'
'            If devuelve = "1" Then 'Cliente Bloqueado por Situación Especial.
'                MsgBox UCase("Proveedor bloqueado por: ") & miSQL & "-" & devuelve, vbInformation, "Situación Especial del proveedor."
'            Else
'                MsgBox miSQL, vbInformation, "Situación Especial del proveedor."
'            End If
'            Exit Sub
'    End If
'
    
    
    CadenaDesdeOtroForm = txtCodigo(5).Text
    Unload Me
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdEnvioMail_Click()
Dim Rs As ADODB.Recordset

    'El proceso constara de varias fases.
    'Fase 1: Montar el select y ver si hay registros
    'Fase 2: Preparar carpetas para los pdf
    'Fase 3: Generar para cada factura (una a una) del select su pdf
    'Fase 4: Enviar por mail, adjuntando los archivos correspondientes
    If Opcionlistado = 315 Then
        If Text1(0).Text = "" Then
            MsgBox "Ponga el asunto", vbExclamation
            Exit Sub
        End If
    Else
        Codigo = ""
        If vParamAplic.PathFacturaE = "" Then
            Codigo = "Falta configurar parametros"
        Else
'            MsgBox vParamAplic.PathFacturaE, vbExclamation
            If Dir(vParamAplic.PathFacturaE & "\", vbDirectory) = "" Then Codigo = "No existe carpeta"
'            MsgBox "todo ok", vbExclamation
        End If
        If Codigo <> "" Then
            MsgBox Codigo, vbExclamation
            Exit Sub
        End If
    End If
    
    'AHora pongo los tipo de facturas
    cadFormula = ""
    cadselect = ""  'ME dira si estan todas o no
    For indCodigo = 0 To Me.ListTipoMov(1000).ListCount - 1
        If Me.ListTipoMov(1000).Selected(indCodigo) Then
            'Esta checkeado
            cadFormula = cadFormula & " OR facturas.codtipom = '" & Trim(Mid(ListTipoMov(1000).List(indCodigo), 1, 3)) & "'"
        Else
            cadselect = "NO"
        End If
    Next indCodigo
    
    If cadFormula = "" Then
        MsgBox "Seleccione algun tipo de factura", vbExclamation
        Exit Sub
    Else
        cadFormula = Mid(cadFormula, 4)
    End If
    'En notabla tendre

    NomTabla = "(" & cadFormula & ")"
    
    
    
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    InicializarVbles
    cadFormula = ""
    cadselect = ""
    
'    'Cadena para seleccion D/H Letra Serie
'    '--------------------------------------------
'    If txtCodigo(0).Text <> "" Or txtCodigo(1).Text <> "" Then
'        Codigo = "facturas.codtipom"
'        'Parametro Desde/Hasta Letra Serie
'        If Not PonerDesdeHasta(Codigo, "T", 0, 1, "") Then Exit Sub
'    End If
        
    'Cadena para seleccion D/H Factura
    '--------------------------------------------
    If txtCodigo(106).Text <> "" Or txtCodigo(107).Text <> "" Then
        Codigo = "facturas.numfactu"
        If Not PonerDesdeHasta(Codigo, "N", 106, 107, "") Then Exit Sub
    End If
    
    'Cadena para seleccion D/H Fecha
    '--------------------------------------------
    If txtCodigo(108).Text <> "" Or txtCodigo(109).Text <> "" Then
        Codigo = "facturas.fecfactu"
        If Not PonerDesdeHasta(Codigo, "F", 108, 109, "") Then Exit Sub
    End If
    
    'Cadena para seleccion D/H Cliente
    '--------------------------------------------
    If txtCodigo(110).Text <> "" Or txtCodigo(111).Text <> "" Then
        Codigo = "facturas.codclien"
        If Not PonerDesdeHasta(Codigo, "N", 110, 111, "") Then Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Eliminamos temporales
    conn.Execute "DELETE from tmpinformes where codusu =" & vUsu.Codigo
    
    If cadselect <> "" Then cadselect = cadselect & " AND "
    cadselect = cadselect & NomTabla
    cadselect = " WHERE " & cadselect

    
    Set Rs = New ADODB.Recordset
    DoEvents
        
    If Opcionlistado = 316 Then
        If Me.Check2.Value = 0 Then
            If cadselect <> "" Then cadselect = cadselect & " AND "
            cadselect = cadselect & " (facturas.enfacturae = 0 )"
        End If
    End If
        
        
    'Ahora insertare en la tabla temporal tmpinformes las facturas que voy a generar pdf
'    Codigo = "insert into tmpinformes (codusu,numalbar,codprove,codartic,numlinea,fechaalb,codalmac,cantidad) "
                                            'codsocio,numfactu, letraser,fecfactu,totalfac
    Codigo = "insert into tmpinformes (codusu,codigo1,importe1, nombre1, fecha1, importe2) "
    Codigo = Codigo & " values ( " & vUsu.Codigo & ","
    
    If Not PrepararCarpetasEnvioMail Then Exit Sub
        
    Screen.MousePointer = vbHourglass

    'Vamos a meter todas las facturas en la tabla temporal para comprobar si tienen mail
    'los clientes
    
    NomTabla = "Select codtipom,numfactu,codclien,fecfactu,totalfac from facturas  " & cadselect
    'El orden vamos a hacerlo por: Tipo documento
    NomTabla = NomTabla & " ORDER BY codtipom, numfactu, fecfactu "
    Rs.Open NomTabla, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    While Not Rs.EOF
        NomTabla = Rs!CodClien & "," & Rs!NumFactu & ",'" & Trim(Rs!codTipoM) & "','" & Format(Rs!FecFactu, FormatoFecha)
        
        'El tipo de informe lo guardare en el ultimo campo
        'El report es el = 12
        NomTabla = NomTabla & "'," & TransformaComasPuntos(CStr(DBLet(Rs!TotalFac, "N"))) & ")"
        conn.Execute Codigo & NomTabla
        NumRegElim = NumRegElim + 1
        Rs.MoveNext
    Wend
    Rs.Close
    
    
    If NumRegElim = 0 Then
        If Opcionlistado = 316 Then
            MsgBox "Ningúna factura para traspasar a FacturaE", vbExclamation
        Else
            MsgBox "Ningun dato a enviar por mail", vbExclamation
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Numero de registros
    NomTabla = NumRegElim
    
    
    If Opcionlistado = 315 Then
        
        'AHora ya tengo todos los datos de las facturas que voy  a imprimir
        'Entonces copruebo si para los clientes si tienen puesto el campo mail o no
        Dim NomCampo As String
        If optEnvioMail(0).Value Then
            'Selecciona mail comercial
            cadselect = "2"  'de maiclie2
            NomCampo = "maiclie2"
        Else
            cadselect = "1"  'de maiclie1
            NomCampo = "maiclie1"
        End If
        '[Monica]27/11/2017: cambiado por un problema de catadau con la sintaxis ¿¿¿¿¿??????
        '                    ahora no agrupo y pongo el having como condicion del where, quito la linea del group by
        cadselect = "Select distinct codclien,maiclie" & cadselect
        cadselect = cadselect & " as email from tmpinformes,clientes where codusu = " & vUsu.Codigo & " and codclien=codigo1"
        cadselect = cadselect & " and " & NomCampo & " is null"
        'cadselect = cadselect & " group by codclien having email is null"
        Rs.Open cadselect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        NumRegElim = 0
        While Not Rs.EOF
            NumRegElim = NumRegElim + 1
            Rs.MoveNext
        Wend
        Rs.Close
        
        If NumRegElim > 0 Then
            If MsgBox("Tiene cliente sin mail. Continuar sin sus datos?", vbQuestion + vbYesNo) = vbNo Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
                
            'Si no salimos borramos
            Rs.Open cadselect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            cadselect = "DELETE from tmpinformes where codusu =" & vUsu.Codigo & " and codigo1 ="
            While Not Rs.EOF
                conn.Execute cadselect & Rs!CodClien
                Rs.MoveNext
            Wend
            Rs.Close
            
            
            cadselect = "Select count(*) from tmpinformes where codusu =" & vUsu.Codigo
            Rs.Open cadselect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            NumRegElim = 0
            If Not Rs.EOF Then
                If Not IsNull(Rs.Fields(0)) Then NumRegElim = DBLet(Rs.Fields(0), "N")
                
            End If
            Rs.Close
            
            If NumRegElim = 0 Then
                'NO hay datos para enviar
                
                Screen.MousePointer = vbDefault
                MsgBox "No hay datos para enviar por mail", vbExclamation
                Exit Sub
            Else
                cadselect = "Hay " & NumRegElim & " facturas para enviar por mail." & vbCrLf & "¿Continuar?"
                If MsgBox(cadselect, vbQuestion + vbYesNo) = vbNo Then NumRegElim = 0
            End If
            If NumRegElim = 0 Then
                Set Rs = Nothing
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            NomTabla = NumRegElim
        
        End If
        
    Else
        cadselect = "Hay " & NumRegElim & " facturas para integrar con facturaE." & vbCrLf & "¿Continuar?"
        If MsgBox(cadselect, vbQuestion + vbYesNo) = vbNo Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
        
        
        
    PonerTamnyosMail True
    MDIppal.visible = False
    'Voy arriesgar.
    'Confio en que no envien por mail mas de 32000 facturas (un integer)
    Label14(22).Caption = "Preparando datos"
    Me.ProgressBar1.Max = CInt(NomTabla)
    Me.ProgressBar1.Value = 0
    
    
    
    NumRegElim = 0
    If GeneracionEnvioMail(Rs) Then NumRegElim = 1
        
    
    'Si ha ido todo bien entonces numregelim=1
    If NumRegElim = 1 Then
        If Opcionlistado = 315 Then
            If optEnvioMail(0).Value Then
                'Selecciona mail comercial
                cadselect = "2"  'de maiclie2
            Else
                cadselect = "1"  'de maiclie1
            End If
            cadselect = "Select nomclien, maiclie" & cadselect
            cadselect = cadselect & " as email,tmpinformes.* from tmpinformes,clientes where codusu = " & vUsu.Codigo & " and codclien=codigo1"
    '        cadSelect = cadSelect & " group by codclien having email is null"
    
            
            frmEMail.DatosEnvio = Text1(0).Text & "|" & Text1(1).Text & "|" & Abs(chkMail.Value) & "|" & cadselect & "|"
            frmEMail.Opcion = 4 'Multienvio de facturacion
            frmEMail.Show vbModal
            
            
            'Para tranquilizar las pantallas, borrar los ficheros generados
            'Confio en que no envien por mail mas de 32000 facturas (un integer)
            Label14(22).Caption = "Restaurando ...."
            Me.ProgressBar1.visible = False
        Else
            'Copiar a parametros
            '
            MsgBox "Proceso finalizado", vbExclamation
                
        End If
        
        Me.Refresh
        DoEvents
        espera 1
        
'[Monica]07/03/2013: solo borramos si volvemos a ejecutar el envio (como en gasolinera)
'        PrepararCarpetasEnvioMail
        Me.ProgressBar1.visible = True
        
    End If
    
    
    'Es para evitar la cantidad de pantallas abriendose y cerrandose
    Me.visible = False
    PonerTamnyosMail False
    espera 1
    Unload Me
    MDIppal.Show

    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case Opcionlistado
            Case 55, 56 '55: Informe de Pedido de Compras (proveedor)
                PonerFoco txtCodigo(73)
            Case 57 '57: Pasar Pedido a Albaran de Compras(Proveedores)
                PonerFoco txtCodigo(47)
            Case 80 'albaran de compras
                PonerFoco txtCodigo(50)
            Case 305, 306 '305: Listado Etiquetas proveedor
                          '306: Listado Cartas a proveedores
                PonerFoco txtCodigo(58)
            Case 307, 308 '307: List. Pendiente de Recibir (COMPRAS)
                          '308: List. Pendiente de Facturar (COMPRAS)
                PonerFoco txtCodigo(65)
                
            Case 310, 311 'Listado Compras por Proveedor/Familia/Articulo
                PonerFoco txtCodigo(90)
                
            Case 315, 316 ' envio de facturas por email y facturacion electronica
                PonerFoco txtCodigo(110)
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim indFrame As Single
Dim devuelve As String
    
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    limpiar Me
    indCodigo = 0
    NomTabla = ""

    'Ocultar todos los Frames de Formulario
    Me.FrameGenAlbCom.visible = False
    Me.FrameEtiqProv.visible = False
    Me.FramePteRecibir.visible = False
    Me.FramePedidos.visible = False
    Me.FrameCompras.visible = False
    Me.FramePasarHco.visible = False
    Me.FrameCambioProveedor.visible = False
    Me.FrameEnvioFacMail.visible = False
    
    CommitConexion
    
    CargarIconos
    
    Select Case Opcionlistado
        'LISTADOS DE FACTURACION
        '-----------------------
        Case 1 ' cambio de proveedor
            PonerFrameCambioProveedorVisible H, W
        
        Case 55, 56
                '55: Informe de Pedido de Compras (Proveedor)
                '56: Informe de Hist. Pedido de Compras (Proveedor)
            PonerFramePedVisible H, W
            indFrame = 12
            If NumCod <> "" Then txtCodigo(73).Text = NumCod
            
            
            
        Case 57 '57: Pedir datos para pasar de Pedido a Albaran (NO IMPRIME LISTADO)
            W = 6315
            H = 4455
            PonerFrameVisible Me.FrameGenAlbCom, True, H, W
            indFrame = 7
            Me.Caption = "Generar Albaran Compras"
            'Poner el trabajador conectado
'--monica
'            Me.txtCodigo(47).Text = PonerTrabajadorConectado(devuelve)
'            Me.txtNombre(47).Text = devuelve
            Me.txtCodigo(49).Text = Format(Now, "dd/mm/yyyy")
        
        
        Case 305, 306 '305: Etiquetas de proveedor
                      '306: Cartas a proveedor
            indFrame = 9
            H = 5325
            W = 7035
            PonerFrameVisible Me.FrameEtiqProv, True, H, W
            Me.Frame2.visible = (Opcionlistado = 306)
            If (Opcionlistado = 306) Then Me.Label9(1).Caption = "Cartas a Proveedores"
            
        Case 307, 308 '307: List. Material Pendiente de recibir (COMPRAS)
                      '308: List. Albaranes ptes de facturar (COMPRAS)
            indFrame = 10
            If Opcionlistado = 307 Then
                Me.Label9(19).Caption = "Material pendiente de recibir"
                H = 5200
            Else
                Me.Label9(19).Caption = "Albaranes pendiente de factura"
                H = 4200
                Me.cmdAceptarPte.Top = 3500
                Me.cmdCancel(10).Top = Me.cmdAceptarPte.Top
            End If
            W = 7035
            PonerFrameVisible Me.FramePteRecibir, True, H, W
            Me.Frame6.visible = (Opcionlistado = 307)
            Me.Frame7.visible = (Opcionlistado = 307)
            
        Case 310, 311 'Listado COMPRAS por proveedor
            indFrame = 16
            H = 5235
            If Opcionlistado = 310 Then
                H = 4325
                Me.cmdAceptarCompras.Top = 3400
                Me.cmdCancel(indFrame).Top = Me.cmdAceptarCompras.Top
                Me.Label9(21).Caption = "Compras por Proveedor"
            End If
            W = 7035
            PonerFrameVisible Me.FrameCompras, True, H, W
            Me.Frame8.visible = (Opcionlistado = 311)
            Me.Frame9.visible = (Opcionlistado = 311)
        
        Case 80, 81 '80: pasar albaranes al historico (ventas)
                        '81: pasar pedidos al historico (ventas)
            H = 4575
            W = 6920
            PonerFrameVisible Me.FramePasarHco, True, H, W
            indFrame = 8
            Me.Caption = "Eliminar"
            Select Case Opcionlistado
                Case 80, 82: Me.Label3(4).Caption = "Pasar Albaran al histórico"
                Case 81: Me.Label3(4).Caption = "Pasar Pedido al histórico"
            End Select
            Me.txtCodigo(50).Text = Format(Now, "dd/mm/yyyy")
'            Me.txtCodigo(51).Text = PonerTrabajadorConectado(devuelve)
'            Me.txtNombre(51).Text = devuelve
            
        Case 315, 316
            indFrame = 18
            
            If Opcionlistado = 316 Then Me.FrameEnvioFacMail.Width = 5535
            
            H = FrameEnvioFacMail.Height
            W = FrameEnvioFacMail.Width
            PonerFrameVisible FrameEnvioFacMail, True, H, W
            CargarComboTipoMov 1000


            chkMail.visible = Opcionlistado = 316 'Solo para facturae
            If Opcionlistado = 316 Then
                cmdEnvioMail.Left = 3240
                cmdCancel(indFrame).Left = 4320
                Label14(16).Caption = "Facturacion E"
                cmdEnvioMail.TabIndex = 474
                Check2.Enabled = True
                Check2.visible = True
            Else
                Label14(16).Caption = "Envio facturas por mail"
                
            End If


    End Select
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
    
End Sub



Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtCodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmMen_DatoSeleccionado(CadenaSeleccion As String)
'formulario de mensajes: frmMensajes
    If CadenaSeleccion <> "" Then
        If Opcionlistado = 305 Or Opcionlistado = 306 Then 'Proveedores
            cadFormula = "{proveedor.codprove} IN [" & CadenaSeleccion & "]"
            cadselect = "proveedor.codprove IN (" & CadenaSeleccion & ")"
        Else 'clientes
            cadFormula = "{clientes.codclien} IN [" & CadenaSeleccion & "]"
            cadselect = "clientes.codclien IN (" & CadenaSeleccion & ")"
        End If
    Else 'no seleccionamos ningun cliente
        cadFormula = ""
        cadselect = ""
    End If
End Sub


Private Sub frmMtoArtic_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Articulos
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoCartasOfe_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Cartas de Oferta
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoCliente_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoFamilia_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Familia de Articulos
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoIncid_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Incidencias
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoProve_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Proveedores
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoSitua_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Situaciones
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de trabajadores
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscarOfer_Click(Index As Integer)

    Select Case Index
        Case 0, 1, 39, 40, 45 'Cod. Carta
            Select Case Index
                Case 0: indCodigo = 2
                Case 1: indCodigo = 13
                Case 39: indCodigo = 63
                Case 40: indCodigo = 64
                Case 45: indCodigo = 81
            End Select
            
            Set frmMtoCartasOfe = New frmFacCartasOferta
            frmMtoCartasOfe.DatosADevolverBusqueda = "0|1|"
            frmMtoCartasOfe.Show vbModal
            Set frmMtoCartasOfe = Nothing
            
        Case 2, 3, 9, 10, 23, 24, 46, 47, 52, 53, 56, 57 'Cod. CLIENTE
            Select Case Index
                Case 2, 3: indCodigo = 7 + Index
                Case 9, 10: indCodigo = 18 + Index
                Case 23, 24: indCodigo = Index + 20
                Case 46, 47: indCodigo = Index + 33
                Case 52, 53: indCodigo = Index + 44
                Case 56, 57: indCodigo = Index + 54
            End Select
            Set frmMtoCliente = New frmClientes
            frmMtoCliente.DatosADevolverBusqueda = "0|2|"
            frmMtoCliente.Show vbModal
            Set frmMtoCliente = Nothing
            
        Case 5, 6, 7, 11, 12, 19, 20, 25, 26  'Cod. AGENTE
            Select Case Index
                Case 4, 5: indCodigo = 7 + Index
                Case 5: indCodigo = 12
                Case 6, 7: indCodigo = 12 + Index
                Case 11, 12: indCodigo = 18 + Index
                Case 19, 20, 25, 26: indCodigo = 20 + Index
            End Select
            If Opcionlistado <> 92 Then
'--monica
'                Set frmMtoAgente = New frmFacAgentesCom
'                frmMtoAgente.DatosaDevolverBusqueda = "0|1|"
'                frmMtoAgente.Show vbModal
'                Set frmMtoAgente = Nothing
            ElseIf Index = 6 Or Index = 7 Then 'Gastos financieros (trabajador)
'--monica
'                Set frmMtoTraba = New frmAdmTrabajadores
'                frmMtoTraba.DatosaDevolverBusqueda = "0|1|"
'                frmMtoTraba.Show vbModal
'                Set frmMtoTraba = Nothing
            End If
            
        Case 8, 28 'cod. TRABAJADOR
            indCodigo = 24
            If Index = 28 Then indCodigo = 51
'--monica
            Set frmTra = New frmManTraba
            frmTra.DatosADevolverBusqueda = "0|1|"
            frmTra.Show vbModal
            Set frmTra = Nothing
            
        Case 13, 14, 30, 31 'cod. ACTIVIDAD
            indCodigo = 20 + Index
            If Index = 30 Or Index = 31 Then indCodigo = Index + 23
'--monica
'            Set frmMtoActiv = New frmFacActividades
'            frmMtoActiv.DatosaDevolverBusqueda = "0|1|"
'            If Not IsNumeric(txtCodigo(indCodigo).Text) Then txtCodigo(indCodigo).Text = ""
'            frmMtoActiv.Show vbModal
'            Set frmMtoActiv = Nothing
            
        Case 15, 16 'cod. ZONA
            indCodigo = 20 + Index
'--monica
'            Set frmMtoZona = New frmFacZonas
'            frmMtoZona.DatosaDevolverBusqueda = "0|1|"
'            frmMtoZona.Show vbModal
'            Set frmMtoZona = Nothing
            
         Case 17, 18 'cod. RUTA
            indCodigo = 20 + Index
'--monica
'            Set frmMtoRuta = New frmFacRutas
'            frmMtoRuta.DatosaDevolverBusqueda = "0|1|"
'            frmMtoRuta.Show vbModal
'            Set frmMtoRuta = Nothing
            
        Case 21, 22, 34 'cod. SITUACION
            indCodigo = 20 + Index
            If Index = 34 Then indCodigo = Index + 23
            Set frmMtoSitua = New frmManSitua
            frmMtoSitua.DatosADevolverBusqueda = "0|1|"
            frmMtoSitua.Show vbModal
            Set frmMtoSitua = Nothing
            
        Case 29 'INCIDENCIAS
            indCodigo = 52
            Set frmMtoIncid = New frmManInciden
            frmMtoIncid.DatosADevolverBusqueda = "0|1|"
            txtCodigo(indCodigo).Text = ""
            frmMtoIncid.Show vbModal
            Set frmMtoIncid = Nothing
            
        Case 32, 33, 37, 38 'Cod POSTAL
            indCodigo = Index + 23
'--monica
'            Set frmCP = New frmCPostal
'            frmCP.DatosaDevolverBusqueda = "0|1|"
'            txtCodigo(indCodigo).Text = ""
'            frmCP.Show vbModal
'            Set frmCP = Nothing
            
        Case 4, 35, 36, 41, 42, 48, 49  'cod. PROVEEDOR
            Select Case Index
                Case 4: indCodigo = Index + 1
                Case 35, 36: indCodigo = Index + 23
                Case 41, 42: indCodigo = Index + 24
                Case 48, 49: indCodigo = Index + 42
            End Select
'            If Index = 35 Or Index = 36 Then indCodigo = Index + 23
'            If Index = 41 Or Index = 42 Then indCodigo = Index + 24
'            If Index = 48 Or Index = 49 Then indCodigo = Index + 42
            Set frmMtoProve = New frmManProve
            frmMtoProve.DatosADevolverBusqueda = "0|1|"
            frmMtoProve.Show vbModal
            Set frmMtoProve = Nothing
            
        Case 43, 44 'cod. ARTICULO
            indCodigo = Index + 24
            Set frmMtoArtic = New frmManArtic
            frmMtoArtic.DatosADevolverBusqueda = "1|" 'Abrimos en Modo Busqueda
            frmMtoArtic.Show vbModal
            Set frmMtoArtic = Nothing
            
        Case 50, 51, 54, 55 'Cod. FAMILIA articulo
            Select Case Index
                Case 50, 51: indCodigo = Index + 44
                Case 54, 55: indCodigo = Index + 46
            End Select
            Set frmMtoFamilia = New frmManFamilias
            frmMtoFamilia.DatosADevolverBusqueda = "0|1|"
            frmMtoFamilia.Show vbModal
            Set frmMtoFamilia = Nothing
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub


Private Sub imgFecha_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   
   '++monica
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim obj As Object

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
   
   Select Case Index
        Case 1 'frameOfertas (indFrame=6)
            indCodigo = 3 'Desde
        Case 2 'frameOfertas (indFrame=6)
            indCodigo = 4 'Hasta
        Case 3 'frameRecordatorio Oferta
            indCodigo = 7 '(Desde)
        Case 4 'frameRecordatorio Oferta
            indCodigo = 8 '(Hasta)
        Case 5 'frameEfectuadas
            indCodigo = 16 'Desde
        Case 6 'frameEfectuadas
            indCodigo = 17 'Hasta
        Case 7 'frameTraspasoHco
            indCodigo = 22 'Desde
        Case 8 'frameTraspasoHco
            indCodigo = 23 'hasta
        Case 9, 10 'FrameGenerarPedido
            indCodigo = Index + 16
        Case 11, 12 'Frame Clientes Inactivos
            indCodigo = 20 + Index
        Case 13 'frame pasar pedido a Albaran de compras (a proveedor)
            indCodigo = 49
        Case 14
            indCodigo = 50
        Case 15, 16
            indCodigo = Index + 54
        Case 17 'Frame Factura Rectificariva
            indCodigo = 72
        Case 18, 19 'Ped. Compras
            indCodigo = Index + 56
        Case 20, 21 'Carta Pedidos
            indCodigo = Index + 57
        Case 22: indCodigo = Index + 60
        Case 23, 24 'Reimprimir facturas
            indCodigo = Index + 62
        Case 25, 26 'Cierre caja TPV
            indCodigo = Index + 63
        Case 27, 28 'Listados estadistica compras
            indCodigo = Index + 65
        Case 29, 30 'Estadistica ventas por familia
            indCodigo = Index + 69
   
        Case 31, 32 'Impresion etiq. clientes. Desde / hasta factura
            indCodigo = Index + 73
        Case 33, 34
            indCodigo = Index + 75
   End Select
   
   
   PonerFormatoFecha txtCodigo(indCodigo)
   If txtCodigo(indCodigo).Text <> "" Then frmF.NovaData = CDate(txtCodigo(indCodigo).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub ListTipoMov_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 33 And KeyCode = 38 Then Exit Sub
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim tabla As String
Dim codCampo As String, NomCampo As String
Dim TipCampo As String, Formato As String
Dim Titulo As String
Dim EsNomCod As Boolean


    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    EsNomCod = False
    TipCampo = "N" 'Casi todos son numericos
    
    Select Case Index
        'FECHA Desde Hasta
        Case 3, 4, 7, 8, 16, 17, 22, 23, 25, 26, 31, 32, 49, 50, 69, 70, 72, 74, 75, 77, 78, 82, 85, 86, 88, 89, 92, 93, 98, 99, 104, 105, 108, 109
            If txtCodigo(Index).Text = "" Then Exit Sub
            PonerFormatoFecha txtCodigo(Index)
            
            'Fecha entrega para Pedido. Poner la semana
            If Index = 26 Then
                'Comprobar que fecha entrega es posterior a la del pedido
                If Not EsFechaIgualPosterior(txtCodigo(25).Text, txtCodigo(26).Text, True, "La Fecha de Entrega debe ser posterior a la Fecha del Pedido.") Then
                    txtCodigo(Index).Text = ""
                    PonerFoco txtCodigo(Index)
                Else
                    txtNombre(4).Text = CalculaSemana(CDate(txtCodigo(26).Text))
                End If
            End If
            
        Case 1, 6, 20, 21, 71, 83, 84, 106, 107 'Nº de OFERTA/FACTURA
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
            End If
        
        Case 2, 13, 63, 64, 81 'CARTA de la Oferta
            EsNomCod = True
            tabla = "scartas"
            codCampo = "codcarta"
            NomCampo = "descarta"
            Formato = "000"
            Titulo = "cod. de Carta"
                    
        Case 9, 10, 27, 28, 43, 44, 79, 80, 96, 97, 110, 111 'Cod. CLIENTE
            EsNomCod = True
            tabla = "clientes"
            codCampo = "codclien"
            NomCampo = "nomclien"
            TipCampo = "N"
            Formato = "000000"
            Titulo = "Cliente"

        Case 11, 12, 18, 19, 29, 30, 39, 40, 45, 46 'Cod. AGENTE
            EsNomCod = True
            Formato = "0000"
            If Opcionlistado = 92 Then 'Gastos tecnicos
                If Index = 18 Or Index = 19 Then
                    'cod agente / cod. trabajador
                    tabla = "straba"
                    codCampo = "codtraba"
                    NomCampo = "nomtraba"
                    Titulo = "Trabajador"
                End If
            Else
                tabla = "sagent"
                codCampo = "codagent"
                NomCampo = "nomagent"
                Titulo = "Agente"
            End If
        
        Case 24, 47, 51 'Cod. TRABAJADOR
            EsNomCod = True
            tabla = "straba"
            codCampo = "codtraba"
            NomCampo = "nomtraba"
            Formato = "0000"
            Titulo = "Trabajador"
            
        Case 33, 34, 53, 54 'Cod ACTIVIDAD
            EsNomCod = True
            tabla = "sactiv"
            codCampo = "codactiv"
            NomCampo = "nomactiv"
            Formato = "000"
            Titulo = "Actividad de Cliente"
            
        Case 35, 36 'cod ZONA
            EsNomCod = True
            tabla = "szonas"
            codCampo = "codzonas"
            NomCampo = "nomzonas"
            Formato = "000"
            Titulo = "Zona de Cliente"
            
        Case 37, 38 'cod RUTA
            EsNomCod = True
            tabla = "srutas"
            codCampo = "codrutas"
            NomCampo = "nomrutas"
            Formato = "000"
            Titulo = "Ruta de Asistencia"
                        
        Case 41, 42, 57 'cod SITUACION
            EsNomCod = True
            tabla = "ssitua"
            codCampo = "codsitua"
            NomCampo = "nomsitua"
            Formato = "00"
            Titulo = "Situación Especial"
            
        Case 52 'cod. Incidencias
            EsNomCod = True
            tabla = "inciden"
            codCampo = "codincid"
            NomCampo = "nomincid"
            TipCampo = "T"
            Titulo = "Incidencias"
            
        Case 55, 56, 60, 61 'cod POSTAL
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "scpostal", "provincia", "cpostal", "CPostal")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = txtCodigo(Index).Text
            
         Case 5, 58, 59, 65, 66, 90, 91 'Cod. PROVEEDOR
            EsNomCod = True
            tabla = "proveedor"
            codCampo = "codprove"
            NomCampo = "nomprove"
            TipCampo = "N"
            Formato = "000000"
            Titulo = "Proveedor"
            
        Case 67, 68 'cod. ARTICULO
            EsNomCod = True
            tabla = "sartic"
            codCampo = "codartic"
            NomCampo = "nomartic"
            TipCampo = "T"
            Titulo = "Artículo"
            
        Case 73  'Nº de Pedido de Compras
            If txtCodigo(Index).Text = "" Then Exit Sub
            If Opcionlistado = 55 Or Opcionlistado = 56 Then
                NomCampo = "numpedpr"
                Titulo = "Proveedor"
            Else
                NomCampo = "numpedcl"
                Titulo = "Cliente"
            End If
            NomCampo = DevuelveDesdeBDNew(cAgro, NomTabla, NomCampo, NomCampo, txtCodigo(Index).Text, "N")
            If NomCampo = "" Then
                MsgBox "No existe el Nº de Pedido de " & Titulo & ": " & txtCodigo(Index).Text, vbInformation
                txtCodigo(Index).Text = ""
                PonerFoco txtCodigo(Index)
            Else
                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
            End If
            
        Case 94, 95, 100, 101 'cod. FAMILIA articulos
            EsNomCod = True
            tabla = "sfamia"
            codCampo = "codfamia"
            NomCampo = "nomfamia"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Familia"
    End Select
    
    If EsNomCod Then
        If TipCampo = "N" Then
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), tabla, NomCampo, codCampo, TipCampo)
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, Formato)
            Else
                txtNombre(Index).Text = ""
            End If
        Else
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), tabla, NomCampo, codCampo, TipCampo)
        End If
    End If
End Sub





Private Function AnyadirParametroDH(cad As String, indD As Byte, indH As Byte) As String
On Error Resume Next
    If txtCodigo(indD).Text <> "" Then
        cad = cad & "desde " & txtCodigo(indD).Text
        If txtNombre(indD).Text <> "" Then cad = cad & " - " & txtNombre(indD).Text
    End If
    If txtCodigo(indH).Text <> "" Then
        cad = cad & "  hasta " & txtCodigo(indH).Text
        If txtNombre(indH).Text <> "" Then cad = cad & " - " & txtNombre(indH).Text
    End If
    AnyadirParametroDH = cad
    If Err.Number <> 0 Then Err.Clear
End Function



Private Sub InicializarVbles()
    cadFormula = ""
    cadselect = ""
    cadParam = ""
    numParam = 0
End Sub


Private Function PonerDesdeHasta(campo As String, Tipo As String, indD As Byte, indH As Byte, param As String) As Boolean
Dim devuelve As String
Dim cad As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    
    'para MySQL
    If Tipo <> "F" Then
        If Not AnyadirAFormula(cadselect, devuelve) Then Exit Function
    Else
        'Fecha para la Base de Datos
        cad = CadenaDesdeHastaBD(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
        If Not AnyadirAFormula(cadselect, cad) Then Exit Function
    End If
    
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            cadParam = cadParam & AnyadirParametroDH(param, indD, indH) & """|"
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
        .Opcion = Opcionlistado
        .Titulo = Titulo
        .NombreRPT = nomRPT
        .ConSubInforme = conSubRPT
        .Show vbModal
    End With
End Sub




Private Sub EnviarEMailMulti(cadWHERE As String, cadTit As String, cadRpt As String, cadTABLA As String)
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim cad1 As String, cad2 As String, Lista As String
Dim Cont As Integer

On Error GoTo EEnviar

    Screen.MousePointer = vbHourglass
    
    If cadTABLA = "proveedor" Then
        'seleccionamos todos los proveedores a los que queremos enviar e-mail
        SQL = "SELECT codprove,nomprove,maiprov1,maiprov2 "
    ElseIf cadTABLA = "clientes" Then
        'seleccionamos todos los clientes a los que queremos enviar e-mail
        SQL = "SELECT codclien,nomclien,maiclie1,maiclie2 "
    End If
    SQL = SQL & "FROM " & cadTABLA
    SQL = SQL & " WHERE " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'creamos una temporal donde guardamos para cada proveedor que SI tiene
    'e-mail, el mail1 o el mail2 al que vamos a enviar
    SQL = "CREATE TEMPORARY TABLE tmpMail ( "
    SQL = SQL & "codusu INT(7) UNSIGNED  DEFAULT '0' NOT NULL, "
    SQL = SQL & "codprove INT(6) UNSIGNED  DEFAULT '0' NOT NULL, "
    SQL = SQL & "nomprove varchar(40)  DEFAULT '' NOT NULL, "
    SQL = SQL & "email varchar(40)  DEFAULT '' NOT NULL) "
    conn.Execute SQL
    
    Cont = 0
    Lista = ""
    
    While Not Rs.EOF
    'para cada cliente/proveedor enviamos un e-mail
        cad1 = DBLet(Rs.Fields(2), "T") 'e-mail administracion
        cad2 = DBLet(Rs.Fields(3), "T") 'e-mail compras
        
        If cad1 = "" And cad2 = "" Then 'no tiene e-mail
'              MsgBox "Sin mail para el proveedor: " & Format(RS!codProve, "000000") & " - " & RS!nomprove, vbExclamation
              Lista = Lista & Format(Rs.Fields(0), "000000") & " - " & Rs.Fields(1) & vbCrLf
        ElseIf cad1 <> "" And cad2 <> "" Then 'tiene 2 e-mail
            'ver a q e-mail se va a enviar (administracion, compras)
            If cadTABLA = "proveedor" Then
                If Me.OptMailCom(0).Value = True Then cad1 = cad2
            Else
                If Me.OptMailCom(1).Value = True Then cad1 = cad2
            End If
        Else 'alguno de los 2 tiene valor
            If cad2 <> "" Then cad1 = cad2  'e-mail para compras
        End If
        
        If cad1 <> "" Then 'HAY email --> ENVIAMOS e-mail
            With frmImprimir
                .OtrosParametros = cadParam
                .NumeroParametros = numParam
                If cadTABLA = "proveedor" Then
                    SQL = "{proveedor.codprove}=" & Rs.Fields(0)
                    .Opcion = 306
                Else
                    SQL = "{clientes.codclien}=" & Rs.Fields(0)
                    .Opcion = 91
                End If
                .FormulaSeleccion = SQL
                .EnvioEMail = True
                CadenaDesdeOtroForm = "GENERANDO"
                .Titulo = cadTit
                .NombreRPT = cadRpt
                .ConSubInforme = True
                .Show vbModal

                If CadenaDesdeOtroForm = "" Then
                'si se ha generado el .pdf para enviar
                    SQL = "INSERT INTO tmpMail (codusu,codprove,nomprove,email)"
                    SQL = SQL & " VALUES (" & vUsu.Codigo & "," & DBSet(Rs.Fields(0), "N") & "," & DBSet(Rs.Fields(1), "T") & "," & DBSet(cad1, "T") & ")"
                    conn.Execute SQL
            
                    Me.Refresh
                    espera 0.4
                    Cont = Cont + 1
                    'Se ha generado bien el documento
                    'Lo copiamos sobre app.path & \temp
                    SQL = Rs.Fields(0) & ".pdf"
                    FileCopy App.path & "\docum.pdf", App.path & "\temp\" & SQL
                End If
            End With
        End If
        Rs.MoveNext
    Wend
    
    Rs.Close
    Set Rs = Nothing
      
    If Cont > 0 Then
        espera 0.4
        If cadTABLA = "proveedor" Then
            SQL = "Carta: " & txtNombre(63).Text & "|"
             SQL = SQL & "Att : " & txtCodigo(62).Text & "|"
        Else
            SQL = "Carta: " & txtNombre(64).Text & "|"
            SQL = SQL & "Att : " & txtCodigo(0).Text & "|"
        End If
       
        frmEMail.Opcion = 2
        frmEMail.DatosEnvio = SQL
        frmEMail.Show vbModal

        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpMail;"
        conn.Execute SQL
        
        'Borrar la carpeta con temporales
        Kill App.path & "\temp\*.pdf"
    End If
    
    Screen.MousePointer = vbDefault
   
    'Mostra mensaje con aquellos proveedores que no tienen e-mail
    If Lista <> "" Then
        If cadTABLA = "sprove" Then
            Lista = "Proveedores sin e-mail:" & vbCrLf & vbCrLf & Lista
        Else
            Lista = "Clientes sin e-mail:" & vbCrLf & vbCrLf & Lista
        End If
        MsgBox Lista, vbInformation
    End If
    
EEnviar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Enviando Informe por e-mail", Err.Description
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpMail;"
        conn.Execute SQL
    End If
End Sub

Private Sub PonerFrameCambioProveedorVisible(H As Integer, W As Integer)
'Frame de Pedidos de Venta y Compra
    W = 7035
    H = 3225
    PonerFrameVisible Me.FrameCambioProveedor, True, H, W
    
End Sub


Private Sub PonerFramePedVisible(H As Integer, W As Integer)
'Frame de Pedidos de Venta y Compra
    W = 6075
    H = 4455
    PonerFrameVisible Me.FramePedidos, True, H, W
    Select Case Opcionlistado
        Case 55 'Cabecera de Pedidos de Compras (a proveedores)
            Me.Label12(0).Caption = "Informe Pedidos compras"
            NomTabla = "scappr"
            NomTablaLin = "slippr"
        Case 56 'Historico de Pedidos Compras
            Me.Label12(0).Caption = "Informe Hist. Pedidos compras"
            NomTabla = "schppr" 'Cabecera  Hco de Pedidos de Compras (a proveedores)
            NomTablaLin = "slhppr"
            If FecEntre <> "" Then txtCodigo(76).Text = FecEntre
    End Select
    
    
    'Ver Fecha Pedido (En Hist.)
    Label12(2).visible = (Opcionlistado = 56)
    txtCodigo(76).visible = (Opcionlistado = 56)
End Sub



Private Sub CargarIconos()
Dim i As Integer
    
    For i = 4 To 4
        Me.imgBuscarOfer(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 28 To 29
        Me.imgBuscarOfer(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 35 To 39
        Me.imgBuscarOfer(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 41 To 44
        Me.imgBuscarOfer(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 48 To 51
        Me.imgBuscarOfer(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 56 To 57
        Me.imgBuscarOfer(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i

End Sub

Private Sub PonerTamnyosMail(peque As Boolean)
    If peque Then
        Me.Height = Me.FrameEnvioMail.Height + 60
        Me.Width = Me.FrameEnvioMail.Width
    Else
        Me.Height = Me.FrameEnvioFacMail.Height
        Me.Width = Me.FrameEnvioFacMail.Width
    End If
    Me.Height = Me.Height + 420
    Me.Width = Me.Width + 120
    Me.FrameEnvioMail.visible = peque
    Me.FrameEnvioFacMail.visible = Not peque
    DoEvents
    Me.Refresh
End Sub


Private Function GeneracionEnvioMail(ByRef Rs As ADODB.Recordset) As Boolean
Dim letraser As String

    On Error GoTo EGeneracionEnvioMail
    
    GeneracionEnvioMail = False

    
    cadselect = "Select * from tmpinformes where codusu =" & vUsu.Codigo & " ORDER BY importe1,codigo1"
    Rs.Open cadselect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CodClien = ""
    While Not Rs.EOF
        
        If Dir(App.path & "\docum.pdf", vbArchive) <> "" Then Kill App.path & "\docum.pdf"
    
        Label14(22).Caption = "Factura: " & Rs!importe1 & " " & Rs!Nombre1
        Label14(22).Refresh
        
        Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
        Dim nomDocu As String 'Nombre de Informe rpt de crystal
        
        indRPT = 12 'Facturas Clientes
        
       If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Function
       'Nombre fichero .rpt a Imprimir
        
        
       cadFormula = "({facturas.codtipom}='" & Trim(Rs!Nombre1) & "') "
       cadFormula = cadFormula & " AND ({facturas.numfactu}=" & Rs!importe1 & ") "
       cadFormula = cadFormula & " AND ({facturas.fecfactu}= Date(" & Year(Rs!fecha1) & "," & Month(Rs!fecha1) & "," & Day(Rs!fecha1) & "))"

   
        With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = True
            .NombreRPT = nomDocu
            .ConSubInforme = True
            .Opcion = 0
            .Titulo = ""
            .Show vbModal
        End With
    
                    
        'Subo el progress bar
        Label14(22).Caption = "Generando PDF"
        Label14(22).Refresh
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
        If (Me.ProgressBar1.Value Mod 25) = 24 Then
            Me.Refresh
            DoEvents
            espera 1
        End If
        Me.Refresh
        DoEvents
        
        If Opcionlistado = 315 Then
            FileCopy App.path & "\docum.pdf", App.path & "\temp\" & Rs!Nombre1 & Format(Rs!importe1, "0000000") & ".pdf" 'RS!importe1 & Format(RS!Codigo1, "0000000") & ".pdf"
        Else
            'Se tiene que llamar base_numserie_numFactura_yyyymmdd.pdf
            ' Sacamos la letra de serie
            letraser = ""
            letraser = DevuelveValor("select letraser from usuarios.stipom where codtipom = " & DBSet(Rs!Nombre1, "T"))
            
            '[Monica]07/03/2013: tengo que incluir el nro de la base de datos, antes era "ariagro_"
            cadFormula = vEmpresa.BDAriagro & "_" & Trim(letraser) & "_" & Rs!importe1 & "_" & Format(Rs!fecha1, "yyyymmdd") & "_F" & Rs!Nombre1 & ".pdf"
            cadFormula = vParamAplic.PathFacturaE & "\" & cadFormula
            
            Label14(22).Caption = cadFormula
            Label14(22).Refresh
        
            FileCopy App.path & "\docum.pdf", cadFormula
            
            'Ha copiado, luego yo la pongo como en facturaE
            cadFormula = "UPDATE facturas set enfacturae=1 WHERE codtipom = '" & Rs!Nombre1 & "' AND numfactu=" & Rs!importe1
            cadFormula = cadFormula & " AND fecfactu='" & Format(Rs!fecha1, FormatoFecha) & "'"
            
            conn.Execute cadFormula
        
        End If
        
        Rs.MoveNext
    Wend
    Rs.Close
    
    Set Rs = Nothing
    GeneracionEnvioMail = True
    Exit Function
EGeneracionEnvioMail:
       MuestraError Err.Number
End Function

Private Sub CargarComboTipoMov(indice As Integer)
'### Combo Tipo Movimiento
'Cargaremos el combo, o bien desde una tabla o con valores fijos
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo

'Lo cargamos con los valores de la tabla stipom que tengan tipo de documento=Albaranes (tipodocu=1)
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim i As Byte

    On Error GoTo ECargaCombo

'    SQL = "select codtipom, nomtipom from stipom where tipodocu=2 " 'Documentos de Facturas
    '3 abril 2007.
    'Mostraba todas las facturas (movimientos que empizan por F, excepto las rectificativas
    'AHora tiene que mostrarlas todas
    'SQL = "select codtipom, nomtipom from stipom where (codtipom like 'F__') and (codtipom<>'FRT')"
    SQL = "select codtipom, nomtipom from usuarios.stipom where (codtipom like 'FA_')  and (tipodocu = 0)"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    i = 0
    
    ListTipoMov(indice).Clear
    
    'LOS TIKCETS NO LOS ENVIO POR MAIL
    While Not Rs.EOF
        ListTipoMov(indice).AddItem Rs.Fields(0).Value & "-" & Rs.Fields(1).Value
        'ListTipoMov(indice).List (ListTipoMov(indice).NewIndex)
        ListTipoMov(indice).Selected((ListTipoMov(indice).NewIndex)) = True
        
        Rs.MoveNext
    Wend
    
    Rs.Close
    Set Rs = Nothing
    
ECargaCombo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


