VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListado2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   11175
   Icon            =   "frmListado2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameRepxDia 
      Height          =   5085
      Left            =   45
      TabIndex        =   153
      Top             =   45
      Width           =   5715
      Begin VB.CommandButton cmdAceptarRepxDia 
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
         Left            =   3135
         TabIndex        =   169
         Top             =   3150
         Width           =   1065
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   7
         Left            =   4320
         TabIndex        =   170
         Top             =   3150
         Width           =   1065
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   2025
         Left            =   105
         TabIndex        =   161
         Top             =   990
         Width           =   5460
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
            Index           =   32
            Left            =   3840
            TabIndex        =   163
            Top             =   435
            Width           =   1350
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
            Index           =   31
            Left            =   1380
            TabIndex        =   162
            Top             =   435
            Width           =   1350
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   1215
            Left            =   135
            TabIndex        =   172
            Top             =   765
            Width           =   5130
            Begin VB.TextBox Text1 
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
               Index           =   6
               Left            =   1620
               MaxLength       =   12
               TabIndex        =   164
               Tag             =   "Tipo Movimiento|T|N|||facturas|codtipom||S|"
               Top             =   90
               Width           =   990
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
               Index           =   0
               Left            =   1215
               MaxLength       =   10
               TabIndex        =   168
               Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
               Top             =   765
               Width           =   1395
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
               Left            =   2655
               Locked          =   -1  'True
               TabIndex        =   173
               Top             =   765
               Width           =   2640
            End
            Begin VB.Label Label1 
               Caption         =   "Tipo Factura"
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
               Index           =   0
               Left            =   210
               TabIndex        =   203
               Top             =   90
               Width           =   1365
            End
            Begin VB.Image imgBuscarG 
               Height          =   240
               Index           =   98
               Left            =   945
               ToolTipText     =   "Buscar Concepto"
               Top             =   765
               Width           =   240
            End
            Begin VB.Label Label1 
               Caption         =   "Cuenta Banco: "
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
               Height          =   195
               Index           =   24
               Left            =   225
               TabIndex        =   174
               Top             =   450
               Width           =   1665
            End
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   5
            Left            =   3540
            Picture         =   "frmListado2.frx":000C
            Top             =   435
            Width           =   240
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   4
            Left            =   1095
            Picture         =   "frmListado2.frx":0097
            Top             =   435
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Reparación:"
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
            Height          =   240
            Index           =   2
            Left            =   360
            TabIndex        =   167
            Top             =   105
            Width           =   1845
         End
         Begin VB.Label Label3 
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
            Height          =   195
            Index           =   29
            Left            =   2835
            TabIndex        =   166
            Top             =   435
            Width           =   600
         End
         Begin VB.Label Label2 
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
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   165
            Top             =   435
            Width           =   600
         End
      End
      Begin VB.Frame FrameContab 
         Caption         =   " Facturas "
         Enabled         =   0   'False
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
         Height          =   620
         Left            =   495
         TabIndex        =   158
         Top             =   45
         Visible         =   0   'False
         Width           =   4725
         Begin VB.OptionButton OptClientes 
            Caption         =   "Clientes"
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
            Left            =   705
            TabIndex        =   160
            Top             =   250
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton OptProve 
            Caption         =   "Proveedores"
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
            Left            =   2505
            TabIndex        =   159
            Top             =   250
            Width           =   1695
         End
      End
      Begin VB.Frame FrameProgress 
         Height          =   1200
         Left            =   450
         TabIndex        =   154
         Top             =   3510
         Visible         =   0   'False
         Width           =   4935
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   400
            Left            =   120
            TabIndex        =   155
            Top             =   640
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   714
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label lblProgess 
            Caption         =   "Iniciando el proceso ..."
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
            Index           =   1
            Left            =   120
            TabIndex        =   157
            Top             =   375
            Width           =   4575
         End
         Begin VB.Label lblProgess 
            Caption         =   "Comprobaciones:"
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
            Index           =   0
            Left            =   120
            TabIndex        =   156
            Top             =   135
            Width           =   4455
         End
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Reparaciones por Día"
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
         Left            =   480
         TabIndex        =   171
         Top             =   580
         Width           =   5055
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   10680
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameEnvioMail 
      Height          =   1215
      Left            =   1800
      TabIndex        =   138
      Top             =   0
      Visible         =   0   'False
      Width           =   6615
      Begin MSComctlLib.ProgressBar PBMail 
         Height          =   375
         Left            =   360
         TabIndex        =   139
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
         TabIndex        =   140
         Top             =   840
         Width           =   5805
      End
   End
   Begin VB.Frame FrameMovArtic 
      Height          =   6795
      Left            =   0
      TabIndex        =   35
      Top             =   90
      Width           =   10635
      Begin VB.CheckBox ChKAcumulados 
         Caption         =   "Sacar Acumulados"
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
         Height          =   240
         Left            =   3000
         TabIndex        =   201
         Top             =   5985
         Visible         =   0   'False
         Width           =   2490
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
         Index           =   86
         Left            =   1485
         TabIndex        =   25
         Top             =   5505
         Width           =   855
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
         Index           =   87
         Left            =   1485
         TabIndex        =   26
         Top             =   5910
         Width           =   855
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
         Index           =   5
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   17
         Top             =   1170
         Width           =   1455
      End
      Begin VB.CommandButton cmdDeselTodos 
         Height          =   435
         Left            =   9000
         Picture         =   "frmListado2.frx":0122
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   740
         Width           =   585
      End
      Begin VB.CommandButton cmdSelTodos 
         Height          =   435
         Left            =   9720
         Picture         =   "frmListado2.frx":080C
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   740
         Width           =   585
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3870
         Left            =   6960
         TabIndex        =   27
         Top             =   1200
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   6826
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Index           =   12
         Left            =   2120
         Locked          =   -1  'True
         TabIndex        =   66
         Text            =   "Text5"
         Top             =   4725
         Width           =   4530
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Index           =   11
         Left            =   2120
         Locked          =   -1  'True
         TabIndex        =   65
         Text            =   "Text5"
         Top             =   4320
         Width           =   4530
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Left            =   2120
         Locked          =   -1  'True
         TabIndex        =   64
         Text            =   "Text5"
         Top             =   2760
         Width           =   4530
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Left            =   2120
         Locked          =   -1  'True
         TabIndex        =   63
         Text            =   "Text5"
         Top             =   2355
         Width           =   4530
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   62
         Text            =   "Text5"
         Top             =   1575
         Width           =   3705
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   61
         Text            =   "Text5"
         Top             =   1170
         Width           =   3705
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
         Index           =   12
         Left            =   1440
         TabIndex        =   24
         Top             =   4725
         Width           =   615
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
         Index           =   11
         Left            =   1440
         TabIndex        =   23
         Top             =   4320
         Width           =   615
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
         Index           =   10
         Left            =   3960
         TabIndex        =   22
         Top             =   3540
         Width           =   1350
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
         Index           =   9
         Left            =   1440
         TabIndex        =   21
         Top             =   3540
         Width           =   1350
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
         Index           =   8
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   20
         Top             =   2760
         Width           =   615
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
         Index           =   7
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   19
         Top             =   2355
         Width           =   615
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
         Index           =   6
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   18
         Top             =   1575
         Width           =   1455
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
         Index           =   3
         Left            =   8115
         TabIndex        =   28
         Top             =   6195
         Width           =   1065
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   3
         Left            =   9240
         TabIndex        =   29
         Top             =   6195
         Width           =   1065
      End
      Begin VB.Label Label3 
         Caption         =   "Rellenando Movimientos...."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   390
         TabIndex        =   202
         Top             =   6345
         Visible         =   0   'False
         Width           =   2520
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   65
         Left            =   1200
         Top             =   5505
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   66
         Left            =   1200
         Top             =   5910
         Width           =   240
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   70
         Left            =   510
         TabIndex        =   137
         Top             =   5505
         Width           =   690
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   69
         Left            =   510
         TabIndex        =   136
         Top             =   5910
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cliente/Proveedor/Socio"
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
         Height          =   240
         Index           =   68
         Left            =   360
         TabIndex        =   135
         Top             =   5175
         Width           =   2385
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipos de Movimiento"
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
         Height          =   240
         Index           =   12
         Left            =   6960
         TabIndex        =   60
         Top             =   960
         Width           =   2025
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   3675
         Picture         =   "frmListado2.frx":0EF6
         Top             =   3540
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1155
         Picture         =   "frmListado2.frx":0F81
         Top             =   3540
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   34
         Left            =   1155
         Top             =   4725
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   33
         Left            =   1155
         Top             =   4320
         Width           =   240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
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
         Height          =   240
         Index           =   11
         Left            =   360
         TabIndex        =   59
         Top             =   3990
         Width           =   825
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   8
         Left            =   510
         TabIndex        =   58
         Top             =   4725
         Width           =   645
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   7
         Left            =   510
         TabIndex        =   57
         Top             =   4320
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Height          =   240
         Index           =   10
         Left            =   360
         TabIndex        =   56
         Top             =   3210
         Width           =   600
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   6
         Left            =   3030
         TabIndex        =   55
         Top             =   3540
         Width           =   600
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   5
         Left            =   510
         TabIndex        =   43
         Top             =   3540
         Width           =   690
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   32
         Left            =   1155
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   31
         Left            =   1155
         Top             =   2355
         Width           =   240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Familia"
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
         Height          =   240
         Index           =   9
         Left            =   360
         TabIndex        =   42
         Top             =   2025
         Width           =   660
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   4
         Left            =   510
         TabIndex        =   41
         Top             =   2760
         Width           =   645
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   0
         Left            =   510
         TabIndex        =   40
         Top             =   2355
         Width           =   690
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   30
         Left            =   1155
         Top             =   1575
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   29
         Left            =   1155
         Top             =   1170
         Width           =   240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Artículo"
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
         Height          =   240
         Index           =   1
         Left            =   360
         TabIndex        =   39
         Top             =   840
         Width           =   750
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Informes Movimiento Artículos"
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
         Left            =   240
         TabIndex        =   38
         Top             =   360
         Width           =   6705
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   33
         Left            =   510
         TabIndex        =   37
         Top             =   1575
         Width           =   645
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   2
         Left            =   510
         TabIndex        =   36
         Top             =   1170
         Width           =   690
      End
   End
   Begin VB.Frame FrameRecalPMP 
      Height          =   5985
      Left            =   0
      TabIndex        =   175
      Top             =   -60
      Width           =   7275
      Begin VB.CommandButton cmdCancel 
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
         Left            =   5805
         TabIndex        =   183
         Top             =   5280
         Width           =   1065
      End
      Begin VB.CommandButton cmdRecalPMP 
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
         Left            =   4590
         TabIndex        =   182
         Top             =   5280
         Width           =   1065
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
         Index           =   25
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   177
         Top             =   1710
         Width           =   1500
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
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   178
         Top             =   2490
         Width           =   1140
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
         Index           =   27
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   179
         Top             =   2895
         Width           =   1140
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
         Index           =   28
         Left            =   1440
         TabIndex        =   180
         Top             =   3735
         Width           =   1140
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
         Index           =   29
         Left            =   1440
         TabIndex        =   181
         Top             =   4140
         Width           =   1140
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   189
         Text            =   "Text5"
         Top             =   1305
         Width           =   3975
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   188
         Text            =   "Text5"
         Top             =   1710
         Width           =   3975
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Left            =   2595
         Locked          =   -1  'True
         TabIndex        =   187
         Text            =   "Text5"
         Top             =   2490
         Width           =   4305
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Index           =   27
         Left            =   2595
         Locked          =   -1  'True
         TabIndex        =   186
         Text            =   "Text5"
         Top             =   2895
         Width           =   4305
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Index           =   28
         Left            =   2595
         Locked          =   -1  'True
         TabIndex        =   185
         Text            =   "Text5"
         Top             =   3735
         Width           =   4305
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Index           =   29
         Left            =   2595
         Locked          =   -1  'True
         TabIndex        =   184
         Text            =   "Text5"
         Top             =   4140
         Width           =   4305
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
         Index           =   24
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   176
         Top             =   1305
         Width           =   1500
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   41
         Left            =   360
         TabIndex        =   200
         Top             =   4980
         Visible         =   0   'False
         Width           =   6120
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   36
         Left            =   465
         TabIndex        =   199
         Top             =   1305
         Width           =   645
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   35
         Left            =   465
         TabIndex        =   198
         Top             =   1710
         Width           =   600
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Recálculo Precio Medio Ponderado"
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
         Left            =   330
         TabIndex        =   197
         Top             =   360
         Width           =   6705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Artículo"
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
         Height          =   240
         Index           =   34
         Left            =   360
         TabIndex        =   196
         Top             =   930
         Width           =   750
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   35
         Left            =   1155
         Top             =   1305
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   22
         Left            =   1155
         Top             =   1710
         Width           =   240
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   32
         Left            =   465
         TabIndex        =   195
         Top             =   2490
         Width           =   645
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   31
         Left            =   465
         TabIndex        =   194
         Top             =   2895
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Familia"
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
         Height          =   240
         Index           =   30
         Left            =   360
         TabIndex        =   193
         Top             =   2160
         Width           =   660
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   21
         Left            =   1155
         Top             =   2490
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   17
         Left            =   1155
         Top             =   2895
         Width           =   240
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   21
         Left            =   465
         TabIndex        =   192
         Top             =   3735
         Width           =   645
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   18
         Left            =   465
         TabIndex        =   191
         Top             =   4140
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
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
         Height          =   240
         Index           =   17
         Left            =   360
         TabIndex        =   190
         Top             =   3405
         Width           =   990
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   9
         Left            =   1140
         Top             =   4140
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   8
         Left            =   1140
         Top             =   3735
         Width           =   240
      End
   End
   Begin VB.Frame FrameHcoMante 
      Height          =   3495
      Left            =   0
      TabIndex        =   141
      Top             =   -120
      Width           =   6495
      Begin VB.CommandButton cmdHcoMante 
         Caption         =   "Aceptar"
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
         Left            =   3930
         TabIndex        =   146
         Top             =   2880
         Width           =   1095
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
         Index           =   112
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   145
         Top             =   2280
         Width           =   780
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Index           =   112
         Left            =   2505
         Locked          =   -1  'True
         TabIndex        =   151
         Text            =   "Text5"
         Top             =   2280
         Width           =   3780
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
         Index           =   111
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   144
         Top             =   1560
         Width           =   780
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Index           =   111
         Left            =   2505
         Locked          =   -1  'True
         TabIndex        =   149
         Text            =   "Text5"
         Top             =   1560
         Width           =   3780
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
         Index           =   110
         Left            =   1725
         TabIndex        =   143
         Top             =   960
         Width           =   1350
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   99
         Left            =   5250
         TabIndex        =   148
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Motivo baja"
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
         Height          =   240
         Index           =   81
         Left            =   240
         TabIndex        =   152
         Top             =   2280
         Width           =   1155
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   90
         Left            =   1440
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
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
         Height          =   240
         Index           =   80
         Left            =   240
         TabIndex        =   150
         Top             =   1560
         Width           =   1065
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   89
         Left            =   1440
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   12
         Left            =   1440
         Picture         =   "frmListado2.frx":100C
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha baja"
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
         Height          =   240
         Index           =   79
         Left            =   240
         TabIndex        =   147
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Paso a mantenimientos anulados"
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
         Index           =   2
         Left            =   240
         TabIndex        =   142
         Top             =   240
         Width           =   5775
      End
   End
   Begin VB.Frame frameListado 
      Height          =   4695
      Left            =   360
      TabIndex        =   10
      Top             =   120
      Width           =   6555
      Begin VB.Frame frameOrdenar 
         Caption         =   "Ordenar por"
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
         Height          =   735
         Left            =   585
         TabIndex        =   92
         Top             =   2640
         Width           =   3555
         Begin VB.OptionButton OptNombre 
            Caption         =   "Descripción"
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
            Left            =   1680
            TabIndex        =   3
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton Optcodigo 
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
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Left            =   2480
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   2040
         Width           =   3300
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Left            =   2480
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   1560
         Width           =   3300
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
         Left            =   1605
         TabIndex        =   1
         Top             =   2040
         Width           =   870
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
         Index           =   1
         Left            =   1605
         TabIndex        =   0
         Top             =   1560
         Width           =   870
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
         Index           =   1
         Left            =   3600
         TabIndex        =   4
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   1
         Left            =   4680
         TabIndex        =   5
         Top             =   3960
         Width           =   975
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1320
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1320
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00972E0B&
         Height          =   240
         Index           =   1
         Left            =   585
         TabIndex        =   16
         Top             =   1200
         Width           =   840
      End
      Begin VB.Label Label1 
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
         Height          =   195
         Index           =   3
         Left            =   585
         TabIndex        =   14
         Top             =   2040
         Width           =   600
      End
      Begin VB.Label Label1 
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
         Height          =   195
         Index           =   2
         Left            =   585
         TabIndex        =   13
         Top             =   1605
         Width           =   645
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Listado Marcas"
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
         TabIndex        =   15
         Top             =   480
         Width           =   5055
      End
   End
   Begin VB.Frame FrameDtosFM 
      Height          =   6225
      Left            =   45
      TabIndex        =   105
      Top             =   45
      Width           =   6915
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   1230
         Left            =   270
         TabIndex        =   124
         Top             =   750
         Width           =   6180
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
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
            Index           =   74
            Left            =   2430
            Locked          =   -1  'True
            TabIndex        =   126
            Text            =   "Text5"
            Top             =   810
            Width           =   3810
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
            Index           =   74
            Left            =   1560
            MaxLength       =   6
            TabIndex        =   98
            Top             =   810
            Width           =   875
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
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
            Index           =   73
            Left            =   2430
            Locked          =   -1  'True
            TabIndex        =   125
            Text            =   "Text5"
            Top             =   405
            Width           =   3810
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
            Index           =   73
            Left            =   1560
            MaxLength       =   6
            TabIndex        =   97
            Top             =   405
            Width           =   875
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   1
            Left            =   1275
            Top             =   810
            Width           =   240
         End
         Begin VB.Label Label3 
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
            Height          =   195
            Index           =   61
            Left            =   585
            TabIndex        =   129
            Top             =   405
            Width           =   600
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   0
            Left            =   1275
            Top             =   405
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00972E0B&
            Height          =   240
            Index           =   44
            Left            =   240
            TabIndex        =   128
            Top             =   120
            Width           =   675
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
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
            Height          =   240
            Index           =   45
            Left            =   585
            TabIndex        =   127
            Top             =   810
            Width           =   705
         End
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   1140
         Left            =   270
         TabIndex        =   118
         Top             =   3060
         Width           =   6135
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
            Index           =   77
            Left            =   1560
            MaxLength       =   4
            TabIndex        =   101
            Top             =   360
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
            Index           =   78
            Left            =   1560
            MaxLength       =   4
            TabIndex        =   102
            Top             =   765
            Width           =   875
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
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
            Index           =   77
            Left            =   2430
            Locked          =   -1  'True
            TabIndex        =   120
            Text            =   "Text5"
            Top             =   360
            Width           =   3765
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
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
            Index           =   78
            Left            =   2430
            Locked          =   -1  'True
            TabIndex        =   119
            Text            =   "Text5"
            Top             =   765
            Width           =   3765
         End
         Begin VB.Label Label3 
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
            Height          =   195
            Index           =   66
            Left            =   585
            TabIndex        =   123
            Top             =   360
            Width           =   600
         End
         Begin VB.Label Label3 
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
            Height          =   195
            Index           =   67
            Left            =   585
            TabIndex        =   122
            Top             =   765
            Width           =   555
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Marca"
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
            Height          =   240
            Index           =   42
            Left            =   240
            TabIndex        =   121
            Top             =   75
            Width           =   600
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   4
            Left            =   1275
            Top             =   360
            Width           =   240
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   5
            Left            =   1275
            Top             =   765
            Width           =   240
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   270
         TabIndex        =   112
         Top             =   4215
         Width           =   6255
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
            Index           =   79
            Left            =   1560
            TabIndex        =   95
            Top             =   360
            Width           =   875
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
            Index           =   80
            Left            =   1560
            TabIndex        =   96
            Top             =   720
            Width           =   875
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
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
            Index           =   79
            Left            =   2445
            Locked          =   -1  'True
            TabIndex        =   114
            Text            =   "Text5"
            Top             =   360
            Width           =   3750
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
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
            Index           =   80
            Left            =   2445
            Locked          =   -1  'True
            TabIndex        =   113
            Text            =   "Text5"
            Top             =   720
            Width           =   3750
         End
         Begin VB.Label Label3 
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
            Height          =   195
            Index           =   65
            Left            =   585
            TabIndex        =   117
            Top             =   360
            Width           =   645
         End
         Begin VB.Label Label3 
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
            Height          =   195
            Index           =   64
            Left            =   585
            TabIndex        =   116
            Top             =   720
            Width           =   600
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
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
            Height          =   240
            Index           =   46
            Left            =   240
            TabIndex        =   115
            Top             =   75
            Width           =   990
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   63
            Left            =   1275
            Top             =   360
            Width           =   240
         End
         Begin VB.Image imgBuscarG 
            Height          =   240
            Index           =   64
            Left            =   1275
            Top             =   720
            Width           =   240
         End
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   12
         Left            =   5400
         TabIndex        =   104
         Top             =   5550
         Width           =   1065
      End
      Begin VB.CommandButton cmdAceptarDtosFM 
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
         Left            =   4230
         TabIndex        =   103
         Top             =   5550
         Width           =   1065
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
         Index           =   75
         Left            =   1830
         MaxLength       =   4
         TabIndex        =   99
         Top             =   2280
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
         Index           =   76
         Left            =   1830
         MaxLength       =   4
         TabIndex        =   100
         Top             =   2685
         Width           =   875
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Index           =   75
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   107
         Text            =   "Text5"
         Top             =   2295
         Width           =   3750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Index           =   76
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   106
         Text            =   "Text5"
         Top             =   2685
         Width           =   3750
      End
      Begin VB.Label Label10 
         Caption         =   "Listado Descuentos Familia/Marca"
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
         Left            =   525
         TabIndex        =   111
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   63
         Left            =   855
         TabIndex        =   110
         Top             =   2280
         Width           =   690
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   62
         Left            =   855
         TabIndex        =   109
         Top             =   2685
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Familia"
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
         Height          =   240
         Index           =   40
         Left            =   510
         TabIndex        =   108
         Top             =   1950
         Width           =   660
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   2
         Left            =   1545
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   3
         Left            =   1545
         Top             =   2685
         Width           =   240
      End
   End
   Begin VB.Frame FrameInfAlmacen 
      Height          =   3495
      Left            =   1560
      TabIndex        =   30
      Top             =   1080
      Width           =   5835
      Begin VB.CommandButton cmdCancel 
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
         Index           =   2
         Left            =   4125
         TabIndex        =   9
         Top             =   2835
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
         Index           =   2
         Left            =   2955
         TabIndex        =   8
         Top             =   2835
         Width           =   1065
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
         Index           =   3
         Left            =   1380
         TabIndex        =   6
         Top             =   1800
         Width           =   1350
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
         Index           =   4
         Left            =   3885
         TabIndex        =   7
         Top             =   1800
         Width           =   1350
      End
      Begin VB.Label Label2 
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
         Height          =   195
         Index           =   6
         Left            =   360
         TabIndex        =   34
         Top             =   1800
         Width           =   645
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   3
         Left            =   2820
         TabIndex        =   33
         Top             =   1800
         Width           =   645
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Informes Almacenes"
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
         Index           =   2
         Left            =   360
         TabIndex        =   32
         Top             =   600
         Width           =   5055
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nº Traspaso"
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
         Height          =   240
         Index           =   1
         Left            =   360
         TabIndex        =   31
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1095
         Picture         =   "frmListado2.frx":1097
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   3600
         Picture         =   "frmListado2.frx":1199
         Top             =   1800
         Width           =   240
      End
   End
   Begin VB.Frame FrameInventario 
      Height          =   6675
      Left            =   0
      TabIndex        =   67
      Top             =   0
      Width           =   7995
      Begin VB.Frame FrameValorar 
         Caption         =   "Valorar Con:"
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
         Height          =   1455
         Left            =   165
         TabIndex        =   87
         Top             =   4320
         Visible         =   0   'False
         Width           =   3075
         Begin VB.OptionButton optPrecioUC 
            Caption         =   "Precio Ultima Compra"
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
            Left            =   240
            TabIndex        =   89
            Top             =   880
            Width           =   2775
         End
         Begin VB.OptionButton optPrecioMP 
            Caption         =   "Precio Medio Ponderado"
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
            Left            =   240
            TabIndex        =   88
            Top             =   450
            Value           =   -1  'True
            Width           =   2775
         End
      End
      Begin VB.Frame FrameOpciones 
         Height          =   1695
         Left            =   4230
         TabIndex        =   130
         Top             =   4365
         Width           =   3510
         Begin VB.CheckBox chkValorado 
            Caption         =   "Valorado"
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
            Left            =   240
            TabIndex        =   134
            Top             =   1320
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkImprimeStock 
            Caption         =   "Imprimir Stock"
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
            Left            =   240
            TabIndex        =   133
            Top             =   960
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.CheckBox chkSinStock 
            Caption         =   "Imprimir Artículos sin Stock"
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
            Left            =   240
            TabIndex        =   132
            Top             =   600
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   3120
         End
         Begin VB.CheckBox chkSaltaPag 
            Caption         =   "Salta pág. en Familia"
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
            Left            =   240
            TabIndex        =   131
            Top             =   240
            Width           =   2775
         End
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
         Index           =   22
         Left            =   5370
         TabIndex        =   52
         Top             =   4440
         Width           =   1350
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Index           =   19
         Left            =   2355
         Locked          =   -1  'True
         TabIndex        =   83
         Text            =   "Text5"
         Top             =   3960
         Width           =   5385
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Index           =   18
         Left            =   2355
         Locked          =   -1  'True
         TabIndex        =   82
         Text            =   "Text5"
         Top             =   3600
         Width           =   5385
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
         Index           =   19
         Left            =   1560
         TabIndex        =   50
         Top             =   3960
         Width           =   735
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
         Index           =   18
         Left            =   1560
         TabIndex        =   49
         Top             =   3600
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   4
         Left            =   6660
         TabIndex        =   54
         Top             =   6135
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
         Index           =   4
         Left            =   5535
         TabIndex        =   53
         Top             =   6135
         Width           =   1065
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
         Index           =   14
         Left            =   1560
         MaxLength       =   16
         TabIndex        =   45
         Top             =   1680
         Width           =   1455
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
         Index           =   15
         Left            =   1560
         MaxLength       =   16
         TabIndex        =   46
         Top             =   2040
         Width           =   1455
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
         Index           =   16
         Left            =   1560
         MaxLength       =   4
         TabIndex        =   47
         Top             =   2640
         Width           =   615
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
         Index           =   17
         Left            =   1560
         MaxLength       =   4
         TabIndex        =   48
         Top             =   3000
         Width           =   615
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
         Index           =   20
         Left            =   2265
         TabIndex        =   51
         Top             =   4440
         Width           =   1350
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
         Index           =   13
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   44
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Left            =   3075
         Locked          =   -1  'True
         TabIndex        =   72
         Text            =   "Text5"
         Top             =   1680
         Width           =   4710
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Left            =   3075
         Locked          =   -1  'True
         TabIndex        =   71
         Text            =   "Text5"
         Top             =   2040
         Width           =   4710
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Index           =   16
         Left            =   2235
         Locked          =   -1  'True
         TabIndex        =   70
         Text            =   "Text5"
         Top             =   2640
         Width           =   5520
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Index           =   17
         Left            =   2235
         Locked          =   -1  'True
         TabIndex        =   69
         Text            =   "Text5"
         Top             =   3000
         Width           =   5520
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Index           =   13
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   68
         Text            =   "Text5"
         Top             =   1080
         Width           =   5640
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   5115
         Picture         =   "frmListado2.frx":129B
         Top             =   4440
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   9
         Left            =   4650
         TabIndex        =   91
         Top             =   4440
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   8
         Left            =   3720
         TabIndex        =   90
         Top             =   4440
         Width           =   600
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   16
         Left            =   1275
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   15
         Left            =   1275
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
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
         Height          =   240
         Index           =   4
         Left            =   240
         TabIndex        =   86
         Top             =   3360
         Width           =   990
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   25
         Left            =   540
         TabIndex        =   85
         Top             =   3960
         Width           =   690
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   24
         Left            =   540
         TabIndex        =   84
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   23
         Left            =   540
         TabIndex        =   81
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   22
         Left            =   540
         TabIndex        =   80
         Top             =   2040
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Artículo"
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
         Height          =   240
         Index           =   2
         Left            =   240
         TabIndex        =   78
         Top             =   1440
         Width           =   750
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   11
         Left            =   1275
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   12
         Left            =   1275
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   20
         Left            =   540
         TabIndex        =   77
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Index           =   19
         Left            =   540
         TabIndex        =   76
         Top             =   3000
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Familia"
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
         Height          =   240
         Index           =   3
         Left            =   240
         TabIndex        =   75
         Top             =   2400
         Width           =   660
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   13
         Left            =   1275
         Top             =   2640
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   14
         Left            =   1275
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inventario"
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
         Height          =   240
         Index           =   5
         Left            =   240
         TabIndex        =   74
         Top             =   4440
         Width           =   1680
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
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
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   73
         Top             =   1080
         Width           =   825
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   10
         Left            =   1275
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   1965
         Picture         =   "frmListado2.frx":1326
         Top             =   4440
         Width           =   240
      End
      Begin VB.Label lbltituloInven 
         Caption         =   "Informe Toma de Inventario Artículos"
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
         Left            =   240
         TabIndex        =   79
         Top             =   360
         Width           =   7575
      End
   End
End
Attribute VB_Name = "frmListado2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcionlistado As Integer

    '==== Listados de ALMACEN ====
    '=============================
    ' 1 .- Listados Marcas.
    ' 2 .- Listado de Almacenes Propios
    ' 3 .- Listado de Tipos de Unidad
    ' 4 .- Listado de Tipos de Artículos
    ' 5 .- Listado de Familias de artículos
    
    ' 6 .- Listado de Artículos
    ' 7 .- Informe de Traspaso de Almacenes
    ' 8 .- Informe de Movimientos de Almacen
    ' 9 .- Listado Busquedas de movimientos de Artículos
    '10 .- Informe de Movimientos de Servicios Varios (SOCIOS, CLIENTES)
    
    '11 .-
    '12 .- Listado Toma de Inventario Articulos
    '13 .- Listado de Diferencias de Inventario Articulos
    '14 .- Actualizar Diferencias de Inventario (No IMPRIME INFORME)
    '15 .- Listado de Articulos Inactivos.
    
    '16 .- Listado Valoracion de Stocks Inventariados
    '17 .- Listado Valoración Stocks
    '18 .- Informe Stocks Maximos y Minimos
    '19 .- Informe de Stocks a una fecharEtiqBulto.rpt
    
    '110 .- Listado de Ubicaciones
    
    '120 .- Recálculo del Precio Medio Ponderado
    
    
    
    '==== Listados de FACTURACION ====
    '=================================
    '20 .- Listado de Actividades de Clientes
    '21 .- Listado de Zonas de Clientes
    '22 .- Listado de Rutas de Asistencia
    '23 .- Listado de Formas de Envío
    '24 .- Listado de Tarifas Ventas
    '25 .-
    
    '26 .-
    '27 .- Listado de Situaciones Especiales
    '28 .- Informe de Tarifas de Articulos
    '29 .- Informe de Promociones de Tarifas
    '30 .- Informe de Precios Especiales
    
    '31 .- Informe de Ofertas
    '32 .- Informe de Recordatorio de Ofertas
    '33 .- Informe de Valoración de Ofertas
    '34 .- Informe de Ofertas Efectuadas
    '35 .- Informe Historico de Ofertas
    
    '36 .- Traspaso de Ofertas al Historico (NO IMPRIME INFORME)
    '37 .- Solicitar datos para pasar de Oferta a Pedido (NO IMPRIME INFORME)
    '38 .- Informe de Pedidos
    '239 .- Hco de Pedidos de venta (Historico)
    '39 .- Orden de Instalacion
    '40 .- Cartas Confirmacion de Pedidos
    
    '41 .- Informe de Pedidos por Articulo
    '42 .- Informe de Disponibilidad de Stocks
    '43 .- Generar Albaran desde Pedido (NO IMPRIME LISTADO)
    '44 .- Informe de Pedidos por Cliente
    '45 .- Informe de Albaran
    
    '46 .- Informe de Clientes Inactivos
    '47 .- Informe de Clientes
    '48 .- Informe de Altas de Nuevos Cliente
    '49 .- Informe de Albaranes por Articulo
    '50 .- Prevision de Facturacion de ALbaranes
    
    '51 .- Informe Incumplimiento Plazos de Entrega
    '52 .- Facturacion de Albaranes (NO IMPRIME LISTADO?)
    '53 .- Informe de Factura
    '54 .- Listado de Descuentos Familia/Marca
    
    '59 .- Informe de Factura ProForma
    '222 .- Informe de Factura Mostrador
    '223 .- Pedir datos para contabilizar facturas CLIENTES
    '224 .- Pedir datos para contabilizar facturas PROVEEDOR
    '225 .- Pedir datos para generar Facturas Rectificativas
    '226 .- Pedir datos para reimprimir Facturas
    '227 .- Informe estadistica Ventas por cliente
    '228 .- Informe estadistica Ventas por Trabajador
    '229 .- Informe estadistica Ventas por meses
    '230 .- Informe estadistica Ventas por familia
    '231 .- Informe detalle facturacion clientes
    
    '240 .- Informe Cierre de Caja del TPV
    
    '245 .- Informe control margenes tarifas
    '246 .- Informe Margen ventas por articulo
    '247 .- Corrección de errores y acutalizacion de tarifas
    
    
    '==== Listados de COMPRAS ====
    '=============================
    '55 .- Informe de Pedido Proveedor
    '56 .- Inf. Historico Pedido Proveedor
    '57 .- Pasa Pedido a Albaran compras (NO IMPRIME LISTADO)
    '58 .- Listado de Proveedores
    
    
    '305 .- Listado Etiquetas de Proveedores
    '306 .- Listado Cartas a Proveedores
    '307 .- Listado Material pendiente de recibir
    '308 .- Listado Albaranes pendientes de facturar
    '309 .- Listado  Precios de Compra
    '310 .- Listado Compras por Proveedor
    '311 .- Listado Compras por Familia
    
    
    
    '==== Listados OTROS ====
    '==================================
    
    '80 .- Pasar Albaranes Ventas al historico (NO IMPRIME)
    '81 .- Pasar Pedidos Ventas al historico (NO IMPRIME)
       
           
    '82 .- Marcar facturar albaranes
    '83 .- Borre avisos cerrados
       
    '90 .- Etiquetas de Clientes
    '91 .- Cartas a Clientes
    
    '92 .- Informe de Gastos Técnicos
    '93 .- Ticket del TPV
        
    '94 .- Etiquetas estanteria
    
    
    '95 .- Etiquetas de bultos
    '96 .- Frecuencias
    '97 .- Eliminar facturas
    '99 .- Traspaso a mantenimientos anulados
    
    
    
    
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

                        ' CADTAG = "A" : me indica que es una factura de venta a socio
    
Public TipoMto As Byte ' solo se utiliza en el informe de movimientos de servicios a socios y clientes
                       ' 0 = es un movimiento de servicios a SOCIO
                       ' 1 = es un movimiento de servicios a CLIENTE
                       
Public DeServicios As Boolean  ' indica si el listado es de servicios a socio/cliente

Private Conexion As Byte
'1.- Conexión a BD Ariagro  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmMtoAlPropios As frmManAlmProp
Attribute frmMtoAlPropios.VB_VarHelpID = -1
Private WithEvents frmMtoTUnidad As frmManTipUnid
Attribute frmMtoTUnidad.VB_VarHelpID = -1
Private WithEvents frmMtoTArticulo As frmManTipArtic
Attribute frmMtoTArticulo.VB_VarHelpID = -1
Private WithEvents frmCta As frmCtasConta
Attribute frmCta.VB_VarHelpID = -1
Private WithEvents frmMtoFamilia As frmManFamilias
Attribute frmMtoFamilia.VB_VarHelpID = -1
Private WithEvents frmMtoProveedor As frmManProve
Attribute frmMtoProveedor.VB_VarHelpID = -1
Private WithEvents frmMtoArticulos As frmManArtic
Attribute frmMtoArticulos.VB_VarHelpID = -1
Private WithEvents frmMtoClientes As frmClientes
Attribute frmMtoClientes.VB_VarHelpID = -1
Private WithEvents frmMtoIncid As frmManInciden
Attribute frmMtoIncid.VB_VarHelpID = -1

Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1

'---- Variables para el INFORME ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadselect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para el frmImprimir
Private cadNomRPT As String 'Nombre del informe a Imprimir
Private conSubRPT As Boolean 'Si el informe tiene subreports
'-----------------------------------


Dim indCodigo As Integer 'indice para txtCodigo

Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean
Dim indFrame As Single
Dim SoloServicios As Boolean

Dim cContaFra As cContabilizarFacturas




Private Sub KEYpress(KeyAscii As Integer)
    Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cmdAceptar_Click(Index As Integer)
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim cadAux As String
Dim bol As Boolean
   InicializarVbles
   
   Select Case Index
   '========= Frame Listados =================================================
    Case 1 'Frame Listados
        If Me.Optcodigo.Value = True Then
            cadAux = Orden1
        Else
            cadAux = Orden2
        End If
        cadParam = "|pOrden=" & cadAux & "|"
        numParam = 1
        
        'Añadir el parametro de Empresa
        cadParam = cadParam & "pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
        
        If Trim(txtCodigo(1).Text) <> "" Or Trim(txtCodigo(2).Text) <> "" Then
            'Cadena para seleccion Desde y Hasta
            If Opcionlistado = 4 Or Opcionlistado = 110 Then
                '4: Listado Tipos de Articulos, 110: List. Ubicaciones
                cadFormula = CadenaDesdeHasta(txtCodigo(1).Text, txtCodigo(2).Text, Codigo, "T")
            Else
                cadFormula = CadenaDesdeHasta(txtCodigo(1).Text, txtCodigo(2).Text, Codigo, "N")
            End If
            
            If cadFormula <> "" Then
                If cadFormula = "Error" Then Exit Sub
                cadAux = ""
                If txtCodigo(1).Text <> "" Then cadAux = "Desde: " & txtCodigo(1).Text & " " & txtNombre(1).Text
                If txtCodigo(2).Text <> "" Then
                    If cadAux <> "" Then cadAux = cadAux & "  -  "
                    cadAux = cadAux & " Hasta: " & txtCodigo(2).Text & " " & txtNombre(2).Text
                End If
                cadParam = cadParam & "pDesde=""" & cadAux & """|"
                numParam = numParam + 1
            End If
        End If
        
    '========= Frame Informes Almacen ========================================
    Case 2 'Frame Informes Almacen
        If Opcionlistado = 7 Then '7: Traspaso Almacen
            indRPT = 1
            cadAux = "scatra"
            cadTitulo = "Informe Traspaso Almacenes"
        ElseIf Opcionlistado = 8 Then '8: Movimientos Almacen
            indRPT = 3
            cadAux = "scamov"
            cadTitulo = "Informe Movimientos Almacen"
        ElseIf Opcionlistado = 10 Then '10: Movimientos de Servicios
            indRPT = 79
            cadAux = "scaser"
            cadTitulo = "Informe Movimientos de Servicios"
            conSubRPT = True
        End If
        
        cadParam = "|"
        If Not PonerParamEmpresa(cadParam, numParam) Then Exit Sub
        If PonerParamRPT(indRPT, cadParam, numParam, cadNomRPT) Then
            'Cadena para seleccion Desde y Hasta DOCUMENTO
            '----------------------------------------------
            If txtCodigo(3).Text <> "" Or txtCodigo(4).Text <> "" Then
                If Not PonerDesdeHasta(Codigo, "N", 3, 4, "") Then Exit Sub
            End If
            If Opcionlistado = 10 Then
                If Not AnyadirAFormula(cadFormula, "{scaser.clisoc}=" & TipoMto) Then Exit Sub
                If Not AnyadirAFormula(cadselect, "{scaser.clisoc}=" & TipoMto) Then Exit Sub
            End If
            If Not HayRegParaInforme(cadAux, cadselect) Then Exit Sub
        End If
                       
                   
                   
    '========= Frame Listado Movimiento de Artículos ========================
    Case 3 'Frame Listado Movimiento de Artículos
        'Nombre fichero .rpt a Imprimir
            ' ordenado por familia/articulo, como inicialmente estaba
        If DeServicios Then
            ' nuevo ordenado por cliente / proveedor / socio
            If ChKAcumulados.Value Then
                cadNomRPT = "rAlmMovimCliSocAcu.rpt"
            Else
                cadNomRPT = "rAlmMovimCliSoc.rpt"
            End If
            
            If Not PonerFormulaYParametrosInf9() Then Exit Sub
            'comprobar que hay datos para mostrar en el Informe
            cadAux = "smoval INNER JOIN sartic ON smoval.codartic=sartic.codartic "
            If Not HayRegParaInforme(cadAux, cadselect) Then Exit Sub
            conSubRPT = True
            If Not CargarTablaTemporal(cadselect) Then Exit Sub
            
            If ChKAcumulados.Value Then
                If Not RellenarTablaTemporalAcum2(cadselect) Then Exit Sub
            End If
            
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
        Else
            cadNomRPT = "rAlmMovim.rpt"
            
            If Not PonerFormulaYParametrosInf9() Then Exit Sub
            'comprobar que hay datos para mostrar en el Informe
            cadAux = "smoval INNER JOIN sartic ON smoval.codartic=sartic.codartic "
            If Not HayRegParaInforme(cadAux, cadselect) Then Exit Sub
            conSubRPT = True
        End If
    
    '========= Frame de Inventario ==========================================
    Case 4 'Frame de Inventario
        If Not ValidarCamposInventario Then Exit Sub
        If Opcionlistado = 19 Then
            cadNomRPT = "rAlmStocksFecha.rpt"
        Else
            'Nombre fichero .rpt a Imprimir
            If vParamAplic.InventarioxProv Then 'Se realiza inventario por Proveedor
                                                'Ordenar por: codprove, codfamia, codartic
                Select Case Opcionlistado
                    Case 12: cadNomRPT = "rAlmInvenxProv.rpt"  'Listado Toma de Inventario
                    Case 13: cadNomRPT = "rAlmInvenxProvDif.rpt"  'Listado Diferencias
                    Case 15: cadNomRPT = "rAlmArtInactivos.rpt"
                    Case 16: cadNomRPT = "rAlmInvenxProvValoracion.rpt"  'Listado Valoracion Stocks Inventariados
                    Case 17: cadNomRPT = "rAlmValoracionxProv.rpt"  'Listado Valoracion Stocks (Por Proveedor)
                End Select
            Else 'Ordenar por Cod. Familia y no por Proveedor. Ordenar por: codfamia, codartic.
                Select Case Opcionlistado
                    Case 12: cadNomRPT = "rAlmInventario.rpt"  'Listado Toma de Inventario
                    Case 13: cadNomRPT = "rAlmInventarioDif.rpt"  'Listado Diferencias
                    Case 15: cadNomRPT = "rAlmArtInactivos.rpt"
                    Case 16: cadNomRPT = "rAlmInvenValoracion.rpt"  'Listado Valoracion Stocks Inventariados
                    Case 17: cadNomRPT = "rAlmValoracion.rpt"  'Listado Valoracion Stocks)
                End Select
            End If
        End If
        Screen.MousePointer = vbHourglass
        DoEvents
        bol = PonerFormulaYParametrosInf12()
        Screen.MousePointer = vbDefault
        If Not bol Then Exit Sub
        
   End Select
    
       
   If Opcionlistado = 14 Then 'Actualizar Inventario (NO IMPRIME INFORME)
'--monica no hay trabajador
'        If Trim(txtCodigo(21).Text) <> "" Then
            'Quitar las llaves:{tabla.codigo} de la cadena consulta
            'para el FormulaSelection del informe Crystal Report y
            'Tendremos la clausula WHERE para insertar en la tabla:sinven
            cadAux = QuitarCaracterACadena(cadFormula, "{")
            cadFormula = QuitarCaracterACadena(cadAux, "}")
            If ActualizarInventario Then
                MsgBox "La Actualización de Inventario se ha realizado correctamente.", vbInformation
            End If
'        Else
'            MsgBox "El campo Trabajador debe tener valor", vbInformation
'            PonerFoco txtCodigo(21)
'            Exit Sub
'        End If
        
   Else 'Listados
        If Opcionlistado = 19 Then cadFormula = ""
        
        LlamarImprimir

        'Realizar otras acciones segun el informe que llame
        Select Case Opcionlistado
            Case 12 'Toma de Inventario
                If frmVisReport.EstaImpreso = True Then
                    PrepararTomaInventario
                End If
            Case 7, 8, 10 ' 7, 8 Movimientos de Almacen
                          ' 10 Movimientos de Servicios
                ActualizarImprimir
            Case 19
                DescargarDatosTMPStockFecha
        End Select
        
   End If
   Screen.MousePointer = vbDefault
End Sub


Private Sub PrepararTomaInventario()
Dim cadAux As String

    On Error GoTo ETomaInv
    
    If MsgBox("¿Impresión correcta para Actualizar Inventario?", vbQuestion + vbYesNo) = vbYes Then
        'Quitar las llaves:{tabla.codigo} de la cadena consulta
        'para el FormulaSelection del informe Crystal Report y
        'Tendremos la clausula WHERE para insertar en la tabla:sinven
'                cadAux = QuitarCaracterACadena(cadFormula, "{")
'                cadFormula = QuitarCaracterACadena(cadAux, "}")
       If CrearTmpInventario(cadselect) Then
            If InsertarInventario Then
                MsgBox "Puede pasar a realizar la Entrada de Inventario Real", vbInformation
            End If
       End If
       cadAux = "DROP TABLE IF EXISTS tmpInven "
       conn.Execute cadAux
    End If
    
ETomaInv:
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub cmdAceptarArtic_Click()
'Listado de Articulos
Dim campo As String
Dim devuelve As String
Dim Opcion As Byte, numOp As Byte
Dim cadFrom As String





    InicializarVbles
    
    'Si es informe=18 de Stocks Maximos y Minimos comprobar
    'que se ha seleccionado un almacen
    Select Case Opcionlistado
    Case 18
        'If OpcionListado = 18 Then
        If txtCodigo(72).Text = "" Then
            MsgBox "Se debe seleccionar un Almacen para el informe.", vbInformation
            Exit Sub
        End If
        cadNomRPT = "rAlmStocksMaxMin.rpt"
        cadFrom = " salmac INNER JOIN sartic ON salmac.codartic=sartic.codartic "
    Case 247
        '
        If txtCodigo(107).Text = "" Or txtNombre(107) = "" Then
            MsgBox "Debe seleccionar una tarifa para el informe.", vbInformation
            Exit Sub
        End If
        
    Case Else
        'El 6
        cadNomRPT = "rAlmListArticulos.rpt"  'Nombre fichero .rpt a Imprimir
        cadFrom = " sartic"
    End Select
    
    '===================================================
    '============ PARAMETROS ===========================
    cadParam = "|"
    'Empresa
    cadParam = cadParam & "pEmpresa=""" & vParam.NombreEmpresa & """|"
    numParam = numParam + 1
    
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion  ALMACEN
    '--------------------------------------------
    If Opcionlistado = 18 And txtCodigo(72).Text <> "" Then
        campo = "{salmac.codalmac}"
        cadFormula = campo & "= " & txtCodigo(72).Text
    Else
        'Es tarifa para la correccion
        If Opcionlistado = 247 And txtCodigo(107).Text <> "" Then
            campo = "{slista.codlista}"
            cadFormula = campo & "= " & txtCodigo(107).Text
        End If
    End If
    
    
    'Cadena para seleccion D/H FAMILIA
    '--------------------------------------------
     If txtCodigo(62).Text <> "" Or txtCodigo(63).Text <> "" Then
        campo = "{sartic.codfamia}"
        'Parametro Desde/Hasta Familila
        devuelve = "pDHFamilia=""Familia: "
        If Not PonerDesdeHasta(campo, "N", 62, 63, devuelve) Then Exit Sub
    End If
    
'    'Cadena para seleccion D/H MARCA
'    '--------------------------------------------
'    If txtCodigo(64).Text <> "" Or txtCodigo(65).Text <> "" Then
'        campo = "{sartic.codmarca}"
'        'Parametro Desde/Hasta Marca
'        devuelve = "pDHMarca=""Marca: "
'        If Not PonerDesdeHasta(campo, "N", 64, 65, devuelve) Then Exit Sub
'    End If
    
    'Cadena para seleccion D/H PROVEEDOR
    '--------------------------------------------
    If txtCodigo(66).Text <> "" Or txtCodigo(67).Text <> "" Then
        campo = "{sartic.codprove}"
        'Parametro Desde/Hasta Proveedor
        devuelve = "pDHProveedor=""Proveedor: "
        If Not PonerDesdeHasta(campo, "N", 66, 67, devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion D/H TIPO ARTICULO
    '--------------------------------------------
    If txtCodigo(68).Text <> "" Or txtCodigo(69).Text <> "" Then
        campo = "{sartic.codtipar}"
        'Parametro Desde/Hasta Tipo Articulo
        devuelve = "pDHTipoArt=""Tipo Articulo: "
        If Not PonerDesdeHasta(campo, "T", 68, 69, devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion D/H ARTICULO
    '--------------------------------------------
    If txtCodigo(70).Text <> "" Or txtCodigo(71).Text <> "" Then
        campo = "{sartic.codartic}"
        'Parametro Desde/Hasta Articulo
        devuelve = "pDHArticulo=""Articulo: "
        If Not PonerDesdeHasta(campo, "T", 70, 71, devuelve) Then Exit Sub
    End If
    
    'Obtener el parametro con el Orden del Informe
    '---------------------------------------------
    Select Case Opcionlistado
    Case 6
    ''''If OpcionListado = 6 Then '6: Listado de Articulos
        numOp = PonerGrupo(1, ListView2.ListItems(1).Text)
        If numOp <> 0 Then Opcion = numOp
        numOp = PonerGrupo(2, ListView2.ListItems(2).Text)
        If numOp <> 0 Then Opcion = numOp
        numOp = PonerGrupo(3, ListView2.ListItems(3).Text)
        If numOp <> 0 Then Opcion = numOp
        numOp = PonerGrupo(4, ListView2.ListItems(4).Text)
        If numOp <> 0 Then Opcion = numOp
        Opcion = Opcion - 1
    
        Select Case Opcion
            Case 1 'El group2 es el Proveedor
                campo = "pTitulo1=""" & ListView2.ListItems(3).Text & """"
                cadParam = cadParam & campo & "|"
                numParam = numParam + 1
                
                campo = "pTitulo2=""" & ListView2.ListItems(4).Text & """"
                cadParam = cadParam & campo & "|"
                numParam = numParam + 1
            Case 2 'El Group3 es el Proveedor
                campo = "pTitulo1=""" & ListView2.ListItems(2).Text & """"
                cadParam = cadParam & campo & "|"
                numParam = numParam + 1
                
                campo = "pTitulo2=""" & ListView2.ListItems(4).Text & """"
                cadParam = cadParam & campo & "|"
                numParam = numParam + 1
            Case 3, 0 'El Group4 es el Proveedor
                      '0 'El Group1 es el Proveedor
                campo = "pTitulo1=""" & ListView2.ListItems(2).Text & """"
                cadParam = cadParam & campo & "|"
                numParam = numParam + 1
                
                campo = "pTitulo2=""" & ListView2.ListItems(3).Text & """"
                cadParam = cadParam & campo & "|"
                numParam = numParam + 1
                
                If Opcion = 0 Then
                    campo = "pTitulo3=""" & ListView2.ListItems(4).Text & """"
                    cadParam = cadParam & campo & "|"
                    numParam = numParam + 1
                End If
        End Select
       
        'Parametro Orden del Informe
        campo = "pOrden=" & Opcion
        cadParam = cadParam & campo & "|"
        numParam = numParam + 1
        
    Case 18
    ''ElseIf OpcionListado = 18 Then
        'Filtrar ademas por stock<stockMin o stock>stockMax
        campo = "{salmac.canstock}"
        If Me.optStockMax Then
            cadFormula = cadFormula & " AND (" & campo & "> {salmac.stockmax})"
        Else
            'David G 30/01/2007
            If optPuntoPedido.Value Then
                cadFormula = cadFormula & " AND (" & campo & "< {salmac.puntoped})"
            Else
                cadFormula = cadFormula & " AND (" & campo & "< {salmac.stockmin})"
            End If
        End If
    
        'Añadir el Parametro de Stocks Maximos o Minimos
        If Me.optStockMax.Value = True Then
            campo = "0"
        Else
            If optPuntoPedido.Value Then
                campo = "2"
            Else
                campo = "1"
            End If
        End If
        cadParam = cadParam & "pStockMax=" & campo & "|"
        numParam = numParam + 1
    Case 247

        'Correccion de importes
        
        'Mostrare el list
        cadselect = QuitarCaracterACadena(cadFormula, "{")
        cadselect = QuitarCaracterACadena(cadselect, "}")
        frmMensajes.cadwhere = cadselect
        frmMensajes.OpcionMensaje = 16
        frmMensajes.vCampos = txtCodigo(107).Text
        frmMensajes.cadWHERE2 = Trim(Me.cmbDecimales.Text)
        'Por no utilizar otra variable
        NumRegElim = 0
        If Me.chkMinimoCorreg.Value = 1 Then NumRegElim = 1
        frmMensajes.Show vbModal
        Exit Sub
    End Select
    
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    cadselect = cadFormula
    If Not HayRegParaInforme(cadFrom, cadselect) Then Exit Sub
        
    LlamarImprimir
End Sub


Private Sub cmdAceptarDtosFM_Click()
'54: Listado de Descuentos Familia/Marca
'309: Listado precio compras
Dim campo As String, Cad As String
Dim tabla As String

    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
        
    If Opcionlistado = 54 Then
        tabla = "sdtofm"
        conSubRPT = True
    ElseIf Opcionlistado = 309 Then
        tabla = "slispr"
        cadTitulo = "Listado Precios de compra"
        cadNomRPT = "rComPrecios.rpt"
        conSubRPT = False
    End If
    
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion D/H FAMILIA
    '----------------------------------
    If txtCodigo(75).Text <> "" Or txtCodigo(76).Text <> "" Then
        campo = "{" & tabla & ".codfamia}"
        If Opcionlistado = 309 Then campo = "{sartic.codfamia}"
        Cad = "pDHFamilia=""Familia: "
        If Not PonerDesdeHasta(campo, "N", 75, 76, Cad) Then Exit Sub
    End If

    If Opcionlistado = 54 Then
        'Cadena para seleccion D/H CLIENTE
        '--------------------------------------------
        If txtCodigo(73).Text <> "" Or txtCodigo(74).Text <> "" Then
            campo = "{sdtofm.codclien}"
            Cad = "pDHCliente=""Cliente: "
            If Not PonerDesdeHasta(campo, "N", 73, 74, Cad) Then Exit Sub
        End If
    
    
        'Cadena para seleccion D/H MARCA
        '--------------------------------------------
        If txtCodigo(77).Text <> "" Or txtCodigo(78).Text <> "" Then
            campo = "{sdtofm.codmarca}"
            Cad = "pDHMarca=""Marca: "
            If Not PonerDesdeHasta(campo, "N", 77, 78, Cad) Then Exit Sub
        End If
    ElseIf Opcionlistado = 309 Then
        'Cadena para seleccion D/H PROVEEDOR
        '--------------------------------------------
        If txtCodigo(79).Text <> "" Or txtCodigo(80).Text <> "" Then
            campo = "{" & tabla & ".codprove}"
            Cad = "pDHProveedor=""Proveedor: "
            If Not PonerDesdeHasta(campo, "N", 79, 80, Cad) Then Exit Sub
        End If
    End If
    
    '==============================================================
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If Opcionlistado = 309 Then tabla = tabla & " INNER JOIN sartic ON " & tabla & ".codartic=sartic.codartic"
    If Not HayRegParaInforme(tabla, cadselect) Then Exit Sub
    
    LlamarImprimir
End Sub




Private Sub cmdAceptarRepxDia_Click()
'Reparaciones por Dia
Dim devuelve As String
Dim param As String
Dim TotalMante As Integer
Dim Rs As ADODB.Recordset
Dim fecha1 As String, fecha2 As String
Dim NomTabla As String

    InicializarVbles

    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1


    Select Case Opcionlistado
        Case 63
            Codigo = "{scarep.fecentre}"
            param = "pDHFecha=""Fecha Rep.: "
            NomTabla = "scarep"
            cadNomRPT = "rRepReparxDia.rpt"
            conSubRPT = True
            cadTitulo = "Reparaciones por día"
        Case 73
            'Añadir el parametro total Mantenim. si estamos en Informe de Altas
            devuelve = "SELECT DISTINCT COUNT(*) FROM scaman "
            Set Rs = New ADODB.Recordset
            Rs.Open devuelve, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not Rs.EOF Then
                TotalMante = Rs.Fields(0).Value
                cadParam = cadParam & "pTotalMante=" & TotalMante & "|"
                numParam = numParam + 1
            End If
            Rs.Close
            Set Rs = Nothing

            'Añadir el Total Mantenim. del Periodo anterior
            fecha1 = Day(txtCodigo(31).Text) & "/" & Month(txtCodigo(31).Text) & "/" & Year(txtCodigo(31).Text) - 1
            fecha2 = Day(txtCodigo(32).Text) & "/" & Month(txtCodigo(32).Text) & "/" & Year(txtCodigo(32).Text) - 1
            Codigo = "scaman.fechaini"
            devuelve = CadenaDesdeHastaBD(fecha1, fecha2, Codigo, "F")
            If devuelve <> "" And devuelve <> "Error" Then
                devuelve = "SELECT DISTINCT COUNT(*) FROM scaman WHERE " & devuelve
                Set Rs = New ADODB.Recordset
                Rs.Open devuelve, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not Rs.EOF Then
                    TotalMante = Rs.Fields(0).Value
                    cadParam = cadParam & "pTotalAnte=" & TotalMante & "|"
                    numParam = numParam + 1
                End If
                Rs.Close
                Set Rs = Nothing
            End If

            '================= FORMULA =========================
            Codigo = "{scaman.fechaini}"
            param = "pDHFecha=""Fecha: "
            NomTabla = "scaman"
            cadNomRPT = "rManListAltas.rpt"
            cadTitulo = "Informe Altas Mantenimientos"

        Case 223
            param = ""
            If Me.OptClientes Then
                Codigo = "{facturas.fecfactu}"
                NomTabla = "facturas"
                If CadTag = "A" Then
                    Codigo = "{facturassocio.fecfactu}"
                    NomTabla = "facturassocio"
                End If
            Else
'++ monica
                If Not Me.OptProve Then
    
                    Codigo = "{tcafpc.fecrecep}"
                    NomTabla = "tcafpc"
                Else
                    Codigo = "{scafpc.fecrecep}"
                    NomTabla = "scafpc"
                End If
           End If
    End Select


    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Desde y Hasta FECHA
    If Opcionlistado = 223 Then
        'comprobar que se han rellenado los dos campos de fecha
        'sino rellenar con fechaini o fechafin del ejercicio
        'que guardamos en vbles Orden1,Orden2
        If txtCodigo(31).Text = "" Then
           txtCodigo(31).Text = Orden1 'fechaini del ejercicio de la conta
        End If

        If txtCodigo(32).Text = "" Then
           txtCodigo(32).Text = Orden2 'fecha fin del ejercicio de la conta
        End If

         'Comprobar que el intervalo de fechas D/H esta dentro del ejercicio de la
         'contabilidad par ello mirar en la BD de la Conta los parámetros
        If Not ComprobarFechasConta(31) Then Exit Sub
        If Not ComprobarFechasConta(32) Then Exit Sub
        
        '++monica: comprobar si es factura de cliente que se ponen los datos de tesoreria
'        If txtCodigo(33).Text = "" Then
'            MsgBox "Debe introducir los datos de tesoreria.", vbExclamation
'            PonerFoco txtCodigo(33)
'            Exit Sub
'        End If
        If Me.OptClientes Then
            If txtCodigo(0).Text = "" Then
                MsgBox "Debe introducir los datos de tesoreria.", vbExclamation
                PonerFoco txtCodigo(0)
                Exit Sub
            End If
        End If
            
    End If

    devuelve = CadenaDesdeHasta(txtCodigo(31).Text, txtCodigo(32).Text, Codigo, "F")
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    'Parametro D/H Fecha
    If devuelve <> "" And param <> "" Then
        cadParam = cadParam & AnyadirParametroDH(param, 31, 32) & """|"
        numParam = numParam + 1
    End If

    '[Monica]24/07/2017: añadimos el tipo de movimiento
    If Text1(6).Text <> "" Then
        If Not AnyadirAFormula(cadFormula, NomTabla & ".codtipom = """ & Text1(6).Text & """") Then Exit Sub
    End If



    '===================================================
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    cadselect = CadenaDesdeHastaBD(txtCodigo(31).Text, txtCodigo(32).Text, Codigo, "F")
    If Opcionlistado = 223 Then cadselect = cadselect & " AND " & NomTabla & ".intconta=0 "
    
    '[Monica]24/07/2017: añadimos el tipo de movimiento
    If Opcionlistado = 223 And Text1(6).Text <> "" Then
        If Not AnyadirAFormula(cadselect, NomTabla & ".codtipom = '" & Text1(6).Text & "'") Then Exit Sub
    End If
    
    
    
    If Not HayRegParaInforme(NomTabla, cadselect) Then Exit Sub

    If Opcionlistado <> 223 Then
        LlamarImprimir
    Else


        '------------------------------------------------------------------------------
        '  LOG de acciones.                      5: Facturas compras
        Set LOG = New cLOG
        LOG.Insertar 5, vUsu, "Contabilizar facturas compras: " & vbCrLf & NomTabla & vbCrLf & cadselect
        Set LOG = Nothing
        '-----------------------------------------------------------------------------


        ContabilizarFacturas NomTabla, cadselect
        TerminaBloquear
         'Eliminar la tabla TMP
        BorrarTMPFacturas
        'Desbloqueamos ya no estamos contabilizando facturas
        If Me.OptClientes.Value Then
            DesBloqueoManual ("VENCON") 'VENtas CONtabilizar
        Else
            If Me.OptProve.Value Then
                DesBloqueoManual ("COMCON") 'COMpras CONtabilizar
            Else
                DesBloqueoManual ("TRACON") 'TRAnsporte CONtabilizar
            End If
        End If
        Me.FrameProgress.visible = False
        Me.FrameRepxDia.Height = 3500
        Me.Height = Me.FrameRepxDia.Height + 350
        
        Unload Me
    End If
End Sub

Private Sub cmdBajar_Click()
'Bajar el item seleccionado del listview2
    BajarItemList Me.ListView2
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub


Private Sub cmdDeselTodos_Click()
Dim i As Byte

    For i = 1 To ListView1.ListItems.Count
        ListView1.ListItems(i).Checked = False
    Next i
End Sub

Private Sub cmdHcoMante_Click()
    Codigo = ""
    For indCodigo = 110 To 112
        If txtCodigo(indCodigo).Text = "" Then Codigo = Codigo & "M"
        If indCodigo > 110 Then If txtNombre(indCodigo).Text = "" Then Codigo = Codigo & "M"
    Next indCodigo
    If Codigo <> "" Then
        MsgBox "Rellene correctamente todos los datos", vbExclamation
        Exit Sub
    End If
    'CUATRO CAMPOS. El primero de control
    CadenaDesdeOtroForm = "OK|" & txtCodigo(110).Text & "|" & txtNombre(111).Text & "|" & txtCodigo(112).Text & "|"
    Unload Me
End Sub


Private Sub cmdRecalPMP_Click()
Dim miSQL As String

        InicializarVbles
    
        'Añadir el parametro de Empresa
        cadParam = cadParam & "pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
            
      
        If txtCodigo(24).Text <> "" Or txtCodigo(25).Text <> "" Then
            miSQL = " Art: "
            Codigo = "{sartic.codartic}"
            If Not PonerDesdeHasta(Codigo, "T", 24, 25, miSQL) Then Exit Sub
        End If
        
        If txtCodigo(28).Text <> "" Or txtCodigo(29).Text <> "" Then
            miSQL = " Prov: "
            Codigo = "{sartic.codprove}"
            If Not PonerDesdeHasta(Codigo, "N", 28, 29, miSQL) Then Exit Sub
        End If
        If txtCodigo(26).Text <> "" Or txtCodigo(27).Text <> "" Then
            miSQL = " Fam: "
            Codigo = "{sartic.codfamia}"
            If Not PonerDesdeHasta(Codigo, "N", 26, 27, miSQL) Then Exit Sub
        End If
        
        

        If cadselect <> "" Then cadselect = " AND " & cadselect
        cadselect = "sartic.codartic=smoval.codartic " & cadselect   'COMPRAS
        
        
        If Not HayRegParaInforme("sartic,smoval", cadselect) Then
            MsgBox "No hay datos  con estos valores", vbExclamation
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        If RecalculoPrecioMedioPonderado Then
            frmMensajes.OpcionMensaje = 24
            frmMensajes.Show vbModal
            Unload Me
        End If
        Screen.MousePointer = vbDefault

End Sub

Private Sub cmdSelTodos_Click()
Dim i As Byte

    For i = 1 To ListView1.ListItems.Count
        ListView1.ListItems(i).Checked = True
    Next i
End Sub


Private Sub cmdSubir_Click()
'Subir el item seleccionado del listview2 una posicion
    SubirItemList Me.ListView2
End Sub




Private Sub Form_Activate()
Dim IndiceFoco As Integer

    If PrimeraVez Then
        PrimeraVez = False
        IndiceFoco = -1
        Select Case Opcionlistado
        Case 1, 2, 3, 4, 61, 20, 21, 22, 23, 24, 27, 58, 110
            '1:Listado de Marcas, 2:Almacenes Propios, 3:Tipos de Unidad
            '4:Tipos de Artículos, 6:Artículos
            '61:Motivos Pen. Rep
            '58:Proveedores, 110:Ubicaciones
             'PonerFoco txtCodigo(1)
             IndiceFoco = 1
        Case 6 '6: Informe de Articulos
            'PonerFoco txtCodigo(62)
            IndiceFoco = 62
        Case 7, 8 '7: Informe Traspaso Almacenes/Historico
                  '8: Informe Movimientos Almacen/Historico
            'PonerFoco txtCodigo(3)
            IndiceFoco = 3
        Case 9 'Informe Movimientos Artículos
            'PonerFoco txtCodigo(5)
            IndiceFoco = 5
        Case 12, 13, 14, 15, 16, 17, 19 '12: Listado Toma de Inventario Articulos
                        '13: Listado Diferencias de Inventario Articulos
                        '14: Actualizar Diferencias de Inventario (No IMPRIME INFORME)
                        '15: Listado Articulos Inactivos
                        '16: Listado Valoracion de Stocks Inventariados
                        '17: Listado Valoración Stocks
                        '19: Inf. Stocks a una Fecha
            'PonerFoco txtCodigo(13)
            IndiceFoco = 13
        Case 18      '18: Informe Stocks MAximos y Minimos
            'PonerFoco txtCodigo(72)
            IndiceFoco = 72
        Case 28, 29, 30 '28: Informe Tarifas de Articulos
                    '29: Informe Promociones
                    '30: Informe Precios Especiales
            'PonerFoco txtCodigo(23)
            IndiceFoco = 23
        Case 31, 73 '31: Informe Ofertas
                    '73: Listado Altas Mantenimientos
            'PonerFoco txtCodigo(31)
            IndiceFoco = 31
        Case 54 'Listado Descuentos Familia/ Marca
            'PonerFoco txtCodigo(73)
            IndiceFoco = 73
        Case 60 '60: Informe Reparacions - Nº Series
            'PonerFoco txtCodigo(37)
            IndiceFoco = 37
        Case 63, 73, 223 '63: Listado Reparaciones x día
                         '223: Contabilizar facturas
            'PonerFoco txtCodigo(31)
            IndiceFoco = 31
        Case 246 '246: Informe margen ventas x articulo
            'PonerFoco txtCodigo(88)
            IndiceFoco = 88
        Case 64, 406 '64: Listado Reparaciones x Cliente
                     '406: List. Frecuencia de Reparaciones
            'PonerFoco txtCodigo(33)
            IndiceFoco = 33
        Case 70, 71, 76, 79 'Listado Mantenimientos
            'PonerFoco txtCodigo(45)
            IndiceFoco = 45
        Case 72 'Informe Fichas Mantenimientos
            'PonerFoco txtCodigo(55)
            IndiceFoco = 55
            
        Case 77
            'PonerFoco txtCodigo(102)
             IndiceFoco = 102
        Case 78
            'PonerFoco txtCodigo(109)
            IndiceFoco = 109
            
        Case 82, 83
            'Marca facturar a 1
            IndiceFoco = 119
            
        Case 309 '309:Listado precios de compra
            'PonerFoco txtCodigo(79)
            IndiceFoco = 79
        Case 407 'Sustitución Nº Serie
            'PonerFoco txtCodigo(81)
            IndiceFoco = 81
        Case 409 'List. Avisos de averias pendientes
            'PonerFoco txtCodigo(82)
            IndiceFoco = 82
            
        Case 99
            'PonerFoco txtCodigo(110)
            IndiceFoco = 110
        Case 247  'y Correccion de listados de precios tarias etc
             'PonerFoco txtCodigo(107)
             IndiceFoco = 107
        End Select
        If IndiceFoco >= 0 Then PonerFoco txtCodigo(IndiceFoco)
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer


    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    limpiar Me

    'Ocultar todos los Frames de Formulario
    frameListado.visible = False
    FrameInfAlmacen.visible = False
    FrameMovArtic.visible = False
    Me.FrameInventario.visible = False
    Me.FrameRepxDia.visible = False
    Me.FrameDtosFM.visible = False
    
    FrameEnvioMail.visible = False
    FrameHcoMante.visible = False
    FrameRecalPMP.visible = False
    
    CommitConexion
    
    CargarIconos
    
    cadTitulo = ""
    cadNomRPT = ""
    
    Select Case Opcionlistado
        Case 1 To 19, 247 'Listado de ALMACEN
            ListadosAlmacen H, W
        Case 100 To 199 'Listados de ALMACEN
            ListadosAlmacen H, W
        Case 300 To 390 'Listados de COMPRAS
            ListadosCompras H, W
    End Select
    
    
    Select Case Opcionlistado
    
        
    'LISTADOS DE REPARACIONES
    '-------------------------
    Case 223, 224
        If Opcionlistado = 224 Then Me.OptClientes = False
        PonerFrameRepxDiaVisible True, H, W
        indFrame = 7
            
        '[Monica]24/07/2017: introducimos el codtipom
        Text1(6).Text = ""
        If CadTag = "A" Or Opcionlistado = 224 Then
            Text1(6).visible = False
            Label1(0).visible = False
        End If
        
        '++monica:15102008
        If Opcionlistado = 223 Then
            txtCodigo(31).Text = Format(Now, "dd/mm/yyyy")
            txtCodigo(32).Text = Format(Now, "dd/mm/yyyy")
        End If
    Case 99
        
        H = Me.FrameHcoMante.Height
        W = Me.FrameHcoMante.Width
        PonerFrameVisible FrameHcoMante, True, H, W
        indFrame = 99
        cadTitulo = "Pasar a mantenimientos anulados"
        conSubRPT = False
        txtCodigo(110).Text = Format(Now, "dd/mm/yyyy")

    End Select
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub

Private Sub Form_Unload(Cancel As Integer)
    NumCod = ""
End Sub




Private Sub frmcta_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtCodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmMtoAlPropios_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Almacenes Propios
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoArticulos_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Articulos
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoClientes_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Clientes
    If indCodigo > 0 Then
        txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
        txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmMtoFamilia_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Familia de Articulos
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoIncid_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoProveedor_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Proveedores
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoTArticulo_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Tipos de Artículo
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoTUnidad_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Tipos de Unidad
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           ' "____________________________________________________________"
'            vCadena = "ORDEN CLIENTE / PROVEEDOR / SOCIO " & vbCrLf & _
'                      "============================================" & vbCrLf & vbCrLf & _
'                      "Si está marcado el orden se cargará una tabla intermedia con los  " & vbCrLf & _
'                      "datos de Cliente / Proveedor / Socio dependiendo del tipo de movimiento." & vbCrLf & _
'                      "" & vbCrLf & vbCrLf & _
'                      "En caso contrario el listado se saca ordenado por familia y artículo. " & vbCrLf & _
'                      "" & vbCrLf & vbCrLf
                      
            vCadena = "ORDEN CLIENTE / SOCIO " & vbCrLf & _
                      "============================================" & vbCrLf & vbCrLf & _
                      "Para que funcione el orden del listado únicamente ha de estar" & vbCrLf & _
                      "marcado el tipo de movimiento SES ( Movimientos de servicios de" & vbCrLf & _
                      "Socio ) y/o el tipo SEC ( Movimientos de servicios de Cliente ) " & vbCrLf & vbCrLf & _
                      "En caso contrario el listado se saca ordenado por familia y artículo. " & vbCrLf & _
                      "" & vbCrLf & vbCrLf
                      
                      
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"
    

End Sub

Private Sub imgBuscar_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    imgBuscar(1).Tag = Index
    indCodigo = Index
    
    Select Case Index
    Case 1, 2 'FrameListado
        Select Case Opcionlistado
            Case 2 'Listado de ALMACENES Propios
                AbrirFrmAlmPropios
            
            Case 3  'Listado de Tipos de Unidad
                Set frmMtoTUnidad = New frmManTipUnid
                frmMtoTUnidad.DatosADevolverBusqueda = "0|1"
                frmMtoTUnidad.DeConsulta = True
                frmMtoTUnidad.Show vbModal
                Set frmMtoTUnidad = Nothing
            
            Case 4  'Listado de Tipos de Articulos
                AbrirFrmTipoArt

            Case 58
                'DAVID
                indCodigo = Index
                Set frmMtoProveedor = New frmManProve
                frmMtoProveedor.DatosADevolverBusqueda = "0|1|"
                frmMtoProveedor.Show vbModal
                Set frmMtoProveedor = Nothing
        End Select
        
    Case 3, 4 'FrameInfAlmacen
            If Opcionlistado = 7 Or Opcionlistado = 8 Then
'            Case 7, 8 '7: Informe de Traspasos de Almacenes
                  '8: Informe de Movimientos de Almacen
                MandaBusquedaPrevia ""
            End If
    End Select
    
    PonerFoco Me.txtCodigo(Index)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgBuscarG_Click(Index As Integer)
'Buscar general: cada index llama a una tabla
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 0, 1, 6, 7, 36, 43, 44, 49, 50, 75, 76, 77, 80, 81, 93, 94  'cod. CLIENTE
            Select Case Index
                Case 0, 1: indCodigo = Index + 73
                Case 6, 7: indCodigo = Index + 27
                Case 36: indCodigo = Index + 20
                Case 43, 44: indCodigo = Index + 4
                Case 49, 50: indCodigo = Index - 12
                Case 75: indCodigo = 0
                Case 76, 77, 80, 81: indCodigo = Index + 22
                Case 93, 94: indCodigo = Index + 24
            End Select
            AbrirFrmClientes
        
        Case 2, 3, 13, 14, 17, 19, 20, 21, 31, 32, 57, 58, 67, 68, 73, 74 'cod. FAMILIA
            Select Case Index
                Case 2, 3: indCodigo = Index + 73
                Case 13, 14: indCodigo = Index + 3
                Case 19, 20: indCodigo = Index + 43
                Case 31, 32: indCodigo = Index - 24
                Case 57, 58: indCodigo = Index - 32
                Case 67, 68, 73, 74: indCodigo = Index + 21
                Case 17: indCodigo = 27
                Case 21: indCodigo = 26
                
            End Select
            Set frmMtoFamilia = New frmManFamilias
            frmMtoFamilia.DatosADevolverBusqueda = "0|1|"
            frmMtoFamilia.Show vbModal
            Set frmMtoFamilia = Nothing
            
            
        Case 90, 91, 92
            indCodigo = 22 + Index
            Set frmMtoIncid = New frmManInciden
            frmMtoIncid.DatosADevolverBusqueda = "0|1|"
            frmMtoIncid.Show vbModal
            Set frmMtoIncid = Nothing
            
            
'        Case 8, 9, 51, 52 'cod. Direc/DPTO
'            Select Case Index
'                Case 8, 9:
'                Case 51, 52: indCodigo = Index - 12
'            End Select
        
        Case 10, 18, 33, 34 'cod. ALMACEN
            Select Case Index
                Case 10: indCodigo = Index + 3
                Case 18: indCodigo = Index + 54
                Case 33, 34: indCodigo = Index - 22
            End Select
            AbrirFrmAlmPropios
            
        Case 11, 12, 22, 27, 28, 29, 30, 35, 61, 62, 69, 70, 71, 72 'cod. ARTICULO
            Select Case Index
                Case 11, 12: indCodigo = Index + 3
                Case 27, 28: indCodigo = Index + 43
                Case 29, 30: indCodigo = Index - 24
                Case 61, 62: indCodigo = Index - 32
                Case 69, 70, 71, 72: indCodigo = Index + 21
                Case 35: indCodigo = 24
                Case 22: indCodigo = 25
            End Select
            Set frmMtoArticulos = New frmManArtic
            frmMtoArticulos.DatosADevolverBusqueda = "0|1|" 'Abrimos en Modo Busqueda
            frmMtoArticulos.Show vbModal
            Set frmMtoArticulos = Nothing
            
        Case 25, 26 'cod TIPO ARTICULO
            indCodigo = Index + 43
            AbrirFrmTipoArt

            
        Case 8, 9, 15, 16, 23, 24, 63, 64 'cod. PROVEEDOR
            Select Case Index
                Case 15, 16: indCodigo = Index + 3
                Case 23, 24: indCodigo = Index + 43
                Case 63, 64: indCodigo = Index + 16
                Case 8, 9: indCodigo = Index + 20
            End Select
            Set frmMtoProveedor = New frmManProve
            frmMtoProveedor.DatosADevolverBusqueda = "0|1|"
            frmMtoProveedor.Show vbModal
            Set frmMtoProveedor = Nothing
            
       Case 98 'cta contable
            indCodigo = 0
            Set frmCta = New frmCtasConta
            frmCta.CodigoActual = txtCodigo(indCodigo)
            frmCta.DatosADevolverBusqueda = "0|1|"
            frmCta.Show vbModal
            Set frmCta = Nothing
            
            
        Case 39, 40, 53, 54 'cod. Nº CONTRATO (= nº mantenimiento)

        Case 87
            indCodigo = 107
    End Select
    PonerFoco txtCodigo(indCodigo)
    Screen.MousePointer = vbDefault
End Sub



Private Sub imgFecha_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim obj As Object

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


   imgFecha(0).Tag = Index
'   Set frmF = New frmCal
   frmF.NovaData = Now
   
   Select Case Index
        Case 0 'frameMovArtic
            indCodigo = 9
        Case 1 'frameMovArtic
            indCodigo = 10
        Case 2 'frameInventario (indFrame=4)
            indCodigo = 20
        Case 3 'frameInventario (indFrame=4)
            indCodigo = 22
        Case 4 'frameReparacionesxDia (indFrame=7)
            indCodigo = 31
        Case 5 'frameReparacionesxDia (indFrame=7)
            indCodigo = 32
        Case 6 'frameReparacionesxClien (indFrame=8)
            indCodigo = 43
        Case 7 'frameReparacionesxClien (indFrame=8)
            indCodigo = 44
        Case 8 'frameMAntenimientos
            indCodigo = 53
        Case 9 'frameMAntenimientos
            indCodigo = 54
        Case 10 'FrameListAvisosPtes
            indCodigo = 82
        Case 11 'FrameListAvisosPtes
            indCodigo = 83
        Case 13, 14
            indCodigo = Index + 102
        Case 15, 16
            indCodigo = Index + 104
        Case 33
            indCodigo = Index
        Case 109
            indCodigo = 109
   End Select
   
   
   PonerFormatoFecha txtCodigo(indCodigo)
   If txtCodigo(indCodigo).Text <> "" Then frmF.NovaData = CDate(txtCodigo(indCodigo).Text)
   
   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtCodigo(indCodigo)
End Sub




Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Optcodigo_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub OptNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.cmdAceptar(1)
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub



Private Sub HabilitarTextoCliente(Habilitar As Boolean)
    If Not Habilitar Then
        txtNombre(10).BackColor = &H80000018
    Else
        txtNombre(10).BackColor = &H80000005
    End If
    txtNombre(10).Locked = Not Habilitar
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim Nregs As Long

    'Quitar espacios en blanco por los lados
    Text1(Index).Text = Trim(Text1(Index).Text)

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
         Case 6 ' tipo de movimiento
            If Text1(Index).Text <> "" Then
                Text1(Index).Text = UCase(Text1(Index).Text)
                Nregs = TotalRegistros("select count(*) from usuarios.stipom stipom where codtipom = " & DBSet(Text1(Index).Text, "T"))
                If Nregs = 0 Then
                    MsgBox "No existe el tipo de movimiento. Reintroduzca.", vbExclamation
                    PonerFoco Text1(Index)
                End If
            End If
    End Select
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub




Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub




Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim tabla As String
Dim codCampo As String, nomCampo As String
Dim TipCampo As String, Formato As String
Dim Titulo As String
Dim EsNomCod As Boolean 'Si es campo Cod-Descripcion llama a PonerNombreDeCod


    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    EsNomCod = False
    If Index = 1 Or Index = 2 Then
    'el mismo frame ( y por tanto los mismos campos) se utilizan para distintos
    'informes. Según de donde llamemos código de una tabla u otra
        Select Case Opcionlistado
            Case 1 'Listado MARCAS
                EsNomCod = True
                tabla = "smarca"
                codCampo = "codmarca"
                nomCampo = "nommarca"
                TipCampo = "N"
                Formato = "0000"
                Titulo = "Marca"
                
            Case 2 'Listado ALMACENES Propios
                EsNomCod = True
                tabla = "salmpr"
                codCampo = "codalmac"
                nomCampo = "nomalmac"
                TipCampo = "N"
                Formato = "000"
                Titulo = "Almacen Propio"
                
            Case 3 'Listado Tipos UNIDADES
                EsNomCod = True
                tabla = "sunida"
                codCampo = "codunida"
                nomCampo = "nomunida"
                TipCampo = "N"
                Formato = "00"
                Titulo = "Tipo Unidad"
                
            Case 4 'Listado Tipos Artículos
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), 1, "stipar", "nomtipar", "codtipar", "Tipo de Artículo", "T")
    
            Case 110 'Listado Ubicaciones Almacen
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "subica", "nomubica", "codubica", "Ubicaciones Almacen", "T")
            
            
            Case 20 'Listado ACTIVIDADES de Clientes
                EsNomCod = True
                tabla = "sactiv"
                codCampo = "codactiv"
                nomCampo = "nomactiv"
                TipCampo = "N"
                Formato = "000"
                Titulo = "Actividad de Cliente"
            
            Case 21 'Listado ZONAS de Clientes
                EsNomCod = True
                tabla = "szonas"
                codCampo = "codzonas"
                nomCampo = "nomzonas"
                TipCampo = "N"
                Formato = "000"
                Titulo = "Zona de Cliente"
            
            Case 22 'Listado RUTAS de Asistencia
                EsNomCod = True
                tabla = "srutas"
                codCampo = "codrutas"
                nomCampo = "nomrutas"
                TipCampo = "N"
                Formato = "000"
                Titulo = "Ruta de Asistencia"
            
            Case 23 'Listado Formas de Envío
                EsNomCod = True
                tabla = "senvio"
                codCampo = "codenvio"
                nomCampo = "nomenvio"
                TipCampo = "N"
                Formato = "000"
                Titulo = "Forma de Envío"
            
            Case 24 'Listado Tarifas Venta
                EsNomCod = True
                tabla = "starif"
                codCampo = "codlista"
                nomCampo = "nomlista"
                TipCampo = "N"
                Formato = "000"
                Titulo = "Tarifa de Venta"
            
            Case 27 'Listado SITUACIONES Especiales
                EsNomCod = True
                tabla = "ssitua"
                codCampo = "codsitua"
                nomCampo = "nomsitua"
                TipCampo = "N"
                Formato = "00"
                Titulo = "Situación Especial"
            
            Case 58 'Listado PROVEEDORES
                EsNomCod = True
                tabla = "proveedor"
                codCampo = "codprove"
                nomCampo = "nomprove"
                TipCampo = "N"
                Formato = "000000"
                Titulo = "Proveedor"
            
            Case 61 'Listado MOTIVOS Pend. Rep.
                EsNomCod = True
                tabla = "smotre"
                codCampo = "codmotre"
                nomCampo = "nommotre"
                TipCampo = "N"
                Formato = "00"
                Titulo = "Motivos Pend. Rep."
        End Select
        
    ElseIf Index = 3 Or Index = 4 Then
         '7: Informe Traspaso Almacenes
         '8: Informe Movimientos Almacen
         txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
    Else
        Select Case Index
        Case 0, 86, 87
            If txtCodigo(Index).Text <> "" Then
                
                If Index = 0 Then
                    txtNombre(Index).Text = PonerNombreCuenta(txtCodigo(Index), 2)
                    If txtNombre(Index).Text = "" Then
                        MsgBox "Número de Cuenta contable no existe en la contabilidad. Reintroduzca.", vbExclamation
                    End If
                Else
                    PonerFormatoEntero txtCodigo(Index)
                    If (Index = 86 Or Index = 87) Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
                End If
            End If
            
        Case 5, 6, 14, 15, 24, 25, 30, 70, 71, 90, 91, 92, 93  'Cod. ARTICULO
            EsNomCod = True
            tabla = "sartic"
            codCampo = "codartic"
            nomCampo = "nomartic"
            TipCampo = "T"
            Titulo = "Artículo"
            txtCodigo(Index).Text = UCase(txtCodigo(Index).Text)
        
        Case 7, 16, 25, 26, 27, 62, 63, 75, 76, 88, 89, 94, 95 'Cod. FAMILIA
            EsNomCod = True
            tabla = "sfamia"
            codCampo = "codfamia"
            nomCampo = "nomfamia"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Familia"
        
        Case 9, 10, 20, 31, 32, 43, 44, 53, 54, 82, 83, 109, 110, 115, 116, 119, 120   'FECHA Desde Hasta
            If txtCodigo(Index).Text <> "" Then
                If Index = 22 And Opcionlistado = 19 Then 'Este campo sera Hora y no Fecha
                    PonerFormatoHora txtCodigo(Index)
                Else
                    PonerFormatoFecha txtCodigo(Index)
                    If Opcionlistado = 223 And txtCodigo(Index).Text <> "" Then
                        'Contabilizar facturas
                        If Not ComprobarFechasConta(Index) Then
                            PonerFoco txtCodigo(Index)
'                        Else '++monica
'                            If OptClientes.Value Then
'                                PonerFoco txtCodigo(0)
'                            Else
'                                cmdCancel(7).SetFocus
'                            End If
                        End If '++
                    End If
                    
                End If
            End If
        
        Case 11, 12, 13, 72 'ALMACENES Propios
            EsNomCod = True
            tabla = "salmpr"
            codCampo = "codalmac"
            nomCampo = "nomalmac"
            TipCampo = "N"
            Formato = "000"
            Titulo = "Almacen Propio"
            
        Case 18, 19, 28, 29, 66, 67, 79, 80 'PROVEEDOR
            EsNomCod = True
            tabla = "proveedor"
            codCampo = "codprove"
            nomCampo = "nomprove"
            TipCampo = "N"
            Formato = "000000"
            Titulo = "Proveedor"
        
        Case 21, 96, 97, 111 'Cod. Operario/Trabajador
            EsNomCod = True
            tabla = "straba"
            codCampo = "codtraba"
            nomCampo = "nomtraba"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Trabajador"
        
        Case 23, 107
            EsNomCod = True
            TipCampo = "N"
            If Opcionlistado = 30 Then 'Precios Especiales
                tabla = "sclien"
                codCampo = "codclien"
                nomCampo = "nomclien"
                Formato = "000000"
                Titulo = "Cliente"
            Else   'Tarifas Precios
                tabla = "starif"
                codCampo = "codlista"
                nomCampo = "nomlista"
                Formato = "000"
                Titulo = "Tarifa de Venta"
            End If
        
        Case 27, 28, 64, 65, 77, 78 'MARCAS
            EsNomCod = True
            tabla = "smarca"
            codCampo = "codmarca"
            nomCampo = "nommarca"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Marca"
        
        Case 31 'Nº de Oferta
            If txtCodigo(Index).Text = "" Then Exit Sub
            codCampo = DevuelveDesdeBDNew(cAgro, "scapre", "numofert", "numofert", txtCodigo(Index).Text, "N")
            If codCampo = "" Then
                MsgBox "No existe el código de Oferta: " & NumCod, vbInformation
                PonerFoco txtCodigo(Index)
            Else
                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
            End If
            
        Case 32, 43 'Carta de la Oferta
            EsNomCod = True
            tabla = "scartas"
            codCampo = "codcarta"
            nomCampo = "descarta"
            TipCampo = "N"
            Formato = "000"
            Titulo = "Cartas para Ofertas"
            
        Case 37, 38, 34, 47, 48, 55, 56, 73, 74, 98, 101, 102, 103, 117, 118 'Cod. CLIENTE
            EsNomCod = True
            tabla = "sclien"
            codCampo = "codclien"
            nomCampo = "nomclien"
            TipCampo = "N"
            Formato = "000000"
            Titulo = "Cliente"
            
        Case 112, 113, 114
            EsNomCod = True
            tabla = "inciden"
            codCampo = "codincid"
            nomCampo = "nomincid"
            TipCampo = "T"
            'Formato = "0000"
            Titulo = "Incidencias"
        
        Case 41, 42, 59, 60 'Nº Contrato
'            If txtCodigo(Index).Text <> "" Then
'                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
'            End If

        Case 45, 46, 106, 108 'ZONAS del Cliente
            EsNomCod = True
            tabla = "szonas"
            codCampo = "codzonas"
            nomCampo = "nomzonas"
            TipCampo = "N"
            Formato = "000"
            Titulo = "Zonas de Clientes"
        
        Case 49, 50 'Cod. AGENTE
            EsNomCod = True
            tabla = "sagent"
            codCampo = "codagent"
            nomCampo = "nomagent"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Agente"
            
        Case 51, 52, 57, 58, 104, 105 'Tipos Contratos/MAntenimientos
            EsNomCod = True
            tabla = "stipco"
            codCampo = "codtipco"
            nomCampo = "nomtipco"
            TipCampo = "T"
            Titulo = "Tipos de Contratos"
            
        Case 61 'Año Ejercicio
            If txtCodigo(Index).Text = "" Then Exit Sub
            If Not IsNumeric(txtCodigo(Index).Text) Then
                MsgBox "El Ejercicio debe ser un Año", vbInformation
                Exit Sub
            End If
        
        Case 68, 69 'Tipos de Articulos
            EsNomCod = True
            tabla = "stipar"
            codCampo = "codtipar"
            nomCampo = "nomtipar"
            TipCampo = "T"
            Titulo = "Tipo de Articulo"
            
        Case 84, 85 'RUTAS del cliente
            EsNomCod = True
            tabla = "srutas"
            codCampo = "codrutas"
            nomCampo = "nomrutas"
            TipCampo = "N"
            Formato = "000"
            Titulo = "Ruta de Asistencia"
        End Select
    End If
    
    If EsNomCod Then
        If TipCampo = "N" Then
            If PonerFormatoEntero(txtCodigo(Index)) Then
                
                
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), tabla, nomCampo, codCampo)
'                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), Tabla, NomCampo, codCampo, Titulo, TipCampo)
            
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, Formato)
            Else
                txtNombre(Index).Text = ""
            End If
        Else
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), tabla, nomCampo, codCampo)
'            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), Tabla, NomCampo, codCampo, Titulo, TipCampo)
        End If
    End If
    
   
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim tabla As String
Dim Titulo As String

    'Llamamos a al form
    Cad = ""
    Conexion = cAgro    'Conexión a BD: Ariges
    Select Case Opcionlistado
        Case 7 'Traspaso de Almacenes
            Cad = Cad & "Nº Trasp|scatra.codtrasp|N|0000000|40·Almacen Origen|scatra.almaorig|N|000|20·Almacen Destino|scatra.almadest|N|000|20·Fecha|scatra.fechatra|F||20·"
            tabla = "scatra"
            Titulo = "Traspaso Almacenes"
        Case 8 'Movimientos de Almacen
            Cad = Cad & "Nº Movim.|scamov.codmovim|N|0000000|40·Almacen|scamov.codalmac|N|000|30·Fecha|scamov|fecmovim|F||30·"
            tabla = "scamov"
            Titulo = "Movimientos Almacen"
        Case 9, 12, 13, 14, 15, 16, 17 '9: Movimientos Articulos
                   '12: Inventario Articulos
                   '14:Actualizar Diferencias de Stock Inventariado
                   '16: Listado Valoracion stock inventariado
            Cad = Cad & "Código|sartic.codartic|T||30·Denominacion|sartic.nomartic|T||70·"
            tabla = "sartic"
            Titulo = "Articulos"
    End Select
          
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
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


Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda

    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        Select Case Opcionlistado
            Case 7, 8 'Informe Traspasos Almacen
                txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaDevuelta, 1), "0000000")
                PonerFoco txtCodigo(indCodigo)
            Case 9, 12, 13, 14, 15, 16, 17 '9: Informe Movimiento Articulos
                                'Inventario Articulos
                                '14: Actualizar diferencias Stock Inventariado
                                '16: Listado Valoracion stock inventariado
                txtCodigo(indCodigo).Text = RecuperaValor(CadenaDevuelta, 1)
                txtNombre(indCodigo).Text = RecuperaValor(CadenaDevuelta, 2)
                PonerFoco txtCodigo(indCodigo)
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerFrameListadoVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para los listados de los mantenimientos de tabla: smarca, stipar,...

    H = 6675
    W = 7995
    PonerFrameVisible Me.frameListado, visible, H, W

    If visible = True Then
        Me.Optcodigo.Value = True
    End If
End Sub


Private Sub PonerFrameInventarioVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Inventario Visible y Ajustado al Formulario, y visualiza los controles
'necesario para cada Informe de Inventario
Dim VerOpcion As Boolean
       
    If visible = True Then
        H = 6675
        W = 7995
        VerOpcion = (Opcionlistado = 15) Or (Opcionlistado = 16) Or (Opcionlistado = 17) Or (Opcionlistado = 19)
        
        If VerOpcion Then
            H = 6900
            Me.cmdAceptar(4).Top = 6360
            Me.cmdCancel(4).Top = 6360
        ElseIf Opcionlistado = 13 Then
            H = 6000
            Me.cmdAceptar(4).Top = 5200
            Me.cmdCancel(4).Top = Me.cmdAceptar(4).Top
        End If
        PonerFrameVisible Me.FrameInventario, visible, H, W

                
        '======================================
        'Valorar con Precios
        If VerOpcion Then
            Me.FrameValorar.visible = VerOpcion
            Me.FrameValorar.Left = 300
            If Opcionlistado = 17 Then
                Me.FrameValorar.Top = 4500
            Else
                Me.FrameValorar.Top = 5000
            End If
            Me.chkSinStock.visible = VerOpcion
        End If
        '====================================
        'Poner el Trabajador
        VerOpcion = (Opcionlistado = 14)
'--monica he quitao el trabajador
'        Me.Label4(7).visible = VerOpcion
'        Me.imgBuscarG(17).visible = VerOpcion
'        Me.txtCodigo(21).visible = VerOpcion
'        Me.txtNombre(21).visible = VerOpcion
'        If VerOpcion Then txtCodigo(21).TabIndex = 47
        
        
        '======================================
        'Fecha Listados
        If Opcionlistado = 15 Then '15: Listado Articulos Inactivos
            Me.Label4(5).Caption = "Fecha Inactividad"
        ElseIf Opcionlistado = 19 Then
            Me.Label4(5).Caption = "Fecha Stock"
        Else
            Me.Label4(5).Caption = "Fecha Inventario"
        End If
        VerOpcion = (Opcionlistado = 12) Or (Opcionlistado = 15) Or (Opcionlistado = 16) Or (Opcionlistado = 19)
        Me.Label4(5).visible = VerOpcion  'campo fecha
        Me.imgFecha(2).visible = VerOpcion
        Me.txtCodigo(20).visible = VerOpcion
        'campo HAsta Fecha
        Me.Label4(8).visible = (Opcionlistado = 16)
        'Si opcionlistado=19 este campo sera la hora
        Me.Label4(9).visible = (Opcionlistado = 16) Or (Opcionlistado = 19)
        If Opcionlistado = 19 Then
            Me.Label4(9).Caption = "Hora"
            Me.Label4(9).Left = 4150
            Me.txtCodigo(22).Left = 4700
        End If
        Me.imgFecha(3).visible = (Opcionlistado = 16)
        Me.txtCodigo(22).visible = (Opcionlistado = 16) Or (Opcionlistado = 19)
        If Opcionlistado = 16 Then
            Me.Label4(8).Left = 2280
            Me.imgFecha(2).Left = 2820
            Me.txtCodigo(20).Left = 3120
            Me.Label4(9).Left = 4680
            Me.imgFecha(3).Left = 5160
            Me.txtCodigo(22).Left = 5430
'            txtCodigo(22).TabIndex = 48
        End If
        
        
        '====================================
        'Activar o no los check de Opcion:
        VerOpcion = (Opcionlistado = 12) Or (Opcionlistado = 13) Or (Opcionlistado = 16) Or (Opcionlistado = 17) Or (Opcionlistado = 19) Or Opcionlistado = 15
                    '12: Toma de Inventario
                    '13: Listado Diferencias stock
        
        Me.FrameOpciones.visible = VerOpcion
        Me.FrameOpciones.Top = 5000
        If Opcionlistado = 13 Then
            Me.FrameOpciones.Top = 5500
            Me.FrameOpciones.BorderStyle = 0
            Me.FrameOpciones.Height = 1000
            '15/06/2009
            Me.FrameValorar.visible = True
            Me.FrameValorar.Top = 4450
            '15/06/2009
        End If
        Me.FrameOpciones.Height = 1000

        Me.chkSaltaPag.visible = VerOpcion
        Me.chkValorado.visible = (Opcionlistado = 16) Or (Opcionlistado = 17)

        
        VerOpcion = (Opcionlistado = 12)
        If VerOpcion Or Opcionlistado = 13 Then Me.FrameOpciones.Left = 700
        '15/06/2009
        If Opcionlistado = 13 Then
            Me.FrameOpciones.Left = 4230
            Me.FrameOpciones.Top = 4250
            Me.FrameOpciones.Height = 900
        End If
        '15/06/2009
        Me.chkImprimeStock.visible = VerOpcion
        Me.chkImprimeStock.Top = 600
        If VerOpcion Then Me.txtCodigo(20).Text = Date
    End If
End Sub




Private Sub PonerFrameRepxDiaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para los listados de las Reparaciones x dia, de tabla: scarep

    H = 4085
    W = 6000
    If Me.OptClientes.Value Then W = 6795
    PonerFrameVisible Me.FrameRepxDia, visible, H, W

    If visible = True Then
        Me.Caption = "AriAgro"
'--monica:siempre son facturas de proveedor
'        Me.FrameContab.visible = (OpcionListado = 223 Or OpcionListado = 224)
        Me.FrameProgress.visible = False
        If Opcionlistado <> 223 And Opcionlistado <> 224 Then
            Me.cmdAceptarRepxDia.Top = 2800
            Me.cmdCancel(7).Top = 2800
        End If
        Select Case Opcionlistado
            Case 63
                Me.lblTitulo(0).Caption = "Reparaciones por Día"
                Me.Label2(2).Caption = "Fecha Reparación:"
                Frame2.Top = 1350
            Case 73
                Me.lblTitulo(0).Caption = "Altas de Mantenimientos"
                Me.Label2(2).Caption = "Fecha Mantenimiento:"
                Frame2.Top = 1350
            Case 223, 224 'Pedir datos para contabilizar facturas
                Me.lblTitulo(0).Caption = "Contabilizar Facturas"
                Me.Label2(2).Caption = "Fecha Factura:"
                '++monica: datos de contabilizacion fact.venta para tesoreria
                Frame1.visible = (Me.OptClientes.Value = True)
                Frame1.Enabled = (Me.OptClientes.Value = True)
                If Me.OptClientes.Value Then
                    Frame2.Top = 1000
                Else
                '++
                    Frame2.Top = 1680
                End If
                If Opcionlistado = 224 Then
                    Me.OptProve.Value = True
                    Opcionlistado = 223
                End If
        End Select
    End If
End Sub



Private Sub ponerFrameArticulosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el informe de Articulos, de tabla: sartic
Dim b As Boolean



    'Hay una opcion mas que mostrara este frame. la 247. Correccion de de tarigfas e importes en articulos
    FrameTapaINCORRECTO.visible = False
    chkMinimoCorreg.visible = False
    b = (Opcionlistado = 6)
    If b Then
        Me.Label9.Caption = "Informe de Articulos"
       
        W = 8595
    Else
        If Opcionlistado = 18 Then
            Me.Label9.Caption = "Informe Stocks Maximos y Minimos"
            Label4(36).Caption = "Almacén"
        Else
            'NUEVA OCPION:  247
            'Corregir tarifas y eso
            chkMinimoCorreg.visible = True
            Me.Label9.Caption = "Verificación tarifas y P.V.P."
            FrameTapaINCORRECTO.visible = True
            Label4(36).Caption = "Tarifa"
            cmbDecimales.ListIndex = 0
        End If
        W = 7395
       
    End If
    H = 6820
    
    
    PonerFrameVisible Me.FrameInfArticulos, visible, H, W
    If visible = True Then
        'visible orden campos si opcionlistado=6
        Me.FrameOrden.visible = b
        Label4(36).visible = Not b

        Me.imgBuscarG(18).visible = Not b
        Me.txtCodigo(72).visible = Not b
        Me.txtNombre(72).visible = Not b
        
        'Visible Frame stocks Max Minimos si opcionlistado=18
        Me.optStockMax.Value = True
        Me.FrameStockMaxMin.visible = Opcionlistado = 18
  
    
    
        'REajustes.
        'El articulo NO se muestra si la opcion es 247
        b = Opcionlistado <> 247
        Label4(38).visible = b
        Label3(51).visible = b
        imgBuscarG(27).visible = b
        txtCodigo(70).visible = b
        txtNombre(70).visible = b
        Label3(54).visible = b
        imgBuscarG(28).visible = b
        txtCodigo(71).visible = b
        txtNombre(71).visible = b
        
        Label4(75).visible = Not b
        cmbDecimales.visible = Not b
    End If
End Sub


Private Sub CargarListView()
'Carga el List View del frame: frameMovimArtic
'con los parametros de la tabla: stipom (Tipos de Movimientos)
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem

    'Los encabezados
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Código", 800
    ListView1.ColumnHeaders.Add , , "Descripción", 2250
    
    Sql = "select * from usuarios.stipom stipom where muevesto = 1 "
    If DeServicios Then Sql = Sql & " and codtipom in ('SES','SEC','RES','REC')"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Set ItmX = ListView1.ListItems.Add
        ItmX.Text = Rs.Fields(0).Value
        ItmX.Checked = True
        ItmX.SubItems(1) = Rs.Fields(1).Value
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
End Sub



Private Sub CargarListViewOrden()
'Carga el List View del frame: frameInfArticulos
'para establecer el orden en que se van a mostrar los datos en el Informe
'Orden: Familia, MArca, Proveedor, Tipo de Articulo, Articulo
Dim ItmX As ListItem

    'Los encabezados
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "Campo", 1600
    Set ItmX = ListView2.ListItems.Add
    ItmX.Text = "Familia"
    Set ItmX = ListView2.ListItems.Add
    ItmX.Text = "Marca"
    Set ItmX = ListView2.ListItems.Add
    ItmX.Text = "Proveedor"
    Set ItmX = ListView2.ListItems.Add
    ItmX.Text = "Tipo Articulo"
End Sub


Private Function PonerFormulaYParametrosInf9() As Boolean
Dim Cad As String
Dim todosMarcados As Boolean
Dim devuelve As String
Dim i As Byte

    PonerFormulaYParametrosInf9 = False
    InicializarVbles
    
    'Parametro EMPRESA
    cadParam = "|pNomEmpre=""" & vParam.NombreEmpresa & """|"
    numParam = 1
        
    'Cadena para seleccion Desde y Hasta ARTICULO
    If txtCodigo(5).Text <> "" Or txtCodigo(6).Text <> "" Then
        Codigo = "{smoval.codartic}"
        devuelve = "pDHArticulo=""Artículo: "
        If Not PonerDesdeHasta(Codigo, "T", 5, 6, devuelve) Then Exit Function
    End If
                    
    'Cadena para seleccion Desde y Hasta FAMILIA
    If txtCodigo(7).Text <> "" Or txtCodigo(8).Text <> "" Then
        Codigo = "{sartic.codfamia}"
        devuelve = "pDHFamilia=""Familia: "
        If Not PonerDesdeHasta(Codigo, "N", 7, 8, devuelve) Then Exit Function
    End If
        
    'Cadena para seleccion Desde y Hasta ALMACEN
    If txtCodigo(11).Text <> "" Or txtCodigo(12).Text <> "" Then
        Codigo = "{smoval.codalmac}"
        devuelve = "pDHAlmacen=""Almacen: "
        If Not PonerDesdeHasta(Codigo, "N", 11, 12, devuelve) Then Exit Function
    End If
    
    
    'Cadena para seleccion Desde y Hasta CLIENTE/PROVEEDOR/SOCIO
    If txtCodigo(86).Text <> "" Or txtCodigo(87).Text <> "" Then
        Codigo = "{smoval.codigope}"
        devuelve = "pDHOperario=""Cliente/Proveedor/Socio: "
        If Not PonerDesdeHasta(Codigo, "N", 86, 87, devuelve) Then Exit Function
    End If
    
        
'    cadSelect = QuitarCaracterACadena(cadFormula, "{")
'    cadSelect = QuitarCaracterACadena(cadSelect, "}")
        
    '=================================================
    'Cadena para seleccion Desde y Hasta FECHA
    If txtCodigo(9).Text <> "" Or txtCodigo(10).Text <> "" Then
        Codigo = "{smoval.fechamov}"
        devuelve = "pDHFecha=""Fecha: "
        If Not PonerDesdeHasta(Codigo, "F", 9, 10, devuelve) Then Exit Function
    End If
        
        
    'Cadena de Seleccion TIPOS de MOVIMIENTOS
    Codigo = "{smoval.detamovi}"
    devuelve = ""
    
    '[Monica]11/10/2011: vemos que criterios de busqueda ha introducido
    If DeServicios Then
        Cad = ""
        devuelve = ""
        For i = 1 To Me.ListView1.ListItems.Count
            If Me.ListView1.ListItems(i).Checked Then
                If Cad = "" Then
                    Cad = Me.ListView1.ListItems(i).Text
                Else
                    Cad = Cad & ", " & Me.ListView1.ListItems(i).Text
                End If
                If devuelve = "" Then
                    devuelve = Codigo & " = """ & Me.ListView1.ListItems(i).Text & """"
                Else
                    devuelve = devuelve & " or " & Codigo & " = """ & Me.ListView1.ListItems(i).Text & """"
                End If
            End If
        Next i
        
        If devuelve <> "" Then 'Hay algun movimiento marcado
            If cadFormula <> "" Then
                cadFormula = cadFormula & " AND " & "(" & devuelve & ")"
'                devuelve = QuitarCaracterACadena(devuelve, "{")
'                devuelve = QuitarCaracterACadena(devuelve, "}")
                cadselect = cadselect & " AND " & "(" & devuelve & ")"
                cadParam = cadParam
            Else
                cadFormula = "(" & devuelve & ")"
'                devuelve = QuitarCaracterACadena(devuelve, "{")
'                devuelve = QuitarCaracterACadena(devuelve, "}")
                cadselect = "(" & devuelve & ")"
            End If
            Cad = "pTiposMov=""Tipos Movimiento: " & Cad
            cadParam = cadParam & Cad & """|"
            numParam = numParam + 1
        Else 'Todos desmarcados
            cadFormula = ""
        End If
    Else
        'Si todos seleccionados no añadir la select
        todosMarcados = True
        i = 1
        While Not i > Me.ListView1.ListItems.Count And todosMarcados
            If Not Me.ListView1.ListItems(i).Checked Then todosMarcados = False
            i = i + 1
        Wend
        SoloServicios = False
        If Not todosMarcados Then
            Cad = ""
            devuelve = ""
            For i = 1 To Me.ListView1.ListItems.Count
                If Me.ListView1.ListItems(i).Checked Then
                    If Cad = "" Then
                        Cad = Me.ListView1.ListItems(i).Text
                    Else
                        Cad = Cad & ", " & Me.ListView1.ListItems(i).Text
                    End If
                    If devuelve = "" Then
                        devuelve = Codigo & " = """ & Me.ListView1.ListItems(i).Text & """"
                    Else
                        devuelve = devuelve & " or " & Codigo & " = """ & Me.ListView1.ListItems(i).Text & """"
                    End If
                End If
            Next i
    
            If devuelve <> "" Then 'Hay algun movimiento marcado
                If cadFormula <> "" Then
                    cadFormula = cadFormula & " AND " & "(" & devuelve & ")"
    '                devuelve = QuitarCaracterACadena(devuelve, "{")
    '                devuelve = QuitarCaracterACadena(devuelve, "}")
                    cadselect = cadselect & " AND " & "(" & devuelve & ")"
                    cadParam = cadParam
                Else
                    cadFormula = "(" & devuelve & ")"
    '                devuelve = QuitarCaracterACadena(devuelve, "{")
    '                devuelve = QuitarCaracterACadena(devuelve, "}")
                    cadselect = "(" & devuelve & ")"
                End If
                Cad = "pTiposMov=""Tipos Movimiento: " & Cad
                cadParam = cadParam & Cad & """|"
                numParam = numParam + 1
            Else 'Todos desmarcados
                Cad = ""
                For i = 1 To ListView1.ListItems.Count
                    If Cad = "" Then
                        Cad = """" & ListView1.ListItems(i).Text & """"
                    Else
                        Cad = Cad & ", """ & ListView1.ListItems(i).Text & """"
                    End If
                Next i
                devuelve = Codigo & " NOT IN [" & Cad & "]"
                Cad = Codigo & " NOT IN (" & Cad & ")"
                Cad = QuitarCaracterACadena(Cad, "{")
                Cad = QuitarCaracterACadena(Cad, "}")
                If cadFormula = "" Then
                    cadFormula = "(" & devuelve & ")"
                    cadselect = "(" & Cad & ")"
                Else
                    cadFormula = cadFormula & " AND " & "(" & devuelve & ")"
                    cadselect = cadselect & " AND " & "(" & Cad & ")"
                End If
            End If
        End If
    End If
    
    If cadFormula = "" Then
        MsgBox "Introduzca algún criterio de selección para el Informe.", vbInformation
        Exit Function
    End If
    PonerFormulaYParametrosInf9 = True
    
End Function


Private Function PonerFormulaYParametrosInf12() As Boolean
Dim Cad As String, cadFrom As String
Dim devuelve As String
Dim ImprStock As String
Dim CodAux As String
Dim strValorado As String
Dim strSinStock As String
Dim bytPrecio As Byte

'    InicializarVbles
    cadParam = "|pNomEmpre=""" & vParam.NombreEmpresa & """|"
    numParam = numParam + 1
    cadFrom = ""
    devuelve = ""
    PonerFormulaYParametrosInf12 = False
    
    
    '15/06/2009
    If Opcionlistado = 13 Then
        If optPrecioUC.Value Then
           cadParam = cadParam & "pPrecio=0|"
        Else
           cadParam = cadParam & "pPrecio=1|"
        End If
        numParam = numParam + 1
    End If
    '15/06/2009
    
    
    '===================================================
    '================= FORMULA =========================
    
    Select Case Opcionlistado
        Case 12, 15, 16, 17, 19
            CodAux = "{salmac."
            cadFrom = "  salmac "
'        Case 15 'Listado articulos inactivos
'            CodAux = "{salmac."
'            cadFrom = "  (salmac LEFT OUTER JOIN smoval ON salmac.codartic=smoval.codartic AND salmac.codalmac=smoval.codalmac) "
'            cadFrom = "salmac"
        Case 13, 14
            CodAux = "{sinven."
            cadFrom = " sinven "
    End Select
    
    'Cadena para seleccion De ALMACEN
    '-----------------------------------
    Codigo = CodAux & "codalmac}"
    If Trim(txtCodigo(13).Text) <> "" Then _
    devuelve = Codigo & " = " & Val(txtCodigo(13).Text)
    If devuelve <> "" Then
        cadFormula = devuelve
        Cad = "pAlmacen= ""Almacen: " & Format(txtCodigo(13).Text, "000") & " " & txtNombre(13).Text
        cadParam = cadParam & Cad & """|"
        numParam = numParam + 1
    End If
            
    'Cadena para seleccion Desde y Hasta ARTICULOS
    '----------------------------------------------
    If txtCodigo(14).Text <> "" Or txtCodigo(15).Text <> "" Then
        Codigo = CodAux & "codartic}"
        devuelve = "pDHArticulo=""Articulo: "
        If Not PonerDesdeHasta(Codigo, "T", 14, 15, devuelve) Then Exit Function
    End If
    
    'Cadena para seleccion Desde y Hasta FAMILIA
    '--------------------------------------------
    If txtCodigo(16).Text <> "" Or txtCodigo(17).Text <> "" Then
        Select Case Opcionlistado
            Case 12, 15, 16, 17, 19: Codigo = "{sartic.codfamia}"
            Case Else: Codigo = "{sinven.codfamia}"
        End Select
        Cad = "pDHFamilia=""Familia: "
        If Not PonerDesdeHasta(Codigo, "N", 16, 17, Cad) Then Exit Function
        cadFrom = cadFrom & " INNER JOIN sartic ON " & CodAux & "codartic=sartic.codartic "
    End If
    
    'Cadena para seleccion Desde y Hasta PROVEEDOR
    '----------------------------------------------
    If txtCodigo(18).Text <> "" Or txtCodigo(19).Text <> "" Then
        Select Case Opcionlistado
            Case 12, 15, 16, 17, 19: Codigo = "{sartic.codprove}"
            Case Else: Codigo = "{sinven.codprove}"
        End Select
        Cad = "pDHProveedor=""Proveedor: "
        If Not PonerDesdeHasta(Codigo, "N", 18, 19, Cad) Then Exit Function
    End If
    'Select para MySQL
    cadselect = QuitarCaracterACadena(cadFormula, "{")
    cadselect = QuitarCaracterACadena(cadselect, "}")
    cadselect = QuitarCaracterACadena(cadselect, "_1")
    cadFrom = QuitarCaracterACadena(cadFrom, "{")
    
    'Cadena para seleccion Desde y Hasta FECHA
    '----------------------------------------------
    If (Opcionlistado = 16) Then
        If txtCodigo(20).Text <> "" Or txtCodigo(22).Text <> "" Then
            'codigo = "{salmac.codartic}"
            Codigo = CodAux & "fechainv}"
            devuelve = CadenaDesdeHasta(txtCodigo(20).Text, txtCodigo(22).Text, Codigo, "F")
    
            If devuelve = "Error" Then Exit Function
            
            If Not AnyadirAFormula(cadFormula, devuelve) Then
                Exit Function
            ElseIf devuelve <> "" Then
                Cad = "pDHFecha=""Fecha: "
                If txtCodigo(20).Text <> "" Then _
                    Cad = Cad & "desde " & txtCodigo(20).Text
                If txtCodigo(22).Text <> "" Then _
                    Cad = Cad & "  hasta " & txtCodigo(22).Text
                cadParam = cadParam & Cad & """|"
                numParam = numParam + 1
                'Para Comprobar si hay registros a Mostrar antes de abrir el Informe
                devuelve = "salmac.fechainv"
                devuelve = CadenaDesdeHastaBD(txtCodigo(20).Text, txtCodigo(22).Text, devuelve, "F")
                AnyadirAFormula cadselect, devuelve
            Else
                'Si no hay fecha de inventario seleccionada coger solo
                'los articulos de los que se haya hecho inventario alguna vez
                devuelve = "not isnull({salmac.fechainv})"
                If Not AnyadirAFormula(cadFormula, devuelve) Then
                    Exit Function
                End If
                devuelve = "not isnull(salmac.fechainv)"
                AnyadirAFormula cadselect, devuelve
            End If
        End If
    End If
    
    'Cadena de seleccion de FECHA de Inactividad
    '------------------------------------------------
    If Opcionlistado = 15 Then '15: Listado de Articulos Inactivos
         If txtCodigo(20).Text <> "" Then _
            Cad = "pFechaInve=""" & txtCodigo(20).Text & """"
        
        'Poner en el parametro pListaArt la lista de Articulos que no tiene
        'un registro de movimiento en la smoval con fecha posterior a la
        'fecha de inactividad
        strValorado = ListaArtActivos(cadselect, txtCodigo(20).Text)
        Cad = "pListaArtic=""" & strValorado & """|"
        cadParam = cadParam & Cad
        numParam = numParam + 1
        
        'Añadir a la formula de seleccion que no sea uno de la lista
        devuelve = " not (" & CodAux & "codartic} in {@pListaArtic})"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
        
        strValorado = QuitarCaracterACadena(strValorado, "[")
        strValorado = QuitarCaracterACadena(strValorado, "]")
        devuelve = " not (salmac.codartic in (" & strValorado & "))"
        AnyadirAFormula cadselect, devuelve
    End If
    
    'Cadena de seleccion de FECHA de Stocks a una Fecha
    '--------------------------------------------------
     If Opcionlistado = 19 Then
        If txtCodigo(20).Text <> "" Then
            Cad = txtCodigo(20).Text
            'Hora
            If txtCodigo(22).Text <> "" Then _
                Cad = Cad & "  " & txtCodigo(22).Text
            cadParam = cadParam & "pFechaStock=""" & Cad & """|"
            numParam = numParam + 1
        End If
     End If
     
    'Cadena para Seleccion de Articulos con Stock<>0
    '------------------------------------------------
    If Opcionlistado = 16 Or Opcionlistado = 17 Or Opcionlistado = 15 Then
        If Me.chkSinStock.Value = 0 Then
            If Opcionlistado = 16 Then
                devuelve = "{salmac.stockinv}<>0"
            Else
                devuelve = CodAux & "canstock}<>0"
            End If
            If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
            
            devuelve = QuitarCaracterACadena(devuelve, "{")
            devuelve = QuitarCaracterACadena(devuelve, "}")
            devuelve = QuitarCaracterACadena(devuelve, "_1")
            AnyadirAFormula cadselect, devuelve
        End If
    ElseIf Opcionlistado = 19 Then
         If Me.chkSinStock.Value Then
            ImprStock = "True"
        Else
            ImprStock = "False"
        End If
        cadParam = cadParam & "pSinStock=" & ImprStock & "|"
        numParam = numParam + 1
    End If
    
       
    '============================================
    '============= PARAMETROS ===================
    If Opcionlistado = 12 Or Opcionlistado = 15 Then
        '12: Toma de Inventario
        '15: Listado Articulos Inactivos
        cadParam = cadParam & "pFechaInve=""" & txtCodigo(20).Text & """|"
        numParam = numParam + 1
    End If
    
    If Opcionlistado = 12 Then
        'Parámetro Imprime Stock (Si/No)
        If Me.chkImprimeStock.Value Then
            ImprStock = "True"
        Else
            ImprStock = "False"
        End If
        cadParam = cadParam & "pImprimeStock=" & ImprStock & "|"
        numParam = numParam + 1
        
        'seleccionar para inventariar los articulos que no tienen control stock
        devuelve = " {sartic.ctrstock} = 1 "
        AnyadirAFormula cadFormula, devuelve
        AnyadirAFormula cadselect, devuelve
        'Laura 03/01/07
        If InStr(cadFrom, "sartic") = 0 Then '03-06-2009 monica:antes not instr(cafrom,"sartic")
            cadFrom = cadFrom & " INNER JOIN sartic ON " & CodAux & "codartic=sartic.codartic "
        End If
    End If
    
    If Opcionlistado = 12 Or Opcionlistado = 13 Or Opcionlistado = 15 Or Opcionlistado = 16 Or Opcionlistado = 17 Or Opcionlistado = 19 Then
        'Parámetro Salta Pag. en Familia (Si/No)
        If Me.chkSaltaPag.Value Then
            ImprStock = "True"
        Else
            ImprStock = "False"
        End If
        cadParam = cadParam & "pSaltaFamilia=" & ImprStock & "|"
        numParam = numParam + 1
    End If
    
    If Opcionlistado = 16 Or Opcionlistado = 17 Then '16: Valoración de Stocks Inventariados
                                                     '17: Valoración Stocks
        'Parámetro Valorado
        If Me.chkValorado.Value Then
            strValorado = "True"
        Else
            strValorado = "False"
        End If
        cadParam = cadParam & "pValorado=" & strValorado & "|"
        numParam = numParam + 1
    End If
    
    If (Opcionlistado = 15) Or (Opcionlistado = 16) Or (Opcionlistado = 17) Or (Opcionlistado = 19) Then
        'Parametro Precio de Valoracion, elegir un Precio para realizar la valoracion: canstock * precio
        If Me.optPrecioMP.Value Then bytPrecio = 1
'        If Me.optPrecioMA.Value Then bytPrecio = 2
        If Me.optPrecioUC.Value Then bytPrecio = 3
'        If Me.optPrecioStd.Value Then bytPrecio = 4
        cadParam = cadParam & "pPrecio=" & bytPrecio & "|"
        numParam = numParam + 1
    End If
    '=====================================================================
    
       
    'comprobar si hay registros para mostrar en el Informe antes de Abrirlo
    If Not HayRegParaInforme(cadFrom, cadselect) Then Exit Function
    
    If Opcionlistado = 19 Then
        cadselect = "Select count(*) FROM " & cadFrom & " WHERE " & cadselect
        cadselect = Replace(cadselect, "count(*)", "*")
        DescargarDatosTMPStockFecha
        If Not CargarTMPStockFecha(cadselect, txtCodigo(20).Text, txtCodigo(22).Text) Then Exit Function
    End If
    
    PonerFormulaYParametrosInf12 = True
End Function


Private Function AnyadirParametroDH(Cad As String, indD As Byte, indH As Byte) As String
On Error Resume Next
    
     If txtCodigo(indD).Text <> "" Then
        Cad = Cad & "desde " & txtCodigo(indD).Text
        If txtNombre(indD).Text <> "" Then Cad = Cad & " - " & txtNombre(indD).Text
    End If
    If txtCodigo(indH).Text <> "" Then
        Cad = Cad & "  hasta " & txtCodigo(indH).Text
        If txtNombre(indH).Text <> "" Then Cad = Cad & " - " & txtNombre(indH).Text
    End If
    
    AnyadirParametroDH = Cad
    If Err.Number <> 0 Then Err.Clear
End Function


Private Function InsertarInventario() As Boolean
'Inserta en la Tabla:sinven los articulos seleccionados para realizar Inventario
'Inserta en la Tabla Hist.: shinve los datos que habia de inventario
'Además Actualiza la Tabla:salmac los campos:fechainv, horainve, statusin
Dim Sql As String, ADonde As String
Dim Rs As ADODB.Recordset
Dim hora As Date

On Error GoTo EInventario:
   
'   If CrearTmpInventario(cadSelect) Then
   

        'Aqui empieza transaccion
        conn.BeginTrans
    
          
    
'        'Insertar en la tabla de Histórico: shinve
'        'Pasamos al Hist. los datos que habia antes de hacer nuevo inventario
'        ADonde = "Insertando datos en Histórico. Tabla: shinve"
'        SQL = "INSERT INTO shinve (codartic,codalmac,fechainv,horainve,existenc) "
'        SQL = SQL & " SELECT salmac.codartic, salmac.codalmac, salmac.fechainv,salmac.horainve,salmac.stockinv "
'        SQL = SQL & " FROM salmac INNER JOIN sartic ON salmac.codartic=sartic.codartic "
'        SQL = SQL & " WHERE " & cadFormula
'        'si no se ha inventariado antes no lo pasamos al historico
'        SQL = SQL & " AND not isnull(salmac.fechainv) "
'        Conn.Execute SQL
'
        
        'Insertar en la tabla de Histórico: shinve
        'Pasamos al Hist. los datos que habia antes de hacer nuevo inventario
        ADonde = "Insertando datos en Histórico. Tabla: shinve"
        Sql = "INSERT INTO shinve (codartic,codalmac,fechainv,horainve,existenc) "
        Sql = Sql & " SELECT codartic,codalmac,fechainv,horainve,stockinv "
'        SQL = SQL & " FROM salmac INNER JOIN sartic ON salmac.codartic=sartic.codartic "
        Sql = Sql & " FROM tmpInven "
'        SQL = SQL & " WHERE " & cadFormula
        'si no se ha inventariado antes no lo pasamos al historico
        Sql = Sql & " WHERE not isnull(fechainv) "
        '--- Laura 03/01/2006
        Sql = Sql & " AND fechainv<>'0000-00-00' AND date(horainve)<>'0000-00-00' "
        '---
        conn.Execute Sql
        
        
        
        
        
        hora = Format(txtCodigo(20).Text & " " & Time, "yyyy-mm-dd hh:mm:ss")
        
'        'Insertamos en la Tabla sinven
'        ADonde = "Insertando datos en Inventario Real. Tabla: sinven"
'        SQL = "INSERT INTO sinven (codartic, codalmac, codfamia, codprove, fechainv, horainve, existenc) "
'        SQL = SQL & "SELECT salmac.codartic, salmac.codalmac, sartic.codfamia, sartic.codprove," & DBSet(txtCodigo(20).Text, "F") & " as fechainv," & DBSet(hora, "FH") & " as horainve, 0 as existenc "
'        SQL = SQL & " FROM salmac INNER JOIN sartic ON salmac.codartic=sartic.codartic "
'        SQL = SQL & " WHERE " & cadFormula
'        'Insertamos los articulos que tiene control de stock
'        SQL = SQL & " AND sartic.ctrstock=1"
'        Conn.Execute SQL

        'Insertamos en la Tabla sinven
        ADonde = "Insertando datos en Inventario Real. Tabla: sinven"
        Sql = "INSERT INTO sinven (codartic, codalmac, codfamia, codprove, fechainv, horainve, existenc) "
        Sql = Sql & "SELECT codartic, codalmac, codfamia, codprove," & DBSet(txtCodigo(20).Text, "F") & " as fechainv," & DBSet(hora, "FH") & " as horainve, 0 as existenc "
'        SQL = SQL & " FROM salmac INNER JOIN sartic ON salmac.codartic=sartic.codartic "
        Sql = Sql & " FROM tmpInven "
'        SQL = SQL & " WHERE " & cadFormula
        'Insertamos los articulos que tiene control de stock
'        SQL = SQL & " AND sartic.ctrstock=1"
        conn.Execute Sql


        
        
'        SQL = "SELECT salmac.codartic, salmac.codalmac, sartic.codfamia, sartic.codprove "
'        SQL = SQL & "FROM salmac INNER JOIN sartic ON salmac.codartic=sartic.codartic "
'        SQL = SQL & " WHERE " & cadFormula
        
        Sql = "SELECT codartic, codalmac, codfamia, codprove "
        Sql = Sql & " FROM tmpInven "
    
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not Rs.EOF
            
        
        
    '        'Insertamos en la Tabla sinven
    '        ADonde = "Insertando datos en Inventario Real. Tabla: sinven"
    '        SQL = "INSERT INTO sinven (codartic, codalmac, codfamia, codprove, fechainv, horainve, existenc) "
    '        SQL = SQL & " VALUES (" & DBSet(Rs.Fields(0).Value, "T") & ", " & Rs.Fields(1).Value & ", "
    '        SQL = SQL & Rs.Fields(2).Value & ", " & Rs.Fields(3).Value & ", '"
    '        'SQL = SQL & Format(txtCodigo(20).Text, "yyyy-mm-dd") & "', '" & hora & "', " & rs.Fields(2).Value & ")"
    '        SQL = SQL & Format(txtCodigo(20).Text, "yyyy-mm-dd") & "', '" & Format(hora, "yyyy-mm-dd hh:mm:ss") & "', 0)"
    '        Conn.Execute SQL
            
            
            
            
            
            
            
            'Actualizamos la tabla salmac ponemos statusin=1 para indicar que se
            'esta realizando inventario y bloquear los articulos para que no se puedan
            'realizar movimientos, traspasos, etc.
            'Actualizamos la Tabla: salmac los campos: fechainv, horainve
            ADonde = "Actualizando datos en Articulos x Almacen"
            Sql = "UPDATE salmac SET fechainv='" & Format(txtCodigo(20).Text, "yyyy-mm-dd") & "', "
            Sql = Sql & " horainve='" & Format(hora, "yyyy-mm-dd hh:mm:ss") & "', " & "statusin=1 , stockinv=0"
            Sql = Sql & " WHERE codartic=" & DBSet(Rs.Fields(0).Value, "T") & " AND "
            Sql = Sql & "codalmac=" & Rs.Fields(1).Value
            conn.Execute Sql
            Rs.MoveNext
        Wend
    
        Rs.Close
        Set Rs = Nothing
'    Else
'        Exit Function
'    End If
    
EInventario:
    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
          Sql = "Insertando Datos de Inventario." & vbCrLf & "--------------------------------------" & vbCrLf
          Sql = Sql & ADonde
          MuestraError Err.Number, Sql, Err.Description
        conn.RollbackTrans
        InsertarInventario = False
    Else
        InsertarInventario = True
        conn.CommitTrans
    End If
End Function


Private Function CrearTmpInventario(cadFormula As String) As Boolean
Dim Sql As String
Dim b As Boolean

    On Error GoTo ECrearInv
    
    b = False
    Sql = "CREATE TEMPORARY TABLE tmpInven ( "
    Sql = Sql & "codartic varchar(16) NOT NULL default '0', "
    Sql = Sql & "codalmac smallint(3) unsigned NOT NULL default '0', "
    Sql = Sql & "codfamia smallint(4) unsigned NOT NULL default '0', "
    Sql = Sql & "codprove int(6) unsigned NOT NULL default '0', "
    Sql = Sql & "fechainv date NOT NULL default '0000-00-00', "
    Sql = Sql & "horainve datetime NOT NULL default '0000-00-00 00:00:00', "
    Sql = Sql & "stockinv decimal(12,2) NOT NULL default '0.00')"
    conn.Execute Sql
    b = True
    
    
    'Seleccionar todos los registros que vamos a inventariar, insertarlos en la TMP
    'y trabajar con estos valores
    Sql = "SELECT salmac.codartic, salmac.codalmac, sartic.codfamia, sartic.codprove,salmac.fechainv,salmac.horainve,salmac.stockinv  "
    Sql = Sql & "FROM salmac INNER JOIN sartic ON salmac.codartic=sartic.codartic "
    Sql = Sql & " WHERE " & cadFormula
    Sql = Sql & " AND sartic.ctrstock=1"

    Sql = " INSERT INTO tmpInven " & Sql
    conn.Execute Sql
    
    
    
ECrearInv:
    If Err.Number <> 0 Then
        Sql = " DROP TABLE IF EXISTS tmpInven;"
        conn.Execute Sql
        b = False
        Err.Clear
    End If
    CrearTmpInventario = b
End Function






Private Function ActualizarInventario() As Boolean
'-----------------------------------------------------------------
'* Modifica en la Tabla: salmac los campos: cansotck, fechainv, horainve,statusin de los articulos seleccionados
'y les asigna los valores de los campos: existenc, fechainv, horainve, false de la tabla: sinven
'* Elimina de la Tabla: sinven los registros seleccinados para actualizar
'* Inserta Movimientos de Articulos en la Tabla: smoval
'-------------------------------------------------------------------
Dim Sql As String, ADonde As String
Dim Rs As ADODB.Recordset
Dim DevStock As String
Dim CanStock As Long, Diferencia As Long
Dim vTipoMov As CTiposMov
'Dim CodTipoMov As String * 3
Dim NumMovim As Long, NumLinea As Long
Dim LetraSerie As String * 1
Dim CadValues As String
Dim bol As Boolean
    
    On Error Resume Next
    
    'Obtener Registros de la Tabla:sinven de los que se va a actualizar el Stock
    Sql = "SELECT * "
    Sql = Sql & " FROM sinven "
    Sql = Sql & " WHERE " & cadFormula
   
    bol = True
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Rs.EOF Then
        bol = False
        ActualizarInventario = False
        MsgBox "No existen Registros en la Tabla: sinven para Actualizar Inventario.", vbExclamation
        Rs.Close
        Set Rs = Nothing
        Exit Function
    End If
    
    
    'Obtener el contador para los movimientos del Almacen que se esta inventariando
    'A cada registro de la tabla sinven se le asignará un numero de linea.
    '----------------------------------------------------------------------------
    Set vTipoMov = New CTiposMov
'    CodTipoMov = "REG"
    If vTipoMov.Leer("DFI") Then  'Se han cargado correctamente los valores de la clase
        'Obtener el contador para el codigo de Movimiento
        LetraSerie = vTipoMov.LetraSerie
        NumMovim = vTipoMov.ConseguirContador("DFI")
        NumLinea = 1
        bol = True
    Else
        bol = False
    End If
    
    If Not bol Then
        Set vTipoMov = Nothing
        Exit Function
    End If
    
   
    On Error GoTo EActualizarInven:
    'Aqui empieza la transaccion
    conn.BeginTrans
    
    While Not Rs.EOF And bol 'Para cada registro de la tabla sinven
    
        'Introducir Movimiento de Entrada/Salida si hay diferencia entre el
        'Stock del Sistema y el Stock Real Inventariado.
        '------------------------------------------------------------------
        ADonde = "Introduciendo Movimiento de Entrada/Salida. Tabla: smoval."
        DevStock = DevuelveDesdeBDNew(cAgro, "salmac", "canstock", "codartic", Rs!codArtic, "T", , "codalmac", Rs!codAlmac, "N")
        If DevStock <> "" Then
            CanStock = CLng(DevStock)
            Diferencia = Rs!existenc - CanStock
            If Diferencia <> 0 Then 'Insertar Movimiento de Entrada/Salida en Almacen
                CadValues = DBSet(Rs!codArtic, "T") & ", " & Rs!codAlmac & ", '" & Format(Rs!fechainv, "yyyy-mm-dd") & "', '"
                CadValues = CadValues & Format(Rs!horainve, "yyyy-mm-dd hh:mm:ss") & "', "
                bol = InsertarMovimArticulos(CadValues, Rs!codArtic, Diferencia, LetraSerie, NumMovim, NumLinea)
                NumLinea = NumLinea + 1
            Else
                bol = True
            End If
        Else
            bol = False
        End If
        
        'Actualizamos la Tabla: salmac
        '           salmac.canstock := existencia Real
        '           salmac.statusin := false (desbloqueamos los articulos )
        '---------------------------------------
        If bol Then
            ADonde = "Actualizando Stock de Articulos en Almacen. Tabla: salmac."
            Sql = "UPDATE salmac, sartic SET salmac.canstock=" & DBSet(Rs!existenc, "N") & ", salmac.statusin=0"
            '[Monica]04/12/2017: nos guardamos el pmp y preciouc del articulo de cuando se hace el inventario
            '                    he tenido que añadir la tabla sartic para sacar los precios
            Sql = Sql & ", salmac.preciomp = sartic.preciomp, salmac.preciouc = sartic.preciouc "
            Sql = Sql & " WHERE salmac.codartic=" & DBSet(Rs!codArtic, "T") & " AND salmac.codalmac=" & Rs!codAlmac
            Sql = Sql & " and salmac.codartic = sartic.codartic "
            conn.Execute Sql
        End If

        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
    If bol Then
'        'Pasamos la tabla de inventario real sinven al historico: shinve
'        'antes de eliminarla
'        ADonde = "Pasando registros de Inventario al Histórico: shinve."
'        SQL = "INSERT INTO shinve (codartic,codalmac,fechainv,horainve,existenc) "
'        SQL = SQL & "SELECT codartic,codalmac,fechainv,horainve,existenc "
'        SQL = SQL & " FROM sinven WHERE " & cadFormula
'        Conn.Execute SQL
    
        'Eliminamos los registros seleccionados de la Tabla: sinven
        '----------------------------------------------------------
        ADonde = "Eliminando registros de Inventario. Tabla: sinven."
        Sql = "DELETE FROM sinven "
        Sql = Sql & " WHERE " & cadFormula
        conn.Execute Sql
        
        
        'Incrementamos el contador para el Tipo De Movimiento
        '-----------------------------------------------------
        ADonde = "Actualizando el contador ."
        'bol = vTipoMov.IncrementarContador(
        vTipoMov.IncrementarContador ("DFI")
    End If
    Set vTipoMov = Nothing
        
EActualizarInven:
    If Err.Number <> 0 Or Not bol Then
         'Hay error , almacenamos y salimos
          Sql = "Actualizar Inventario." & vbCrLf & "----------------------------" & vbCrLf
          Sql = Sql & ADonde
          MuestraError Err.Number, Sql, Err.Description
          conn.RollbackTrans
          ActualizarInventario = False
          Set vTipoMov = Nothing
    Else
        ActualizarInventario = True
        conn.CommitTrans
    End If
End Function


Private Function InsertarMovimArticulos(CadValues As String, codArtic As String, Cantidad As Long, LetraSerie As String, NumMovim As Long, NumLinea As Long) As Boolean
Dim vImporte As Single, vPrecioVenta As String
Dim tipoMov As Byte
Dim Sql As String
On Error Resume Next
         
        'Obtener el precio de venta del articulo
        '++monica añadido el tipo de precio para el movimiento, antes solo el pmp
        If vParamAplic.TipoPrecio = 0 Then 'precio medio ponderado
             vPrecioVenta = DevuelveDesdeBDNew(cAgro, "sartic", "preciomp", "codartic", codArtic, "T")
        Else ' ultimo precio
             vPrecioVenta = DevuelveDesdeBDNew(cAgro, "sartic", "preciouc", "codartic", codArtic, "T")
        End If
        If vPrecioVenta <> "" Then
            vImporte = Cantidad * CSng(vPrecioVenta)
        Else
            vImporte = 0
        End If
        
        'Tipo de Movimiento de Almacen
        If Cantidad > 0 Then 'Insertar Movimiento de Entrada en Almacen
            tipoMov = 1
        ElseIf Cantidad < 0 Then 'Insertar Movimiento de Salida en Almacen
            tipoMov = 0
        End If
        
        Sql = "INSERT INTO smoval (codartic, codalmac, fechamov, horamovi, tipomovi, detamovi, cantidad, impormov, codigope, letraser, document, numlinea) "
        Sql = Sql & " VALUES (" & CadValues & tipoMov & ", '" & "DFI" & "', " & DBSet(Cantidad, "N") & ", " & DBSet(vImporte, "N") & ", 0,'" '--monica he quitado el trabajador  & Val(txtCodigo(21).Text) & ", '"
        Sql = Sql & LetraSerie & "', " & NumMovim & ", " & NumLinea & ")"
        conn.Execute Sql
        
        If Err.Number <> 0 Then
             'Hay error , almacenamos y salimos
            InsertarMovimArticulos = False
        Else
            InsertarMovimArticulos = True
        End If
    
End Function


Private Function ValidarCamposInventario() As Boolean
'Comprobar que los campos requeridos tienen valor antes de abrir el listado
Dim b As Boolean
        b = True
        'campo almacen debe tener valor
        If Trim(txtCodigo(13).Text) = "" Then
             MsgBox "El campo Almacen debe tener valor.", vbInformation
             PonerFoco txtCodigo(13)
             b = False
        End If
    
        'fecha de inventario debe tener valor
        If b Then
            If (Opcionlistado = 12 Or Opcionlistado = 15 Or Opcionlistado = 19) And Trim(txtCodigo(20).Text) = "" Then
                 MsgBox "El campo Fecha debe tener valor.", vbInformation
                 PonerFoco txtCodigo(20)
                 b = False
            End If
        End If
        ValidarCamposInventario = b
End Function



Private Function ListaArtActivos(cadwhere As String, FechaIn As String) As String
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Lista As String
'Devuelve una cadena con la concatenacion de todos los articulos que
'no debe seleccionar ya que si tienen movimientos con fecha posterior
'a FechaIn.
'ej:    "[""00000004"", ""00000033""]"

    Lista = "["
    
    Sql = "SELECT distinct smoval.codartic from smoval "
    If InStr(cadwhere, "sartic") > 0 Then Sql = Sql & " INNER JOIN sartic ON smoval.codartic=sartic.codartic "
    Sql = Sql & " WHERE " & Replace(cadwhere, "salmac", "smoval")
    If cadwhere <> "" Then Sql = Sql & " AND "
    Sql = Sql & " fechamov>='" & Format(FechaIn, FormatoFecha) & "' "
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
'        lista = lista & """" & RS.Fields(0).Value & """"
        Lista = Lista & DBSet(Rs.Fields(0).Value, "T")
        Rs.MoveNext
        If Not Rs.EOF Then Lista = Lista & ", "
    Wend
    Lista = Lista & "]"
    ListaArtActivos = Lista
    Rs.Close
    Set Rs = Nothing
End Function



Private Sub ActualizarImprimir()
Dim i As Long
Dim Desde As Long, Hasta As Long
Dim Sql As String

    Select Case Opcionlistado
    Case 7  'TRASPASO ALMACEN
        If frmVisReport.EstaImpreso = True Then
        'Desde=-1 si estamos en Historico ya que aqui no se introducen valores Desde/Hasta
            If Trim(txtCodigo(3).Text) <> "" Then Desde = CLng(txtCodigo(3).Text)
            If Trim(txtCodigo(4).Text) <> "" Then Hasta = CLng(txtCodigo(4).Text)
            For i = Desde To Hasta
                Sql = "UPDATE scatra SET situacio=1" 'Impreso
                Sql = Sql & " WHERE codtrasp=" & i
                conn.Execute Sql
            Next i
        End If
        
    Case 8  'MOVIMIENTO ALMACEN
        If frmVisReport.EstaImpreso = True Then
           'Desde=-1 si estamos en Historico ya que aqui no se introducen valores Desde/Hasta
           If Trim(txtCodigo(3).Text) <> "" Then Desde = CLng(txtCodigo(3).Text)
           If Trim(txtCodigo(4).Text) <> "" Then Hasta = CLng(txtCodigo(4).Text)
           For i = Desde To Hasta
                Sql = "UPDATE scamov SET situacio=1"
                Sql = Sql & " WHERE codmovim=" & i
                conn.Execute Sql
           Next i
        End If
        
    Case 10  'MOVIMIENTO DE SERVICIOS DE ALMACEN
        If frmVisReport.EstaImpreso = True Then
           'Desde=-1 si estamos en Historico ya que aqui no se introducen valores Desde/Hasta
           If Trim(txtCodigo(3).Text) <> "" Then Desde = CLng(txtCodigo(3).Text)
           If Trim(txtCodigo(4).Text) <> "" Then Hasta = CLng(txtCodigo(4).Text)
           For i = Desde To Hasta
                Sql = "UPDATE scaser SET situacio=1"
                Sql = Sql & " WHERE codservi=" & i
                conn.Execute Sql
           Next i
        End If
    
    End Select
End Sub


Private Sub InicializarVbles()
    cadFormula = ""
    cadselect = ""
    cadParam = ""
    numParam = 0
'    cadTitulo = ""
'    cadNomRPT = ""
    conSubRPT = False
End Sub


Private Function PonerDesdeHasta(campo As String, Tipo As String, indD As Byte, indH As Byte, param As String) As Boolean
Dim devuelve As String
Dim Cad As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    
    If Tipo <> "F" Then
        'Fecha para Crystal Report
        If Not AnyadirAFormula(cadselect, devuelve) Then Exit Function
    Else
        'Fecha para la Base de Datos
        Cad = CadenaDesdeHastaBD(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
        If Not AnyadirAFormula(cadselect, Cad) Then Exit Function
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
        .Titulo = cadTitulo
        .NombreRPT = cadNomRPT
        .ConSubInforme = conSubRPT
        .Show vbModal
    End With
End Sub


Private Function PonerGrupo(numGrupo As Byte, cadgrupo As String) As Byte
Dim campo As String
Dim nomCampo As String

    campo = "pGroup" & numGrupo & "="
    nomCampo = "pGroup" & numGrupo & "Name="
    PonerGrupo = 0
    
    Select Case cadgrupo
        Case "Familia"
            cadParam = cadParam & campo & "{sartic.codfamia}" & "|"
            If numGrupo = 1 Then
                cadParam = cadParam & nomCampo & " ""FAMILIA: "" & " & " totext({sartic.codfamia},""0000"") & " & """  """ & " & {sfamia.nomfamia}" & "|"
            Else
                cadParam = cadParam & nomCampo & " totext({sartic.codfamia},""0000"") & " & """ """ & " & {sfamia.nomfamia}" & "|"
            End If
            numParam = numParam + 1
        Case "Marca"
            cadParam = cadParam & campo & "{sartic.codmarca}" & "|"
            If numGrupo = 1 Then
                cadParam = cadParam & nomCampo & " ""MARCA: "" & " & " totext({sartic.codmarca},""0000"") & " & """  """ & " & {smarca.nommarca}" & "|"
            Else
                cadParam = cadParam & nomCampo & " totext({sartic.codmarca},""0000"") & " & """ """ & " & {smarca.nommarca}" & "|"
            End If
            numParam = numParam + 1
        Case "Proveedor"
            cadParam = cadParam & campo & "{sartic.codprove}" & "|"
            If numGrupo = 1 Then
                cadParam = cadParam & nomCampo & " ""PROVEEDOR: "" & " & " totext({sartic.codprove},""000000"") & " & """  """ & " & {sprove.nomprove}" & "|"
            Else
                cadParam = cadParam & nomCampo & " totext({sartic.codprove},""000000"") & " & """ """ & " & {sprove.nomprove}" & "|"
            End If
            numParam = numParam + 1
            PonerGrupo = numGrupo
        Case "Tipo Articulo"
            cadParam = cadParam & campo & "{sartic.codtipar}" & "|"
            If numGrupo = 1 Then
                cadParam = cadParam & nomCampo & " ""TIPO ARTICULO: "" & " & " {sartic.codtipar} & " & """  """ & " & {stipar.nomtipar}" & "|"
            Else
                cadParam = cadParam & nomCampo & " {sartic.codtipar} & " & """ """ & " & {stipar.nomtipar}" & "|"
            End If
            numParam = numParam + 1
    End Select

End Function



Private Sub AbrirFrmAlmPropios()
    Set frmMtoAlPropios = New frmManAlmProp
    frmMtoAlPropios.DatosADevolverBusqueda = "0|1|"
    frmMtoAlPropios.DeConsulta = True
    frmMtoAlPropios.Show vbModal
    Set frmMtoAlPropios = Nothing
End Sub


Private Sub AbrirFrmTipoArt()
'Tipos de Articulos
    Set frmMtoTArticulo = New frmManTipArtic
    frmMtoTArticulo.DatosADevolverBusqueda = "0|1|"
    frmMtoTArticulo.DeConsulta = True
    frmMtoTArticulo.Show vbModal
    Set frmMtoTArticulo = Nothing
End Sub

Private Sub AbrirFrmClientes()
'Clientes
    Set frmMtoClientes = New frmClientes
    frmMtoClientes.DatosADevolverBusqueda = "0|1|"
    frmMtoClientes.Show vbModal
    Set frmMtoClientes = Nothing
End Sub


Private Function ComprobarFechasConta(ind As Integer) As Boolean
'comprobar que el periodo de fechas a contabilizar esta dentro del
'periodo de fechas del ejercicio de la contabilidad
Dim FechaIni As String, FechaFin As String
Dim Cad As String
Dim Rs As ADODB.Recordset
    
On Error GoTo EComprobar

    ComprobarFechasConta = False
    
    If txtCodigo(ind).Text <> "" Then
        FechaIni = "Select fechaini,fechafin From parametros"
        Set Rs = New ADODB.Recordset
        Rs.Open FechaIni, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        If Not Rs.EOF Then
            FechaIni = DBLet(Rs!FechaIni, "F")
            FechaFin = DateAdd("yyyy", 1, DBLet(Rs!FechaFin, "F"))
            'nos guardamos los valores
            Orden1 = FechaIni
            Orden2 = FechaFin
        
            If Not EntreFechas(FechaIni, txtCodigo(ind).Text, FechaFin) Then
                 Cad = "El período de contabilización debe estar dentro del ejercicio:" & vbCrLf & vbCrLf
                 Cad = Cad & "    Desde: " & FechaIni & vbCrLf
                 Cad = Cad & "    Hasta: " & FechaFin
                 MsgBox Cad, vbExclamation
                 txtCodigo(ind).Text = ""
            Else
                ComprobarFechasConta = True
            End If
        End If
        Rs.Close
        Set Rs = Nothing
    Else
        ComprobarFechasConta = True
    End If
    
EComprobar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Fechas", Err.Description
End Function



Private Sub ContabilizarFacturas(cadTABLA As String, cadwhere As String)
'Contabiliza Facturas de Clientes o de Proveedores
Dim Sql As String
Dim b As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste As String

    If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then
        Sql = "VENCON" 'contabilizar facturas de venta
    Else
        If cadTABLA = "scafpc" Then
            Sql = "COMCON" 'contabilizar facturas de compra
        Else
            Sql = "TRACON" 'contabilizar facturas de trasnporte
        End If
    End If
    
    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se pueden Contabilizar Facturas. Hay otro usuario contabilizando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If


     'comprobar que se han rellenado los dos campos de fecha
     'sino rellenar con fechaini o fechafin del ejercicio
     'que guardamos en vbles Orden1,Orden2
     If txtCodigo(31).Text = "" Then
        txtCodigo(31).Text = Orden1 'fechaini del ejercicio de la conta
     End If

     If txtCodigo(32).Text = "" Then
        txtCodigo(32).Text = Orden2 'fecha fin del ejercicio de la conta
     End If


     'Comprobar que el intervalo de fechas D/H esta dentro del ejercicio de la
     'contabilidad par ello mirar en la BD de la Conta los parámetros
     If Not ComprobarFechasConta(32) Then Exit Sub



    'comprobar si existen en Ariagro facturas anteriores al periodo solicitado
    'sin contabilizar
    If Me.txtCodigo(31).Text <> "" Then 'anteriores a fechadesde
        Sql = "SELECT COUNT(*) FROM " & cadTABLA
        If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then
            Sql = Sql & " WHERE fecfactu <"
        ElseIf cadTABLA = "scafpc" Or cadTABLA = "tcafpc" Then
            Sql = Sql & " WHERE fecrecep <"
        End If
        Sql = Sql & DBSet(txtCodigo(31), "F") & " AND intconta=0 "
        If RegistrosAListar(Sql) > 0 Then
            MsgBox "Hay Facturas anteriores sin contabilizar.", vbExclamation
            Exit Sub
        End If
    End If


'    'Laura: 11/10/2006 bloquear los registros q vamos a contabilizar
'    If Not BloqueaRegistro(cadTabla, cadWhere) Then
'        MsgBox "No se pueden Contabilizar Facturas. Hay registros bloqueados.", vbExclamation
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If


    '==========================================================
    'REALIZAR COMPROBACIONES ANTES DE CONTABILIZAR FACTURAS
    '==========================================================

'    Me.lblProgess(0).Caption = "Comprobaciones: "
'    CargarProgres Me.ProgressBar1, 100

    'Cargar tabla TEMP con las Facturas que vamos a Trabajar
    b = CrearTMPFacturas(cadTABLA, cadwhere)
    If Not b Then Exit Sub
    

    'Laura: 11/10/2006 bloquear los registros q vamos a contabilizar
'    TerminaBloquear
    Sql = cadTABLA & " INNER JOIN tmpFactu ON " & cadTABLA
    If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then
        Sql = Sql & ".codtipom=tmpFactu.codtipom AND "
    Else
        If cadTABLA = "scafpc" Then
            Sql = Sql & ".codprove=tmpFactu.codprove AND "
        Else
            Sql = Sql & ".codtrans=tmpFactu.codtrans AND "
        End If
    End If
    
    Sql = Sql & cadTABLA & ".numfactu=tmpFactu.numfactu AND " & cadTABLA & ".fecfactu=tmpFactu.fecfactu "
    If Not BloqueaRegistro(Sql, cadwhere) Then
        MsgBox "No se pueden Contabilizar Facturas. Hay registros bloqueados.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If


    '---- Preparamos la pantalla de Contabilizar
    'Visualizar la barra de Progreso
    Me.FrameRepxDia.Height = 5100
    Me.Height = Me.FrameRepxDia.Height
    Me.FrameProgress.visible = True
    Me.FrameProgress.Top = 3350

    Me.lblProgess(0).Caption = "Comprobaciones: "
    CargarProgres Me.ProgressBar1, 100


    'comprobar que todas las LETRAS SERIE existen en la contabilidad y en Ariges
    '--------------------------------------------------------------------------
    IncrementarProgres Me.ProgressBar1, 10
    If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then
        Me.lblProgess(1).Caption = "Comprobando letras de serie ..."
        b = ComprobarLetraSerie(cadTABLA)
    End If
    IncrementarProgres Me.ProgressBar1, 10
    Me.Refresh
    If Not b Then Exit Sub


    'comprobar que no haya Nº FACTURAS en la contabilidad para esa fecha
    'que ya existan
    '-----------------------------------------------------------------------
    If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then
        Me.lblProgess(1).Caption = "Comprobando Nº Facturas en contabilidad ..."
        Sql = "anofaccl>=" & Year(txtCodigo(31).Text) & " AND anofaccl<= " & Year(txtCodigo(32).Text)
        b = ComprobarNumFacturas_new(cadTABLA, Sql)
    End If
    IncrementarProgres Me.ProgressBar1, 20
    Me.Refresh
    If Not b Then Exit Sub


    'comprobar que todas las CUENTAS de los distintos clientes que vamos a
    'contabilizar existen en la Conta: sclien.codmacta IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgess(1).Caption = "Comprobando Cuentas Contables en contabilidad ..."
    b = ComprobarCtaContable_new(cadTABLA, 1)
    IncrementarProgres Me.ProgressBar1, 20
    Me.Refresh
    If Not b Then Exit Sub


    'comprobar que todas las CUENTAS de venta de la familia de los articulos que vamos a
    'contabilizar existen en la Conta: sfamia.ctaventa IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then
        Me.lblProgess(1).Caption = "Comprobando Cuentas Ctbles Ventas en contabilidad ..."
    Else
        If cadTABLA = "scafpc" Then
            Me.lblProgess(1).Caption = "Comprobando Cuentas Ctbles Compras en contabilidad ..."
        End If
    End If
    If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Or cadTABLA = "scafpc" Then
        b = ComprobarCtaContable_new(cadTABLA, 2)
    Else
        ' comprobamos que existe las cuentas vparamaplic.ctatrareten
        b = ComprobarCtaContable_new(cadTABLA, 4)
    End If
    IncrementarProgres Me.ProgressBar1, 20
    Me.Refresh
    If Not b Then Exit Sub


    'comprobar que todas las CUENTAS de venta de las variedades que vamos a
    'contabilizar existen en la Conta: (variedades.raiz + tipomer.codtimer + variedades.codvarie) IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then
        Me.lblProgess(1).Caption = "Comprobando Cuentas Ctbles Ventas en contabilidad ..."
        b = ComprobarCtaContable_new(cadTABLA, 8)
    End If
'    IncrementarProgres Me.ProgressBar1, 20
    Me.Refresh
    If Not b Then Exit Sub



    'comprobar que todas las CUENTAS de transporte interior y exterior de las variedades que vamos a
    'contabilizar existen en la Conta: (variedades.ctatrainterior + variedades.ctatraexporta) IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    If cadTABLA = "tcafpc" Then
        Me.lblProgess(1).Caption = "Comprobando Cuentas Ctbles de Variedades en contabilidad ..."
        b = ComprobarCtaContable_new(cadTABLA, 5) ' cuentas de variedades de transporte
        Me.Refresh
        If Not b Then Exit Sub
        b = ComprobarCtaContable_new(cadTABLA, 9) ' cuentas de portes de vuelta
        If Not b Then Exit Sub
        b = ComprobarCtaContable_new(cadTABLA, 12) ' cuentas de variedades de comisionistas
        If Not b Then Exit Sub
    End If
'    IncrementarProgres Me.ProgressBar1, 20


    'comprobar que todos las TIPO IVA de las distintas fecturas que vamos a
    'contabilizar existen en la Conta: scafac.codigiv1,codigiv2,codigiv3 IN (conta.tiposiva.codigiva)
    '--------------------------------------------------------------------------
    Me.lblProgess(1).Caption = "Comprobando Tipos de IVA en contabilidad ..."
    b = ComprobarTiposIVA(cadTABLA)
    IncrementarProgres Me.ProgressBar1, 10
    Me.Refresh
    If Not b Then Exit Sub


    'comprobar si hay contabilidad ANALITICA: conta.parametros.autocoste=1
    'y verificar que las cuentas de sfamia.ctaventa empiezan por el digito
    'de conta.parametros.grupogto o conta.parametros.grupovta
    'obtener el centro de coste del usuario para insertarlo en linfact
    If vEmpresa.TieneAnalitica Then  'hay contab. analitica
       Me.lblProgess(1).Caption = "Comprobando Contabilidad Analítica ..."
       If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Or cadTABLA = "scafpc" Then
           b = ComprobarCtaContable_new(cadTABLA, 3)
       Else
           b = ComprobarCtaContable_new(cadTABLA, 7)
           If b Then b = ComprobarCtaContable_new(cadTABLA, 10)
           If b Then b = ComprobarCtaContable_new(cadTABLA, 13)
        If Not b Then Exit Sub
           
           
       End If
       If Not b Then Exit Sub

       '(si tiene analítica requiere un centro de coste para insertar en conta.linfact)
'--monica:no tenemos trabajadores que tengan asociado un centro de coste está en variedades
       CCoste = ""
       
       If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then
            b = ComprobarCCoste_new(CCoste, cadTABLA, 1)
            If Not b Then Exit Sub
            b = ComprobarCCoste_new(CCoste, cadTABLA, 2)
            If Not b Then Exit Sub
            b = ComprobarCCoste_new(CCoste, cadTABLA, 3)
            If Not b Then Exit Sub
       Else
            b = ComprobarCCoste_new(CCoste, cadTABLA)
            If Not b Then Exit Sub
       End If
'++monica
        CCoste = ""
    End If
    IncrementarProgres Me.ProgressBar1, 10
    Me.Refresh


    If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then b = ComprobarFormadePago(cadTABLA)
    If Not b Then Exit Sub
        

    '===========================================================================
    'CONTABILIZAR FACTURAS
    '===========================================================================
    Me.lblProgess(0).Caption = "Contabilizar Facturas: "
    CargarProgres Me.ProgressBar1, 10
    Me.lblProgess(1).Caption = "Insertando Facturas en Contabilidad..."



    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    LOG.Insertar 3, vUsu, "Contabilizar facturas: " & vbCrLf & cadTABLA & vbCrLf & cadwhere
    Set LOG = Nothing
    '-----------------------------------------------------------------------------




    '---- Crear tabla TEMP para los posible errores de facturas
    tmpErrores = CrearTMPErrFact(cadTABLA)

    '---- Pasar las Facturas a la Contabilidad
    b = PasarFacturasAContab(cadTABLA, CCoste)

    '---- Mostrar ListView de posibles errores (si hay)
    If Not b Then
        If tmpErrores Then
            'Cargar un listview con la tabla TEMP de Errores y mostrar
            'las facturas que fallaron
            frmMensajes.OpcionMensaje = 10
            frmMensajes.Show vbModal
        Else
            MsgBox "No pueden mostrarse los errores.", vbInformation
        End If
    Else
        MsgBox "El proceso ha finalizado correctamente.", vbInformation
    End If

    'Este bien o mal, si son proveedores abriremos el listado
    'Imprimimiremos un listado de contabilizacion de facturas
    '------------------------------------------------------
    If cadTABLA = "scafpc" Or cadTABLA = "tcafpc" Then
        If DevuelveValor("Select count(*) from tmpinformes where codusu = " & vUsu.Codigo) > 0 Then
            InicializarVbles
            cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
            numParam = numParam + 1
            
            cadParam = "|pDHFecha=""" & vUsu.Nombre & "   Hora: " & Format(Now, "hh:mm") & """|"
            numParam = numParam + 1
            cadFormula = "({tmpinformes.codusu} =" & vUsu.Codigo & ")"
            conSubRPT = False
            If cadTABLA = "scafpc" Then
                cadTitulo = "Listado contabilizacion FRAPRO"
                cadNomRPT = "rContabPRO.rpt"
            Else
                cadTitulo = "Listado contabilizacion FRATRA"
                cadNomRPT = "rContabTRA.rpt"
            End If
            
            LlamarImprimir
        End If
    End If


    '---- Eliminar tabla TEMP de Errores
    BorrarTMPErrFact

End Sub


Private Function PasarFacturasAContab(cadTABLA As String, CCoste As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim b As Boolean
Dim i As Integer
Dim NumFactu As Integer
Dim Codigo1 As String

    On Error GoTo EPasarFac

    PasarFacturasAContab = False

    '---- Obtener el total de Facturas a Insertar en la contabilidad
    Sql = "SELECT count(*) "
    Sql = Sql & " FROM " & cadTABLA & " INNER JOIN tmpFactu "
    If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then
        Codigo1 = "codtipom"
    Else
        If cadTABLA = "scafpc" Then
            Codigo1 = "codprove"
        Else
            Codigo1 = "codtrans"
        End If
    End If
    Sql = Sql & " ON " & cadTABLA & "." & Codigo1 & "=tmpFactu." & Codigo1
    Sql = Sql & " AND " & cadTABLA & ".numfactu=tmpFactu.numfactu AND " & cadTABLA & ".fecfactu=tmpFactu.fecfactu "

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        NumFactu = Rs.Fields(0)
    Else
        NumFactu = 0
    End If
    Rs.Close
    Set Rs = Nothing


    'Modificacion como David
    '-----------------------------------------------------------
    ' Mosrtaremos para cada factura de PROVEEDOR
    ' que numregis le ha asignado
    Sql = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
    conn.Execute Sql


    Set cContaFra = New cContabilizarFacturas
    
    If Not cContaFra.EstablecerValoresInciales(ConnConta) Then
        'NO ha establcedio los valores de la conta.  Le dejaremos seguir, avisando que
        ' obviamente, no va a contabilizar las FRAS
        Sql = "Si continua, las facturas se insertaran en el registro, pero no serán contabilizadas" & vbCrLf
        Sql = Sql & "en este momento. Deberán ser contabilizadas desde el ARICONTA" & vbCrLf & vbCrLf
        Sql = Sql & Space(50) & "¿Continuar?"
        If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
    End If




    '---- Pasar cada una de las facturas seleccionadas a la Conta
    If NumFactu > 0 Then
        CargarProgres Me.ProgressBar1, NumFactu

        'seleccinar todas las facturas que hemos insertado en la temporal (las que vamos a contabilizar)
        Sql = "SELECT * "
        Sql = Sql & " FROM tmpFactu "

        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenStatic, adLockPessimistic, adCmdText
        i = 1

        b = True
        'pasar a contabilidad cada una de las facturas seleccionadas
        While Not Rs.EOF
            If cadTABLA = "facturas" Or cadTABLA = "facturassocio" Then
                Sql = cadTABLA & "." & Codigo1 & "=" & DBSet(Rs.Fields(0), "T") & " and numfactu=" & Rs!NumFactu
                Sql = Sql & " and fecfactu=" & DBSet(Rs!FecFactu, "F")
                If PasarFactura(Sql, CCoste, txtCodigo(0).Text, cadTABLA, cContaFra) = False And b Then b = False
            Else
                If cadTABLA = "scafpc" Then
                    Sql = cadTABLA & "." & Codigo1 & "=" & DBSet(Rs.Fields(0), "N") & " and numfactu=" & DBSet(Rs!NumFactu, "T")
                    Sql = Sql & " and fecfactu=" & DBSet(Rs!FecFactu, "F")
                    If PasarFacturaProv(Sql, CCoste, Orden2, cContaFra) = False And b Then b = False
                Else
                    Sql = cadTABLA & "." & Codigo1 & "=" & DBSet(Rs.Fields(0), "N") & " and numfactu=" & DBSet(Rs!NumFactu, "T")
                    Sql = Sql & " and fecfactu=" & DBSet(Rs!FecFactu, "F")
                    If PasarFacturaTrans(Sql, CCoste, Orden2, cContaFra) = False And b Then b = False
                End If
            End If

            '---- Laura 26/10/2006
            'Al pasar cada factura al hacer el commit desbloqueamos los registros
            'que teniamos bloqueados y los volvemos a bloquear
            'Laura: 11/10/2006 bloquear los registros q vamos a contabilizar
            Sql = cadTABLA & " INNER JOIN tmpFactu ON " & cadTABLA & "." & Codigo1 & "=tmpFactu." & Codigo1 & " AND " & cadTABLA & ".numfactu=tmpFactu.numfactu AND " & cadTABLA & ".fecfactu=tmpFactu.fecfactu "
            If Not BloqueaRegistro(Sql, cadTABLA & "." & Codigo1 & "=tmpFactu." & Codigo1 & " AND " & cadTABLA & ".numfactu=tmpFactu.numfactu AND " & cadTABLA & ".fecfactu=tmpFactu.fecfactu") Then
'                MsgBox "No se pueden Contabilizar Facturas. Hay registros bloqueados.", vbExclamation
'                Screen.MousePointer = vbDefault
'                Exit Sub
            End If
            '----

            IncrementarProgres Me.ProgressBar1, 1
            Me.lblProgess(1).Caption = "Insertando Facturas en Contabilidad...   (" & i & " de " & NumFactu & ")"
            Me.Refresh
            i = i + 1
            Rs.MoveNext
        Wend

        Rs.Close
        Set Rs = Nothing
    End If

EPasarFac:
    If Err.Number <> 0 Then b = False

    If b Then
        PasarFacturasAContab = True
    Else
        PasarFacturasAContab = False
    End If
End Function



Private Sub ListadosAlmacen(H As Integer, W As Integer)
    'LISTADOS DE ALMACENES
    '---------------------
    Select Case Opcionlistado
        Case 1   'Listados de Marcas
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado de Marcas"
            indFrame = 1
            Codigo = "{smarca.codmarca}"
            Orden1 = "{smarca.codmarca}"
            Orden2 = "{smarca.nommarca}"
            cadTitulo = "Listado Marcas"
            cadNomRPT = "rAlmMarcas.rpt"
            conSubRPT = False
            
        Case 2   'Listado de Almacenes Propios
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado de Almacenes"
            indFrame = 1
            Codigo = "{salmpr.codalmac}"
            Orden1 = "{salmpr.codalmac}"
            Orden2 = "{salmpr.nomalmac}"
            cadTitulo = "Listado Almacenes Propios"
            cadNomRPT = "rAlmAPropios.rpt"
            conSubRPT = False
            
        Case 3   'Listado de Tipos de Unidad
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado Tipos de Unidad"
            indFrame = 1
            Codigo = "{sunida.codunida}"
            Orden1 = "{sunida.codunida}"
            Orden2 = "{sunida.nomunida}"
            cadTitulo = "Listado Tipos de Unidad"
            cadNomRPT = "rAlmTUnidad.rpt"
            conSubRPT = False
            
        Case 4   'Listado de Tipos de Artículos
            PonerFrameListadoVisible True, H, W
            Me.lblTitulo(1).Caption = "Listado Tipos de Artículos"
            indFrame = 1
            Codigo = "{stipar.codtipar}"
            Orden1 = "{stipar.codtipar}"
            Orden2 = "{stipar.nomtipar}"
            txtCodigo(1).Tag = CadTag
            txtCodigo(2).Tag = CadTag
            cadTitulo = "Listado Tipos de Artículos"
            cadNomRPT = "rAlmTArticulo.rpt"
            conSubRPT = False
            
        Case 6    'Listado de Artículo
            ponerFrameArticulosVisible True, H, W
            CargarListViewOrden
            Codigo = "{sartic"
            indFrame = 11
            cadTitulo = "Listado de Artículos"
            
            
        Case 18, 247 'Informe Stocks Maximos y Minimos   'OPCION: 247 es este tb
            ponerFrameArticulosVisible True, H, W
            Codigo = "{salmac"
            indFrame = 11
            
        Case 7, 8 '7: Informe de Traspasos de Almacen
                  '8: Informe de Movimientos de Almacen
            If Opcionlistado = 7 Then
                Me.lblTitulo(2).Caption = "Informe Traspaso de Almacen"
                Me.Label2(1).Caption = "Nº Traspaso"
                Codigo = "{scatra.codtrasp}"
            Else
                Me.lblTitulo(2).Caption = "Informe Movimientos de Almacen"
                Me.Label2(1).Caption = "Nº Movimiento"
                Codigo = "{scamov.codmovim}"
            End If
            H = 3495
            W = 5835
            PonerFrameVisible Me.FrameInfAlmacen, True, H, W
            indFrame = 2
            If NumCod <> "" Then
                txtCodigo(3).Text = NumCod
                txtCodigo(4).Text = NumCod
            End If
            
        Case 9 'Informe Movimiento Artículos
            W = 10700
            H = 6795
            PonerFrameVisible Me.FrameMovArtic, True, H, W
            indFrame = 3
            Codigo = "{smoval.codartic}"
            If DeServicios Then
                cadTitulo = "Informe de Servicios"
                lblTitulo(3).Caption = "Informe Movimientos de Servicios"
                Label3(68).Caption = "Cliente/Socio"
            Else
                cadTitulo = "Informe Movimientos Articulos"
            End If
            conSubRPT = True
            ChKAcumulados.visible = True
            ChKAcumulados.Enabled = True
            CargarListView
            
        Case 10 'Impresion de albaranes de movimientos de servicios
            Me.lblTitulo(2).Caption = "Informe Movimientos de Servicios"
            Me.Label2(1).Caption = "Nº Movimiento"
            Codigo = "{scaser.codservi}"
            H = 3495
            W = 5835
            PonerFrameVisible Me.FrameInfAlmacen, True, H, W
            indFrame = 2
            If NumCod <> "" Then
                txtCodigo(3).Text = NumCod
                txtCodigo(4).Text = NumCod
            End If
            
        Case 12 '12: Listado Toma de Inventario Articulos
            PonerFrameInventarioVisible True, H, W
            indFrame = 4
            Me.chkImprimeStock.visible = True
            Me.lbltituloInven.Caption = "Listado Toma de Inventario Articulos"
            cadTitulo = "Toma Inventario Articulos"
            'codigo = "{salmac.codalmac}"
            
        Case 13 '13: Listado Diferencias de Inventario Articulos
            PonerFrameInventarioVisible True, H, W
            indFrame = 4
            Me.lbltituloInven.Caption = "Listado Diferencias de Inventario Articulos"
            'codigo = "{sinven.codalmac}"
            cadTitulo = "Diferencias Inventario Articulos"
            
        Case 14 '14: Actualizar Direfencias Inventario (NO IMPRIME INFORME)
            PonerFrameInventarioVisible True, H, W
            indFrame = 4
            Me.lbltituloInven.Caption = "Actualizar Diferencias de Inventario de Articulos"
            Me.Caption = "Inventario de Articulos"
            
        Case 15 '15: Listado de Articulos Inactivos
            PonerFrameInventarioVisible True, H, W
            indFrame = 4
            Me.lbltituloInven.Caption = "Listado Articulos Inactivos"
            cadTitulo = "Listado Articulos Inactivos"
    
        Case 16 '16 .- Listado Valoracion de Stocks Inventariados
            PonerFrameInventarioVisible True, H, W
            indFrame = 4
            Me.lbltituloInven.Caption = "Listado Valoración Stocks Inventariados"
            cadTitulo = "Listado Valoración Stocks Inventariados"
            
        Case 17 '17 .- Listado Valoración Stocks
            PonerFrameInventarioVisible True, H, W
            indFrame = 4
            Me.lbltituloInven.Caption = "Listado Valoración Stocks"
            cadTitulo = "Listado Valoración Stocks"
            
        Case 19 '19 .- Inf. Stocks a una Fecha
            PonerFrameInventarioVisible True, H, W
            indFrame = 4
            Me.lbltituloInven.Caption = "Informe Stocks a una Fecha"
            cadTitulo = "Stocks a una Fecha"
            
        Case 120 '120 .- Recálculo del Precio Medio Ponderado
            H = 5985
            W = 7275
            PonerFrameVisible Me.FrameRecalPMP, True, H, W
            cadTitulo = "Recálculo del Precio Medio Ponderado"
        
    End Select
End Sub




Private Sub ListadosCompras(H As Integer, W As Integer)
'=============================================
'==== Listados de COMPRAS

    Select Case Opcionlistado
        Case 309 '309: Listado precios de compra
            H = 4450
            W = 6920
            PonerFrameVisible Me.FrameDtosFM, True, H, W
            Me.Frame4.visible = True
            Me.Frame4.Top = 840
            Me.Frame5.visible = False
            Me.Frame6.visible = False
            Me.cmdAceptarDtosFM.Top = 3500
            Me.cmdCancel(12).Top = Me.cmdAceptarDtosFM.Top
            indFrame = 12 '6
    End Select
End Sub

Private Sub CargarIconos()
Dim i As Integer
    
    For i = 1 To 4
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i

    For i = 0 To 5
        Me.imgBuscarG(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 8 To 17
        Me.imgBuscarG(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 21 To 22
        Me.imgBuscarG(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 29 To 35
        Me.imgBuscarG(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 63 To 66
        Me.imgBuscarG(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 89 To 90
        Me.imgBuscarG(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 98 To 98
        Me.imgBuscarG(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    

End Sub


'
Private Function CargarTablaTemporal(vWhere As String) As Boolean
Dim Sql As String
Dim SQL1 As String
Dim Rs As ADODB.Recordset
Dim SqlValues As String


    On Error GoTo eCargarTablaTemporal
    
    CargarTablaTemporal = False

    Sql = "delete from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
    conn.Execute Sql

    Sql = "SELECT `smoval`.`fechamov`, `smoval`.`tipomovi`, `sartic`.`codfamia`, `sfamia`.`nomfamia`, `smoval`.`detamovi`, `smoval`.`cantidad`, `smoval`.`codartic`, `sartic`.`nomartic`, `smoval`.`codigope`, `smoval`.`impormov`, `smoval`.`document` "
    Sql = Sql & " FROM   (`smoval` `smoval` INNER JOIN `sartic` `sartic` ON `smoval`.`codartic`=`sartic`.`codartic`) INNER JOIN `sfamia` `sfamia` ON `sartic`.`codfamia`=`sfamia`.`codfamia` "
    If vWhere <> "" Then Sql = Sql & " where " & vWhere
    Sql = Sql & " ORDER BY `smoval`.`detamovi`, `smoval`.`codigope`, `sartic`.`codfamia`, `smoval`.`codartic`, `smoval`.`fechamov` DESC"

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SqlValues = ""
    While Not Rs.EOF
        SqlValues = SqlValues & "(" & vUsu.Codigo & "," & DBSet(Rs!Fechamov, "F") & "," & DBSet(Rs!tipomovi, "N") & "," & DBSet(Rs!codfamia, "N") & "," & DBSet(Rs!detamovi, "T") & ","
        SqlValues = SqlValues & DBSet(Rs!Cantidad, "N") & "," & DBSet(Rs!codArtic, "T") & "," & DBSet(Rs!codigope, "N") & "," & DBSet(Rs!impormov, "N") & ","
        SqlValues = SqlValues & DBSet(Rs!Document, "N") & ","
        
        If DBLet(Rs!detamovi, "T") = "RES" Or DBLet(Rs!detamovi, "T") = "SES" Then
            SqlValues = SqlValues & "0),"
        Else
            SqlValues = SqlValues & "1),"
        End If
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    ' quitamos la ultima coma
    If SqlValues <> "" Then
        SqlValues = Mid(SqlValues, 1, Len(SqlValues) - 1)
    
        Sql = "insert into tmpinformes (codusu, fecha1, campo1, importe4, nombre1, importe1, nombre2, codigo1, importe2, importe3, campo2) values "
    
        conn.Execute Sql & SqlValues
    End If
    
    CargarTablaTemporal = True
    Exit Function
    
eCargarTablaTemporal:
    MuestraError Err.Number, "Carga Tabla Temporal"
End Function


Private Function RecalculoPrecioMedioPonderado() As Boolean
Dim codArtic As String
Dim existencia As Currency
Dim PrecioMP As Currency
Dim Aux As Currency
Dim YaEstaFijadoPrecio As Boolean

Dim miSQL As String


    On Error GoTo eRecalculoPrecioMedioPonderado
    RecalculoPrecioMedioPonderado = False
    
    Label3(41).Caption = "Preparando datos"
    Label3(41).Refresh
    
    conn.Execute "Delete from tmpinformes where codusu = " & vUsu.Codigo
    
    Set miRsAux = New ADODB.Recordset
    miSQL = "Select sartic.codartic,preciomp,nomartic,codprove,codfamia,smoval.* from sartic,smoval WHERE " & cadselect
    miSQL = miSQL & " ORDER BY sartic.codartic,fechamov,horamovi"
    
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    codArtic = "@DABIZ"
    cadParam = "INSERT INTO tmpInformes(codusu,codigo1,campo1,campo2,nombre1,nombre2,precio1,precio2) VALUES "  'LOS ISNERTS
    cadTitulo = ""
    NumRegElim = 0
    cadNomRPT = ""
    While Not miRsAux.EOF
        If codArtic <> miRsAux!codArtic Then
            Label3(41).Caption = miRsAux!NomArtic
            Label3(41).Refresh
            
            
           ' If miRsAux!codArtic = "106980" Then Stop
            
            If codArtic <> "@DABIZ" Then
                'Metemos en tmp
                If YaEstaFijadoPrecio Then
                    cadNomRPT = cadNomRPT & DBSet(PrecioMP, "N") & ")" 'metemos el calulado
                    cadTitulo = cadTitulo & ", " & cadNomRPT
                
                    If Len(cadTitulo) > 3000 Then
                        'INSERTO TMP
                        cadTitulo = Mid(cadTitulo, 2) 'quito la coma
                        cadTitulo = cadParam & cadTitulo
                        conn.Execute cadTitulo
                        cadTitulo = ""
                    End If
                End If
            End If
            
            YaEstaFijadoPrecio = False
            
            '(codusu,codigo1,campo1,nombre1,nombre2,importeb1,importeb2
            '                 codpro artic   nomart   preciopmp  calculado
            NumRegElim = NumRegElim + 1
            cadNomRPT = "(" & vUsu.Codigo & "," & NumRegElim & "," & miRsAux!codProve & "," & miRsAux!codfamia & "," & DBSet(miRsAux!codArtic, "T")
            cadNomRPT = cadNomRPT & "," & DBSet(miRsAux!NomArtic, "T") & "," & DBSet(miRsAux!PrecioMP, "N") & ","
            
            
            'NUEVO ARTICULO

            If miRsAux!detamovi = "ALC" Then
                'If miRsAux!Cantidad = 0 Then Stop
                existencia = miRsAux!Cantidad
                PrecioMP = miRsAux!impormov / existencia
                YaEstaFijadoPrecio = True
            Else
                'El primer movimiento NO es una venta
                
            End If
        
        
            codArtic = miRsAux!codArtic
        
        Else
            If miRsAux!detamovi = "ALC" Then
                'NO es nuevo. otra linea
                If existencia < 0 Then
                    'No habia nada
                    existencia = existencia + miRsAux!Cantidad
                    PrecioMP = miRsAux!impormov / miRsAux!Cantidad
    
                
                Else
                    If (existencia + miRsAux!Cantidad) <> 0 Then
                        PrecioMP = Round2(((existencia * PrecioMP) + miRsAux!impormov) / (existencia + miRsAux!Cantidad), 4)
                        
                    Else
                        'Stop
                    End If
                End If
                YaEstaFijadoPrecio = True
            Else
                'Es otro tipo de moviemiento
                If YaEstaFijadoPrecio Then
                    'Si no esta fijado el precio NO hago nada
                    Aux = miRsAux!Cantidad
                    If miRsAux!tipomovi = 0 Then Aux = -Aux
                    existencia = existencia + Aux
                Else
                    
                    
                End If
                
            End If
        
        End If
        
        miRsAux.MoveNext
        
    Wend
    miRsAux.Close
    
    
    If cadNomRPT <> "" Then
        If YaEstaFijadoPrecio Then
            cadNomRPT = cadNomRPT & DBSet(PrecioMP, "N") & ")" 'metemos el calulado
            cadTitulo = cadTitulo & ", " & cadNomRPT
        End If
        
        'INSERTO TMP
        If cadTitulo <> "" Then
            cadTitulo = Mid(cadTitulo, 2) 'quito la coma
            cadTitulo = cadParam & cadTitulo
            conn.Execute cadTitulo
        End If
    End If


    'Ajuste final. Borro los que el precio sea igual
    Label3(41).Caption = "Ajuste datos"
    Label3(41).Refresh
    miSQL = "DELETE from tmpinformes WHERE codusu = " & vUsu.Codigo & " AND precio1 = precio2 "
    conn.Execute miSQL
    
    miSQL = "Select count(*) from tmpinformes WHERE codusu = " & vUsu.Codigo
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If DBLet(miRsAux.Fields(0), "N") > 0 Then RecalculoPrecioMedioPonderado = True
    End If
    miRsAux.Close

eRecalculoPrecioMedioPonderado:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set miRsAux = Nothing
    Label3(41).Caption = ""
End Function

'
Private Function RellenarTablaTemporalAcum(vWhere As String) As Boolean
Dim Sql As String
Dim SQL1 As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim SqlValues As String

Dim CliSocAnt As Integer
Dim ClienteAnt As Long
Dim FamiliaAnt As Long
Dim ArticAnt As String
Dim FechaAnt As Date
Dim ImporteAnt As Currency
Dim TipoMAnt As Integer
Dim AlbarAnt As Long
Dim DetaMovAnt As String

Dim Acumulado As Long
Dim Importe As Currency
Dim FechaIni As Date
Dim Precio As Currency
Dim Entro As Boolean
    On Error GoTo eRellenarTablaTemporalAcum
    
    RellenarTablaTemporalAcum = False

    Label3(13).visible = True
    Screen.MousePointer = vbHourglass
    DoEvents


    Sql = "select * from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
    Sql = Sql & " order by campo2, codigo1, importe4, nombre2, fecha1, importe3 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        CliSocAnt = DBLet(Rs!campo2, "N")
        ClienteAnt = DBLet(Rs!Codigo1, "N")
        FamiliaAnt = DBLet(Rs!importe4, "N")
        ArticAnt = DBLet(Rs!Nombre2, "T")
        FechaAnt = DBLet(Rs!fecha1, "F")
        ImporteAnt = DBLet(Rs!importe2, "N")
        TipoMAnt = DBLet(Rs!campo1, "N")
        AlbarAnt = DBLet(Rs!importe3, "N")
        DetaMovAnt = DBLet(Rs!nombre1, "T")
        
'         Insertar movimiento de inicio
        Sql2 = "select sum(if(tipomovi = 1, cantidad,0)) - sum(if(tipomovi = 0,cantidad,0)) "
        Sql2 = Sql2 & " from smoval where codigope = " & DBSet(ClienteAnt, "N")
        Sql2 = Sql2 & " and codartic = " & DBSet(ArticAnt, "T")
        Sql2 = Sql2 & " and detamovi = " & DBSet(Rs!nombre1, "T")
        Sql2 = Sql2 & " and fechamov < " & DBSet(Rs!fecha1, "F")
        
        If txtCodigo(9).Text <> "" Then Sql2 = Sql2 & " and fechamov < " & DBSet(txtCodigo(9).Text, "F")
        

        Acumulado = DevuelveValor(Sql2)
        Precio = PonerPrecio(ArticAnt, CStr(ClienteAnt))
        Importe = Round2(Acumulado * Precio, 2)

        If txtCodigo(9).Text <> "" Then
            FechaIni = DateAdd("d", -1, CDate(txtCodigo(9).Text))
        Else
            FechaIni = DateAdd("d", -1, FechaAnt)
        End If

        Sql2 = "insert into tmpinformes (codusu, campo2, codigo1, importe4, nombre2, fecha1, nombre1, importeb1, importe2, precio1) values ("
        Sql2 = Sql2 & vUsu.Codigo & "," & DBSet(CliSocAnt, "N") & "," & DBSet(ClienteAnt, "N") & ","
        Sql2 = Sql2 & DBSet(FamiliaAnt, "N") & "," & DBSet(ArticAnt, "T") & "," & DBSet(FechaIni, "F") & ","
        Sql2 = Sql2 & DBSet(Rs!nombre1, "T") & ","
        Sql2 = Sql2 & DBSet(Acumulado, "N") & ","
        Sql2 = Sql2 & DBSet(Importe, "N") & "," & DBSet(Precio, "N") & ")"

        conn.Execute Sql2
        
        FechaIni = FechaIni + 1
        
        While FechaIni < DBLet(Rs!fecha1) And Precio <> 0
            Sql2 = "insert into tmpinformes (codusu, campo2, codigo1, importe4, nombre2, nombre1, fecha1,importeb1, importe2, precio1) values ("
            Sql2 = Sql2 & vUsu.Codigo & "," & DBSet(CliSocAnt, "N") & "," & DBSet(ClienteAnt, "N") & ","
            Sql2 = Sql2 & DBSet(FamiliaAnt, "N") & "," & DBSet(ArticAnt, "T") & "," & DBSet(DetaMovAnt, "T") & ","
            Sql2 = Sql2 & DBSet(FechaIni, "F") & "," & DBSet(Acumulado, "N") & ","
            Sql2 = Sql2 & DBSet(Importe, "N") & "," & DBSet(Precio, "N") & ")"

            conn.Execute Sql2

            FechaIni = FechaIni + 1
        Wend
        
        
    End If
    
    
    While Not Rs.EOF
        If CliSocAnt <> DBLet(Rs!campo2, "N") Or ClienteAnt <> DBLet(Rs!Codigo1, "N") Or FamiliaAnt <> DBLet(Rs!importe4, "N") Or _
           Trim(ArticAnt) <> DBLet(Rs!Nombre2, "T") Then
            
            ' falta insertar los restantes movimientos hasta la fecha fin
            If txtCodigo(10).Text <> "" Then
                FechaAnt = FechaAnt + 1
                While FechaAnt <= txtCodigo(10).Text And Precio <> 0
                    Sql2 = "insert into tmpinformes (codusu, campo2, codigo1, importe4, nombre2, fecha1, nombre1, importeb1, importe2, precio1) values ("
                    Sql2 = Sql2 & vUsu.Codigo & "," & DBSet(CliSocAnt, "N") & "," & DBSet(ClienteAnt, "N") & ","
                    Sql2 = Sql2 & DBSet(FamiliaAnt, "N") & "," & DBSet(ArticAnt, "T") & ","
                    Sql2 = Sql2 & DBSet(FechaAnt, "F") & "," & DBSet(DetaMovAnt, "T") & "," & DBSet(Acumulado, "N") & ","
                    Sql2 = Sql2 & DBSet(Importe, "N") & "," & DBSet(Precio, "N") & ")"
            
                    conn.Execute Sql2
                    
                    FechaAnt = FechaAnt + 1
                Wend
            End If
            
            ' Inicializamos variables
            CliSocAnt = DBLet(Rs!campo2, "N")
            ClienteAnt = DBLet(Rs!Codigo1, "N")
            FamiliaAnt = DBLet(Rs!importe4, "N")
            ArticAnt = DBLet(Rs!Nombre2, "T")
            FechaAnt = DBLet(Rs!fecha1, "F")
            DetaMovAnt = DBLet(Rs!nombre1, "T")
            
            'Insertar movimiento de inicio
            Sql2 = "select sum(if(tipomovi = 1, cantidad,0)) - sum(if(tipomovi = 0,cantidad,0)) "
            Sql2 = Sql2 & " from smoval where codigope = " & DBSet(ClienteAnt, "N")
            Sql2 = Sql2 & " and codartic = " & DBSet(ArticAnt, "T")
            Sql2 = Sql2 & " and detamovi = " & DBSet(Rs!nombre1, "T")
            Sql2 = Sql2 & " and fechamov < " & DBSet(Rs!fecha1, "F")
            If txtCodigo(9).Text <> "" Then Sql2 = Sql2 & " and fechamov < " & DBSet(txtCodigo(9).Text, "F")
    
            Acumulado = DevuelveValor(Sql2)
            Precio = PonerPrecio(ArticAnt, CStr(ClienteAnt))
            Importe = Round2(Acumulado * Precio, 2)
    
            If txtCodigo(9).Text <> "" Then
                FechaIni = DateAdd("d", -1, CDate(txtCodigo(9).Text))
            Else
                FechaIni = DateAdd("d", -1, FechaAnt)
            End If
    
            Sql2 = "insert into tmpinformes (codusu, campo2, codigo1, importe4, nombre2, nombre1, fecha1,importeb1, importe2, precio1) values ("
            Sql2 = Sql2 & vUsu.Codigo & "," & DBSet(CliSocAnt, "N") & "," & DBSet(ClienteAnt, "N") & ","
            Sql2 = Sql2 & DBSet(FamiliaAnt, "N") & "," & DBSet(ArticAnt, "T") & "," & DBSet(DetaMovAnt, "T") & ","
            Sql2 = Sql2 & DBSet(FechaIni, "F") & "," & DBSet(Acumulado, "N") & ","
            Sql2 = Sql2 & DBSet(Importe, "N") & "," & DBSet(Precio, "N") & ")"
    
            conn.Execute Sql2
            
            FechaIni = FechaIni + 1
            While FechaIni < DBLet(Rs!fecha1)
                If Precio <> 0 Then
                    Sql2 = "insert into tmpinformes (codusu, campo2, codigo1, importe4, nombre2, nombre1, fecha1,importeb1, importe2, precio1) values ("
                    Sql2 = Sql2 & vUsu.Codigo & "," & DBSet(CliSocAnt, "N") & "," & DBSet(ClienteAnt, "N") & ","
                    Sql2 = Sql2 & DBSet(FamiliaAnt, "N") & "," & DBSet(ArticAnt, "T") & "," & DBSet(DetaMovAnt, "T") & ","
                    Sql2 = Sql2 & DBSet(FechaIni, "F") & "," & DBSet(Acumulado, "N") & ","
                    Sql2 = Sql2 & DBSet(Importe, "N") & "," & DBSet(Precio, "N") & ")"
    
                    conn.Execute Sql2
                End If
                FechaIni = FechaIni + 1
            Wend
        End If
    
    
        If FechaIni = DBLet(Rs!fecha1) Then
            ' modificamos unicamente el acumulado y el importe
            If DBLet(Rs!campo1) = 0 Then ' salida
                Acumulado = Acumulado - DBLet(Rs!importe1)
            Else ' entrada
                Acumulado = Acumulado + DBLet(Rs!importe1)
            End If
            
            Precio = PonerPrecio(ArticAnt, CStr(Rs!Codigo1))
            Importe = Round2(Acumulado * Precio, 2) ' ClienteAnt)), 2)
            
            Sql2 = "update tmpinformes set importeb1 = " & DBSet(Acumulado, "N")
            Sql2 = Sql2 & ", importe2 = " & DBSet(Importe, "N")
            Sql2 = Sql2 & ", precio1 = " & DBSet(Precio, "N")
            Sql2 = Sql2 & " where codusu = " & vUsu.Codigo
            Sql2 = Sql2 & " and fecha1 = " & DBSet(FechaIni, "F")
            Sql2 = Sql2 & " and campo2 = " & DBSet(Rs!campo2, "N") 'DBSet(CliSocAnt, "N")
            Sql2 = Sql2 & " and codigo1 = " & DBSet(Rs!Codigo1, "N") 'DBSet(ClienteAnt, "N")
            Sql2 = Sql2 & " and importe4 = " & DBSet(Rs!importe4, "N") 'DBSet(FamiliaAnt, "N")
            Sql2 = Sql2 & " and nombre2 = " & DBSet(Rs!Nombre2, "T") 'DBSet(ArticAnt, "T")
            Sql2 = Sql2 & " and campo1 = " & DBSet(Rs!campo1, "N") 'DBSet(TipoMAnt, "N")
            Sql2 = Sql2 & " and importe3 = " & DBSet(Rs!importe3, "N") 'DBSet(AlbarAnt, "N")
        
            conn.Execute Sql2
            
            FechaIni = FechaIni + 1
        Else
            FechaIni = FechaIni + 1
        End If

        Entro = False

        While FechaIni < DBLet(Rs!fecha1)
            ' solo si tiene importe el movimiento
            If DBLet(Importe, "N") <> 0 Then
                Sql2 = "insert into tmpinformes (codusu, campo2, codigo1, importe4, nombre2, fecha1, nombre1, importeb1, importe2, precio1) values ( "
                Sql2 = Sql2 & vUsu.Codigo & "," & DBSet(CliSocAnt, "N") & "," & DBSet(ClienteAnt, "N") & ","
                Sql2 = Sql2 & DBSet(FamiliaAnt, "N") & "," & DBSet(ArticAnt, "T") & ","
                Sql2 = Sql2 & DBSet(FechaIni, "F") & "," & DBSet(Rs!nombre1, "F") & "," & DBSet(Acumulado, "N") & ","
                Sql2 = Sql2 & DBSet(Importe, "N") & "," & DBSet(Precio, "N") & ")"
                
                conn.Execute Sql2
            End If
            
            FechaIni = FechaIni + 1
            Entro = True
        Wend
        
        If Entro Then
            If FechaIni = DBLet(Rs!fecha1) Then
                ' modificamos unicamente el acumulado y el importe
                If DBLet(Rs!campo1) = 0 Then ' salida
                    Acumulado = Acumulado - DBLet(Rs!importe1)
                Else ' entrada
                    Acumulado = Acumulado + DBLet(Rs!importe1)
                End If
                
                Precio = PonerPrecio(ArticAnt, CStr(Rs!Codigo1))
                Importe = Round2(Acumulado * Precio, 2) ' ClienteAnt)), 2)
                
                Sql2 = "update tmpinformes set importeb1 = " & DBSet(Acumulado, "N")
                Sql2 = Sql2 & ", importe2 = " & DBSet(Importe, "N")
                Sql2 = Sql2 & ", precio1 = " & DBSet(Precio, "N")
                Sql2 = Sql2 & " where codusu = " & vUsu.Codigo
                Sql2 = Sql2 & " and fecha1 = " & DBSet(FechaIni, "F")
                Sql2 = Sql2 & " and campo2 = " & DBSet(Rs!campo2, "N") 'DBSet(CliSocAnt, "N")
                Sql2 = Sql2 & " and codigo1 = " & DBSet(Rs!Codigo1, "N") 'DBSet(ClienteAnt, "N")
                Sql2 = Sql2 & " and importe4 = " & DBSet(Rs!importe4, "N") 'DBSet(FamiliaAnt, "N")
                Sql2 = Sql2 & " and nombre2 = " & DBSet(Rs!Nombre2, "T") 'DBSet(ArticAnt, "T")
                Sql2 = Sql2 & " and campo1 = " & DBSet(Rs!campo1, "N") 'DBSet(TipoMAnt, "N")
                Sql2 = Sql2 & " and importe3 = " & DBSet(Rs!importe3, "N") 'DBSet(AlbarAnt, "N")
            
                conn.Execute Sql2
                
            End If
        End If
        
        FechaAnt = Rs!fecha1
        ImporteAnt = Rs!importe1
        TipoMAnt = Rs!campo1
        AlbarAnt = Rs!importe3
        
        FechaIni = FechaAnt
        
        Rs.MoveNext
    Wend
    
    If txtCodigo(10).Text <> "" Then
        FechaAnt = FechaAnt + 1
        While FechaAnt <= txtCodigo(10).Text And Importe <> 0
            Sql2 = "insert into tmpinformes (codusu, campo2, codigo1, importe4, nombre2, nombre1, fecha1,importeb1, importe2, precio1) values ("
            Sql2 = Sql2 & vUsu.Codigo & "," & DBSet(CliSocAnt, "N") & "," & DBSet(ClienteAnt, "N") & ","
            Sql2 = Sql2 & DBSet(FamiliaAnt, "N") & "," & DBSet(ArticAnt, "T") & "," & DBSet(DetaMovAnt, "T") & ","
            Sql2 = Sql2 & DBSet(FechaAnt, "F") & "," & DBSet(Acumulado, "N") & ","
            Sql2 = Sql2 & DBSet(Importe, "N") & "," & DBSet(Precio, "N") & ")"
    
            conn.Execute Sql2
            
            FechaAnt = FechaAnt + 1
        Wend
    End If
        
    Set Rs = Nothing
    
    Screen.MousePointer = vbDefault
    Label3(13).visible = False
    DoEvents
    
    RellenarTablaTemporalAcum = True
    Exit Function
    
eRellenarTablaTemporalAcum:
    Screen.MousePointer = vbDefault
    Label3(13).visible = False
    DoEvents
    MuestraError Err.Number, "Rellenar Tabla Temporal"
End Function

Private Function PonerPrecio(Articulo As String, Cliente As String) As Currency
Dim Sql As String
Dim Precio As Currency

    On Error Resume Next

    Sql = "select precioar from clientes_precio where codclien =  " & DBSet(Cliente, "N")
    Sql = Sql & " and codartic = " & DBSet(Articulo, "T")
    
    If TotalRegistrosConsulta(Sql) = 0 Then
        Sql = "select preciove from sartic where codartic = " & DBSet(Articulo, "T")
    End If
    
    Precio = DevuelveValor(Sql)
    
    PonerPrecio = Precio

End Function



Private Function RellenarTablaTemporalAcum2(vWhere As String) As Boolean
Dim Sql As String
Dim SQL1 As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim SqlValues As String

Dim CliSocAnt As Integer
Dim ClienteAnt As Long
Dim FamiliaAnt As Long
Dim ArticAnt As String
Dim FechaAnt As Date
Dim ImporteAnt As Currency
Dim TipoMAnt As Integer
Dim AlbarAnt As Long
Dim DetaMovAnt As String

Dim Acumulado As Long
Dim Importe As Currency
Dim FechaIni As Date
Dim Precio As Currency
Dim Entro As Boolean
    On Error GoTo eRellenarTablaTemporalAcum
    
    RellenarTablaTemporalAcum2 = False

    Label3(13).visible = True
    Screen.MousePointer = vbHourglass
    DoEvents


    Sql = "select * from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
    Sql = Sql & " order by campo2, codigo1, importe4, nombre2, fecha1, importe3 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        CliSocAnt = DBLet(Rs!campo2, "N")
        ClienteAnt = DBLet(Rs!Codigo1, "N")
        FamiliaAnt = DBLet(Rs!importe4, "N")
        ArticAnt = DBLet(Rs!Nombre2, "T")
        FechaAnt = DBLet(Rs!fecha1, "F")
        ImporteAnt = DBLet(Rs!importe2, "N")
        TipoMAnt = DBLet(Rs!campo1, "N")
        AlbarAnt = DBLet(Rs!importe3, "N")
        DetaMovAnt = DBLet(Rs!nombre1, "T")
        
'         Insertar movimiento de inicio
        Sql2 = "select sum(if(tipomovi = 1, cantidad,0)) - sum(if(tipomovi = 0,cantidad,0)) "
        Sql2 = Sql2 & " from smoval where codigope = " & DBSet(ClienteAnt, "N")
        Sql2 = Sql2 & " and codartic = " & DBSet(ArticAnt, "T")
        Sql2 = Sql2 & " and detamovi = " & DBSet(Rs!nombre1, "T")
        Sql2 = Sql2 & " and fechamov < " & DBSet(Rs!fecha1, "F")
        
        If txtCodigo(9).Text <> "" Then Sql2 = Sql2 & " and fechamov < " & DBSet(txtCodigo(9).Text, "F")
        

        Acumulado = DevuelveValor(Sql2)
        Precio = PonerPrecio(ArticAnt, CStr(ClienteAnt))
        Importe = Round2(Acumulado * Precio, 2)

        If txtCodigo(9).Text <> "" Then
            FechaIni = DateAdd("d", -1, CDate(txtCodigo(9).Text))
        Else
            FechaIni = DateAdd("d", -1, FechaAnt)
        End If

        Sql2 = "insert into tmpinformes (codusu, campo2, codigo1, importe4, nombre2, fecha1, nombre1, importeb1, importe2, precio1, importe3) values ("
        Sql2 = Sql2 & vUsu.Codigo & "," & DBSet(CliSocAnt, "N") & "," & DBSet(ClienteAnt, "N") & ","
        Sql2 = Sql2 & DBSet(FamiliaAnt, "N") & "," & DBSet(ArticAnt, "T") & "," & DBSet(FechaIni, "F") & ","
        Sql2 = Sql2 & DBSet(Rs!nombre1, "T") & ","
        Sql2 = Sql2 & DBSet(Acumulado, "N") & ","
        Sql2 = Sql2 & DBSet(Importe, "N") & "," & DBSet(Precio, "N") & ",0)"

        conn.Execute Sql2
        
        FechaIni = FechaIni + 1
        
        While FechaIni < DBLet(Rs!fecha1)
            If Precio <> 0 Then
                Sql2 = "insert into tmpinformes (codusu, campo2, codigo1, importe4, nombre2, nombre1, fecha1,importeb1, importe2, precio1, importe3) values ("
                Sql2 = Sql2 & vUsu.Codigo & "," & DBSet(CliSocAnt, "N") & "," & DBSet(ClienteAnt, "N") & ","
                Sql2 = Sql2 & DBSet(FamiliaAnt, "N") & "," & DBSet(ArticAnt, "T") & "," & DBSet(DetaMovAnt, "T") & ","
                Sql2 = Sql2 & DBSet(FechaIni, "F") & "," & DBSet(Acumulado, "N") & ","
                Sql2 = Sql2 & DBSet(Importe, "N") & "," & DBSet(Precio, "N") & ",0)"
    
                conn.Execute Sql2
            End If
            FechaIni = FechaIni + 1
        Wend
    End If
    
    While Not Rs.EOF
        If CliSocAnt <> DBLet(Rs!campo2, "N") Or ClienteAnt <> DBLet(Rs!Codigo1, "N") Or FamiliaAnt <> DBLet(Rs!importe4, "N") Or _
           Trim(ArticAnt) <> DBLet(Rs!Nombre2, "T") Then
            
            ' falta insertar los restantes movimientos hasta la fecha fin
            If txtCodigo(10).Text <> "" Then
                While FechaIni <= txtCodigo(10).Text
                    Sql2 = "insert into tmpinformes (codusu, campo2, codigo1, importe4, nombre2, fecha1, nombre1, importeb1, importe2, precio1, importe3) values ("
                    Sql2 = Sql2 & vUsu.Codigo & "," & DBSet(CliSocAnt, "N") & "," & DBSet(ClienteAnt, "N") & ","
                    Sql2 = Sql2 & DBSet(FamiliaAnt, "N") & "," & DBSet(ArticAnt, "T") & ","
                    Sql2 = Sql2 & DBSet(FechaIni, "F") & "," & DBSet(DetaMovAnt, "T") & "," & DBSet(Acumulado, "N") & ","
                    Sql2 = Sql2 & DBSet(Importe, "N") & "," & DBSet(Precio, "N") & ",0)"
            
                    conn.Execute Sql2
                    
                    FechaIni = FechaIni + 1
                Wend
            End If
            
            ' Inicializamos variables
            CliSocAnt = DBLet(Rs!campo2, "N")
            ClienteAnt = DBLet(Rs!Codigo1, "N")
            FamiliaAnt = DBLet(Rs!importe4, "N")
            ArticAnt = DBLet(Rs!Nombre2, "T")
            DetaMovAnt = DBLet(Rs!nombre1, "T")
            
            'Insertar movimiento de inicio
            Sql2 = "select sum(if(tipomovi = 1, cantidad,0)) - sum(if(tipomovi = 0,cantidad,0)) "
            Sql2 = Sql2 & " from smoval where codigope = " & DBSet(ClienteAnt, "N")
            Sql2 = Sql2 & " and codartic = " & DBSet(ArticAnt, "T")
            Sql2 = Sql2 & " and detamovi = " & DBSet(Rs!nombre1, "T")
            Sql2 = Sql2 & " and fechamov < " & DBSet(Rs!fecha1, "F")
            If txtCodigo(9).Text <> "" Then Sql2 = Sql2 & " and fechamov < " & DBSet(txtCodigo(9).Text, "F")
    
            Acumulado = DevuelveValor(Sql2)
            Precio = PonerPrecio(ArticAnt, CStr(ClienteAnt))
            Importe = Round2(Acumulado * Precio, 2)
    
            If txtCodigo(9).Text <> "" Then
                FechaIni = DateAdd("d", -1, CDate(txtCodigo(9).Text))
            Else
                FechaIni = DateAdd("d", -1, FechaAnt)
            End If
    
            Sql2 = "insert into tmpinformes (codusu, campo2, codigo1, importe4, nombre2, nombre1, fecha1,importeb1, importe2, precio1, importe3) values ("
            Sql2 = Sql2 & vUsu.Codigo & "," & DBSet(CliSocAnt, "N") & "," & DBSet(ClienteAnt, "N") & ","
            Sql2 = Sql2 & DBSet(FamiliaAnt, "N") & "," & DBSet(ArticAnt, "T") & "," & DBSet(DetaMovAnt, "T") & ","
            Sql2 = Sql2 & DBSet(FechaIni, "F") & "," & DBSet(Acumulado, "N") & ","
            Sql2 = Sql2 & DBSet(Importe, "N") & "," & DBSet(Precio, "N") & ",0)"
    
            conn.Execute Sql2
            
            FechaIni = FechaIni + 1
            While FechaIni < DBLet(Rs!fecha1)
                If Precio <> 0 Then
                    Sql2 = "insert into tmpinformes (codusu, campo2, codigo1, importe4, nombre2, nombre1, fecha1,importeb1, importe2, precio1, importe3) values ("
                    Sql2 = Sql2 & vUsu.Codigo & "," & DBSet(CliSocAnt, "N") & "," & DBSet(ClienteAnt, "N") & ","
                    Sql2 = Sql2 & DBSet(FamiliaAnt, "N") & "," & DBSet(ArticAnt, "T") & "," & DBSet(DetaMovAnt, "T") & ","
                    Sql2 = Sql2 & DBSet(FechaIni, "F") & "," & DBSet(Acumulado, "N") & ","
                    Sql2 = Sql2 & DBSet(Importe, "N") & "," & DBSet(Precio, "N") & ",0)"
    
                    conn.Execute Sql2
                End If
                FechaIni = FechaIni + 1
            Wend
            
        End If
    

        If FechaIni = DBLet(Rs!fecha1) Then
            Rs.MoveNext
            If Not Rs.EOF Then
                If FechaIni <> DBLet(Rs!fecha1) Then FechaIni = FechaIni + 1
            End If
        Else
            While FechaIni < DBLet(Rs!fecha1)
                Sql2 = "insert into tmpinformes (codusu, campo2, codigo1, importe4, nombre2, fecha1, nombre1, importeb1, importe2, precio1, importe3) values ( "
                Sql2 = Sql2 & vUsu.Codigo & "," & DBSet(CliSocAnt, "N") & "," & DBSet(ClienteAnt, "N") & ","
                Sql2 = Sql2 & DBSet(FamiliaAnt, "N") & "," & DBSet(ArticAnt, "T") & ","
                Sql2 = Sql2 & DBSet(FechaIni, "F") & "," & DBSet(Rs!nombre1, "F") & "," & DBSet(Acumulado, "N") & ","
                Sql2 = Sql2 & DBSet(Importe, "N") & "," & DBSet(Precio, "N") & ",0)"
                
                conn.Execute Sql2
                
                FechaIni = FechaIni + 1
            Wend
            FechaIni = DBLet(Rs!fecha1)
           
        End If
       
    Wend
    
    If txtCodigo(10).Text <> "" Then
        FechaIni = FechaIni + 1
        While FechaIni <= txtCodigo(10).Text And Importe <> 0
            Sql2 = "insert into tmpinformes (codusu, campo2, codigo1, importe4, nombre2, nombre1, fecha1,importeb1, importe2, precio1, importe3) values ("
            Sql2 = Sql2 & vUsu.Codigo & "," & DBSet(CliSocAnt, "N") & "," & DBSet(ClienteAnt, "N") & ","
            Sql2 = Sql2 & DBSet(FamiliaAnt, "N") & "," & DBSet(ArticAnt, "T") & "," & DBSet(DetaMovAnt, "T") & ","
            Sql2 = Sql2 & DBSet(FechaIni, "F") & "," & DBSet(Acumulado, "N") & ","
            Sql2 = Sql2 & DBSet(Importe, "N") & "," & DBSet(Precio, "N") & ",0)"
    
            conn.Execute Sql2
            
            FechaIni = FechaIni + 1
        Wend
    End If
        
    Set Rs = Nothing
    
    ' recorremos el recordset unicamente con el calculo del acumulado
    Sql = "select * from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
    Sql = Sql & " order by campo2, codigo1, importe4, nombre2, fecha1, importe3 "
        
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        CliSocAnt = DBLet(Rs!campo2, "N")
        ClienteAnt = DBLet(Rs!Codigo1, "N")
        FamiliaAnt = DBLet(Rs!importe4, "N")
        ArticAnt = DBLet(Rs!Nombre2, "T")
    End If
    
    While Not Rs.EOF
        If CliSocAnt <> DBLet(Rs!campo2, "N") Or ClienteAnt <> DBLet(Rs!Codigo1, "N") Or FamiliaAnt <> DBLet(Rs!importe4, "N") Or _
           Trim(ArticAnt) <> DBLet(Rs!Nombre2, "T") Then
            
            Acumulado = DBLet(Rs!importeb1)
            
            CliSocAnt = DBLet(Rs!campo2, "N")
            ClienteAnt = DBLet(Rs!Codigo1, "N")
            FamiliaAnt = DBLet(Rs!importe4, "N")
            ArticAnt = DBLet(Rs!Nombre2, "T")
        
        End If
        
        If DBLet(Rs!importe1) <> 0 And Not IsNull(Rs!importe1) Then
            If DBLet(Rs!campo1) = 0 Then ' salida
                Acumulado = Acumulado - DBLet(Rs!importe1)
            Else ' entrada
                Acumulado = Acumulado + DBLet(Rs!importe1)
            End If
        End If
            
        Precio = PonerPrecio(ArticAnt, CStr(Rs!Codigo1))
        Importe = Round2(Acumulado * Precio, 2) ' ClienteAnt)), 2)
        
        Sql2 = "update tmpinformes set importeb1 = " & DBSet(Acumulado, "N")
        Sql2 = Sql2 & ", importe2 = " & DBSet(Importe, "N")
        Sql2 = Sql2 & ", precio1 = " & DBSet(Precio, "N")
        Sql2 = Sql2 & " where codusu = " & vUsu.Codigo
        Sql2 = Sql2 & " and fecha1 = " & DBSet(Rs!fecha1, "F")
        Sql2 = Sql2 & " and campo2 = " & DBSet(Rs!campo2, "N") 'DBSet(CliSocAnt, "N")
        Sql2 = Sql2 & " and codigo1 = " & DBSet(Rs!Codigo1, "N") 'DBSet(ClienteAnt, "N")
        Sql2 = Sql2 & " and importe4 = " & DBSet(Rs!importe4, "N") 'DBSet(FamiliaAnt, "N")
        Sql2 = Sql2 & " and nombre2 = " & DBSet(Rs!Nombre2, "T") 'DBSet(ArticAnt, "T")
        Sql2 = Sql2 & " and importe3 = " & DBSet(Rs!importe3, "N") 'DBSet(AlbarAnt, "N")
    
        conn.Execute Sql2
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    Screen.MousePointer = vbDefault
    Label3(13).visible = False
    DoEvents
    
    RellenarTablaTemporalAcum2 = True
    Exit Function
    
eRellenarTablaTemporalAcum:
    Screen.MousePointer = vbDefault
    Label3(13).visible = False
    DoEvents
    MuestraError Err.Number, "Rellenar Tabla Temporal"
End Function





