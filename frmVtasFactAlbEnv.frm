VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmVtasFactAlbEnv 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   7530
   Icon            =   "frmVtasFactAlbEnv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameFacturar 
      Height          =   6285
      Left            =   30
      TabIndex        =   26
      Top             =   -30
      Width           =   7395
      Begin VB.Frame FrameProgress 
         Height          =   1050
         Left            =   300
         TabIndex        =   56
         Top             =   4980
         Visible         =   0   'False
         Width           =   4695
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   345
            Left            =   120
            TabIndex        =   57
            Top             =   600
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   609
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label lblProgess 
            Caption         =   "Iniciando el proceso ..."
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   59
            Top             =   350
            Width           =   4335
         End
         Begin VB.Label lblProgess 
            Caption         =   "Facturando:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   58
            Top             =   135
            Width           =   4215
         End
      End
      Begin VB.Frame Frame4 
         Height          =   4065
         Left            =   300
         TabIndex        =   38
         Top             =   780
         Width           =   6855
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   34
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   62
            Top             =   300
            Width           =   1215
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   42
            Left            =   2940
            Locked          =   -1  'True
            TabIndex        =   51
            Text            =   "Text5"
            Top             =   3210
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   42
            Left            =   2160
            MaxLength       =   6
            TabIndex        =   34
            Tag             =   "Forma Pago|N|N|0|999|scaalb|codforpa|000||"
            Top             =   3210
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   43
            Left            =   2940
            Locked          =   -1  'True
            TabIndex        =   50
            Text            =   "Text5"
            Top             =   3570
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   43
            Left            =   2160
            MaxLength       =   6
            TabIndex        =   35
            Tag             =   "Forma Pago|N|N|0|999|scaalb|codforpa|000||"
            Top             =   3570
            Width           =   735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   41
            Left            =   2160
            MaxLength       =   6
            TabIndex        =   33
            Tag             =   "Cod. Cliente|N|N|0|999999|scaalb|codclien|000000||"
            Top             =   2730
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   41
            Left            =   2940
            Locked          =   -1  'True
            TabIndex        =   46
            Text            =   "Text5"
            Top             =   2730
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   40
            Left            =   2160
            MaxLength       =   6
            TabIndex        =   32
            Tag             =   "Cod. Cliente|N|N|0|999999|scaalb|codclien|000000||"
            Top             =   2370
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   40
            Left            =   2940
            Locked          =   -1  'True
            TabIndex        =   45
            Text            =   "Text5"
            Top             =   2370
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   38
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   30
            Top             =   1650
            Width           =   1215
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   39
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   31
            Top             =   1980
            Width           =   1215
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   36
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   28
            Top             =   810
            Width           =   1215
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   37
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   29
            Top             =   1170
            Width           =   1215
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   14
            Left            =   1860
            Picture         =   "frmVtasFactAlbEnv.frx":000C
            ToolTipText     =   "Buscar fecha"
            Top             =   300
            Width           =   240
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Factura"
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
            Left            =   240
            TabIndex        =   63
            Top             =   270
            Width           =   1035
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   22
            Left            =   1860
            ToolTipText     =   "Buscar forma pago"
            Top             =   3210
            Width           =   240
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Forma pago"
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
            Index           =   3
            Left            =   240
            TabIndex        =   54
            Top             =   3000
            Width           =   855
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
            Index           =   48
            Left            =   1335
            TabIndex        =   53
            Top             =   3210
            Width           =   450
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   23
            Left            =   1860
            ToolTipText     =   "Buscar forma pago"
            Top             =   3570
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
            Index           =   49
            Left            =   1335
            TabIndex        =   52
            Top             =   3570
            Width           =   420
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   20
            Left            =   1860
            ToolTipText     =   "Buscar cliente"
            Top             =   2370
            Width           =   240
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   21
            Left            =   1860
            ToolTipText     =   "Buscar cliente"
            Top             =   2730
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
            Index           =   50
            Left            =   1335
            TabIndex        =   49
            Top             =   2730
            Width           =   420
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
            Index           =   51
            Left            =   1335
            TabIndex        =   48
            Top             =   2370
            Width           =   450
         End
         Begin VB.Label Label10 
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
            Index           =   2
            Left            =   240
            TabIndex        =   47
            Top             =   2220
            Width           =   495
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
            Index           =   37
            Left            =   1350
            TabIndex        =   44
            Top             =   1980
            Width           =   420
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   12
            Left            =   1860
            Picture         =   "frmVtasFactAlbEnv.frx":0097
            ToolTipText     =   "Buscar fecha"
            Top             =   1665
            Width           =   240
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Albaran"
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
            TabIndex        =   43
            Top             =   1440
            Width           =   1035
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
            Index           =   46
            Left            =   1350
            TabIndex        =   42
            Top             =   1650
            Width           =   450
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   13
            Left            =   1860
            Picture         =   "frmVtasFactAlbEnv.frx":0122
            ToolTipText     =   "Buscar fecha"
            Top             =   1995
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
            Index           =   36
            Left            =   1380
            TabIndex        =   41
            Top             =   1170
            Width           =   420
         End
         Begin VB.Label Label10 
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
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   40
            Top             =   600
            Width           =   780
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
            Index           =   45
            Left            =   1380
            TabIndex        =   39
            Top             =   810
            Width           =   450
         End
      End
      Begin VB.CommandButton cmdAceptarFac 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5220
         TabIndex        =   36
         Top             =   5670
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   6
         Left            =   6300
         TabIndex        =   37
         Top             =   5670
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Facturación de Albaranes Envases"
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
         Left            =   360
         TabIndex        =   27
         Top             =   240
         Width           =   6615
      End
      Begin VB.Label Label10 
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
         Index           =   10
         Left            =   120
         TabIndex        =   61
         Top             =   3360
         Width           =   6855
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7800
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FramePreFacturar 
      Height          =   5775
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   7035
      Begin VB.CheckBox chkResumenForpa 
         Caption         =   "Resumen forma de pago"
         Height          =   195
         Left            =   3840
         TabIndex        =   60
         Top             =   4560
         Width           =   2295
      End
      Begin VB.CheckBox chkSoloFacturar 
         Caption         =   "Solo Albaranes para facturar"
         Height          =   375
         Left            =   3840
         TabIndex        =   11
         Top             =   4080
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.Frame Frame7 
         Caption         =   "Tipo Informe"
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
         Height          =   735
         Left            =   420
         TabIndex        =   55
         Top             =   4020
         Width           =   3135
         Begin VB.OptionButton OptDetalle 
            Caption         =   "Resumen"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   10
            Top             =   300
            Width           =   1335
         End
         Begin VB.OptionButton OptDetalle 
            Caption         =   "Detalle"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   9
            Top             =   300
            Width           =   1455
         End
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   26
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptarPreFac 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   12
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   5
         Left            =   5160
         TabIndex        =   13
         Top             =   5040
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   27
         Left            =   3900
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   30
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "Text5"
         Top             =   3240
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   30
         Left            =   1740
         MaxLength       =   3
         TabIndex        =   7
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   31
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   3600
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   31
         Left            =   1740
         MaxLength       =   3
         TabIndex        =   8
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   29
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   6
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   29
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "Text5"
         Top             =   2520
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   28
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   5
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   28
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "Text5"
         Top             =   2160
         Width           =   3615
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
         Index           =   44
         Left            =   3120
         TabIndex        =   25
         Top             =   1440
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   10
         Left            =   1440
         Picture         =   "frmVtasFactAlbEnv.frx":01AD
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Prefacturación de Albaranes"
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
         Left            =   360
         TabIndex        =   24
         Top             =   480
         Width           =   6375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Albaran"
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
         Index           =   43
         Left            =   480
         TabIndex        =   23
         Top             =   1200
         Width           =   1035
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
         Index           =   42
         Left            =   920
         TabIndex        =   22
         Top             =   1440
         Width           =   450
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   11
         Left            =   3600
         Picture         =   "frmVtasFactAlbEnv.frx":0238
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   16
         Left            =   1440
         Top             =   3260
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Formas de Pago"
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
         Index           =   41
         Left            =   480
         TabIndex        =   21
         Top             =   3000
         Width           =   1155
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
         Index           =   40
         Left            =   915
         TabIndex        =   20
         Top             =   3240
         Width           =   450
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   17
         Left            =   1440
         Top             =   3620
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
         Index           =   39
         Left            =   915
         TabIndex        =   19
         Top             =   3600
         Width           =   420
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
         Index           =   35
         Left            =   920
         TabIndex        =   18
         Top             =   2520
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   15
         Left            =   1440
         Top             =   2550
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
         Index           =   34
         Left            =   920
         TabIndex        =   17
         Top             =   2160
         Width           =   450
      End
      Begin VB.Label Label4 
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
         Index           =   33
         Left            =   480
         TabIndex        =   16
         Top             =   1920
         Width           =   495
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   14
         Left            =   1440
         Top             =   2160
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmVtasFactAlbEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DatoSeleccionado(CadenaSeleccion As String)

Public OpcionListado As Integer
'(ver opciones en frmListado)
      
      
      
'Alguna opcion mas
'                   1000.-  Es cuando paso pedido a albaran y este a factura en el mismo proceso
'                   1001.-  Facturar un unico albaran
      
      
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir


Public CodClien As String 'Para seleccionar inicialmente las ofertas del Proveedor

'#Laura 14/11/2006 Recuperar facturas Alzira
Public EstaRecupFact As Boolean ' si esta recuperando facturas (para albaranes de mostrador)


'Private HaDevueltoDatos As Boolean
Private NomTabla As String
Private NomTablaLin As String

'Private WithEvents frmMtoCartasOfe As frmFacCartasOferta
Private WithEvents frmCli As frmClientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmArt As frmManArtic
Attribute frmArt.VB_VarHelpID = -1
Private WithEvents frmFPago As frmManFpago
Attribute frmFPago.VB_VarHelpID = -1


'Private WithEvents frmB As frmBuscaGrid  'Busquedas
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1

'----- Variables para el INforme ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String
Private numParam As Byte
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private Titulo As String 'Titulo informe que se pasa a frmImprimir
Private nomRPT As String 'nombre del fichero .rpt a imprimir
Private conSubRPT As Boolean 'si tiene subinformes para enlazarlos a las tablas correctas
'-------------------------------------


Dim indCodigo As Integer 'indice para txtCodigo

Dim PrimeraVez As Boolean


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub





Private Sub chkSoloFacturar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub





Private Sub cmdAceptarFac_Click()
'Facturacion de Albaranes
Dim campo As String, cad As String
Dim cadFrom As String
Dim cadSQL As String 'Para seleccionar los Albaranes del rango seleccion
                      'que no se van a facturar
Dim CambiamosConta As Boolean
    
    InicializarVbles
    cadFrom = ""
    CambiamosConta = False
    '--- Comprobar q los campos tienen valor
    If Trim(txtCodigo(34).Text) = "" Then 'Fecha factura
        MsgBox "El campo Fecha Factura debe tener valor.", vbExclamation
        Exit Sub
    End If
    
    
    '--- Seleccinar los Albaranes que cumplen los criterios introducidos
    'Desde/Hasta Nº ALBARAN
    '-------------------------
    If txtCodigo(36).Text <> "" Or txtCodigo(37).Text <> "" Then
        campo = NomTabla & ".numalbar"
        cad = ""
        If Not PonerDesdeHasta(campo, "N", 36, 37, cad) Then Exit Sub
    End If

    'Desde/Hasta FECHA del ALBARAN
    '--------------------------------------------
    If txtCodigo(38).Text <> "" Or txtCodigo(39).Text <> "" Then
        'Para MySQL
        campo = "scaalb.fechaalb"
        cad = CadenaDesdeHastaBD(txtCodigo(38).Text, txtCodigo(39).Text, campo, "F")
        If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
        'Para Crystal Report
        campo = "{scaalb.fechaalb}"
        cad = "pDHFecha=""Fecha: "
        If Not PonerDesdeHasta(campo, "F", 38, 39, cad) Then Exit Sub
    End If

    'Cadena para seleccion D/H CLIENTE
    '----------------------------------------
    If txtCodigo(40).Text <> "" Or txtCodigo(41).Text <> "" Then
        campo = "scaalb.codclien"
        cad = ""
        If Not PonerDesdeHasta(campo, "N", 40, 41, cad) Then Exit Sub
    End If

    'Cadena para seleccion FORMA PAGO
    '--------------------------------------------
    If txtCodigo(42).Text <> "" Or txtCodigo(43).Text <> "" Then
        campo = "scaalb.codforpa"
        cad = " "
        If Not PonerDesdeHasta(campo, "N", 42, 43, cad) Then Exit Sub
    End If

    
    cadSQL = cadSelect
    'Seleccionar los Albaranes que tiene scaalb.factursn=1
    cad = " {scaalb.factursn=1} "
    If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
    AnyadirAFormula cadFormula, cad
    
    
    '--- Comprobar q se han Seleccionados registros de Albaran con esos criterios
    cad = "Select count(*) " ' & NomTabla & " INNER JOIN " & nomTablaLin
    If cadFrom = "" Then cadFrom = " scaalb inner join clientes on scaalb.codclien = clientes.codclien "
    cad = cad & " FROM " & cadFrom

    If Not HayRegParaInforme(cadFrom, cadSelect) Then Exit Sub
    
    'Verificar si con los criterios seleccionados (PARA VENTAS)
    'seleccionar si quedan en el rango Albaranes que no se van a Facturar
    'y mostrar mensaje
    
    'Seleccionar los Albaranes que tiene scaalb.factursn=0
    campo = " scaalb.factursn=0 "
    If Not AnyadirAFormula(cadSQL, campo) Then Exit Sub
    cadSQL = cad & " WHERE " & cadSQL
    If RegistrosAListar(cadSQL) > 0 Then
        'Mostrar los Albaranes que no se van a Facturar
        cadSQL = Replace(cadSQL, "count(*)", "scaalb.codtipom,scaalb.numalbar,scaalb.fechaalb,scaalb.codclien,clientes.nomclien")
        frmMensajes.OpcionMensaje = 12
        frmMensajes.cadWhere = cadSQL
        frmMensajes.Show vbModal
        If frmMensajes.vCampos = "0" Then Exit Sub
    End If
    
    cad = cad & " WHERE " & cadSelect
    'Pasar Albaranes a Facturas
    If InStr(cad, "clientes") <> 0 Then 'hay JOIN con sclien
        cad = Replace(cad, "count(*)", "scaalb.*, clientes.nomclien,clientes.domclien,clientes.codpobla,clientes.pobclien,clientes.proclien,clientes.cifclien,clientes.telclie1")
    Else
        cad = Replace(cad, "count(*)", "*")
    End If


    Me.Height = Me.Height + 300
    Me.FrameFacturar.Height = Me.FrameFacturar.Height + 300
    Me.FrameProgress.visible = True
'--monica
'    Me.FrameProgress.Top = 6250
    Me.ProgressBar1.Left = 200
    Me.ProgressBar1.Value = 0
    Me.lblProgess(1).Caption = "Inicializando el proceso..."
        
    'proceso normal
    Screen.MousePointer = vbHourglass
     
    '------------------------------------------------------------------------------
    '  LOG de acciones.
    Set LOG = New cLOG
    campo = "Albaran: " & CodClien & " : " & NumCod
    LOG.Insertar 2, vUsu, campo
    Set LOG = Nothing
    '-----------------------------------------------------------------------------

    campo = "" ' txtCSB(0).Text & "|" & txtCSB(1).Text & "|" & txtCSB(2).Text & "|"
    TraspasoAlbaranesFacturas cad, cadSelect, txtCodigo(34).Text, "", Me.ProgressBar1, Me.lblProgess(1), True, CodClien, campo

    Screen.MousePointer = vbDefault
    
    If CambiamosConta Then
       'Reestablecer la conexion con la antigua conta
'       AbrirConexionConta False
    End If
    Me.Height = Me.Height - 300
    Me.FrameFacturar.Height = Me.FrameFacturar.Height - 300
    Me.FrameProgress.visible = False
End Sub



'#### Laura 14/11/2006 Recuperar facturas ALZIRA
Private Function ComprobarCliente_RecuperarFac(cadSelAlb As String, fecFac As String, numFac As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim codMacta1 As String 'cliente factura ariges
Dim codMacta2 As String 'cliente factura conta
Dim LEtra As String

    On Error GoTo ErrCompCliente
    ComprobarCliente_RecuperarFac = False
    
    'codmacta del cliente del albaran a facturar en Ariges
    Sql = "select scaalb.codclien,sclien.codmacta"
    Sql = Sql & " from scaalb inner join sclien on scaalb.codclien=sclien.codclien "
    Sql = Sql & " Where " & cadSelAlb
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        codMacta1 = DBLet(Rs!Codmacta, "T")
    
    End If
    Set Rs = Nothing
    
    
    'codmacta en la contabilidad
    LEtra = ObtenerLetraSerie("FAV")
    Sql = "SELECT codmacta FROM cabfact "
    Sql = Sql & " WHERE numserie=" & DBSet(LEtra, "T") & " AND codfaccl=" & numFac & " AND anofaccl=" & Year(fecFac)
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        codMacta2 = DBLet(Rs!Codmacta, "T")
    End If
    Set Rs = Nothing
    
    If codMacta1 <> "" And codMacta2 <> "" Then
        If codMacta1 = codMacta2 Then
            ComprobarCliente_RecuperarFac = True
        Else
            ComprobarCliente_RecuperarFac = False
            MsgBox "La cuenta contable en la factura de Contabilidad no coincide con la del cliente del Albaran", vbExclamation
        End If
    Else
        ComprobarCliente_RecuperarFac = False
        MsgBox "No se ha podido leer la cuenta contable del cliente", vbExclamation
    End If
    
    Exit Function
    
ErrCompCliente:
    ComprobarCliente_RecuperarFac = False
    MuestraError Err.Number, "Comprobar cliente", Err.Description
End Function
'#####

Private Sub cmdAceptarPreFac_Click()
'Prevision de Facturacion de Albaranes
Dim campo As String, cad As String
Dim b As Boolean
Dim indice As Integer

    InicializarVbles
    b = (OpcionListado = 50)
    
    If (Not b) Or (b And (CodClien = "ALV" Or CodClien = "AL1")) Then
        'Pasar nombre de la Empresa como parametro
        cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
    End If
    
    
    'Desde/Hasta FECHA del ALBARAN
    '--------------------------------------------
    If Trim(txtCodigo(26).Text) <> "" Or Trim(txtCodigo(27).Text) <> "" Then
        If b And (CodClien <> "ALV" And CodClien <> "AL1") Then
            campo = "scaalb.fechaalb"
            cad = "FECHA: "
            cadFormula = CadenaDesdeHastaBD(txtCodigo(26).Text, txtCodigo(27).Text, campo, "F")
            cadParam = cadParam & AnyadirParametroDH(cad, 26, 27) & """|"
        Else
            'Para MySQL
            campo = "scaalb.fechaalb"
            cadSelect = CadenaDesdeHastaBD(txtCodigo(26).Text, txtCodigo(27).Text, campo, "F")
            'Para Crystal Report
            campo = "{scaalb.fechaalb}"
            cad = "pDHFecha=""Fecha: "
            If Not PonerDesdeHasta(campo, "F", 26, 27, cad) Then Exit Sub
        End If
    End If

    'Cadena para seleccion CLIENTE
    '--------------------------------------------
    If txtCodigo(28).Text <> "" Or txtCodigo(29).Text <> "" Then
        If b And (CodClien <> "ALV" And CodClien <> "AL1") Then
            campo = "scaalb.codclien"
            cad = "CLIENTE: "
        Else
            campo = "{scaalb.codclien}"
            cad = "pDHCliente=""Cliente: "
        End If
        If Not PonerDesdeHasta(campo, "N", 28, 29, cad) Then Exit Sub
    End If

    If b Then 'opcionlistado=50
        'Cadena para seleccion FORMA PAGO
        '--------------------------------------------
        If txtCodigo(30).Text <> "" Or txtCodigo(31).Text <> "" Then
            If b And (CodClien <> "ALV" And CodClien <> "AL1") Then
                campo = "scaalb.codforpa"
                cad = "F. PAGO: "
            Else
                campo = "{scaalb.codforpa}"
                cad = "pDHForpa=""Forma Pago: "
            End If
            If Not PonerDesdeHasta(campo, "N", 30, 31, cad) Then Exit Sub
        End If
        
        cad = " {scaalb.codtipom}='" & CodClien & "' "
        If Not AnyadirAFormula(cadFormula, cad) Then Exit Sub
        If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
        
        'Seleccionar los que esten marcados para facturar
        'Seleccionar solo aquellos que el campo scaalb.factursn=1
        If Me.chkSoloFacturar.Value = 1 Then
            cad = " {scaalb.factursn}=1 "
            If Not AnyadirAFormula(cadFormula, cad) Then Exit Sub
            If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
        End If
    End If
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If Not HayRegParaInforme("scaalb", cadSelect) Then Exit Sub
    
        
    If OpcionListado = 50 And (CodClien = "ALV" Or CodClien = "AL1") Then
    
        If chkResumenForpa.Value = 1 Then
            'VAMOS A MOSTRAR LA HOJA RESUMEN DE FORMAS DE PAGO
            conn.Execute "DELETE FROM tmpinformes where codusu =" & vUsu.codigo
        
            If Me.OptDetalle(0).Value Then
                Titulo = "SELECT scaalb.codforpa, sum(slialb.importel)," & vUsu.codigo
                Titulo = Titulo & " FROM   ((scaalb scaalb INNER JOIN sclien sclien ON scaalb.codclien=sclien.codclien) INNER JOIN slialb slialb ON (scaalb.codtipom=slialb.codtipom) AND (scaalb.numalbar=slialb.numalbar)) INNER JOIN starif starif ON sclien.codtarif=starif.codlista"
            
            Else
                Titulo = "SELECT  codforpa ,sum(slialb.importel)," & vUsu.codigo
                Titulo = Titulo & " FROM   slialb slialb INNER JOIN scaalb scaalb ON (slialb.codtipom=scaalb.codtipom) AND (slialb.numalbar=scaalb.numalbar)"
            End If
    
            If cadSelect <> "" Then Titulo = Titulo & " WHERE " & cadSelect
            
            Titulo = Titulo & " GROUP BY codforpa"
            Titulo = "INSERT INTO tmpinformes (codigo1,importe1,codusu) " & Titulo
            conn.Execute Titulo
        End If
    
    
        Titulo = "Previsión Facturación Ventas"
        '-- Si estan activos los servicios hay diferentes posibilidades y el título
        '   las refleja, la variabele 'indice' lleva la información del combo seleccionado y
        '   ha sido cargada un poco más arriba [SERVICIOS]
        conSubRPT = True
        If Me.OptDetalle(0).Value = True Then
            nomRPT = "rFacPrevFactDetalle.rpt"
        Else
            nomRPT = "rFacPrevFactResum.rpt"
        End If
        
        cad = "pCodUsu=" & vUsu.codigo & "|"
        cadParam = cadParam & cad
        numParam = numParam + 1
        
        '-- Ahora el título depende de los tipos de albaranes seleccionados [SERVICIOS]
        cad = "pTitulo=""" & Titulo & """|"
        cadParam = cadParam & cad
        numParam = numParam + 1
        
        
        '--  Mostrara , o no, el subreport con el resumen por forma pago
        cad = "pVerForpa=" & Abs(chkResumenForpa.Value) & "|"
        cadParam = cadParam & cad
        numParam = numParam + 1
        
        
        On Error GoTo EPreFact
        cad = "delete from tmpstockfec where codusu=" & vUsu.codigo
        conn.Execute cad
        
        'Insertar total bonificaciones por cliente,articulo en una TEMPORAL
        cad = "SELECT " & vUsu.codigo & " as codusu,  slialb.codartic,scaalb.codclien,sum(slialb.cantidad) as stock "
        cad = cad & "FROM (((scaalb INNER JOIN slialb ON scaalb.codtipom=slialb.codtipom and scaalb.numalbar=slialb.numalbar) "
        cad = cad & " INNER JOIN sbonif ON slialb.codartic=sbonif.codartic ) "
        cad = cad & " INNER JOIN sclien ON scaalb.codclien=sclien.codclien) "
        cad = cad & " INNER JOIN starif ON sclien.codtarif=starif.codlista "
        cad = cad & "WHERE " & cadSelect
        cad = cad & " AND starif.bonifica=1 "
        cad = cad & " GROUP BY scaalb.codclien,slialb.codartic"
        
        cad = "INSERT INTO tmpstockfec (codusu,codartic,codalmac,stock) " & cad
        conn.Execute cad
    End If
    
    If b And (CodClien <> "ALV" And CodClien <> "AL1") Then 'OpcionListado = 50 'NO Imprime, mostrar resultado en pantalla
        frmMensajes.cadWhere = cadSelect
        frmMensajes.vCampos = cadParam
        frmMensajes.OpcionMensaje = 6 'Prefacturacion Albaranes
        frmMensajes.Show vbModal
    Else
        LlamarImprimir
    End If
    
    If OpcionListado = 50 And (CodClien = "ALV" Or CodClien = "AL1") Then
        cad = "delete from tmpstockfec where codusu=" & vUsu.codigo
        conn.Execute cad
    End If
EPreFact:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Informe Prefacturación", Err.Description
    End If
End Sub



Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
     
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 50 '50: Prevision de Facturacion Albaranes (NO IMPRIME LISTADO)
                PonerFoco txtCodigo(26)
            Case 52 '52: Facturacion de Albaranes
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim I As Integer
Dim indFrame As Single

    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    limpiar Me

    'Ocultar todos los Frames de Formulario
    Me.FramePreFacturar.visible = False
    Me.FrameFacturar.visible = False
    
    
    For I = 14 To 17
        Me.imgBuscarOfer(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
    For I = 20 To 23
        Me.imgBuscarOfer(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I

    CommitConexion
    
    NomTabla = "scaalb"
    NomTablaLin = "slialb"
        
    Select Case OpcionListado
        'LISTADOS DE FACTURACION
        '-----------------------
        Case 50 '50: Prevision Facturacion de Albaranes (NO IMPRIME LISTADO)
            PonerFramePreFacVisible True, H, W
            indFrame = 5 'solo para el boton cancelar
            chkResumenForpa.visible = OpcionListado = 50
        Case 52 '52: Facturacion de Albaranes
            PonerFrameFacVisible True, H, W
            txtCodigo(34).Text = Format(Now, "dd/mm/yyyy")
            txtCodigo(39).Text = Format(CDate(txtCodigo(34).Text) - 1, "dd/mm/yyyy")
            indFrame = 6
            
            NomTabla = "scaalb"
            NomTablaLin = "slialb"
            
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



Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Articulos
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFPago_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Formas de Pabo
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscarOfer_Click(Index As Integer)
    Select Case Index
            
        Case 14, 15, 20, 21 'Cod. CLIENTE
            Select Case Index
                Case 11, 12: indCodigo = Index + 9
                Case 14, 15: indCodigo = Index + 14
                Case 20, 21: indCodigo = Index + 20
                Case 27, 28: indCodigo = Index + 21
                Case 32: indCodigo = 8
            End Select
            Set frmCli = New frmClientes
            frmCli.DatosADevolverBusqueda = "0|2|"
            If Not IsNumeric(txtCodigo(indCodigo).Text) Then txtCodigo(indCodigo).Text = ""
            frmCli.Show vbModal
            Set frmCli = Nothing
            
        Case 16, 17, 22, 23 'Forma de PAGO
            Select Case Index
                Case 16, 17: indCodigo = Index + 14
                Case 22, 23: indCodigo = Index + 20
                Case 29, 30: indCodigo = Index + 21
            End Select
            Set frmFPago = New frmManFpago
            frmFPago.DatosADevolverBusqueda = "0|1|"
            If Not IsNumeric(txtCodigo(indCodigo).Text) Then txtCodigo(indCodigo).Text = ""
            frmFPago.Show vbModal
            Set frmFPago = Nothing
            
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub


Private Sub imgFecha_Click(Index As Integer)
   
'++monica

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
        Case 10 'FramePreFacturar
            indCodigo = 26
        Case 11 'FramePreFacturar
            indCodigo = 27
        Case 12 'Frame Factura
            indCodigo = 38
        Case 13 'Frame Factura
            indCodigo = 39
        Case 14 'FrameFactura
            indCodigo = 34
   
   End Select
   
   PonerFormatoFecha txtCodigo(indCodigo)
   If txtCodigo(indCodigo).Text <> "" Then frmF.NovaData = CDate(txtCodigo(indCodigo).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtCodigo(indCodigo)


'++
'
'
'
'
'
'   Screen.MousePointer = vbHourglass
'   Set frmF = New frmCal
'   frmF.Fecha = Now
'
'
'   Select Case Index
'        Case 10 'FramePreFacturar
'            indCodigo = 26
'        Case 11 'FramePreFacturar
'            indCodigo = 27
'        Case 12 'Frame Factura
'            indCodigo = 38
'        Case 13 'Frame Factura
'            indCodigo = 39
'        Case 14 'FrameFactura
'            indCodigo = 34
'   End Select
'
'   PonerFormatoFecha txtCodigo(indCodigo)
'   If txtCodigo(indCodigo).Text <> "" Then frmF.Fecha = CDate(txtCodigo(indCodigo).Text)
'
'   Screen.MousePointer = vbDefault
'   frmF.Show vbModal
'   Set frmF = Nothing
'   PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub OptTipoInf_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub OptDetalle_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
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
Dim devuelve As String
Dim codCampo As String, nomCampo As String
Dim Tabla As String
      
    Select Case Index
        Case 1 'Importe (Decimal(12,2))
            PonerFormatoDecimal txtCodigo(Index), 1
            
        
        'FECHA Desde Hasta
        Case 26, 27, 34, 38, 39
            If txtCodigo(Index).Text <> "" Then
                PonerFormatoFecha txtCodigo(Index)
            End If
           
        
        Case 36, 37  'Nº de Pedido / Albaran
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
            End If
            

        Case 28, 29, 40, 41 'Cod. CLIENTE
            If PonerFormatoEntero(txtCodigo(Index)) Then
                nomCampo = "nomclien"
                Tabla = "clientes"
                codCampo = "codclien"
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), Tabla, nomCampo, codCampo, "N")
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
            Else
                txtNombre(Index).Text = ""
            End If
            
        Case 30, 31, 42, 43 'Cod. Formas de PAGO
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "forpago", "nomforpa", "codforpa", "N")
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            Else
                txtNombre(Index).Text = ""
            End If
        
    End Select
End Sub




Private Sub PonerFramePreFacVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame del Prevision Facturacion Albaran Visible y Ajustado al Formulario, y visualiza los controles
Dim b As Boolean
Dim cad As String

    H = 5600
    If OpcionListado = 51 Then 'Inf. Incum. plazos entrega
        H = 5300
        Me.cmdAceptarPreFac.Top = 4600
        Me.cmdCancel(5).Top = Me.cmdAceptarPreFac.Top
    End If
    W = 7040
    'Ajustar Tamaño del Frame para ajustar tamaño de Formulario al del Frame
    PonerFrameVisible Me.FramePreFacturar, visible, H, W
    If visible = True Then
        b = (OpcionListado = 50)
        Label4(41).visible = b
        Me.imgBuscarOfer(16).visible = b
        Me.imgBuscarOfer(17).visible = b
        Me.txtCodigo(30).visible = b
        Me.txtCodigo(31).visible = b
        Me.txtNombre(30).visible = b
        Me.txtNombre(31).visible = b
        'solo albaranes a facturar
        Me.chkSoloFacturar.visible = b
        Me.chkSoloFacturar.Value = 1
        
        'Detalle o resumen
        Me.Frame7.visible = b And (CodClien = "ALV" Or CodClien = "AL1")
        Me.OptDetalle(0).Value = True
        
        If Not b Then
            Me.Label9(0).Caption = "Incum. plazos entrega"
        Else 'Prevision Facturacion
            Select Case CodClien 'aqui guardamos el tipo de movimiento
                Case "ALV", "AL1": cad = "" ' antes "(Ventas)" [SERVICIOS]
                Case "ALR": cad = "(Reparaciones)"
                Case "ALM": cad = "(Mantenimientos)"
            End Select
            Me.Label9(0).Caption = "Previsión de facturación " & cad
        End If
    End If
End Sub


Private Sub PonerFrameFacVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Facturacion de Albaran Visible y Ajustado al Formulario, y visualiza los controles
Dim cad As String

    H = 6285
    W = 7395
    
    If visible = True Then
         Select Case CodClien 'aqui guardamos el tipo de movimiento
            Case "ALV", "AL1": cad = "(Ventas)"
                
        End Select
        
        Me.Label10(0).Caption = "Facturación de Albaranes Envases " & cad
        Me.Caption = "Facturación"
    End If
    
    PonerFrameVisible Me.FrameFacturar, visible, H, W
End Sub


Private Function AnyadirParametroDH(cad As String, indD As Byte, indH As Byte) As String
On Error Resume Next

    If txtCodigo(indD).Text <> "" And txtCodigo(indH).Text <> "" Then
        If txtCodigo(indD).Text = txtCodigo(indH).Text Then
            cad = cad & txtCodigo(indD).Text
            If txtNombre(indD).Text <> "" Then cad = cad & " - " & txtNombre(indD).Text
            AnyadirParametroDH = cad
            Exit Function
        End If
    End If
    
    If txtCodigo(indD).Text <> "" Then
        cad = cad & "desde " & txtCodigo(indD).Text
        If txtNombre(indD).Text <> "" Then cad = cad & " - " & txtNombre(indD).Text
    End If
    If txtCodigo(indH).Text <> "" Then
        cad = cad & "  hasta " & txtCodigo(indH).Text
        If txtNombre(indH).Text <> "" Then cad = cad & " - " & txtNombre(indH).Text
    End If
    AnyadirParametroDH = cad
End Function


Private Function PonerDesdeHasta(campo As String, tipo As String, indD As Byte, indH As Byte, param As String) As Boolean
Dim devuelve As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, tipo)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    If tipo <> "F" Then
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
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


Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
End Sub


Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = OpcionListado
        .Titulo = Titulo
        .ConSubInforme = conSubRPT
        .NombreRPT = nomRPT  'nombre del informe
        .Show vbModal
    End With
End Sub

Private Sub txtCodigo_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
           Case 15, 16 'ARTICULO
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "sartic", "nomartic", "codartic", "Articulo", "T")
            If txtNombre(Index).Text = "" And txtCodigo(Index) <> "" Then Cancel = True
    End Select
End Sub

Private Function ObtenerClientes(cadW As String, Importe As String) As String
Dim Sql As String
Dim Rs As ADODB.Recordset

    On Error GoTo EClientes
    
    cadW = Replace(cadW, "{", "")
    cadW = Replace(cadW, "}", "")
    
    Sql = "select codclien,nomclien,sum(baseimp1),sum(baseimp2),sum(baseimp3),sum(baseimp1)+ sum(if(isnull(baseimp2),0,baseimp2))+ sum(if(isnull(baseimp3),0,baseimp3)) as BaseImp"
    Sql = Sql & " From scafac "
    If cadW <> "" Then Sql = Sql & " where " & cadW
    Sql = Sql & " group by codclien "
    If Importe <> "" Then Sql = Sql & "having baseimp>" & Importe
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""
    While Not Rs.EOF
'        If RS!BaseImp >= CCur(Importe) Then
            Sql = Sql & Rs!CodClien & ","
'        End If
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    If Sql <> "" Then
        Sql = Mid(Sql, 1, Len(Sql) - 1)
        Sql = "( {scafac.codclien} IN [" & Sql & "] )"
    End If
    ObtenerClientes = Sql
    
EClientes:
   If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
End Function



