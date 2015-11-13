VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmModConf 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   11070
   Icon            =   "frmModConf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameInfArticulos 
      Height          =   8220
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   10950
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   25
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   90
         Text            =   "Text5"
         Top             =   7020
         Width           =   4005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   25
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   18
         Tag             =   "Palet|N|N|0|999|forfaits|codpalet|000||"
         Top             =   7020
         Width           =   675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   24
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   87
         Text            =   "Text5"
         Top             =   6690
         Width           =   4005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   24
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   17
         Tag             =   "Palet|N|N|0|999|forfaits|codpalet|000||"
         Top             =   6690
         Width           =   675
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   23
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   16
         Tag             =   "Marca|N|S|0|999|forfaits|codmarca|000||"
         Top             =   6330
         Width           =   675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   23
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   85
         Text            =   "Text5"
         Top             =   6330
         Width           =   4005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   22
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   15
         Tag             =   "Marca|N|S|0|999|forfaits|codmarca|000||"
         Top             =   6030
         Width           =   675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   22
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   82
         Text            =   "Text5"
         Top             =   6030
         Width           =   4005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   21
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   14
         Tag             =   "Presentacion|N|N|0|999|forfaits|codprese|000||"
         Top             =   5640
         Width           =   675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   21
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   80
         Text            =   "Text5"
         Top             =   5640
         Width           =   4005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   20
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   13
         Tag             =   "Presentacion|N|N|0|999|forfaits|codprese|000||"
         Top             =   5340
         Width           =   675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   20
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   77
         Text            =   "Text5"
         Top             =   5340
         Width           =   4005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   19
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   12
         Tag             =   "Confeccion|N|N|0|999|forfaits|codtipco|000||"
         Top             =   4950
         Width           =   675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   19
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   75
         Text            =   "Text5"
         Top             =   4950
         Width           =   4005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   11
         Tag             =   "Confeccion|N|N|0|999|forfaits|codtipco|000||"
         Top             =   4650
         Width           =   675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   18
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   72
         Text            =   "Text5"
         Top             =   4650
         Width           =   4005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   10
         Tag             =   "Medida|N|N|0|999|forfaits|codmedid|000||"
         Top             =   4230
         Width           =   675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   17
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   70
         Text            =   "Text5"
         Top             =   4230
         Width           =   4005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   9
         Tag             =   "Medida|N|N|0|999|forfaits|codmedid|000||"
         Top             =   3930
         Width           =   675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   16
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   67
         Text            =   "Text5"
         Top             =   3930
         Width           =   4005
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   15
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   65
         Text            =   "Text5"
         Top             =   3510
         Width           =   4005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   15
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   8
         Tag             =   "Capacidad|N|N|0|999|forfaits|codcapac|000||"
         Top             =   3510
         Width           =   675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   14
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   62
         Text            =   "Text5"
         Top             =   3210
         Width           =   4005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   14
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   7
         Tag             =   "Capacidad|N|N|0|999|forfaits|codcapac|000||"
         Top             =   3210
         Width           =   675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   13
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   60
         Text            =   "Text5"
         Top             =   2820
         Width           =   4005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   6
         Tag             =   "Envase|N|N|0|999|forfaits|codtipen|000||"
         Top             =   2820
         Width           =   675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   12
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   "Text5"
         Top             =   2520
         Width           =   4005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   5
         Tag             =   "Envase|N|N|0|999|forfaits|codtipen|000||"
         Top             =   2520
         Width           =   675
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   10
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "Variedad|N|S|||forfaits|codvarie|000000||"
         Top             =   1830
         Width           =   1035
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   11
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "Variedad|N|S|||forfaits|codvarie|000000||"
         Top             =   2115
         Width           =   1035
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   10
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   53
         Text            =   "Text5"
         Top             =   1830
         Width           =   3645
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   11
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   52
         Text            =   "Text5"
         Top             =   2115
         Width           =   3645
      End
      Begin VB.Frame Frame3 
         Caption         =   "Nuevo Coste"
         ForeColor       =   &H00972E0B&
         Height          =   1755
         Left            =   6960
         TabIndex        =   37
         Top             =   3720
         Width           =   3495
         Begin VB.TextBox txtCodigo 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   2160
            MaxLength       =   16
            TabIndex        =   38
            Top             =   210
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.TextBox txtCodigo 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   2160
            MaxLength       =   16
            TabIndex        =   39
            Top             =   510
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.TextBox txtCodigo 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   2160
            MaxLength       =   16
            TabIndex        =   40
            Top             =   810
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.TextBox txtCodigo 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   2160
            MaxLength       =   16
            TabIndex        =   41
            Top             =   1110
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.TextBox txtCodigo 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   2160
            MaxLength       =   16
            TabIndex        =   42
            Top             =   1410
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Label Label3 
            Caption         =   "Label2"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Width           =   1875
         End
         Begin VB.Label Label3 
            Caption         =   "Label2"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   48
            Top             =   540
            Width           =   1875
         End
         Begin VB.Label Label3 
            Caption         =   "Label2"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   47
            Top             =   840
            Width           =   1875
         End
         Begin VB.Label Label3 
            Caption         =   "Label2"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   45
            Top             =   1140
            Width           =   1875
         End
         Begin VB.Label Label3 
            Caption         =   "Label2"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   43
            Top             =   1440
            Width           =   1875
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Variación"
         ForeColor       =   &H00972E0B&
         Height          =   1755
         Left            =   6960
         TabIndex        =   26
         Top             =   1740
         Width           =   3495
         Begin VB.TextBox txtCodigo 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   2160
            MaxLength       =   16
            TabIndex        =   31
            Top             =   1410
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.TextBox txtCodigo 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   2160
            MaxLength       =   16
            TabIndex        =   30
            Top             =   1110
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.TextBox txtCodigo 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   2160
            MaxLength       =   16
            TabIndex        =   29
            Top             =   810
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.TextBox txtCodigo 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   2160
            MaxLength       =   16
            TabIndex        =   28
            Top             =   510
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.TextBox txtCodigo 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   2160
            MaxLength       =   16
            TabIndex        =   27
            Top             =   210
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   36
            Top             =   1440
            Width           =   1875
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   35
            Top             =   1140
            Width           =   1875
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   34
            Top             =   840
            Width           =   1875
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   33
            Top             =   540
            Width           =   1875
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1875
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo"
         ForeColor       =   &H00972E0B&
         Height          =   555
         Left            =   6960
         TabIndex        =   25
         Top             =   990
         Width           =   3495
         Begin VB.OptionButton Opcion 
            Caption         =   "Nuevo Coste"
            Height          =   255
            Index           =   5
            Left            =   2010
            TabIndex        =   20
            Top             =   210
            Width           =   1275
         End
         Begin VB.OptionButton Opcion 
            Caption         =   "Variación"
            Height          =   255
            Index           =   4
            Left            =   570
            TabIndex        =   19
            Top             =   210
            Width           =   1155
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   71
         Left            =   3195
         Locked          =   -1  'True
         TabIndex        =   51
         Text            =   "Text5"
         Top             =   1365
         Width           =   3255
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   70
         Left            =   3195
         Locked          =   -1  'True
         TabIndex        =   50
         Text            =   "Text5"
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   71
         Left            =   1710
         MaxLength       =   16
         TabIndex        =   2
         Top             =   1365
         Width           =   1455
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   70
         Left            =   1710
         MaxLength       =   16
         TabIndex        =   1
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   405
         Left            =   8280
         TabIndex        =   44
         Top             =   7275
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   405
         Left            =   9450
         TabIndex        =   46
         Top             =   7275
         Width           =   975
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   15
         Left            =   1425
         ToolTipText     =   "Buscar palet"
         Top             =   7020
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   20
         Left            =   870
         TabIndex        =   91
         Top             =   7020
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   14
         Left            =   1425
         ToolTipText     =   "Buscar palet"
         Top             =   6720
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Confección"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   7
         Left            =   480
         TabIndex        =   89
         Top             =   4470
         Width           =   810
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   19
         Left            =   870
         TabIndex        =   88
         Top             =   6720
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   18
         Left            =   870
         TabIndex        =   86
         Top             =   6330
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   13
         Left            =   1425
         ToolTipText     =   "Buscar marca"
         Top             =   6330
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   17
         Left            =   870
         TabIndex        =   84
         Top             =   6060
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Palet"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   6
         Left            =   480
         TabIndex        =   83
         Top             =   6540
         Width           =   360
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   12
         Left            =   1425
         ToolTipText     =   "Buscar marca"
         Top             =   6060
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   16
         Left            =   870
         TabIndex        =   81
         Top             =   5640
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   11
         Left            =   1425
         ToolTipText     =   "Buscar presentación"
         Top             =   5640
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   870
         TabIndex        =   79
         Top             =   5370
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Marca"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   5
         Left            =   480
         TabIndex        =   78
         Top             =   5820
         Width           =   450
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   10
         Left            =   1425
         ToolTipText     =   "Buscar presentacion"
         Top             =   5370
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   870
         TabIndex        =   76
         Top             =   4950
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   1425
         ToolTipText     =   "Buscar confección"
         Top             =   4950
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   870
         TabIndex        =   74
         Top             =   4680
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Presentación"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   73
         Top             =   5160
         Width           =   930
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   1425
         ToolTipText     =   "Buscar confección"
         Top             =   4680
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   870
         TabIndex        =   71
         Top             =   4230
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   1425
         ToolTipText     =   "Buscar medida"
         Top             =   4230
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   11
         Left            =   870
         TabIndex        =   69
         Top             =   3960
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Medida"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   3
         Left            =   510
         TabIndex        =   68
         Top             =   3750
         Width           =   525
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1425
         ToolTipText     =   "Buscar medida"
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1425
         ToolTipText     =   "Buscar capacidad"
         Top             =   3510
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   10
         Left            =   870
         TabIndex        =   66
         Top             =   3510
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1425
         ToolTipText     =   "Buscar capacidad"
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Capacidad"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   2
         Left            =   510
         TabIndex        =   64
         Top             =   3030
         Width           =   765
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   9
         Left            =   870
         TabIndex        =   63
         Top             =   3240
         Width           =   465
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1425
         ToolTipText     =   "Buscar envase"
         Top             =   2820
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   8
         Left            =   870
         TabIndex        =   61
         Top             =   2820
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1425
         ToolTipText     =   "Buscar envase"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Envase"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   1
         Left            =   510
         TabIndex        =   59
         Top             =   2310
         Width           =   540
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   7
         Left            =   870
         TabIndex        =   58
         Top             =   2520
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   6
         Left            =   870
         TabIndex        =   56
         Top             =   1830
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   5
         Left            =   870
         TabIndex        =   55
         Top             =   2115
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   0
         Left            =   510
         TabIndex        =   54
         Top             =   1620
         Width           =   630
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1425
         ToolTipText     =   "Buscar variedad"
         Top             =   1830
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1425
         ToolTipText     =   "Buscar variedad"
         Top             =   2115
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cambio Costes de Confecciones"
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
         TabIndex        =   24
         Top             =   300
         Width           =   6735
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   28
         Left            =   1425
         ToolTipText     =   "Buscar forfait"
         Top             =   1365
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   27
         Left            =   1425
         ToolTipText     =   "Buscar forfait"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Confección"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   38
         Left            =   510
         TabIndex        =   23
         Top             =   870
         Width           =   810
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   54
         Left            =   870
         TabIndex        =   22
         Top             =   1365
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   51
         Left            =   870
         TabIndex        =   21
         Top             =   1080
         Width           =   465
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   8730
      Top             =   5580
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmModConf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmConf As frmManForfaits
Attribute frmConf.VB_VarHelpID = -1

Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmVar As frmManVariedad ' mantenimiento de variedades
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmEnv As frmManTipEnv 'tipos de envases
Attribute frmEnv.VB_VarHelpID = -1
Private WithEvents frmCap As frmManCapEnv 'capacidad
Attribute frmCap.VB_VarHelpID = -1
Private WithEvents frmMed As frmManMedEnv 'medidas
Attribute frmMed.VB_VarHelpID = -1
Private WithEvents frmPal As frmManPaleConf 'palets
Attribute frmPal.VB_VarHelpID = -1
Private WithEvents frmPres As frmManPresConf 'presentacion
Attribute frmPres.VB_VarHelpID = -1
Private WithEvents frmMar As frmManMarcas 'marcas
Attribute frmMar.VB_VarHelpID = -1
Private WithEvents frmConf2 As frmManFConf 'confeccion
Attribute frmConf2.VB_VarHelpID = -1

'---- Variables para el INFORME ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para el frmImprimir
Private conSubRPT As Boolean 'Si el informe tiene subreports
Private cadNombreRPT As String 'Nombre del informe
'-----------------------------------

Dim TipCod As String
Dim indCodigo As Integer 'indice para txtCodigo

Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim CambioHorientacionPapel As Boolean ' indicamos si se va a imprimir en landscape

Dim PrimeraVez As Boolean
Dim indFrame As Single
Dim Salir As Boolean

Private Sub KEYpress(KeyAscii As Integer)
    Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cmdAceptar_Click()
'Listado de Articulos
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim campo As String
Dim Opcion As Byte, numOp As Byte


    If Not DatosOk Then Exit Sub
    
    InicializarVbles
    
    cadTABLA = "forfaits"
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    '====================================================
    '================= FORMULA ==========================
    
    'Cadena para seleccion D/H Confeccion
    '--------------------------------------------
    cDesde = Trim(txtCodigo(70).Text)
    cHasta = Trim(txtCodigo(71).Text)
    nDesde = txtNombre(70).Text
    nHasta = txtNombre(71).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & cadTABLA & ".codforfait}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHConfeccion= """) Then Exit Sub
    End If

    'Cadena para seleccion D/H Variedad
    '--------------------------------------------
    cDesde = Trim(txtCodigo(10).Text)
    cHasta = Trim(txtCodigo(11).Text)
    nDesde = txtNombre(10).Text
    nHasta = txtNombre(11).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & cadTABLA & ".codvarie}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad= """) Then Exit Sub
    End If

    'Cadena para seleccion D/H Envase
    '--------------------------------------------
    cDesde = Trim(txtCodigo(12).Text)
    cHasta = Trim(txtCodigo(13).Text)
    nDesde = txtNombre(12).Text
    nHasta = txtNombre(13).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & cadTABLA & ".codtipen}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHEnvase= """) Then Exit Sub
    End If

    'Cadena para seleccion D/H Capacidad
    '--------------------------------------------
    cDesde = Trim(txtCodigo(14).Text)
    cHasta = Trim(txtCodigo(15).Text)
    nDesde = txtNombre(14).Text
    nHasta = txtNombre(15).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & cadTABLA & ".codcapac}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCapacidad= """) Then Exit Sub
    End If

    'Cadena para seleccion D/H Medida
    '--------------------------------------------
    cDesde = Trim(txtCodigo(16).Text)
    cHasta = Trim(txtCodigo(17).Text)
    nDesde = txtNombre(16).Text
    nHasta = txtNombre(17).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & cadTABLA & ".codmedid}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHMedida= """) Then Exit Sub
    End If

    'Cadena para seleccion D/H Confeccion
    '--------------------------------------------
    cDesde = Trim(txtCodigo(18).Text)
    cHasta = Trim(txtCodigo(19).Text)
    nDesde = txtNombre(18).Text
    nHasta = txtNombre(19).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & cadTABLA & ".codtipco}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHConfeccion= """) Then Exit Sub
    End If

    'Cadena para seleccion D/H Presentacion
    '--------------------------------------------
    cDesde = Trim(txtCodigo(20).Text)
    cHasta = Trim(txtCodigo(21).Text)
    nDesde = txtNombre(20).Text
    nHasta = txtNombre(21).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & cadTABLA & ".codprese}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHPresentacion= """) Then Exit Sub
    End If
    
    'Cadena para seleccion D/H Marca
    '--------------------------------------------
    cDesde = Trim(txtCodigo(22).Text)
    cHasta = Trim(txtCodigo(23).Text)
    nDesde = txtNombre(22).Text
    nHasta = txtNombre(23).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & cadTABLA & ".codmarca}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHMarca= """) Then Exit Sub
    End If

    'Cadena para seleccion D/H Palet
    '--------------------------------------------
    cDesde = Trim(txtCodigo(24).Text)
    cHasta = Trim(txtCodigo(25).Text)
    nDesde = txtNombre(24).Text
    nHasta = txtNombre(25).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & cadTABLA & ".codpalet}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHPalet= """) Then Exit Sub
    End If


    If HayRegParaInforme(cadTABLA, cadSelect) Then
        If Not BloqueaRegistro(cadTABLA, cadSelect) Then
            MsgBox "No se pueden Actualizar costes. Hay registros bloqueados.", vbExclamation
            Screen.MousePointer = vbDefault
        Else
            If ProcesarCambios(cadTABLA, cadSelect) Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                cmdCancel_Click
            End If
        End If
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(70)
        Me.Opcion(4).Value = True
        PonerFoco txtCodigo(70)
        If Salir = True Then Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim I As Integer
Dim Maxim As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    limpiar Me

    Salir = False

    For I = 27 To 28
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
    For I = 0 To 15
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I

    'Ocultar todos los Frames de Formulario
    Me.FrameInfArticulos.visible = False
    
    CommitConexion
    
    cadTitulo = ""
    cadNombreRPT = ""
    
    ListadosAlmacen H, W
    
    Maxim = DevuelveValor("select count(*) from nombcoste")
    If Maxim > 5 Then
        MsgBox "Se ha superado el número maximo permitido de costes. Revise.", vbExclamation
        Salir = True
        Exit Sub
    End If
    
    For I = 0 To 4
        Label2(I).Caption = ""
        Label3(I).Caption = ""
    Next I
    
    For I = 0 To Maxim - 1
        Label2(I).Caption = DevuelveValor("select denominacion from nombcoste where codcoste = " & DBSet(I + 1, "N"))
        Label3(I).Caption = Label2(I).Caption
        txtCodigo(I).visible = True
        txtCodigo(I).Enabled = True
        txtCodigo(I + 5).visible = True
        txtCodigo(I + 5).Enabled = True
    Next I
    
'    Frame3.Enabled = False
'    For i = 5 To 9
'        txtCodigo(i).Text = ""
'    Next i
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub

Private Sub frmCap_DatoSeleccionado(CadenaSeleccion As String)
'Capacidades
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'capacidad
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmConf2_DatoSeleccionado(CadenaSeleccion As String)
'Confecciones
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'codconfe
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmEnv_DatoSeleccionado(CadenaSeleccion As String)
'tipos de envase
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'codtipen
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmMar_DatoSeleccionado(CadenaSeleccion As String)
'Marcas
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'codmarca
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmMed_DatoSeleccionado(CadenaSeleccion As String)
'Medidas de envase
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'codmedid
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmPal_DatoSeleccionado(CadenaSeleccion As String)
'Tipos de Palets
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'codpalet
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmPres_DatoSeleccionado(CadenaSeleccion As String)
'Presentacion
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'codpresentacion
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub


Private Sub frmConf_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de confecciones
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtCodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de variedades
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
'Buscar general: cada index llama a una tabla
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 27, 28 'cod. de confeccion
            indCodigo = Index + 43
            Set frmConf = New frmManForfaits
            frmConf.DatosADevolverBusqueda = "0|1|" 'Abrimos en Modo Busqueda
            frmConf.DeConsulta = True
            frmConf.Show vbModal
            Set frmConf = Nothing
            
        Case 0, 1 'variedades
            indCodigo = Index + 10
            Set frmVar = New frmManVariedad
            frmVar.DatosADevolverBusqueda = "0|1|" 'Abrimos en Modo Busqueda
            frmVar.DeConsulta = True
            frmVar.Show vbModal
            Set frmVar = Nothing
        
        Case 2, 3 'tipo de envases
            indCodigo = Index + 10
            Set frmEnv = New frmManTipEnv
            frmEnv.DatosADevolverBusqueda = "0|1|"
            frmEnv.CodigoActual = txtCodigo(indCodigo).Text
            frmEnv.Show vbModal
            Set frmEnv = Nothing
            PonerFoco txtCodigo(indCodigo)
        Case 4, 5 'capacidad
            indCodigo = Index + 10
            Set frmCap = New frmManCapEnv
            frmCap.DatosADevolverBusqueda = "0|1|"
            frmCap.CodigoActual = txtCodigo(indCodigo).Text
            frmCap.Show vbModal
            Set frmCap = Nothing
            PonerFoco txtCodigo(indCodigo)
        Case 6, 7 'medidas de envases
            indCodigo = Index + 10
            Set frmMed = New frmManMedEnv
            frmMed.DatosADevolverBusqueda = "0|1|"
            frmMed.CodigoActual = txtCodigo(indCodigo).Text
            frmMed.Show vbModal
            Set frmMed = Nothing
            PonerFoco txtCodigo(indCodigo)
        Case 8, 9 'Confeccion
            indCodigo = Index + 10
            Set frmConf2 = New frmManFConf
            frmConf2.DatosADevolverBusqueda = "0|1|"
            frmConf2.CodigoActual = txtCodigo(indCodigo).Text
            frmConf2.Show vbModal
            Set frmConf2 = Nothing
            PonerFoco txtCodigo(indCodigo)
        Case 10, 11 'Presentacion
            indCodigo = Index + 10
            Set frmPres = New frmManPresConf
            frmPres.DatosADevolverBusqueda = "0|1|"
            frmPres.CodigoActual = txtCodigo(indCodigo).Text
            frmPres.Show vbModal
            Set frmPres = Nothing
            PonerFoco txtCodigo(indCodigo)
        Case 12, 13 'Marca
            indCodigo = Index + 10
            Set frmMar = New frmManMarcas
            frmMar.DatosADevolverBusqueda = "0|1|"
            frmMar.CodigoActual = txtCodigo(indCodigo).Text
            frmMar.Show vbModal
            Set frmMar = Nothing
            PonerFoco txtCodigo(indCodigo)
        Case 14, 15 'Palet
            indCodigo = Index + 10
            Set frmPal = New frmManPaleConf
            frmPal.DatosADevolverBusqueda = "0|1|"
            frmPal.CodigoActual = txtCodigo(indCodigo).Text
            frmPal.Show vbModal
            Set frmPal = Nothing
            PonerFoco txtCodigo(indCodigo)
            
    End Select
    PonerFoco txtCodigo(indCodigo)
    Screen.MousePointer = vbDefault
End Sub




Private Sub Opcion_Click(Index As Integer)
Dim I As Integer

    If Index = 4 Then
        Frame3.Enabled = False
        For I = 5 To 9
            txtCodigo(I).Text = ""
        Next I
        Frame2.Enabled = True
        PonerFoco txtCodigo(0)
    Else
        Frame2.Enabled = False
        For I = 0 To 4
            txtCodigo(I).Text = ""
        Next I
        Frame3.Enabled = True
        PonerFoco txtCodigo(5)
    End If
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 70: KEYBusqueda KeyAscii, 27 'forfait
            Case 71: KEYBusqueda KeyAscii, 28 'forfait
            Case 10: KEYBusqueda KeyAscii, 0 'variedad
            Case 11: KEYBusqueda KeyAscii, 1 'variedad
        
            Case 12: KEYBusqueda KeyAscii, 2 'envase
            Case 13: KEYBusqueda KeyAscii, 3 'envase
            Case 14: KEYBusqueda KeyAscii, 4 'capacidad
            Case 15: KEYBusqueda KeyAscii, 5 'capacidad
            Case 16: KEYBusqueda KeyAscii, 6 'medida
            Case 17: KEYBusqueda KeyAscii, 7 'medida
            Case 18: KEYBusqueda KeyAscii, 8 'confeccion
            Case 19: KEYBusqueda KeyAscii, 9 'confeccion
            Case 20: KEYBusqueda KeyAscii, 10 'presentacion
            Case 21: KEYBusqueda KeyAscii, 11 'presentacion
            Case 22: KEYBusqueda KeyAscii, 12 'marca
            Case 23: KEYBusqueda KeyAscii, 13 'marca
            Case 24: KEYBusqueda KeyAscii, 14 'palet
            Case 25: KEYBusqueda KeyAscii, 15 'palet
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
Dim Tabla As String
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
        
    Select Case Index
        Case 0, 1, 2, 3, 4 ' porcentaje ( 2 decimales )
            PonerFormatoDecimal txtCodigo(Index), 4
            
            If Index = 4 Then cmdAceptar.SetFocus
            
        Case 5, 6, 7, 8, 9 ' importe coste ( 4 decimales )
            PonerFormatoDecimal txtCodigo(Index), 7
    
            If Index = 9 Then cmdAceptar.SetFocus
            
        Case 10, 11 ' variedades
            PonerFormatoEntero txtCodigo(Index)
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "variedades", "nomvarie", "codvarie", "N")
        
        Case 70, 71  'Cod. forfait
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "forfaits", "nomconfe", "codforfait", "T")
            
        Case 12, 13 'envase
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "confenva", "nomtipen")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
        
        Case 14, 15 'capacidad
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "capacida", "nomcapac")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
        
        Case 16, 17 'medida
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "confmedi", "nommedid")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
        
        Case 18, 19 'confeccion
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "conftipo", "nomtipco")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
        
        Case 20, 21 'presentacion
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "confpres", "nomprese")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
        
        Case 22, 23 'marca
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "marcas", "nommarca")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
        
        Case 24, 25 'palet
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "confpale", "nompalet")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            
    End Select
    
End Sub

'Private Sub frmB_Selecionado(CadenaDevuelta As String)
''Formulario para Busqueda
'
'    If CadenaDevuelta <> "" Then
'        HaDevueltoDatos = True
'        Screen.MousePointer = vbHourglass
'        Select Case OpcionListado
'            Case 7, 8 'Informe Traspasos Almacen
'                txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaDevuelta, 1), "0000000")
'                PonerFoco txtCodigo(indCodigo)
'            Case 9, 12, 13, 14, 15, 16, 17 '9: Informe Movimiento Articulos
'                                'Inventario Articulos
'                                '14: Actualizar diferencias Stock Inventariado
'                                '16: Listado Valoracion stock inventariado
'                txtCodigo(indCodigo).Text = RecuperaValor(CadenaDevuelta, 1)
'                txtNombre(indCodigo).Text = RecuperaValor(CadenaDevuelta, 2)
'                PonerFoco txtCodigo(indCodigo)
'        End Select
'    End If
'    Screen.MousePointer = vbDefault
'End Sub

Private Sub ponerFrameArticulosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el informe de Articulos, de tabla: sartic
Dim b As Boolean

    b = True
    H = 8220
    W = 10950
    
    PonerFrameVisible Me.FrameInfArticulos, visible, H, W

End Sub



Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
    conSubRPT = False
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
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        devuelve2 = CadenaDesdeHastaBD(codD, codH, Codigo, TipCod)
        If devuelve2 = "Error" Then Exit Function
        If Not AnyadirAFormula(cadSelect, devuelve2) Then Exit Function
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
        .EnvioEMail = False
        .NombreRPT = cadNombreRPT
        .ConSubinforme = True
        .Opcion = 0 'Opcion
        .CambioHorientacionPapel = CambioHorientacionPapel
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
        Case "Forfait"
            cadParam = cadParam & campo & "{forfaits.codforfait}" & "|"
'            If numGrupo = 1 Then
'                cadParam = cadParam & nomCampo & "|"
'            End If
            numParam = numParam + 1
            
        Case "NomForfait"
            cadParam = cadParam & campo & "{forfaits.nomconfe}" & "|"
            numParam = numParam + 1
    End Select

End Function


Private Function ComprobarFechasConta(ind As Integer) As Boolean
'comprobar que el periodo de fechas a contabilizar esta dentro del
'periodo de fechas del ejercicio de la contabilidad
Dim FechaIni As String, FechaFin As String
Dim cad As String
Dim Rs As ADODB.Recordset
    
On Error GoTo EComprobar

    ComprobarFechasConta = False
    
    If txtCodigo(ind).Text <> "" Then
        FechaIni = "Select fechaini,fechafin From parametros"
        Set Rs = New ADODB.Recordset
        Rs.Open FechaIni, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        If Not Rs.EOF Then
            FechaIni = DBLet(Rs!FechaIni, "F")
            FechaFin = DBLet(Rs!FechaFin, "F") + 365
            'nos guardamos los valores
            Orden1 = FechaIni
            Orden2 = FechaFin
        
            If Not EntreFechas(FechaIni, txtCodigo(ind).Text, FechaFin) Then
                 cad = "El período de contabilización debe estar dentro del ejercicio:" & vbCrLf & vbCrLf
                 cad = cad & "    Desde: " & FechaIni & vbCrLf
                 cad = cad & "    Hasta: " & FechaFin
                 MsgBox cad, vbExclamation
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

Private Sub ListadosAlmacen(H As Integer, W As Integer)
   'Listado de Artículo
    ponerFrameArticulosVisible True, H, W
End Sub


Private Function DatosOk() As Boolean
    DatosOk = False
        
    If Opcion(4).Value = True Then ' variacion coste
        If txtCodigo(0).Text = "" And txtCodigo(1).Text = "" And txtCodigo(2).Text = "" And _
           txtCodigo(3).Text = "" And txtCodigo(4).Text = "" Then
            MsgBox "Debe introducir algún valor de variación coste. Revise.", vbExclamation
            Exit Function
        End If
    Else  ' nuevo coste
        If txtCodigo(5).Text = "" And txtCodigo(6).Text = "" And txtCodigo(7).Text = "" And _
           txtCodigo(8).Text = "" And txtCodigo(9).Text = "" Then
            MsgBox "Debe introducir algún valor de variación coste. Revise.", vbExclamation
            Exit Function
        End If
    End If
    
    DatosOk = True

End Function

Private Function ProcesarCambios(nTabla As String, cadSelect As String) As Boolean
Dim vSQL As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset
Dim RS3 As ADODB.Recordset
Dim Importe As String
Dim Sql As String
Dim Sql1 As String

Dim Albaran As Long
Dim Linea As Long

Dim Codigiva As String

    On Error GoTo eProcesarCambios


    ProcesarCambios = False
    
    conn.BeginTrans

    If cadSelect = "" Then cadSelect = "(1=1)"
    
    nTabla = QuitarCaracterACadena(nTabla, "{")
    nTabla = QuitarCaracterACadena(nTabla, "}")

    If cadSelect <> "" Then
        cadSelect = QuitarCaracterACadena(cadSelect, "{")
        cadSelect = QuitarCaracterACadena(cadSelect, "}")
        cadSelect = QuitarCaracterACadena(cadSelect, "_1")
    End If

    vSQL = "select codforfait from " & nTabla
    If cadSelect <> "" Then vSQL = vSQL & " where " & cadSelect

    If Opcion(4).Value Then 'variacion coste
        If txtCodigo(0).Text <> "" Then
            Sql = "update forfaits_costes set importes = importes + round(importes * " & DBSet(txtCodigo(0).Text, "N") & " / 100, 4) "
            Sql = Sql & " where codcoste = 1 and codforfait in (" & vSQL & ")"
            
            conn.Execute Sql
        End If
        If txtCodigo(1).Text <> "" Then
            Sql = "update forfaits_costes set importes = importes + round(importes * " & DBSet(txtCodigo(1).Text, "N") & " / 100, 4) "
            Sql = Sql & " where codcoste = 2 and codforfait in (" & vSQL & ")"
            
            conn.Execute Sql
        End If
        If txtCodigo(2).Text <> "" Then
            Sql = "update forfaits_costes set importes = importes + round(importes * " & DBSet(txtCodigo(2).Text, "N") & " / 100, 4) "
            Sql = Sql & " where codcoste = 3 and codforfait in (" & vSQL & ")"
            
            conn.Execute Sql
        End If
        If txtCodigo(3).Text <> "" Then
            Sql = "update forfaits_costes set importes = importes + round(importes * " & DBSet(txtCodigo(3).Text, "N") & " / 100, 4) "
            Sql = Sql & " where codcoste = 4 and codforfait in (" & vSQL & ")"
            
            conn.Execute Sql
        End If
        If txtCodigo(4).Text <> "" Then
            Sql = "update forfaits_costes set importes = importes + round(importes * " & DBSet(txtCodigo(4).Text, "N") & " / 100, 4) "
            Sql = Sql & " where codcoste = 5 and codforfait in (" & vSQL & ")"
            
            conn.Execute Sql
        End If
    Else ' nuevo coste
        If txtCodigo(5).Text <> "" Then
            Sql = "update forfaits_costes set importes = " & DBSet(txtCodigo(5).Text, "N")
            Sql = Sql & " where codcoste = 1 and codforfait in (" & vSQL & ")"
            
            conn.Execute Sql
        End If
        If txtCodigo(6).Text <> "" Then
            Sql = "update forfaits_costes set importes = " & DBSet(txtCodigo(6).Text, "N")
            Sql = Sql & " where codcoste = 2 and codforfait in (" & vSQL & ")"
            
            conn.Execute Sql
        End If
        If txtCodigo(7).Text <> "" Then
            Sql = "update forfaits_costes set importes = " & DBSet(txtCodigo(7).Text, "N")
            Sql = Sql & " where codcoste = 3 and codforfait in (" & vSQL & ")"
            
            conn.Execute Sql
        End If
        If txtCodigo(8).Text <> "" Then
            Sql = "update forfaits_costes set importes = " & DBSet(txtCodigo(8).Text, "N")
            Sql = Sql & " where codcoste = 4 and codforfait in (" & vSQL & ")"
            
            conn.Execute Sql
        End If
        If txtCodigo(9).Text <> "" Then
            Sql = "update forfaits_costes set importes = " & DBSet(txtCodigo(9).Text, "N")
            Sql = Sql & " where codcoste = 5 and codforfait in (" & vSQL & ")"
            
            conn.Execute Sql
        End If
    End If
       
    conn.CommitTrans
    ProcesarCambios = True
    Exit Function
    
eProcesarCambios:
    conn.RollbackTrans
    MuestraError Err.Number, "Procesar Cambios", Err.Description
End Function


