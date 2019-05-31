VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmVtasOrdenCarga 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   11940
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   12615
   Icon            =   "frmVtasOrdenCarga.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11940
   ScaleWidth      =   12615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameCobros 
      Height          =   11970
      Left            =   0
      TabIndex        =   34
      Top             =   -60
      Width           =   12525
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
         Left            =   2145
         MaxLength       =   100
         TabIndex        =   27
         Top             =   9720
         Width           =   7365
      End
      Begin VB.Frame Frame4 
         Caption         =   "FACTURACION DEL TRANSPORTE A : "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   420
         TabIndex        =   57
         Top             =   10065
         Width           =   9300
         Begin VB.TextBox txtCodigo 
            Height          =   360
            Index           =   30
            Left            =   120
            MaxLength       =   100
            TabIndex        =   31
            Top             =   1290
            Width           =   8985
         End
         Begin VB.TextBox txtCodigo 
            Height          =   360
            Index           =   29
            Left            =   120
            MaxLength       =   100
            TabIndex        =   30
            Top             =   930
            Width           =   8985
         End
         Begin VB.TextBox txtCodigo 
            Height          =   360
            Index           =   28
            Left            =   120
            MaxLength       =   100
            TabIndex        =   29
            Top             =   570
            Width           =   8985
         End
         Begin VB.TextBox txtCodigo 
            Height          =   360
            Index           =   18
            Left            =   120
            MaxLength       =   100
            TabIndex        =   28
            Top             =   210
            Width           =   8985
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "ENTREGA 2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3090
         Left            =   390
         TabIndex        =   50
         Top             =   6615
         Width           =   9285
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
            Index           =   33
            Left            =   1755
            MaxLength       =   250
            TabIndex        =   24
            Top             =   1965
            Width           =   7440
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
            Index           =   27
            Left            =   1755
            MaxLength       =   250
            TabIndex        =   23
            Top             =   1605
            Width           =   7440
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
            Index           =   26
            Left            =   1755
            MaxLength       =   250
            TabIndex        =   21
            Top             =   885
            Width           =   7425
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
            Left            =   1755
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   26
            Top             =   2700
            Width           =   7395
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
            Left            =   1755
            MaxLength       =   20
            TabIndex        =   25
            Top             =   2340
            Width           =   2790
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
            Left            =   1755
            MaxLength       =   250
            TabIndex        =   22
            Top             =   1260
            Width           =   7440
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
            Left            =   1755
            MaxLength       =   250
            ScrollBars      =   2  'Vertical
            TabIndex        =   20
            Top             =   540
            Width           =   7425
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
            Left            =   1755
            MaxLength       =   20
            TabIndex        =   18
            Top             =   165
            Width           =   2850
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
            Index           =   12
            Left            =   5940
            MaxLength       =   30
            TabIndex        =   19
            Top             =   165
            Width           =   3240
         End
         Begin VB.Label Label4 
            Caption         =   "Observaciones"
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
            Height          =   255
            Index           =   17
            Left            =   150
            TabIndex        =   56
            Top             =   2700
            Width           =   2355
         End
         Begin VB.Label Label4 
            Caption         =   "Teléfono"
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
            Height          =   255
            Index           =   15
            Left            =   150
            TabIndex        =   55
            Top             =   2355
            Width           =   1860
         End
         Begin VB.Label Label4 
            Caption         =   "Lugar"
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
            Height          =   255
            Index           =   14
            Left            =   180
            TabIndex        =   54
            Top             =   1260
            Width           =   1860
         End
         Begin VB.Label Label4 
            Caption         =   "Mercancia"
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
            Height          =   255
            Index           =   13
            Left            =   180
            TabIndex        =   53
            Top             =   570
            Width           =   1365
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Carga"
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
            Height          =   255
            Index           =   12
            Left            =   180
            TabIndex        =   52
            Top             =   195
            Width           =   960
         End
         Begin VB.Label Label4 
            Caption         =   "Hora Carga"
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
            Height          =   255
            Index           =   11
            Left            =   4995
            TabIndex        =   51
            Top             =   195
            Width           =   960
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "ENTREGA 1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3090
         Left            =   390
         TabIndex        =   43
         Top             =   3540
         Width           =   9285
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
            Left            =   1725
            MaxLength       =   100
            ScrollBars      =   2  'Vertical
            TabIndex        =   15
            Top             =   1890
            Width           =   7470
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
            Left            =   1725
            MaxLength       =   100
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   1545
            Width           =   7470
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
            Index           =   23
            Left            =   1725
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   855
            Width           =   7470
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
            Index           =   11
            Left            =   5895
            MaxLength       =   30
            TabIndex        =   10
            Top             =   165
            Width           =   3285
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
            Index           =   10
            Left            =   1725
            MaxLength       =   20
            TabIndex        =   9
            Top             =   165
            Width           =   2940
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
            Index           =   9
            Left            =   1725
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   510
            Width           =   7470
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
            Left            =   1725
            MaxLength       =   100
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   1200
            Width           =   7470
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
            Left            =   1725
            MaxLength       =   20
            TabIndex        =   16
            Top             =   2250
            Width           =   2970
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
            Left            =   1725
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   17
            Top             =   2625
            Width           =   7440
         End
         Begin VB.Label Label4 
            Caption         =   "Hora Carga"
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
            Height          =   255
            Index           =   10
            Left            =   4995
            TabIndex        =   49
            Top             =   165
            Width           =   960
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Carga"
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
            Height          =   255
            Index           =   9
            Left            =   180
            TabIndex        =   48
            Top             =   195
            Width           =   960
         End
         Begin VB.Label Label4 
            Caption         =   "Mercancia"
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
            Height          =   255
            Index           =   8
            Left            =   180
            TabIndex        =   47
            Top             =   510
            Width           =   1365
         End
         Begin VB.Label Label4 
            Caption         =   "Lugar"
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
            Height          =   255
            Index           =   7
            Left            =   180
            TabIndex        =   46
            Top             =   1170
            Width           =   1860
         End
         Begin VB.Label Label4 
            Caption         =   "Teléfono"
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
            Index           =   4
            Left            =   150
            TabIndex        =   45
            Top             =   2265
            Width           =   1860
         End
         Begin VB.Label Label4 
            Caption         =   "Observaciones"
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
            Index           =   1
            Left            =   150
            TabIndex        =   44
            Top             =   2655
            Width           =   2355
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "RECOGIDA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3075
         Left            =   390
         TabIndex        =   36
         Top             =   450
         Width           =   9285
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
            Left            =   1725
            MaxLength       =   250
            TabIndex        =   6
            Top             =   1920
            Width           =   7425
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
            Index           =   21
            Left            =   1725
            MaxLength       =   250
            TabIndex        =   5
            Top             =   1575
            Width           =   7425
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
            Left            =   1725
            MaxLength       =   250
            TabIndex        =   3
            Top             =   885
            Width           =   7425
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
            Left            =   1725
            MaxLength       =   100
            TabIndex        =   8
            Top             =   2655
            Width           =   7425
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
            Left            =   1725
            MaxLength       =   20
            TabIndex        =   7
            Top             =   2280
            Width           =   2925
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
            Index           =   7
            Left            =   1725
            MaxLength       =   250
            TabIndex        =   4
            Top             =   1230
            Width           =   7425
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
            Left            =   1725
            MaxLength       =   250
            TabIndex        =   2
            Top             =   540
            Width           =   7425
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
            Index           =   2
            Left            =   1725
            MaxLength       =   20
            TabIndex        =   0
            Top             =   195
            Width           =   2940
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
            Index           =   8
            Left            =   5985
            MaxLength       =   30
            TabIndex        =   1
            Top             =   195
            Width           =   3240
         End
         Begin VB.Label Label4 
            Caption         =   "Observaciones"
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
            Height          =   255
            Index           =   2
            Left            =   150
            TabIndex        =   42
            Top             =   2655
            Width           =   2355
         End
         Begin VB.Label Label4 
            Caption         =   "Teléfono"
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
            Height          =   255
            Index           =   6
            Left            =   150
            TabIndex        =   41
            Top             =   2295
            Width           =   1860
         End
         Begin VB.Label Label4 
            Caption         =   "Lugar"
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
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   40
            Top             =   1230
            Width           =   1860
         End
         Begin VB.Label Label4 
            Caption         =   "Mercancia"
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
            Height          =   255
            Index           =   3
            Left            =   180
            TabIndex        =   39
            Top             =   600
            Width           =   1365
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Carga"
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
            Height          =   255
            Index           =   16
            Left            =   180
            TabIndex        =   38
            Top             =   225
            Width           =   960
         End
         Begin VB.Label Label4 
            Caption         =   "Hora Carga"
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
            Height          =   255
            Index           =   5
            Left            =   4995
            TabIndex        =   37
            Top             =   225
            Width           =   960
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
         Left            =   11205
         TabIndex        =   33
         Top             =   11400
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
         Left            =   10020
         TabIndex        =   32
         Top             =   11400
         Width           =   1065
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   0
         Left            =   90
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   10125
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Precio"
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
         Height          =   255
         Index           =   18
         Left            =   540
         TabIndex        =   58
         Top             =   9735
         Width           =   1860
      End
      Begin VB.Label Label1 
         Caption         =   "Orden de Carga"
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
         TabIndex        =   35
         Top             =   105
         Width           =   5160
      End
   End
End
Attribute VB_Name = "frmVtasOrdenCarga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor Monica +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public NumCod As String 'Para indicar el numero de albaran
Public NomTrans As String 'indicamos el nombre de transportista
Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public DatosADevolverBusqueda As String    'Tindrà el nº de text que vol que torne, empipat
Public Event DatoSeleccionado(CadenaSeleccion As String)
    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

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
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim Tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean

Dim CodTipoMov As String
'Codigo tipo de movimiento en función del valor en la tabla de parámetros: stipom


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim Contador As Long
Dim vTipoMov As CTiposMov
Dim Mens As String
Dim ContCMR As String
Dim Sql As String


    InicializarVbles
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "pOrden=""" & NumCod & """|"
    numParam = numParam + 1
    
    cadParam = cadParam & "pRecFecha=""" & txtCodigo(2).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pRecHora=""" & txtCodigo(8).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pRecMer=""" & txtCodigo(3).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pRecMer1=""" & txtCodigo(20).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pRecLugar=""" & txtCodigo(7).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pRecLugar1=""" & txtCodigo(21).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pRecLugar2=""" & txtCodigo(22).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pRecTfno=""" & txtCodigo(0).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pRecObs=""" & txtCodigo(4).Text & """|"
    numParam = numParam + 1
    
    cadParam = cadParam & "pEnt1Fecha=""" & txtCodigo(10).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pEnt1Hora=""" & txtCodigo(11).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pEnt1Mer=""" & txtCodigo(9).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pEnt1Mer1=""" & txtCodigo(23).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pEnt1Lugar=""" & txtCodigo(6).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pEnt1Lugar1=""" & txtCodigo(24).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pEnt1Lugar2=""" & txtCodigo(25).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pEnt1Tfno=""" & txtCodigo(5).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pEnt1Obs=""" & txtCodigo(1).Text & """|"
    numParam = numParam + 1
    
    cadParam = cadParam & "pEnt2Fecha=""" & txtCodigo(13).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pEnt2Hora=""" & txtCodigo(12).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pEnt2Mer=""" & txtCodigo(14).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pEnt2Mer1=""" & txtCodigo(26).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pEnt2Lugar=""" & txtCodigo(15).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pEnt2Lugar1=""" & txtCodigo(27).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pEnt2Lugar2=""" & txtCodigo(33).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pEnt2Tfno=""" & txtCodigo(16).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pEnt2Obs=""" & txtCodigo(17).Text & """|"
    numParam = numParam + 1
    
    cadParam = cadParam & "pPrecio=""" & txtCodigo(19).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pFacTrans=""" & txtCodigo(18).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pFacTrans1=""" & txtCodigo(28).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pFacTrans2=""" & txtCodigo(29).Text & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pFacTrans3=""" & txtCodigo(30).Text & """|"
    numParam = numParam + 1
    
    'Nombre fichero .rpt a Imprimir
    
    indRPT = 10 'Impresion de Orden de Carga
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
    
    'Nombre fichero .rpt a Imprimir
    cadNombreRPT = nomDocu
    
    cadTitulo = "Orden de Carga"
    
    LlamarImprimir
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(2)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection
Dim Sql As String

    PrimeraVez = True
    limpiar Me

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, H, W
    indFrame = 5
    Tabla = "pedidos"
    
    imgAyuda(0).Picture = frmPpal.ImageListB.ListImages(10).Picture

    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.CmdCancel.Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub



Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           ' "____________________________________________________________"
            vCadena = "Si dejamos en blanco Facturacion del transporte a, imprime los datos " & vbCrLf & _
                      "de la empresa." & vbCrLf & vbCrLf & _
                      "En caso contrario se imprime lo que se indique en estas cuatro casillas." & vbCrLf & _
                      vbCrLf
                      
                      
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"
    
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
    
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 11970 '10485
        Me.FrameCobros.Width = 12525 '11085
        W = Me.FrameCobros.Width
        H = Me.FrameCobros.Height
    End If
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
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .Opcion = 0
        .EnvioEMail = False
        .ConSubInforme = True
        .Show vbModal
    End With
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


Private Function HayRegistros(cTabla As String, cWhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
Dim Rs As ADODB.Recordset

    Sql = "Select * FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    Sql = Sql & " group by 1 "
    Sql = Sql & " having sum(totalfac) > " & DBSet(txtCodigo(6).Text, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Rs.EOF Then
        MsgBox "No hay datos para mostrar en el Informe.", vbInformation
        HayRegistros = False
    Else
        HayRegistros = True
    End If

End Function


Private Function CargarVariedades(Mens As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim b As Boolean

    Sql = "delete from tmpinformes where codusu =" & vUsu.Codigo
    conn.Execute Sql
    
    Sql = "select sum(pesobrut) from albaran_variedad where numalbar = " & NumCod
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        If DBLet(Rs.Fields(0).Value, "N") > vParamAplic.LimPesoCMR Then
            Mens = "Repartir Pesos Brutos"
            b = RepartirBrutos(Mens, DBLet(Rs.Fields(0).Value, "N"))
        Else
            Mens = "Cargar Pesos Brutos"
            b = CargarBrutos(Mens)
        End If
    End If
    Set Rs = Nothing
    
    CargarVariedades = b
    
End Function

Private Function RepartirBrutos(Mens As String, SumaBrutos As Currency) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Importe As Currency
Dim CarImporte As String
Dim CadValues As String

    On Error GoTo eRepartirBrutos

    RepartirBrutos = False
    If SumaBrutos = 0 Then Exit Function

    Sql = "select numlinea, pesobrut from albaran_variedad where numalbar = " & NumCod
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Sql2 = "insert into tmpinformes (codusu, campo1, importe1) values "
    
    While Not Rs.EOF
        Importe = Round2(vParamAplic.LimPesoCMR * DBLet(Rs.Fields(1).Value, "N") / (SumaBrutos * 10), 0)
        CarImporte = CStr(Importe) & "0"
        
        CadValues = "(" & vUsu.Codigo & "," & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(CarImporte, "N") & ")"
        
        conn.Execute Sql2 & CadValues
    
        Rs.MoveNext
    Wend
    
    Rs.Close
    
    Set Rs = Nothing

eRepartirBrutos:
    If Err.Number = 0 Then
        RepartirBrutos = True
    End If
End Function

Private Function CargarBrutos(Mens As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Importe As Currency
Dim CarImporte As String
Dim CadValues As String
    
    On Error GoTo eCargarBrutos
    
    Sql = "select numlinea, pesobrut from albaran_variedad where numalbar = " & NumCod
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Sql2 = "insert into tmpinformes (codusu, campo1, importe1) values "
    
    While Not Rs.EOF
        Importe = Round2(DBLet(Rs.Fields(1).Value, "N") / 10, 0)
        CarImporte = CStr(Importe) & "0"
        
        CadValues = "(" & vUsu.Codigo & "," & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(CarImporte, "N") & ")"
        
        conn.Execute Sql2 & CadValues
        
        Rs.MoveNext
    Wend

eCargarBrutos:
    If Err.Number <> 0 Then
        Mens = Mens & vbCrLf & Err.Description
        CargarBrutos = False
    Else
        CargarBrutos = True
    End If
End Function

Private Function InsertarTemporal() As Boolean
Dim Sql As String
    
    On Error GoTo eInsertarTemporal
    
    InsertarTemporal = True

    conn.Execute "delete from tmpcmr where codusu = " & vUsu.Codigo
    
    Sql = "insert into tmpcmr(numlinea, codusu, numalbar) select numlinea, "
    Sql = Sql & vUsu.Codigo & "," & NumCod & " from tmpcopiascmr "
    conn.Execute Sql

eInsertarTemporal:
    If Err.Number <> 0 Then InsertarTemporal = False
End Function

Private Function ProvAgenciaTransporte() As String
Dim Sql As String
Dim Rs As ADODB.Recordset
    
    ProvAgenciaTransporte = ""
    
    Sql = "select protrans from agencias, albaran where albaran.numalbar = " & DBSet(NumCod, "N")
    Sql = Sql & " and albaran.codtrans = agencias.codtrans"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        ProvAgenciaTransporte = UCase(DBLet(Rs.Fields(0).Value, "T"))
    End If
    
End Function

