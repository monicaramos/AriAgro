VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmVtasInformes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   11595
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   10455
   Icon            =   "frmVtasInformes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11595
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCobros 
      Height          =   11700
      Left            =   -45
      TabIndex        =   24
      Top             =   -90
      Width           =   10515
      Begin VB.CheckBox Check1 
         Caption         =   "Comisionista l�neas Albar�n"
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
         Index           =   11
         Left            =   6570
         TabIndex        =   104
         Top             =   10260
         Width           =   3585
      End
      Begin VB.CheckBox Check1 
         Caption         =   "S�lo No Cobrados"
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
         Index           =   10
         Left            =   6570
         TabIndex        =   103
         Top             =   7125
         Width           =   3120
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Declaraci�n de Ventas"
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
         Left            =   6570
         TabIndex        =   102
         Top             =   9195
         Width           =   3435
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Detalle de Albaranes"
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
         Index           =   0
         Left            =   6570
         TabIndex        =   101
         Top             =   6450
         Width           =   3120
      End
      Begin VB.CheckBox Check1 
         Caption         =   "S�lo Facturados"
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
         Index           =   1
         Left            =   6570
         TabIndex        =   100
         Top             =   6780
         Width           =   3120
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Incluir Portes/Comis. en Gastos"
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
         Index           =   2
         Left            =   6570
         TabIndex        =   99
         Top             =   7470
         Width           =   3570
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Gastos Costes Reales"
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
         Index           =   3
         Left            =   6570
         TabIndex        =   98
         Top             =   7800
         Width           =   3120
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Agrupar Costes Confecci�n"
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
         Index           =   4
         Left            =   6570
         TabIndex        =   97
         Top             =   8160
         Width           =   3435
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Rendimiento por Calibre"
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
         Index           =   5
         Left            =   6570
         TabIndex        =   96
         Top             =   8490
         Width           =   3435
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Informe Albaranes Entrada"
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
         Index           =   7
         Left            =   6570
         TabIndex        =   95
         Top             =   8835
         Width           =   3435
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Exportar a Excel"
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
         Left            =   6570
         TabIndex        =   94
         Top             =   9915
         Width           =   3585
      End
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
         Left            =   5085
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Tag             =   "Tipo Variedad|N|N|||variedades|tipovariedad||N|"
         Top             =   10530
         Width           =   1710
      End
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
         Index           =   0
         Left            =   1935
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Tag             =   "Tipo Variedad|N|N|||variedades|tipovariedad||N|"
         Top             =   10530
         Width           =   1620
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Evoluci�n de Precios"
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
         Index           =   6
         Left            =   6570
         TabIndex        =   91
         Top             =   9555
         Width           =   3585
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
         Index           =   19
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   87
         Text            =   "Text5"
         Top             =   10065
         Width           =   3360
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
         Index           =   18
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   86
         Text            =   "Text5"
         Top             =   9690
         Width           =   3360
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
         Index           =   19
         Left            =   1785
         MaxLength       =   3
         TabIndex        =   19
         Top             =   10065
         Width           =   860
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
         Index           =   18
         Left            =   1785
         MaxLength       =   3
         TabIndex        =   18
         Top             =   9690
         Width           =   860
      End
      Begin VB.Frame FrameOrdenar 
         Caption         =   "Ordenar Por:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   6000
         Left            =   6255
         TabIndex        =   69
         Top             =   285
         Width           =   4080
         Begin VB.OptionButton optList1 
            Caption         =   "Cliente - Fecha"
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
            Index           =   15
            Left            =   270
            TabIndex        =   106
            Top             =   5670
            Width           =   3660
         End
         Begin VB.OptionButton optList1 
            Caption         =   "Comisionista - Variedad - Fecha"
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
            Index           =   14
            Left            =   270
            TabIndex        =   105
            Top             =   5310
            Width           =   3615
         End
         Begin VB.OptionButton optList1 
            Caption         =   "Variedad - Variedad Comercial"
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
            Index           =   13
            Left            =   270
            TabIndex        =   85
            Top             =   4965
            Width           =   3645
         End
         Begin VB.OptionButton optList1 
            Caption         =   "Gastos Variedad-Confecci�n"
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
            Index           =   12
            Left            =   270
            TabIndex        =   84
            Top             =   3900
            Width           =   3585
         End
         Begin VB.OptionButton optList1 
            Caption         =   "Tipos de Mercado"
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
            Index           =   11
            Left            =   270
            TabIndex        =   81
            Top             =   4605
            Width           =   3540
         End
         Begin VB.OptionButton optList1 
            Caption         =   "Salidas por Calibre"
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
            Index           =   10
            Left            =   270
            TabIndex        =   80
            Top             =   4260
            Width           =   3540
         End
         Begin VB.OptionButton optList1 
            Caption         =   "Gastos Confecci�n-Variedad"
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
            Index           =   9
            Left            =   270
            TabIndex        =   79
            Top             =   3540
            Width           =   3585
         End
         Begin VB.OptionButton optList1 
            Caption         =   "Pa�s - Variedad"
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
            Index           =   8
            Left            =   270
            TabIndex        =   78
            Top             =   3195
            Width           =   3540
         End
         Begin VB.OptionButton optList1 
            Caption         =   "Variedad - Pa�s"
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
            Index           =   7
            Left            =   270
            TabIndex        =   77
            Top             =   2835
            Width           =   3540
         End
         Begin VB.OptionButton optList1 
            Caption         =   "Marca - Variedad"
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
            Index           =   6
            Left            =   270
            TabIndex        =   76
            Top             =   2490
            Width           =   3540
         End
         Begin VB.OptionButton optList1 
            Caption         =   "Variedad - Marca"
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
            Index           =   5
            Left            =   270
            TabIndex        =   75
            Top             =   2130
            Width           =   3540
         End
         Begin VB.OptionButton optList1 
            Caption         =   "Confecci�n - Variedad"
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
            Index           =   4
            Left            =   270
            TabIndex        =   74
            Top             =   1770
            Width           =   3540
         End
         Begin VB.OptionButton optList1 
            Caption         =   "Variedad - Confecci�n"
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
            Index           =   3
            Left            =   270
            TabIndex        =   73
            Top             =   1425
            Width           =   3540
         End
         Begin VB.OptionButton optList1 
            Caption         =   "Cliente - Destino - Variedad"
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
            Index           =   2
            Left            =   270
            TabIndex        =   72
            Top             =   1065
            Width           =   3540
         End
         Begin VB.OptionButton optList1 
            Caption         =   "Variedad - Fecha"
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
            Index           =   0
            Left            =   270
            TabIndex        =   71
            Top             =   360
            Value           =   -1  'True
            Width           =   3360
         End
         Begin VB.OptionButton optList1 
            Caption         =   "Variedad - Cliente - Destino"
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
            Index           =   1
            Left            =   270
            TabIndex        =   70
            Top             =   720
            Width           =   3540
         End
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
         Index           =   14
         Left            =   1785
         MaxLength       =   3
         TabIndex        =   16
         Top             =   8700
         Width           =   860
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
         Index           =   15
         Left            =   1785
         MaxLength       =   3
         TabIndex        =   17
         Top             =   9075
         Width           =   860
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
         Index           =   14
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   65
         Text            =   "Text5"
         Top             =   8700
         Width           =   3360
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
         Index           =   15
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   64
         Text            =   "Text5"
         Top             =   9075
         Width           =   3360
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
         Left            =   1785
         MaxLength       =   3
         TabIndex        =   14
         Top             =   7800
         Width           =   860
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
         Index           =   13
         Left            =   1785
         MaxLength       =   3
         TabIndex        =   15
         Top             =   8175
         Width           =   860
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
         Index           =   12
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   60
         Text            =   "Text5"
         Top             =   7800
         Width           =   3360
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
         Index           =   13
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   59
         Text            =   "Text5"
         Top             =   8175
         Width           =   3360
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
         Left            =   1785
         MaxLength       =   3
         TabIndex        =   12
         Top             =   6810
         Width           =   860
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
         Left            =   1785
         MaxLength       =   3
         TabIndex        =   13
         Top             =   7185
         Width           =   860
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
         Index           =   10
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   55
         Text            =   "Text5"
         Top             =   6810
         Width           =   3360
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
         Index           =   11
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   54
         Text            =   "Text5"
         Top             =   7185
         Width           =   3360
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
         Index           =   7
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   50
         Text            =   "Text5"
         Top             =   4275
         Width           =   3360
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
         Index           =   6
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   49
         Text            =   "Text5"
         Top             =   3900
         Width           =   3360
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
         Left            =   1785
         MaxLength       =   4
         TabIndex        =   7
         Top             =   4275
         Width           =   860
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
         Index           =   6
         Left            =   1785
         MaxLength       =   4
         TabIndex        =   6
         Top             =   3915
         Width           =   860
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
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   0
         Top             =   1020
         Width           =   860
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
         Left            =   1785
         MaxLength       =   3
         TabIndex        =   1
         Top             =   1410
         Width           =   860
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
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "Text5"
         Top             =   1020
         Width           =   3360
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
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "Text5"
         Top             =   1410
         Width           =   3360
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
         Index           =   9
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "Text5"
         Top             =   6240
         Width           =   2685
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
         Index           =   8
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "Text5"
         Top             =   5865
         Width           =   2685
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
         Left            =   1785
         MaxLength       =   6
         TabIndex        =   11
         Top             =   6240
         Width           =   1545
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
         Left            =   1785
         MaxLength       =   16
         TabIndex        =   10
         Top             =   5865
         Width           =   1545
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
         Index           =   3
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "Text5"
         Top             =   2385
         Width           =   3360
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
         Index           =   2
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "Text5"
         Top             =   1995
         Width           =   3360
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
         Left            =   1785
         MaxLength       =   6
         TabIndex        =   3
         Top             =   2385
         Width           =   860
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
         Left            =   1785
         MaxLength       =   6
         TabIndex        =   2
         Top             =   1995
         Width           =   860
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
         Index           =   17
         Left            =   1785
         MaxLength       =   10
         TabIndex        =   9
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   5265
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
         Index           =   16
         Left            =   1785
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   4860
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
         Left            =   9180
         TabIndex        =   23
         Top             =   11040
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
         Left            =   8010
         TabIndex        =   22
         Top             =   11040
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
         Index           =   4
         Left            =   1785
         MaxLength       =   6
         TabIndex        =   4
         Top             =   2955
         Width           =   860
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
         Left            =   1785
         MaxLength       =   6
         TabIndex        =   5
         Top             =   3330
         Width           =   860
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
         TabIndex        =   26
         Text            =   "Text5"
         Top             =   2955
         Width           =   3360
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
         TabIndex        =   25
         Text            =   "Text5"
         Top             =   3330
         Width           =   3360
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   360
         TabIndex        =   82
         Top             =   11040
         Width           =   7410
         _ExtentX        =   13070
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   5
         Left            =   6150
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   10230
         Width           =   240
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   4
         Left            =   6180
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   7125
         Width           =   240
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   0
         Left            =   6150
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   9540
         Width           =   240
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   2
         Left            =   6150
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   9885
         Width           =   240
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   3
         Left            =   6150
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   9195
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Variedad"
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
         Index           =   32
         Left            =   3600
         TabIndex        =   93
         Top             =   10575
         Width           =   1350
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   1
         Left            =   6150
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   8820
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Mercancia"
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
         Index           =   31
         Left            =   360
         TabIndex        =   92
         Top             =   10530
         Width           =   1500
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   17
         Left            =   1485
         MouseIcon       =   "frmVtasInformes.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar comisionista"
         Top             =   10065
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   16
         Left            =   1485
         MouseIcon       =   "frmVtasInformes.frx":015E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar comisionista"
         Top             =   9690
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Comisionista"
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
         TabIndex        =   90
         Top             =   9405
         Width           =   1215
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
         Height          =   195
         Index           =   29
         Left            =   720
         TabIndex        =   89
         Top             =   10065
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
         Height          =   195
         Index           =   28
         Left            =   720
         TabIndex        =   88
         Top             =   9690
         Width           =   690
      End
      Begin VB.Label Label4 
         Caption         =   "Cargando tabla temporal"
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
         Index           =   27
         Left            =   360
         TabIndex        =   83
         Top             =   11310
         Width           =   7350
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
         Height          =   195
         Index           =   26
         Left            =   720
         TabIndex        =   68
         Top             =   8700
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
         Height          =   195
         Index           =   25
         Left            =   720
         TabIndex        =   67
         Top             =   9075
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pa�s"
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
         Index           =   24
         Left            =   360
         TabIndex        =   66
         Top             =   8415
         Width           =   390
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   14
         Left            =   1485
         MouseIcon       =   "frmVtasInformes.frx":02B0
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar pa�s"
         Top             =   8700
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   15
         Left            =   1485
         MouseIcon       =   "frmVtasInformes.frx":0402
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar pa�s"
         Top             =   9075
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
         Height          =   195
         Index           =   23
         Left            =   720
         TabIndex        =   63
         Top             =   7800
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
         Height          =   195
         Index           =   22
         Left            =   720
         TabIndex        =   62
         Top             =   8175
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Mercado"
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
         Index           =   21
         Left            =   360
         TabIndex        =   61
         Top             =   7515
         Width           =   1650
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   12
         Left            =   1485
         MouseIcon       =   "frmVtasInformes.frx":0554
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar mercado"
         Top             =   7800
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   13
         Left            =   1485
         MouseIcon       =   "frmVtasInformes.frx":06A6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar mercado"
         Top             =   8175
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
         Height          =   195
         Index           =   20
         Left            =   765
         TabIndex        =   58
         Top             =   6810
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
         Height          =   195
         Index           =   19
         Left            =   765
         TabIndex        =   57
         Top             =   7185
         Width           =   645
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
         Index           =   18
         Left            =   360
         TabIndex        =   56
         Top             =   6525
         Width           =   600
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   10
         Left            =   1485
         MouseIcon       =   "frmVtasInformes.frx":07F8
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar marca"
         Top             =   6810
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   11
         Left            =   1485
         MouseIcon       =   "frmVtasInformes.frx":094A
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar marca"
         Top             =   7185
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   1485
         MouseIcon       =   "frmVtasInformes.frx":0A9C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar destino"
         Top             =   4275
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1485
         MouseIcon       =   "frmVtasInformes.frx":0BEE
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar destino"
         Top             =   3900
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Destino"
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
         TabIndex        =   53
         Top             =   3600
         Width           =   735
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
         Height          =   195
         Index           =   10
         Left            =   735
         TabIndex        =   52
         Top             =   4275
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
         Height          =   195
         Index           =   9
         Left            =   735
         TabIndex        =   51
         Top             =   3900
         Width           =   690
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
         Height          =   195
         Index           =   8
         Left            =   765
         TabIndex        =   48
         Top             =   1020
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
         Height          =   195
         Index           =   7
         Left            =   765
         TabIndex        =   47
         Top             =   1395
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
         Height          =   240
         Index           =   6
         Left            =   360
         TabIndex        =   46
         Top             =   735
         Width           =   525
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1485
         MouseIcon       =   "frmVtasInformes.frx":0D40
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   1020
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1485
         MouseIcon       =   "frmVtasInformes.frx":0E92
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   1395
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Informes de Ventas"
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
         Left            =   270
         TabIndex        =   43
         Top             =   225
         Width           =   6240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   1485
         MouseIcon       =   "frmVtasInformes.frx":0FE4
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar forfait"
         Top             =   6240
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   1485
         MouseIcon       =   "frmVtasInformes.frx":1136
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar forfait"
         Top             =   5850
         Width           =   240
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
         Height          =   240
         Index           =   5
         Left            =   360
         TabIndex        =   42
         Top             =   5580
         Width           =   645
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
         Height          =   195
         Index           =   4
         Left            =   765
         TabIndex        =   41
         Top             =   6240
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
         Height          =   195
         Index           =   3
         Left            =   765
         TabIndex        =   40
         Top             =   5865
         Width           =   690
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1485
         MouseIcon       =   "frmVtasInformes.frx":1288
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   2370
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1485
         MouseIcon       =   "frmVtasInformes.frx":13DA
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   1995
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
         Height          =   240
         Index           =   2
         Left            =   360
         TabIndex        =   37
         Top             =   1710
         Width           =   855
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
         Height          =   195
         Index           =   1
         Left            =   765
         TabIndex        =   36
         Top             =   2370
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
         Height          =   195
         Index           =   0
         Left            =   765
         TabIndex        =   35
         Top             =   1995
         Width           =   690
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Index           =   16
         Left            =   360
         TabIndex        =   32
         Top             =   4590
         Width           =   1815
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
         Height          =   195
         Index           =   15
         Left            =   765
         TabIndex        =   31
         Top             =   4860
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
         Height          =   195
         Index           =   14
         Left            =   765
         TabIndex        =   30
         Top             =   5265
         Width           =   645
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1485
         Picture         =   "frmVtasInformes.frx":152C
         ToolTipText     =   "Buscar fecha"
         Top             =   4860
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   1485
         Picture         =   "frmVtasInformes.frx":15B7
         ToolTipText     =   "Buscar fecha"
         Top             =   5265
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
         Height          =   195
         Index           =   13
         Left            =   735
         TabIndex        =   29
         Top             =   2955
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
         Height          =   195
         Index           =   12
         Left            =   735
         TabIndex        =   28
         Top             =   3330
         Width           =   645
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
         Index           =   11
         Left            =   360
         TabIndex        =   27
         Top             =   2655
         Width           =   675
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1485
         MouseIcon       =   "frmVtasInformes.frx":1642
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   2955
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1485
         MouseIcon       =   "frmVtasInformes.frx":1794
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   3330
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7920
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmVtasInformes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MANOLO +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar n� oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexi�n a BD Ariges  2.- Conexi�n a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmCla As frmManClases 'Clases
Attribute frmCla.VB_VarHelpID = -1
Private WithEvents frmVar As frmManVariedad 'Variedad
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmCli As frmClientes 'Clientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmDes As frmDestCli 'Destinos de Clientes
Attribute frmDes.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
'Private WithEvents frmFor As frmManForfaits 'Forfaits
Private WithEvents frmFor As frmBasico2
Attribute frmFor.VB_VarHelpID = -1
Private WithEvents frmMar As frmManMarcas 'Marcas
Attribute frmMar.VB_VarHelpID = -1
Private WithEvents frmTMe As frmManTipMerc 'Tipos de Mercado
Attribute frmTMe.VB_VarHelpID = -1
Private WithEvents frmPais As frmManPaises 'Paises
Attribute frmPais.VB_VarHelpID = -1
Private WithEvents frmComis As frmManAgencias 'comisionistas
Attribute frmComis.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'mensajes
Attribute frmMens.VB_VarHelpID = -1
Private WithEvents frmMensCate As frmMensajes 'mensajes
Attribute frmMensCate.VB_VarHelpID = -1
Private WithEvents frmMensContr As frmMensajes 'Contratos
Attribute frmMensContr.VB_VarHelpID = -1


Private WithEvents frmMensClase As frmMensajes 'mensajes
Attribute frmMensClase.VB_VarHelpID = -1
Private WithEvents frmMensVariedad As frmMensajes 'mensajes
Attribute frmMensVariedad.VB_VarHelpID = -1
Private WithEvents frmMensCliente As frmMensajes 'mensajes
Attribute frmMensCliente.VB_VarHelpID = -1
Private WithEvents frmMensDestino As frmMensajes 'mensajes
Attribute frmMensDestino.VB_VarHelpID = -1
Private WithEvents frmMensForfait As frmMensajes 'mensajes
Attribute frmMensForfait.VB_VarHelpID = -1
Private WithEvents frmMensMarca As frmMensajes 'mensajes
Attribute frmMensMarca.VB_VarHelpID = -1
Private WithEvents frmMensMercado As frmMensajes 'mensajes
Attribute frmMensMercado.VB_VarHelpID = -1
Private WithEvents frmMensPais As frmMensajes 'mensajes
Attribute frmMensPais.VB_VarHelpID = -1
Private WithEvents frmMensComisionista As frmMensajes 'mensajes
Attribute frmMensComisionista.VB_VarHelpID = -1


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

Dim PrimeraVez As Boolean
Dim ConSubInforme As Boolean

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub Check1_Click(Index As Integer)
    '[Monica]26/06/2012: si es informe albaranes entradas no es declaracion de ventas
    If Index = 7 And Check1(7).Value = 1 Then Check1(9).Value = 0
    If Index = 9 And Check1(9).Value = 1 Then Check1(7).Value = 0
End Sub

Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim vSQL As String

Dim vSqlClase As String
Dim vsqlVariedad As String
Dim vSqlCliente As String
Dim vsqlDestino As String
Dim vSqlForfait As String
Dim vSqlMarca As String
Dim vSqlMercado As String
Dim vSqlPais As String
Dim vSqlComisionista As String
Dim Tipo As Byte



    InicializarVbles
    
    '========= PARAMETROS  =============================
    'A�adir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
     '======== FORMULA  ====================================
    'Seleccionar registros de la empresa conectada
'    Codigo = "{" & tabla & ".codempre}=" & vEmpresa.codEmpre
'    If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
'    If Not AnyadirAFormula(cadSelect, Codigo) Then Exit Sub
    
    'D/H Clase
    cDesde = Trim(txtCodigo(0).Text)
    cHasta = Trim(txtCodigo(1).Text)
    nDesde = txtNombre(0).Text
    nHasta = txtNombre(1).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{variedades.codclase}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase= """) Then Exit Sub
    End If
    
    vSqlClase = ""
    If txtCodigo(0).Text <> "" Then vSqlClase = vSqlClase & " and clases.codclase >= " & DBSet(txtCodigo(0).Text, "N")
    If txtCodigo(1).Text <> "" Then vSqlClase = vSqlClase & " and clases.codclase <= " & DBSet(txtCodigo(1).Text, "N")
    
'    vSql = ""
'    If txtCodigo(0).Text <> "" Then vSql = vSql & " and variedades.codclase >= " & DBSet(txtCodigo(0).Text, "N")
'    If txtCodigo(1).Text <> "" Then vSql = vSql & " and variedades.codclase <= " & DBSet(txtCodigo(1).Text, "N")

    'D/H Variedades
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    nDesde = txtNombre(2).Text
    nHasta = txtNombre(3).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{albaran_variedad.codvarie}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad= """) Then Exit Sub
    End If
    
    vsqlVariedad = ""
    If txtCodigo(2).Text <> "" Then vsqlVariedad = vsqlVariedad & " and variedades.codvarie >= " & DBSet(txtCodigo(2).Text, "N")
    If txtCodigo(3).Text <> "" Then vsqlVariedad = vsqlVariedad & " and variedades.codvarie <= " & DBSet(txtCodigo(3).Text, "N")
    
    '[Monica]23/02/2011: a�adimos la condicion del tipo de mercancia de la variedad
    If vsqlVariedad <> "" Then
        Select Case Combo1(0).ListIndex
            Case 0 ' cooperativa
                vsqlVariedad = vsqlVariedad & " and variedades.tipovariedad = 0"
            Case 1 ' ajenas
                vsqlVariedad = vsqlVariedad & " and variedades.tipovariedad = 1"
            Case 2 ' todas
                ' sin condicion
        End Select
    Else
        Select Case Combo1(0).ListIndex
            Case 0 ' cooperativa
                If Not AnyadirAFormula(cadFormula, "{variedades.tipovariedad} = 0") Then Exit Sub
                If Not AnyadirAFormula(cadselect, "{variedades.tipovariedad} = 0") Then Exit Sub
            Case 1 ' ajenas
                If Not AnyadirAFormula(cadFormula, "{variedades.tipovariedad} = 1") Then Exit Sub
                If Not AnyadirAFormula(cadselect, "{variedades.tipovariedad} = 1") Then Exit Sub
            Case 2 ' todas
                ' sin condicion
        End Select
    End If
    
    
    'D/H Cliente
    cDesde = Trim(txtCodigo(4).Text)
    cHasta = Trim(txtCodigo(5).Text)
    nDesde = txtNombre(4).Text
    nHasta = txtNombre(5).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{albaran.codclien}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCliente= """) Then Exit Sub
    End If
    
    vSqlCliente = ""
    If txtCodigo(4).Text <> "" Then vSqlCliente = vSqlCliente & " and clientes.codclien >= " & DBSet(txtCodigo(4).Text, "N")
    If txtCodigo(5).Text <> "" Then vSqlCliente = vSqlCliente & " and clientes.codclien <= " & DBSet(txtCodigo(5).Text, "N")
    
    
    'D/H Destino
    cDesde = Trim(txtCodigo(6).Text)
    cHasta = Trim(txtCodigo(7).Text)
    nDesde = txtNombre(6).Text
    nHasta = txtNombre(7).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{albaran.coddesti}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHDestino= """) Then Exit Sub
    End If
    
    vsqlDestino = ""
    If txtCodigo(6).Text <> "" Then vsqlDestino = vsqlDestino & " and destinos.coddesti >= " & DBSet(txtCodigo(6).Text, "N")
    If txtCodigo(7).Text <> "" Then vsqlDestino = vsqlDestino & " and destinos.coddesti <= " & DBSet(txtCodigo(7).Text, "N")

    
    
    'D/H Fecha albaran
    cDesde = Trim(txtCodigo(16).Text)
    cHasta = Trim(txtCodigo(17).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{albaran.fechaalb}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    'D/H Forfaits
    cDesde = Trim(txtCodigo(8).Text)
    cHasta = Trim(txtCodigo(9).Text)
    nDesde = txtNombre(8).Text
    nHasta = txtNombre(9).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{albaran_variedad.codforfait}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHForfait= """) Then Exit Sub
    End If
    
    vSqlForfait = ""
    If txtCodigo(8).Text <> "" Then vSqlForfait = vSqlForfait & " and forfaits.codforfait >= " & DBSet(txtCodigo(8).Text, "T")
    If txtCodigo(9).Text <> "" Then vSqlForfait = vSqlForfait & " and forfaits.codforfait <= " & DBSet(txtCodigo(9).Text, "T")
    
    'D/H Marca
    cDesde = Trim(txtCodigo(10).Text)
    cHasta = Trim(txtCodigo(11).Text)
    nDesde = txtNombre(10).Text
    nHasta = txtNombre(11).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{albaran_variedad.codmarca}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHMarca= """) Then Exit Sub
    End If
    vSqlMarca = ""
    If txtCodigo(10).Text <> "" Then vSqlMarca = vSqlMarca & " and marcas.codmarca >= " & DBSet(txtCodigo(10).Text, "N")
    If txtCodigo(11).Text <> "" Then vSqlMarca = vSqlMarca & " and marcas.codmarca <= " & DBSet(txtCodigo(11).Text, "N")
    
    'D/H Tipo de Mercado
    cDesde = Trim(txtCodigo(12).Text)
    cHasta = Trim(txtCodigo(13).Text)
    nDesde = txtNombre(12).Text
    nHasta = txtNombre(13).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{albaran.codtimer}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHMercado= """) Then Exit Sub
    End If
    vSqlMercado = ""
    If txtCodigo(12).Text <> "" Then vSqlMercado = vSqlMercado & " and tipomer.codtimer >= " & DBSet(txtCodigo(12).Text, "N")
    If txtCodigo(13).Text <> "" Then vSqlMercado = vSqlMercado & " and tipomer.codtimer <= " & DBSet(txtCodigo(13).Text, "N")
    
    'D/H pais
    cDesde = Trim(txtCodigo(14).Text)
    cHasta = Trim(txtCodigo(15).Text)
    nDesde = txtNombre(14).Text
    nHasta = txtNombre(15).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{destinos.codpaise}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHPais= """) Then Exit Sub
    End If
    vSqlPais = ""
    If txtCodigo(14).Text <> "" Then vSqlPais = vSqlPais & " and paises.codpaise >= " & DBSet(txtCodigo(14).Text, "N")
    If txtCodigo(15).Text <> "" Then vSqlPais = vSqlPais & " and paises.codpaise <= " & DBSet(txtCodigo(15).Text, "N")
    
    'D/H comisionista
    cDesde = Trim(txtCodigo(18).Text)
    cHasta = Trim(txtCodigo(19).Text)
    nDesde = txtNombre(18).Text
    nHasta = txtNombre(19).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        '[Monica]12/12/2013: si esta marcado se refiere al de las l�neas
        If Check1(11).Value = 1 Then
            Codigo = "{albaran_variedad.codcomis}"
        Else
            Codigo = "{albaran.codcomis}"
        End If
        
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHComisionista= """) Then Exit Sub
    End If
    vSqlComisionista = ""
    If txtCodigo(18).Text <> "" Then vSqlComisionista = vSqlComisionista & " and agencias.codtrans >= " & DBSet(txtCodigo(18).Text, "N")
    If txtCodigo(19).Text <> "" Then vSqlComisionista = vSqlComisionista & " and agencias.codtrans <= " & DBSet(txtCodigo(19).Text, "N")
    
    'tipo de variedad
    Select Case Combo1(1).ListIndex
        Case 0 ' todas
        Case Is >= 1 '0=convencional 1=bilogico 2=reconversion
            If Not AnyadirAFormula(cadFormula, "{albaran_variedad.codtipo} = " & Combo1(1).ListIndex - 1) Then Exit Sub
            If Not AnyadirAFormula(cadselect, "{albaran_variedad.codtipo} = " & Combo1(1).ListIndex - 1) Then Exit Sub
    End Select
    
    
    If vSqlClase <> "" And txtCodigo(0).Text <> txtCodigo(1).Text Then
        Set frmMensClase = New frmMensajes
    
        frmMensClase.OpcionMensaje = 21
        frmMensClase.Label5 = "Clases"
        frmMensClase.cadwhere = vSqlClase
        frmMensClase.Show vbModal
    
        Set frmMensClase = Nothing
    End If
    
    If vsqlVariedad <> "" And txtCodigo(2).Text <> txtCodigo(3).Text Then
        Set frmMensVariedad = New frmMensajes
    
        frmMensVariedad.OpcionMensaje = 21
        frmMensVariedad.Label5 = "Variedades"
        frmMensVariedad.cadwhere = vsqlVariedad
        frmMensVariedad.Show vbModal
    
        Set frmMensVariedad = Nothing
    End If
    
    If vSqlCliente <> "" And txtCodigo(4).Text <> txtCodigo(5).Text Then
        Set frmMensCliente = New frmMensajes
    
        frmMensCliente.OpcionMensaje = 21
        frmMensCliente.Label5 = "Clientes"
        frmMensCliente.cadwhere = vSqlCliente
        frmMensCliente.Show vbModal
    
        Set frmMensCliente = Nothing
    End If
    
    If vsqlDestino <> "" And txtCodigo(6).Text <> txtCodigo(7).Text And txtCodigo(4).Text = txtCodigo(5).Text And txtCodigo(4).Text <> "" Then
        Set frmMensDestino = New frmMensajes

        frmMensDestino.OpcionMensaje = 21
        frmMensDestino.Label5 = "Destinos"
        frmMensDestino.cadwhere = vsqlDestino & " and destinos.codclien = " & txtCodigo(4).Text
        frmMensDestino.Show vbModal

        Set frmMensDestino = Nothing
    End If
    
    If vSqlForfait <> "" And txtCodigo(8).Text <> txtCodigo(9).Text Then
        Set frmMensForfait = New frmMensajes
    
        frmMensForfait.OpcionMensaje = 21
        frmMensForfait.Label5 = "Forfaits"
        frmMensForfait.cadwhere = vSqlForfait
        frmMensForfait.Show vbModal
    
        Set frmMensForfait = Nothing
    End If
    
    If vSqlMarca <> "" And txtCodigo(10).Text <> txtCodigo(11).Text Then
        Set frmMensMarca = New frmMensajes
    
        frmMensMarca.OpcionMensaje = 21
        frmMensMarca.Label5 = "Marcas"
        frmMensMarca.cadwhere = vSqlMarca
        frmMensMarca.Show vbModal
    
        Set frmMensMarca = Nothing
    End If
    
    If vSqlMercado <> "" And txtCodigo(12).Text <> txtCodigo(13).Text Then
        Set frmMensMercado = New frmMensajes
    
        frmMensMercado.OpcionMensaje = 21
        frmMensMercado.Label5 = "Mercados"
        frmMensMercado.cadwhere = vSqlMercado
        frmMensMercado.Show vbModal
    
        Set frmMensMercado = Nothing
    End If
    
    If vSqlPais <> "" And txtCodigo(14).Text <> txtCodigo(15).Text Then
        Set frmMensPais = New frmMensajes
    
        frmMensPais.OpcionMensaje = 21
        frmMensPais.Label5 = "Paises"
        frmMensPais.cadwhere = vSqlPais
        frmMensPais.Show vbModal
    
        Set frmMensPais = Nothing
    End If
    
    If vSqlComisionista <> "" And txtCodigo(18).Text <> txtCodigo(19).Text Then
        Set frmMensComisionista = New frmMensajes
    
        frmMensComisionista.OpcionMensaje = 21
        frmMensComisionista.Label5 = "Comisionistas"
        frmMensComisionista.cadwhere = vSqlComisionista
        frmMensComisionista.Show vbModal
    
        Set frmMensComisionista = Nothing
    End If
    
    '[Monica]26/06/2013: sacamos cuales son las distintas categorias que aparecen
    Set frmMensCate = New frmMensajes
    
    frmMensCate.OpcionMensaje = 21
    frmMensCate.Label5 = "Categorias"
    frmMensCate.cadwhere = ""
    frmMensCate.Show vbModal
    
    Set frmMensCate = Nothing
    
    
    '[Monica]17/10/2016: nuevo informe de albaranes por categoria
'[Monica]25/04/2018: para cualquier tipo de informe
'    If optList1(15).Value Then
        Set frmMensContr = New frmMensajes
        
        frmMensContr.OpcionMensaje = 21
        frmMensContr.Label5 = "Contratos"
        frmMensContr.cadwhere = ""
        frmMensContr.Show vbModal
        
        Set frmMensContr = Nothing
'    End If
    
    ' detalle de albaranes
    If Check1(0).Value Then
        cadParam = cadParam & "pDetalle=1|"
    Else
        cadParam = cadParam & "pDetalle=0|"
    End If
    numParam = numParam + 1
    
    ' incluir gastos de portes
    If Check1(2).Value Then ' se incluyen los gastos de portes
        cadParam = cadParam & "pGastosPor=1|"
    Else
        cadParam = cadParam & "pGastosPor=0|"
    End If
    numParam = numParam + 1
    
    ' solo facturados
    If Check1(1).Value Then
        cadParam = cadParam & "pTipo=""S�lo Facturados""|"
    Else
        cadParam = cadParam & "pTipo=""Facturados y No Facturados""|"
    End If
    cadTABLA = tabla & " INNER JOIN albaran_variedad ON albaran.numalbar = albaran_variedad.numalbar "
    cadTABLA = "(" & cadTABLA & ") INNER JOIN variedades ON albaran_variedad.codvarie = variedades.codvarie "
    cadTABLA = "(" & cadTABLA & ") INNER JOIN destinos ON albaran.codclien = destinos.codclien and albaran.coddesti = destinos.coddesti "
    cadTABLA = "(" & cadTABLA & ") INNER JOIN forfaits ON albaran_variedad.codforfait = forfaits.codforfait "
    ' solo los facturados
    If Check1(1).Value Then
        cadTABLA = "(" & cadTABLA & ") INNER JOIN facturas_variedad ON albaran_variedad.numalbar = facturas_variedad.numalbar and albaran_variedad.numlinea = facturas_variedad.numlinealbar "
    Else
        cadTABLA = "(" & cadTABLA & ") LEFT JOIN facturas_variedad ON albaran_variedad.numalbar = facturas_variedad.numalbar and albaran_variedad.numlinea = facturas_variedad.numlinealbar "
    End If
    
    cadFormula = "{tmpinfventas.codusu} = " & vUsu.Codigo
'    If optList1(10).Value Then
'        If ProcesarCambiosCalibres(cadTabla, cadSelect) Then
'            cadTitulo = "Albaranes de Venta"
'            cadNombreRPT = "rAlbaranVta11.rpt"
'            LlamarImprimir
'        End If
'    Else

    ' ++monica: 16/03/2009
    ' a�adido: el listado donde las variedades son difentes a las variedades comerciales
    If optList1(13).Value Then
        If cadselect <> "" Then cadselect = cadselect & " and "
        cadselect = cadselect & " albaran_variedad.codvarie <> albaran_variedad.codvarco "
    End If
    ' ++

    If HayRegistros(cadTABLA, cadselect) Then
        If Check1(6).Value = 1 Then ' proceso de informe de evolucion de precios
            If ProcesarCambiosEvolucion(cadTABLA, cadselect) Then
                cadTitulo = "Evoluci�n Precios Albaranes de Venta"
                ConSubInforme = False
                If optList1(0).Value Then
                    cadNombreRPT = "rAlbaranVta1a.rpt"
                    cadParam = cadParam & "pOrden=""Variedad - Fecha""|"
                    numParam = numParam + 1
                End If
                
                If optList1(2).Value Then
                    cadNombreRPT = "rAlbaranVta3a.rpt"
                    cadParam = cadParam & "pOrden=""Cliente - Destino - Variedad""|"
                    numParam = numParam + 1
                End If
            
                LlamarImprimir
            End If
            
            Exit Sub
        End If
        
        
        If ProcesarCambios(cadTABLA, cadselect) Then
            '[Monica]16/11/2011: en el caso de la salida a Excel
            If Check1(8).Value Then
                '[Monica]19/11/2015: insertamos las calidades para el caso de catadau sacar una hoja diferente
                If vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Then
                    If CargarTemporal Then
                        If HayRegistros("tmpinformes", "codusu=" & vUsu.Codigo) Then
                            If Dir(App.path & "\Ventas.exe", vbArchive) <> "" And _
                               Dir(App.path & "\PlantillaVtaCatadau.xls", vbArchive) <> "" And _
                               Dir(App.path & "\ControlVtas.cfg", vbArchive) <> "" Then
                                Shell App.path & "\ventas.exe /E|" & vUsu.CadenaConexion & "|" & vUsu.Codigo & "|", vbNormalFocus
                            Else
                                MsgBox "No tiene los ficheros necesarios para realizar el proceso. Llame a Ariadna", vbExclamation
                            End If
                        
                        End If
                    
                    End If
                Else
                    If Dir(App.path & "\Ventas.exe", vbArchive) <> "" And _
                       Dir(App.path & "\PlantillaVta.xls", vbArchive) <> "" And _
                       Dir(App.path & "\ControlVtas.cfg", vbArchive) <> "" Then
                        Shell App.path & "\ventas.exe /E|" & vUsu.CadenaConexion & "|" & vUsu.Codigo & "|", vbNormalFocus
                    Else
                        MsgBox "No tiene los ficheros necesarios para realizar el proceso. Llame a Ariadna", vbExclamation
                    End If
                End If
            Else
              
                cadTitulo = "Albaranes de Venta"
                ConSubInforme = False
                
                If optList1(0).Value Then
                     'Nombre fichero .rpt a Imprimir
                     cadNombreRPT = "rAlbaranVta1.rpt"
    '[Monica]19/09/2011: en el caso de pedir un informe de evolucion de precios no provisional/definitivo
    '                 If Check1(6).Value = 1 Then
    '                    cadNombreRPT = "rAlbaranVta1a.rpt"
    '                    If Option1(1).Value Then
    '                        cadParam = cadParam & "pPrecDef=1|"
    '                    Else
    '                        cadParam = cadParam & "pPrecDef=0|"
    '                    End If
    '                    numParam = numParam + 1
    '                 End If
    
                     If Check1(7).Value = 1 Then
                        cadNombreRPT = "rAlbaranVta1b.rpt"
                     End If
                     
                     If Check1(9).Value = 1 Then
                        cadTitulo = "Declaraci�n de Ventas"
                        cadNombreRPT = "rAlbaranVta1c.rpt"
                     End If
                     
                     cadParam = cadParam & "pOrden=""Variedad - Fecha""|"
                     numParam = numParam + 1
                End If
                If optList1(1).Value Then
                     cadNombreRPT = "rAlbaranVta2.rpt"
                     cadParam = cadParam & "pOrden=""Variedad - Cliente - Destino""|"
                     numParam = numParam + 1
                End If
                If optList1(2).Value Then
                    cadNombreRPT = "rAlbaranVta3.rpt"
    '[Monica]19/09/2011: Cambiamos el informe de valorado por precio provisional o definitivo
    '                If Check1(6).Value = 1 Then
    '                    cadNombreRPT = "rAlbaranVta3a.rpt"
    '                    ahora es un informe de evolucion de precios
    '                    If Option1(1).Value Then
    '                        cadParam = cadParam & "pPrecDef=1|"
    '                    Else
    '                        cadParam = cadParam & "pPrecDef=0|"
    '                    End If
    '                    numParam = numParam + 1
    '                End If
                    cadParam = cadParam & "pOrden=""Cliente - Destino - Variedad""|"
                    numParam = numParam + 1
                End If
                If optList1(3).Value Then
                    cadNombreRPT = "rAlbaranVta4.rpt"
                    cadParam = cadParam & "pOrden=""Variedad - Confecci�n""|"
                    numParam = numParam + 1
                End If
                If optList1(4).Value Then
                    cadNombreRPT = "rAlbaranVta5.rpt"
                    cadParam = cadParam & "pOrden=""Confecci�n - Variedad""|"
                    numParam = numParam + 1
                End If
                If optList1(5).Value Then
                    cadNombreRPT = "rAlbaranVta6.rpt"
                    cadParam = cadParam & "pOrden=""Variedad - Marca""|"
                    numParam = numParam + 1
                End If
                If optList1(6).Value Then
                    cadNombreRPT = "rAlbaranVta7.rpt"
                    cadParam = cadParam & "pOrden=""Marca - Variedad""|"
                    numParam = numParam + 1
                End If
                If optList1(7).Value Then
                    cadNombreRPT = "rAlbaranVta8.rpt"
                    cadParam = cadParam & "pOrden=""Variedad - Pa�s""|"
                    numParam = numParam + 1
                End If
                If optList1(8).Value Then
                    cadNombreRPT = "rAlbaranVta9.rpt"
                    cadParam = cadParam & "pOrden=""Pa�s - Variedad""|"
                    numParam = numParam + 1
                End If
                If optList1(9).Value Then
                    cadTitulo = "Gastos Confecci�n Variedad"
                    
                    If Check1(4).Value = False Then ' si no agrupamos los gastos de confeccion
                        cadNombreRPT = "rAlbaranVta10.rpt"
                        ConSubInforme = True
                        
                        If NroGastosMayoraCuatro(cadTABLA, cadselect) Then
                            cadNombreRPT = "rAlbaranVta10a.rpt"
                        End If
                    Else
                        cadNombreRPT = "rAlbaranVta10b.rpt"
                        ConSubInforme = True
                    End If
                End If
                 If optList1(10).Value Then
                    cadTitulo = "Salidas por Calibre"
                    cadNombreRPT = "rAlbaranVta11.rpt"
                End If
                If optList1(11).Value Then
                    cadTitulo = "Tipos de Mercado"
                    If Check1(0).Value Then
                        cadNombreRPT = "rAlbaranVta12a.rpt"
                    Else
                        cadNombreRPT = "rAlbaranVta12.rpt"
                    End If
                End If
                If optList1(12).Value Then
                    cadTitulo = "Gastos Variedad Confecci�n"
                    If Check1(4).Value = False Then ' si no agrupamos los gastos por confeccion
                        cadNombreRPT = "rAlbaranVta13.rpt"
                        ConSubInforme = True
                    
                        If NroGastosMayoraCuatro(cadTABLA, cadselect) Then
                            cadNombreRPT = "rAlbaranVta13a.rpt"
                        End If
                    Else
                        cadNombreRPT = "rAlbaranVta13b.rpt"
                        ConSubInforme = True
                    End If
                End If
                If optList1(13).Value Then
                    cadTitulo = "Variedad distinta de Variedad Comercial"
                    cadNombreRPT = "rAlbaranVta14.rpt"
                    ConSubInforme = True
                End If
                
                '[Monica]12/12/2013: en el caso de comisionista
                If optList1(14).Value Then
                    cadTitulo = "Comisionista - Variedad - Fecha"
                    cadNombreRPT = "rAlbaranVta16.rpt"
                    ConSubInforme = True
                    cadParam = cadParam & "pOrden=""Comisionista - Variedad - Fecha""|"
                    numParam = numParam + 1
                End If
                
                '[Monica]17/10/2016: en el caso de contrato
                If optList1(15).Value Then
                    cadTitulo = "Clientes - Fecha"
                    cadNombreRPT = "rAlbaranVta17.rpt"
                    ConSubInforme = True
                    cadParam = cadParam & "pOrden=""Clientes - Fecha""|"
                    numParam = numParam + 1
                End If
                              
                
                If Check1(5).Value Then
                    '[Monica]27/02/2012: si estamos en rdto por calibre hemos de prorratear todos los gastos por el peso neto
                    '      nueva funcion de ProcesarCambiosGastos
                    If ProcesarCambiosGastos Then
                        cadFormula = "{tmpinformes.codusu} =" & vUsu.Codigo
                        cadTitulo = "Rendimiento por Calibre"
                        cadNombreRPT = "rAlbaranVta15.rpt"
                        ConSubInforme = False
                    End If
                End If
                
                '[Monica]05/03/2013: si solo quiere los que no estan cobrados,
                '                    hemos de eliminar los datos de albaranes que esten marcados como cobrados
                If (optList1(0).Value Or optList1(2).Value) And Check1(10).Value = 1 Then
                    If Not EliminarCobrados Then Exit Sub
                    
                    If Not HayRegistros("tmpinfventas", "codusu = " & vUsu.Codigo) Then Exit Sub
                End If
                
                
                LlamarImprimir
            End If
      End If
    End If
    
End Sub

Private Function EliminarCobrados() As Boolean
Dim Sql As String
    
    On Error GoTo eEliminarCobrados
        
    EliminarCobrados = False
        
    Sql = "delete from tmpinfventas where codusu = " & vUsu.Codigo
    Sql = Sql & " and cobrado = 1 "

    conn.Execute Sql
        
    EliminarCobrados = True
    Exit Function
    
eEliminarCobrados:
    MuestraError Err.Number, "Seleccionar los No cobrados", Err.Description
End Function


Private Function ProcesarCambiosGastos()
Dim Sql As String
Dim Sql3 As String
Dim Sql2 As String
Dim Sql4 As String

Dim Rs As ADODB.Recordset
Dim RS3 As ADODB.Recordset

Dim Gastos As Currency
Dim vGastos As Currency
Dim GastosAc As Currency
Dim TPesoNeto As Long
Dim ImporteFac As Currency
Dim UltimaLinea1 As Long
Dim Diferencia As Currency
Dim HayReg As Long

    On Error GoTo eProcesarCambiosGastos
    

    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    Sql = "select * from tmpinfventas where codusu = " & vUsu.Codigo
    Sql = Sql & " order by numalbar, numlinea "
    
    
    Label4(27).visible = True
    Pb1.visible = True
    Label4(27).Caption = "Cargando tabla temporal: prorrateo de gastos por kilos"
    DoEvents
        
    HayReg = TotalRegistrosConsulta(Sql)
    
    Pb1.Max = HayReg
    Pb1.Value = 0
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                                            'albaran,numlinea,numline1,gasto por linea de calibre, importe de venta
    Sql2 = "insert into tmpinformes (codusu, codigo1, campo1, campo2, importe1, importe2) values "
    
    cadselect = ""
    
    While Not Rs.EOF
        IncrementarProgresNew Pb1, 1
    
    
        TPesoNeto = DBLet(Rs!Pesoneto, "N")
        Gastos = DBLet(Rs!Gastos, "N") ' DBLet(Rs!GastosPortes, "N") + DBLet(Rs!GastosEnvases, "N") + DBLet(Rs!gastoscomisiones, "N")
    
        Sql3 = "select * from albaran_calibre where numalbar = " & DBSet(Rs!NumAlbar, "N") & " and numlinea = " & DBSet(Rs!NumLinea, "N")
        Sql3 = Sql3 & " order by numline1 "
        
        Set RS3 = New ADODB.Recordset
        RS3.Open Sql3, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        GastosAc = 0
    
        While Not RS3.EOF
            vGastos = 0
            If TPesoNeto <> 0 Then
                vGastos = Round2(RS3!Pesoneto * Gastos / TPesoNeto, 2)
            End If
            GastosAc = GastosAc + vGastos
            UltimaLinea1 = DBLet(RS3!numline1, "N")
        
            Sql4 = "Select sum(impornet) from facturas_calibre where numalbar = " & DBSet(Rs!NumAlbar, "N")
            Sql4 = Sql4 & " and numlinealbar = " & DBSet(Rs!NumLinea, "N")
            Sql4 = Sql4 & " and numline1albar = " & DBSet(RS3!numline1, "N")
            
            ImporteFac = DevuelveValor(Sql4)
        
            
            '[Monica]31/03/2017: para el caso de Montifrut las facturas pueden estar en la otra bbdd con lo que tenemos que
            '                    jugar con el precio definitivo
            If vParamAplic.Cooperativa = 12 Then
                Sql4 = "Select count(*) from facturas_calibre where numalbar = " & DBSet(Rs!NumAlbar, "N")
                Sql4 = Sql4 & " and numlinealbar = " & DBSet(Rs!NumLinea, "N")
                Sql4 = Sql4 & " and numline1albar = " & DBSet(RS3!numline1, "N")
                If TotalRegistros(Sql4) = 0 Then
                    ImporteFac = Round2(DBLet(RS3!Pesoneto, "N") * DBLet(RS3!preciopro, "N"), 2)
                End If
            End If
        
        
            cadselect = "(" & vUsu.Codigo & "," & DBSet(Rs!NumAlbar, "N") & "," & DBSet(Rs!NumLinea, "N") & ","
            cadselect = cadselect & DBSet(RS3!numline1, "N") & "," & DBSet(vGastos, "N") & ","
            cadselect = cadselect & DBSet(ImporteFac, "N") & ")"
        
            conn.Execute Sql2 & cadselect
        
            RS3.MoveNext
        Wend
        
        Set RS3 = Nothing
        
        If GastosAc <> Gastos Then
            Diferencia = Gastos - GastosAc
            Sql3 = "update tmpinformes set importe1 = importe1 + " & DBSet(Diferencia, "N")
            Sql3 = Sql3 & " where codusu = " & vUsu.Codigo
            Sql3 = Sql3 & " and codigo1 = " & DBSet(Rs!NumAlbar, "N")
            Sql3 = Sql3 & " and campo1 = " & DBSet(Rs!NumLinea, "N")
            Sql3 = Sql3 & " and campo2 = " & DBSet(UltimaLinea1, "N")
        
            conn.Execute Sql3
        End If
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    Label4(27).visible = False
    Pb1.visible = False
    DoEvents
    
    ProcesarCambiosGastos = True
    Exit Function

eProcesarCambiosGastos:
    MuestraError Err.Number, "Procesar Cambios Gastos", Err.Description
    ProcesarCambiosGastos = False
    Label4(27).visible = False
    Pb1.visible = False
    DoEvents
End Function

Private Function ProcesarCambios(cadTABLA, cadwhere As String) As Boolean
Dim Sql As String
Dim SQL1 As String
Dim Sql2 As String
Dim i As Integer
Dim HayReg As Long
Dim b As Boolean
Dim Rs As ADODB.Recordset
Dim Rsx As ADODB.Recordset
Dim TotalGastos As Currency
Dim TotalGastosReales As Currency
Dim PesoCaja As Currency
Dim PesoReal As Currency
Dim ImpVenta As Currency
Dim Facturado As Byte
Dim Cobrado As Byte
Dim cadTabla2 As String

Dim Coste1 As Integer
Dim Coste2 As Integer
Dim Coste3 As Integer
Dim Coste4 As Integer
Dim Coste5 As Integer

Dim Gasto1 As Currency
Dim Gasto2 As Currency
Dim Gasto3 As Currency
Dim Gasto4 As Currency
Dim Gasto5 As Currency
Dim Costes As Integer
Dim GastosEnvases As Currency
Dim GastosPortes As Currency
Dim GastosComision As Currency

Dim Sql3 As String
Dim PreProv As Currency

On Error GoTo eProcesarCambios

    HayReg = 0
    
    ProcesarCambios = False
    
    conn.Execute "delete from tmpinfventas where codusu = " & DBSet(vUsu.Codigo, "N")
        
    If cadwhere <> "" Then
        cadwhere = QuitarCaracterACadena(cadwhere, "{")
        cadwhere = QuitarCaracterACadena(cadwhere, "}")
        cadwhere = QuitarCaracterACadena(cadwhere, "_1")
    End If
        
    SQL1 = "select albaran.fechaalb, albaran.numalbar, albaran_variedad.numlinea, "
    SQL1 = SQL1 & "albaran_variedad.numcajas, albaran_variedad.pesoneto, albaran_variedad.preciopro, "
    SQL1 = SQL1 & "sum(facturas_variedad.impornet), forfaits.kiloscaj, albaran_variedad.preciodef from " & cadTABLA
    SQL1 = SQL1 & " where (1 = 1) "
    If cadwhere <> "" Then SQL1 = SQL1 & " and " & cadwhere
    SQL1 = SQL1 & " group by 1, 2, 3, 4, 5, 6, 8, 9"
    SQL1 = SQL1 & " order by 1, 2, 3, 4, 5, 6, 8, 9"
        
    Set Rs = New ADODB.Recordset
    Rs.Open SQL1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Label4(27).Caption = "Cargando tabla temporal"
    Label4(27).visible = True
    Pb1.visible = True
    DoEvents
        
    '[Monica]25/04/2018: cambiado fuera del while
    If Me.optList1(9).Value Or Me.optList1(12).Value Or Check1(8).Value Then
        If Me.Check1(3).Value = 0 Then '++monica:240608 a�adida la condicion caso de que los gastos no sean los reales
            cadTabla2 = "(" & cadTABLA & ") inner join albaran_costes on albaran_variedad.numalbar = albaran_costes.numalbar "
            cadTabla2 = cadTabla2 & " and albaran_variedad.numlinea = albaran_costes.numlinea "
            
            Sql2 = "select count(distinct albaran_costes.codcoste) from " & cadTabla2
            If cadwhere <> "" Then Sql2 = Sql2 & " where " & cadwhere
            
            Costes = DevuelveValor(Sql2)
            If CCur(Costes) > 5 Then
                MsgBox "El numero de costes distintos es superior a cinco y no cabe en el listado", vbExclamation
                ProcesarCambios = False
                Exit Function
            End If
        Else
            cadTabla2 = "(" & cadTABLA & ") inner join albaran_costreal on albaran_variedad.numalbar = albaran_costreal.numalbar "
            cadTabla2 = cadTabla2 & " and albaran_variedad.numlinea = albaran_costreal.numlinea "
            
            Sql2 = "select count(distinct albaran_costreal.codcoste) from " & cadTabla2
            Sql2 = Sql2 & cadwhere
            
            Costes = DevuelveValor(Sql2)
            If CCur(Costes) > 5 Then
                MsgBox "El numero de costes distintos es superior a cinco y no cabe en el listado", vbExclamation
                ProcesarCambios = False
                Exit Function
            End If
        End If
    End If
        
        
        
    HayReg = TotalRegistrosConsulta(SQL1)
    
    Pb1.Max = HayReg
    Pb1.Value = 0
'[Monica]25/02/2013: antes
'    Coste1 = -1
'    Coste2 = -1
'    Coste3 = -1
'    Coste4 = -1
'    Coste5 = -1
'ahora
    Coste1 = 1
    Coste2 = 2
    Coste3 = 3
    Coste4 = 4
    Coste5 = 5
    
    Sql = ""
    
    While Not Rs.EOF
        IncrementarProgresNew Pb1, 1
    
        Sql2 = "select sum(if(impcoste is null,0,impcoste)) from albaran_costes where numalbar = "
        Sql2 = Sql2 & DBLet(Rs.Fields(1).Value, "N") & " and numlinea = "
        Sql2 = Sql2 & DBLet(Rs.Fields(2).Value, "N")
        
        If Me.Check1(3).Value = 1 Then '++monica:240608 en caso de que los costes sean los reales
            Sql2 = Sql2 & " and (albaran_costes.tipogasto = 1 or albaran_costes.tipogasto = 4) "
        End If
        
        If Not (optList1(9).Value Or optList1(12).Value) Then
            If Me.Check1(2).Value = 0 Then
                Sql2 = Sql2 & " and (albaran_costes.tipogasto < 2 or albaran_costes.tipogasto = 4) "
            End If
        End If

        TotalGastos = DevuelveValor(Sql2)
        
        
        '[Monica]27/10/2011: Si no hay gastos de portes o comisiones se prorratean
'        If Not (optList1(9).Value Or optList1(12).Value) Then '[Monica]23/12/2011: fallaba esta condicion
            If Me.Check1(2).Value = 1 Then
                Sql2 = "select sum(if(impcoste is null,0,impcoste)) from albaran_costes where numalbar = "
                Sql2 = Sql2 & DBLet(Rs.Fields(1).Value, "N") & " and numlinea = "
                Sql2 = Sql2 & DBLet(Rs.Fields(2).Value, "N")
                Sql2 = Sql2 & " and tipogasto = 2 "
                
                GastosPortes = DevuelveValor(Sql2)
                
                '[Monica]27/10/2011: Si no hay gastos de portes pq no hay facturas de transporte hemos de valorarlo
                '                    con el porte provisional de cabecera de albaran
                '                    Le pasamos a la funcion: (albaran, linea, tipo)
                If GastosPortes = 0 Then
                    GastosPortes = ProrrateoPortesComisProvisional(DBLet(Rs.Fields(1).Value, "N"), DBLet(Rs.Fields(2).Value, "N"), 0)
                    TotalGastos = TotalGastos + GastosPortes
                End If
                
                
                Sql2 = "select sum(if(impcoste is null,0,impcoste)) from albaran_costes where albaran_costes.numalbar = " & DBLet(Rs.Fields(1).Value, "N")
                Sql2 = Sql2 & " and albaran_costes.numlinea = " & DBLet(Rs.Fields(2).Value, "N")
                Sql2 = Sql2 & " and albaran_costes.tipogasto = 3 "
                GastosComision = DevuelveValor(Sql2)
                
                '[Monica]27/10/2011: Si no hay gastos de portes pq no hay facturas de transporte hemos de valorarlo
                '                    con el porte provisional de cabecera de albaran
                '                    Le pasamos a la funcion: (albaran, linea, tipo)
                If GastosComision = 0 Then
                    GastosComision = ProrrateoPortesComisProvisional(DBLet(Rs.Fields(1).Value, "N"), DBLet(Rs.Fields(2).Value, "N"), 1)
                    TotalGastos = TotalGastos + GastosComision
                End If
                
                '[Monica]18/11/2011: a�adimos la condicion de que no salga a execl
                If Check1(4).Value = 0 And Check1(8).Value = 0 Then
                    GastosPortes = GastosPortes + GastosComision
                End If
            End If
'        End If
        ' fin de a�adido 27/10/2011
        
        
        '++monica:240608 a�adido para el caso de los costes reales
        If Me.Check1(3).Value = 1 Then
            Sql2 = "select sum(if(impcoste is null,0,impcoste)) from albaran_costreal where numalbar = "
            Sql2 = Sql2 & DBLet(Rs.Fields(1).Value, "N") & " and numlinea = "
            Sql2 = Sql2 & DBLet(Rs.Fields(2).Value, "N")
            
            TotalGastosReales = DevuelveValor(Sql2)
            
            TotalGastos = TotalGastos + TotalGastosReales
        End If
        
'        Sql2 = "select forfaits.kiloscaj from forfaits, albaran_variedad where "
'        Sql2 = Sql2 & " albaran_variedad.numalbar = " & DBSet(Rs.Fields(1).Value, "N")
'        Sql2 = Sql2 & " and albaran_variedad.numlinea = " & DBSet(Rs.Fields(2).Value, "N")
'        Sql2 = Sql2 & " and albaran_variedad.codforfait = forfaits.codforfait "
'
'        PesoCaja = DevuelveValor(Sql2)
        PesoCaja = DBLet(Rs!kiloscaj, "N")
        PesoReal = Round2(PesoCaja * DBLet(Rs.Fields(3).Value, "N"), 2)
        
        
'[Monica]19/09/2011: sustituido todo esto por el tema de precio provisional o precio definitivo o precio facturado
'        ImpVenta = 0
'        PreProv = 0
'        If Check1(6).Value = 1 Then
'        '[Monica]10/12/2010: jugamos con el precio provisional en los listados Variedad-Fecha y
'        ' Cliente-Destino-Variedad. Es un informe Provisional
'            If Option1(0).Value Then ' Precio Provisional
'                PreProv = DBLet(RS.Fields(5).Value, "N")
'                ImpVenta = Round2(DBLet(RS.Fields(4).Value, "N") * DBLet(RS.Fields(5).Value, "N"), 2)
'                Facturado = 0
'                Cobrado = 0
'            Else ' Precio Definitivo
'                PreProv = DBLet(RS.Fields(8).Value, "N")
'                ImpVenta = Round2(DBLet(RS.Fields(4).Value, "N") * DBLet(RS.Fields(8).Value, "N"), 2)
'                Facturado = 0
'                Cobrado = 0
'            End If
'        Else
'            '[Monica]10/12/2010: lo que hacia antes
'            If Not IsNull(RS.Fields(6).Value) Then
'                ImpVenta = RS.Fields(6).Value
'                Facturado = 1
'                ' solo en este caso miro si esta o no cobrada en tesoreria
'                '[Monica]16/04/2010:antes FacturaCobradaTesoreria
'                'Cobrado = FacturaCobradaTesoreria(DBLet(RS.Fields(1).Value, "N"), DBLet(RS.Fields(2).Value, "N"))
'                Cobrado = AlbaranCobradoTesoreria(DBLet(RS.Fields(1).Value, "N"), DBLet(RS.Fields(2).Value, "N"))
'            Else
'                ImpVenta = Round2(DBLet(RS.Fields(4).Value, "N") * DBLet(RS.Fields(5).Value, "N"), 2)
'                Facturado = 0
'                Cobrado = 0
'            End If
'        End If
'
' sustituido lo anterior por lo siquiente:
        PreProv = 0
        If Not IsNull(Rs.Fields(6).Value) Then
            ImpVenta = Rs.Fields(6).Value
            Facturado = 2 ' facturado
            ' solo en este caso miro si esta o no cobrada en tesoreria
            '[Monica]16/04/2010:antes FacturaCobradaTesoreria
            'Cobrado = FacturaCobradaTesoreria(DBLet(RS.Fields(1).Value, "N"), DBLet(RS.Fields(2).Value, "N"))
            Cobrado = AlbaranCobradoTesoreria(DBLet(Rs.Fields(1).Value, "N"), DBLet(Rs.Fields(2).Value, "N"))
            PreProv = 0
            If DBLet(Rs.Fields(4).Value, "N") <> 0 Then
                PreProv = Round2(ImpVenta / DBLet(Rs.Fields(4).Value, "N"), 4)
            End If
        Else
            If DBLet(Rs.Fields(8).Value, "N") <> 0 Then
                ImpVenta = Round2(DBLet(Rs.Fields(4).Value, "N") * DBLet(Rs.Fields(8).Value, "N"), 2)
                Facturado = 1 ' definitivo
                Cobrado = 0
                PreProv = DBLet(Rs.Fields(8).Value, "N")
            Else
                ImpVenta = Round2(DBLet(Rs.Fields(4).Value, "N") * DBLet(Rs.Fields(5).Value, "N"), 2)
                Facturado = 0 ' provisional
                Cobrado = 0
                PreProv = DBLet(Rs.Fields(5).Value, "N")
            End If
        End If
        
        Gasto1 = 0
        Gasto2 = 0
        Gasto3 = 0
        Gasto4 = 0
        Gasto5 = 0
        
        '[Monica]16/11/2011: A�adida la condicion de check1(8).value ( salida a hoja excel )
        'calculo para informe Gastos de Confeccion : rAlbaranVta10
        If Me.optList1(9).Value Or Me.optList1(12).Value Or Check1(8).Value Then
            If Me.Check1(3).Value = 0 Then '++monica:240608 a�adida la condicion caso de que los gastos no sean los reales
'[Monica]25/04/2018: fuera del while
'                cadTabla2 = "(" & cadTABLA & ") inner join albaran_costes on albaran_variedad.numalbar = albaran_costes.numalbar "
'                cadTabla2 = cadTabla2 & " and albaran_variedad.numlinea = albaran_costes.numlinea "
'
'                Sql2 = "select count(distinct albaran_costes.codcoste) from " & cadTabla2
'                If cadWHERE <> "" Then Sql2 = Sql2 & " where " & cadWHERE
'
'                Costes = DevuelveValor(Sql2)
'                If CCur(Costes) > 5 Then
'                    MsgBox "El numero de costes distintos es superior a cinco y no cabe en el listado", vbExclamation
'                    ProcesarCambios = False
'                    Exit Function
'                End If
                
                Sql2 = "select codcoste, impcoste from albaran_costes where albaran_costes.numalbar = " & DBSet(Rs.Fields(1).Value, "N")
                Sql2 = Sql2 & " and albaran_costes.numlinea = " & DBSet(Rs.Fields(2).Value, "N")
                Sql2 = Sql2 & " and albaran_costes.tipogasto = 0 "
                
                Set Rsx = New ADODB.Recordset
                Rsx.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                While Not Rsx.EOF
                    '[Monica]25/02/2013: Pueden poner en al confeccion unicamente el coste coste que tiene
                    'If Coste1 = -1 Or Coste1 = DBLet(Rsx.Fields(0).Value, "N") Then
                    If Coste1 = DBLet(Rsx.Fields(0).Value, "N") Then
                        Coste1 = DBLet(Rsx.Fields(0).Value, "N")
                        Gasto1 = DBLet(Rsx.Fields(1).Value, "N")
                    Else
                        '[Monica]25/02/2013: Pueden poner en al confeccion unicamente el coste coste que tiene
                        'If Coste2 = -1 Or Coste2 = DBLet(Rsx.Fields(0).Value, "N") Then
                        If Coste2 = DBLet(Rsx.Fields(0).Value, "N") Then
                            Coste2 = DBLet(Rsx.Fields(0).Value, "N")
                            Gasto2 = DBLet(Rsx.Fields(1).Value, "N")
                        Else
                            '[Monica]25/02/2013: Pueden poner en al confeccion unicamente el coste coste que tiene
                            'If Coste3 = -1 Or Coste3 = DBLet(Rsx.Fields(0).Value, "N") Then
                            If Coste3 = DBLet(Rsx.Fields(0).Value, "N") Then
                                Coste3 = DBLet(Rsx.Fields(0).Value, "N")
                                Gasto3 = DBLet(Rsx.Fields(1).Value, "N")
                            Else
                                '[Monica]25/02/2013: Pueden poner en al confeccion unicamente el coste coste que tiene
                                'If Coste4 = -1 Or Coste4 = DBLet(Rsx.Fields(0).Value, "N") Then
                                If Coste4 = DBLet(Rsx.Fields(0).Value, "N") Then
                                    Coste4 = DBLet(Rsx.Fields(0).Value, "N")
                                    Gasto4 = DBLet(Rsx.Fields(1).Value, "N")
                                Else
                                    '[Monica]25/02/2013: Pueden poner en al confeccion unicamente el coste coste que tiene
                                    'If Coste5 = -1 Or Coste5 = DBLet(Rsx.Fields(0).Value, "N") Then
                                    If Coste5 = DBLet(Rsx.Fields(0).Value, "N") Then
                                        Coste5 = DBLet(Rsx.Fields(0).Value, "N")
                                        Gasto5 = DBLet(Rsx.Fields(1).Value, "N")
                                    End If
                                End If
                            End If
                        End If
                   End If
                   Rsx.MoveNext
                Wend
                
                Sql2 = "select sum(if(impcoste is null,0,impcoste)) from albaran_costes where albaran_costes.numalbar = " & DBLet(Rs.Fields(1).Value, "N")
                Sql2 = Sql2 & " and albaran_costes.numlinea = " & DBLet(Rs.Fields(2).Value, "N")
                Sql2 = Sql2 & " and albaran_costes.tipogasto = 1 "
                GastosEnvases = DevuelveValor(Sql2)
                
                '[Monica] 15/06/2010 a�adido costes paletizacion
                Sql2 = "select sum(if(impcoste is null,0,impcoste)) from albaran_costes where albaran_costes.numalbar = " & DBLet(Rs.Fields(1).Value, "N")
                Sql2 = Sql2 & " and albaran_costes.numlinea = " & DBLet(Rs.Fields(2).Value, "N")
                Sql2 = Sql2 & " and albaran_costes.tipogasto = 4 "
                GastosEnvases = GastosEnvases + DevuelveValor(Sql2)

                
                Sql2 = "select sum(if(impcoste is null,0,impcoste)) from albaran_costes where albaran_costes.numalbar = " & DBLet(Rs.Fields(1).Value, "N")
                Sql2 = Sql2 & " and albaran_costes.numlinea = " & DBLet(Rs.Fields(2).Value, "N")
                Sql2 = Sql2 & " and albaran_costes.tipogasto = 2 "
                GastosPortes = DevuelveValor(Sql2)
                
                '[Monica]27/10/2011: Si no hay gastos de portes pq no hay facturas de transporte hemos de valorarlo
                '                    con el porte provisional de cabecera de albaran
                '                    Le pasamos a la funcion: (albaran, linea, tipo)
                If GastosPortes = 0 Then
                    GastosPortes = ProrrateoPortesComisProvisional(DBLet(Rs.Fields(1).Value, "N"), DBLet(Rs.Fields(2).Value, "N"), 0)
                End If
                
                
                Sql2 = "select sum(if(impcoste is null,0,impcoste)) from albaran_costes where albaran_costes.numalbar = " & DBLet(Rs.Fields(1).Value, "N")
                Sql2 = Sql2 & " and albaran_costes.numlinea = " & DBLet(Rs.Fields(2).Value, "N")
                Sql2 = Sql2 & " and albaran_costes.tipogasto = 3 "
                GastosComision = DevuelveValor(Sql2)
                
                '[Monica]27/10/2011: Si no hay gastos de portes pq no hay facturas de transporte hemos de valorarlo
                '                    con el porte provisional de cabecera de albaran
                '                    Le pasamos a la funcion: (albaran, linea, tipo)
                If GastosComision = 0 Then
                    GastosComision = ProrrateoPortesComisProvisional(DBLet(Rs.Fields(1).Value, "N"), DBLet(Rs.Fields(2).Value, "N"), 1)
                End If
                
                '[Monica]18/11/2011: a�adimos la condicion de que no salga a excel
                If Check1(4).Value = 0 And Check1(8).Value = 0 Then
                    GastosPortes = GastosPortes + GastosComision
                End If
                
            Else '++monica:240608 caso de que los gastos sean reales a�adido todo el else
            
'[Monica]25/04/2018: fuera del while
'                cadTabla2 = "(" & cadTABLA & ") inner join albaran_costreal on albaran_variedad.numalbar = albaran_costreal.numalbar "
'                cadTabla2 = cadTabla2 & " and albaran_variedad.numlinea = albaran_costreal.numlinea "
'                Sql2 = "select count(distinct albaran_costreal.codcoste) from " & cadTabla2
'                Sql2 = Sql2 & cadWHERE
'
'                Costes = DevuelveValor(Sql2)
'                If CCur(Costes) > 5 Then
'                    MsgBox "El numero de costes distintos es superior a cinco y no cabe en el listado", vbExclamation
'                    ProcesarCambios = False
'                    Exit Function
'                End If
                
                Sql2 = "select codcoste, impcoste from albaran_costreal where albaran_costreal.numalbar = " & DBLet(Rs.Fields(1).Value, "N")
                Sql2 = Sql2 & " and albaran_costreal.numlinea = " & DBLet(Rs.Fields(2).Value, "N")
                Sql2 = Sql2 & " and albaran_costreal.tipogasto = 0 "
                
                Set Rsx = New ADODB.Recordset
                Rsx.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                While Not Rsx.EOF
                    '[Monica]25/02/2013: Pueden poner en al confeccion unicamente el coste coste que tiene
                    'If Coste1 = -1 Or Coste1 = DBLet(Rsx.Fields(0).Value, "N") Then
                    If Coste1 = DBLet(Rsx.Fields(0).Value, "N") Then
                        Coste1 = DBLet(Rsx.Fields(0).Value, "N")
                        Gasto1 = DBLet(Rsx.Fields(1).Value, "N")
                    Else
                        '[Monica]25/02/2013: Pueden poner en al confeccion unicamente el coste coste que tiene
                        'If Coste2 = -1 Or Coste2 = DBLet(Rsx.Fields(0).Value, "N") Then
                        If Coste2 = DBLet(Rsx.Fields(0).Value, "N") Then
                            Coste2 = DBLet(Rsx.Fields(0).Value, "N")
                            Gasto2 = DBLet(Rsx.Fields(1).Value, "N")
                        Else
                            '[Monica]25/02/2013: Pueden poner en al confeccion unicamente el coste coste que tiene
                            'If Coste3 = -1 Or Coste3 = DBLet(Rsx.Fields(0).Value, "N") Then
                            If Coste3 = DBLet(Rsx.Fields(0).Value, "N") Then
                                Coste3 = DBLet(Rsx.Fields(0).Value, "N")
                                Gasto3 = DBLet(Rsx.Fields(1).Value, "N")
                            Else
                                '[Monica]25/02/2013: Pueden poner en al confeccion unicamente el coste coste que tiene
                                'If Coste4 = -1 Or Coste4 = DBLet(Rsx.Fields(0).Value, "N") Then
                                If Coste4 = DBLet(Rsx.Fields(0).Value, "N") Then
                                    Coste4 = DBLet(Rsx.Fields(0).Value, "N")
                                    Gasto4 = DBLet(Rsx.Fields(1).Value, "N")
                                Else
                                    '[Monica]25/02/2013: Pueden poner en al confeccion unicamente el coste coste que tiene
                                    'If Coste5 = -1 Or Coste5 = DBLet(Rsx.Fields(0).Value, "N") Then
                                    If Coste5 = DBLet(Rsx.Fields(0).Value, "N") Then
                                        Coste5 = DBLet(Rsx.Fields(0).Value, "N")
                                        Gasto5 = DBLet(Rsx.Fields(1).Value, "N")
                                    End If
                                End If
                            End If
                        End If
                   End If
                   Rsx.MoveNext
                Wend
                
                Sql2 = "select sum(if(impcoste is null,0,impcoste)) from albaran_costes where albaran_costes.numalbar = " & DBLet(Rs.Fields(1).Value, "N")
                Sql2 = Sql2 & " and albaran_costes.numlinea = " & DBLet(Rs.Fields(2).Value, "N")
                Sql2 = Sql2 & " and albaran_costes.tipogasto = 1 "
                GastosEnvases = DevuelveValor(Sql2)
                
                '[Monica] 15/06/2010 a�adido costes paletizacion
                Sql2 = "select sum(if(impcoste is null,0,impcoste)) from albaran_costes where albaran_costes.numalbar = " & DBLet(Rs.Fields(1).Value, "N")
                Sql2 = Sql2 & " and albaran_costes.numlinea = " & DBLet(Rs.Fields(2).Value, "N")
                Sql2 = Sql2 & " and albaran_costes.tipogasto = 4 "
                GastosEnvases = GastosEnvases + DevuelveValor(Sql2)
                
                
                Sql2 = "select sum(if(impcoste is null,0,impcoste)) from albaran_costes where albaran_costes.numalbar = " & DBLet(Rs.Fields(1).Value, "N")
                Sql2 = Sql2 & " and albaran_costes.numlinea = " & DBLet(Rs.Fields(2).Value, "N")
                Sql2 = Sql2 & " and albaran_costes.tipogasto = 2 "
                GastosPortes = DevuelveValor(Sql2)
            
                '[Monica]27/10/2011: Si no hay gastos de portes pq no hay facturas de transporte hemos de valorarlo
                '                    con el porte provisional de cabecera de albaran
                '                    Le pasamos a la funcion: (albaran, linea, tipo)
                If GastosPortes = 0 Then
                    GastosPortes = ProrrateoPortesComisProvisional(DBLet(Rs.Fields(1).Value, "N"), DBLet(Rs.Fields(2).Value, "N"), 0)
                End If
            
            
                Sql2 = "select sum(if(impcoste is null,0,impcoste)) from albaran_costes where albaran_costes.numalbar = " & DBLet(Rs.Fields(1).Value, "N")
                Sql2 = Sql2 & " and albaran_costes.numlinea = " & DBLet(Rs.Fields(2).Value, "N")
                Sql2 = Sql2 & " and albaran_costes.tipogasto = 3 "
                GastosComision = DevuelveValor(Sql2)
                
                '[Monica]27/10/2011: Si no hay gastos de portes pq no hay facturas de transporte hemos de valorarlo
                '                    con el porte provisional de cabecera de albaran
                '                    Le pasamos a la funcion: (albaran, linea, tipo)
                If GastosComision = 0 Then
                    GastosComision = ProrrateoPortesComisProvisional(DBLet(Rs.Fields(1).Value, "N"), DBLet(Rs.Fields(2).Value, "N"), 1)
                End If
                
                '[Monica]18/11/2011: a�adimos la condicion de que no salga a excel
                If Check1(4).Value = 0 And Check1(8).Value = 0 Then
                    GastosPortes = GastosPortes + GastosComision
                End If
            End If
        End If
        
        Sql = Sql & "(" & DBSet(vUsu.Codigo, "N") & ","
        Sql = Sql & DBSet(Rs.Fields(0).Value, "F") & "," & DBSet(Rs.Fields(1).Value, "N") & "," & DBSet(Rs.Fields(2).Value, "N") & ","
        Sql = Sql & DBSet(Rs.Fields(3).Value, "N") & "," 'numero de cajas
        Sql = Sql & DBSet(PesoReal, "N") & "," & DBSet(Rs.Fields(4).Value, "N") & "," 'peso neto
        Sql = Sql & DBSet(TotalGastos, "N") & "," & DBSet(ImpVenta, "N") & "," ' importe de venta
        Sql = Sql & DBSet(Facturado, "N") & ","  'facturado o no, pasa a ser : 0=provisional 1=definitivo 2=facturado
        Sql = Sql & DBSet(Cobrado, "N") & "," 'cobrado o no
        Sql = Sql & DBSet(Coste1, "N") & "," & DBSet(Gasto1, "N") & "," 'coste1 gasto1
        Sql = Sql & DBSet(Coste2, "N") & "," & DBSet(Gasto2, "N") & "," 'coste2 gasto2
        Sql = Sql & DBSet(Coste3, "N") & "," & DBSet(Gasto3, "N") & "," 'coste3 gasto3
        Sql = Sql & DBSet(Coste4, "N") & "," & DBSet(Gasto4, "N") & "," 'coste4 gasto4
        Sql = Sql & DBSet(Coste5, "N") & "," & DBSet(Gasto5, "N") & "," 'coste5 gasto5
        Sql = Sql & DBSet(GastosPortes, "N") & "," ' gastos portes
        Sql = Sql & DBSet(GastosComision, "N") & "," ' gastos comisiones
        Sql = Sql & DBSet(GastosEnvases, "N") & "," ' gastos envases
        Sql = Sql & DBSet(PreProv, "N") & ")," ' precio provisional para el informe provisional
        
'        Conn.Execute Sql
      
        Rs.MoveNext
    Wend
    
    If Sql <> "" Then
        ' quitamos la ultima coma
        Sql = Mid(Sql, 1, Len(Sql) - 1)
    
        Sql3 = "insert into tmpinfventas (codusu, fecalbar, numalbar, numlinea, numcajas, pesoreal, pesoneto, gastos, impventa, facturado, cobrado, "
        Sql3 = Sql3 & " codigo1, gastos1, codigo2, gastos2, codigo3, gastos3, codigo4, gastos4, codigo5, gastos5, gastosportes, gastoscomisiones, gastosenvases, precioprovisional) values "
        Sql3 = Sql3 & Sql
        
        conn.Execute Sql3
    End If
    
    Sql3 = "update tmpinfventas a, tmpinfventas b set a.codigo5 = b.codigo5 "
    Sql3 = Sql3 & " where b.codusu = " & vUsu.Codigo
    Sql3 = Sql3 & " and a.codusu = " & vUsu.Codigo
    Sql3 = Sql3 & " and  b.codigo5 > 0 and a.codigo5 = -1"
    
    conn.Execute Sql3
    
    ProcesarCambios = True

    Label4(27).visible = False
    Pb1.visible = False
    
eProcesarCambios:
    If Err.Number <> 0 Then
        ProcesarCambios = False
    End If
End Function


Private Function ProrrateoPortesComisProvisional(Albaran As String, Linea As String, Tipo As Byte) As Currency
'Tipo = 0 portes
'     = 1 comisiones
Dim CADENA As String
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim KilosTot As Long
Dim KilosNet As Long
Dim PortesTot As Currency

    If Tipo = 0 Then
        Sql = "select portespre from albaran where numalbar = " & DBSet(Albaran, "N")
    Else
        Sql = "select comisionespre from albaran where numalbar = " & DBSet(Albaran, "N")
    End If
    PortesTot = DevuelveValor(Sql)
    
    Sql = "select sum(pesoneto) from albaran_variedad where numalbar = " & DBSet(Albaran, "N")
    KilosTot = DevuelveValor(Sql)
    
    Sql = "select pesoneto from albaran_variedad where numalbar = " & DBSet(Albaran, "N") & " and numlinea = " & DBSet(Linea, "N")
    KilosNet = DevuelveValor(Sql)
    
    ProrrateoPortesComisProvisional = 0
    If KilosTot <> 0 Then
        ProrrateoPortesComisProvisional = Round2(PortesTot * KilosNet / KilosTot, 2)
    End If
    

End Function


Private Function ProcesarCambiosCalibres(cadTABLA, cadwhere As String) As Boolean
Dim Sql As String
Dim SQL1 As String
Dim Sql2 As String
Dim i As Integer
Dim HayReg As Long
Dim b As Boolean
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim PesoCaja As Currency
Dim PesoReal As Currency
Dim ImpVenta As Currency
Dim Facturado As Byte
Dim cadTabla2 As String
Dim Incluido As Byte ' me indica si he podido incluir el calibre

Dim VariedadAnt As Integer

Dim Calibre, Neto

On Error GoTo eProcesarCambiosCalibres

    HayReg = 0
    
    conn.Execute "delete from tmpinfventas where codusu = " & DBSet(vUsu.Codigo, "N")
        
    If cadwhere <> "" Then
        cadwhere = QuitarCaracterACadena(cadwhere, "{")
        cadwhere = QuitarCaracterACadena(cadwhere, "}")
        cadwhere = QuitarCaracterACadena(cadwhere, "_1")
    End If
        
    SQL1 = "select albaran_variedad.codvarie, albaran.fechaalb, albaran.numalbar, albaran_variedad.numlinea, "
    SQL1 = SQL1 & "albaran_variedad.numcajas, albaran_variedad.pesoneto from " & cadTABLA
    SQL1 = SQL1 & cadwhere
    SQL1 = SQL1 & " order by 1, 2, 3, 4, 5, 6"
        
    Set Rs = New ADODB.Recordset
    Rs.Open SQL1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Label4(27).visible = True
    Pb1.visible = True
        
    HayReg = TotalRegistrosConsulta(SQL1)
    Pb1.Max = HayReg
            
    If Not Rs.EOF Then
        Calibre = Array(-1, -1, -1, -1, -1, -1, -1, -1, -1)
        VariedadAnt = DBLet(Rs.Fields(0).Value, "N")
    End If
            
    While Not Rs.EOF
        If VariedadAnt <> DBLet(Rs.Fields(0).Value, "N") Then
            Calibre = Array(-1, -1, -1, -1, -1, -1, -1, -1, -1)
            VariedadAnt = DBLet(Rs.Fields(0).Value, "N")
        End If
        IncrementarProgresNew Pb1, 1
    
        Sql2 = "select codcalib, sum(pesoneto) from albaran_calibre where numalbar = "
        Sql2 = Sql2 & DBSet(Rs.Fields(2).Value, "N") & " and numlinea = "
        Sql2 = Sql2 & DBSet(Rs.Fields(3).Value, "N")
        Sql2 = Sql2 & " group by 1 "
        Sql2 = Sql2 & " order by 1 "
        
        Set Rs1 = New ADODB.Recordset
        Rs1.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
        Neto = Array(0, 0, 0, 0, 0, 0, 0, 0, 0)
        
        While Not Rs1.EOF
            Incluido = 0
            For i = 0 To 8
                If Calibre(i) = -1 Or Calibre(i) = DBLet(Rs1.Fields(0).Value, "N") Then
                    Calibre(i) = DBLet(Rs1.Fields(0).Value, "N")
                    Neto(i) = Neto(i) + DBLet(Rs1.Fields(1).Value, "N")
                    Incluido = 1
                    Exit For
                End If
            Next i
            
            Rs1.MoveNext
        Wend
        Set Rs1 = Nothing
        
        Sql = "insert into tmpinfventas (codusu, fecalbar, numalbar, numlinea, numcajas, pesoneto, "
        Sql = Sql & " calibre1, neto1, calibre2, neto2, calibre3, neto3, calibre4, neto4, calibre5, neto5, "
        Sql = Sql & " calibre6, neto6, calibre7, neto7, calibre8, neto8, calibre9, neto9, impcalibres) values (" & DBSet(vUsu.Codigo, "N") & ","
        Sql = Sql & DBSet(Rs.Fields(1).Value, "F") & "," & DBSet(Rs.Fields(2).Value, "N") & "," & DBSet(Rs.Fields(3).Value, "N") & ","
        Sql = Sql & DBSet(Rs.Fields(4).Value, "N") & "," 'numero de cajas
        Sql = Sql & DBSet(Rs.Fields(5).Value, "N") & "," 'peso neto
        Sql = Sql & DBSet(Calibre(0), "N") & "," & DBSet(Neto(0), "N") & "," ' calibre 1
        Sql = Sql & DBSet(Calibre(1), "N") & "," & DBSet(Neto(1), "N") & "," ' calibre 2
        Sql = Sql & DBSet(Calibre(2), "N") & "," & DBSet(Neto(2), "N") & "," ' calibre 3
        Sql = Sql & DBSet(Calibre(3), "N") & "," & DBSet(Neto(3), "N") & "," ' calibre 4
        Sql = Sql & DBSet(Calibre(4), "N") & "," & DBSet(Neto(4), "N") & "," ' calibre 5
        Sql = Sql & DBSet(Calibre(5), "N") & "," & DBSet(Neto(5), "N") & "," ' calibre 6
        Sql = Sql & DBSet(Calibre(6), "N") & "," & DBSet(Neto(6), "N") & "," ' calibre 7
        Sql = Sql & DBSet(Calibre(7), "N") & "," & DBSet(Neto(7), "N") & "," ' calibre 8
        Sql = Sql & DBSet(Calibre(8), "N") & "," & DBSet(Neto(8), "N") & "," ' calibre 9
        Sql = Sql & DBSet(Incluido, "N") & ")" ' si se han podido incluir todos los calibres
        
        conn.Execute Sql
      
        Rs.MoveNext
    Wend
    
    Label4(27).visible = False
    Pb1.visible = False
    
    ProcesarCambiosCalibres = HayRegistros("tmpinfventas", "codusu = " & vUsu.Codigo)
    
    Exit Function
    
eProcesarCambiosCalibres:
    If Err.Number <> 0 Then
        ProcesarCambiosCalibres = False
    End If
End Function


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(0)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection
Dim ExporExcel As Boolean

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
     For H = 0 To imgBuscar.Count - 1
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Next H

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, H, W
    indFrame = 5
    tabla = "albaran"
    
    Label4(27).visible = False
    Pb1.visible = False
    
    ActivarAyuda True
    CargaCombo
    Combo1(0).ListIndex = 2
    Combo1(1).ListIndex = 0
    
    '[Monica]18/11/2011: a�adida la pagina de excel
    If vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Then
        ExporExcel = Dir(App.path & "\Ventas.exe", vbArchive) <> "" And _
                             Dir(App.path & "\PlantillaVtaCatadau.xls", vbArchive) <> "" And _
                             Dir(App.path & "\ControlVtas.cfg", vbArchive) <> ""
    
    Else
        ExporExcel = Dir(App.path & "\Ventas.exe", vbArchive) <> "" And _
                             Dir(App.path & "\PlantillaVta.xls", vbArchive) <> "" And _
                             Dir(App.path & "\ControlVtas.cfg", vbArchive) <> ""
    End If
    Check1(8).Enabled = ExporExcel
    Check1(8).visible = ExporExcel
    imgAyuda(2).Picture = frmPpal.ImageListB.ListImages(10).Picture
    imgAyuda(2).visible = ExporExcel
    imgAyuda(2).Enabled = ExporExcel
    
    optList1_Click (0)
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmCla_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clases
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
    PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub frmComis_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Comisionista
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmDes_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Destinos
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") '[Monica]05/12/2018: ampliamos mascara
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFor_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de forfaits
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMar_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Marcas
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        Sql = " {variedades.codvarie} in (" & CadenaSeleccion & ")"
        Sql2 = " {variedades.codvarie} in [" & CadenaSeleccion & "]"
    Else
        Sql = " {variedades.codvarie} = -1 "
    End If
    If Not AnyadirAFormula(cadselect, Sql) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub


Private Sub frmMensCate_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String

    If SeleccionadosTodos Then
        ' no hacemos nada

    Else
        If CadenaSeleccion <> "" Then
            '[Monica]17/06/2013: a�adida la categoria
            If CategoriaValorNulo Then
                Sql = " ({albaran_variedad.categori} is null or {albaran_variedad.categori} ='' or {albaran_variedad.categori} in (" & CadenaSeleccion & "))"
                Sql2 = " (isnull({albaran_variedad.categori}) or {albaran_variedad.categori} ='' or {albaran_variedad.categori} in [" & CadenaSeleccion & "])"
            Else
                Sql = " {albaran_variedad.categori} in (" & CadenaSeleccion & ")"
                Sql2 = " {albaran_variedad.categori} in [" & CadenaSeleccion & "]"
            End If
        Else
            If CategoriaValorNulo Then
                Sql = " ({albaran_variedad.categori} is null or {albaran_variedad.categori} ='') "
                Sql2 = " (isnull({albaran_variedad.categori}) or {albaran_variedad.categori} ='') "
            Else
                Sql = " {albaran_variedad.categori} = '-1' "
            End If
        End If
        If Not AnyadirAFormula(cadselect, Sql) Then Exit Sub
        If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub
    End If
End Sub

Private Sub frmMensContr_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String

    If SeleccionadosTodos Then
        ' no hacemos nada

    Else
        If CadenaSeleccion <> "" Then
            '[Monica]17/10/2016: a�adido el contrato
            If CategoriaValorNulo Then
                Sql = " ({albaran.nrocontra} is null or {albaran.nrocontra} ='' or {albaran.nrocontra} in (" & CadenaSeleccion & "))"
                Sql2 = " (isnull({albaran.nrocontra}) or {albaran.nrocontra} ='' or {albaran.nrocontra} in [" & CadenaSeleccion & "])"
            Else
                Sql = " {albaran.nrocontra} in (" & CadenaSeleccion & ")"
                Sql2 = " {albaran.nrocontra} in [" & CadenaSeleccion & "]"
            End If
        Else
            If CategoriaValorNulo Then
                Sql = " ({albaran.nrocontra} is null or {albaran.nrocontra} ='') "
                Sql2 = " (isnull({albaran.nrocontra}) or {albaran.nrocontra} ='') "
            Else
                Sql = " {albaran.nrocontra} = '-1' "
            End If
        End If
        If Not AnyadirAFormula(cadselect, Sql) Then Exit Sub
        If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub
    End If
End Sub


Private Sub frmMensClase_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        Sql = " {variedades.codclase} in (" & CadenaSeleccion & ")"
        Sql2 = " {variedades.codclase} in [" & CadenaSeleccion & "]"
    Else
        Sql = " {variedades.codclase} = -1 "
    End If
    If Not AnyadirAFormula(cadselect, Sql) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub

Private Sub frmMensCliente_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        Sql = " {albaran.codclien} in (" & CadenaSeleccion & ")"
        Sql2 = " {albaran.codclien} in [" & CadenaSeleccion & "]"
    Else
        Sql = " {albaran.codclien} = -1 "
    End If
    If Not AnyadirAFormula(cadselect, Sql) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub

Private Sub frmMensComisionista_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        Sql = " {albaran.codcomis} in (" & CadenaSeleccion & ")"
        Sql2 = " {albaran.codcomis} in [" & CadenaSeleccion & "]"
    Else
        Sql = " {albaran.codcomis} = -1 "
    End If
    
    '[Monica]12/12/2013: si esta marcado se refiere al de las l�neas
    If Check1(11).Value = 1 Then
        Sql = Replace(Sql, "{albaran.codcomis}", "{albaran_variedad.codcomis}")
        Sql2 = Replace(Sql2, "{albaran.codcomis}", "{albaran_variedad.codcomis}")
    End If
    
    If Not AnyadirAFormula(cadselect, Sql) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub


Private Sub frmMensDestino_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        Sql = " {destinos.coddesti} in (" & CadenaSeleccion & ")"
        Sql2 = " {destinos.coddesti} in [" & CadenaSeleccion & "]"
    Else
        Sql = " {destinos.coddesti} = -1 "
    End If
    If Not AnyadirAFormula(cadselect, Sql) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub

Private Sub frmMensForfait_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        Sql = " {forfaits.codforfait} in (" & CadenaSeleccion & ")"
        Sql2 = " {forfaits.codforfait} in [" & CadenaSeleccion & "]"
    Else
        Sql = " {forfaits.codforfait} = -1 "
    End If
    If Not AnyadirAFormula(cadselect, Sql) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub

Private Sub frmMensMarca_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        Sql = " {albaran_variedad.codmarca} in (" & CadenaSeleccion & ")"
        Sql2 = " {albaran_variedad.codmarca} in [" & CadenaSeleccion & "]"
    Else
        Sql = " {albaran_variedad.codmarca} = -1 "
    End If
    If Not AnyadirAFormula(cadselect, Sql) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub

Private Sub frmMensMercado_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        Sql = " {albaran.codtimer} in (" & CadenaSeleccion & ")"
        Sql2 = " {albaran.codtimer} in [" & CadenaSeleccion & "]"
    Else
        Sql = " {albaran.codtimer} = -1 "
    End If
    If Not AnyadirAFormula(cadselect, Sql) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub

Private Sub frmMensPais_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        Sql = " {destinos.codpaise} in (" & CadenaSeleccion & ")"
        Sql2 = " {dsetinos.codpaise} in [" & CadenaSeleccion & "]"
    Else
        Sql = " {destinos.codpaise} = -1 "
    End If
    If Not AnyadirAFormula(cadselect, Sql) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub


End Sub

Private Sub frmMensVariedad_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        Sql = " {variedades.codvarie} in (" & CadenaSeleccion & ")"
        Sql2 = " {variedades.codvarie} in [" & CadenaSeleccion & "]"
    Else
        Sql = " {variedades.codvarie} = -1 "
    End If
    If Not AnyadirAFormula(cadselect, Sql) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub

Private Sub frmPais_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Paises
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTMe_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Tipo de mercado
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Variedades
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
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
    imgFec(0).Tag = Index + 16 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtCodigo(Index + 16).Text <> "" Then frmC.NovaData = txtCodigo(Index + 16).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtCodigo(CByte(imgFec(0).Tag))
    ' ***************************
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1 'CLASE
            AbrirFrmClases (Index)
        
        Case 2, 3 'VARIEDADES
            AbrirFrmVariedades (Index)
        
        Case 4, 5 'CLIENTES
            AbrirFrmClientes (Index)
        
        Case 6, 7 'DESTINOS
            AbrirFrmDestinos (Index)
        
        Case 8, 9 'FORFAIS
            AbrirFrmForfaits (Index)
        
        Case 10, 11 'MARCAS
            AbrirFrmMarcas (Index)
        
        Case 12, 13 'TIPOS DE MERCADO
            AbrirFrmMercados (Index)
    
        Case 14, 15 'PAIS
            AbrirFrmPais (Index)
            
        Case 16, 17 'COMISIONISTA
            AbrirFrmComisionista (Index + 2)
    
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub Optcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub OptNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub optList1_Click(Index As Integer)
    VisualizarChecks (Index)
End Sub

Private Sub optList1_KeyPress(Index As Integer, KeyAscii As Integer)
    VisualizarChecks (Index)
End Sub


Private Sub VisualizarChecks(Index As Integer)
    If Index = 9 Or Index = 12 Then
        Check1(4).Enabled = (optList1(Index).Value = True)
    Else
        Check1(4).Enabled = False
        Check1(4).Value = 0
    End If
    If Index = 0 Or Index = 2 Then
        Check1(6).Enabled = (optList1(Index).Value = True)
    Else
        Check1(6).Enabled = False
        Check1(6).Value = 0
    End If
    If Index = 0 Then
        Check1(7).Enabled = (optList1(Index).Value = True)
        Check1(9).Enabled = (optList1(Index).Value = True)
    Else
        Check1(7).Enabled = False
        Check1(7).Value = 0
        Check1(9).Enabled = False
        Check1(9).Value = 0
    End If
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'14/02/2007 antes
'    KEYpress KeyAscii
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'clase desde
            Case 1: KEYBusqueda KeyAscii, 1 'clase hasta
            Case 2: KEYBusqueda KeyAscii, 2 'variedad desde
            Case 3: KEYBusqueda KeyAscii, 3 'variedad hasta
            Case 4: KEYBusqueda KeyAscii, 4 'cliente desde
            Case 5: KEYBusqueda KeyAscii, 5 'cliente hasta
            Case 6: KEYBusqueda KeyAscii, 6 'destino desde
            Case 7: KEYBusqueda KeyAscii, 7 'destino hasta
            Case 16: KEYFecha KeyAscii, 0 'fecha desde
            Case 17: KEYFecha KeyAscii, 1 'fecha hasta
            Case 8: KEYBusqueda KeyAscii, 8 'forfaits desde
            Case 9: KEYBusqueda KeyAscii, 9 'forfaits hasta
            Case 10: KEYBusqueda KeyAscii, 10 'marca desde
            Case 11: KEYBusqueda KeyAscii, 11 'marca hasta
            Case 12: KEYBusqueda KeyAscii, 12 'tipo de mercado desde
            Case 13: KEYBusqueda KeyAscii, 13 'tipo de mercado hasta
            Case 14: KEYBusqueda KeyAscii, 14 'pais desde
            Case 15: KEYBusqueda KeyAscii, 15 'pais hasta
            Case 18: KEYBusqueda KeyAscii, 16 'comisionista desde
            Case 19: KEYBusqueda KeyAscii, 17 'comisionista hasta
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

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
            
        Case 0, 1 'CLASE
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "clases", "nomclase", "codclase", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
        
        Case 2, 3 'VARIEDAD
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
         
        Case 4, 5 'CLIENTE
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "clientes", "nomclien", "codclien", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
            
            ' solo se puede introducir destino si cliente desde y hasta son iguales
            txtCodigo(6).Enabled = (txtCodigo(4).Text = txtCodigo(5).Text)
            imgBuscar(6).Enabled = txtCodigo(6).Enabled
            imgBuscar(7).Enabled = txtCodigo(7).Enabled
            If Not txtCodigo(6).Enabled Then
                txtCodigo(6).Text = ""
                txtNombre(6).Text = ""
            End If
            txtCodigo(7).Enabled = (txtCodigo(4).Text = txtCodigo(5).Text)
            If Not txtCodigo(7).Enabled Then
                txtCodigo(7).Text = ""
                txtNombre(7).Text = ""
            End If
            
            If Index = 5 Then
                If txtCodigo(6).Enabled Then
                    PonerFoco txtCodigo(6)
                Else
                    PonerFoco txtCodigo(16)
                End If
            End If
            
        Case 6, 7 'DESTINO
            If txtCodigo(4).Text <> "" And txtCodigo(4).Text = txtCodigo(5).Text Then
                txtNombre(Index).Text = DevuelveDesdeBDNew(cAgro, "destinos", "nomdesti", "codclien", txtCodigo(4).Text, "N", , "coddesti", txtCodigo(Index).Text, "N")
                '[Monica]05/12/2018: ampliamos mascara
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
            End If
            
        Case 16, 17 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 8, 9 'FORFAITS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "forfaits", "nomconfe", "codforfait", "T")
            
        Case 10, 11 'MARCA
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "marcas", "nommarca", "codmarca", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
        
        Case 12, 13 'TIPO DE MERCADO
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "tipomer", "nomtimer", "codtimer", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
        
        Case 14, 15 'PAIS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "paises", "nompaise", "codpaise", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
                        
        Case 18, 19 'COMISIONISTAS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "agencias", "nomtrans", "codtrans", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 11700
        Me.FrameCobros.Width = 10515
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
        .EnvioEMail = False
        .NombreRPT = cadNombreRPT
        .ConSubInforme = ConSubInforme
        .Opcion = 1
        .Show vbModal
    End With
End Sub

Private Sub AbrirFrmClientes(indice As Integer)
    indCodigo = indice
    Set frmCli = New frmClientes
    frmCli.DatosADevolverBusqueda = "0|2|"
    frmCli.Show vbModal
    Set frmCli = Nothing
End Sub

Private Sub AbrirFrmClases(indice As Integer)
    indCodigo = indice
    Set frmCla = New frmManClases
    frmCla.DatosADevolverBusqueda = "0|1|"
    frmCla.Show vbModal
    Set frmCla = Nothing
End Sub

Private Sub AbrirFrmVariedades(indice As Integer)
    indCodigo = indice
    Set frmVar = New frmManVariedad
    frmVar.DatosADevolverBusqueda = "0|1|"
    frmVar.DeConsulta = True
    frmVar.CodigoActual = txtCodigo(indCodigo)
    frmVar.Show vbModal
    Set frmVar = Nothing
End Sub

Private Sub AbrirFrmDestinos(indice As Integer)
    indCodigo = indice
    Set frmDes = New frmDestCli
    frmDes.DatosADevolverBusqueda = "0|1|"
'    frmDes.DeConsulta = True
    frmDes.Cliente = txtCodigo(4).Text
    frmDes.CodigoActual = txtCodigo(indCodigo)
    frmDes.Show vbModal
    Set frmDes = Nothing
End Sub

Private Sub AbrirFrmForfaits(indice As Integer)
    indCodigo = indice
'    Set frmFor = New frmManForfaits
'    frmFor.DatosADevolverBusqueda = "0|1|"
'    frmFor.DeConsulta = True
'    frmFor.CodigoActual = txtCodigo(indCodigo)
'    frmFor.Show vbModal
'    Set frmFor = Nothing

    Set frmFor = New frmBasico2

    AyudaForfaits frmFor
    
    Set frmFor = Nothing


End Sub
 
Private Sub AbrirFrmMarcas(indice As Integer)
    indCodigo = indice
    Set frmMar = New frmManMarcas
    frmMar.DatosADevolverBusqueda = "0|1|"
    frmMar.DeConsulta = True
    frmMar.CodigoActual = txtCodigo(indCodigo)
    frmMar.Show vbModal
    Set frmMar = Nothing
End Sub

Private Sub AbrirFrmMercados(indice As Integer)
    indCodigo = indice
    Set frmTMe = New frmManTipMerc
    frmTMe.DatosADevolverBusqueda = "0|1|"
    frmTMe.DeConsulta = True
    frmTMe.CodigoActual = txtCodigo(indCodigo)
    frmTMe.Show vbModal
    Set frmTMe = Nothing
End Sub

Private Sub AbrirFrmPais(indice As Integer)
    indCodigo = indice
    Set frmPais = New frmManPaises
    frmPais.DatosADevolverBusqueda = "0|1|"
    frmPais.DeConsulta = True
    frmPais.CodigoActual = txtCodigo(indCodigo)
    frmPais.Show vbModal
    Set frmPais = Nothing
End Sub

Private Sub AbrirFrmComisionista(indice As Integer)
    indCodigo = indice
    Set frmComis = New frmManAgencias
    frmComis.DatosADevolverBusqueda = "0|1|"
    frmComis.DeConsulta = True
    frmComis.CodigoActual = txtCodigo(indCodigo)
    frmComis.Show vbModal
    Set frmComis = Nothing
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
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Rs.EOF Then
        MsgBox "No hay datos para mostrar en el Informe.", vbInformation
        HayRegistros = False
    Else
        HayRegistros = True
    End If

End Function


Private Function NroGastosMayoraCuatro(cadTABLA As String, cadwhere As String) As Boolean
Dim Sql As String
Dim cadTabla2 As String
Dim cadWHERE2 As String

    NroGastosMayoraCuatro = False

    cadTabla2 = "(" & cadTABLA & ") inner join albaran_costes on albaran_variedad.numalbar = albaran_costes.numalbar "
    cadTabla2 = cadTabla2 & " and albaran_variedad.numlinea = albaran_costes.numlinea "

    If cadwhere <> "" Then
        cadWHERE2 = QuitarCaracterACadena(cadwhere, "{")
        cadWHERE2 = QuitarCaracterACadena(cadWHERE2, "}")
        cadWHERE2 = QuitarCaracterACadena(cadWHERE2, "_1")
    End If

'    Sql = "select count(*) from nombcoste"
    Sql = "select count(distinct albaran_costes.codcoste) from " & cadTabla2
    If cadWHERE2 <> "" Then Sql = Sql & " where " & cadWHERE2
    
    
    NroGastosMayoraCuatro = (TotalRegistros(Sql) > 4)

End Function


Private Sub ActivarAyuda(sn As Boolean)
Dim i As Integer

    For i = 0 To imgAyuda.Count - 1
        imgAyuda(i).Picture = frmPpal.ImageListB.ListImages(10).Picture
        imgAyuda(i).visible = sn
        imgAyuda(i).Enabled = sn
    Next i
End Sub


Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           ' "____________________________________________________________"
            vCadena = "S�lo se utiliza si el listado est� ordenado por Variedad-Fecha " & vbCrLf & _
                      "o por Cliente-Destino-Variedad. " & vbCrLf & vbCrLf & _
                      "Informe que saca la evoluci�n de precios de los albaranes de " & vbCrLf & _
                      "salida. Precio provisional, definitivo y precio de factura y " & vbCrLf & _
                      "margenes entre los distintos precios." & vbCrLf & vbCrLf & _
                      "El Importe de Venta de todos los informes se obtiene con el " & vbCrLf & _
                      "precio facturado si lo tiene y si no con el precio definitivo " & vbCrLf & _
                      "si lo tiene y si no con el precio provisional."
                      
        Case 1
           ' "____________________________________________________________"
            vCadena = "S�lo se utiliza si el listado est� ordenado por Variedad-Fecha. " & vbCrLf & _
                      vbCrLf & _
                      "Saca la informaci�n del albar�n de venta junto con el nro de traza " & vbCrLf & _
                      "donde se indican cuales son los albaranes de entrada asociados." & vbCrLf & _
                      "" & vbCrLf

        Case 2
           ' "____________________________________________________________"
            vCadena = "Saca a una hoja excel todos los datos detallados de los albaranes " & vbCrLf & _
                      "seleccionados. " & vbCrLf & _
                      vbCrLf & _
                      "Ignora el orden seleccionado y lo ordena por nro.albar�n y fecha." & vbCrLf & _
                      "Tendr� en cuenta el punto de S�lo Facturados si est� seleccionado. " & vbCrLf & vbCrLf
                      
            '[Monica]20/11/2015: solo en el caso de catadau es una excel diferente para anecoop
            If vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Then
                vCadena = "Saca a una hoja excel todos los datos de expedientes a nivel de  " & vbCrLf & _
                          "calibres. " & vbCrLf & _
                          vbCrLf & _
                          "Ignora el orden seleccionado y lo ordena por nro.albar�n y fecha." & vbCrLf & vbCrLf
            End If
                      

        Case 3
           ' "____________________________________________________________"
            vCadena = "S�lo se utiliza si el listado est� ordenado por Variedad-Fecha. " & vbCrLf & _
                      vbCrLf & _
                      "Saca la informaci�n del albar�n de venta junto con el tipo de  " & vbCrLf & _
                      "variedad." & vbCrLf & _
                      "" & vbCrLf

        Case 4
           ' "____________________________________________________________"
            vCadena = "S�lo se utiliza si el listado est� ordenado por Variedad-Fecha y " & vbCrLf & _
                      "por Cliente-Destino-Variedad.  " & vbCrLf & vbCrLf & _
                      "Saca �nicamente los datos de los albaranes no Cobrados. " & vbCrLf & _
                      "" & vbCrLf

        Case 5
           ' "____________________________________________________________"
            vCadena = "S�lo se utiliza si el listado est� ordenado por Comisionista-Variedad-Fecha." & vbCrLf & _
                      "Si est� marcado se refiere al comisionista de las l�neas de albar�n.  " & vbCrLf & vbCrLf & _
                      "En caso contrario se refiere al comisionista del albar�n. " & vbCrLf & _
                      "" & vbCrLf



    End Select
    MsgBox vCadena, vbInformation, "Descripci�n de Ayuda"
    
End Sub

' ********* si n'hi han combos a la cap�alera ************
Private Sub CargaCombo()
Dim i As Integer
Dim Cad As String
Dim Rs As ADODB.Recordset

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For i = 0 To Combo1.Count - 1
        Combo1(i).Clear
    Next i
    
    Combo1(0).AddItem "Cooperativa"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Ajena"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "Todas"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2

    'Tipo de variedad
    Combo1(1).AddItem "Todas"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    
    Cad = "SELECT * FROM tipovarie ORDER BY codtipo"
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not Rs.EOF
        Combo1(1).AddItem Rs!nomtipo
        Combo1(1).ItemData(Combo1(1).NewIndex) = Rs!codtipo + 1
        Rs.MoveNext
    Wend
    Rs.Close

End Sub




Private Function ProcesarCambiosEvolucion(cadTABLA, cadwhere As String) As Boolean
Dim Sql As String
Dim SQL1 As String
Dim Sql2 As String
Dim i As Integer
Dim HayReg As Long
Dim b As Boolean
Dim Rs As ADODB.Recordset
Dim Rsx As ADODB.Recordset
Dim TotalGastos As Currency
Dim TotalGastosReales As Currency
Dim PesoCaja As Currency
Dim PesoReal As Currency
Dim ImpVenta As Currency
Dim Facturado As Byte
Dim Cobrado As Byte
Dim cadTabla2 As String

Dim Coste1 As Integer
Dim Coste2 As Integer
Dim Coste3 As Integer
Dim Coste4 As Integer
Dim Coste5 As Integer

Dim Gasto1 As Currency
Dim Gasto2 As Currency
Dim Gasto3 As Currency
Dim Gasto4 As Currency
Dim Gasto5 As Currency
Dim Costes As Integer
Dim GastosEnvases As Currency
Dim GastosPortes As Currency
Dim GastosComision As Currency

Dim Sql3 As String
Dim PreProv As Currency
Dim PreDef As Currency
Dim PreFact As Currency

On Error GoTo eProcesarCambiosEvolucion

    HayReg = 0
    
    ProcesarCambiosEvolucion = False
    
    conn.Execute "delete from tmpinfventas where codusu = " & DBSet(vUsu.Codigo, "N")
        
    If cadwhere <> "" Then
        cadwhere = QuitarCaracterACadena(cadwhere, "{")
        cadwhere = QuitarCaracterACadena(cadwhere, "}")
        cadwhere = QuitarCaracterACadena(cadwhere, "_1")
    End If
        
    SQL1 = "select albaran.fechaalb, albaran.numalbar, albaran_variedad.numlinea, "
    SQL1 = SQL1 & "sum(facturas_variedad.impornet) from " & cadTABLA
    SQL1 = SQL1 & " where (1 = 1) "
    If cadwhere <> "" Then SQL1 = SQL1 & " and " & cadwhere
    SQL1 = SQL1 & " group by 1, 2, 3"
    SQL1 = SQL1 & " order by 1, 2, 3"
        
    Set Rs = New ADODB.Recordset
    Rs.Open SQL1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Label4(27).visible = True
    Pb1.visible = True
        
    HayReg = TotalRegistrosConsulta(SQL1)
    
    Pb1.Max = HayReg
    Pb1.Value = 0
    
    Sql = ""
    
    While Not Rs.EOF
        IncrementarProgresNew Pb1, 1
    
        ImpVenta = DBLet(Rs.Fields(3).Value, "N")
        
        Cobrado = 0
        If Not IsNull(Rs.Fields(3).Value) Then Cobrado = AlbaranCobradoTesoreria(DBLet(Rs.Fields(1).Value, "N"), DBLet(Rs.Fields(2).Value, "N"))

        
        Sql = Sql & "(" & DBSet(vUsu.Codigo, "N") & ","
        Sql = Sql & DBSet(Rs.Fields(1).Value, "N") & "," ' albaran
        Sql = Sql & DBSet(Rs.Fields(2).Value, "N") & "," ' linea
        Sql = Sql & DBSet(Rs.Fields(0).Value, "F") & "," ' fechaalbaran
        Sql = Sql & DBSet(ImpVenta, "N") & "," 'importe facturado
        Sql = Sql & DBSet(Cobrado, "N") & ")," ' cobrado
        
'        Conn.Execute Sql
      
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    
    If Sql <> "" Then
        ' quitamos la ultima coma
        Sql = Mid(Sql, 1, Len(Sql) - 1)
                                                '                              impfacturado
        Sql3 = "insert into tmpinfventas (codusu, numalbar, numlinea, fecalbar, impventa, cobrado ) values "
        Sql3 = Sql3 & Sql
        
        conn.Execute Sql3
    End If
    
    ProcesarCambiosEvolucion = True

    Label4(27).visible = False
    Pb1.visible = False
    
eProcesarCambiosEvolucion:
    If Err.Number <> 0 Then
        ProcesarCambiosEvolucion = False
    End If
End Function


Private Function CargarTemporal() As Boolean
Dim Sql As String
Dim SqlIns As String
Dim Rs As ADODB.Recordset

    On Error GoTo eCargarTemporal

    CargarTemporal = False


    Sql = "delete from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
    conn.Execute Sql
    
    Sql = "select " & vUsu.Codigo & ", albaran.fechaalb, albaran_calibre.numcajas, albaran_calibre.codcalib, albaran_calibre.numalbar, albaran_calibre.numlinea, albaran_calibre.numline1, albaran_calibre.pesoneto, sum(facturas_calibre.impornet) importe "
    Sql = Sql & " from ((tmpinfventas inner join albaran on tmpinfventas.numalbar = albaran.numalbar) inner join albaran_calibre on tmpinfventas.numalbar = albaran_calibre.numalbar and tmpinfventas.numlinea = albaran_calibre.numlinea)  "
    Sql = Sql & " left join facturas_calibre on albaran_calibre.numalbar = facturas_calibre.numalbar and albaran_calibre.numlinea = facturas_calibre.numlinealbar and albaran_calibre.numline1 = facturas_calibre.numline1albar "
    Sql = Sql & " where tmpinfventas.codusu = " & DBSet(vUsu.Codigo, "N")
    Sql = Sql & " group by 1,2,3,4,5,6,7,8 order by 1,2,3,4 "
                                            'fecalbar, numcajas, codcalib, numalbar,  numlinea,  numline1,  pesoneto, importe
    SqlIns = "insert into tmpinformes (codusu, fecha1, importe1, importe2, importeb1, importeb2, importeb3, importe3, importe4 )     "
    SqlIns = SqlIns & Sql
    conn.Execute SqlIns
    
'    ' quiere la clase
'    SqlIns = "update tmpinfventas, albaran_variedad, variedades, clases "
'    SqlIns = SqlIns & " set tmpinfventas.importe5 = variedades.codclase, tmpinfventas.nombre1 = clases.nomclase "
'    SqlIns = SqlIns & " where tmpinfventas.codusu = " & DBSet(vUsu.Codigo, "N")
'    SqlIns = SqlIns & " and tmpinfventas.importeb1 = albaran_variedad.numalbar "
'    SqlIns = SqlIns & " and tmpinfventas.importeb2 = albaran_variedad.numlinea "
'    SqlIns = SqlIns & " and albaran_variedad.codvarie = variedades.codvarie "
'    SqlIns = SqlIns & " and variedades.codclase = clases.codclase "
'
'    conn.Execute SqlIns
'
    
    
    
    CargarTemporal = True
    Exit Function
    
eCargarTemporal:
    MuestraError Err.Number, "Cargar Temporal", Err.Description
End Function
