VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmComFacturar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturas Compra Proveedores"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   13935
   Icon            =   "frmComFacturar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   13935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkVistaPrevia 
      Caption         =   "Vista previa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   11205
      TabIndex        =   56
      Top             =   195
      Width           =   1605
   End
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   135
      TabIndex        =   54
      Top             =   45
      Width           =   1830
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   55
         Top             =   180
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Pedir Datos"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver Albaranes"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Generar Facturas"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameIntro 
      Height          =   1590
      Left            =   135
      TabIndex        =   7
      Top             =   765
      Width           =   13650
      Begin VB.TextBox Text1 
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
         Left            =   7770
         MaxLength       =   15
         TabIndex        =   4
         Tag             =   "Concepto|T|N|||scafpc|confacpr|||"
         Text            =   "Text1"
         Top             =   390
         Width           =   5565
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Contabiliz."
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
         Left            =   6045
         TabIndex        =   49
         Top             =   1080
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Tesoreria"
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
         Left            =   6060
         TabIndex        =   48
         Top             =   720
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1240
      End
      Begin VB.TextBox Text2 
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
         Index           =   3
         Left            =   1530
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   14
         Text            =   "Text2"
         Top             =   1050
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Left            =   3420
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Recepción|F|N|||scafpc|fecrecep|dd/mm/yyyy|N|"
         Top             =   400
         Width           =   1485
      End
      Begin VB.TextBox Text2 
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
         Left            =   8850
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   12
         Text            =   "Text2"
         Top             =   1000
         Width           =   4470
      End
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
         Index           =   5
         Left            =   8130
         MaxLength       =   5
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1000
         Width           =   660
      End
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
         Index           =   3
         Left            =   550
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "Cod. Proveedor|N|N|0|999999|scafpc|codprove|000000|S|"
         Text            =   "Text1"
         Top             =   1050
         Width           =   960
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Factura|F|N|||scafpc|fecfactu|dd/mm/yyyy|S|"
         Top             =   400
         Width           =   1485
      End
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
         Index           =   0
         Left            =   240
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Nº Factura|T|N|||scafpc|numfactu||S|"
         Text            =   "Text1 7"
         Top             =   400
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Concepto"
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
         Left            =   7815
         TabIndex        =   53
         Top             =   135
         Width           =   1215
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   5025
         Picture         =   "frmComFacturar.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   150
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   2790
         Picture         =   "frmComFacturar.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   135
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   7815
         ToolTipText     =   "Buscar banco propio"
         Top             =   1035
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   240
         ToolTipText     =   "Buscar proveedor"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fec. Recepción"
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
         Left            =   3420
         TabIndex        =   13
         Top             =   150
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Cta. Prevista Pago"
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
         Left            =   7815
         TabIndex        =   11
         Top             =   750
         Width           =   1980
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   795
         Width           =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "Fec.Factura"
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
         Index           =   29
         Left            =   1560
         TabIndex        =   9
         Top             =   150
         Width           =   1500
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Factura"
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
         Index           =   28
         Left            =   240
         TabIndex        =   8
         Top             =   150
         Width           =   1095
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   2160
      MaxLength       =   4
      TabIndex        =   51
      Text            =   "Text1"
      Top             =   1110
      Width           =   660
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   4
      Left            =   2880
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   50
      Text            =   "Text2"
      Top             =   1110
      Width           =   3615
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4050
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7144
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
      NumItems        =   0
   End
   Begin VB.Frame FrameFactura 
      Height          =   4150
      Left            =   8010
      TabIndex        =   15
      Top             =   2310
      Width           =   5760
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "&Generar"
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
         Left            =   1485
         TabIndex        =   35
         Top             =   3660
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
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
         Left            =   315
         TabIndex        =   36
         Top             =   3660
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   3540
         MaxLength       =   15
         TabIndex        =   42
         Tag             =   "Importe IVA 1|N|N|0||scafac|imporiv1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   1440
         Width           =   2070
      End
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
         Index           =   8
         Left            =   3540
         MaxLength       =   15
         TabIndex        =   41
         Tag             =   "Base Imponible 3|N|N|0||scafac|baseimp3|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   990
         Width           =   2070
      End
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
         Index           =   7
         Left            =   3540
         MaxLength       =   15
         TabIndex        =   40
         Tag             =   "Base Imponible 2 |N|N|0||scafac|baseimp2|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   615
         Width           =   2070
      End
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
         Left            =   3540
         MaxLength       =   15
         TabIndex        =   38
         Tag             =   "Base Imponible 1|N|N|0||scafac|baseimp1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   240
         Width           =   2070
      End
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
         Index           =   12
         Left            =   345
         MaxLength       =   5
         TabIndex        =   34
         Tag             =   "Codigo IVA 3|N|S|0|99|scafac|codiva3|00|N|"
         Text            =   "Text1 7"
         Top             =   3015
         Width           =   585
      End
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
         Index           =   11
         Left            =   345
         MaxLength       =   5
         TabIndex        =   33
         Tag             =   "Codigo IVA 2|N|S|0|99|scafac|codiva2|00|N|"
         Text            =   "Text1 7"
         Top             =   2640
         Width           =   585
      End
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
         Index           =   10
         Left            =   345
         MaxLength       =   5
         TabIndex        =   32
         Tag             =   "Codigo IVA 1|N|S|0|99|scafac|codiva1|00|N|"
         Text            =   "Text1 7"
         Top             =   2280
         Width           =   555
      End
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
         Index           =   16
         Left            =   1665
         MaxLength       =   15
         TabIndex        =   25
         Tag             =   "Base Imponible 1|N|N|0||scafac|baseimp1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2280
         Width           =   1710
      End
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
         Index           =   13
         Left            =   975
         MaxLength       =   5
         TabIndex        =   24
         Tag             =   "% IVA 1|N|S|0|99.90|scafac|porciva1|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   2280
         Width           =   660
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   3540
         MaxLength       =   15
         TabIndex        =   23
         Tag             =   "Importe IVA 1|N|N|0||scafac|imporiv1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2280
         Width           =   2115
      End
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
         Index           =   17
         Left            =   1665
         MaxLength       =   15
         TabIndex        =   22
         Tag             =   "Base Imponible 2 |N|N|0||scafac|baseimp2|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2640
         Width           =   1710
      End
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
         Index           =   14
         Left            =   975
         MaxLength       =   5
         TabIndex        =   21
         Tag             =   "& IVA 2|N|S|0|99.90|scafac|porciva2|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   2640
         Width           =   660
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   3540
         MaxLength       =   15
         TabIndex        =   20
         Tag             =   "Importe IVA 2|N|N|0||scafac|imporiv2|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2640
         Width           =   2115
      End
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
         Index           =   18
         Left            =   1665
         MaxLength       =   15
         TabIndex        =   19
         Tag             =   "Base Imponible 3|N|N|0||scafac|baseimp3|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   3015
         Width           =   1710
      End
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
         Index           =   15
         Left            =   990
         MaxLength       =   5
         TabIndex        =   18
         Tag             =   "% IVA 3|N|S|0|99.90|scafac|porciva3|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   3015
         Width           =   660
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   3540
         MaxLength       =   15
         TabIndex        =   17
         Tag             =   "Importe IVA 3|N|N|0||scafac|imporiv3|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   3015
         Width           =   2115
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
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
         Left            =   3570
         MaxLength       =   15
         TabIndex        =   16
         Tag             =   "Total Factura|N|N|0||scafac|totalfac|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   3720
         Width           =   2130
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   90
         ToolTipText     =   "Buscar codigo iva"
         Top             =   3030
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   90
         ToolTipText     =   "Buscar codigo iva"
         Top             =   2655
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   90
         ToolTipText     =   "Buscar codigo iva"
         Top             =   2310
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   3435
         TabIndex        =   47
         Top             =   990
         Width           =   135
      End
      Begin VB.Line Line2 
         X1              =   3555
         X2              =   5585
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Label Label1 
         Caption         =   "Importe dto.gnral"
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
         Left            =   1665
         TabIndex        =   46
         Top             =   990
         Width           =   1800
      End
      Begin VB.Label Label1 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   3435
         TabIndex        =   45
         Top             =   615
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "Importe dto.pp"
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
         Left            =   1665
         TabIndex        =   44
         Top             =   615
         Width           =   1935
      End
      Begin VB.Line Line1 
         X1              =   1665
         X2              =   5595
         Y1              =   1845
         Y2              =   1845
      End
      Begin VB.Label Label1 
         Caption         =   "Base Imponible"
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
         Left            =   1665
         TabIndex        =   43
         Top             =   1440
         Width           =   1845
      End
      Begin VB.Label Label1 
         Caption         =   "Bruto Factura"
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
         Left            =   1665
         TabIndex        =   39
         Top             =   240
         Width           =   1650
      End
      Begin VB.Label Label1 
         Caption         =   "Cod."
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
         Left            =   390
         TabIndex        =   37
         Top             =   2025
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Base Imponible"
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
         Left            =   1665
         TabIndex        =   31
         Top             =   1995
         Width           =   1665
      End
      Begin VB.Label Label1 
         Caption         =   "Importe IVA"
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
         Index           =   33
         Left            =   3540
         TabIndex        =   30
         Top             =   2025
         Width           =   1965
      End
      Begin VB.Label Label1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   37
         Left            =   3375
         TabIndex        =   29
         Top             =   2025
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   36
         Left            =   11880
         TabIndex        =   28
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "TOTAL FACTURA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   39
         Left            =   3555
         TabIndex        =   27
         Top             =   3450
         Width           =   2160
      End
      Begin VB.Label Label1 
         Caption         =   "% IVA"
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
         Index           =   41
         Left            =   975
         TabIndex        =   26
         Top             =   2025
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   8190
      Top             =   5625
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   13245
      TabIndex        =   57
      Top             =   135
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ayuda"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Operador"
      Height          =   255
      Index           =   1
      Left            =   1845
      TabIndex        =   52
      Top             =   900
      Width           =   735
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   1845
      Picture         =   "frmComFacturar.frx":0122
      ToolTipText     =   "Buscar trabajador"
      Top             =   1125
      Width           =   240
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnPedirDatos 
         Caption         =   "&Pedir Datos"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnVerAlbaran 
         Caption         =   "&Ver Albaranes"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnGenerarFac 
         Caption         =   "&Generar Factura"
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
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
Attribute VB_Name = "frmComFacturar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)


'========== VBLES PRIVADAS ====================
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1

Private WithEvents frmProv As frmManProve
Attribute frmProv.VB_VarHelpID = -1
'Private WithEvents frmT As frmAdmTrabajadores  'Form Mto Trabajadores
Private WithEvents frmBanPr As frmManBanco 'Mto de Bancos propios
Attribute frmBanPr.VB_VarHelpID = -1
Private WithEvents frmTipIva As frmTipIVAConta  'Tipos de IVA de la contabilidad
Attribute frmTipIva.VB_VarHelpID = -1


Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Mantenimiento Lineas
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

'cadena donde se almacena la WHERE para la seleccion de los albaranes
'marcados para facturar
Dim cadWHERE As String

'Cuando vuelve del formulario de ver los albaranes seleccionados hay que volver
'a cargar los datos de los albaranes
Dim VerAlbaranes As Boolean

Dim PrimeraVez As Boolean

'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------

Dim dtoGn As Currency
Dim dtoPP As Currency
Dim Forpa As Integer
Dim indCodigo As Integer


Private vProve As CProveedor


Private Sub cmdCancelar_Click()
Dim i As Integer

    FrameIntro.Enabled = False
    ListView1.Enabled = True
    FrameFactura.Enabled = False
    
    
    BloquearTxt Text1(10), True
    BloquearTxt Text1(11), True
    BloquearTxt Text1(12), True
    
    
    For i = 3 To 5
        imgBuscar(i).Enabled = False
        imgBuscar(i).visible = False
    Next i
    
    
    Me.cmdCancelar.Enabled = False
    Me.cmdCancelar.visible = False
    Me.cmdGenerar.Enabled = False
    Me.cmdGenerar.visible = False

End Sub

Private Sub cmdGenerar_Click()
Dim i As Integer

    FrameIntro.Enabled = False
    ListView1.Enabled = True
    FrameFactura.Enabled = False


    For i = 3 To 5
        imgBuscar(i).Enabled = False
        imgBuscar(i).visible = False
    Next i
    
    
    Me.cmdCancelar.Enabled = False
    Me.cmdCancelar.visible = False
    Me.cmdGenerar.Enabled = False
    Me.cmdGenerar.visible = False


    BotonFacturar
    Set vProve = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If VerAlbaranes Then RefrescarAlbaranes
    VerAlbaranes = False
End Sub


Private Sub Form_Load()
Dim i As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    ' ICONITOS DE LA BARRA
'    With Me.Toolbar1
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 1   'Pedir Datos
'        .Buttons(2).Image = 3   'Ver albaranes
'        .Buttons(3).Image = 15   'Generar FActura
'        .Buttons(6).Image = 11   'Salir
'    End With
    ' ICONITOS DE LA BARRA
    
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Pedir Datos
        .Buttons(2).Image = 3   'Ver albaranes
        .Buttons(3).Image = 15   'Generar FActura
    End With
    
    'cargar IMAGES de busqueda
    Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    For i = 2 To 5
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    
    Me.FrameFactura.Enabled = False
    For i = 3 To 5
        Me.imgBuscar(i).visible = False
        Me.imgBuscar(i).Enabled = False
    Next i
    
    
    LimpiarCampos   'Limpia los campos TextBox
    InicializarListView
   
    '## A mano
    NombreTabla = "scafpc" 'cabecera facturas compras a proveedor
    NomTablaLineas = "slifpc" 'Tabla lineas de Facturacion
    Ordenacion = " ORDER BY scafpc.codprove, scafpc.numfactu, scafpc.fecfactu "
    
    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    CadenaConsulta = "Select * from " & NombreTabla
    CadenaConsulta = CadenaConsulta & " where numfactu=-1"
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
'    Else
'        PonerModo 1
    End If
    PrimeraVez = False
End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me
    'Aqui va el especifico de cada form es
    '### a mano
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    DesBloqueoManual "RECFAC"
    TerminaBloquear
'    DesBloqueoManual ("scaalp")
End Sub


Private Sub frmBanPr_DatoSeleccionado(CadenaSeleccion As String)
    'Form de Mantenimiento de Bancos Propios
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 1)
    Text1(5).Text = Format(Text1(5).Text, "0000")
    Text2(5).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmF_Selec(vFecha As Date)
Dim indice As Byte
    indice = CByte(Me.imgFecha(0).Tag)
    Text1(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmProv_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Proveedores
Dim indice As Byte
    
    indice = 3
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Proveedor
    FormateaCampo Text1(indice)
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom proveedor
End Sub

Private Sub frmTipIva_DatoSeleccionado(CadenaSeleccion As String)
'Tipos de IVA (de la Contabilidad)
    Text1(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codigiva
    FormateaCampo Text1(indCodigo)
    Text1(indCodigo + 3).Text = RecuperaValor(CadenaSeleccion, 3) '% iva
    RecalcularDatosFactura
End Sub

'Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
''Form Mantenimiento de Trabajadores
'Dim Indice As Byte
'    Indice = 4
'    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Trabajador
'    FormateaCampo Text1(Indice)
'    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
'End Sub


Private Sub imgBuscar_Click(Index As Integer)
Dim indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. Proveedor
            Set frmProv = New frmManProve
            frmProv.DatosADevolverBusqueda = "0|"
            frmProv.Show vbModal
            Set frmProv = Nothing
            indice = 3
'--monica
'        Case 1 'Operador. Trabajador
'            Indice = 4
'            Set frmT = New frmAdmTrabajadores
'            frmT.DatosADevolverBusqueda = "0"
'            frmT.Show vbModal
'            Set frmT = Nothing
       
       Case 2 'Bancos Propios
            indice = 5
            Set frmBanPr = New frmManBanco
            frmBanPr.DatosADevolverBusqueda = "0|1|"
            frmBanPr.Show vbModal
            Set frmBanPr = Nothing
    
        Case 3, 4, 5 ' codigos de iva de contabilidad
            indCodigo = Index + 7
        
            Set frmTipIva = New frmTipIVAConta
            frmTipIva.DeConsulta = True
            frmTipIva.DatosADevolverBusqueda = "0|1|2|"
            frmTipIva.CodigoActual = Text1(indCodigo).Text
            frmTipIva.Show vbModal
            Set frmTipIva = Nothing
        
    
    
    End Select
    
    PonerFoco Text1(indice)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer)
Dim indice As Byte
Dim esq As Long
Dim dalt As Long
Dim menu As Long
Dim obj As Object

   If Modo = 2 Or Modo = 0 Then Exit Sub
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
   
   frmF.NovaData = Now
   indice = Index + 1
   Me.imgFecha(0).Tag = indice
   
   PonerFormatoFecha Text1(indice)
   If Text1(indice).Text <> "" Then frmF.NovaData = CDate(Text1(indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(indice)
End Sub

Private Sub ListView1_ItemCheck(ByVal item As MSComctlLib.ListItem)
'Cuando se selecciona un albaran de la lista
Dim i As Integer
Dim Cad As String
Dim TipoFP As Integer 'Forma de pago
Dim TipoDtoPP As Currency 'descuento pronto pago
Dim tipoDtoGn As Currency 'descuento general

    Screen.MousePointer = vbHourglass
    
    Set ListView1.SelectedItem = item
    
    'Inicializamos a cero
    TipoFP = 0
    TipoDtoPP = 0
    tipoDtoGn = 0
    
    'cuando seleccionamos un check vemos si lo podemos seleccionar
    'ya que si ya habia algun albaran selecionado tendremos que comprobar
    'que son de la misma forpa, dtoppago y dtognral.
    'si esto no se cumple no se pueden agrupar en la misma factura
    For i = 1 To ListView1.ListItems.Count
        If item.Index <> i Then
            If ListView1.ListItems(i).Checked Then
                'ya habia otro albaran seleccionado
                TipoFP = ListView1.ListItems(i).SubItems(2)
                TipoDtoPP = CCur(ListView1.ListItems(i).SubItems(4))
                tipoDtoGn = CCur(ListView1.ListItems(i).SubItems(5))
                Exit For
            End If
        End If
    Next i
    
    If Not (TipoFP = 0 And TipoDtoPP = 0 And tipoDtoGn = 0) Then
    'si ya habia un albaran seleccionado, comprobar que es del mismo tipo
        If item.SubItems(2) <> TipoFP Or item.SubItems(4) <> TipoDtoPP Or item.SubItems(5) <> tipoDtoGn Then
            MsgBox "Se debe seleccionar albaranes de la misma Forma de Pago y Descuentos", vbExclamation
            ListView1.SelectedItem.Checked = False
            Screen.MousePointer = vbDefault
            ListView1.SetFocus
            Exit Sub
        End If
    Else
    End If
    
    ' Calculamos los datos de factura
    If Not VerAlbaranes Then CalcularDatosFactura
    Screen.MousePointer = vbDefault
End Sub


Private Sub mnGenerarFac_Click()
Dim i As Integer

'    If Text1(16).Text = "" And Text1(17).Text = "" And Text1(18).Text = "" Then
'        If MsgBox("Se va a generar una factura a cero. ¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
'    End If


    FrameIntro.Enabled = False
    ListView1.Enabled = False
    FrameFactura.Enabled = True
    
    For i = 6 To 23
        BloquearTxt Text1(i), True
    Next i
    
    For i = 6 To 23
        Text1(i).Enabled = False
    Next i
    
    
    BloquearTxt Text1(10), (Text1(16).Text = "")
    BloquearTxt Text1(11), (Text1(17).Text = "")
    BloquearTxt Text1(12), (Text1(18).Text = "")
    
    imgBuscar(3).Enabled = (Text1(16).Text <> "")
    imgBuscar(4).Enabled = (Text1(17).Text <> "")
    imgBuscar(5).Enabled = (Text1(18).Text <> "")
    imgBuscar(3).visible = (Text1(16).Text <> "")
    imgBuscar(4).visible = (Text1(17).Text <> "")
    imgBuscar(5).visible = (Text1(18).Text <> "")
    
    Me.cmdCancelar.Enabled = True
    Me.cmdCancelar.visible = True
    Me.cmdGenerar.Enabled = True
    Me.cmdGenerar.visible = True
    
    PonerFoco Text1(10)
    
'    BotonFacturar
'    Set vProve = Nothing
'    Screen.MousePointer = vbDefault
End Sub

Private Sub mnPedirDatos_Click()
    BotonPedirDatos
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub


Private Sub mnVerAlbaran_Click()
    BotonVerAlbaranes
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)
        
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1, 2 'Fecha factura, fecha recepcion
            '[Monica]28/08/2013: controlamos que esté dentro de campaña
            PonerFormatoFecha Text1(Index), True
            If Text1(Index) <> "" Then
                ' No debe existir el número de factura para el proveedor en hco
                If ExisteFacturaEnHco Then
                    InicializarListView
                End If
            End If
            
        Case 3 'Cod Proveedor
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "proveedor", "nomprove", "codprove")
                
                Text1(5).Text = DevuelveDesdeBDNew(cAgro, "proveedor", "codbanpr", "codprove", Text1(3).Text, "N")
                If Text1(5).Text <> "" Then Text2(5).Text = DevuelveDesdeBDNew(cAgro, "banpropi", "nombanpr", "codbanpr", Text1(5).Text, "N")
                
                ' No debe existir el número de factura para el proveedor en hco
                If ExisteFacturaEnHco Then
                    InicializarListView
                Else
                    'comprobamos que no haya nadie recepcionando facturas de ese proveedor
                    DesBloqueoManual ("RECFAC")
                    If Not BloqueoManual("RECFAC", Text1(3).Text) Then
                        MsgBox "No se puede recepcionar factura de ese proveedor. Hay otro usuario recepcionando.", vbExclamation
                        BotonPedirDatos
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    Else
                        CargarAlbaranes
                    End If
                    
                End If
                
            Else
                Text2(Index).Text = ""
            End If
            
'--monica
'        Case 4 'Cod Trabajador
'            If PonerFormatoEntero(Text1(Index)) Then
'                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba")
'            Else
'                Text2(Index).Text = ""
'            End If

        Case 5 'Cta Prevista de PAgo
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "banpropi", "nombanpr", "codbanpr")
                Text1(Index).Text = Format(Text1(Index).Text, "0000")
            Else
                Text2(Index).Text = ""
            End If
            
        Case 10, 11, 12 ' codigo de iva
            If PonerFormatoEntero(Text1(Index)) Then
                Text1(Index + 3).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", Text1(Index).Text, "N")
            Else
                Text1(Index + 3).Text = ""
            End If
        
        
            RecalcularDatosFactura
            
    End Select
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte, Numreg As Byte
Dim b As Boolean
On Error GoTo EPonerModo

    Modo = Kmodo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
        
                 
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    'si estamos en modificar bloquea las compos que son clave primaria
    BloquearText1 Me, Modo
    
    'Importes siempre bloqueados
    For i = 6 To 22
        BloquearTxt Text1(i), True
    Next i
    'Campo B.Imp y Imp. IVA siempre en azul
    Text1(9).BackColor = &HFFFFC0 'Base imponible
    Text1(19).BackColor = &HFFFFC0 'Total Iva 1
    Text1(20).BackColor = &HFFFFC0 'Iva 2
    Text1(21).BackColor = &HFFFFC0 'IVa 3
    Text1(22).BackColor = &HC0C0FF    'Total factura
    
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2 And Modo <> 5)
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = b
    Next i
                    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
'    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
'Comprobar que los datos del frame de introduccion son correctos antes de cargar datos
Dim vtag As CTag
Dim Cad As String
Dim i As Byte

    On Error GoTo EDatosOK
    DatosOk = False
    
    ' deben de introducirse todos los datos del frame
    For i = 0 To 5
        If Text1(i).Text = "" And i <> 4 Then
            If Text1(i).Tag <> "" Then
                Set vtag = New CTag
                If vtag.Cargar(Text1(i)) Then
                    Cad = vtag.Nombre
                Else
                    Cad = "Campo"
                End If
                Set vtag = Nothing
            Else
                Cad = "Campo"
                If i = 5 Then Cad = "Cta. Prev. Pago"
            End If
            MsgBox Cad & " no puede estar vacio. Reintroduzca", vbExclamation
            PonerFoco Text1(i)
            Exit Function
        End If
    Next i
        
    'comprobar que la fecha de la factura sea anterior a la fecha de recepcion
    If Not EsFechaIgualPosterior(Text1(1).Text, Text1(2).Text, True, "La fecha de recepción debe ser igual o posterior a la fecha de la factura.") Then
        Exit Function
    End If
    
    'Comprobar que la fecha de RECEPCION esta dentro de los ejercicios contables
    '[Monica]20/06/2017: como David
    If vParamAplic.NumeroConta <> 0 Then
        ResultadoFechaContaOK = EsFechaOKConta(CDate(Text1(2).Text))
        If ResultadoFechaContaOK > 0 Then
            If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
            Exit Function
        End If
    End If
    
    
    'comprobar que se han seleccionado lineas para facturar
    If cadWHERE = "" Then
        MsgBox "Debe seleccionar albaranes para facturar.", vbExclamation
        Exit Function
    End If
    
    
    ' No debe existir el número de factura para el proveedor en hco
    If ExisteFacturaEnHco Then Exit Function
    
    
    'todos los albaranes seleccionados deben tener la misma: forma pago, dto ppago, dto gnral
    Cad = "select count(distinct codforpa,dtoppago,dtognral) from scaalp "
    Cad = Cad & " WHERE " & Replace(cadWHERE, "slialp", "scaalp")
    If RegistrosAListar(Cad) > 1 Then
        MsgBox "No se puede facturar albaranes con distintas: forma de pago, dto gral, dto ppago.", vbExclamation
        Exit Function
    End If
    
    
    'Si la forpa es TRANSFERENCIA entonces compruebo la si tiene cta bancaria
    Cad = "select distinct (codforpa) from scaalp "
    Cad = Cad & " WHERE " & Replace(cadWHERE, "slialp", "scaalp")
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = miRsAux.Fields(0)
    miRsAux.Close
    
    
    
    'Ahora buscamos el tipforpa del codforpa
    Cad = "Select tipoforp from forpago where codforpa=" & Cad
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    If miRsAux.EOF Then
        MsgBox "Error en el TIPO de forma de pago", vbExclamation
    Else
        i = 1
        Cad = miRsAux.Fields(0)
        If Val(Cad) = vbFPTransferencia Then
            'Compruebo que la forpa es transferencia
            i = 2
        End If
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    
    
    If i = 2 Then
        'La forma de pago es transferencia. Debo comprobar que existe la cuenta bancaria
        'del proveedor
        If vProve.CuentaBan = "" Or vProve.DigControl = "" Or vProve.Sucursal = "" Or vProve.Banco = "" Then
            Cad = "Cuenta bancaria incorrecta. Forma de pago: transferencia.    ¿Continuar?"
            If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then i = 0
        End If
    End If
    
    'Si i=0 es que o esta mal la forpa o no quiere seguir pq no tiene cuenta bancaria
    If i > 0 Then DatosOk = True
    Exit Function
    
EDatosOK:
    DatosOk = False
    MuestraError Err.Number, "Comprobar datos correctos", Err.Description
End Function



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Pedir datos
             mnPedirDatos_Click
             
        Case 2 'Ver Albaranes
            mnVerAlbaran_Click
            
        Case 3 'Generar Factura
            mnGenerarFac_Click

    End Select
End Sub


Private Sub PonerOpcionesMenu()
Dim J As Byte

    PonerOpcionesMenuGeneral Me
    
    J = Val(Me.mnPedirDatos.HelpContextID)
    If J < vUsu.Nivel Then Me.mnPedirDatos.Enabled = False
    
    J = Val(Me.mnGenerarFac.HelpContextID)
    If J < vUsu.Nivel Then Me.mnGenerarFac.Enabled = False
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

 
Private Sub BotonPedirDatos()
Dim Nombre As String


    FrameIntro.Enabled = True
    ListView1.Enabled = False
    FrameFactura.Enabled = False
    

    'Vaciamos todos los Text
    LimpiarCampos
    'Vaciamos el ListView
    InicializarListView
    
    'Como no habrá albaranes seleccionados vaciamos la cadwhere
    cadWHERE = ""
    
    PonerModo 3
    
    'fecha recepcion
    Text1(2).Text = Format(Now, "dd/mm/yyyy")
    
'--monica
'    'poner trabajador conectado como operador
'    Text1(4).Text = PonerTrabajadorConectado(Nombre)
'    Text2(4).Text = Nombre
    
    'desbloquear los registros de la saalp (si hay bloquedos)
    TerminaBloquear
    
    'si vamos
    'desBloqueo Manual de las tablas
'    DesBloqueoManual ("scaalp")
    
    PonerFoco Text1(0)
End Sub

Private Sub BotonModificar()
Dim Nombre As String

    PonerModo 4
    
    PonerFoco Text1(10)

End Sub




Private Sub BotonVerAlbaranes()

    If Not SeleccionaRegistros Then Exit Sub
    
    VerAlbaranes = True
    
    frmComEntAlbaranes.cadSelAlbaranes = cadWHERE
    frmComEntAlbaranes.EsHistorico = False
    frmComEntAlbaranes.Show vbModal
    frmComEntAlbaranes.cadSelAlbaranes = ""
End Sub
    


Private Sub CargarAlbaranes()
'Recupera de la BD y muestra en el Listview todos los albaranes de compra
'que tiene el proveedor introducido.
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
On Error GoTo ECargar

    ListView1.ListItems.Clear
    If VerAlbaranes = False Then cadWHERE = ""
    
    'si no hay proveedor salir
    If Text1(3).Text = "" Then Exit Sub
    
    SQL = "SELECT scaalp.numalbar,scaalp.fechaalb,scaalp.codforpa,forpago.nomforpa,scaalp.dtoppago,scaalp.dtognral, "
    SQL = SQL & " sum(slialp.importel) as bruto "
    SQL = SQL & " FROM (scaalp LEFT OUTER JOIN forpago ON scaalp.codforpa=forpago.codforpa) "
    SQL = SQL & " INNER JOIN slialp ON scaalp.numalbar = slialp.numalbar  AND scaalp.fechaalb=slialp.fechaalb AND scaalp.codprove=slialp.codprove "
    SQL = SQL & " WHERE scaalp.codprove =" & Text1(3).Text
    SQL = SQL & " GROUP BY scaalp.numalbar, scaalp.fechaalb, scaalp.codforpa, scaalp.dtoppago,scaalp.dtognral "
    SQL = SQL & " ORDER BY scaalp.numalbar"

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    InicializarListView
    
    While Not Rs.EOF
        Set ItmX = ListView1.ListItems.Add()
        ItmX.Text = Rs!NumAlbar
        ItmX.SubItems(1) = Format(Rs!FechaAlb, "dd/mm/yyyy")
        ItmX.SubItems(2) = Format(Rs!Codforpa, "000")
        ItmX.SubItems(3) = Rs!nomforpa
        ItmX.SubItems(4) = Format(Rs!DtoPPago, "#0.00")
        ItmX.SubItems(5) = Format(Rs!DtoGnral, "#0.00")
        ItmX.SubItems(6) = Format(Rs!Bruto, "#,###,#0.00") '(RAFA/ALZIRA) 12092006
        'Sig
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    ListView1.Enabled = True

ECargar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando Albaranes", Err.Description
End Sub


Private Sub InicializarListView()
'Inicializa las columnas del List view

    ListView1.ListItems.Clear
    
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "NºAlbaran", 1200
    ListView1.ColumnHeaders.Add , , "Fecha", 1350, 2
    ListView1.ColumnHeaders.Add , , "FPag", 750
    ListView1.ColumnHeaders.Add , , "Descripción", 1500
    ListView1.ColumnHeaders.Add , , "DtoPP", 725, 2
    ListView1.ColumnHeaders.Add , , "DtoGr", 725, 2
    ListView1.ColumnHeaders.Add , , "Imp.Bruto", 1450, 1
End Sub



Private Sub CalcularDatosFactura()
Dim i As Integer
Dim SQL As String
Dim cadAux As String
Dim vFactu As CFacturaCom

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For i = 6 To 22
         Text1(i).Text = ""
    Next i
    
    cadAux = ""
    cadWHERE = ""
    
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then
        'para cada albaran seleccionado para la factura
            Forpa = ListView1.ListItems(i).SubItems(2)
            dtoPP = ListView1.ListItems(i).SubItems(4)
            dtoGn = ListView1.ListItems(i).SubItems(5)
            SQL = "(numalbar=" & DBSet(ListView1.ListItems(i).Text, "T") & " and "
            SQL = SQL & "fechaalb=" & DBSet(ListView1.ListItems(i).SubItems(1), "F") & ")"
            If cadAux = "" Then
                cadAux = SQL
            Else
                cadAux = cadAux & " OR " & SQL
            End If
        End If
    Next i
    
    If cadAux <> "" Then
    'se han seleccionado albaranes para facturar
    'Esta el la cadena WHERE de los albaranes seleccionados para obtener
    'el bruto de las lineas de los albaranes agrupadas por tipo de iva
        cadWHERE = "slialp.codprove=" & Val(Text1(3).Text)
        cadWHERE = cadWHERE & " AND (" & cadAux & ")"
    Else
        Exit Sub
    End If
    
    
    If Not SeleccionaRegistros Then Exit Sub
    
    If Not BloqueaRegistro("scaalp", cadWHERE) Then
        ListView1.SelectedItem.Checked = False
    End If
    
    Set vFactu = New CFacturaCom
    vFactu.DtoPPago = dtoPP
    vFactu.DtoGnral = dtoGn
    If vFactu.CalcularDatosFactura(cadWHERE, "scaalp", "slialp") Then
        Text1(6).Text = vFactu.BrutoFac
        Text1(7).Text = vFactu.ImpPPago
        Text1(8).Text = vFactu.ImpGnral
        Text1(9).Text = vFactu.BaseImp
        Text1(10).Text = vFactu.TipoIVA1
        Text1(11).Text = vFactu.TipoIVA2
        Text1(12).Text = vFactu.TipoIVA3
        Text1(13).Text = vFactu.PorceIVA1
        Text1(14).Text = vFactu.PorceIVA2
        Text1(15).Text = vFactu.PorceIVA3
        Text1(16).Text = vFactu.BaseIVA1
        Text1(17).Text = vFactu.BaseIVA2
        Text1(18).Text = vFactu.BaseIVA3
        Text1(19).Text = vFactu.ImpIVA1
        Text1(20).Text = vFactu.ImpIVA2
        Text1(21).Text = vFactu.ImpIVA3
        Text1(22).Text = vFactu.TotalFac
        
        For i = 6 To 22
            FormateaCampo Text1(i)
        Next i
        'Quitar ceros de linea IVA 2
        If Val(Text1(14).Text) = 0 And Val(Text1(11).Text) = 0 Then
            For i = 11 To 20 Step 3
                Text1(i).Text = QuitarCero(CCur(Text1(i).Text))
            Next i
        End If
        'Quitar ceros de linea IVA 3
        If Val(Text1(15).Text) = 0 And Val(Text1(12).Text) = 0 Then
            For i = 12 To 21 Step 3
                Text1(i).Text = QuitarCero(CCur(Text1(i).Text))
            Next i
        End If
        
    Else
        MuestraError Err.Number, "Calculando Factura", Err.Description
    End If
    Set vFactu = Nothing
   
End Sub


Private Sub RecalcularDatosFactura()
Dim i As Integer
Dim SQL As String
Dim cadAux As String
Dim TotalFactura As Currency
    
Dim ImpBImIVA As Currency
Dim impiva As Currency
Dim ImpIVA1 As Currency
Dim ImpIVA2 As Currency
Dim ImpIVA3 As Currency
    
    On Error GoTo eRecalcularDatosFactura
    
    If Not SeleccionaRegistros Then Exit Sub
    
    If Not BloqueaRegistro("scaalp", cadWHERE) Then
        ListView1.SelectedItem.Checked = False
    End If
    
    TotalFactura = 0
    If Text1(16).Text <> "" Then
        cadAux = Text1(13).Text
        ImpBImIVA = CCur(ImporteSinFormato(Text1(16).Text))
        If cadAux = "" Then cadAux = "0"
        impiva = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
        
        ImpIVA1 = impiva
        
        
        'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
        'los vamos acumulando
        TotalFactura = TotalFactura + ImpBImIVA + impiva
    End If
    
    If Text1(17).Text <> "" Then
        cadAux = Text1(14).Text
        ImpBImIVA = CCur(ImporteSinFormato(Text1(17).Text))
        If cadAux = "" Then cadAux = "0"
        impiva = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
        
        ImpIVA2 = impiva
        
        
        'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
        'los vamos acumulando
        TotalFactura = TotalFactura + ImpBImIVA + impiva
    End If
    
    
    If Text1(18).Text <> "" Then
        cadAux = Text1(15).Text
        ImpBImIVA = CCur(ImporteSinFormato(Text1(18).Text))
        If cadAux = "" Then cadAux = "0"
        impiva = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
        
        ImpIVA3 = impiva
        
        
        'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
        'los vamos acumulando
        TotalFactura = TotalFactura + ImpBImIVA + impiva
    End If
        
'        Text1(6).Text = vFactu.BrutoFac
'        Text1(7).Text = vFactu.ImpPPago
'        Text1(8).Text = vFactu.ImpGnral
'        Text1(9).Text = vFactu.BaseImp
'        Text1(10).Text = vFactu.TipoIVA1
'        Text1(11).Text = vFactu.TipoIVA2
'        Text1(12).Text = vFactu.TipoIVA3
'        Text1(13).Text = vFactu.PorceIVA1
'        Text1(14).Text = vFactu.PorceIVA2
'        Text1(15).Text = vFactu.PorceIVA3
'        Text1(16).Text = vFactu.BaseIVA1
'        Text1(17).Text = vFactu.BaseIVA2
'        Text1(18).Text = vFactu.BaseIVA3
        
        Text1(19).Text = ImpIVA1
        Text1(20).Text = ImpIVA2
        Text1(21).Text = ImpIVA3
        Text1(22).Text = TotalFactura
        
        For i = 19 To 22
            FormateaCampo Text1(i)
        Next i
        'Quitar ceros de linea IVA 2
        If Val(Text1(14).Text) = 0 And Val(Text1(11).Text) = 0 Then
            For i = 11 To 20 Step 3
                Text1(i).Text = ""
            Next i
        End If
        'Quitar ceros de linea IVA 3
        If Val(Text1(15).Text) = 0 And Val(Text1(12).Text) = 0 Then
            For i = 12 To 21 Step 3
                Text1(i).Text = ""
            Next i
        End If
        Exit Sub
        
   
eRecalcularDatosFactura:
    MuestraError Err.Number, "Recalculando Datos de Factura", Err.Description
End Sub





Private Function SeleccionaRegistros() As Boolean
'Comprueba que se seleccionan albaranes en la base de datos
'es decir que hay albaranes marcados
'cuando se van marcando albaranes se van añadiendo el la cadena cadWhere
Dim SQL As String

    On Error GoTo ESel
    SeleccionaRegistros = False
    
    If cadWHERE = "" Then Exit Function
    cadWHERE = Replace(cadWHERE, "slialp", "scaalp")
    
    SQL = "Select count(*) FROM scaalp"
    SQL = SQL & " WHERE " & cadWHERE
    If RegistrosAListar(SQL) <> 0 Then SeleccionaRegistros = True
    Exit Function
    
ESel:
    SeleccionaRegistros = False
    MuestraError Err.Number, "No hay seleccionados Albaranes", Err.Description
End Function


Private Sub BotonFacturar()
Dim vFactu As CFacturaCom
Dim Cad As String
Dim i As Integer


    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    ' preguntamos antes de recepcionar
    If MsgBox("¿ Desea generar la Factura de Proveedor ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    Cad = ""
    If Text1(3).Text = "" Then
        Cad = "Falta proveedor"
    Else
        If Not IsNumeric(Text1(3).Text) Then Cad = "Campo proveedor debe ser numérico"
    End If
    If Cad <> "" Then
        MsgBox Cad, vbExclamation
        Exit Sub
    End If
        
        
        
    Set vProve = New CProveedor
    
    'Tiene que ller los datos del proveedor
    If Not vProve.LeerDatos(Text1(3).Text) Then Exit Sub
    
    
    If Not DatosOk Then
        Exit Sub
    End If
    
        'Pasar los Albaranes seleccionados con cadWHERE a una factura
        Set vFactu = New CFacturaCom
        vFactu.Proveedor = Text1(3).Text
        vFactu.NumFactu = Text1(0).Text
        vFactu.FecFactu = Text1(1).Text
        vFactu.FecRecep = Text1(2).Text
        vFactu.Trabajador = Text1(4).Text
        vFactu.BancoPr = Text1(5).Text
        vFactu.BrutoFac = ImporteFormateado(Text1(6).Text)
        vFactu.ForPago = Forpa
        vFactu.DtoPPago = dtoPP
        vFactu.DtoGnral = dtoGn
        vFactu.ImpPPago = ImporteFormateado(Text1(7).Text)
        vFactu.ImpGnral = ImporteFormateado(Text1(8).Text)
        vFactu.BaseIVA1 = ImporteFormateado(Text1(16).Text)
        vFactu.BaseIVA2 = ImporteFormateado(Text1(17).Text)
        vFactu.BaseIVA3 = ImporteFormateado(Text1(18).Text)
        vFactu.TipoIVA1 = ComprobarCero(Text1(10).Text)
        vFactu.TipoIVA2 = ComprobarCero(Text1(11).Text)
        vFactu.TipoIVA3 = ComprobarCero(Text1(12).Text)
        vFactu.PorceIVA1 = ComprobarCero(Text1(13).Text)
        vFactu.PorceIVA2 = ComprobarCero(Text1(14).Text)
        vFactu.PorceIVA3 = ComprobarCero(Text1(15).Text)
        vFactu.ImpIVA1 = ImporteFormateado(Text1(19).Text)
        vFactu.ImpIVA2 = ImporteFormateado(Text1(20).Text)
        vFactu.ImpIVA3 = ImporteFormateado(Text1(21).Text)
        vFactu.TotalFac = ImporteFormateado(Text1(22).Text)
        vFactu.Concepto = Text1(23).Text
        
        
        'Si el proveedor tiene CTA BANCARIA se la asigno
        vFactu.CCC_Entidad = vProve.Banco
        vFactu.CCC_Oficina = vProve.Sucursal
        vFactu.CCC_CC = vProve.DigControl
        vFactu.CCC_CTa = vProve.CuentaBan
        '[Monica]22/11/2013: si el proveedor tiene cuenta bancaria se la asigno
        vFactu.CCC_Iban = vProve.Iban
        
        If vFactu.TraspasoAlbaranesAFactura(cadWHERE) Then BotonPedirDatos
        Set vFactu = Nothing
    
    
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Function ExisteFacturaEnHco() As Boolean
'Comprobamos si la factura ya existe en la tabla de Facturas a Proveedor: scafpc
Dim Cad As String

    ExisteFacturaEnHco = False
    'Tiene que tener valor los 3 campos de clave primaria antes de comprobar
    If Not (Text1(0).Text <> "" And Text1(1).Text <> "" And Text1(3).Text <> "") Then Exit Function
    
    ' No debe existir el número de factura para el proveedor en hco
    Cad = "SELECT count(*) FROM scafpc "
    Cad = Cad & " WHERE codprove=" & Text1(3).Text & " AND numfactu=" & DBSet(Text1(0).Text, "T") & " AND year(fecfactu)=" & Year(Text1(1).Text)
    If RegistrosAListar(Cad) > 0 Then
        MsgBox "Factura de proveedor ya existente. Reintroduzca.", vbExclamation
        ExisteFacturaEnHco = True
        Exit Function
    End If
End Function

Private Sub RefrescarAlbaranes()
Dim i As Integer
Dim SQL As String
Dim Itm As ListItem
Dim Rs As ADODB.Recordset
    

    For i = 1 To ListView1.ListItems.Count
        SQL = "SELECT scaalp.numalbar,scaalp.fechaalb,scaalp.codforpa,forpago.nomforpa,scaalp.dtoppago,scaalp.dtognral, "
        SQL = SQL & " sum(slialp.importel) as bruto "
        SQL = SQL & " FROM (scaalp LEFT OUTER JOIN forpago ON scaalp.codforpa=forpago.codforpa) "
        SQL = SQL & " INNER JOIN slialp ON scaalp.numalbar = slialp.numalbar  AND scaalp.fechaalb=slialp.fechaalb AND scaalp.codprove=slialp.codprove "
        SQL = SQL & " WHERE scaalp.codprove =" & Text1(3).Text & " AND scaalp.numalbar=" & DBSet(ListView1.ListItems(i).Text, "T") & " AND scaalp.fechaalb=" & DBSet(ListView1.ListItems(i).SubItems(1), "F")
        SQL = SQL & " GROUP BY scaalp.numalbar, scaalp.fechaalb, scaalp.codforpa, scaalp.dtoppago,scaalp.dtognral "
        SQL = SQL & " ORDER BY scaalp.numalbar"

        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

        If Not Rs.EOF Then 'Actualizamos los datos de este item en el list
            ListView1.ListItems(i).SubItems(2) = Rs!Codforpa
            ListView1.ListItems(i).SubItems(3) = Rs!nomforpa
            ListView1.ListItems(i).SubItems(4) = Rs!DtoPPago
            ListView1.ListItems(i).SubItems(5) = Rs!DtoGnral
            ListView1.ListItems(i).SubItems(6) = Rs!Bruto

        End If
        
        If ListView1.ListItems(i).Checked Then 'comprobamos otra vez el chek y recalculamos factura
            Set Itm = ListView1.ListItems(i)
            ListView1_ItemCheck Itm
        End If

        Rs.Close
        Set Rs = Nothing
    Next i
    
    'recalcular el total de la factura
     For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then
            CalcularDatosFactura
            Exit For
        End If
     Next i
End Sub






