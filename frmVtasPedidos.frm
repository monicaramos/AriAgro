VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmVtasPedidos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedidos de Clientes"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   16680
   Icon            =   "frmVtasPedidos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   16680
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
      Left            =   13995
      TabIndex        =   96
      Top             =   270
      Width           =   1605
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   5745
      TabIndex        =   94
      Top             =   45
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   95
         Top             =   180
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Primero"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Anterior"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Siguiente"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Último"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3795
      TabIndex        =   92
      Top             =   45
      Width           =   1875
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   210
         TabIndex        =   93
         Top             =   180
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Impresión Proveedor"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Orden de Carga"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Generar Albarán"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   135
      TabIndex        =   90
      Top             =   45
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   91
         Top             =   180
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
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
               Object.Width           =   1e-4
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Buscar"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver Todos"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Imprimir"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Salir"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2745
      Left            =   135
      TabIndex        =   23
      Top             =   810
      Width           =   16385
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
         Index           =   16
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   75
         Text            =   "Text2"
         Top             =   2295
         Width           =   4800
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
         Left            =   1575
         MaxLength       =   6
         TabIndex        =   8
         Tag             =   "Cod.Almacen|N|N|0|999|pedidos|codalmac|000||"
         Text            =   "Text1"
         Top             =   2295
         Width           =   890
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Impreso"
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
         Left            =   14130
         TabIndex        =   71
         Tag             =   "Impresor|N|N|||pedidos|impresor|0||"
         Top             =   1620
         Width           =   1095
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
         Left            =   9495
         MaxLength       =   4
         TabIndex        =   15
         Tag             =   "Nro.Acta|N|S|||pedidos|nroactas|##0||"
         Text            =   "Text3"
         Top             =   1170
         Width           =   1140
      End
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
         Index           =   12
         Left            =   7560
         MaxLength       =   15
         TabIndex        =   14
         Tag             =   "Nro.Contrato|T|S|||pedidos|nrocontra|||"
         Text            =   "123456789012345"
         Top             =   1170
         Width           =   1905
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
         Left            =   11250
         MaxLength       =   3
         TabIndex        =   16
         Tag             =   "Nro.Palets|N|S|||pedidos|totpalet|##0||"
         Text            =   "Text3"
         Top             =   1170
         Width           =   1140
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
         Left            =   1575
         MaxLength       =   6
         TabIndex        =   7
         Tag             =   "Cod.Agencia|N|N|0|999|pedidos|codtrans|000||"
         Text            =   "Text1"
         Top             =   1935
         Width           =   890
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
         Index           =   6
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   62
         Text            =   "Text2"
         Top             =   1935
         Width           =   4800
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
         Left            =   1575
         MaxLength       =   6
         TabIndex        =   6
         Tag             =   "Tipo Mercado|N|N|0|999|pedidos|codtimer|000||"
         Text            =   "Text1"
         Top             =   1575
         Width           =   890
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
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   60
         Text            =   "Text2"
         Top             =   1575
         Width           =   4800
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
         Left            =   12555
         MaxLength       =   10
         TabIndex        =   12
         Tag             =   "Fecha Albarán|F|S|||pedidos|fechaalb|dd/mm/yyyy||"
         Top             =   450
         Width           =   1450
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
         Index           =   4
         Left            =   1575
         MaxLength       =   6
         TabIndex        =   5
         Tag             =   "Cod.Destino|N|N|0|9999|pedidos|coddesti|0000||"
         Text            =   "Text1"
         Top             =   1215
         Width           =   890
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
         Index           =   4
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   58
         Text            =   "Text2"
         Top             =   1215
         Width           =   4800
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
         Index           =   2
         Left            =   2835
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Carga|F|N|||pedidos|fechacar|dd/mm/yyyy||"
         Top             =   450
         Width           =   1450
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
         Left            =   9495
         MaxLength       =   40
         TabIndex        =   10
         Tag             =   "Matricula Remolque|T|S|||pedidos|matrirem|||"
         Text            =   "Text3"
         Top             =   450
         Width           =   1575
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
         Index           =   1
         Left            =   1305
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Pedido|F|N|||pedidos|fechaped|dd/mm/yyyy||"
         Top             =   450
         Width           =   1450
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
         Left            =   7560
         MaxLength       =   12
         TabIndex        =   9
         Tag             =   "Matricula Vehiculo|T|S|||pedidos|matriveh|||"
         Top             =   450
         Width           =   1575
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
         Index           =   9
         Left            =   11250
         MaxLength       =   7
         TabIndex        =   11
         Tag             =   "Nº Albaran|N|S|||pedidos|numalbar|0000000||"
         Text            =   "Text3"
         Top             =   450
         Width           =   1140
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
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   39
         Text            =   "Text2"
         Top             =   855
         Width           =   4800
      End
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
         Height          =   645
         Index           =   15
         Left            =   7560
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Tag             =   "Observaciones|T|S|||pedidos|observac|||"
         Top             =   1980
         Width           =   8085
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
         Left            =   5445
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "Situacion|N|N|||pedidos|situacio|0||"
         Top             =   450
         Width           =   1845
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
         Left            =   1590
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "Cod. Cliente|N|N|0|999999|pedidos|codclien|000000||"
         Text            =   "Text1"
         Top             =   855
         Width           =   890
      End
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
         Index           =   14
         Left            =   14085
         MaxLength       =   12
         TabIndex        =   13
         Tag             =   "Referencia Cl|T|S|||pedidos|refclien|||"
         Text            =   "Text3"
         Top             =   450
         Width           =   1545
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
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
         Left            =   225
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Nº Pedido|N|S|||pedidos|numpedid|0000000|S|"
         Text            =   "Text1 7"
         Top             =   450
         Width           =   980
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Pedido"
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
         Left            =   225
         TabIndex        =   86
         Top             =   180
         Width           =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "Situación"
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
         Index           =   27
         Left            =   5445
         TabIndex        =   85
         Top             =   180
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Albarán"
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
         Left            =   11250
         TabIndex        =   84
         Top             =   225
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "Mat.Vehículo"
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
         Left            =   7560
         TabIndex        =   83
         Top             =   225
         Width           =   1410
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Ped"
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
         Left            =   1305
         TabIndex        =   82
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Mat.Remolque"
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
         Left            =   9495
         TabIndex        =   81
         Top             =   225
         Width           =   1500
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   3
         Left            =   2835
         TabIndex        =   80
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "F.Albarán"
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
         Left            =   12555
         TabIndex        =   79
         Top             =   225
         Width           =   1140
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   9
         Left            =   225
         TabIndex        =   76
         Top             =   2340
         Width           =   810
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1305
         ToolTipText     =   "Buscar Agencia"
         Top             =   2340
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Nro.Acta"
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
         Index           =   18
         Left            =   9495
         TabIndex        =   66
         Top             =   900
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "Nro.Contrato"
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
         Index           =   17
         Left            =   7560
         TabIndex        =   65
         Top             =   900
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Nro.Palets"
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
         Index           =   16
         Left            =   11250
         TabIndex        =   64
         Top             =   900
         Width           =   1155
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1305
         ToolTipText     =   "Buscar Agencia"
         Top             =   1980
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Agencia "
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
         Left            =   225
         TabIndex        =   63
         Top             =   1980
         Width           =   810
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1305
         ToolTipText     =   "Buscar T.Mercado"
         Top             =   1620
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "T.Mercado"
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
         Left            =   225
         TabIndex        =   61
         Top             =   1620
         Width           =   1080
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   13725
         Picture         =   "frmVtasPedidos.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1305
         ToolTipText     =   "Buscar Destino"
         Top             =   1260
         Width           =   240
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   8
         Left            =   225
         TabIndex        =   59
         Top             =   1260
         Width           =   765
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   4185
         Picture         =   "frmVtasPedidos.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   2475
         Picture         =   "frmVtasPedidos.frx":0122
         ToolTipText     =   "Buscar fecha"
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   9135
         ToolTipText     =   "Zoom descripción"
         Top             =   1710
         Width           =   240
      End
      Begin VB.Label Label29 
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
         Height          =   255
         Left            =   7560
         TabIndex        =   38
         Top             =   1710
         Width           =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "Refer.Cliente"
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
         Left            =   14085
         TabIndex        =   37
         Top             =   180
         Width           =   1425
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   0
         Left            =   225
         TabIndex        =   24
         Top             =   900
         Width           =   810
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1305
         ToolTipText     =   "Buscar Cliente"
         Top             =   900
         Width           =   240
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   17
      Left            =   4695
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   88
      Text            =   "Text2"
      Top             =   1290
      Width           =   2640
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   17
      Left            =   3840
      MaxLength       =   6
      TabIndex        =   87
      Text            =   "Text1"
      Top             =   1290
      Width           =   780
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   8
      Left            =   12015
      MaxLength       =   30
      TabIndex        =   78
      Tag             =   "Unidades|N|N|0||pedidos_calibre|unidades|#,##0||"
      Text            =   "unidades"
      Top             =   6885
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtAux3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   16
      Left            =   9135
      MaxLength       =   30
      TabIndex        =   77
      Tag             =   "Unidades|N|S|||pedidos_variedad|unidades|#,##0|N|"
      Text            =   "unidades"
      Top             =   4905
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   12645
      MaxLength       =   30
      TabIndex        =   70
      Tag             =   "Peso Neto|N|N|0||pedidos_calibre|pesoneto|###,##0||"
      Text            =   "pesoneto"
      Top             =   6885
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtAux3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   15
      Left            =   10530
      MaxLength       =   30
      TabIndex        =   69
      Tag             =   "Prec.Profes.|N|S|||pedidos_variedad|preciopro|#0.0000|N|"
      Text            =   "precio prof"
      Top             =   4905
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAux3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   14
      Left            =   9810
      MaxLength       =   30
      TabIndex        =   68
      Tag             =   "Total Palets|N|S|||pedidos_variedad|totpalet|##0|N|"
      Text            =   "tot.palet"
      Top             =   4905
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.TextBox txtAux3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   13
      Left            =   8460
      MaxLength       =   30
      TabIndex        =   67
      Tag             =   "Num.Cajas|N|S|||pedidos_variedad|numcajas|#,##0|N|"
      Text            =   "num.caj"
      Top             =   4905
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtAux3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   12
      Left            =   5850
      MaxLength       =   30
      TabIndex        =   50
      Text            =   "nom forf"
      Top             =   4905
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.TextBox txtAux3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   11
      Left            =   4590
      MaxLength       =   30
      TabIndex        =   49
      Text            =   "nom marca"
      Top             =   4905
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAux3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   6
      Left            =   4365
      MaxLength       =   30
      TabIndex        =   48
      Tag             =   "Marca|N|N|||pedidos_variedad|codmarca|000||"
      Text            =   "marca"
      Top             =   4905
      Visible         =   0   'False
      Width           =   405
   End
   Begin MSComctlLib.Toolbar ToolAux 
      Height          =   390
      Index           =   0
      Left            =   135
      TabIndex        =   47
      Top             =   3690
      Width           =   1110
      _ExtentX        =   1958
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
   End
   Begin VB.TextBox txtAux3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   10
      Left            =   7695
      MaxLength       =   30
      TabIndex        =   46
      Tag             =   "Peso Neto|N|S|||pedidos_variedad|pesoneto|###,##0|N|"
      Text            =   "peso neto"
      Top             =   4905
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.TextBox txtAux3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   9
      Left            =   6930
      MaxLength       =   30
      TabIndex        =   45
      Tag             =   "Peso Bruto|N|N|||pedidos_variedad|pesobrut|###,##0||"
      Text            =   "peso bruto"
      Top             =   4905
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAux3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   8
      Left            =   6120
      MaxLength       =   30
      TabIndex        =   44
      Tag             =   "Categoria|T|S|||pedidos_variedad|categori|||"
      Text            =   "categ"
      Top             =   4905
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAux3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   5310
      MaxLength       =   30
      TabIndex        =   43
      Tag             =   "Forfait|T|N|||pedidos_variedad|codforfait|||"
      Text            =   "forfait"
      Top             =   4905
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.TextBox txtAux3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   3600
      MaxLength       =   30
      TabIndex        =   42
      Text            =   "nom.var.comer"
      Top             =   4905
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAux3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   2835
      MaxLength       =   30
      TabIndex        =   41
      Tag             =   "Variedad Comercial|N|N|||pedidos_variedad|codvarco|||"
      Text            =   "var.comer."
      Top             =   4905
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAux3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   2160
      MaxLength       =   30
      TabIndex        =   40
      Text            =   "nomvarie"
      Top             =   4905
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   11415
      MaxLength       =   30
      TabIndex        =   34
      Tag             =   "Num.Cajas|N|N|0||pedidos_calibre|numcajas|#,##0||"
      Text            =   "numcajas"
      Top             =   6885
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   10665
      MaxLength       =   5
      TabIndex        =   33
      Text            =   "nomca"
      Top             =   6885
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   6060
      MaxLength       =   12
      TabIndex        =   32
      Tag             =   "Num.Palet|N|N|||pedidos_calibre|numpalet||S|"
      Text            =   "numpedid"
      Top             =   6885
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   7020
      MaxLength       =   12
      TabIndex        =   31
      Tag             =   "Num.Linea|N|N|||pedidos_calibre|numlinea|00|N|"
      Text            =   "numlinea"
      Top             =   6885
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   7830
      MaxLength       =   12
      TabIndex        =   30
      Tag             =   "Num.Linea 1|N|N|||pedidos_calibre|numline1||N|"
      Text            =   "numline1"
      Top             =   6885
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   8910
      MaxLength       =   12
      TabIndex        =   29
      Tag             =   "Variedad|N|N|||pedidos_calibre|codvarie|000000|N|"
      Text            =   "variedad"
      Top             =   6885
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   9975
      MaxLength       =   5
      TabIndex        =   28
      Tag             =   "Calibre|N|N|||pedidos_calibre|codcalib|00|N|"
      Text            =   "calib"
      Top             =   6885
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtAux3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   450
      MaxLength       =   7
      TabIndex        =   27
      Tag             =   "Num.Pedido|N|N|||pedidos_variedad|numpedid||S|"
      Text            =   "numpedi"
      Top             =   4905
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtAux3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   810
      MaxLength       =   15
      TabIndex        =   26
      Tag             =   "Num.Linea|N|N|||pedidos_variedad|numlinea|00|S|"
      Text            =   "numlinea"
      Top             =   4905
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAux3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   1485
      MaxLength       =   30
      TabIndex        =   25
      Tag             =   "Variedad|N|N|||pedidos_variedad|codvarie||N|"
      Text            =   "variedad"
      Top             =   4905
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   165
      TabIndex        =   21
      Top             =   8055
      Width           =   2175
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   22
         Top             =   135
         Width           =   1755
      End
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
      Left            =   15435
      TabIndex        =   19
      Top             =   8145
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
      Left            =   14265
      TabIndex        =   18
      Top             =   8160
      Width           =   1065
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
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
      Left            =   15435
      TabIndex        =   20
      Top             =   8145
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data3 
      Height          =   330
      Left            =   3000
      Top             =   1080
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   945
      Top             =   7425
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
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   270
      Top             =   7425
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmVtasPedidos.frx":01AD
      Height          =   2055
      Left            =   135
      TabIndex        =   36
      Top             =   4140
      Width           =   16350
      _ExtentX        =   28840
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmVtasPedidos.frx":01C2
      Height          =   1680
      Left            =   5940
      TabIndex        =   35
      Top             =   6345
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   2963
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   16035
      TabIndex        =   97
      Top             =   210
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
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3300
      Left            =   7695
      TabIndex        =   51
      Top             =   4095
      Visible         =   0   'False
      Width           =   3930
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
         Index           =   1
         Left            =   225
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   56
         Text            =   "Text2"
         Top             =   1260
         Width           =   3360
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
         Index           =   7
         Left            =   225
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   72
         Text            =   "Text2"
         Top             =   2655
         Width           =   3360
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
         Index           =   2
         Left            =   225
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   55
         Text            =   "Text2"
         Top             =   1980
         Width           =   3360
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
         Index           =   0
         Left            =   225
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   53
         Text            =   "Text2"
         Top             =   585
         Width           =   3360
      End
      Begin VB.Label Label1 
         Caption         =   "Categoria"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   225
         TabIndex        =   73
         Top             =   2430
         Width           =   1035
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   225
         TabIndex        =   57
         Top             =   1710
         Width           =   1500
      End
      Begin VB.Label Label1 
         Caption         =   "Variedad Comercial"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   12
         Left            =   225
         TabIndex        =   54
         Top             =   315
         Width           =   1530
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   1
         Left            =   225
         TabIndex        =   52
         Top             =   990
         Width           =   945
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo Palet"
      Height          =   255
      Index           =   10
      Left            =   3840
      TabIndex        =   89
      Top             =   1020
      Width           =   765
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   5
      Left            =   4650
      ToolTipText     =   "Buscar Cliente"
      Top             =   1020
      Width           =   240
   End
   Begin VB.Label Label3 
      Caption         =   "123456789012345"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   645
      Left            =   540
      TabIndex        =   74
      Top             =   6885
      Width           =   5190
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
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
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         HelpContextID   =   2
         Shortcut        =   ^I
         Visible         =   0   'False
      End
      Begin VB.Menu mnImprimirProv 
         Caption         =   "&Impresión Proveedor"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnOrdenCarga 
         Caption         =   "&Orden de Carga"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnCMR 
         Caption         =   "&CMR"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnGenerarAlb 
         Caption         =   "&Generar Albarán"
         Shortcut        =   ^G
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
Attribute VB_Name = "frmVtasPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'Si se llama de la busqueda en el frmAlmMovimArticulos se accede
'a las tablas del Albaran o de Facturas de movimiento seleccionado (solo consulta)
Public hcoCodMovim As String 'cod. movim
Public hcoCodTipoM As String 'Codigo detalle de Movimiento(ALC)
Public hcoFechaMov As String 'fecha del movimiento

'========== VBLES PRIVADAS ====================
Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmLPed As frmVtasLinPedidos 'Lineas de variedades de pedidos
Attribute frmLPed.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmMen As frmMensajes  'Mensajes: palets asociados al pedido
Attribute frmMen.VB_VarHelpID = -1

Private WithEvents frmCli As frmClientes 'Form Mto de Clientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmTra As frmManAgencias 'Form Mto de Agencias de Transporte
Attribute frmTra.VB_VarHelpID = -1
Private WithEvents frmMer As frmManTipMerc 'Form Mto de Tipos de Mercado
Attribute frmMer.VB_VarHelpID = -1
Private WithEvents frmDest As frmDestCli 'Form Mto de destinos de clientes
Attribute frmDest.VB_VarHelpID = -1
Private WithEvents frmList As frmListadoPed 'Form listado de pedidos
Attribute frmList.VB_VarHelpID = -1
Private WithEvents frmAlm As frmManAlmProp 'Form mto de almacenes propios
Attribute frmAlm.VB_VarHelpID = -1
Private WithEvents frmPal As frmManPaleConf 'Palets de confreccion
Attribute frmPal.VB_VarHelpID = -1

Private WithEvents frmOrden As frmVtasOrdenCarga 'Orden de carga
Attribute frmOrden.VB_VarHelpID = -1
Private WithEvents frmPed As frmBasico2
Attribute frmPed.VB_VarHelpID = -1

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

Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim CodTipoMov As String
'Codigo tipo de movimiento en función del valor en la tabla de parámetros: stipom

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean

Dim EsCabecera As Boolean
'Para saber en MandaBusquedaPrevia si busca en la tabla scapla o en la tabla sdirec


Dim EsDeVarios As Boolean
'Si el cliente mostrado es de Varios o No

'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1


Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos
Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal
Dim indice As Byte
' utilizado para buscar por checks
Private BuscaChekc As String

Dim AlbDePalet As Boolean 'Si se va a generar el Pedido partiendo del pedido o del palet
Dim Continuar As Boolean
Dim FechaAlb As String 'Para cuando vuelve de pedir datos para Generar Albaran, saber la fecha que se introdujo
Private CadenaSQL As String 'Para crear consulta de Generar Albaran a partir del Pedido
Dim ImprimeAlb As Boolean 'Para saber cuando vuelve de Generar ALbaran si se ha solicitado Imprimir Albaran o no
Dim Incidencia As String

Private Sub Check1_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "check1(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "check1(" & Index & ")|"
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim i As Integer

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda

        Case 3  'AÑADIR
            If DatosOk Then InsertarCabecera
'                If InsertarDesdeForm2(Me, 2, "Frame2") Then
'                    Data1.RecordSource = "Select * from " & NombreTabla & Ordenacion
'                    PosicionarData
'                End If
'            Else
'                ModificaLineas = 0
'            End If
        

        Case 4  'MODIFICAR
            If DatosOk Then
               If ModificaDesdeFormulario2(Me, 2, "Frame2") Then
                    espera 0.2
                    TerminaBloquear
                    PosicionarData
'                    FormatoDatosTotales
'                    i = Data3.Recordset.AbsolutePosition
                    PonerCampos
                    PonerCamposLineas
'                    SituarDataPosicion Data3, CLng(i), ""
                End If
            End If
            
         Case 5 'InsertarModificar LINEAS
'            If ModificaLineas = 2 Then 'MODIFICAR lineas
'                If ModificarLinea Then
'                    TerminaBloquear
'                    CargaGrid DataGrid1, Data2, True
'                    ModificaLineas = 0
'                    PonerBotonCabecera True
'                    BloquearTxt Text2(16), True
'
'                    LLamaLineas Modo, 0, "DataGrid1"
'                    PosicionarData
'                Else
'                    TerminaBloquear
'                End If
'                Me.DataGrid1.Enabled = True
'            End If
    End Select
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 0, 1, 3 'Busqueda, Insertar
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
            LLamaLineas Modo, 0, "DataGrid2"
            PonerFoco Text1(0)
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
            
        Case 5 'Lineas Detalle
            TerminaBloquear
            BloquearTxt Text2(16), True
            If ModificaLineas = 1 Then 'INSERTAR
                ModificaLineas = 0
                DataGrid1.AllowAddNew = False
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            End If
            ModificaLineas = 0
            LLamaLineas Modo, 0, "DataGrid1"
            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
    End Select
End Sub
Private Sub BotonAnyadir()

    LimpiarCampos 'Huida els TextBox
    PonerModo 3
    
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la capçalera *******
'    Text1(0).Text = SugerirCodigoSiguienteStr("palets", "numpalet")
'    FormateaCampo Text1(0)
    Combo1(0).ListIndex = -1
    
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    Text1(2).Text = Format(Now, "dd/mm/yyyy")
        
    LimpiarDataGrids
    
    PonerFoco Text1(1) '*** 1r camp visible que siga PK ***
    
    ' *** si n'hi han camps de descripció a la capçalera ***
    'PosarDescripcions
End Sub


Private Sub BotonBuscar()
Dim anc As Single

    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        PonerModo 1
        
        'poner los txtaux para buscar por lineas de albaran
        anc = DataGrid2.Top
        If DataGrid2.Row < 0 Then
            anc = anc + 440
        Else
            anc = anc + DataGrid2.RowTop(DataGrid2.Row) + 20
        End If
        LLamaLineas Modo, anc, "DataGrid2"
        
        
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbLightBlue 'vbYellow
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbLightBlue 'vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
    If chkVistaPrevia.Value = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia ""
    Else
        LimpiarCampos
        LimpiarDataGrids
        CadenaConsulta = "Select pedidos.* "
        CadenaConsulta = CadenaConsulta & "from " & NombreTabla
'        CadenaConsulta = CadenaConsulta & " WHERE scafac.codtipom='" & CodTipoMov & "'"
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index, True
    PonerCampos
End Sub


Private Sub BotonModificar()
Dim NroAlbar As String
Dim Cad As String

'--Monica: cambiado por lo de abajo
'    'solo se puede modificar el pedido si no tiene albaran asociado
'    If DBLet(Data1.Recordset!numalbar, "N") <> 0 Then
'        TerminaBloquear
'        Exit Sub
'    End If

    NroAlbar = NroAlbaranAsignado(Data1.Recordset!numpedid, 1)
    If NroAlbar <> "" Then
        Cad = "Este pedido está asociado al albarán " & NroAlbar & "." & vbCrLf
        Cad = Cad & "                ¿ Desea continuar ?"
        If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            TerminaBloquear
            Exit Sub
        End If
    End If

    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    PonerFoco Text1(1) '*** 1r camp visible que siga PK ***
        
End Sub


Private Sub BotonModificarLinea()
'Modificar una linea
Dim vWhere As String
Dim anc As Single
Dim J As Byte

    On Error GoTo eModificarLinea


'     'solo se puede modificar la factura si no esta contabilizada
'    If FactContabilizada Then
'        TerminaBloquear
'        Exit Sub
'    End If

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then  '1= Insertar
        TerminaBloquear
        Exit Sub
    End If
    
    If Data2.Recordset.EOF Then
        TerminaBloquear
        Exit Sub
    End If
    
    vWhere = ObtenerWhereCP(False)
    vWhere = vWhere & " AND codtipoa='" & Data3.Recordset.Fields!codtipoa & "' AND numalbar=" & Data3.Recordset.Fields!NumAlbar
    vWhere = vWhere & " and numlinea=" & Data2.Recordset!NumLinea
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then
        TerminaBloquear
        Exit Sub
    End If

    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        J = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, J
        DataGrid1.Refresh
    End If
    
'    anc = ObtenerAlto(Me.DataGrid1)
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 210
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 10
    End If

    For J = 0 To 2
        txtAux(J).Text = DataGrid1.Columns(J + 5).Text
    Next J
    Text2(16).Text = DataGrid1.Columns(J + 5).Text
    For J = J + 1 To 9
        txtAux(J - 1).Text = DataGrid1.Columns(J + 5).Text
    Next J
    
    ModificaLineas = 2 'Modificar
    LLamaLineas ModificaLineas, anc, "DataGrid1"
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerBotonCabecera False
    BloquearTxt Text2(16), False 'Campo Ampliacion Linea
    PonerFoco txtAux(4)
    Me.DataGrid1.Enabled = False

eModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub LLamaLineas(xModo As Byte, Optional alto As Single, Optional grid As String)
Dim jj As Integer
Dim b As Boolean

    Select Case grid
        Case "DataGrid1"
            DeseleccionaGrid Me.DataGrid1
            'PonerModo xModo + 1
    
            b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Lineas
    
            For jj = 0 To txtAux.Count - 1
                If jj = 4 Or jj = 6 Or jj = 7 Or jj = 8 Then
                    txtAux(jj).Height = DataGrid1.RowHeight
                    txtAux(jj).Top = alto
                    txtAux(jj).visible = b
                End If
            Next jj
            
        Case "DataGrid2"
            DeseleccionaGrid Me.DataGrid2
            b = (xModo = 1)
             For jj = 0 To txtAux3.Count - 1
                txtAux3(jj).Height = DataGrid2.RowHeight
                txtAux3(jj).Top = alto - 200
                txtAux3(jj).visible = b
            Next jj
    End Select
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Facturas (scafac)
' y los registros correspondientes de las tablas cab. albaranes (scafac1)
' y las lineas de la factura (slifac)
Dim Cad As String
Dim NroAlbar As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
        
'--Monica: cambiado por lo de abajo
'    'solo se puede modificar el pedido si no tiene numero de albaran
'    If DBLet(Data1.Recordset!numalbar, "N") <> 0 Then Exit Sub
    NroAlbar = NroAlbaranAsignado(Data1.Recordset!numpedid, 1)
    If NroAlbar <> "" Then
        Cad = "Este pedido está asociado al albarán " & NroAlbar & "." & vbCrLf
        Cad = Cad & "                ¿ Desea continuar ?"
        If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Sub
        End If
    End If
    
    Cad = "Cabecera de Pedidos." & vbCrLf
    Cad = Cad & "-------------------------------------      " & vbCrLf & vbCrLf
    Cad = Cad & "Va a eliminar el Pedido:            "
    Cad = Cad & vbCrLf & "Nº Pedido:  " & Format(Text1(0).Text, "0000000")
    Cad = Cad & vbCrLf & "Fecha:  " & Format(Text1(1).Text, "dd/mm/yyyy")

    Cad = Cad & vbCrLf & vbCrLf & " ¿Desea Eliminarlo? "

    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
'        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
'        NumPedElim = Data1.Recordset.Fields(1).Value
        
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
        End If
        
'        'Devolvemos contador, si no estamos actualizando
'        Set vTipoMov = New CTiposMov
'        vTipoMov.DevolverContador CodTipoMov, NumPedElim
'        Set vTipoMov = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
EEliminar:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminar Albaran", Err.Description
End Sub


'Private Sub BloqueaText3()
'Dim i As Byte
'    'bloquear los Text3 que son las lineas de scafac1
'    For i = 0 To 3
'        BloquearTxt Text3(i), (Modo <> 4)
'    Next i
'    If Me.FrameObserva.visible Then
'        For i = 9 To 13
'            BloquearTxt Text3(i), (Modo <> 4)
'        Next i
'    End If
'    For i = 4 To 8
'        BloquearTxt Text3(i), True
'    Next i
'
'    'datos venta TPV
'    BloquearTxt Text3(14), True
'    BloquearTxt Text3(15), True
'End Sub


Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim Cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        DataGrid2.Enabled = True
        If Not Data1.Recordset.EOF Then _
            Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

    Else 'Se llama desde algún Prismatico de otro Form al Mantenimiento de Trabajadores
        If Data1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        Cad = Data1.Recordset.Fields(0) & "|"
        Cad = Cad & Data1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(Cad)
        Unload Me
    End If
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

' doble click en el grid de variedades
Private Sub DataGrid2_DblClick()
    If Data3.Recordset.EOF Then Exit Sub

    Set frmLPed = New frmVtasLinPedidos
    
    frmLPed.ModoExt = 0
    frmLPed.Pedido = Data3.Recordset.Fields(0).Value
    frmLPed.Linea = Data3.Recordset.Fields(1).Value
    frmLPed.Show vbModal
    
    Set frmLPed = Nothing
End Sub

Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim i As Byte

'    If LastCol = -1 Then Exit Sub

    'Datos de la tabla pedidos_calibres
    If Not Data3.Recordset.EOF Then
        'Datos de la tabla palets_calibres
        CargaGrid DataGrid1, Data2, True
    Else
        'Datos de la tabla palets_calibres
        CargaGrid DataGrid1, Data2, False
    End If
    
'    CargaForaGrid
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
    'Viene de DblClick en frmAlmMovimArticulos y carga el form con los valores
    If hcoCodMovim <> "" And Not Data1.Recordset.EOF Then PonerCadenaBusqueda
    
'    PonerCadenaBusqueda
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
Dim i As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
     'Icono de busqueda
    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next kCampo

    ' ICONITOS DE LA BARRA
'    btnPrimero = 16
'    With Me.Toolbar1
'        .HotImageList = frmPpal.imgListComun_OM
'        .DisabledImageList = frmPpal.imgListComun_BN
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 1   'Botón Buscar
'        .Buttons(2).Image = 2   'Botón Todos
'        .Buttons(4).Image = 3   'Insertar
'        .Buttons(5).Image = 4   'Modificar
'        .Buttons(6).Image = 5   'Borrar
'        .Buttons(8).Image = 10  'Impresión de Pedido
'        .Buttons(9).Image = 26  'Impresión de Proveedor
'
'        .Buttons(10).Image = 24  'Orden de Carga
'        .Buttons(11).Image = 23 'Impresion CMR
'        .Buttons(12).Image = 17 'Generar Albaran
'        .Buttons(14).Image = 11  'Salir
'
'        .Buttons(btnPrimero).Image = 6  'Primero
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
'        .Buttons(btnPrimero + 3).Image = 9 'Último
'    End With
    
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'l'1 i el 2 son separadors
        .Buttons(5).Image = 1   'Buscar
        .Buttons(6).Image = 2   'Totss
        'el 5 i el 6 son separadors
        .Buttons(1).Image = 3   'Insertar
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Borrar
        'el 10  son separadors
        .Buttons(8).Image = 10  'Imprimir
    End With
    
    With Me.Toolbar5
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26  'Impresión de Proveedor
        .Buttons(2).Image = 24  'Orden de Carga
        .Buttons(3).Image = 17 'Generar Albaran
    End With
    
    '[Monica]05/06/2018: si no se hace la impresion de proveedor se rediseñan los toolbar
    If vParamAplic.Cooperativa <> 15 Then
        Me.Toolbar5.Width = Me.Toolbar5.Width - 450
        Me.FrameBotonGnral2.Width = Me.FrameBotonGnral2.Width - 450
        Me.FrameDesplazamiento.Left = FrameDesplazamiento.Left - 450
    End If
    
    ' desplazamiento
    With Me.ToolbarDes
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 6
        .Buttons(2).Image = 7
        .Buttons(3).Image = 8
        .Buttons(4).Image = 9
    End With
    
'    ' La Ayuda
'    With Me.ToolbarAyuda
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 12
'    End With
    
    
    
    ' ******* si n'hi han llínies *******
    'ICONETS DE LES BARRES ALS TABS DE LLÍNIA
    For kCampo = 0 To ToolAux.Count - 1
        With Me.ToolAux(kCampo)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
        End With
    Next kCampo
   ' ***********************************
   'IMAGES para zoom
    For i = 0 To Me.imgZoom.Count - 1
        Me.imgZoom(i).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    Next i
    
    LimpiarCampos   'Limpia los campos TextBox
    CargaCombo
    
    If vParamAplic.Cooperativa = 15 Then
        CodTipoMov = "PE1" 'hcoCodTipoM
    Else
        CodTipoMov = "PEV" 'hcoCodTipoM
    End If
    VieneDeBuscar = False
    
        
    '## A mano
    NombreTabla = "pedidos"
    NomTablaLineas = "pedidos_variedad" 'Tabla lineas de variedades
    Ordenacion = " ORDER BY pedidos.numpedid"
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    CadenaConsulta = "select * from pedidos "
    If hcoCodMovim <> "" Then
        CadenaConsulta = CadenaConsulta & " where numpedid =" & hcoCodMovim
    Else
        CadenaConsulta = CadenaConsulta & " where numpedid = -1"
    End If
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        BotonBuscar
    End If
'        CargaGrid DataGrid1, Data2, False
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    PrimeraVez = False
   
End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.Combo1(0).ListIndex = -1
    Me.Check1(0).Value = 0
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    If Modo = 4 Then TerminaBloquear
End Sub


Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod almacen
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nombre del almacen
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadB = Aux
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
    Screen.MousePointer = vbDefault
End Sub



Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    If imgFec(0).Tag < 2 Then
        Text1(CByte(imgFec(0).Tag) + 1).Text = Format(vFecha, "dd/mm/yyyy") '<===
    Else
        Text1(CByte(imgFec(0).Tag) + 8).Text = Format(vFecha, "dd/mm/yyyy") '<===
    End If
    ' ********************************************
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000") 'Cod Cliente
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 3) 'Nombre del cliente
End Sub

Private Sub frmDest_DatoSeleccionado(CadenaSeleccion As String)
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Destino
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nombre del destino
    
    Text1(5) = DevuelveDesdeBDNew(cAgro, "destinos", "codtimer", "codclien", Text1(3).Text, "N", , "coddesti", Text1(4), "N")
    Text1(5) = Format(Text1(5), "000")
    Text2(5) = DevuelveDesdeBDNew(cAgro, "tipomer", "nomtimer", "codtimer", Text1(5).Text, "N")
    
    MostrarCadena Text1(3), Text1(4)
    
End Sub

Private Sub frmList_DatoSeleccionado(CadenaSeleccion As String)
'Cuando pasa de Pedido -> Albaran
'Aqui devuelve los valores que se introducen desde el Form de Listado de Pedido
'para generar el Albaran
Dim vSQL As String

    'Construimos parte de la SQL para insertar en tabla de Albaranes(scaalb)
    
    FechaAlb = RecuperaValor(CadenaSeleccion, 1)
    vSQL = ""
    vSQL = " '" & Format(FechaAlb, FormatoFecha) & "' as fechaalb, " 'Fecha Albaran
    vSQL = vSQL & Text1(3).Text & ", " 'codigo cliente
    vSQL = vSQL & Text1(4).Text & ", " 'codigo destino
    vSQL = vSQL & Text1(6).Text & ", " 'agencia de transporte
    vSQL = vSQL & DBSet(Text1(7).Text, "T") & "," ' matricula de vehiculo
    vSQL = vSQL & DBSet(Text1(8).Text, "T") & "," ' matricula de remolque
    vSQL = vSQL & DBSet(Text1(14).Text, "T") & "," ' referencia cliente
    vSQL = vSQL & Text1(5).Text & ", " ' tipo de mercado
    vSQL = vSQL & DBSet(Text1(11).Text, "N") & "," ' total paletss
    vSQL = vSQL & ValorNulo & "," 'portes previstos
    vSQL = vSQL & DBSet(Text1(12).Text, "T") & ", " 'nro de contrato
    vSQL = vSQL & DBSet(Text1(13).Text, "N") & ", " ' nro actas
    vSQL = vSQL & Text1(0).Text & " as numpedid, '" 'Nº Pedido
    vSQL = vSQL & Format(Text1(1).Text, FormatoFecha) & "' as fechaped, " 'Fecha Pedido
    vSQL = vSQL & DBSet(Text1(15).Text, "T") & ", " 'observaciones
    vSQL = vSQL & "0," 'pasa a aridoc
    vSQL = vSQL & DBSet(Text1(16).Text, "N") ' codigo de almacen
    CadenaSQL = vSQL
    
    'Se almacena aqui si el usuario quiere imprimir el Albaran tras generarlo
    ImprimeAlb = CBool(RecuperaValor(CadenaSeleccion, 2))
    Incidencia = RecuperaValor(CadenaSeleccion, 3)
End Sub

' devolvemos la linea del datagrid en donde estabamos
Private Sub frmLPed_DatoSeleccionado(CadenaSeleccion As String)
Dim vWhere As String
             
   PonerCamposLineas
   
   If CadenaSeleccion = "" Then Exit Sub
             
   vWhere = "(numpedid = " & RecuperaValor(CadenaSeleccion, 1) & " and numlinea = " & RecuperaValor(CadenaSeleccion, 2) & ")"
   SituarDataMULTI Data3, vWhere, "" ', Indicador
   
   PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
   PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
   

End Sub

Private Sub frmMen_DatoSeleccionado(CadenaSeleccion As String)
    Continuar = (CadenaSeleccion = "1")
End Sub

Private Sub frmMer_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Tipos de Mercado
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Mercado
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Mercado
End Sub

Private Sub frmPal_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de paelets de confeccion
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod palet
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Descripcion
End Sub

Private Sub frmPed_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String
Dim Aux As String
      
    If CadenaSeleccion <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        CadB = "numpedid = " & RecuperaValor(CadenaSeleccion, 1)
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Agencias de Transporte
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Agencias de Transporte
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Descripcion
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(indice).Text = vCampo
End Sub

Private Sub imgBuscar_Click(Index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. de Cliente
            indice = 3
            PonerFoco Text1(indice)
            Set frmCli = New frmClientes
            frmCli.DatosADevolverBusqueda = "0|1|2|"
            frmCli.Show vbModal
            Set frmCli = Nothing
            PonerFoco Text1(indice)
        
        Case 1 'Cod. de Destino de Cliente
            If Text1(3) = "" Then Exit Sub
            
            indice = 4
            PonerFoco Text1(indice)
            Set frmDest = New frmDestCli
            frmDest.Cliente = Text1(3)
            frmDest.DatosADevolverBusqueda = "0|1|"
            frmDest.Show vbModal
            Set frmDest = Nothing
            PonerFoco Text1(indice)
            
        Case 2 ' Tipo de Mercado
            indice = 5
            PonerFoco Text1(indice)
            Set frmMer = New frmManTipMerc
            frmMer.DatosADevolverBusqueda = "0|1|2|"
            frmMer.Show vbModal
            Set frmMer = Nothing
            PonerFoco Text1(indice)
        
        Case 3 ' Agencia de Transporte
            indice = 6
            PonerFoco Text1(indice)
            Set frmTra = New frmManAgencias
            frmTra.DatosADevolverBusqueda = "0|1|2|"
            frmTra.Show vbModal
            Set frmTra = Nothing
            PonerFoco Text1(indice)
    
        Case 4 ' Almacén
            indice = 16
            PonerFoco Text1(indice)
            Set frmAlm = New frmManAlmProp
            frmAlm.DatosADevolverBusqueda = "0|1|"
            frmAlm.Show vbModal
            Set frmAlm = Nothing
            PonerFoco Text1(indice)
    
        Case 5 ' Numero de palet
            indice = 17
            PonerFoco Text1(indice)
            Set frmPal = New frmManPaleConf
            frmPal.DatosADevolverBusqueda = "0|1|"
            frmPal.Show vbModal
            Set frmPal = Nothing
            PonerFoco Text1(indice)
    
    
    End Select
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub imgFec_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
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
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40

    If Index < 2 Then
        imgFec(0).Tag = Index '<===
        ' *** repasar si el camp es txtAux o Text1 ***
        If Text1(Index + 1).Text <> "" Then frmC.NovaData = Text1(Index + 1).Text
    Else
        imgFec(0).Tag = Index '<===
        ' *** repasar si el camp es txtAux o Text1 ***
        If Text1(Index + 8).Text <> "" Then frmC.NovaData = Text1(Index + 8).Text
    End If
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    If Index < 2 Then
        PonerFoco Text1(CByte(imgFec(0).Tag) + 1) '<===
    Else
        PonerFoco Text1(CByte(imgFec(0).Tag) + 8) '<===
    End If
    ' ********************************************
End Sub


Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        indice = 15
        frmZ.pTitulo = "Observaciones del Pedido"
        frmZ.pValor = Text1(indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco Text1(indice)
    End If
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub


Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Pedido
'         BotonEliminarLinea
    Else   'Eliminar Pedido
         BotonEliminar
    End If
End Sub

Private Sub mnGenerarAlb_Click()
'Pasar una Pedido a Albaran
Dim Resp As Byte
Dim b As Boolean
Dim cadMen As String

    'Comprobar que hay un Pedido seleccionado
    If Data1.Recordset.EOF Then Exit Sub
    
    If DBLet(Data1.Recordset!NumAlbar, "N") <> 0 Then
        MsgBox "Este pedido ya tiene asociado el albarán: " & DBLet(Data1.Recordset!NumAlbar, "N"), vbExclamation
        Exit Sub
    End If

    If TienePalets(Data1.Recordset!numpedid) Then
        'preguntar si utiliza o no palet
        Resp = MsgBox("¿Utiliza Palets?", vbYesNoCancel)
        If Resp = vbCancel Then Exit Sub
    
        If Resp = vbYes Then
            AlbDePalet = True 'VIENE DEL PALET
        Else
            AlbDePalet = False
        End If
    Else
        ' no tiene palets asociados
        AlbDePalet = False
    End If
    
    Continuar = True
    If AlbDePalet Then 'VIENE DEL PALET
        'mostramos los palets asociados al pedido
        Set frmMen = New frmMensajes
        frmMen.vCampos = DBLet(Data1.Recordset!numpedid, "N")
        frmMen.cadwhere = "select * from palets where numpedid = " & DBLet(Data1.Recordset!numpedid, "N")
        frmMen.OpcionMensaje = 19 'Palets asociados al pedido
        frmMen.Show vbModal
        Set frmMen = Nothing
    Else ' 23/11/2009: viene de pedido compruebo que tengan lineas
        Continuar = (TotalRegistros("select count(*) from pedidos_variedad where numpedid = " & DBLet(Data1.Recordset!numpedid, "N")) <> 0)
        If Not Continuar Then
            MsgBox "El pedido no tiene lineas. Revise.", vbExclamation
        End If
    End If
    If Continuar Then
        Screen.MousePointer = vbHourglass
        GenerarAlbaran
    End If
End Sub

Private Sub mnImprimir_Click()
'Imprimir Factura
    
    If Data1.Recordset.EOF Then Exit Sub
    
    BotonImprimir 0
End Sub


Private Sub mnImprimirProv_Click()
'Imprimir Factura
    
    If Data1.Recordset.EOF Then Exit Sub
    
    BotonImprimir 1
End Sub



Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub


Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
        'bloquea la tabla cabecera de factura: scafac
        If BLOQUEADesdeFormulario(Me) Then
            'bloquear la tabla cabecera de albaranes de la factura: scafac1
            If BloqueaAlbxFac Then
                If BloqueaLineasFac Then BotonModificarLinea
            End If
        End If
         
    Else   'Modificar Pedido
        'bloquea la tabla cabecera de factura: pedidos
        If BLOQUEADesdeFormulario(Me) Then
            'bloquear la tabla cabecera de albaranes de la factura: scafac1
            BotonModificar
        End If
    End If
End Sub


Private Function BloqueaAlbxFac() As Boolean
'bloquea todos los albaranes de la factura
Dim Sql As String

    On Error GoTo EBloqueaAlb
    
    BloqueaAlbxFac = False
    'bloquear cabecera albaranes x factura
    Sql = "select * FROM scafac1 "
    Sql = Sql & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute Sql, , adCmdText
    BloqueaAlbxFac = True

EBloqueaAlb:
    If Err.Number <> 0 Then BloqueaAlbxFac = False
End Function


Private Function BloqueaLineasFac() As Boolean
'bloquea todas las lineas de la factura
Dim Sql As String

    On Error GoTo EBloqueaLin

    BloqueaLineasFac = False
    'bloquear cabecera albaranes x factura
    Sql = "select * FROM slifac "
    Sql = Sql & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute Sql, , adCmdText
    BloqueaLineasFac = True

EBloqueaLin:
    If Err.Number <> 0 Then BloqueaLineasFac = False
End Function


Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    If (Modo = 5) Then 'Modo 5: Mto Lineas
        '1:Insertar linea, 2: Modificar
        If ModificaLineas = 1 Or ModificaLineas = 2 Then cmdCancelar_Click
        cmdRegresar_Click
        Exit Sub
    End If
    Unload Me
End Sub


Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub


Private Sub Text1_Change(Index As Integer)
    If Index = 9 Then HaCambiadoCP = True 'Cod. Postal
End Sub

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
'    If Index = 9 Then HaCambiadoCP = False 'CPostal
'    If Index = 1 And Modo = 1 Then
'        SendKeys "{tab}"
'        Exit Sub
'    End If
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 15 Or (Index = 15 And Text1(15).Text = "") Then KEYpress KeyAscii
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
Dim devuelve As String
Dim cadMen As String
Dim Sql As String

        
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1, 2, 10 'Fecha pedido y fecha de carga
            '[Monica]28/08/2013: controlamos que esté dentro de campaña
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index), True
        
        Case 3 'Cliente
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "clientes", "nomclien")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Cliente: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCli = New frmClientes
                        frmCli.DatosADevolverBusqueda = "0|1|"
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmCli.Show vbModal
                        Set frmCli = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                Else
                    ' mostramos en el label3 la cadena
                    MostrarCadena Text1(Index), Text1(4)
                End If
            End If
                
        Case 4 ' Destino del cliente
            If PonerFormatoEntero(Text1(Index)) Then
                If Text1(3).Text <> "" Then
                    Text2(Index).Text = DevuelveDesdeBDNew(cAgro, "destinos", "nomdesti", "codclien", Text1(3), "N", , "coddesti", Text1(4), "N")
                    If Text2(Index).Text = "" Then
                        cadMen = "No existe el Destino: " & Text1(Index).Text & vbCrLf
                        cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                        If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                            Set frmCli = New frmClientes
                            frmCli.DatosADevolverBusqueda = "0|1|"
                            Text1(Index).Text = ""
                            TerminaBloquear
                            frmCli.Show vbModal
                            Set frmCli = Nothing
                            If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                        Else
                            Text1(Index).Text = ""
                        End If
                        PonerFoco Text1(Index)
                    Else
                        ' traemos el tipo de mercado del destino
                        Text1(5).Text = DevuelveDesdeBDNew(cAgro, "destinos", "codtimer", "codclien", Text1(3), "N", , "coddesti", Text1(4), "N")
                        PonerFormatoEntero Text1(5)
                        If Text1(5) <> "" Then
                            Text2(5).Text = PonerNombreDeCod(Text1(5), "tipomer", "nomtimer", "codtimer", "N")
                        End If
                        ' mostramos en el label3 la cadena
                        MostrarCadena Text1(3), Text1(4)
                    End If
                Else
                    MsgBox "Debe introducir previamente el cliente.", vbExclamation
                    PonerFoco Text1(3)
                End If
            End If
            
        Case 5 'Tipo de Mercado
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "tipomer", "nomtimer")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Tipo de Mercado: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmMer = New frmManTipMerc
                        frmMer.DatosADevolverBusqueda = "0|1|"
                        frmMer.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmMer.Show vbModal
                        Set frmMer = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            End If
                
        Case 6 'Agencia de Transporte
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "agencias", "nomtrans")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Agencia de Transporte: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmTra = New frmManAgencias
                        frmTra.DatosADevolverBusqueda = "0|1|"
                        frmTra.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmTra.Show vbModal
                        Set frmTra = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            End If
            
        Case 16 ' Almacen
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "salmpr", "nomalmac")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Almacén: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmAlm = New frmManAlmProp
                        frmAlm.DatosADevolverBusqueda = "0|1|"
                        frmAlm.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmAlm.Show vbModal
                        Set frmAlm = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            End If
        
        Case 17 'Tipo de palet
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "confpale", "nompalet")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Tipo de Palet: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmPal = New frmManPaleConf
                        frmPal.DatosADevolverBusqueda = "0|1|"
                        frmPal.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmPal.Show vbModal
                        Set frmPal = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            End If
        
    End Select
End Sub


Private Sub HacerBusqueda()
Dim CadB As String
Dim cadAux As String
    
'    '--- Laura 12/01/2007
'    cadAux = Text1(5).Text
'    If Text1(4).Text <> "" Then Text1(5).Text = ""
'    '---
    
'    '--- Laura 12/01/2007
'    Text1(5).Text = cadAux
'    '---
'    CadB = ObtenerBusqueda(Me)
    CadB = ObtenerBusqueda2(Me, BuscaChekc)

    If chkVistaPrevia = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select pedidos.* from " & NombreTabla & " LEFT JOIN pedidos_variedad ON pedidos.numpedid=pedidos_variedad.numpedid "
        CadenaConsulta = CadenaConsulta & " WHERE " & CadB & " GROUP BY pedidos.numpedid " & Ordenacion
'        CadenaConsulta = "select palets.* from " & NombreTabla
'        CadenaConsulta = CadenaConsulta & " WHERE " & CadB & " GROUP BY palets.numpalet " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim tabla As String
Dim Titulo As String
Dim Desc As String, devuelve As String
'    'Llamamos a al form
'    '##A mano
'    Cad = ""
'    Cad = Cad & "Nº.Pedido|pedidos.numpedid|N||15·"
'
'    If vParamAplic.Cooperativa <> 18 Then
'        Cad = Cad & "Cliente|pedidos.codclien|N||10·" 'ParaGrid(Text1(3), 10, "Cliente")
'        Cad = Cad & "Nombre Cliente|clientes.nomclien|N||45·"
'    Else
'        Cad = Cad & "Nombre Cliente|clientes.nomclien|N||40·"
'        Cad = Cad & "Ref.Cliente|pedidos.refclien|N||15·"
'    End If
'    Cad = Cad & ParaGrid(Text1(1), 15, "F.Pedido")
'    Cad = Cad & ParaGrid(Text1(2), 15, "F.Carga")
'    tabla = NombreTabla & " INNER JOIN clientes ON pedidos.codclien=clientes.codclien "
'
'    Titulo = "Pedidos"
'    devuelve = "0|"
'
'    If Cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = Cad
'        frmB.vtabla = tabla
'        frmB.vSQL = CadB
'        HaDevueltoDatos = False
'        '###A mano
'        frmB.vDevuelve = "0|1|"
'        frmB.vDevuelve = devuelve
'        frmB.vTitulo = Titulo
'        frmB.vSelElem = 0
''        frmB.vConexionGrid = cAgro  'Conexión a BD: Ariagro
'        If Not EsCabecera Then frmB.Label1.FontSize = 11
''        frmB.vBuscaPrevia = chkVistaPrevia
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
''        If EsCabecera Then
''            PonerCadenaBusqueda
''            Text1(0).Text = Format(Text1(0).Text, "0000000")
''        End If
'        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
'        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
'            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
'        End If
'    End If
'    Screen.MousePointer = vbDefault

    Set frmPed = New frmBasico2
    
    AyudaPedidos frmPed, Text1(0), CadB
    
    Set frmPed = Nothing

End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass

    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        If Modo = 1 Then
            PonerFoco Text1(kCampo)
'            Text1(0).BackColor = vbYellow
        End If
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
        LLamaLineas Modo, 0, "DataGrid2"
        PonerCampos
    End If


    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCamposLineas()
'Carga el grid de los AlbaranesxFactura, es decir, la tabla scafac1 de la factura seleccionada
Dim b As Boolean
Dim b2 As Boolean

    On Error GoTo EPonerLineas

    If Data1.Recordset.EOF Then Exit Sub

    Screen.MousePointer = vbHourglass
    
    'Datos de la tabla albaranes x factura: scafac1
    CargaGrid DataGrid2, Data3, True
    '++monica
    If Data3.Recordset.RecordCount > 0 Then
        CargaGrid DataGrid1, Data2, True
    Else
        CargaGrid DataGrid1, Data2, False
    End If
    '++
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim BrutoFac As Single

    On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 2, "Frame2"
    
'    FormatoDatosTotales
    
    'poner descripcion campos
    Modo = 4
    
    Text2(3).Text = PonerNombreDeCod(Text1(3), "clientes", "nomclien", "codclien", "N") 'cliente
    Text2(4).Text = DevuelveDesdeBDNew(cAgro, "destinos", "nomdesti", "codclien", Text1(3), "N", , "coddesti", Text1(4), "N") 'destino
    Text2(5).Text = PonerNombreDeCod(Text1(5), "tipomer", "nomtimer", "codtimer", "N") 'tipo de mercado
    Text2(6).Text = PonerNombreDeCod(Text1(6), "agencias", "nomtrans", "codtrans", "N") 'agencia
    Text2(16).Text = PonerNombreDeCod(Text1(16), "salmpr", "nomalmac", "codalmac", "N") 'almacen
    Text2(17).Text = PonerNombreDeCod(Text1(17), "confpale", "nompalet", "codpalet", "N") 'palets
    
    MostrarCadena Text1(3), Text1(4)

    Modo = 2
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Ofertas
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario

    If Err.Number <> 0 Then Err.Clear
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte, NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo

    'Actualiza Iconos Insertar,Modificar,Eliminar
    '## No tiene el boton modificar y no utiliza la funcion general
'    ActualizarToolbar Modo, Kmodo
    BuscaChekc = ""
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Or hcoCodMovim <> "" Then
        cmdRegresar.visible = b
    Else
        cmdRegresar.visible = False
    End If
        
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, Numreg
    DesplazamientoVisible b And Data1.Recordset.RecordCount > 1
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    'si estamos en modificar bloquea las compos que son clave primaria
    BloquearText1 Me, Modo
    BloquearCombo Me, Modo
    For i = 9 To 10
        BloquearTxt Text1(i), Not (Modo = 1)
        Text1(i).Enabled = (Modo = 1)
    Next i
    Me.Check1(0).Enabled = (Modo = 1)
    
    b = (Modo <> 1)
    'Campos Nº Pedido bloqueado y en azul
    BloquearTxt Text1(0), b, True
'    BloquearTxt Text1(3), b 'referencia
    
    
    'bloquear los Text3 que son las lineas de scafac1
'    BloqueaText3
    
    'Si no es modo lineas Boquear los TxtAux
    For i = 0 To txtAux.Count - 1
        BloquearTxt txtAux(i), (Modo <> 5)
    Next i
'    BloquearTxt txtAux(8), True
    
    'Si no es modo Busqueda Bloquear los TxtAux3 (son los txtaux de los variedades de palets)
'    For i = 0 To txtAux3.Count - 1
'        BloquearTxt txtAux3(i), True '(Modo <> 1)
'    Next i
    For i = 0 To 7
        BloquearTxt txtAux3(i), True
        txtAux3(i).Enabled = False
    Next i
    For i = 11 To 12
        BloquearTxt txtAux3(i), True
        txtAux3(i).Enabled = False
    Next i
    For i = 8 To 10
        BloquearTxt txtAux3(i), (Modo <> 1)
        txtAux3(i).Enabled = (Modo = 1)
    Next i
    For i = 13 To 15
        BloquearTxt txtAux3(i), (Modo <> 1)
        txtAux3(i).Enabled = (Modo = 1)
    Next i
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    BloquearImgBuscar Me, Modo, ModificaLineas
    BloquearImgFec Me, 0, Modo
    BloquearImgFec Me, 1, Modo
    
    imgFec(2).Enabled = (Modo = 1)
    imgFec(2).visible = (Modo = 1)
    
    If Modo <> 4 Then Label3.Caption = ""
    
'    Me.imgBuscar(1).visible = False
                    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub DesplazamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim b As Boolean

    On Error GoTo EDatosOK

    DatosOk = False
    
'    ComprobarDatosTotales

    'comprobamos datos OK de la tabla scafac
    b = CompForm2(Me, 2, "Frame2") ' , 1) 'Comprobar formato datos ok de la cabecera: opcion=1
    If Not b Then Exit Function
    
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
Dim b As Boolean
Dim i As Byte

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    b = True

    For i = 0 To txtAux.Count - 1
        If i = 4 Or i = 6 Or i = 7 Then
            If txtAux(i).Text = "" Then
                MsgBox "El campo " & txtAux(i).Tag & " no puede ser nulo", vbExclamation
                b = False
                PonerFoco txtAux(i)
                Exit Function
            End If
        End If
    Next i
            
    DatosOkLinea = b
    
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 16 And KeyCode = 40 Then 'campo Amliacion Linea y Flecha hacia abajo
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 16 And KeyAscii = 13 Then 'campo Amliacion Linea y ENTER
        PonerFocoBtn Me.cmdAceptar
    End If
End Sub


Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
Dim NroAlbar As String
Dim Cad As String
    
    NroAlbar = NroAlbaranAsignado(Data1.Recordset!numpedid, 1)
    If NroAlbar <> "" Then
        Cad = "Este pedido está asociado al albarán " & NroAlbar & "." & vbCrLf
        Cad = Cad & "                ¿ Desea continuar ?"
        If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Sub
        End If
    End If

    If BloqueaRegistro(NombreTabla, "numpedid = " & Data1.Recordset!numpedid) Then
'    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
        Select Case Button.Index
            Case 1 'añadir variedad
                Set frmLPed = New frmVtasLinPedidos
                
                frmLPed.ModoExt = 3
                frmLPed.Pedido = Data1.Recordset.Fields(0).Value
                frmLPed.Show vbModal
            
                Set frmLPed = Nothing
            Case 2 'modificar variedad
                Set frmLPed = New frmVtasLinPedidos
                
                frmLPed.ModoExt = 4
                frmLPed.Pedido = Data3.Recordset.Fields(0).Value
                frmLPed.Linea = Data3.Recordset.Fields(1).Value
                frmLPed.Show vbModal
                
                Set frmLPed = Nothing
                
            Case 3 ' boton eliminar linea de variedades
                BotonEliminarLinea
            Case Else
        End Select
        PonerCampos
        TerminaBloquear
    End If
End Sub


Private Sub BotonEliminarLinea()
Dim Cad As String

    On Error GoTo EEliminarLinea

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
    ' ***************************************************************************

    ' *************** canviar la pregunta ****************
    Cad = "¿Seguro que desea eliminar la Variedad?"
    Cad = Cad & vbCrLf & "Pedido: " & Data3.Recordset.Fields(0)
    Cad = Cad & vbCrLf & "Variedad: " & Data3.Recordset.Fields(3)
    
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        On Error GoTo EEliminarLinea
        Screen.MousePointer = vbHourglass
        NumRegElim = Data3.Recordset.AbsolutePosition
        If Not EliminarLinea Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data3, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
    End If
    Screen.MousePointer = vbDefault
    
EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Variedad de Pedido", Err.Description

End Sub



'Private Sub Text3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    KEYdown KeyCode
'End Sub
'
'Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
'    KEYpress KeyAscii
'End Sub
'
'Private Sub Text3_LostFocus(Index As Integer)
'    Select Case Index
'        Case 0, 1, 2 'trabajador
'            Text2(Index).Text = PonerNombreDeCod(Text3(Index), conAri, "straba", "nomtraba", "codtraba", "Cod. Trabajador", "N")
'        Case 3 'cod. envio
'            Text2(Index).Text = PonerNombreDeCod(Text3(Index), conAri, "senvio", "nomenvio", "codenvio", "Cod. Envio", "N")
'            If Screen.ActiveControl.TabIndex <> 27 Then PonerFocoBtn Me.cmdAceptar
'        Case 13 'observa 5
'            PonerFocoBtn Me.cmdAceptar
'    End Select
'End Sub
'
Private Sub mnOrdenCarga_Click()
'Imprimir la Orden de Carga
    
    If Data1.Recordset.EOF Then Exit Sub
    
    BotonOrdenCarga
End Sub

Private Sub mnCMR_Click()
'Imprimir la Orden de Carga
    
    If Data1.Recordset.EOF Then Exit Sub
    
    BotonCMR
End Sub

Private Sub BotonOrdenCarga()
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadselect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String

    If Text1(0).Text = "" Then
        MsgBox "Debe seleccionar un Pedido para Imprimir.", vbInformation
        Exit Sub
    End If
    
    cadFormula = ""
    cadParam = ""
    cadselect = ""
    numParam = 0
    
    If vParamAplic.Cooperativa = 15 Then
        Set frmOrden = New frmVtasOrdenCarga
        
        frmOrden.NumCod = Mid(Text1(0).Text, 4, 4) & "/" & Year(CDate(Text1(1).Text))
        frmOrden.Show vbModal
        
        Set frmOrden = Nothing
        
        Exit Sub
    End If
        
    
    
    '===================================================
    '============ PARAMETROS ===========================
    indRPT = 10 'Impresion de Orden de Carga
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
      
    cadParam = cadParam & "pEmpresa='" & vEmpresa.nomempre & "'|"
    numParam = numParam + 1
      
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de albaran
    '---------------------------------------------------
    If Text1(0).Text <> "" Then
        'Nº pedido
        devuelve = "{palets.numpedid}=" & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        devuelve = "numpedid = " & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadselect, devuelve) Then Exit Sub
    End If
    
    If Not HayRegParaInforme("palets", cadselect) Then Exit Sub
     
     With frmImprimir
          '[Monica]24/01/2012: añadido la siguientes 3 lineas para el envio por el outlook
            .outClaveNombreArchiv = Format(Text1(0).Text, "0000000")
            .outCodigoCliProv = Text1(3).Text
            '[Monica]06/05/2015: destino para sacar email
            .outCodigoDestino = Text1(4).Text
            .outTipoDocumento = 6
     
     
     
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 0
            .ConSubInforme = True
            .Titulo = "Orden de Carga"
            .Show vbModal
    End With
End Sub


Private Sub BotonCMR()
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadselect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String

'    If Text1(0).Text = "" Then
'        MsgBox "Debe seleccionar un Albarán para Imprimir.", vbInformation
'        Exit Sub
'    End If
'
'    cadFormula = ""
'    cadParam = ""
'    cadSelect = ""
'    numParam = 0
'
'    '===================================================
'    '============ PARAMETROS ===========================
'    indRPT = 11 'Impresion de CMR
'    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
'
'    'Nombre fichero .rpt a Imprimir
'    frmImprimir.NombreRPT = nomDocu
'
'    '===================================================
'    '================= FORMULA =========================
'    'Cadena para seleccion Nº de albaran
'    '---------------------------------------------------
'    If Text1(0).Text <> "" Then
'        'Nº palet
'        devuelve = "{" & NombreTabla & ".numalbar}=" & Val(Text1(0).Text)
'        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
'        devuelve = "numalbar = " & Val(Text1(0).Text)
'        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
'    End If
'
'    If Not HayRegParaInforme(NombreTabla, cadSelect) Then Exit Sub
'
'     With frmImprimir
'            .FormulaSeleccion = cadFormula
'            .OtrosParametros = cadParam
'            .NumeroParametros = numParam
'            .SoloImprimir = False
'            .EnvioEMail = False
'            .Opcion = 0
'            .Titulo = "Impresión CMR"
'            .Show vbModal
'    End With
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Añadir
            mnNuevo_Click

        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
        Case 5  'Buscar
            mnBuscar_Click
        Case 6  'Todos
            BotonVerTodos
        Case 8  ' Impresion de pedido
            mnImprimir_Click
'        Case 9  ' Impresion de proveedor
'            mnImprimirProv_Click
'        Case 10  ' Orden Carga
'            mnOrdenCarga_Click
'        Case 11 'CMR
'            mnCMR_Click
'        Case 12 'General Albaran
'            mnGenerarAlb_Click
'        Case 14    'Salir
'            mnSalir_Click
'        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
'            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub ActualizarToolbar(Modo As Byte, Kmodo As Byte)
'Modo: Modo antiguo
'Kmodo: Modo que se va a poner

    If (Modo = 5) And (Kmodo <> 5) Then
        'El modo antigu era modificando las lineas
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
'        Toolbar1.Buttons(5).Image = 3
'        Toolbar1.Buttons(5).ToolTipText = "Nuevo Albaran"
        '-- Modificar
        Toolbar1.Buttons(5).Image = 4
        Toolbar1.Buttons(5).ToolTipText = "Modificar Factura"
        '-- eliminar
        Toolbar1.Buttons(6).Image = 5
        Toolbar1.Buttons(6).ToolTipText = "Eliminar Factura"
    End If
    If Kmodo = 5 Then
        'Ponemos nuevos dibujitos y tal y tal
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
'        Toolbar1.Buttons(5).Image = 12
'        Toolbar1.Buttons(5).ToolTipText = "Nueva linea"
        '-- Modificar
        Toolbar1.Buttons(5).Image = 13
        Toolbar1.Buttons(5).ToolTipText = "Modificar linea factura"
        '-- eliminar
        Toolbar1.Buttons(6).Image = 14
        Toolbar1.Buttons(6).ToolTipText = "Eliminar linea factura"
    End If
End Sub
    
    
'Private Function ModificarLinea() As Boolean
''Modifica un registro en la tabla de lineas de Albaran: slialb
'Dim SQL As String
'Dim vWhere As String
'Dim b As Boolean
'
'    On Error GoTo EModificarLinea
'
'    ModificarLinea = False
'    If Data2.Recordset.EOF Then Exit Function
'
'    vWhere = ObtenerWhereCP(True)
'    vWhere = vWhere & " AND codtipoa='" & Data3.Recordset.Fields!codtipoa & "' "
'    vWhere = vWhere & " AND numalbar=" & Data3.Recordset.Fields!numalbar
'    vWhere = vWhere & " AND numlinea=" & Data2.Recordset.Fields!numlinea
'
'    If DatosOkLinea() Then
'        SQL = "UPDATE slifac SET "
'        SQL = SQL & " ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
'        SQL = SQL & "precioar = " & DBSet(txtAux(4).Text, "N") & ", "
'        SQL = SQL & "dtoline1= " & DBSet(txtAux(6).Text, "N") & ", dtoline2= " & DBSet(txtAux(7).Text, "N") & ", "
'        SQL = SQL & "importel = " & DBSet(txtAux(8).Text, "N") & ", "
'        SQL = SQL & "origpre='" & txtAux(5) & "'"
'        SQL = SQL & vWhere
'    End If
'
'    If SQL <> "" Then
'        'actualizar la factura y vencimientos
'        b = ModificarFactura(SQL)
'
'        ModificarLinea = b
'    End If
'
'EModificarLinea:
'    If Err.Number <> 0 Then
'        MuestraError Err.Number, "Modificar Lineas Factura" & vbCrLf & Err.Description
'        b = False
'    End If
'    ModificarLinea = b
'End Function


Private Sub PonerBotonCabecera(b As Boolean)
'Pone el boton de Regresar a la Cabecera si pasamos a MAntenimiento de Lineas
'o Pone los botones de Aceptar y cancelar en Insert,update o delete lineas
    On Error Resume Next

    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdRegresar.visible = b
    Me.cmdRegresar.Caption = "Cabecera"
    If b Then
        Me.lblIndicador.Caption = "Líneas " & TituloLinea
        PonerFocoBtn Me.cmdRegresar
    End If
    'Habilitar las opciones correctas del menu segun Modo
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
    DataGrid2.Enabled = Not b
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim b As Boolean
Dim Opcion As Byte
Dim Sql As String

    On Error GoTo ECargaGrid

    b = DataGrid1.Enabled
    If vDataGrid.Name = "DataGrid1" Then
        Opcion = 1
    Else
        Opcion = 2
    End If
    Sql = MontaSQLCarga(enlaza, Opcion)
    CargaGridGnral vDataGrid, vData, Sql, PrimeraVez
    
    vDataGrid.RowHeight = 270
    
    CargaGrid2 vDataGrid, vData
    vDataGrid.ScrollBars = dbgAutomatic
    
     b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
     vDataGrid.Enabled = Not b
    
   
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim tots As String
    
    On Error GoTo ECargaGrid

    Select Case vDataGrid.Name
        Case "DataGrid1" 'Pedidos_calibres
'           SQL = "SELECT numpedid, numlinea, numline1, codvarie, codcalib, nomcalib, numcajas, pesoneto
            tots = "N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux(3)|T|Variedad|1000|;"
            tots = tots & "S|txtAux(4)|T|Calibre|1000|;S|txtAux(5)|T|Nombre Calibre|4000|;S|txtAux(6)|T|Cajas|1200|;S|txtAux(8)|T|Uds|1200|;S|txtAux(7)|T|Peso Neto|1500|;"
            arregla tots, DataGrid1, Me
'            DataGrid1.Columns(11).Alignment = dbgCenter
'            DataGrid1.Columns(12).Alignment = dbgRight
'            DataGrid1.Columns(13).Alignment = dbgRight
'            DataGrid1.Columns(14).Alignment = dbgRight
                       
         Case "DataGrid2" 'pedidos_variedad
'           SQL = "SELECT numpedid, numlinea, codvarie, nomvarie1, codvarco, nomvarie2, codmarca, nommarca, codforfait, nomforfait, categori, pesobrut, totpalet, preciopro, numcajas, pesoneto
            tots = "N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux3(3)|T|Variedad Real|2000|;N||||0|;"
            tots = tots & "S|txtAux3(5)|T|Var.Comercial|2000|;N||||0|;S|txtAux3(11)|T|Marca|2300|;N||||0|;S|txtAux3(12)|T|Forfait|2330|;S|txtAux3(8)|T|Cat.|500|;"
            tots = tots & "S|txtAux3(9)|T|Peso Bruto|1400|;S|txtAux3(14)|T|Palets|800|;S|txtAux3(15)|T|Pr.Prov.|1000|;S|txtAux3(13)|T|Cajas|1000|;S|txtAux3(16)|T|Uds|1000|;S|txtAux3(10)|T|Peso Neto|1400|;"
            arregla tots, DataGrid2, Me, 350
            
            DataGrid2.Columns(3).Alignment = dbgLeft
            DataGrid2.Columns(5).Alignment = dbgLeft
            DataGrid2.Columns(7).Alignment = dbgLeft
            DataGrid2.Columns(9).Alignment = dbgLeft
                     
'            DataGrid2_RowColChange 1, 1
    End Select
    
    vDataGrid.HoldFields
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  ' Impresion de proveedor
            mnImprimirProv_Click
        Case 2  ' Orden Carga
            mnOrdenCarga_Click
        Case 3 'General Albaran
            mnGenerarAlb_Click
    End Select
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

Private Sub TxtAux_Change(Index As Integer)
    If Index = 6 And ModificaLineas = 2 Then 'Precio y Modo Borrar Lineas
        txtAux(5).Text = "M"
    End If
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)

    'Quitar espacios en blanco
    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    Select Case Index
        Case 4 'Precio
             'Tipo 2: Decimal(10,4)
             If txtAux(Index).Text <> "" Then PonerFormatoDecimal txtAux(Index), 2
            
        Case 6, 7 'Descuentos
            PonerFormatoDecimal txtAux(Index), 4 'Tipo 4: Decimal(4,2)
            If Index = 7 Then PonerFoco Me.Text2(16)
            
        Case 8 'Importe Linea
            PonerFormatoDecimal txtAux(Index), 3 'Tipo 3: Decimal(10,2)
    End Select
    
'    If (Index = 3 Or Index = 4 Or Index = 6 Or Index = 7) Then 'Cant., Precio, Dto1, Dto2
'        If txtAux(1).Text = "" Then Exit Sub
'        txtAux(8).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(6).Text, txtAux(7).Text, vParamAplic.TipoDtos)
'        PonerFormatoDecimal txtAux(8), 1
'    End If
End Sub


Private Sub BotonMtoLineas(numTab As Integer, Cad As String)
    If Me.DataGrid1.visible Then
        If Me.Data2.Recordset.RecordCount < 1 Then
            MsgBox "El Pedido no tiene lineas.", vbInformation
            Exit Sub
        End If
        TituloLinea = Cad
    End If
    ModificaLineas = 0
    PonerModo 5
    PonerBotonCabecera True
End Sub


Private Function Eliminar() As Boolean
Dim Sql As String, LEtra As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
    
    On Error GoTo FinEliminar

    b = False
    If Data1.Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        

    'Eliminar en tablas de factura de Ariges
    '------------------------------------------
    Sql = " " & ObtenerWhereCP(True)

    'Lineas de calibres (pedidos_calibre)
    conn.Execute "Delete from pedidos_calibre " & Sql

    'Lineas de variedades
    conn.Execute "Delete from pedidos_variedad " & Sql
    
    'Cabecera de palets (pedidos)
    conn.Execute "Delete from " & NombreTabla & Sql
    
    'Decrementar contador si borramos el ult. palet
    Set vTipoMov = New CTiposMov
    vTipoMov.DevolverContador CodTipoMov, Val(Text1(0).Text)
    Set vTipoMov = Nothing
    
    b = True
    
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Pedido", Err.Description
        b = False
    End If
    If Not b Then
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function

Private Function EliminarLinea() As Boolean
Dim Sql As String, LEtra As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
    
    On Error GoTo FinEliminar

    b = False
    If Data3.Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        

    'Eliminar en tablas de paltes_variedad y pedidos_calibre
    '------------------------------------------
    Sql = " where numpedid = " & Data3.Recordset.Fields(0)
    Sql = Sql & " and numlinea = " & Data3.Recordset.Fields(1)

    'Lineas de calibres (pedidos_calibre)
    conn.Execute "Delete from pedidos_calibre " & Sql

    'Lineas de variedades
    conn.Execute "Delete from pedidos_variedad " & Sql
    
    b = True
    
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Variedad del Pedido", Err.Description
        b = False
    End If
    If Not b Then
        conn.RollbackTrans
        EliminarLinea = False
    Else
        conn.CommitTrans
        EliminarLinea = True
    End If
End Function

Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next

    CargaGrid DataGrid2, Data3, False
    CargaGrid DataGrid1, Data2, False
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP(False) & ")"
        If SituarDataMULTI(Data1, vWhere, Indicador) Then
             PonerModo 2
             lblIndicador.Caption = Indicador
        Else
             LimpiarCampos
             'Poner los grid sin apuntar a nada
             LimpiarDataGrids
             PonerModo 0
         End If
    Else
        'El Data esta vacio, desde el modo de inicio se pulsa Insertar
        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Function ObtenerWhereCP(conWhere As Boolean) As String
Dim Sql As String

    On Error Resume Next
    
    Sql = " numpedid= " & Text1(0).Text
    If conWhere Then Sql = " WHERE " & Sql
    ObtenerWhereCP = Sql
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo cadena WHERE.", Err.Description
End Function


Private Function MontaSQLCarga(enlaza As Boolean, Opcion As Byte) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim Sql As String
    
    If Opcion = 1 Then
        Sql = "SELECT numpedid, numlinea, numline1, pedidos_calibre.codvarie, pedidos_calibre.codcalib, nomcalib, numcajas, unidades, pesoneto "
        Sql = Sql & " FROM pedidos_calibre, calibres WHERE pedidos_calibre.codvarie = calibres.codvarie and "
        Sql = Sql & " pedidos_calibre.codcalib = calibres.codcalib "
    ElseIf Opcion = 2 Then
        Sql = "SELECT pedidos_variedad.numpedid, numlinea, pedidos_variedad.codvarie, a.nomvarie as nomvarie1, pedidos_variedad.codvarco, "
        Sql = Sql & " b.nomvarie as nomvarie2, pedidos_variedad.codmarca, marcas.nommarca, pedidos_variedad.codforfait, forfaits.nomconfe, "
        Sql = Sql & " categori, pesobrut, totpalet, preciopro, numcajas, unidades, pesoneto "
        Sql = Sql & " FROM pedidos_variedad, variedades a, variedades b, marcas, forfaits " 'lineas de variedades del pedido
        Sql = Sql & " WHERE pedidos_variedad.codvarie = a.codvarie "
        Sql = Sql & " and pedidos_variedad.codvarco = b.codvarie"
        Sql = Sql & " and pedidos_variedad.codmarca = marcas.codmarca "
        Sql = Sql & " and pedidos_variedad.codforfait = forfaits.codforfait "
    End If
    
    If enlaza Then
        Sql = Sql & " and " & ObtenerWhereCP(False)
        If Opcion = 1 Then Sql = Sql & " AND numlinea=" & Data3.Recordset.Fields!NumLinea
    Else
        Sql = Sql & " and numpedid = -1"
    End If
    Sql = Sql & " ORDER BY numpedid"
    If Opcion = 1 Then Sql = Sql & ", numlinea "
    MontaSQLCarga = Sql
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean, bAux As Boolean
Dim i As Integer

        b = ((Modo = 2) Or (Modo = 0)) And (hcoCodMovim = "")  'Or (Modo = 5 And ModificaLineas = 0)
        'Buscar
        Toolbar1.Buttons(5).Enabled = b
        Me.mnBuscar.Enabled = b
        'Vore Tots
        Toolbar1.Buttons(6).Enabled = b
        Me.mnVerTodos.Enabled = b
        'Añadir
        Toolbar1.Buttons(1).Enabled = b
        Me.mnModificar.Enabled = b
        
        
        b = (Modo = 2 And Data1.Recordset.RecordCount > 0)
        'Modificar
        Toolbar1.Buttons(2).Enabled = b
        Me.mnModificar.Enabled = b
        'eliminar
        Toolbar1.Buttons(3).Enabled = (Modo = 2)
        Me.mnEliminar.Enabled = (Modo = 2)
        'Impresión de pedido
        Toolbar1.Buttons(8).Enabled = (Modo = 2 And Data1.Recordset.RecordCount > 0) Or (hcoCodMovim <> "")
        Me.mnImprimir.Enabled = (Modo = 2 And Data1.Recordset.RecordCount > 0) Or (hcoCodMovim <> "")
        
        'Impresión de pedido proveedor
        Toolbar5.Buttons(1).visible = (vParamAplic.Cooperativa = 15)
        Me.mnImprimirProv.visible = (vParamAplic.Cooperativa = 15)
        Toolbar5.Buttons(1).Enabled = ((Modo = 2 And Data1.Recordset.RecordCount > 0) Or (hcoCodMovim <> "")) And vParamAplic.Cooperativa = 15
        Me.mnImprimirProv.Enabled = ((Modo = 2 And Data1.Recordset.RecordCount > 0) Or (hcoCodMovim <> "")) And vParamAplic.Cooperativa = 15
        
        'Orden de Carga
        Toolbar5.Buttons(2).Enabled = b
        Me.mnOrdenCarga.Enabled = b
'[Monica]06/06/2018: el generar cmr esta en albaranes
'        'Generar CMR
'        Toolbar1.Buttons(11).Enabled = b
'        Me.mnCMR.Enabled = b
        'Generar Albaran
        Toolbar5.Buttons(3).Enabled = b
        Me.mnGenerarAlb.Enabled = b
        

    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
'++monica: si insertamos lo he quitado
'    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
    b = (Modo = 4 Or Modo = 2) And (hcoCodMovim = "")
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.Data3.Recordset.RecordCount > 0)
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i


End Sub


Private Sub BotonImprimir(Opcion As Byte)
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadselect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String

    If Text1(0).Text = "" Then
        MsgBox "Debe seleccionar un Pedido para Imprimir.", vbInformation
        Exit Sub
    End If
    
    cadFormula = ""
    cadParam = ""
    cadselect = ""
    numParam = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    If Opcion = 0 Then
        indRPT = 7 'Impresion de Pedido
    Else
        indRPT = 104 ' impresion de pedido para proveedor
    End If
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
      
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de pedido
    '---------------------------------------------------
    If Text1(0).Text <> "" Then
        'Nº palet
        devuelve = "{" & NombreTabla & ".numpedid}=" & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        devuelve = "numpedid = " & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadselect, devuelve) Then Exit Sub
    End If
    
    If Not HayRegParaInforme(NombreTabla, cadselect) Then Exit Sub
     
     With frmImprimir
          '[Monica]24/01/2012: añadido la siguientes 3 lineas para el envio por el outlook
            .outClaveNombreArchiv = Format(Text1(0).Text, "0000000")
            .outCodigoCliProv = Text1(3).Text
            '[Monica]06/05/2015: destino para sacar email
            .outCodigoDestino = Text1(4).Text
            .outTipoDocumento = 3
            
            
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 0
            If Opcion = 0 Then
                .Titulo = "Impresión de Pedido"
            Else
                .Titulo = "Impresión de Pedido para Proveedor"
            End If
            .ConSubInforme = True
            .Show vbModal
    End With
End Sub


Private Sub TxtAux3_GotFocus(Index As Integer)
    ConseguirFoco txtAux3(Index), Modo
End Sub

Private Sub TxtAux3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index <> 0 And KeyCode <> 38 Then KEYdown KeyCode
End Sub

Private Sub TxtAux3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub TxtAux3_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux3(Index), Modo) Then Exit Sub
End Sub


Private Function ObtenerSelFactura() As String
Dim Cad As String
Dim Rs As ADODB.Recordset

    On Error Resume Next

    Cad = ""
    '******************************************************
    'laura: esto se puede comentar, ya no hay movimiento FTI en la smoval
    If hcoCodTipoM = "FTI" Then
        'no hay albaran directamente va a factura de ticket
        
        'ver si lo encontramos como factura: codtipom, numfactu,fecfactu
        Cad = "SELECT COUNT(*) FROM scafac "
        Cad = Cad & " WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
        If RegistrosAListar(Cad) > 0 Then
            Cad = " WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
        Else
            Cad = ""
        End If
    End If
    '******************************************************
        
    If Cad = "" Then
        'En la smoval estaba e mov. de ALbaran
        Cad = "SELECT codtipom,numfactu,fecfactu FROM scafac1 "
        Cad = Cad & " WHERE codtipoa=" & DBSet(hcoCodTipoM, "T") & " AND numalbar=" & hcoCodMovim & " AND fechaalb=" & DBSet(hcoFechaMov, "F")
        
        Set Rs = New ADODB.Recordset
        Rs.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then 'where para la factura
            Cad = " WHERE codtipom='" & Rs!codTipoM & "' AND numfactu= " & Rs!NumFactu & " AND fecfactu=" & DBSet(Rs!FecFactu, "F")
        Else
            Cad = " WHERE numfactu=-1"
        End If
        Rs.Close
        Set Rs = Nothing
    End If
    ObtenerSelFactura = Cad
End Function




Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim i As Byte
    
    Combo1(0).Clear
    
    Combo1(0).AddItem "Original"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    
    Combo1(0).AddItem "Modificado"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    
    Combo1(0).AddItem "Anulado"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    
End Sub

Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim Sql As String

    On Error GoTo EInsertarCab
    
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
        Sql = CadenaInsertarDesdeForm(Me)
        If Sql <> "" Then
            If InsertarOferta(Sql, vTipoMov) Then
                CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                PonerCadenaBusqueda
                PonerModo 2
                'Ponerse en Modo Insertar Lineas
'                BotonMtoLineas 0, "Variedades"
'                BotonAnyadirLinea
                Set frmLPed = New frmVtasLinPedidos
                
                frmLPed.ModoExt = 3
                frmLPed.Pedido = CLng(Text1(0).Text)
                frmLPed.Show vbModal
                
                Set frmLPed = Nothing
            End If
        End If
        Text1(0).Text = Format(Text1(0).Text, "0000000")
    End If
    Set vTipoMov = Nothing
    
EInsertarCab:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Function InsertarOferta(vSQL As String, vTipoMov As CTiposMov) As Boolean
Dim MenError As String
Dim bol As Boolean, Existe As Boolean
Dim cambiaSQL As Boolean
Dim devuelve As String

    On Error GoTo EInsertarOferta
    
    bol = True
    
    cambiaSQL = False
    'Comprobar si mientras tanto se incremento el contador de Pedidos
    'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
    Do
        devuelve = DevuelveDesdeBDNew(cAgro, NombreTabla, "numpedid", "numpedid", Text1(0).Text, "N")
        If devuelve <> "" Then
            'Ya existe el contador incrementarlo
            Existe = True
            vTipoMov.IncrementarContador (CodTipoMov)
            Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
            cambiaSQL = True
        Else
            Existe = False
        End If
    Loop Until Not Existe
    If cambiaSQL Then vSQL = CadenaInsertarDesdeForm(Me)
    
    
    'Aqui empieza transaccion
    conn.BeginTrans
    MenError = "Error al insertar en la tabla Cabecera de Pedidos (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
'    'Actualizar los datos del cliente si es de varios
'    If EsDeVarios Then
'        'Si es cliente de varios actualizar datos cliente en tabla:sclvar
'        MenError = "Modificando datos cliente varios"
'        bol = ActualizarClienteVarios(Text1(4).Text, Text1(6).Text)
'    End If
'
'    If bol Then
'        'Actualizar el campo fechamov (ult. movimiento) de la tabla de clientes (sclien)
'        MenError = "Actualizando Fecha Movimiento del Cliente."
'        bol = ActualizarFecMovCliente
        
        MenError = "Error al actualizar el contador del Pedido."
    '    bol = vTipoMov.IncrementarContador("REG")
        vTipoMov.IncrementarContador (CodTipoMov)
'    End If
    
EInsertarOferta:
        If Err.Number <> 0 Then
            MenError = "Insertando Pedido." & vbCrLf & "----------------------------" & vbCrLf & MenError
            MuestraError Err.Number, MenError, Err.Description
            bol = False
        End If
        If bol Then
            conn.CommitTrans
            InsertarOferta = True
        Else
            conn.RollbackTrans
            InsertarOferta = False
        End If
End Function

Private Sub MostrarCadena(clien As String, desti As String)
Dim Sql As String

    If clien = "" Or desti = "" Then Exit Sub

    Sql = DevuelveDesdeBDNew(cAgro, "destinos", "codcaden", "codclien", clien, "N", , "coddesti", desti, "N")
    If Sql <> "" Then
        Label3.Caption = DevuelveDesdeBDNew(cAgro, "cadenas", "nomcaden", "codcaden", Sql, "N")
    Else
        Label3.Caption = ""
    End If

End Sub


Private Function InicializarCStockAlbar(ByRef vCStock As CStock, TipoM As String, Optional NumLinea As String, Optional ByRef Rs As ADODB.Recordset) As Boolean
'Para comprobar stock al pasar de Pedido a Albaran de Venta
On Error Resume Next
    
'    vCStock.tipoMov = TipoM
'    vCStock.DetaMov = "ALV"
'    vCStock.Trabajador = CInt(Text1(4).Text) 'En codigope ponemos el Cliente
'    vCStock.Documento = Text1(0).Text
'    vCStock.codArtic = Rs!codArtic
'    vCStock.codAlmac = CInt(Rs!codAlmac)
'
'    If AlbDePalet Then
'        vCStock.Cantidad = CSng(Rs!Cantidad)
'        If Rs.Fields.Count > 3 Then 'Si no se selecciona el campo importe de la tabla es que solo vamos a comprobar stock y no se necesita
'            vCStock.Importe = CCur(Rs!ImporteL)
'        End If
'    Else
'        vCStock.Cantidad = CSng(Rs!servidas)
'        'Si se va a Insertar en alguna linea obtener el importe
'        'Si solo vamos a comprobar stock no hace falta el importe
'        If Rs.Fields.Count > 4 Then
'            vCStock.Importe = CCur(CalcularImporte(Rs!servidas, Rs!precioar, Rs!dtoline1, Rs!dtoline2, vParamAplic.TipoDtos))
'        End If
'    End If
'
'    vCStock.LineaDocu = CInt(ComprobarCero(numlinea))
'
'    If Err.Number <> 0 Then
'        MsgBox "No se han podido inicializar la clase para actualizar Stock", vbExclamation
'        InicializarCStockAlbar = False
'    Else
'        InicializarCStockAlbar = True
'    End If
End Function


Private Sub GenerarAlbaran()
Dim numPed As Long 'Nº Pedido
Dim NumAlb As String 'Nº Albaran
Dim Sql As String
Dim frmAlb As frmVtasAlbaranes

    'Pedir: fecha de albaran y si se quiere imprimir
    CadenaSQL = ""
    Set frmList = New frmListadoPed
    frmList.Opcionlistado = 43
    frmList.NumCod = CodTipoMov
    frmList.Show vbModal
    Set frmList = Nothing
    If CadenaSQL = "" Then Exit Sub
    

    NumRegElim = Data1.Recordset.AbsolutePosition
    numPed = Data1.Recordset!numpedid

    If PasarPedidoAAlbaran(CadenaSQL, NumAlb) Then
        espera 0.4

        MsgBox "El Pedido de Venta Nº: " & Format(numPed, "0000000") & vbCrLf & vbCrLf & "ha generado el Albaran Nº: " & Format(NumAlb, "0000000")
        
        Set frmAlb = New frmVtasAlbaranes
        frmAlb.NumAlbar = NumAlb
        frmAlb.Show vbModal
        
        PosicionarData
        PonerCampos
        If Not Data2.Recordset.EOF Then ' 23/11/2009 solo si no tiene lineas no cargamos el grid
            CargaGrid DataGrid1, Data2, True
        End If
        Screen.MousePointer = vbDefault

        'Imprimer albaran si se solicitó
        If ImprimeAlb Then ImprimirAlbaran 45, NumAlb

    End If
End Sub

Private Function PasarPedidoAAlbaran(vSQL As String, NumAlb As String) As Boolean
'IN -> vSQL: cadena para el Select con los datos obtenidos en frmList
'OUT -> numAlb: Nº de Albaran de Venta que se ha insertado
Dim bol As Boolean
Dim MenError As String
Dim devuelve As String
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim cCli As CCliente

    On Error GoTo EGenPedido

    bol = False
        
    'Aqui empieza transaccion
    conn.BeginTrans
    
    'Insertar en tablas de Albaranes el Pedido (albaran, albaran_variedad, albaran_calibre)
    bol = InsertarAlbaran(vSQL, MenError, NumAlb)
    
    If bol Then bol = ActualizarCabPedido(NumAlb)
    
    
EGenPedido:
    If Err.Number <> 0 Or Not bol Then
        MenError = "Pasando Pedido a Albaran." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        bol = False
    End If
    If bol Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
    End If
    PasarPedidoAAlbaran = bol
End Function

Private Function InsertarAlbaran(vSQL As String, MenError As String, NumAlb As String) As Boolean
'Devuelve el mensaje de error si se produce
Dim bol As Boolean, Existe As Boolean
Dim devuelve As String
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim codTipoM As String

    On Error GoTo EInsertarAlbaran
    
    bol = False
    InsertarAlbaran = bol
    
    'Obtener el Contador de ALBARAN
    '[Monica]02/07/2012: antes cogiamos el tipo de movimiento de parametros ahora lo cogemos de clientes
    'codTipoM = vParamAplic.CodTipomAlb ' "ALV"
    
    codTipoM = DevuelveValor("select codtipalb from clientes where codclien = " & DBSet(Text1(3).Text, "N"))
    
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(codTipoM) Then
        'Comprobar si mientras tanto se incremento el contador de albaranes
        Do
            NumAlb = vTipoMov.ConseguirContador(codTipoM)
            devuelve = DevuelveDesdeBDNew(cAgro, "albaran", "numalbar", "numalbar", NumAlb, "N")
            If devuelve <> "" Then
                'Ya existe el contador incrementarlo
                Existe = True
                vTipoMov.IncrementarContador (codTipoM)
                NumAlb = vTipoMov.ConseguirContador(codTipoM)
            Else
                Existe = False
            End If
        Loop Until Not Existe
            
    Else 'No existe el tipo de Movimiento
        Set vTipoMov = Nothing
        Exit Function
    End If
    
    'Acabar la sql con el contador seleccionado
    devuelve = vSQL
    vSQL = "INSERT INTO albaran (numalbar,fechaalb,codclien,coddesti,codtrans,matriveh,matrirem,"
    vSQL = vSQL & "refclien,codtimer,totpalet,portespre,nrocontra,nroactas,numpedid,fechaped,observac,"
    vSQL = vSQL & "pasaridoc, codalmac) "
    vSQL = vSQL & "SELECT " & NumAlb & " as numalbar, " & devuelve
    vSQL = vSQL & " FROM " & NombreTabla & " WHERE numpedid=" & Text1(0).Text

    'Insertar Cabecera
    MenError = "Error al insertar en la tabla Cabecera de Albaranes (albaranes)."
    conn.Execute vSQL, , adCmdText
    
    '[Monica]02/07/2012: cogemos el tipo de movimiento de parametros para las inserciones en almacen
    codTipoM = vParamAplic.CodTipomAlb ' "ALV"
    
    'Insertar Lineas (albaran_variedad, albaran_calibre, albaran_costes)
    MenError = "Error al insertar en la tabla Lineas de Albaran (albaran_variedad)."
    If Not InsertarLineasAlbaran(codTipoM, MenError, NumAlb) Then Exit Function
    
    'Insertar Lineas (albaran_palets)
    MenError = "Error al insertar en la tabla Lineas de Albaran (albaran_palets)."
    If Not InsertarPaletsAlbaran(Text1(0), NumAlb) Then Exit Function
    
    '[Monica]02/07/2012: antes cogiamos el tipo de movimiento de parametros ahora lo cogemos de clientes
    codTipoM = DevuelveValor("select codtipalb from clientes where codclien = " & DBSet(Text1(3).Text, "N"))
    
    
    MenError = "Error al actualizar el contador del ALbaran."
'    bol = vTipoMov.IncrementarContador("REG")
    vTipoMov.IncrementarContador (codTipoM)
    Set vTipoMov = Nothing
    bol = True
    
EInsertarAlbaran:
        If Err.Number <> 0 Then bol = False
        InsertarAlbaran = bol
End Function

Private Function InsertarLineasAlbaran(TipoM As String, MenError As String, NumAlb As String) As Boolean
'Inserta en la tabla de lineas de albaran (slialb)
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset

Dim ImpLinea As String

Dim NumLin As Integer
Dim NumLin1 As Integer
Dim b As Boolean

Dim NumCajas As Long
Dim PesoBruto As Long
Dim Pesoneto As Long
Dim VariedadAnt As Integer
Dim VarComAnt As Integer
Dim MarcaAnt As Integer
Dim ForfaitAnt As String
Dim CategoriAnt As String

    On Error GoTo eInsertarLineasAlbaran

    If AlbDePalet Then
        If Not vParamAplic.PaseAlbarAgrupCalib Then
            b = InsertarVariedades(MenError, NumAlb)
        Else
            b = InsertarVariedadesSinAgrupar(MenError, NumAlb)
        End If
    Else
        ' copiamos el pedido tal cual
        Sql = "select * from pedidos_variedad "
        Sql = Sql & " WHERE " & ObtenerWhereCP(False)
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

        b = True
        While Not Rs.EOF And b 'Para cada linea de pedido insertar una de albaran
'albaran_variedad:numalbar,numlinea,codvarie,codvarco,codforfait,codmarca,categori,totpalet,numcajas,pesobrut,pesoneto,preciopro,preciodef,codincid,impcomis,observac
'pedidos_variedad:numpedid,numlinea,codvarie,codvarco,codforfait,codmarca,categori,totpalet,numcajas,pesobrut,pesoneto,preciopro
            Sql = "INSERT INTO albaran_variedad (numalbar,numlinea,codvarie,codvarco,codforfait,codmarca,categori, "
            Sql = Sql & "totpalet,numcajas,pesobrut,pesoneto,preciopro,preciodef,codincid,impcomis,observac, unidades, codpalet) "
            Sql = Sql & " VALUES(" & NumAlb & ", " & Rs!NumLinea & " , "
            Sql = Sql & DBLet(Rs!codvarie, "N") & ", " & DBSet(Rs!codvarco, "N") & ", " & DBSet(Rs!codforfait, "T") & ", " & DBSet(Rs!Codmarca, "N") & ", "
            Sql = Sql & DBSet(Rs!categori, "T") & ", " & DBSet(Rs!TotPalet, "N") & ", " & DBSet(Rs!NumCajas, "N") & ", " & DBSet(Rs!pesobrut, "N") & ", "
            Sql = Sql & DBSet(Rs!Pesoneto, "N") & ", " & DBSet(Rs!preciopro, "N") & ", " & ValorNulo & "," & DBSet(Incidencia, "N") & "," & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & DBSet(Rs!Unidades, "N") & "," & DBSet(Rs!CodPalet, "N") & " )"
            MenError = "Error al insertar en la tabla Lineas de Albaran (albaran_variedad)."
            conn.Execute Sql
            
'albaran_calibre:numalbar,numlinea,numline1,codvarie,codcalib,numcajas,pesobrut,pesoneto
'pedidos_calibre:numpedid,numlinea,numline1,codvarie,codcalib,numcajas,pesoneto
            Sql = "INSERT INTO albaran_calibre (numalbar,numlinea,numline1,codvarie,codcalib,numcajas,pesobrut,pesoneto, unidades) "
            Sql = Sql & " select " & NumAlb & ", numlinea, numline1,codvarie,codcalib,numcajas," & ValorNulo & ", pesoneto, unidades "
            Sql = Sql & " from pedidos_calibre where numpedid = " & Text1(0).Text
            Sql = Sql & " and numlinea = " & Rs!NumLinea
            
            MenError = "Error al insertar en la tabla Calibres de Albaran (albaran_calibre)."
            conn.Execute Sql
            
            MenError = "Error al Actualizar Costes."
            b = ActualizarCostes(CLng(NumAlb), DBLet(Rs!NumLinea, "N"), True, DBLet(Rs!codforfait, "T"), DBLet(Rs!CodPalet, "N"))
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
    End If


eInsertarLineasAlbaran:
    If Err.Number <> 0 Or Not b Then
        InsertarLineasAlbaran = False
    Else
        InsertarLineasAlbaran = True
    End If
End Function


Private Function InsertarVariedades(MenError As String, NumAlb As String) As Boolean
'Inserta en la tabla de lineas de albaran (slialb)
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset

Dim ImpLinea As String

Dim NumLin As Integer
Dim NumLin1 As Integer
Dim b As Boolean

Dim NumCajas As Long
Dim PesoBruto As Long
Dim Pesoneto As Long

Dim PesoBrutoLin As Long
Dim PesoNetoLin As Long

Dim VariedadAnt As Integer
Dim VarComAnt As Integer
Dim MarcaAnt As Integer
Dim ForfaitAnt As String
Dim CategoriAnt As String

    On Error GoTo eInsertarVariedades

        'Insertar en la tabla de Pedido, los registros seleccionados de la tabla de Palets
        
        Sql = ""
        Sql = "SELECT palets_variedad.codvarie, palets_variedad.codvarco, palets_variedad.codmarca, "
        Sql = Sql & " palets_variedad.codforfait, palets_variedad.categori, palets.codpalet, "
        Sql = Sql & " sum(pesobrut),sum(pesoneto),sum(numcajas) "
        Sql = Sql & " FROM palets, palets_variedad WHERE palets.numpedid=" & DBSet(Text1(0).Text, "N")
        Sql = Sql & " and palets.numpalet = palets_variedad.numpalet "
        Sql = Sql & " GROUP BY 1,2,3,4,5,6"
        Sql = Sql & " ORDER BY 1,2,3,4,5,6"
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        NumLin = 0
        b = True
        While Not Rs.EOF And b 'agrupamos las lineas de palets_variedad
            NumLin = NumLin + 1
            
            Sql = "INSERT INTO albaran_variedad(numalbar,numlinea,codvarie,codvarco,codforfait,codmarca,"
            Sql = Sql & "categori,totpalet,numcajas,pesobrut,pesoneto,preciopro,preciodef,"
            Sql = Sql & "codincid,impcomis,observac, codpalet) VALUES "
            Sql = Sql & "(" & NumAlb & "," & NumLin & "," & DBSet(Rs.Fields(0), "N") & "," & DBSet(Rs.Fields(1), "N") & ","
            Sql = Sql & DBSet(Rs.Fields(3), "T") & "," & DBSet(Rs.Fields(2), "N") & ","
            Sql = Sql & DBSet(Rs.Fields(4), "T") & "," & ValorNulo & ","
            Sql = Sql & DBSet(Rs.Fields(8), "N") & "," & DBSet(Rs.Fields(6), "N") & ","
            Sql = Sql & DBSet(Rs.Fields(7), "N") & "," & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & DBSet(Incidencia, "N") & "," & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & DBSet(Rs.Fields(5), "N") & ")"
    
            conn.Execute Sql
            
            ' en cuantos palets aparece esta linea
            Sql = "select count(distinct palets.numpalet) from palets, palets_variedad where palets.numpedid = " & DBSet(Text1(0).Text, "N")
            Sql = Sql & " and palets_variedad.codvarie = " & DBSet(Rs.Fields(0).Value, "N")
            Sql = Sql & " and palets_variedad.codvarco = " & DBSet(Rs.Fields(1).Value, "N")
            Sql = Sql & " and palets_variedad.codforfait = " & DBSet(Rs.Fields(3).Value, "T")
            Sql = Sql & " and palets_variedad.codmarca = " & DBSet(Rs.Fields(2).Value, "N")
            
            If DBSet(Rs.Fields(4).Value, "T") = ValorNulo Then
                Sql = Sql & " and palets_variedad.categori is null "
            Else
                Sql = Sql & " and palets_variedad.categori = " & DBSet(Rs.Fields(4).Value, "T")
            End If
            '[Monica] 15/06/2010 añadido costes paletizacion
            Sql = Sql & " and palets.codpalet = " & DBSet(Rs.Fields(5).Value, "N")
            
            Sql = Sql & " and palets.numpalet = palets_variedad.numpalet "
            
            Set Rs1 = New ADODB.Recordset
            Rs1.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not Rs1.EOF Then
                If DBLet(Rs1.Fields(0).Value, "N") <> 0 Then
                    Sql = "update albaran_variedad set totpalet = " & DBSet(Rs1.Fields(0).Value, "N")
                    Sql = Sql & " where numalbar = " & DBSet(NumAlb, "N")
                    Sql = Sql & " and numlinea = " & DBSet(NumLin, "N")
                    
                    conn.Execute Sql
                End If
            End If
            Set Rs1 = Nothing
            
            
            'Insertar en la tabla de albaranes, los registros seleccionados de la tabla de Palets
            Sql = ""
            Sql = "SELECT palets_variedad.codvarie, palets_variedad.codvarco, palets_variedad.codmarca, "
            Sql = Sql & " palets_variedad.codforfait, palets_variedad.categori, palets_calibre.codcalib, "
            Sql = Sql & " sum(pesobrut),sum(pesoneto),sum(palets_variedad.numcajas),sum(palets_calibre.numcajas) "
            Sql = Sql & " FROM palets, palets_variedad, palets_calibre WHERE palets.numpedid=" & Text1(0).Text
            Sql = Sql & " and palets_variedad.codvarie = " & DBSet(Rs.Fields(0).Value, "N")
            Sql = Sql & " and palets_variedad.codvarco = " & DBSet(Rs.Fields(1).Value, "N")
            Sql = Sql & " and palets_variedad.codforfait = " & DBSet(Rs.Fields(3).Value, "T")
            Sql = Sql & " and palets_variedad.codmarca = " & DBSet(Rs.Fields(2).Value, "N")
            If DBSet(Rs.Fields(4).Value, "T") = ValorNulo Then
                Sql = Sql & " and palets_variedad.categori is null "
            Else
                Sql = Sql & " and palets_variedad.categori = " & DBSet(Rs.Fields(4).Value, "T")
            End If
            '[Monica] 15/06/2010 añadido costes paletizacion
            Sql = Sql & " and palets.codpalet = " & DBSet(Rs.Fields(5).Value, "N")
            
            Sql = Sql & " and palets.numpalet = palets_variedad.numpalet "
            Sql = Sql & " and palets_variedad.numpalet = palets_calibre.numpalet "
            Sql = Sql & " and palets_variedad.numlinea = palets_calibre.numlinea "
            Sql = Sql & " GROUP BY 1,2,3,4,5,6"
            Sql = Sql & " ORDER BY 1,2,3,4,5,6"
            
            Set Rs1 = New ADODB.Recordset
            Rs1.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            PesoBruto = 0
            Pesoneto = 0
            NumCajas = 0
            NumLin1 = 0
            While Not Rs1.EOF
                NumLin1 = NumLin1 + 1
'22-09-2008
'                PesoBruto = PesoBruto + DBLet(Rs1.Fields(6).Value, "N")
'                PesoNeto = PesoNeto + DBLet(Rs1.Fields(7).Value, "N")
'                NumCajas = NumCajas + DBLet(Rs1.Fields(8).Value, "N")
                PesoBrutoLin = 0
                If DBLet(Rs1.Fields(8).Value, "N") <> 0 Then
                    PesoBrutoLin = Round2(DBLet(Rs1.Fields(6).Value, "N") * DBLet(Rs1.Fields(9).Value, "N") / DBLet(Rs1.Fields(8).Value, "N"), 0)
                End If
                PesoBruto = PesoBruto + PesoBrutoLin
                PesoNetoLin = 0
                If DBLet(Rs1.Fields(8).Value, "N") <> 0 Then
                    PesoNetoLin = Round2(DBLet(Rs1.Fields(7).Value, "N") * DBLet(Rs1.Fields(9).Value, "N") / DBLet(Rs1.Fields(8).Value, "N"), 0)
                End If
                Pesoneto = Pesoneto + PesoNetoLin
                
                ' insertamos en la tabla albaran calibres
                Sql = "INSERT INTO albaran_calibre(numalbar, numlinea, numline1, codvarie, codcalib, "
                Sql = Sql & "numcajas, pesobrut, pesoneto) VALUES ("
                Sql = Sql & NumAlb & "," & NumLin & "," & NumLin1 & "," & DBSet(Rs1.Fields(0).Value, "N") & ","
                Sql = Sql & DBSet(Rs1.Fields(5).Value, "N") & "," & DBSet(Rs1.Fields(9).Value, "N") & ","
                Sql = Sql & DBSet(PesoBrutoLin, "N") & "," & DBSet(PesoNetoLin, "N") & ")"
                
                MenError = "Error al insertar en la tabla Calibres de Albaran (albaran_calibre)."
                conn.Execute Sql
                
                Rs1.MoveNext
            Wend
            
            Set Rs1 = Nothing
            
            ' redondeamos en la ultima linea de calibres
            Sql = "select sum(pesobrut),sum(pesoneto) "
            Sql = Sql & " FROM palets, palets_variedad WHERE palets.numpedid=" & Text1(0).Text
            Sql = Sql & " and palets_variedad.codvarie = " & DBSet(Rs.Fields(0).Value, "N")
            Sql = Sql & " and palets_variedad.codvarco = " & DBSet(Rs.Fields(1).Value, "N")
            Sql = Sql & " and palets_variedad.codforfait = " & DBSet(Rs.Fields(3).Value, "T")
            Sql = Sql & " and palets_variedad.codmarca = " & DBSet(Rs.Fields(2).Value, "N")
            If DBSet(Rs.Fields(4).Value, "T") = ValorNulo Then
                Sql = Sql & " and palets_variedad.categori is null "
            Else
                Sql = Sql & " and palets_variedad.categori = " & DBSet(Rs.Fields(4).Value, "T")
            End If
            '[Monica] 15/06/2010 añadido costes paletizacion
            Sql = Sql & " and palets.codpalet = " & DBSet(Rs.Fields(5).Value, "N")
            
            Sql = Sql & " and palets.numpalet = palets_variedad.numpalet "
            
            Set Rs1 = New ADODB.Recordset
            Rs1.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not Rs1.EOF Then
                If DBLet(Rs1.Fields(0).Value, "N") <> PesoBruto Or DBLet(Rs1.Fields(1).Value, "N") <> Pesoneto Then
                    Sql = "update albaran_calibre set pesobrut = pesobrut + " & DBLet(Rs1.Fields(0).Value, "N") - PesoBruto
                    Sql = Sql & ", pesoneto = pesoneto + " & DBLet(Rs1.Fields(1).Value, "N") - Pesoneto
                    Sql = Sql & " where albaran_calibre.numalbar  = " & NumAlb
                    Sql = Sql & " and albaran_calibre.numlinea = " & NumLin
                    Sql = Sql & " and albaran_calibre.numline1 = " & NumLin1
                    
                    conn.Execute Sql
                End If
            End If
            Set Rs1 = Nothing
             
'22-09-2008
'            ' actualizamos el numero de cajas, peso bruto y peso neto
'            sql = "UPDATE albaran_variedad SET numcajas = " & DBLet(NumCajas, "N")
'            sql = sql & ", pesobrut = " & DBSet(PesoBruto, "N")
'            sql = sql & ", pesoneto = " & DBSet(PesoNeto, "N")
'            sql = sql & " where numalbar = " & DBSet(NumAlb, "N")
'            sql = sql & " and numlinea = " & DBSet(NumLin, "N")
'
'            conn.Execute sql
            
            MenError = "Error al Actualizar Costes."
            If b Then b = ActualizarCostes(CLng(NumAlb), NumLin, True, DBLet(Rs!codforfait, "T"), DBLet(Rs!CodPalet, "N"))
            
            Rs.MoveNext
        Wend
        Set Rs = Nothing
        
eInsertarVariedades:
    If Err.Number <> 0 Or Not b Then
        InsertarVariedades = False
    Else
        InsertarVariedades = True
    End If
End Function






Private Function InsertarPaletsAlbaran(numPed As String, NumAlb As String) As Boolean
'Inserta en la tabla de lineas de albaran (albaran_palets)
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim NumLin As Integer

Dim Rs As ADODB.Recordset

    On Error GoTo eInsertarPaletsAlbaran

    Sql2 = "INSERT INTO albaran_palets (numalbar, numlinea, numpalet) VALUES "
    
    Sql = "Select distinct numpalet from palets where numpedid = " & DBLet(numPed, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumLin = 0
    While Not Rs.EOF
        NumLin = NumLin + 1
        
        Sql3 = "(" & DBLet(NumAlb, "N") & ", " & DBLet(NumLin, "N") & ", " & DBLet(Rs.Fields(0).Value, "N") & ")"
        
        conn.Execute Sql2 & Sql3
    
        Rs.MoveNext
    Wend


eInsertarPaletsAlbaran:
    If Err.Number <> 0 Then
        InsertarPaletsAlbaran = False
    Else
        InsertarPaletsAlbaran = True
    End If
End Function





Private Function EliminarPedido(numPed As Long) As Boolean
'Eliminar las lineas y la Cabecera de un Pedido. Tablas: scaped, sliped
Dim Sql As String

    On Error GoTo EEliminarPed

     Sql = " WHERE  numpedcl=" & numPed

    'Lineas de Pedido
    conn.Execute "Delete from " & NomTablaLineas & Sql

    'Cabecera
    conn.Execute "Delete from " & NombreTabla & Sql

EEliminarPed:
    If Err.Number <> 0 Then
        EliminarPedido = False
    Else
        EliminarPedido = True
    End If
End Function



Private Sub PosicionarDataTrasEliminar()
'Despues Eliminar y hacer refresh del Data, situar el Data en el registro siguiente
    If SituarDataTrasEliminar(Data1, NumRegElim) Then
        PonerCampos
    Else
        LimpiarCampos
        LimpiarDataGrids
        PonerModo 0
    End If
End Sub


Private Sub ImprimirAlbaran(Opcion As Integer, NumAlbar As String)
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadselect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String

    If NumAlbar = "" Then
        MsgBox "Debe seleccionar un Albarán para Imprimir.", vbInformation
        Exit Sub
    End If
    
    cadFormula = ""
    cadParam = ""
    cadselect = ""
    numParam = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    indRPT = 9 'Impresion de Albaran
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
      
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de albaran
    '---------------------------------------------------
    If NumAlbar <> "" Then
        'Nº palet
        devuelve = "{albaran.numalbar}=" & Val(NumAlbar)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        devuelve = "numalbar = " & Val(NumAlbar)
        If Not AnyadirAFormula(cadselect, devuelve) Then Exit Sub
    End If
    
    If Not HayRegParaInforme("albaran", cadselect) Then Exit Sub
     
     With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .ConSubInforme = True
            .Opcion = 0
            .Titulo = "Impresión de Albarán"
            .Show vbModal
    End With
End Sub

Private Function InsertarMovStock(NumAlb As String) As Boolean
Dim vCStock As CStock
Dim b As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String

    On Error Resume Next

    InsertarMovStock = False
    
    Set vCStock = New CStock
    b = True
'--monica
'    SQL = "select * from sliped WHERE " & Replace(ObtenerWhereCP, NombreTabla, NomTablaLineas)
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    vCStock.Fechamov = FechaAlb
    
    'para cada linea del Pedido Insertar en smoval y Actualizar Stock en salmac
    While (Not Rs.EOF) And b
        'si hay control de stock
'        SQL = DevuelveDesdeBDNew(conAri, "sartic", "ctrstock", "codartic", RS!codartic, "T")
'        If Val(SQL) = 1 Then
            If Not InicializarCStockAlbar(vCStock, "S", CStr(Rs!NumLinea), Rs) Then Exit Function
            vCStock.Documento = NumAlb
            If vCStock.Cantidad <> 0 Then
                'en actualizar stock comprobamos si el articulo tiene control de stock
                    b = vCStock.ActualizarStock
            End If
'        End If
        Rs.MoveNext
    Wend
    Set vCStock = Nothing
    Rs.Close
    Set Rs = Nothing
    
    InsertarMovStock = b
    
End Function


Private Function ActualizarCabPedido(NumAlb As String) As Boolean
Dim Sql As String

    On Error Resume Next

    Sql = "UPDATE pedidos SET numalbar= " & DBSet(NumAlb, "N") & ", fechaalb = " & DBSet(FechaAlb, "F")
    Sql = Sql & " WHERE " & ObtenerWhereCP(False)
    conn.Execute Sql
    
    If Err.Number <> 0 Then
        ActualizarCabPedido = False
    Else
        ActualizarCabPedido = True
    End If
End Function


Private Function InsertarLineasCalibres(Palet As Long, Variedad As Integer) As Boolean
Dim Sql As String

    Sql = "SELECT palets_calibre.codvarie, palets_calibre.codcalib, sum(numcajas)"
    Sql = Sql & " FROM palets_calibre, palets_variedad, palets WHERE palets.numpedid=" & Palet
    Sql = Sql & " and palets_calibre.codvarie = " & Variedad
    Sql = Sql & " and palets.numpalet = palets_variedad.numpalet "
    Sql = Sql & " and palets.numpalet = palets_calibre.numpalet "
    Sql = Sql & " GROUP BY 1,2"
    Sql = Sql & " ORDER BY 1,2"

End Function

Private Function TienePalets(Pedido As Long) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset

    TienePalets = False

    Sql = "select * from palets WHERE numpedid = " & DBSet(Pedido, "N")
    Set Rs = New ADODB.Recordset
    
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            TienePalets = True
        End If
    End If
End Function


Private Function InsertarVariedadesSinAgrupar(MenError As String, NumAlb As String) As Boolean
'Inserta en la tabla de lineas de albaran (slialb)
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset

Dim ImpLinea As String

Dim NumLin As Integer
Dim NumLin1 As Integer
Dim b As Boolean

Dim NumCajas As Long
Dim PesoBruto As Long
Dim Pesoneto As Long

Dim PesoBrutoLin As Long
Dim PesoNetoLin As Long

Dim VariedadAnt As Integer
Dim VarComAnt As Integer
Dim MarcaAnt As Integer
Dim ForfaitAnt As String
Dim CategoriAnt As String


Dim PesoBrutoVar As String
Dim PesoNetoVar As String
Dim NumCajasVar As String


    On Error GoTo eInsertarVariedades

        'Insertar en la tabla de Pedido, los registros seleccionados de la tabla de Palets
        
        Sql = ""
        Sql = "SELECT palets_calibre.codcalib, palets_variedad.codvarie, palets_variedad.codvarco, palets_variedad.codmarca, "
        Sql = Sql & " palets_variedad.codforfait, palets_variedad.categori, palets.codpalet, "
        Sql = Sql & " pesobrut,pesoneto,sum(palets_calibre.numcajas) "
        Sql = Sql & " FROM palets, palets_variedad, palets_calibre WHERE palets.numpedid=" & DBSet(Text1(0).Text, "N")
        Sql = Sql & " and palets.numpalet = palets_variedad.numpalet "
        Sql = Sql & " and palets_variedad.numpalet = palets_calibre.numpalet "
        Sql = Sql & " and palets_variedad.numlinea = palets_calibre.numlinea "
        Sql = Sql & " GROUP BY 2,3,4,5,6,7,1"
        Sql = Sql & " ORDER BY 2,3,4,5,6,7,1"
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        NumLin = 0
        b = True
        While Not Rs.EOF And b 'agrupamos las lineas de palets_variedad
            NumLin = NumLin + 1
            
            Sql3 = "select sum(pesobrut) from palets_variedad, palets, palets_calibre where palets.numpedid = " & Text1(0).Text
            Sql3 = Sql3 & " and palets_variedad.codvarie = " & DBSet(Rs.Fields(1).Value, "N")
            Sql3 = Sql3 & " and palets_variedad.codvarco = " & DBSet(Rs.Fields(2).Value, "N")
            Sql3 = Sql3 & " and palets_variedad.codforfait = " & DBSet(Rs.Fields(4).Value, "T")
            Sql3 = Sql3 & " and palets_variedad.codmarca = " & DBSet(Rs.Fields(3).Value, "N")
            Sql3 = Sql3 & " and palets_calibre.codcalib = " & DBSet(Rs.Fields(0).Value, "N")
            Sql3 = Sql3 & " and palets_variedad.numpalet = palets.numpalet"
            Sql3 = Sql3 & " and palets_variedad.numpalet = palets_calibre.numpalet"
            Sql3 = Sql3 & " and palets_variedad.numlinea = palets_calibre.numlinea"
            If DBSet(Rs.Fields(5).Value, "T") = ValorNulo Then
                Sql3 = Sql3 & " and palets_variedad.categori is null "
            Else
                Sql3 = Sql3 & " and palets_variedad.categori = " & DBSet(Rs.Fields(5).Value, "T")
            End If

            '[Monica] 15/06/2010 añadido costes paletizacion
            Sql3 = Sql3 & " and palets.codpalet = " & DBSet(Rs.Fields(6).Value, "N")



            PesoBrutoVar = DevuelveValor(Sql3)
            
            Sql3 = "select sum(pesoneto) from palets_variedad, palets, palets_calibre where palets.numpedid = " & Text1(0).Text
            Sql3 = Sql3 & " and palets_variedad.codvarie = " & DBSet(Rs.Fields(1).Value, "N")
            Sql3 = Sql3 & " and palets_variedad.codvarco = " & DBSet(Rs.Fields(2).Value, "N")
            Sql3 = Sql3 & " and palets_variedad.codforfait = " & DBSet(Rs.Fields(4).Value, "T")
            Sql3 = Sql3 & " and palets_variedad.codmarca = " & DBSet(Rs.Fields(3).Value, "N")
            Sql3 = Sql3 & " and palets_calibre.codcalib = " & DBSet(Rs.Fields(0).Value, "N")
            Sql3 = Sql3 & " and palets_variedad.numpalet = palets.numpalet"
            Sql3 = Sql3 & " and palets_variedad.numpalet = palets_calibre.numpalet"
            Sql3 = Sql3 & " and palets_variedad.numlinea = palets_calibre.numlinea"
            If DBSet(Rs.Fields(5).Value, "T") = ValorNulo Then
                Sql3 = Sql3 & " and palets_variedad.categori is null "
            Else
                Sql3 = Sql3 & " and palets_variedad.categori = " & DBSet(Rs.Fields(5).Value, "T")
            End If
            '[Monica] 15/06/2010 añadido costes paletizacion
            Sql3 = Sql3 & " and palets.codpalet = " & DBSet(Rs.Fields(6).Value, "N")

            
            PesoNetoVar = DevuelveValor(Sql3)
            
            Sql3 = "select sum(palets_variedad.numcajas) from palets_variedad, palets, palets_calibre where palets.numpedid = " & Text1(0).Text
            Sql3 = Sql3 & " and palets_variedad.codvarie = " & DBSet(Rs.Fields(1).Value, "N")
            Sql3 = Sql3 & " and palets_variedad.codvarco = " & DBSet(Rs.Fields(2).Value, "N")
            Sql3 = Sql3 & " and palets_variedad.codforfait = " & DBSet(Rs.Fields(4).Value, "T")
            Sql3 = Sql3 & " and palets_variedad.codmarca = " & DBSet(Rs.Fields(3).Value, "N")
            Sql3 = Sql3 & " and palets_calibre.codcalib = " & DBSet(Rs.Fields(0).Value, "N")
            Sql3 = Sql3 & " and palets_variedad.numpalet = palets.numpalet"
            Sql3 = Sql3 & " and palets_variedad.numpalet = palets_calibre.numpalet"
            Sql3 = Sql3 & " and palets_variedad.numlinea = palets_calibre.numlinea"
            If DBSet(Rs.Fields(5).Value, "T") = ValorNulo Then
                Sql3 = Sql3 & " and palets_variedad.categori is null "
            Else
                Sql3 = Sql3 & " and palets_variedad.categori = " & DBSet(Rs.Fields(5).Value, "T")
            End If
            '[Monica] 15/06/2010 añadido costes paletizacion
            Sql3 = Sql3 & " and palets.codpalet = " & DBSet(Rs.Fields(6).Value, "N")

            
            NumCajasVar = DevuelveValor(Sql3)
        
            If NumCajasVar <> DBLet(Rs.Fields(8).Value, "N") And NumCajasVar <> 0 Then
                PesoNetoVar = Round2(CCur(PesoNetoVar) * DBLet(Rs.Fields(8).Value, "N") / CCur(NumCajasVar), 0)
            End If
        
            
            Sql = "INSERT INTO albaran_variedad(numalbar,numlinea,codvarie,codvarco,codforfait,codmarca,"
            Sql = Sql & "categori,totpalet,numcajas,pesobrut,pesoneto,preciopro,preciodef,"
            Sql = Sql & "codincid,impcomis,observac, codpalet) VALUES "
            Sql = Sql & "(" & NumAlb & "," & NumLin & "," & DBSet(Rs.Fields(1), "N") & "," & DBSet(Rs.Fields(2), "N") & ","
            Sql = Sql & DBSet(Rs.Fields(4), "T") & "," & DBSet(Rs.Fields(3), "N") & ","
            Sql = Sql & DBSet(Rs.Fields(5), "T") & "," & ValorNulo & ","
            Sql = Sql & DBSet(Rs.Fields(8).Value, "N") & "," & DBSet(PesoBrutoVar, "N") & ","
            Sql = Sql & DBSet(PesoNetoVar, "N") & "," & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & DBSet(Incidencia, "N") & "," & ValorNulo & "," & ValorNulo & ","
            Sql = Sql & DBSet(Rs.Fields(6).Value, "N") & ")"
            
'            Sql = Sql & DBSet(RS.Fields(8), "N") & "," & DBSet(RS.Fields(6), "N") & ","
'            Sql = Sql & DBSet(RS.Fields(7), "N") & "," & ValorNulo & "," & ValorNulo & ","
'            Sql = Sql & DBSet(Incidencia, "N") & "," & ValorNulo & "," & ValorNulo & ")"
    
            conn.Execute Sql
            
            ' en cuantos palets aparece esta linea
            Sql = "select count(distinct palets.numpalet) from palets, palets_variedad, palets_calibre where palets.numpedid = " & DBSet(Text1(0).Text, "N")
            Sql = Sql & " and palets_variedad.codvarie = " & DBSet(Rs.Fields(1).Value, "N")
            Sql = Sql & " and palets_variedad.codvarco = " & DBSet(Rs.Fields(2).Value, "N")
            Sql = Sql & " and palets_variedad.codforfait = " & DBSet(Rs.Fields(4).Value, "T")
            Sql = Sql & " and palets_variedad.codmarca = " & DBSet(Rs.Fields(3).Value, "N")
            Sql = Sql & " and palets_calibre.codcalib = " & DBSet(Rs.Fields(0).Value, "N")
            
            If DBSet(Rs.Fields(5).Value, "T") = ValorNulo Then
                Sql = Sql & " and palets_variedad.categori is null "
            Else
                Sql = Sql & " and palets_variedad.categori = " & DBSet(Rs.Fields(5).Value, "T")
            End If
            '[Monica] 15/06/2010 añadido costes paletizacion
            Sql = Sql & " and palets.codpalet = " & DBSet(Rs.Fields(6).Value, "N")
            
            Sql = Sql & " and palets.numpalet = palets_variedad.numpalet "
            
            Set Rs1 = New ADODB.Recordset
            Rs1.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not Rs1.EOF Then
                If DBLet(Rs1.Fields(0).Value, "N") <> 0 Then
                    Sql = "update albaran_variedad set totpalet = " & DBSet(Rs1.Fields(0).Value, "N")
                    Sql = Sql & " where numalbar = " & DBSet(NumAlb, "N")
                    Sql = Sql & " and numlinea = " & DBSet(NumLin, "N")
                    
                    conn.Execute Sql
                End If
            End If
            Set Rs1 = Nothing
            
            
            ' insertamos en la tabla albaran calibres
            Sql = "INSERT INTO albaran_calibre(numalbar, numlinea, numline1, codvarie, codcalib, "
            Sql = Sql & "numcajas, pesobrut, pesoneto) VALUES ("
            Sql = Sql & NumAlb & "," & NumLin & ",1," & DBSet(Rs.Fields(1).Value, "N") & ","
            Sql = Sql & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(Rs.Fields(8).Value, "N") & ","
            Sql = Sql & DBSet(PesoBrutoVar, "N") & "," & DBSet(PesoNetoVar, "N") & ")"
            
            MenError = "Error al insertar en la tabla Calibres de Albaran (albaran_calibre)."
            conn.Execute Sql
            
            
            
            
            
'            'Insertar en la tabla de albaranes, los registros seleccionados de la tabla de Palets
'            Sql = ""
'            Sql = "SELECT palets_variedad.codvarie, palets_variedad.codvarco, palets_variedad.codmarca, "
'            Sql = Sql & " palets_variedad.codforfait, palets_variedad.categori, palets_calibre.codcalib, "
'            Sql = Sql & " sum(pesobrut),sum(pesoneto),sum(palets_variedad.numcajas),sum(palets_calibre.numcajas) "
'            Sql = Sql & " FROM palets, palets_variedad, palets_calibre WHERE palets.numpedid=" & Text1(0).Text
'            Sql = Sql & " and palets_variedad.codvarie = " & DBSet(Rs.Fields(1).Value, "N")
'            Sql = Sql & " and palets_variedad.codvarco = " & DBSet(Rs.Fields(2).Value, "N")
'            Sql = Sql & " and palets_variedad.codforfait = " & DBSet(Rs.Fields(4).Value, "T")
'            Sql = Sql & " and palets_variedad.codmarca = " & DBSet(Rs.Fields(3).Value, "N")
'            Sql = Sql & " and palets_calibre.codcalib = " & DBSet(Rs.Fields(0).Value, "N")
'            If DBSet(Rs.Fields(5).Value, "T") = ValorNulo Then
'                Sql = Sql & " and palets_variedad.categori is null "
'            Else
'                Sql = Sql & " and palets_variedad.categori = " & DBSet(Rs.Fields(5).Value, "T")
'            End If
'            Sql = Sql & " and palets.numpalet = palets_variedad.numpalet "
'            Sql = Sql & " and palets_variedad.numpalet = palets_calibre.numpalet "
'            Sql = Sql & " and palets_variedad.numlinea = palets_calibre.numlinea "
'            Sql = Sql & " GROUP BY 1,2,3,4,5,6"
'            Sql = Sql & " ORDER BY 1,2,3,4,5,6"
'
'            Set Rs1 = New ADODB.Recordset
'            Rs1.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'            PesoBruto = 0
'            Pesoneto = 0
'            NumCajas = 0
'            NumLin1 = 0
'            While Not Rs1.EOF
'                NumLin1 = NumLin1 + 1
''22-09-2008
''                PesoBruto = PesoBruto + DBLet(Rs1.Fields(6).Value, "N")
''                PesoNeto = PesoNeto + DBLet(Rs1.Fields(7).Value, "N")
''                NumCajas = NumCajas + DBLet(Rs1.Fields(8).Value, "N")
'
'
'                PesoBrutoLin = 0
'                If CCur(NumCajasVar) <> 0 Then
'                    PesoBrutoLin = Round2(CCur(PesoBrutoVar) * DBLet(Rs1.Fields(9).Value, "N") / CCur(NumCajasVar), 0)
'                End If
'                PesoBruto = PesoBruto + PesoBrutoLin
'                PesoNetoLin = 0
'                If CCur(NumCajasVar) <> 0 Then
'                    PesoNetoLin = Round2(CCur(PesoNetoVar) * DBLet(Rs1.Fields(9).Value, "N") / CCur(NumCajasVar), 0)
'                End If
'                Pesoneto = Pesoneto + PesoNetoLin
'
''                PesoBrutoLin = 0
''                If DBLet(RS1.Fields(8).Value, "N") <> 0 Then
''                    PesoBrutoLin = Round2(DBLet(RS1.Fields(6).Value, "N") * DBLet(RS1.Fields(9).Value, "N") / DBLet(RS1.Fields(8).Value, "N"), 0)
''                End If
''                PesoBruto = PesoBruto + PesoBrutoLin
''                PesoNetoLin = 0
''                If DBLet(RS1.Fields(8).Value, "N") <> 0 Then
''                    PesoNetoLin = Round2(DBLet(RS1.Fields(7).Value, "N") * DBLet(RS1.Fields(9).Value, "N") / DBLet(RS1.Fields(8).Value, "N"), 0)
''                End If
''                Pesoneto = Pesoneto + PesoNetoLin
'
'                ' insertamos en la tabla albaran calibres
'                Sql = "INSERT INTO albaran_calibre(numalbar, numlinea, numline1, codvarie, codcalib, "
'                Sql = Sql & "numcajas, pesobrut, pesoneto) VALUES ("
'                Sql = Sql & NumAlb & "," & NumLin & "," & NumLin1 & "," & DBSet(Rs1.Fields(0).Value, "N") & ","
'                Sql = Sql & DBSet(Rs1.Fields(5).Value, "N") & "," & DBSet(Rs1.Fields(9).Value, "N") & ","
'                Sql = Sql & DBSet(PesoBrutoLin, "N") & "," & DBSet(PesoNetoLin, "N") & ")"
'
'                MenError = "Error al insertar en la tabla Calibres de Albaran (albaran_calibre)."
'                conn.Execute Sql
'
'                Rs1.MoveNext
'            Wend
'
'            Set Rs1 = Nothing
'
'            ' redondeamos en la ultima linea de calibres
'            Sql = "select sum(pesobrut),sum(pesoneto) "
'            Sql = Sql & " FROM palets, palets_variedad WHERE palets.numpedid=" & Text1(0).Text
'            Sql = Sql & " and palets_variedad.codvarie = " & DBSet(Rs.Fields(0).Value, "N")
'            Sql = Sql & " and palets_variedad.codvarco = " & DBSet(Rs.Fields(1).Value, "N")
'            Sql = Sql & " and palets_variedad.codforfait = " & DBSet(Rs.Fields(3).Value, "T")
'            Sql = Sql & " and palets_variedad.codmarca = " & DBSet(Rs.Fields(2).Value, "N")
'
'            If DBSet(Rs.Fields(4).Value, "T") = ValorNulo Then
'                Sql = Sql & " and palets_variedad.categori is null "
'            Else
'                Sql = Sql & " and palets_variedad.categori = " & DBSet(Rs.Fields(4).Value, "T")
'            End If
'            Sql = Sql & " and palets.numpalet = palets_variedad.numpalet "
'
'            Set Rs1 = New ADODB.Recordset
'            Rs1.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'            If Not Rs1.EOF Then
'                If CCur(PesoBrutoVar) <> PesoBruto Or CCur(PesoNetoVar) <> Pesoneto Then
'                    Sql = "update albaran_calibre set pesobrut = pesobrut + " & DBLet(Rs1.Fields(0).Value, "N") - PesoBruto
'                    Sql = Sql & ", pesoneto = pesoneto + " & DBLet(Rs1.Fields(1).Value, "N") - Pesoneto
'                    Sql = Sql & " where albaran_calibre.numalbar  = " & NumAlb
'                    Sql = Sql & " and albaran_calibre.numlinea = " & NumLin
'                    Sql = Sql & " and albaran_calibre.numline1 = " & NumLin1
'
'                    conn.Execute Sql
'                End If
'            End If
'            Set Rs1 = Nothing
'
''22-09-2008
''            ' actualizamos el numero de cajas, peso bruto y peso neto
''            sql = "UPDATE albaran_variedad SET numcajas = " & DBLet(NumCajas, "N")
''            sql = sql & ", pesobrut = " & DBSet(PesoBruto, "N")
''            sql = sql & ", pesoneto = " & DBSet(PesoNeto, "N")
''            sql = sql & " where numalbar = " & DBSet(NumAlb, "N")
''            sql = sql & " and numlinea = " & DBSet(NumLin, "N")
''
''            conn.Execute sql
            
            MenError = "Error al Actualizar Costes."
            If b Then b = ActualizarCostes(CLng(NumAlb), NumLin, True, DBLet(Rs!codforfait, "T"), DBLet(Rs!CodPalet, "N"))
            
            Rs.MoveNext
        Wend
        Set Rs = Nothing



eInsertarVariedades:
    If Err.Number <> 0 Or Not b Then
        InsertarVariedadesSinAgrupar = False
    Else
        InsertarVariedadesSinAgrupar = True
    End If
End Function

