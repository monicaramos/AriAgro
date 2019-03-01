VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmManForfaits 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Forfaits"
   ClientHeight    =   10905
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   17730
   Icon            =   "frmManForfaits.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10905
   ScaleWidth      =   17730
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
      Index           =   0
      Left            =   14895
      TabIndex        =   85
      Top             =   240
      Width           =   1605
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   135
      TabIndex        =   83
      Top             =   45
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   84
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
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3795
      TabIndex        =   81
      Top             =   45
      Width           =   1785
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   210
         TabIndex        =   82
         Top             =   180
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Expandir Operaciones"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cambio Costes"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Duplicar Confección"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   5655
      TabIndex        =   79
      Top             =   45
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   80
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
   Begin VB.Frame Frame2 
      Height          =   3465
      Index           =   0
      Left            =   135
      TabIndex        =   23
      Top             =   810
      Width           =   17250
      Begin VB.TextBox text1 
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
         Left            =   8295
         MaxLength       =   9
         TabIndex        =   10
         Tag             =   "Precio por Kilo|N|S|||forfaits|preciokilonom|###0.0000||"
         Top             =   2985
         Width           =   1440
      End
      Begin VB.TextBox text1 
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
         Left            =   8280
         MaxLength       =   16
         TabIndex        =   9
         Tag             =   "Cajas por Palet|N|S|||forfaits|cajaspalet|##,##0||"
         Top             =   2505
         Width           =   1440
      End
      Begin VB.TextBox text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Index           =   3
         Left            =   180
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Tag             =   "Observaciones|T|S|||forfaits|observac|||"
         Top             =   1800
         Width           =   6600
      End
      Begin VB.TextBox text1 
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
         Left            =   8280
         MaxLength       =   16
         TabIndex        =   8
         Tag             =   "Peso Caja|N|S|||forfaits|pesocaja|#0.00||"
         Top             =   2055
         Width           =   1440
      End
      Begin VB.TextBox text2 
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
         Left            =   2295
         TabIndex        =   58
         Text            =   "12345678901234567890"
         Top             =   1125
         Width           =   4515
      End
      Begin VB.TextBox text1 
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
         Left            =   1350
         MaxLength       =   6
         TabIndex        =   2
         Tag             =   "Variedad|N|S|||forfaits|codvarie|000000||"
         Text            =   "123456"
         Top             =   1125
         Width           =   900
      End
      Begin VB.TextBox text2 
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
         Left            =   12825
         TabIndex        =   56
         Top             =   2970
         Width           =   4065
      End
      Begin VB.TextBox text1 
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
         Left            =   12150
         MaxLength       =   4
         TabIndex        =   17
         Tag             =   "Palet|N|N|0|999|forfaits|codpalet|000||"
         Text            =   "123"
         Top             =   2970
         Width           =   630
      End
      Begin VB.TextBox text2 
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
         Left            =   12825
         TabIndex        =   54
         Top             =   2520
         Width           =   4065
      End
      Begin VB.TextBox text1 
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
         Left            =   12150
         MaxLength       =   4
         TabIndex        =   16
         Tag             =   "Marca|N|S|0|999|forfaits|codmarca|000||"
         Text            =   "123"
         Top             =   2520
         Width           =   630
      End
      Begin VB.TextBox text2 
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
         Left            =   12825
         TabIndex        =   52
         Top             =   2070
         Width           =   4065
      End
      Begin VB.TextBox text1 
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
         Left            =   12150
         MaxLength       =   4
         TabIndex        =   15
         Tag             =   "Presentacion|N|N|0|999|forfaits|codprese|000||"
         Text            =   "123"
         Top             =   2070
         Width           =   630
      End
      Begin VB.TextBox text2 
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
         Left            =   12825
         TabIndex        =   50
         Top             =   1605
         Width           =   4065
      End
      Begin VB.TextBox text1 
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
         Left            =   12150
         MaxLength       =   4
         TabIndex        =   14
         Tag             =   "Confeccion|N|N|0|999|forfaits|codtipco|000||"
         Text            =   "123"
         Top             =   1605
         Width           =   630
      End
      Begin VB.TextBox text2 
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
         Left            =   12825
         TabIndex        =   48
         Top             =   1155
         Width           =   4065
      End
      Begin VB.TextBox text1 
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
         Left            =   12150
         MaxLength       =   4
         TabIndex        =   13
         Tag             =   "Medida|N|N|0|999|forfaits|codmedid|000||"
         Text            =   "123"
         Top             =   1155
         Width           =   630
      End
      Begin VB.TextBox text2 
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
         Left            =   12825
         TabIndex        =   46
         Top             =   705
         Width           =   4065
      End
      Begin VB.TextBox text1 
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
         Left            =   12150
         MaxLength       =   4
         TabIndex        =   12
         Tag             =   "Capacidad|N|N|0|999|forfaits|codcapac|000||"
         Text            =   "123"
         Top             =   705
         Width           =   630
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   3915
         MaxLength       =   16
         TabIndex        =   18
         Top             =   1980
         Width           =   1035
      End
      Begin VB.TextBox text1 
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
         Left            =   8280
         MaxLength       =   16
         TabIndex        =   7
         Tag             =   "Kilos Unidad|N|S|0|999.99|forfaits|kilosuni|##0.00||"
         Top             =   1605
         Width           =   1440
      End
      Begin VB.TextBox text1 
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
         Left            =   8280
         MaxLength       =   16
         TabIndex        =   6
         Tag             =   "Kilos Caja|N|S|0|999.99|forfaits|kiloscaj|##0.00||"
         Top             =   1125
         Width           =   1440
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
         ItemData        =   "frmManForfaits.frx":000C
         Left            =   8280
         List            =   "frmManForfaits.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "Facturar por|N|N|0||forfaits|facturar|||"
         Top             =   660
         Width           =   1440
      End
      Begin VB.TextBox text1 
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
         Left            =   1350
         MaxLength       =   40
         TabIndex        =   1
         Tag             =   "Nombre|T|N|||forfaits|nomconfe|||"
         Text            =   "1234657890123456798012345678901234567890"
         Top             =   675
         Width           =   5460
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
         ItemData        =   "frmManForfaits.frx":002C
         Left            =   8280
         List            =   "frmManForfaits.frx":0036
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Tag             =   "Caja / Kilo|N|N|0|1|forfaits|cajakilo|||"
         Top             =   195
         Width           =   1440
      End
      Begin VB.TextBox text2 
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
         Left            =   12825
         TabIndex        =   26
         Top             =   255
         Width           =   4065
      End
      Begin VB.TextBox text1 
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
         Left            =   12150
         MaxLength       =   4
         TabIndex        =   11
         Tag             =   "Envase|N|N|0|999|forfaits|codtipen|000||"
         Text            =   "123"
         Top             =   255
         Width           =   630
      End
      Begin VB.TextBox text1 
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
         Left            =   1350
         MaxLength       =   16
         TabIndex        =   0
         Tag             =   "Código  forfait|T|N|||forfaits|codforfait||S|"
         Text            =   "1234568790123456"
         Top             =   270
         Width           =   1890
      End
      Begin VB.Label Label1 
         Caption         =   "Precio/Kilo"
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
         Left            =   7065
         TabIndex        =   76
         Top             =   2985
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Cajas/Palet"
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
         Left            =   7065
         TabIndex        =   75
         Top             =   2535
         Width           =   1185
      End
      Begin VB.Label Label40 
         Caption         =   "Códigos EAN"
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
         Left            =   4995
         TabIndex        =   74
         Top             =   270
         Width           =   1320
      End
      Begin VB.Image imgBuscar 
         Height          =   330
         Index           =   8
         Left            =   6435
         ToolTipText     =   "Códigos EAN asociados"
         Top             =   225
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Peso Caja"
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
         Left            =   7065
         TabIndex        =   72
         Top             =   2085
         Width           =   1185
      End
      Begin VB.Label Label10 
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
         Height          =   255
         Left            =   180
         TabIndex        =   59
         Top             =   1125
         Width           =   870
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   1080
         ToolTipText     =   "Buscar Colectivo"
         Top             =   1125
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Palet"
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
         Left            =   10530
         TabIndex        =   57
         Top             =   2970
         Width           =   870
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   11880
         ToolTipText     =   "Buscar Palet"
         Top             =   2970
         Width           =   240
      End
      Begin VB.Label Label8 
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
         Left            =   10530
         TabIndex        =   55
         Top             =   2520
         Width           =   870
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   11880
         ToolTipText     =   "Buscar Marca"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Presentación"
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
         Left            =   10530
         TabIndex        =   53
         Top             =   2070
         Width           =   1320
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   11880
         ToolTipText     =   "Buscar Presentación"
         Top             =   2070
         Width           =   240
      End
      Begin VB.Label Label5 
         Caption         =   "Confección"
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
         Left            =   10530
         TabIndex        =   51
         Top             =   1605
         Width           =   1320
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   11880
         ToolTipText     =   "Buscar Confección"
         Top             =   1605
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Medida"
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
         Left            =   10530
         TabIndex        =   49
         Top             =   1155
         Width           =   870
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   11880
         ToolTipText     =   "Buscar Medida"
         Top             =   1155
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Capacidad"
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
         Left            =   10530
         TabIndex        =   47
         Top             =   705
         Width           =   1275
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   11880
         ToolTipText     =   "Buscar Capacidad"
         Top             =   705
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Kilos/Unidad"
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
         Left            =   7065
         TabIndex        =   45
         Top             =   1650
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Kilos/Caja"
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
         Left            =   7065
         TabIndex        =   44
         Top             =   1155
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Facturar por "
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
         Left            =   7065
         TabIndex        =   43
         Top             =   705
         Width           =   1200
      End
      Begin VB.Label Label6 
         Caption         =   "Nombre"
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
         Left            =   180
         TabIndex        =   42
         Top             =   675
         Width           =   780
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
         Left            =   180
         TabIndex        =   29
         Top             =   1530
         Width           =   1575
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   1755
         ToolTipText     =   "Zoom descripción"
         Top             =   1530
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Confección por "
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
         Left            =   7065
         TabIndex        =   28
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label18 
         Caption         =   "Envase"
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
         Left            =   10530
         TabIndex        =   27
         Top             =   255
         Width           =   870
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   11880
         ToolTipText     =   "Buscar Envase"
         Top             =   255
         Width           =   240
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   24
         Top             =   270
         Width           =   1095
      End
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
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
      Left            =   13590
      TabIndex        =   78
      Text            =   "Text3"
      Top             =   10305
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
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
      Left            =   9155
      TabIndex        =   70
      Text            =   "Text3"
      Top             =   10305
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
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
      Left            =   15615
      TabIndex        =   68
      Text            =   "Text3"
      Top             =   9855
      Width           =   1485
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
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
      Left            =   9155
      TabIndex        =   66
      Text            =   "Text3"
      Top             =   9855
      Width           =   1575
   End
   Begin VB.Frame FrameAux1 
      Caption         =   "Costes"
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
      Height          =   5385
      Left            =   11340
      TabIndex        =   36
      Top             =   4320
      Width           =   6070
      Begin VB.CommandButton btnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1350
         MaskColor       =   &H00000000&
         TabIndex        =   65
         ToolTipText     =   "Buscar Coste"
         Top             =   2610
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
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
         Height          =   330
         Index           =   3
         Left            =   1710
         TabIndex        =   64
         Text            =   "nombre"
         Top             =   2610
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   9
         Left            =   810
         MaxLength       =   2
         TabIndex        =   38
         Tag             =   "Codigo coste|N|N|1|99|forfaits_costes|codcoste|00|S|"
         Text            =   "linea"
         Top             =   2610
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   10
         Left            =   3690
         MaxLength       =   9
         TabIndex        =   39
         Tag             =   "Importes|N|N|||forfaits_costes|importes|###0.0000||"
         Text            =   "impcoste"
         Top             =   2610
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   8
         Left            =   360
         MaxLength       =   16
         TabIndex        =   37
         Tag             =   "Codigo Forfaits|T|N|||forfaits_costes|codforfait||S|"
         Text            =   "codfor"
         Top             =   2610
         Visible         =   0   'False
         Width           =   555
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   1
         Left            =   90
         TabIndex        =   40
         Top             =   270
         Width           =   1200
         _ExtentX        =   2117
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
      Begin MSDataGridLib.DataGrid DataGridAux 
         Bindings        =   "frmManForfaits.frx":004C
         Height          =   2390
         Index           =   1
         Left            =   90
         TabIndex        =   41
         Top             =   675
         Width           =   5910
         _ExtentX        =   10425
         _ExtentY        =   4207
         _Version        =   393216
         AllowUpdate     =   0   'False
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   19
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
            AllowFocus      =   0   'False
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc AdoAux 
         Height          =   375
         Index           =   1
         Left            =   1350
         Top             =   180
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
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
         Caption         =   "AdoAux(1)"
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
   End
   Begin VB.Frame FrameAux0 
      Caption         =   "Envases"
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
      Height          =   5430
      Left            =   180
      TabIndex        =   30
      Top             =   4320
      Width           =   11010
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   3
         Left            =   450
         MaxLength       =   16
         TabIndex        =   73
         Tag             =   "Linea|N|N|||forfaits_envases|numlinea|00|S|"
         Text            =   "lin"
         Top             =   2565
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.TextBox txtAux2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
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
         Height          =   330
         Index           =   1
         Left            =   5850
         TabIndex        =   63
         Text            =   "importe"
         Top             =   2610
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtAux2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
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
         Height          =   330
         Index           =   0
         Left            =   4815
         TabIndex        =   62
         Text            =   "precio"
         Top             =   2610
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton btnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1620
         MaskColor       =   &H00000000&
         TabIndex        =   61
         ToolTipText     =   "Buscar Envase"
         Top             =   2610
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
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
         Height          =   330
         Index           =   2
         Left            =   1830
         TabIndex        =   60
         Top             =   2610
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   2
         Left            =   3690
         MaxLength       =   9
         TabIndex        =   33
         Tag             =   "Cantidad|N|N|||forfaits_envases|cantidad|###0.0000||"
         Text            =   "cantidad"
         Top             =   2610
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   1
         Left            =   675
         MaxLength       =   16
         TabIndex        =   32
         Tag             =   "Codigo Articulo|T|N|||forfaits_envases|codartic||N|"
         Text            =   "articuloarticulo"
         Top             =   2565
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   0
         Left            =   45
         MaxLength       =   16
         TabIndex        =   31
         Tag             =   "Codigo Forfaits|T|N|||forfaits_envases|codforfait||S|"
         Text            =   "codforfaits"
         Top             =   2565
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   0
         Left            =   135
         TabIndex        =   34
         Top             =   225
         Width           =   1200
         _ExtentX        =   2117
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
      Begin MSAdodcLib.Adodc AdoAux 
         Height          =   375
         Index           =   0
         Left            =   3720
         Top             =   225
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
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
         Caption         =   "AdoAux(0)"
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
      Begin MSDataGridLib.DataGrid DataGridAux 
         Bindings        =   "frmManForfaits.frx":0064
         Height          =   4545
         Index           =   0
         Left            =   135
         TabIndex        =   35
         Top             =   630
         Width           =   10595
         _ExtentX        =   18680
         _ExtentY        =   8017
         _Version        =   393216
         AllowUpdate     =   0   'False
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   19
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
            AllowFocus      =   0   'False
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   240
      TabIndex        =   21
      Top             =   10035
      Width           =   2865
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
         Left            =   120
         TabIndex        =   22
         Top             =   180
         Width           =   2655
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
      Left            =   16335
      TabIndex        =   20
      Top             =   10305
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
      Left            =   15135
      TabIndex        =   19
      Top             =   10290
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   1980
      Top             =   6120
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Left            =   16320
      TabIndex        =   25
      Top             =   10290
      Visible         =   0   'False
      Width           =   1065
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   16935
      TabIndex        =   86
      Top             =   180
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
   Begin VB.Label Label11 
      Caption         =   "Total por Kilo: "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   240
      Index           =   3
      Left            =   11925
      TabIndex        =   77
      Top             =   10350
      Width           =   1725
   End
   Begin VB.Line Line1 
      X1              =   8865
      X2              =   10935
      Y1              =   10215
      Y2              =   10215
   End
   Begin VB.Label Label11 
      Caption         =   "Coste Total Forfait: "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   240
      Index           =   2
      Left            =   6675
      TabIndex        =   71
      Top             =   10350
      Width           =   2445
   End
   Begin VB.Label Label11 
      Caption         =   "TOTAL IMPORTE COSTES: "
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
      Left            =   11925
      TabIndex        =   69
      Top             =   9900
      Width           =   2400
   End
   Begin VB.Label Label11 
      Caption         =   "TOTAL IMPORTE ENVASES: "
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
      Index           =   0
      Left            =   6675
      TabIndex        =   67
      Top             =   9900
      Width           =   2445
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
         HelpContextID   =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         HelpContextID   =   2
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnExpandirOperaciones 
         Caption         =   "Expandir &Operaciones"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnCambioCostes 
         Caption         =   "&Cambio Costes"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnDuplicarConfec 
         Caption         =   "&Duplicar Confección"
         Shortcut        =   ^D
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
Attribute VB_Name = "frmManForfaits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MANOLO                   -+-+
' +-+- Menú: CLIENTES                  -+-+
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

' +-+-+-+- DISSENY +-+-+-+-
' 1. Posar tots els controls al formulari
' 2. Posar els index correlativament
' 3. Si n'hi han botons de buscar repasar el ToolTipText
' 4. Alliniar els camps numérics a la dreta i el resto a l'esquerra
' 5. Posar els TAGs
' (si es INTEGER: si PK => mínim 1; si no PK => mínim 0; màxim => 99; format => 00)
' (si es DECIMAL; mínim => 0; màxim => 99.99; format => #,###,###,##0.00)
' (si es DATE; format => dd/mm/yyyy)
' 6. Posar els MAXLENGTHs
' 7. Posar els TABINDEXs

Option Explicit

'Dim T1 As Single

Public DatosADevolverBusqueda As String    'Tindrà el nº de text que vol que torne, empipat
Public Event DatoSeleccionado(CadenaSeleccion As String)
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean

' *** declarar els formularis als que vaig a cridar ***
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1

Private WithEvents frmArt As frmManArtic 'articulos
Attribute frmArt.VB_VarHelpID = -1
Private WithEvents frmVar As frmManVariedad 'variedades
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
Private WithEvents frmConf As frmManFConf 'confeccion
Attribute frmConf.VB_VarHelpID = -1


Private WithEvents frmNCoste As frmManNomCoste 'nombre de coste
Attribute frmNCoste.VB_VarHelpID = -1
Private WithEvents frmOpEnv As frmOperaEnv 'Operaciones masivas sobre envases
Attribute frmOpEnv.VB_VarHelpID = -1
Private WithEvents frmCEan As frmCodEAN 'Codigos Ean
Attribute frmCEan.VB_VarHelpID = -1

Private WithEvents frmFor As frmBasico2 'Mandabusquedaprevia
Attribute frmFor.VB_VarHelpID = -1


'*****************************************************
Private Modo As Byte
'*************** MODOS ********************
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'   5.-  Manteniment Llinies

'+-+-Variables comuns a tots els formularis+-+-+

Dim ModoLineas As Byte
'1.- Afegir,  2.- Modificar,  3.- Borrar,  0.-Passar control a Llínies

Dim NumTabMto As Integer 'Indica quin nº de Tab està en modo Mantenimient
Dim TituloLinea As String 'Descripció de la llínia que està en Mantenimient
Dim PrimeraVez As Boolean

Private CadenaConsulta As String 'SQL de la taula principal del formulari
Private Ordenacion As String
Private NombreTabla As String  'Nom de la taula
Private NomTablaLineas As String 'Nom de la Taula de llínies del Mantenimient en que estem

Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

'Private VieneDeBuscar As Boolean
'Per a quan torna 2 poblacions en el mateix codi Postal. Si ve de pulsar prismatic
'de búsqueda posar el valor de població seleccionada i no tornar a recuperar de la Base de Datos

Dim btnPrimero As Byte 'Variable que indica el nº del Botó PrimerRegistro en la Toolbar1
'Dim CadAncho() As Boolean  'array, per a quan cridem al form de llínies
Dim indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim CadB As String

Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 0 'Articulos
            Set frmArt = New frmManArtic
            frmArt.DatosADevolverBusqueda = "0|1|"
            frmArt.CodigoActual = txtAux(1).Text
            frmArt.Show vbModal
            Set frmArt = Nothing
            PonerFoco txtAux(1)
        Case 1 'Costes
            Set frmNCoste = New frmManNomCoste
            frmNCoste.DatosADevolverBusqueda = "0|1|"
            frmNCoste.CodigoActual = txtAux(9).Text
            frmNCoste.Show vbModal
            Set frmNCoste = Nothing
            PonerFoco txtAux(9)
    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub


Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BÚSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm2(Me, 1) Then
'                    text2(9).Text = PonerNombreCuenta(text1(9), Modo, text1(0).Text)
        
                    Data1.RecordSource = "Select * from " & NombreTabla & Ordenacion
                    TerminaBloquear
                    PosicionarData
                End If
            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario2(Me, 1) Then
                    TerminaBloquear
                    PosicionarData
                End If
            Else
                ModoLineas = 0
            End If
        ' *** si n'hi han llínies ***
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 1 'afegir llínia
                    InsertarLinea
                Case 2 'modificar llínies
                    If ModificarLinea Then
                        PosicionarData
                    Else
                        PonerFoco txtAux(12)
                    End If
            End Select
            'nuevo calculamos los totales de lineas
            CalcularTotales
        ' **************************
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
'    If PrimeraVez Then PrimeraVez = False
    If PrimeraVez Then
        PrimeraVez = False
        If DatosADevolverBusqueda = "" Then
            PonerModo 0
        Else
            If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                BotonAnyadir
            Else
                PonerModo 1 'búsqueda
                ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
                Text1(0).BackColor = vbYellow 'codforfait
                ' ****************************************************************************
            End If
        End If
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia(0).Value
    If Modo = 4 Then TerminaBloquear
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim i As Integer

    PrimeraVez = True
    
    ' ICONETS DE LA BARRA
'    btnPrimero = 18 'index del botó "primero"
'    With Me.Toolbar1
'        .HotImageList = frmPpal.imgListComun_OM
'        .DisabledImageList = frmPpal.imgListComun_BN
'        .ImageList = frmPpal.imgListComun
'        'l'1 i el 2 son separadors
'        .Buttons(3).Image = 1   'Buscar
'        .Buttons(4).Image = 2   'Totss
'        'el 5 i el 6 son separadors
'        .Buttons(7).Image = 3   'Insertar
'        .Buttons(8).Image = 4   'Modificar
'        .Buttons(9).Image = 5   'Borrar
'        .Buttons(11).Image = 19   'Expandir Añadir, Borrar y Modificar
'        .Buttons(12).Image = 26   'cambio de costes de confeccion
'        .Buttons(13).Image = 16  'Duplicar Confecciones
'        'el 10 i el 11 son separadors
'
'        .Buttons(14).Image = 10  'Imprimir
'        .Buttons(15).Image = 11  'Eixir
'        'el 13 i el 14 son separadors
'        .Buttons(btnPrimero).Image = 6  'Primer
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Següent
'        .Buttons(btnPrimero + 3).Image = 9 'Últim
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
        .Buttons(1).Image = 19   'Expandir Añadir, Borrar y Modificar
        .Buttons(2).Image = 26   'cambio de costes de confeccion
        .Buttons(3).Image = 16  'Duplicar Confecciones
    End With
    
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
    For i = 0 To ToolAux.Count - 1
        With Me.ToolAux(i)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
        End With
    Next i
    ' ***********************************
    
    'cargar IMAGES de busqueda
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    Me.imgBuscar(8).Picture = frmPpal.imgListComun.ListImages(21).Picture
   
    'carga IMAGES de mail
'    For i = 0 To Me.imgMail.Count - 1
'        Me.imgMail(i).Picture = frmPpal.imgListImages16.ListImages(2).Picture
'    Next i
    
    'IMAGES para zoom
    For i = 0 To Me.imgZoom.Count - 1
        Me.imgZoom(i).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    Next i
    
    CargaCombo
    
    
    LimpiarCampos   'Neteja els camps TextBox
    ' ******* si n'hi han llínies *******
    DataGridAux(0).ClearFields
    DataGridAux(1).ClearFields
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "forfaits"
    Ordenacion = " ORDER BY codforfait"
    
    'Mirem com està guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    Data1.ConnectionString = conn
    '***** cambiar el nombre de la PK de la cabecera *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where codforfait='-1'"
    Data1.Refresh
       
    CargaGrid 0, False
    CargaGrid 1, False
       
    ModoLineas = 0
       
    ' *** si n'hi han combos (capçalera o llínies) ***
    CargaCombo
    
'    If DatosADevolverBusqueda = "" Then
'        PonerModo 0
'    Else
'        PonerModo 1 'búsqueda
'        ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
'        Text1(0).BackColor = vbYellow 'codforfait
'        ' ****************************************************************************
'    End If
End Sub

Private Sub LimpiarCampos()
    On Error Resume Next
    
    limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    
    ' *** si n'hi han combos a la capçalera ***
    Me.Combo1(0).ListIndex = -1
    Me.Combo1(1).ListIndex = -1
    ' *****************************************

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub LimpiarCamposLin(FrameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, FrameAux  'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO s'habiliten, o no, els diversos camps del
'   formulari en funció del modo en que anem a treballar
Private Sub PonerModo(Kmodo As Byte, Optional indFrame As Integer)
Dim i As Integer, NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo
 
    'Actualisa Iconos Insertar,Modificar,Eliminar
    'ActualizarToolbar Modo, Kmodo
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo, ModoLineas
       
    'Modo 2. N'hi han datos i estam visualisant-los
    '=========================================
    'Posem visible, si es formulari de búsqueda, el botó "Regresar" quan n'hi han datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = (Modo = 2)
    Else
        cmdRegresar.visible = False
    End If
    
    Text1(5).Enabled = True
    Combo1(1).Enabled = True
    
    
    '=======================================
    b = (Modo = 2)
    'Posar Fleches de desplasament visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Només es per a saber que n'hi ha + d'1 registre
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, Numreg
    DesplazamientoVisible b And Data1.Recordset.RecordCount > 1
    '---------------------------------------------
    
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
       
    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
    'Si estem en Insertar a més neteja els camps Text1
    BloquearText1 Me, Modo
    
    '*** si n'hi han combos a la capçalera ***
    BloquearCombo Me, Modo
    '**************************
    
    ' ***** bloquejar tots els controls visibles de la clau primaria de la capçalera ***
'    If Modo = 4 Then _
'        BloquearTxt Text1(0), True 'si estic en  modificar, bloqueja la clau primaria
    ' **********************************************************************************
    
    ' **** si n'hi han imagens de buscar en la capçalera *****
    BloquearImgBuscar Me, Modo, ModoLineas
    BloquearImgZoom Me, Modo, ModoLineas
    ' ********************************************************
    Me.imgBuscar(8).Enabled = (Modo = 2)
    Me.imgBuscar(8).visible = (Modo = 2)
    Me.Label40.visible = (Modo = 2)
        
    'si cajakilo = K entonces facturar es siempre K y el valor de kiloscaj se pone por defecto
    'igual que kilosuni
    'si cajakilo = C entonces facturar puede ser K (poner valor en kiloscaj) o U (poner valor
    ' en kilosuni
    If (Modo = 3 Or Modo = 4) Then
        If Combo1(0).ListIndex = 1 Then Combo1(1).ListIndex = 1
        b = (Combo1(0).ListIndex <> 1)
        Combo1(1).Enabled = b
        Text1(5).Enabled = b
        If Combo1(1).Enabled Then Combo1(1).SetFocus
    End If
    'fin
    
    
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    PonerLongCampos

    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 0, False
        CargaGrid 1, False
    End If
    
    b = (Modo = 4) Or (Modo = 2)
    DataGridAux(0).Enabled = b
    DataGridAux(1).Enabled = b
      
    ' ****** si n'hi han combos a la capçalera ***********************
     If (Modo = 0) Or (Modo = 2) Or (Modo = 5) Then
        Combo1(0).Enabled = False
        Combo1(0).BackColor = &H80000018 'groc
        Combo1(1).Enabled = False
        Combo1(1).BackColor = &H80000018 'groc
    ElseIf (Modo = 1) Or (Modo = 3) Or (Modo = 4) Then
        Combo1(0).Enabled = True
        Combo1(0).BackColor = &H80000005 'blanc
        Combo1(1).Enabled = True
        Combo1(1).BackColor = &H80000005 'blanc
    End If
    ' ****************************************************************
    
    PonerModoOpcionesMenu (Modo) 'Activar opcions menú según modo
    PonerOpcionesMenu   'Activar opcions de menú según nivell
                        'de permisos de l'usuari

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
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los TEXT1
    PonerLongCamposGnral Me, Modo, 1
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub

Private Sub PonerModoOpcionesMenu(Modo)
'Actives unes Opcions de Menú i Toolbar según el modo en que estem
Dim b As Boolean, bAux As Boolean
Dim i As Byte
    
    'Barra de CAPÇALERA
    '------------------------------------------
    'b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    b = (Modo = 2 Or Modo = 0)
    'Buscar
    Toolbar1.Buttons(5).Enabled = b
    Me.mnBuscar.Enabled = b
    'Vore Tots
    Toolbar1.Buttons(6).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    'Insertar
    Toolbar1.Buttons(1).Enabled = b And Not DeConsulta
    Me.mnNuevo.Enabled = b And Not DeConsulta
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnEliminar.Enabled = b
    'Duplicar Confecciones
    Toolbar5.Buttons(3).Enabled = b
    Me.mnDuplicarConfec.Enabled = b
    
    'Expandir operaciones
    Toolbar5.Buttons(1).Enabled = True And Not DeConsulta
    Me.mnExpandirOperaciones.Enabled = True And Not DeConsulta
    'Cambio Costes
    Toolbar5.Buttons(2).Enabled = True And Not DeConsulta
    Me.mnCambioCostes.Enabled = True And Not DeConsulta
    'Imprimir
    Toolbar1.Buttons(8).Enabled = True And Not DeConsulta
    Me.mnImprimir.Enabled = True And Not DeConsulta
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
'++monica: si insertamos lo he quitado
'    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
    b = (Modo = 4 Or Modo = 2) And Not DeConsulta
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.AdoAux(i).Recordset.RecordCount > 0)
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i
    
End Sub

Private Sub Desplazamiento(Index As Integer)
'Botons de Desplaçament; per a desplaçar-se pels registres de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index, True
    PonerCampos
End Sub

Private Function MontaSQLCarga(Index As Integer, enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basant-se en la informació proporcionada pel vector de camps
'   crea un SQl per a executar una consulta sobre la base de datos que els
'   torne.
' Si ENLAZA -> Enlaça en el data1
'           -> Si no el carreguem sense enllaçar a cap camp
'--------------------------------------------------------------------
Dim Sql As String
Dim Tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
               
        Case 0 'ENVASES
            Sql = "SELECT forfaits_envases.codforfait, forfaits_envases.codartic, sartic.nomartic, "
            Sql = Sql & " forfaits_envases.cantidad, "
            If vParamAplic.TipoPrecio = 0 Then
                Sql = Sql & "sartic.preciomp, round(forfaits_envases.cantidad*sartic.preciomp,4) "
            Else
                Sql = Sql & "sartic.preciouc, round(forfaits_envases.cantidad*sartic.preciouc,4) "
            End If
            
            Sql = Sql & ", forfaits_envases.numlinea FROM forfaits_envases, sartic "
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE forfaits_envases.codforfait = '-1'"
            End If
            Sql = Sql & " and forfaits_envases.codartic = sartic.codartic"
            Sql = Sql & " ORDER BY forfaits_envases.numlinea"
               
        Case 1 'COSTES
            Sql = "SELECT forfaits_costes.codforfait, forfaits_costes.codcoste, nombcoste.denominacion ,forfaits_costes.importes "
            Sql = Sql & " FROM forfaits_costes, nombcoste"
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE forfaits_costes.codforfait = '-1'"
            End If
            Sql = Sql & " and forfaits_costes.codcoste = nombcoste.codcoste"
            Sql = Sql & " ORDER BY forfaits_costes.codcoste"
            
    End Select
    
    MontaSQLCarga = Sql
End Function

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Articulos
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 1) 'codartic
    txtAux2(2).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
    VisualizaPrecio
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim Aux As String
    
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabem quins camps son els que mos torna
        'Creem una cadena consulta i posem els datos
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadB = Aux
        '   Com la clau principal es única, en posar el sql apuntant
        '   al valor retornat sobre la clau ppal es suficient
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        ' **********************************
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmCap_DatoSeleccionado(CadenaSeleccion As String)
'Capacidades
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'capacidad
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub


Private Sub frmConf_DatoSeleccionado(CadenaSeleccion As String)
'Confecciones
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codconfe
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmEnv_DatoSeleccionado(CadenaSeleccion As String)
'tipos de envase
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codtipen
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmFor_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String
Dim Aux As String
      
    If CadenaSeleccion <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        CadB = "codforfait = " & DBSet(RecuperaValor(CadenaSeleccion, 1), "T")
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmMar_DatoSeleccionado(CadenaSeleccion As String)
'Marcas
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codmarca
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmMed_DatoSeleccionado(CadenaSeleccion As String)
'Medidas de envase
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codmedid
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmNCoste_DatoSeleccionado(CadenaSeleccion As String)
'Costes
    txtAux(9).Text = RecuperaValor(CadenaSeleccion, 1) 'codcoste
    txtAux2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmPal_DatoSeleccionado(CadenaSeleccion As String)
'Tipos de Palets
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codpalet
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmPres_DatoSeleccionado(CadenaSeleccion As String)
'Presentacion
    Text1(11).Text = RecuperaValor(CadenaSeleccion, 1) 'codpresentacion
    Text2(11).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Variedades
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1) 'codvariedad
    Text2(2).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(indice).Text = vCampo
End Sub

Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        indice = 3
        frmZ.pTitulo = "Observaciones de la Confección"
        frmZ.pValor = Text1(indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco Text1(indice)
    End If
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
    Combo1(0).ListIndex = -1 'quan busque, per defecte no seleccione cap tipo de client
    Combo1(1).ListIndex = -1
End Sub

Private Sub mnCambioCostes_Click()
    Screen.MousePointer = vbHourglass
    frmModConf.Show vbModal
    
    If Not Data1.Recordset.EOF Then
        PosicionarData
        PonerCampos
'        PonerCamposLineas
    End If

    
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnDuplicarConfec_Click()
    If Data1.Recordset.RecordCount = 0 Then Exit Sub

    Screen.MousePointer = vbHourglass
    frmListConfeccion.Opcionlistado = 1
    frmListConfeccion.NumCod = Text1(0).Text
    frmListConfeccion.Show vbModal
    Screen.MousePointer = vbDefault
    If Confeccion <> "" Then
        CadB = "codforfait = " & DBSet(Confeccion, "T")
        If chkVistaPrevia(0) = 1 Then
            MandaBusquedaPrevia CadB
        ElseIf CadB <> "" Then
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
            PonerCadenaBusqueda
        End If
        
        Confeccion = ""
    End If
End Sub

Private Sub mnExpandirOperaciones_Click()
    Set frmOpEnv = New frmOperaEnv
    frmOpEnv.DatosADevolverBusqueda = "0|1|"
    frmOpEnv.CodigoActual = Text1(0).Text
    frmOpEnv.Show vbModal
    Set frmOpEnv = Nothing
    PonerFoco Text1(0)
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    Screen.MousePointer = vbHourglass
    frmListConfeccion.Opcionlistado = 0
    frmListConfeccion.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnModificar_Click()
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************
    
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 5  'Búscar
           mnBuscar_Click
        Case 6  'Tots
            mnVerTodos_Click
        Case 1  'Nou
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
        Case 11  'Expandir operaciones
            mnExpandirOperaciones_Click
        Case 12 ' Cambio Costes
            mnCambioCostes_Click
        Case 13 ' Duplicar Confeccion
            mnDuplicarConfec_Click
        Case 8 'Imprimir
            mnImprimir_Click
        Case 15    'Eixir
            mnSalir_Click
            
        Case btnPrimero To btnPrimero + 3 'Fleches Desplaçament
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub BotonBuscar()
Dim i As Integer
' ***** Si la clau primaria de la capçalera no es Text1(0), canviar-ho en <=== *****
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        PonerFoco Text1(0) ' <===
        Text1(0).BackColor = vbLightBlue 'vbYellow ' <===
        ' *** si n'hi han combos a la capçalera ***
        For i = 0 To Combo1.Count - 1
            Combo1(i).ListIndex = -1
        Next i
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbLightBlue
            PonerFoco Text1(kCampo)
        End If
    End If
' ******************************************************************************
End Sub

Private Sub HacerBusqueda()

    CadB = ObtenerBusqueda2(Me, 1)
    
    If chkVistaPrevia(0) = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        ' *** foco al 1r camp visible de la capçalera que siga clau primaria ***
        PonerFoco Text1(0)
        ' **********************************************************************
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)

    Set frmFor = New frmBasico2
    
    AyudaForfaits frmFor
    
    Set frmFor = Nothing
   
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String
Dim Aux As String
Dim i As Integer
Dim J As Integer

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    Cad = ""
    i = 0
    Do
        J = i + 1
        i = InStr(J, DatosADevolverBusqueda, "|")
        If i > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, i - J)
            J = Val(Aux)
            Cad = Cad & Text1(J).Text & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub

Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        PonerModo 2
        'Data1.Recordset.MoveLast
        Data1.Recordset.MoveFirst
        PonerCampos
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub

EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub BotonVerTodos()
'Vore tots
    LimpiarCampos 'Neteja els Text1
    CadB = ""
    
    If chkVistaPrevia(0).Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

Private Sub BotonAnyadir()
Dim NumF As String
    LimpiarCampos 'Huida els TextBox
    PonerModo 3
    
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la capçalera *******
'    text1(0).Text = SugerirCodigoSiguienteStr("forfaits", "codforfait")
'    FormateaCampo text1(0)
    '******************** canviar taula i camp **************************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        NumF = NuevoCodigo
    Else
        NumF = ""
    End If
    '********************************************************************


       
    PosicionarCombo Combo1(0), 0
    PosicionarCombo Combo1(1), 0
        
            
    Text1(0) = NumF
    PonerFoco Text1(0) '*** 1r camp visible que siga PK ***
    
    ' *** si n'hi han camps de descripció a la capçalera ***
    'PosarDescripcions

End Sub

Private Sub BotonModificar()

    PonerModo 4

    ' *** bloquejar els camps visibles de la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco Text1(1)
End Sub

Private Sub BotonEliminar()
Dim Cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
    ' ***************************************************************************

    ' *************** canviar la pregunta ****************
    Cad = "¿Seguro que desea eliminar el Forfait?"
    Cad = Cad & vbCrLf & "Código: " & Data1.Recordset.Fields(0)
    Cad = Cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(1)
    
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Proveedor", Err.Description
End Sub

Private Sub PonerCampos()
Dim i As Integer
Dim codpobla As String, despobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la capçalera
    
    ' *** si n'hi han llínies en datagrids ***
    'For i = 0 To DataGridAux.Count - 1
    For i = 0 To 1
            CargaGrid i, True
            If Not AdoAux(i).Recordset.EOF Then _
                PonerCamposForma2 Me, AdoAux(i), 2, "FrameAux" & i
    Next i

    
    ' ************* configurar els camps de les descripcions de la capçalera *************
    Text2(2).Text = PonerNombreDeCod(Text1(2), "variedades", "nomvarie")
    Text2(7).Text = PonerNombreDeCod(Text1(7), "confenva", "nomtipen")
    Text2(8).Text = PonerNombreDeCod(Text1(8), "capacida", "nomcapac")
    Text2(9).Text = PonerNombreDeCod(Text1(9), "confmedi", "nommedid")
    Text2(10).Text = PonerNombreDeCod(Text1(10), "conftipo", "nomtipco")
    Text2(11).Text = PonerNombreDeCod(Text1(11), "confpres", "nomprese")
    Text2(12).Text = PonerNombreDeCod(Text1(12), "marcas", "nommarca")
    Text2(13).Text = PonerNombreDeCod(Text1(13), "confpale", "nompalet")
    ' ********************************************************************************
    
    CalcularTotales
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu
End Sub

Private Sub cmdCancelar_Click()
Dim i As Integer
Dim V

    Select Case Modo
        Case 1, 3 'Búsqueda, Insertar
                LimpiarCampos
                If Data1.Recordset.EOF Then
                    PonerModo 0
                Else
                    PonerModo 2
                    PonerCampos
                End If
                ' *** foco al primer camp visible de la capçalera ***
                PonerFoco Text1(0)

        Case 4  'Modificar
                TerminaBloquear
                PonerModo 2
                PonerCampos
                ' *** primer camp visible de la capçalera ***
                PonerFoco Text1(0)
        
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 1 'afegir llínia
                    ModoLineas = 0
                    ' *** les llínies que tenen datagrid (en o sense tab) ***
                    If NumTabMto = 0 Or NumTabMto = 1 Or NumTabMto = 2 Or NumTabMto = 4 Then
                        DataGridAux(NumTabMto).AllowAddNew = False
                        ' **** repasar si es diu Data1 l'adodc de la capçalera ***
                        'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'Modificar
                        LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                        DataGridAux(NumTabMto).Enabled = True
                        DataGridAux(NumTabMto).SetFocus

                        ' *** si n'hi han camps de descripció dins del grid, els neteje ***
                        'txtAux2(2).text = ""

                    End If
                    
'                    ' *** si n'hi han tabs ***
'                    SituarTab (NumTabMto + 1)
                    
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        AdoAux(NumTabMto).Recordset.MoveFirst
                    End If

                Case 2 'modificar llínies
                    ModoLineas = 0
                    
                    ' *** si n'hi han tabs ***
'                    SituarTab (NumTabMto + 1)
                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                    PonerModo 4
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        ' *** l'Index de Fields es el que canvie de la PK de llínies ***
                        V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
                        AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
                        ' ***************************************************************
                    End If

            End Select
            
            TerminaBloquear
            
            PosicionarData
            
            ' *** si n'hi han llínies en grids i camps fora d'estos ***
            If Not AdoAux(NumTabMto).Recordset.EOF Then
                DataGridAux_RowColChange NumTabMto, 1, 1
            Else
                LimpiarCamposFrame NumTabMto
            End If
    End Select
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm2(Me, 1)
    If Not b Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
    If (Modo = 3) Then 'insertar
        'comprobar si existe ya el cod. del campo clave primaria
        If ExisteCP(Text1(0)) Then b = False
    End If
    
    If Modo = 3 Or Modo = 4 Then
        Select Case Combo1(0).ListIndex
            Case 0  'caja
                Select Case Combo1(1).ListIndex
                    Case 0  'unidad
                        If CCur(Text1(5).Text) = 0 Then
                            MsgBox "El campo Kilos/Unidad debe tener un valor superior a cero", vbExclamation
                            b = False
                        End If
                    Case 1  'kilo
                        If CCur(Text1(4).Text) = 0 Then
                            MsgBox "El campo Kilos/Caja debe tener un valor superior a cero", vbExclamation
                            b = False
                        End If
                End Select
            Case 1  'kilo
                If ComprobarCero(Text1(4).Text) = 0 Then
                    MsgBox "El campo Kilos/Caja debe tener un valor superior a cero", vbExclamation
                    b = False
                End If
            
        End Select
    
    End If
    ' ************************************************************************************
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    Cad = "(codforfait=" & DBSet(Text1(0).Text, "T") & ")"
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(Data1, cad, Indicador) Then
    If SituarData(Data1, Cad, Indicador) Then
        If ModoLineas <> 1 Then PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
End Sub

Private Function Eliminar() As Boolean
Dim vWhere As String

    On Error GoTo FinEliminar

    conn.BeginTrans
    ' ***** canviar el nom de la PK de la capçalera, repasar codEmpre *******
    vWhere = " WHERE codforfait=" & DBSet(Data1.Recordset!codforfait, "T")
        
    ' ***** elimina les llínies ****
    conn.Execute "DELETE FROM forfaits_envases " & vWhere
        
    conn.Execute "DELETE FROM forfaits_costes " & vWhere
        
    'Eliminar la CAPÇALERA
    conn.Execute "Delete from " & NombreTabla & vWhere
       
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub

' *** si n'hi han combos a la capçalera ***
Private Sub Combo1_GotFocus(Index As Integer)
    If Modo = 1 Then Combo1(Index).BackColor = vbYellow
  
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim b As Boolean

    If Combo1(Index).BackColor = vbYellow Then Combo1(Index).BackColor = vbWhite
    
    'si cajakilo = K entonces facturar es siempre K y el valor de kiloscaj se pone por defecto
    'igual que kilosuni
    'si cajakilo = C entonces facturar puede ser K (poner valor en kiloscaj) o U (poner valor
    ' en kilosuni
    If Index = 0 And (Modo = 3 Or Modo = 4) Then
        If Combo1(0).ListIndex = 1 Then Combo1(1).ListIndex = 1
        b = (Combo1(0).ListIndex <> 1)
        Combo1(1).Enabled = b
        Text1(5).Enabled = b
        If Combo1(1).Enabled Then Combo1(1).SetFocus
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    
    ' ***************** configurar els LostFocus dels camps de la capçalera *****************
    Select Case Index
        Case 0 'codigo de forfait
            Text1(Index).Text = UCase(Text1(Index).Text)
        
        Case 2 'Variedad
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "variedades", "nomvarie")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Variedad: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmVar = New frmManVariedad
                        frmVar.DatosADevolverBusqueda = "0|1|"
                        frmVar.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmVar.Show vbModal
                        Set frmVar = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 7 'Envase
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index) = PonerNombreDeCod(Text1(Index), "confenva", "nomtipen")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Tipo de Envase: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmEnv = New frmManTipEnv
                        frmEnv.DatosADevolverBusqueda = "0|1|"
                        frmEnv.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmEnv.Show vbModal
                        Set frmEnv = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
                
        Case 8 'Capacidad
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index) = PonerNombreDeCod(Text1(Index), "capacida", "nomcapac")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Capacidad de Envase: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCap = New frmManCapEnv
                        frmCap.DatosADevolverBusqueda = "0|1|"
                        frmCap.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmCap.Show vbModal
                        Set frmCap = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
        
        Case 9 'medida
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index) = PonerNombreDeCod(Text1(Index), "confmedi", "nommedid")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Medida de Envase: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmMed = New frmManMedEnv
                        frmMed.DatosADevolverBusqueda = "0|1|"
                        frmMed.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmMed.Show vbModal
                        Set frmMed = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
        
        Case 10 'confeccion
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index) = PonerNombreDeCod(Text1(Index), "conftipo", "nomtipco")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Confección: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmConf = New frmManFConf
                        frmConf.DatosADevolverBusqueda = "0|1|"
                        frmConf.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmConf.Show vbModal
                        Set frmConf = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
        
        Case 11 'Presentacion
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index) = PonerNombreDeCod(Text1(Index), "confpres", "nomprese")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Presentación de la Confección: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmPres = New frmManPresConf
                        frmPres.DatosADevolverBusqueda = "0|1|"
                        frmPres.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmPres.Show vbModal
                        Set frmPres = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
        
        
        Case 12 'marcas
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index) = PonerNombreDeCod(Text1(Index), "marcas", "nommarca")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Marca: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmMar = New frmManMarcas
                        frmMar.DatosADevolverBusqueda = "0|1|"
                        frmMar.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmMar.Show vbModal
                        Set frmMar = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
        
        Case 13 'palets
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index) = PonerNombreDeCod(Text1(Index), "confpale", "nompalet")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Palet: " & Text1(Index).Text & vbCrLf
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
            Else
                Text2(Index).Text = ""
            End If
        
        
            
        Case 4, 5, 6 'kilos/caja, kilo/unidad, desviacion
            If Modo = 1 Then Exit Sub
            If Index = 4 And Combo1(0).ListIndex = 1 Then
                Text1(5).Text = Text1(4).Text
                PonerFormatoDecimal Text1(5), 4
            End If
            PonerFormatoDecimal Text1(Index), 4
            
        Case 14
            If Modo = 1 Then Exit Sub
            PonerFormatoDecimal Text1(Index), 4
            
        Case 15 'cajas por palet
            PonerFormatoEntero Text1(Index)
            
        Case 16 ' precio por kilo para nominas
            If Modo = 1 Then Exit Sub
            PonerFormatoDecimal Text1(Index), 7
        
        
        
    End Select
        ' ***************************************************************************
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 7: KEYBusqueda KeyAscii, 0 'envase
                Case 8: KEYBusqueda KeyAscii, 1 'capacidad
                Case 9: KEYBusqueda KeyAscii, 2 'medida
                Case 10: KEYBusqueda KeyAscii, 3 'confeccion
                Case 11: KEYBusqueda KeyAscii, 4 'presentacion
                Case 12: KEYBusqueda KeyAscii, 5 'marca
                Case 13: KEYBusqueda KeyAscii, 6 'palet
                Case 2: KEYBusqueda KeyAscii, 7 'variedad
            End Select
        End If
    Else
        If Index <> 3 Or (Index = 3 And Text1(3).Text = "") Then KEYpress KeyAscii
    End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub



'************* LLINIES: ****************************
Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
'-- pon el bloqueo aqui
    'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
    Select Case Button.Index
        Case 1
            BotonAnyadirLinea Index
        Case 2
            BotonModificarLinea Index
        Case 3
            BotonEliminarLinea Index
        Case Else
    End Select
    'End If
End Sub

Private Sub BotonEliminarLinea(Index As Integer)
Dim Sql As String
Dim vWhere As String
Dim Eliminar As Boolean

    On Error GoTo Error2

    ModoLineas = 3 'Posem Modo Eliminar Llínia
    
    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index

    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar(Index) Then Exit Sub
    NumTabMto = Index
    Eliminar = False
   
    vWhere = ObtenerWhereCab(True)
    
    ' ***** independentment de si tenen datagrid o no,
    ' canviar els noms, els formats i el DELETE *****
    Select Case Index
        Case 0 'envases
            Sql = "¿Seguro que desea eliminar el Envase?"
            Sql = Sql & vbCrLf & "Envase: " & AdoAux(Index).Recordset!codArtic
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM forfaits_envases "
                Sql = Sql & vWhere & " AND numlinea= " & AdoAux(Index).Recordset!NumLinea
            End If
            
        Case 1 'coste
            Sql = "¿Seguro que desea eliminar el Coste Confección?"
            Sql = Sql & vbCrLf & "Nombre: " & AdoAux(Index).Recordset!codCoste
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM forfaits_costes "
                Sql = Sql & vWhere & " AND codcoste= " & AdoAux(Index).Recordset!codCoste
            End If
            
    End Select

    If Eliminar Then
        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        conn.Execute Sql
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        CargaGrid Index, True
        If Not SituarDataTrasEliminar(AdoAux(Index), NumRegElim, True) Then
'            PonerCampos
            
        End If
        CalcularTotales
        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
        ' *** si n'hi han tabs ***
'        SituarTab (NumTabMto + 1)
    End If
    
    ModoLineas = 0
    PosicionarData
    
    Exit Sub
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando linea", Err.Description
End Sub

Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vtabla As String
Dim anc As Single
Dim i As Integer
    
    ModoLineas = 1 'Posem Modo Afegir Llínia
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index
    
    ' *** bloquejar la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case Index
        Case 0: vtabla = "forfaits_envases"
        Case 1: vtabla = "forfaits_costes"
    End Select
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 0, 1 ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            
            If Index = 0 Then NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", vWhere)

            AnyadirLinea DataGridAux(Index), AdoAux(Index)
    
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
            
            LLamaLineas Index, ModoLineas, anc
        
            Select Case Index
                ' *** valor per defecte a l'insertar i formateig de tots els camps ***
                Case 0 'envases
                    txtAux(0).Text = Text1(0).Text 'codforfait
                    txtAux(3).Text = NumF 'codartic
                    txtAux(1).Text = ""
                    txtAux(2).Text = ""
                    txtAux2(0).Text = ""
                    txtAux2(1).Text = ""
                    txtAux2(2).Text = ""
                    BloquearTxt txtAux(0), False
                    BloquearTxt txtAux(3), False
                    BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux0"
                    PonerFoco txtAux(1)
                Case 1 'coste confeccion
                    txtAux(8).Text = Text1(0).Text 'codforfait
                    txtAux(9).Text = "" 'NumF 'codcoste
                    txtAux(10).Text = ""
                    txtAux2(3).Text = ""
                    For i = 9 To 9
                        BloquearTxt txtAux(i), False
                    Next i
                    BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux1"
                    PonerFoco txtAux(9)
            End Select
            
    End Select
End Sub

Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim i As Integer
    Dim J As Integer
    
    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub
    
    ModoLineas = 2 'Modificar llínia
       
    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index
    ' *** bloqueje la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
  
    Select Case Index
        Case 0, 1 ' *** pose els index de llínies que tenen datagrid (en o sense tab) ***
            If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
                i = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
                DataGridAux(Index).Scroll 0, i
                DataGridAux(Index).Refresh
            End If
              
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
    End Select
    
    Select Case Index
        ' *** valor per defecte al modificar dels camps del grid ***
        Case 0 ' envases
        
            For J = 0 To 1
                txtAux(J).Text = DataGridAux(Index).Columns(J).Text
            Next J
            txtAux2(2).Text = DataGridAux(Index).Columns(2).Text
            txtAux(2).Text = DataGridAux(Index).Columns(3).Text
            txtAux2(0).Text = DataGridAux(Index).Columns(4).Text
            txtAux2(1).Text = DataGridAux(Index).Columns(5).Text
            txtAux(3).Text = DataGridAux(Index).Columns(6).Text
            BloquearTxt txtAux(0), True
            BloquearTxt txtAux(3), True
            
            BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux0"
            
        Case 1 'coste confeccion
            For J = 8 To 9
                txtAux(J).Text = DataGridAux(Index).Columns(J - 8).Text
            Next J
            txtAux(10).Text = DataGridAux(Index).Columns(3).Text
            txtAux2(3).Text = DataGridAux(Index).Columns(2).Text
            
            For i = 9 To 9
                BloquearTxt txtAux(i), True
            Next i
            BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux1"
            
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 0 'envases
            PonerFoco txtAux(1)
        Case 1 'coste confeccion
            PonerFoco txtAux(10)
    End Select
    ' ***************************************************************************************
End Sub

Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim b As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    DeseleccionaGrid DataGridAux(Index)
       
    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
    Select Case Index
        Case 0 'envases
             For jj = 1 To 2
                txtAux(jj).visible = b
                txtAux(jj).Top = alto
            Next jj
            For jj = 0 To 2
                txtAux2(jj).visible = b
                txtAux2(jj).Top = alto
            Next jj
            btnBuscar(0).visible = b
            btnBuscar(0).Top = alto
            
        Case 1 'coste confeccion
            For jj = 9 To 10
                txtAux(jj).visible = b
                txtAux(jj).Top = alto
            Next jj
            txtAux2(3).visible = b
            txtAux2(3).Top = alto
            btnBuscar(1).visible = b
            btnBuscar(1).Top = alto
            
    End Select
End Sub

' ********* si n'hi han combos a la capçalera ************
Private Sub CargaCombo()
Dim Ini As Integer
Dim fin As Integer
Dim i As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For i = 0 To 1
        Combo1(i).Clear
    Next i
    
    Combo1(0).AddItem "Caja"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Kilo"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    
    Combo1(1).AddItem "Unidad"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Kilo"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    
End Sub

Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Expandir operaciones
            mnExpandirOperaciones_Click
        Case 2 ' Cambio Costes
            mnCambioCostes_Click
        Case 3 ' Duplicar Confeccion
            mnDuplicarConfec_Click
    End Select
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
    
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    

    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 1 ' codigo de envase
            If txtAux(Index) <> "" Then
                txtAux2(2).Text = PonerNombreDeCod(txtAux(Index), "sartic", "nomartic")
                VisualizaPrecio
                If txtAux2(2).Text = "" Then
                    cadMen = "No existe el Envase: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmArt = New frmManArtic
                        frmArt.DatosADevolverBusqueda = "0|1|"
                        frmArt.NuevoCodigo = txtAux(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmArt.Show vbModal
                        Set frmArt = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                End If
            Else
                txtAux2(2).Text = ""
            End If
        
        Case 9 ' codigo de coste
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(3).Text = PonerNombreDeCod(txtAux(Index), "nombcoste", "denominacion")
                If txtAux2(3).Text = "" Then
                    cadMen = "No existe el Código de Coste: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmNCoste = New frmManNomCoste
                        frmNCoste.DatosADevolverBusqueda = "0|1|"
                        frmNCoste.NuevoCodigo = txtAux(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmNCoste.Show vbModal
                        Set frmNCoste = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                End If
            Else
                txtAux2(3).Text = ""
            End If
        
        Case 2, 10 ' cantidad e importes
            If txtAux(Index).Text <> "" Then
                If PonerFormatoDecimal(txtAux(Index), 7) Then
                    CalcularPrecio
                    cmdAceptar.SetFocus
                End If
            End If
    End Select
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
   If Not txtAux(Index).MultiLine Then ConseguirFocoLin txtAux(Index)
End Sub


Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not txtAux(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not txtAux(Index).MultiLine Then
        If KeyAscii = teclaBuscar Then
            If Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) Then
                Select Case Index
                    Case 1: 'articulo
                        KeyAscii = 0
                        btnBuscar_Click (0)
                    Case 9: 'coste
                        KeyAscii = 0
                        btnBuscar_Click (1)
                End Select
            End If
        Else
            KEYpress KeyAscii
        End If
    End If
End Sub

Private Function DatosOkLlin(nomFrame As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim b As Boolean
Dim Cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte

    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False
        
    b = CompForm2(Me, 2, nomFrame) 'Comprovar formato datos ok
    If Not b Then Exit Function
    
    ' ******************************************************************************
    DatosOkLlin = b
    
EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function SepuedeBorrar(ByRef Index As Integer) As Boolean
    SepuedeBorrar = False
    
    ' *** si cal comprovar alguna cosa abans de borrar ***
'    Select Case Index
'        Case 0 'cuentas bancarias
'            If AdoAux(Index).Recordset!ctaprpal = 1 Then
'                MsgBox "No puede borrar una Cuenta Principal. Seleccione antes otra cuenta como Principal", vbExclamation
'                Exit Function
'            End If
'    End Select
    ' ****************************************************
    
    SepuedeBorrar = True
End Function

Private Sub imgBuscar_Click(Index As Integer)
    TerminaBloquear
    indice = Index + 7
     Select Case Index
        Case 0 'tipo de envases
            Set frmEnv = New frmManTipEnv
            frmEnv.DatosADevolverBusqueda = "0|1|"
            frmEnv.CodigoActual = Text1(7).Text
            frmEnv.Show vbModal
            Set frmEnv = Nothing
            PonerFoco Text1(7)
        Case 1 'capacidad
            Set frmCap = New frmManCapEnv
            frmCap.DatosADevolverBusqueda = "0|1|"
            frmCap.CodigoActual = Text1(8).Text
            frmCap.Show vbModal
            Set frmCap = Nothing
            PonerFoco Text1(8)
        Case 2 'medidas de envases
            Set frmMed = New frmManMedEnv
            frmMed.DatosADevolverBusqueda = "0|1|"
            frmMed.CodigoActual = Text1(9).Text
            frmMed.Show vbModal
            Set frmMed = Nothing
            PonerFoco Text1(9)
        Case 3 'Confeccion
            Set frmConf = New frmManFConf
            frmConf.DatosADevolverBusqueda = "0|1|"
            frmConf.CodigoActual = Text1(10).Text
            frmConf.Show vbModal
            Set frmConf = Nothing
            PonerFoco Text1(10)
        Case 4 'Presentacion
            Set frmPres = New frmManPresConf
            frmPres.DatosADevolverBusqueda = "0|1|"
            frmPres.CodigoActual = Text1(11).Text
            frmPres.Show vbModal
            Set frmPres = Nothing
            PonerFoco Text1(11)
        Case 5 'Marca
            Set frmMar = New frmManMarcas
            frmMar.DatosADevolverBusqueda = "0|1|"
            frmMar.CodigoActual = Text1(12).Text
            frmMar.Show vbModal
            Set frmMar = Nothing
            PonerFoco Text1(12)
        Case 6 'Palet
            Set frmPal = New frmManPaleConf
            frmPal.DatosADevolverBusqueda = "0|1|"
            frmPal.CodigoActual = Text1(13).Text
            frmPal.Show vbModal
            Set frmPal = Nothing
            PonerFoco Text1(13)
        Case 7 'Variedades
            Set frmVar = New frmManVariedad
            frmVar.DatosADevolverBusqueda = "0|1|"
            frmVar.CodigoActual = Text1(2).Text
            frmVar.Show vbModal
            Set frmVar = Nothing
            PonerFoco Text1(2)
        Case 8 'codigos ean de este forfait
            Set frmCEan = New frmCodEAN
            frmCEan.Tipo = 1
            frmCEan.CodigoActual = CStr(Me.Data1.Recordset!codforfait)
            frmCEan.Show vbModal
            Set frmCEan = Nothing
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub

Private Sub DataGridAux_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
Dim i As Byte

    If ModoLineas <> 1 Then
        Select Case Index
            Case 0 'cuentas bancarias
                If DataGridAux(Index).Columns.Count > 2 Then
'                    txtAux(11).Text = DataGridAux(Index).Columns("direccio").Text
'                    txtAux(12).Text = DataGridAux(Index).Columns("observac").Text
                End If
                
            Case 1 'departamentos
                If DataGridAux(Index).Columns.Count > 2 Then
'                    txtAux(21).Text = DataGridAux(Index).Columns(5).Text
'                    txtAux(22).Text = DataGridAux(Index).Columns(6).Text
'                    txtAux(23).Text = DataGridAux(Index).Columns(8).Text
'                    txtAux(24).Text = DataGridAux(Index).Columns(15).Text
'                    txtAux2(22).Text = DataGridAux(Index).Columns(7).Text
                End If
                
        End Select
        
    Else 'vamos a Insertar
        Select Case Index
            Case 0 'cuentas bancarias
'                txtAux(11).Text = ""
'                txtAux(12).Text = ""
            Case 1 'departamentos
                For i = 21 To 24
'                   txtAux(i).Text = ""
                Next i
'               txtAux2(22).Text = ""
            Case 2 'Tarjetas
'               txtAux(50).Text = ""
'               txtAux(51).Text = ""
        End Select
    End If
End Sub

' ***** si n'hi han varios nivells de tabs *****
'Private Sub SituarTab(numTab As Integer)
'    On Error Resume Next
'
'    SSTab1.Tab = numTab
'
'    If Err.Number <> 0 Then Err.Clear
'End Sub
' **********************************************

Private Sub CargaFrame(Index As Integer, enlaza As Boolean)
Dim tip As Integer
Dim i As Byte

    AdoAux(Index).ConnectionString = conn
    AdoAux(Index).RecordSource = MontaSQLCarga(Index, enlaza)
    AdoAux(Index).CursorType = adOpenDynamic
    AdoAux(Index).LockType = adLockPessimistic
    AdoAux(Index).Refresh
    
    If Not AdoAux(Index).Recordset.EOF Then
        PonerCamposForma2 Me, AdoAux(Index), 2, "FrameAux" & Index
    Else
        ' *** si n'hi han tabs sense datagrids, li pose els valors als camps ***
        NetejaFrameAux "FrameAux3" 'neteja només lo que te TAG
    End If
End Sub

' *** si n'hi han tabs sense datagrids ***
Private Sub NetejaFrameAux(nom_frame As String)
Dim Control As Object
    
    For Each Control In Me.Controls
        If (Control.Tag <> "") Then
            If (Control.Container.Name = nom_frame) Then
                If TypeOf Control Is TextBox Then
                    Control.Text = ""
                ElseIf TypeOf Control Is ComboBox Then
                    Control.ListIndex = -1
                End If
            End If
        End If
    Next Control

End Sub

Private Sub CargaGrid(Index As Integer, enlaza As Boolean)
Dim b As Boolean
Dim i As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)

    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
    
    Select Case Index
        Case 0 'envases
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;" 'codforfait
            tots = tots & "S|txtAux(1)|T|Envase|1900|;S|btnBuscar(0)|B|||;"
            tots = tots & "S|txtAux2(2)|T|Denominación|4200|;S|txtAux(2)|T|Cantidad|1200|;"
            tots = tots & "S|txtAux2(0)|T|Precio|1200|;S|txtAux2(1)|T|Importe|1500|;N||||0|;"
            
            arregla tots, DataGridAux(Index), Me, 350
        
            DataGridAux(0).Columns(4).NumberFormat = "###,##0.0000"
            DataGridAux(0).Columns(4).Alignment = dbgRight
            DataGridAux(0).Columns(5).NumberFormat = "###,##0.0000"
            DataGridAux(0).Columns(5).Alignment = dbgRight
        
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
        Case 1 'Costes
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;S|txtAux(9)|T|Coste|1000|;S|btnBuscar(1)|B||195|;" 'codsocio,numlinea
            tots = tots & "S|txtAux2(3)|T|Denominación|2800|;"
            tots = tots & "S|txtAux(10)|T|Importe|1500|;"

            arregla tots, DataGridAux(Index), Me, 350
            
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
    End Select
    
    DataGridAux(Index).ScrollBars = dbgAutomatic
      
    ' **** si n'hi han llínies en grids i camps fora d'estos ****
'    If Not AdoAux(Index).Recordset.EOF Then
'        DataGridAux_RowColChange Index, 1, 1
'    Else
''        LimpiarCamposFrame Index
'    End If
      
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub

Private Sub InsertarLinea()
'Inserta registre en les taules de Llínies
Dim nomFrame As String
Dim b As Boolean

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomFrame = "FrameAux0" 'envases
        Case 1: nomFrame = "FrameAux1" 'costes
    End Select
    
    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomFrame) Then
            b = BLOQUEADesdeFormulario2(Me, Data1, 1)
            Select Case NumTabMto
                Case 0, 1 ' *** els index de les llinies en grid (en o sense tab) ***
                     CargaGrid NumTabMto, True
                    If b Then BotonAnyadirLinea NumTabMto
            End Select
           
'            SituarTab (NumTabMto + 1)
        End If
    End If
End Sub

Private Function ModificarLinea() As Boolean
'Modifica registre en les taules de Llínies
Dim nomFrame As String
Dim V As Integer
    
    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomFrame = "FrameAux0" 'envases
        Case 1: nomFrame = "FrameAux1" 'costes
    End Select
    ModificarLinea = False
    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomFrame) Then
            ModoLineas = 0
            
            Select Case NumTabMto
                Case 0
                    V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
                Case 1
                    V = AdoAux(NumTabMto).Recordset.Fields(2) 'el 2 es el nº de llinia
            End Select
            CargaGrid NumTabMto, True
            
            ' *** si n'hi han tabs ***
'            SituarTab (NumTabMto + 1)

            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
            PonerFocoGrid Me.DataGridAux(NumTabMto)
            AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
            
            LLamaLineas NumTabMto, 0
            ModificarLinea = True
        End If
    End If
End Function

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " codforfait=" & DBSet(Me.Data1.Recordset!codforfait, "T")
    
    ObtenerWhereCab = vWhere
End Function

'' *** neteja els camps dels tabs de grid que
''estan fora d'este, i els camps de descripció ***
Private Sub LimpiarCamposFrame(Index As Integer)
    On Error Resume Next
 
'    Select Case Index
'        Case 0 'Cuentas Bancarias
'            txtAux(11).Text = ""
'            txtAux(12).Text = ""
'        Case 1 'Departamentos
'            txtAux(21).Text = ""
'            txtAux(22).Text = ""
'            txtAux2(22).Text = ""
'            txtAux(23).Text = ""
'            txtAux(24).Text = ""
'        Case 2 'Tarjetas
'            txtAux(50).Text = ""
'            txtAux(51).Text = ""
'        Case 4 'comisiones
'            txtAux2(2).Text = ""
'    End Select
'
    If Err.Number <> 0 Then Err.Clear
End Sub

'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGridAux_GotFocus(Index As Integer)
'  WheelHook DataGridAux(Index)
'End Sub
'Private Sub DataGridAux_LostFocus(Index As Integer)
'  WheelUnHook
'End Sub

Private Sub CalcularPrecio()
    txtAux2(1).Text = ""
    If txtAux(2).Text <> "" And txtAux2(0).Text <> "" Then
        txtAux2(1).Text = Round2(ImporteSinFormato(txtAux(2).Text) * ImporteSinFormato(txtAux2(0).Text), 4)
    End If
End Sub

Private Sub VisualizaPrecio()
    Select Case vParamAplic.TipoPrecio
        Case 0
            txtAux2(0).Text = DevuelveDesdeBDNew(cAgro, "sartic", "preciomp", "codartic", txtAux(1), "T")
        Case 1
            txtAux2(0).Text = DevuelveDesdeBDNew(cAgro, "sartic", "preciouc", "codartic", txtAux(1), "T")
    End Select
End Sub

Private Sub CalcularTotales()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim TotalEnvases As String
Dim TotalCostes As String
Dim Valor As Currency
Dim Kilos As Currency

    On Error Resume Next

    'total importes de envases para ese forfait
    Sql = "select sum(round(cantidad * "
    If vParamAplic.TipoPrecio = 0 Then 'precio medio ponderado
        Sql = Sql & " preciomp,4))"
    Else 'precio ultima compra
        Sql = Sql & " preciouc,4))"
    End If
    
    Sql = Sql & " from forfaits_envases, sartic where codforfait = " & DBSet(Text1(0).Text, "T")
    Sql = Sql & " and forfaits_envases.codartic = sartic.codartic"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    TotalEnvases = 0
    If Not Rs.EOF Then
        If Rs.Fields(0).Value > 0 Then TotalEnvases = Rs.Fields(0).Value
    End If
    Rs.Close
    Set Rs = Nothing
    
    Text3(0).Text = TotalEnvases
    Valor = CCur(TransformaPuntosComas(Text3(0).Text))
    If Valor <> 0 Then
        Text3(0).Text = Format(Valor, "###,###,##0.0000")
    Else
        Text3(0).Text = ""
    End If
    
    
    'total costes para ese forfait
    Sql = "select sum(importes) "
    Sql = Sql & " from forfaits_costes where codforfait = " & DBSet(Text1(0).Text, "T")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    TotalCostes = 0
    If Not Rs.EOF Then
        If Rs.Fields(0).Value > 0 Then TotalCostes = Rs.Fields(0).Value
    End If
    Rs.Close
    Set Rs = Nothing
    
    Text3(1).Text = TotalCostes
    Valor = CCur(TransformaPuntosComas(Text3(1).Text))
    If Valor <> 0 Then
        Text3(1).Text = Format(Valor, "###,###,##0.0000")
    Else
        Text3(1).Text = ""
    End If
    
    Text3(2).Text = Round(CCur(TotalEnvases) + CCur(TotalCostes), 4)
    Valor = CCur(TransformaPuntosComas(Text3(2).Text))
    If Valor <> 0 Then
        Text3(2).Text = Format(Valor, "###,###,##0.0000")
    Else
        Text3(2).Text = ""
    End If
    
    '[Monica]17/10/2017: sacamos el coste por caja o por kilo dependiendo de como sea la confeccion
    Text3(3).Text = ""
    If Combo1(0).ListIndex = 0 Then ' caja
        Kilos = CCur(TransformaPuntosComas(ComprobarCero(Text1(4).Text)))
        If Kilos <> 0 Then Text3(3).Text = Round2(Valor / Kilos, 4)
    Else 'kilo
'        Kilos = CCur(TransformaPuntosComas(ComprobarCero(Text1(5).Text)))
'        If Kilos <> 0 Then Text3(3).Text = Round2(Valor / Kilos, 4)
        Text3(3).Text = Valor
    End If
    If CCur(ComprobarCero(Text3(3).Text)) <> 0 Then
        Text3(3).Text = Format(Text3(3).Text, "###,###,##0.0000")
    Else
        Text3(3).Text = ""
    End If
    
    
    If Err.Number <> 0 Then
        Err.Clear
    End If

End Sub
