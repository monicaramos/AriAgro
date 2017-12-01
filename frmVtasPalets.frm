VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmVtasPalets 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gestión de Palets"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   12120
   Icon            =   "frmVtasPalets.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmVtasPalets.frx":000C
   ScaleHeight     =   7650
   ScaleWidth      =   12120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   13
      Left            =   9180
      MaxLength       =   30
      TabIndex        =   74
      Tag             =   "NumCajas|N|N|||palets_variedad|numcajas|#,##0|N|"
      Text            =   "Cajas"
      Top             =   4500
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   0
      Left            =   8460
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   72
      Text            =   "Text2"
      Top             =   5175
      Width           =   1245
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
      Height          =   1410
      Left            =   7110
      TabIndex        =   66
      Top             =   5580
      Width           =   3930
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   1845
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   71
         Text            =   "Text2"
         Top             =   540
         Width           =   1515
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   6
         Left            =   1845
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   70
         Text            =   "Text2"
         Top             =   945
         Width           =   1515
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   5
         Left            =   1845
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   68
         Text            =   "Text2"
         Top             =   180
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "Total Tara Palet"
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
         Height          =   330
         Left            =   180
         TabIndex        =   73
         Top             =   990
         Width           =   1500
      End
      Begin VB.Label Label1 
         Caption         =   "Total Taras Envases"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   12
         Left            =   180
         TabIndex        =   69
         Top             =   225
         Width           =   1530
      End
      Begin VB.Line Line1 
         X1              =   1530
         X2              =   3600
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Tara Palet"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   67
         Top             =   585
         Width           =   945
      End
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   2
      Left            =   9900
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   62
      Text            =   "Text2"
      Top             =   5175
      Width           =   1245
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   1
      Left            =   7110
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   60
      Text            =   "Text2"
      Top             =   5175
      Width           =   1275
   End
   Begin VB.TextBox txtAux3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   12
      Left            =   7455
      MaxLength       =   30
      TabIndex        =   59
      Tag             =   "Marca|N|N|||palets_variedad|codmarca|000||"
      Text            =   "nom forf"
      Top             =   4500
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.TextBox txtAux3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   11
      Left            =   5925
      MaxLength       =   30
      TabIndex        =   58
      Tag             =   "Marca|N|N|||palets_variedad|codmarca|000||"
      Text            =   "nom marca"
      Top             =   4500
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
      Left            =   5625
      MaxLength       =   30
      TabIndex        =   51
      Tag             =   "Marca|N|N|||palets_variedad|codmarca|000||"
      Text            =   "marca"
      Top             =   4500
      Visible         =   0   'False
      Width           =   405
   End
   Begin MSComctlLib.Toolbar ToolAux 
      Height          =   390
      Index           =   0
      Left            =   135
      TabIndex        =   50
      Top             =   2880
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
      Height          =   315
      Index           =   10
      Left            =   11055
      MaxLength       =   30
      TabIndex        =   49
      Tag             =   "Peso Neto|N|S|||palets_variedad|pesoneto|###,##0|N|"
      Text            =   "peso neto"
      Top             =   4500
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAux3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   9
      Left            =   10110
      MaxLength       =   30
      TabIndex        =   48
      Tag             =   "Peso Bruto|N|N|||palets_variedad|pesobrut|###,##0||"
      Text            =   "peso bruto"
      Top             =   4500
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAux3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   8
      Left            =   8220
      MaxLength       =   30
      TabIndex        =   47
      Tag             =   "Categoria|T|S|||palets_variedad|categori|||"
      Text            =   "categ"
      Top             =   4500
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAux3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   7
      Left            =   6825
      MaxLength       =   30
      TabIndex        =   46
      Tag             =   "Forfait|N|N|||palets_variedad|codforfait|||"
      Text            =   "forfait"
      Top             =   4500
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.TextBox txtAux3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   5
      Left            =   4710
      MaxLength       =   30
      TabIndex        =   45
      Text            =   "nom.var.comer"
      Top             =   4500
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAux3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   4
      Left            =   3810
      MaxLength       =   30
      TabIndex        =   44
      Tag             =   "Variedad Comercial|N|N|||palets_variedad|codvarco|||"
      Text            =   "var.comer."
      Top             =   4500
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAux3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   3
      Left            =   2865
      MaxLength       =   30
      TabIndex        =   43
      Text            =   "nomvarie"
      Top             =   4500
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   6
      Left            =   6375
      MaxLength       =   30
      TabIndex        =   35
      Tag             =   "Num.Cajas|N|N|0||palets_calibre|numcajas|#,##0||"
      Text            =   "numcajas"
      Top             =   6570
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   5
      Left            =   5625
      MaxLength       =   5
      TabIndex        =   34
      Text            =   "nomca"
      Top             =   6570
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   255
      MaxLength       =   12
      TabIndex        =   33
      Tag             =   "Num.Palet|N|N|||palets_calibre|numpalet||S|"
      Text            =   "numpalet"
      Top             =   6570
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   1215
      MaxLength       =   12
      TabIndex        =   32
      Tag             =   "Num.Linea|N|N|||palets_calibre|numlinea|00|N|"
      Text            =   "numlinea"
      Top             =   6570
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   2025
      MaxLength       =   12
      TabIndex        =   31
      Tag             =   "Num.Linea 1|N|N|||palets_calibre|numline1||N|"
      Text            =   "numline1"
      Top             =   6570
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   3
      Left            =   3105
      MaxLength       =   12
      TabIndex        =   30
      Tag             =   "Variedad|N|N|||palets_calibre|codvarie|000000|N|"
      Text            =   "variedad"
      Top             =   6570
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   4
      Left            =   4935
      MaxLength       =   5
      TabIndex        =   29
      Tag             =   "Calibre|N|N|||palets_calibre|codcalib|00|N|"
      Text            =   "calib"
      Top             =   6570
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtAux3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   450
      MaxLength       =   7
      TabIndex        =   28
      Tag             =   "Num.Palet|N|N|||palets_variedad|numpalet||S|"
      Text            =   "numpale"
      Top             =   4500
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtAux3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   1035
      MaxLength       =   15
      TabIndex        =   27
      Tag             =   "Num.Linea|N|N|||palets_variedad|numlinea|00|S|"
      Text            =   "numlinea"
      Top             =   4500
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAux3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   26
      Tag             =   "Variedad|N|N|||palets_variedad|variedad||N|"
      Text            =   "variedad"
      Top             =   4500
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Height          =   2235
      Left            =   135
      TabIndex        =   23
      Top             =   540
      Width           =   11820
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   20
         Left            =   8145
         MaxLength       =   6
         TabIndex        =   86
         Tag             =   "Cod.Camara|N|N|0|999|palets|codcamara|000||"
         Text            =   "Text1"
         Top             =   1710
         Width           =   780
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   20
         Left            =   8955
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   85
         Text            =   "Text2"
         Top             =   1710
         Width           =   1785
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   19
         Left            =   10770
         MaxLength       =   15
         TabIndex        =   84
         Tag             =   "IdPalet|N|S|||palets|idpalet|||"
         Text            =   "Text3"
         Top             =   1710
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   18
         Left            =   10770
         MaxLength       =   2
         TabIndex        =   13
         Tag             =   "Lin.Coste|N|N|||palets|linsalida|00||"
         Text            =   "Text3"
         Top             =   1095
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   17
         Left            =   9360
         MaxLength       =   2
         TabIndex        =   12
         Tag             =   "Lin.Coste|N|N|||palets|linentrada|00||"
         Text            =   "Text3"
         Top             =   1095
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   16
         Left            =   10770
         MaxLength       =   2
         TabIndex        =   8
         Tag             =   "Lin.Coste|N|N|||palets|codlinconf|00||"
         Text            =   "Text3"
         Top             =   450
         Width           =   915
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   13
         Left            =   5445
         MaxLength       =   10
         TabIndex        =   9
         Tag             =   "Fecha Confeccion|F|N|||palets|fechaconf|dd/mm/yyyy||"
         Top             =   1080
         Width           =   1245
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   10
         Left            =   9360
         MaxLength       =   40
         TabIndex        =   7
         Text            =   "Text3"
         Top             =   450
         Width           =   915
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   9
         Left            =   6885
         MaxLength       =   8
         TabIndex        =   5
         Text            =   "Text3"
         Top             =   450
         Width           =   810
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   8145
         MaxLength       =   10
         TabIndex        =   6
         Tag             =   "Fecha Fin|F|S|||palets|fechafin|dd/mm/yyyy||"
         Top             =   450
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   9495
         MaxLength       =   40
         TabIndex        =   55
         Tag             =   "Hora Fin|FH|S|||palets|horafin|yyyy-mm-dd hh:mm:ss||"
         Text            =   "Text3"
         Top             =   450
         Width           =   810
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   5445
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Fecha Inicio|F|N|||palets|fechaini|dd/mm/yyyy||"
         Top             =   450
         Width           =   1245
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   8
         Left            =   6885
         MaxLength       =   40
         TabIndex        =   52
         Tag             =   "Hora Inicio|FH|N|||palets|horaini|yyyy-mm-dd hh:mm:ss||"
         Text            =   "Text3"
         Top             =   450
         Width           =   810
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   2970
         TabIndex        =   2
         Tag             =   "Nº Pedido|N|S|||palets|numpedid|000000||"
         Text            =   "Text3"
         Top             =   450
         Width           =   765
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   6300
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   41
         Text            =   "Text2"
         Top             =   1710
         Width           =   1785
      End
      Begin VB.TextBox Text1 
         Height          =   1005
         Index           =   7
         Left            =   225
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Tag             =   "Observaciones|T|S|||palets|observac|||"
         Top             =   1080
         Width           =   5025
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   3915
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "Tipo Mercancia|N|N|||palets|tipmercan|0||"
         Top             =   450
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   5445
         MaxLength       =   6
         TabIndex        =   14
         Tag             =   "Cod. Palet|N|N|0|999|palets|codpalet|000||"
         Text            =   "Text1"
         Top             =   1710
         Width           =   780
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1410
         MaxLength       =   2
         TabIndex        =   1
         Tag             =   "Lin.Confe|N|N|||palets|linconfe|00||"
         Text            =   "Text3"
         Top             =   450
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   0
         Left            =   225
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Nº Palet|N|S|||palets|numpalet|0000000|S|"
         Text            =   "Text1 7"
         Top             =   450
         Width           =   980
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   12
         Left            =   6885
         MaxLength       =   8
         TabIndex        =   10
         Text            =   "Text3"
         Top             =   1080
         Width           =   810
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   11
         Left            =   8130
         MaxLength       =   40
         TabIndex        =   11
         Text            =   "Text3"
         Top             =   1080
         Width           =   810
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   14
         Left            =   6885
         MaxLength       =   40
         TabIndex        =   78
         Tag             =   "Hora Inicio conf|FH|N|||palets|horaiconf|yyyy-mm-dd hh:mm:ss||"
         Text            =   "Text3"
         Top             =   1080
         Width           =   810
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   15
         Left            =   8145
         MaxLength       =   40
         TabIndex        =   79
         Tag             =   "Hora Fin Confec|FH|N|||palets|horafconf|yyyy-mm-dd hh:mm:ss||"
         Text            =   "Text3"
         Top             =   1080
         Width           =   810
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   8955
         ToolTipText     =   "Buscar Palet"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cámara"
         Height          =   255
         Index           =   19
         Left            =   8145
         TabIndex        =   87
         Top             =   1440
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "IdPalet"
         Height          =   255
         Index           =   18
         Left            =   10770
         TabIndex        =   83
         Top             =   1470
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "L.Salida"
         Height          =   255
         Index           =   17
         Left            =   10770
         TabIndex        =   82
         Top             =   870
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "L.Entrada"
         Height          =   255
         Index           =   16
         Left            =   9360
         TabIndex        =   81
         Top             =   870
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Lin.Coste"
         Height          =   255
         Index           =   15
         Left            =   10770
         TabIndex        =   80
         Top             =   195
         Width           =   675
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   11460
         ToolTipText     =   "Buscar lineas confección"
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   2370
         ToolTipText     =   "Buscar lineas confección"
         Top             =   210
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   3480
         ToolTipText     =   "Buscar Pedidos sin albarán"
         Top             =   210
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   6435
         Picture         =   "frmVtasPalets.frx":0A0E
         ToolTipText     =   "Buscar fecha"
         Top             =   810
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Hora Inicio"
         Height          =   255
         Index           =   14
         Left            =   6885
         TabIndex        =   77
         Top             =   855
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "F.Confección"
         Height          =   255
         Index           =   13
         Left            =   5445
         TabIndex        =   76
         Top             =   855
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Hora Fin"
         Height          =   255
         Index           =   8
         Left            =   8145
         TabIndex        =   75
         Top             =   855
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Fin"
         Height          =   255
         Index           =   3
         Left            =   8145
         TabIndex        =   57
         Top             =   225
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Hora Fin"
         Height          =   255
         Index           =   4
         Left            =   9330
         TabIndex        =   56
         Top             =   225
         Width           =   915
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   8955
         Picture         =   "frmVtasPalets.frx":0A99
         ToolTipText     =   "Buscar fecha"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicio"
         Height          =   255
         Index           =   29
         Left            =   5445
         TabIndex        =   54
         Top             =   225
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Hora Inicio"
         Height          =   255
         Index           =   2
         Left            =   6885
         TabIndex        =   53
         Top             =   225
         Width           =   780
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   6435
         Picture         =   "frmVtasPalets.frx":0B24
         ToolTipText     =   "Buscar fecha"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Pedido"
         Height          =   255
         Index           =   6
         Left            =   2970
         TabIndex        =   42
         Top             =   225
         Width           =   540
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   1440
         ToolTipText     =   "Zoom descripción"
         Top             =   810
         Width           =   240
      End
      Begin VB.Label Label29 
         Caption         =   "Observaciones"
         Height          =   255
         Left            =   225
         TabIndex        =   40
         Top             =   810
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Linea Confec."
         Height          =   255
         Index           =   5
         Left            =   1395
         TabIndex        =   39
         Top             =   225
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Mercancia"
         Height          =   255
         Index           =   27
         Left            =   3915
         TabIndex        =   38
         Top             =   225
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Palet"
         Height          =   255
         Index           =   0
         Left            =   5445
         TabIndex        =   25
         Top             =   1440
         Width           =   810
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   6300
         ToolTipText     =   "Buscar Palet"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Palet"
         Height          =   255
         Index           =   28
         Left            =   225
         TabIndex        =   24
         Top             =   180
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   7065
      Width           =   2175
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   20
         Top             =   180
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9990
      TabIndex        =   17
      Top             =   7170
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8820
      TabIndex        =   16
      Top             =   7170
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Añadir"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Lineas Factura"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir "
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8400
         TabIndex        =   22
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   9990
      TabIndex        =   18
      Top             =   7155
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmVtasPalets.frx":0BAF
      Height          =   2025
      Left            =   135
      TabIndex        =   36
      Top             =   4950
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   3572
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmVtasPalets.frx":0BC4
      Height          =   1410
      Left            =   135
      TabIndex        =   37
      Top             =   3375
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   2487
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   945
      Top             =   7200
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
      Top             =   7200
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
   Begin VB.Label Label1 
      Caption         =   "Tara Envase"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   11
      Left            =   9900
      TabIndex        =   65
      Top             =   4950
      Width           =   1530
   End
   Begin VB.Label Label1 
      Caption         =   "="
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   10
      Left            =   9765
      TabIndex        =   64
      Top             =   5220
      Width           =   180
   End
   Begin VB.Label Label1 
      Caption         =   "Tara Envase"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   9
      Left            =   8460
      TabIndex        =   63
      Top             =   4950
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Total Cajas"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   7155
      TabIndex        =   61
      Top             =   4950
      Width           =   900
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
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
      Begin VB.Menu mnLineas 
         Caption         =   "&Lineas"
         Enabled         =   0   'False
         HelpContextID   =   2
         Shortcut        =   ^L
         Visible         =   0   'False
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
Attribute VB_Name = "frmVtasPalets"
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
Private WithEvents frmLPal As frmVtasLinPalets 'Lineas de variedades de palets
Attribute frmLPal.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmMen As frmMensajes 'Pedidos que no tienen asociado un nro de albaran
Attribute frmMen.VB_VarHelpID = -1

Private WithEvents frmMPal As frmManPaleConf 'Form Mto de Palets de confeccion
Attribute frmMPal.VB_VarHelpID = -1
Private WithEvents frmMCam As frmManCamara 'Form Mto de Camaras
Attribute frmMCam.VB_VarHelpID = -1
Private WithEvents frmBas  As frmBasico ' Lineas de confeccion
Attribute frmBas.VB_VarHelpID = -1
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

Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
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
            PonerFoco text1(0)
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco text1(0)
            
        Case 5 'Lineas Detalle
            TerminaBloquear
            BloquearTxt text2(16), True
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
    
    text1(2).Text = Format(Now, "dd/mm/yyyy")
    text1(3).Text = Format(Now, "dd/mm/yyyy")
'    text1(13).Text = Format(Now, "dd/mm/yyyy")
        
    LimpiarDataGrids
    CalcularTaraEnvase "-1"
    
    PonerFoco text1(1) '*** 1r camp visible que siga PK ***
    
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
        PonerFoco text1(0)
        text1(0).BackColor = vbYellow
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            text1(kCampo).Text = ""
            text1(kCampo).BackColor = vbYellow
            PonerFoco text1(kCampo)
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
        CadenaConsulta = "Select palets.* "
        CadenaConsulta = CadenaConsulta & "from " & NombreTabla
'        CadenaConsulta = CadenaConsulta & " WHERE scafac.codtipom='" & CodTipoMov & "'"
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index
    PonerCampos
End Sub


Private Sub BotonModificar()
Dim NroAlbar As String
Dim Cad As String
    NroAlbar = NroAlbaranAsignado(Data1.Recordset!numpalet, 0)
    If NroAlbar <> "" Then
        Cad = "El pedido asociado a este palet se encuentra asignado al albarán " & NroAlbar & "." & vbCrLf
        Cad = Cad & "                     ¿ Desea continuar ?"
        If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            TerminaBloquear
            Exit Sub
        End If
    End If

    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
        
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
    text2(16).Text = DataGrid1.Columns(J + 5).Text
    For J = J + 1 To 9
        txtAux(J - 1).Text = DataGrid1.Columns(J + 5).Text
    Next J
    
    ModificaLineas = 2 'Modificar
    LLamaLineas ModificaLineas, anc, "DataGrid1"
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerBotonCabecera False
    BloquearTxt text2(16), False 'Campo Ampliacion Linea
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
    
    NroAlbar = NroAlbaranAsignado(Data1.Recordset!numpalet, 0)
    If NroAlbar <> "" Then
        Cad = "El pedido asociado a este palet se encuentra asignado al albarán " & NroAlbar & "." & vbCrLf
        Cad = Cad & "         ¿ Desea continuar ?"
        If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Sub
        End If
    End If

    
    Cad = "Cabecera de Palets." & vbCrLf
    Cad = Cad & "-------------------------------------      " & vbCrLf & vbCrLf
    Cad = Cad & "Va a eliminar el Palet:            "
    Cad = Cad & vbCrLf & "Nº Palet:  " & Format(text1(0).Text, "0000000")
    Cad = Cad & vbCrLf & "Fecha:  " & Format(text1(2).Text, "dd/mm/yyyy")

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


Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim i As Byte

'    If LastCol = -1 Then Exit Sub

    If Not Data3.Recordset.EOF Then
        'Datos de la tabla palets_calibres
        CargaGrid DataGrid1, Data2, True
        CalcularTaraEnvase Data3.Recordset.Fields(1)
    Else
        'Datos de la tabla palets_calibres
        CargaGrid DataGrid1, Data2, False
        CalcularTaraEnvase "-1"
    End If
    
    
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
    'Viene de DblClick en frmAlmMovimArticulos y carga el form con los valores
'    If hcoCodMovim <> "" And Not Data1.Recordset.EOF Then PonerCadenaBusqueda
    
'    PonerCadenaBusqueda
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
Dim i As Integer

    'Icono del formulario
'    Me.Icon = frmPpal.Icon
    
     'Icono de busqueda
    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next kCampo

    ' ICONITOS DE LA BARRA
    btnPrimero = 15
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(4).Image = 3   'Insertar
        .Buttons(5).Image = 4   'Modificar
        .Buttons(6).Image = 5   'Borrar
        .Buttons(9).Image = 10 'Mto Lineas Ofertas
        .Buttons(10).Image = 10 'Imprimir Pedido
        .Buttons(12).Image = 11  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    
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
    
    CodTipoMov = "PAL" 'hcoCodTipoM
    VieneDeBuscar = False
    
    
    '[Monica]01/03/2017
    text1(19).Enabled = (vParamAplic.Cooperativa = 9)
    text1(19).visible = (vParamAplic.Cooperativa = 9)
    Label1(18).visible = (vParamAplic.Cooperativa = 9)
        
    '## A mano
    NombreTabla = "palets"
    NomTablaLineas = "palets_variedad" 'Tabla lineas de variedades
    Ordenacion = " ORDER BY palets.numpalet"
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    Data1.ConnectionString = conn
    Data1.RecordSource = "select * from palets where numpalet = -1 "
    Data1.Refresh
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
    PonerModo 0
    LimpiarDataGrids
    
    If DatosADevolverBusqueda <> "" Then
        text1(0).Text = DatosADevolverBusqueda
        HacerBusqueda
    End If
'        CargaGrid DataGrid1, Data2, False
    'Poner los grid sin apuntar a nada
    PrimeraVez = False
   
End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.Combo1(0).ListIndex = -1
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    If Modo = 4 Then TerminaBloquear
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        CadB = ""
        Aux = ValorDevueltoFormGrid(text1(0), CadenaDevuelta, 1)
        CadB = Aux
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Clientes
    text1(4).Text = RecuperaValor(CadenaSeleccion, 1)  'Cod Clien
End Sub

Private Sub frmBas_DatoSeleccionado(CadenaSeleccion As String)
    text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) ' codigo de linea de confeccion
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    text1(indice).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

' devolvemos la linea del datagrid en donde estabamos
Private Sub frmLPal_DatoSeleccionado(CadenaSeleccion As String)
Dim vWhere As String
             
   PonerCamposLineas
   
   If CadenaSeleccion = "" Then Exit Sub
             
   vWhere = "(numpalet = " & RecuperaValor(CadenaSeleccion, 1) & " and numlinea = " & RecuperaValor(CadenaSeleccion, 2) & ")"
   SituarDataMULTI Data3, vWhere, "" ', Indicador
   
   PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
   PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
   
End Sub

Private Sub frmMen_DatoSeleccionado(CadenaSeleccion As String)
    text1(5).Text = CadenaSeleccion
End Sub

Private Sub frmMPal_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Palets de confecciones
    text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Palet
    text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Palets
    text2(3).Text = RecuperaValor(CadenaSeleccion, 3) 'Peso Palet confeccion
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     text1(indice).Text = vCampo
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim Cad As String

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. de Palet
            indice = 4
            PonerFoco text1(4)
            Set frmMPal = New frmManPaleConf
            frmMPal.DatosADevolverBusqueda = "0|1|2|"
            frmMPal.Show vbModal
            Set frmMPal = Nothing
            PonerFoco text1(indice)
            
        Case 1 'Ayuda de pedidos que no tengan asignado nro de albaran
            'mostramos los palets asociados al pedido
            Set frmMen = New frmMensajes
            
            Cad = "select * from pedidos, clientes, destinos where numalbar is null "
            Cad = Cad & " and pedidos.codclien = clientes.codclien and "
            Cad = Cad & " pedidos.codclien = destinos.codclien and pedidos.coddesti = destinos.coddesti"
            
            frmMen.cadWHERE = Cad
            
            frmMen.OpcionMensaje = 20 'Pedidos que no tienen asociados un nro de albaran
            frmMen.Show vbModal
            Set frmMen = Nothing
            
            
        Case 2, 3 ' 2-Lineas de confeccion
                  ' 3-Lineas de coste de confeccion
            If Index = 2 Then
                indice = 1
            Else
                indice = 16
            End If
            PonerFoco text1(indice)
            
            Set frmBas = New frmBasico
            frmBas.DatosADevolverBusqueda = "0|1|"
            frmBas.DeConsulta = True
            frmBas.CodigoActual = text1(indice).Text
            frmBas.CadenaTots = "S|txtAux(0)|T|Código|800|;S|txtAux(1)|T|Descripción|3930|;"
            frmBas.CadenaConsulta = "SELECT cclinconf.codlinconf, cclinconf.nomlinconf "
            frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM cclinconf "
            frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
            frmBas.Tag1 = "Código|N|N|0|9999|cclinconf|codlinconf|00|S|"
            frmBas.Tag2 = "Descripción|T|N|||cclinconf|nomlinconf|||"
            frmBas.Maxlen1 = 2
            frmBas.Maxlen2 = 40
            frmBas.tabla = "cclinconf"
            frmBas.CampoCP = "codlinconf"
            frmBas.Report = "rManCCLineasConf.rpt"
            frmBas.Caption = "Lineas de Confección"
            frmBas.Show vbModal
            Set frmBas = Nothing
            
'            Set frmCCLConf = New frmCCManLineasConf
'            frmCCLConf.DatosADevolverBusqueda = "0|"
'            frmCCLConf.Show vbModal
'            Set frmCCLConf = Nothing
            
            PonerFoco text1(indice)
        
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

    imgFec(0).Tag = Index '<===
    Select Case Index
        Case 0, 1
            indice = Index + 2
        Case 2
            indice = 13
    End Select
    ' *** repasar si el camp es txtAux o Text1 ***
    If text1(indice).Text <> "" Then frmC.NovaData = text1(indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco text1(indice) '<===
    ' ********************************************
End Sub

Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        indice = 7
        frmZ.pTitulo = "Observaciones del Palet"
        frmZ.pValor = text1(indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco text1(indice)
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


Private Sub mnImprimir_Click()
'Imprimir Factura
    
    If Data1.Recordset.EOF Then Exit Sub
    
    BotonImprimir
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnLineas_Click()
    BotonMtoLineas 1, "Facturas"
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
        'bloquea la tabla cabecera de factura: scafac
        If BLOQUEADesdeFormulario(Me) Then
            'bloquear la tabla cabecera de albaranes de la factura: scafac1
            BotonModificar
        End If
    End If
End Sub


Private Function BloqueaAlbxFac() As Boolean
'bloquea todos los albaranes de la factura
Dim SQL As String

    On Error GoTo EBloqueaAlb
    
    BloqueaAlbxFac = False
    'bloquear cabecera albaranes x factura
    SQL = "select * FROM scafac1 "
    SQL = SQL & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute SQL, , adCmdText
    BloqueaAlbxFac = True

EBloqueaAlb:
    If Err.Number <> 0 Then BloqueaAlbxFac = False
End Function


Private Function BloqueaLineasFac() As Boolean
'bloquea todas las lineas de la factura
Dim SQL As String

    On Error GoTo EBloqueaLin

    BloqueaLineasFac = False
    'bloquear cabecera albaranes x factura
    SQL = "select * FROM slifac "
    SQL = SQL & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute SQL, , adCmdText
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
    If Index = 9 Then HaCambiadoCP = False 'CPostal
    If Index = 1 And Modo = 1 Then
        SendKeys "{tab}"
        Exit Sub
    End If
    ConseguirFoco text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 7 Or (Index = 7 And text1(7).Text = "") Then KEYpress KeyAscii
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
Dim SQL As String
        
    If Not PerderFocoGnral(text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1, 16 ' codigo de linea de confeccion
            If Modo = 1 Then Exit Sub
            SQL = DevuelveDesdeBDNew(cAgro, "cclinconf", "codlinconf", "codlinconf", text1(Index).Text, "N")
            If SQL = "" Then
                MsgBox "No existe la línea de confección. Revise.", vbExclamation
                PonerFoco text1(Index)
            End If
    
        Case 2, 3 'Fecha inicio y fecha de fin
            If Modo = 1 Then Exit Sub
            If text1(Index).Text <> "" Then
                '[Monica]28/08/2013: controlamos que esté dentro de campaña
                PonerFormatoFecha text1(Index), True
                If Modo = 3 And Index = 2 And text1(13).Text = "" Then text1(13).Text = text1(2).Text
            End If
                
        Case 13
            If Modo = 1 Then Exit Sub
            '[Monica]28/08/2013: controlamos que esté dentro de campaña
            If text1(Index).Text <> "" Then PonerFormatoFecha text1(Index), True
                
        Case 4 'Confeccion del Palet
            If PonerFormatoEntero(text1(Index)) Then
                text2(Index).Text = PonerNombreDeCod(text1(Index), "confpale", "nompalet")
                If text2(Index).Text = "" Then
                    cadMen = "No existe el Palet: " & text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmMPal = New frmManPaleConf
                        frmMPal.DatosADevolverBusqueda = "0|1|"
                        frmMPal.NuevoCodigo = text1(Index).Text
                        text1(Index).Text = ""
                        TerminaBloquear
                        frmMPal.Show vbModal
                        Set frmMPal = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        text1(Index).Text = ""
                    End If
                    PonerFoco text1(Index)
                Else
                    text2(3).Text = DevuelveDesdeBDNew(cAgro, "confpale", "pesopale", "codpalet", text1(4).Text, "N")
                    If text2(3).Text <> "" Then PonerFormatoDecimal text2(3), 4
                End If
            Else
                text2(Index).Text = ""
                text2(3).Text = ""
            End If
            If Modo = 4 Then CalcularTaraEnvase 1
            
        Case 0 'numero de palet
            PonerFormatoEntero text1(Index)
        
        Case 1 'linea de confeccion
            PonerFormatoEntero text1(Index)
        
'        Case 5 'numero de pedido
'            PonerFormatoEntero Text1(Index)
            
        Case 9, 10 'formato hora
            If Modo = 1 Then Exit Sub
            PonerFormatoHora text1(Index)
            
            If Modo = 3 Then
                ' si estamos insertando y no hay nada aun de horas de confeccion ponemos
                ' las de arriba inicio.
                If text1(12).Text = "" Then text1(12).Text = text1(9).Text
                If text1(11).Text = "" Then text1(11).Text = text1(10).Text
            End If
            
        Case 11, 12 'formato hora
            If Modo = 1 Then Exit Sub
            PonerFormatoHora text1(Index)
        
        '[Monica]01/12/2017:
        Case 20 'Camara
            If PonerFormatoEntero(text1(Index)) Then
                text2(Index).Text = PonerNombreDeCod(text1(Index), "camaras", "nomcamara")
                If text2(Index).Text = "" Then
                    cadMen = "No existe la Cámara: " & text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmMCam = New frmManCamara
                        frmMCam.DatosADevolverBusqueda = "0|1|"
                        frmMCam.NuevoCodigo = text1(Index).Text
                        text1(Index).Text = ""
                        TerminaBloquear
                        frmMCam.Show vbModal
                        Set frmMPal = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        text1(Index).Text = ""
                    End If
                    PonerFoco text1(Index)
                End If
            Else
                text2(Index).Text = ""
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
    If text1(9).Text <> "" Then
        text1(8).Text = text1(9).Text
        text1(8).Tag = Replace(text1(8).Tag, "FH", "FHH")
    End If
    If text1(10).Text <> "" Then
        text1(6).Text = text1(10).Text
        text1(6).Tag = Replace(text1(6).Tag, "FH", "FHH")
    End If
    If text1(12).Text <> "" Then
        text1(14).Text = text1(12).Text
        text1(14).Tag = Replace(text1(14).Tag, "FH", "FHH")
    End If
    If text1(11).Text <> "" Then
        text1(15).Text = text1(11).Text
        text1(15).Tag = Replace(text1(15).Tag, "FH", "FHH")
    End If
    
    
    CadB = ObtenerBusqueda(Me) ' antes obtenerbusqueda3(me,false)
    text1(8).Tag = Replace(text1(8).Tag, "FHH", "FH")
    text1(6).Tag = Replace(text1(6).Tag, "FHH", "FH")
    
    text1(14).Tag = Replace(text1(14).Tag, "FHH", "FH")
    text1(15).Tag = Replace(text1(15).Tag, "FHH", "FH")
    
'    '--- Laura 12/01/2007
'    Text1(5).Text = cadAux
'    '---
    
    If chkVistaPrevia = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select palets.* from " & NombreTabla & " LEFT JOIN palets_variedad ON palets.numpalet=palets_variedad.numpalet "
        CadenaConsulta = CadenaConsulta & " WHERE " & CadB & " GROUP BY palets.numpalet " & Ordenacion
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
    'Llamamos a al form
    '##A mano
    Cad = ""
    Cad = Cad & "Nº.Palet|palets.numpalet|N||15·"
    
    Cad = Cad & ParaGrid(text1(1), 10, "Conf.")
    Cad = Cad & "Palet|confpale.nompalet|N||35·"
    Cad = Cad & ParaGrid(text1(2), 15, "F.Inicio")
    Cad = Cad & ParaGrid(text1(3), 15, "F.Fin")
    tabla = NombreTabla & " INNER JOIN confpale ON palets.codpalet=confpale.codpalet "
    
    Titulo = "Palets"
    devuelve = "0|"
           
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vtabla = tabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vDevuelve = devuelve
        frmB.vTitulo = Titulo
        frmB.vSelElem = 0
'        frmB.vConexionGrid = cAgro  'Conexión a BD: Ariagro
        If Not EsCabecera Then frmB.Label1.FontSize = 11
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
'        If EsCabecera Then
'            PonerCadenaBusqueda
'            Text1(0).Text = Format(Text1(0).Text, "0000000")
'        End If
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If HaDevueltoDatos Then
''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
''                cmdRegresar_Click
        Else   'de ha devuelto datos, es decir NO ha devuelto datos
            PonerFoco text1(kCampo)
        End If
    End If
    Screen.MousePointer = vbDefault
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
            PonerFoco text1(kCampo)
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
        CalcularTaraEnvase Data3.Recordset.Fields(1)
    Else
        CargaGrid DataGrid1, Data2, False
        CalcularTaraEnvase "-1"
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
    text1(9).Text = Mid(text1(8).Text, 12, 8)
    text1(10).Text = Mid(text1(6).Text, 12, 8)
    text1(12).Text = Mid(text1(14).Text, 12, 8)
    text1(11).Text = Mid(text1(15).Text, 12, 8)
    
    text2(3).Text = DevuelveDesdeBDNew(cAgro, "confpale", "pesopale", "codpalet", text1(4).Text, "N")
    If text2(3).Text <> "" Then PonerFormatoDecimal text2(3), 4
    
'    FormatoDatosTotales
    
    'poner descripcion campos
    Modo = 4
    text2(4) = PonerNombreDeCod(text1(4), "confpale", "nompalet", "codpalet", "N") 'palet de confeccion
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
Dim i As Byte, Numreg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo

    'Actualiza Iconos Insertar,Modificar,Eliminar
    '## No tiene el boton modificar y no utiliza la funcion general
'    ActualizarToolbar Modo, Kmodo
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
    Else
        cmdRegresar.visible = False
    End If
        
    'Poner Flechas de desplazamiento visibles
    Numreg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then Numreg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, Numreg
          
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    'si estamos en modificar bloquea las compos que son clave primaria
    BloquearText1 Me, Modo
    BloquearCombo Me, Modo
    For i = 9 To 10
        BloquearTxt text1(i), Not (Modo = 3 Or Modo = 4 Or Modo = 1)
    Next i
    For i = 11 To 12
        BloquearTxt text1(i), Not (Modo = 3 Or Modo = 4 Or Modo = 1)
    Next i
    b = (Modo <> 1)
    'Campos Nº Factura bloqueado y en azul
    BloquearTxt text1(0), b, True
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
    For i = 8 To 10
        BloquearTxt txtAux3(i), (Modo <> 1)
    Next i
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    BloquearImgBuscar Me, Modo, ModificaLineas
    BloquearImgFec Me, 0, Modo
    BloquearImgFec Me, 1, Modo
    BloquearImgFec Me, 2, Modo
    
    Me.imgBuscar(1).Enabled = ((Modo = 3) Or (Modo = 4))
    Me.imgBuscar(1).visible = ((Modo = 3) Or (Modo = 4))
    
                    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
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
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim b As Boolean
Dim SQL As String

    On Error GoTo EDatosOK

    DatosOk = False
    
'    ComprobarDatosTotales

    'concatenamos en el text1(6) y text1(8) la fechahora
    text1(8).Text = Format(text1(2).Text, "dd/mm/yyyy") & " " & Format(text1(9).Text, "HH:MM:SS")
    If text1(3).Text <> "" And text1(10).Text <> "" Then
        text1(6).Text = Format(text1(3).Text, "dd/mm/yyyy") & " " & Format(text1(10).Text, "HH:MM:SS")
    Else
        text1(6).Text = ""
    End If
    
    If text1(13).Text <> "" And text1(12).Text <> "" Then
        text1(14).Text = Format(text1(13).Text, "dd/mm/yyyy") & " " & Format(text1(12).Text, "HH:MM:SS")
    Else
        text1(14).Text = ""
    End If
    
    If text1(13).Text <> "" And text1(11).Text <> "" Then
        text1(15).Text = Format(text1(13).Text, "dd/mm/yyyy") & " " & Format(text1(11).Text, "HH:MM:SS")
    Else
        text1(15).Text = ""
    End If
    
    'comprobamos datos OK de la tabla palets
    b = CompForm2(Me, 2, "Frame2") ' , 1) 'Comprobar formato datos ok de la cabecera: opcion=1
    If Not b Then Exit Function
    
    
    ' comprobamos los rangos de fechas
    If b And text1(3).Text <> "" Then
        If CDate(text1(2).Text) > CDate(text1(3).Text) Then
            MsgBox "La fecha de inicio no puede ser superior a la fecha fin. Revise.", vbExclamation
            b = False
            PonerFoco text1(9)
        End If
    End If
    
    If b And text1(6).Text <> "" Then
        If CDate(text1(8).Text) > CDate(text1(6).Text) Then
            MsgBox "La hora de inicio no puede ser superior a la de fin. Revise.", vbExclamation
            b = False
            PonerFoco text1(9)
        End If
    End If
    
    If b And text1(15).Text <> "" Then
        If CDate(text1(14).Text) > CDate(text1(15).Text) Then
            MsgBox "La hora de inicio de confección no puede ser superior a la de fin. Revise.", vbExclamation
            b = False
            PonerFoco text1(12)
        End If
    End If
    
    
    
    'comprobamos que el numero de pedido existe si no es nulo
    If b And text1(5).Text <> "" Then
        SQL = ""
        SQL = DevuelveDesdeBDNew(cAgro, "pedidos", "numpedid", "numpedid", text1(5), "N")
        If SQL = "" Then
            MsgBox "El número de pedido no existe en la tabla de pedidos. Reintroduzca.", vbExclamation
            text1(5).Text = ""
            b = False
            PonerFoco text1(5)
        End If
    End If
    
    
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

    NroAlbar = NroAlbaranAsignado(Data1.Recordset!numpalet, 0)
    If NroAlbar <> "" Then
        Cad = "El pedido asociado a este palet se encuentra asignado al albarán " & NroAlbar & "." & vbCrLf
        Cad = Cad & "                     ¿ Desea continuar ?"
        If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Sub
        End If
    End If

    If BloqueaRegistro(NombreTabla, "numpalet = " & Data1.Recordset!numpalet) Then
'    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
        Select Case Button.Index
            Case 1 'añadir variedad
                Set frmLPal = New frmVtasLinPalets
                
                frmLPal.ModoExt = 3
                frmLPal.Palet = Data1.Recordset.Fields(0).Value
                frmLPal.Show vbModal
            
                Set frmLPal = Nothing
            Case 2 'modificar variedad
                Set frmLPal = New frmVtasLinPalets
                
                frmLPal.ModoExt = 4
                frmLPal.Palet = Data3.Recordset.Fields(0).Value
                frmLPal.Linea = Data3.Recordset.Fields(1).Value
                frmLPal.Show vbModal
                
                Set frmLPal = Nothing
                
            Case 3 ' boton eliminar linea de variedades
                BotonEliminarLinea
            Case Else
        End Select
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
    Cad = Cad & vbCrLf & "Palet: " & Data3.Recordset.Fields(0)
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
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Variedad de Palet", Err.Description

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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Buscar
            mnBuscar_Click
        Case 2  'Todos
            BotonVerTodos
        
        Case 4  'Añadir
            mnNuevo_Click

        Case 5  'Modificar
            mnModificar_Click
        Case 6  'Borrar
            mnEliminar_Click
        Case 9  'Lineas
            mnLineas_Click
        Case 10 'Imprimir Albaran
            mnImprimir_Click
        Case 12    'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
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
Dim SQL As String

    On Error GoTo ECargaGRid

    b = DataGrid1.Enabled
    If vDataGrid.Name = "DataGrid1" Then
        Opcion = 1
    Else
        Opcion = 2
    End If
    SQL = MontaSQLCarga(enlaza, Opcion)
    CargaGridGnral vDataGrid, vData, SQL, PrimeraVez
    
    vDataGrid.RowHeight = 270
    
    CargaGrid2 vDataGrid, vData
    vDataGrid.ScrollBars = dbgAutomatic
    
     b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
     vDataGrid.Enabled = Not b
    
    Exit Sub
    
ECargaGRid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim tots As String
    
    On Error GoTo ECargaGRid

    Select Case vDataGrid.Name
        Case "DataGrid1" 'Palets_calibres
'           SQL = "SELECT numpale, numlinea, numline1, codvarie, codcalib, nomcalib, numcajas
            tots = "N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux(3)|T|Variedad|1000|;"
            tots = tots & "S|txtAux(4)|T|Calibre|1000|;S|txtAux(5)|T|Nombre Calibre|2500|;S|txtAux(6)|T|Cajas|1500|;"
            arregla tots, DataGrid1, Me
'            DataGrid1.Columns(11).Alignment = dbgCenter
'            DataGrid1.Columns(12).Alignment = dbgRight
'            DataGrid1.Columns(13).Alignment = dbgRight
'            DataGrid1.Columns(14).Alignment = dbgRight
                       
         Case "DataGrid2" 'palets_variedad
'           SQL = "SELECT numpale, numlinea, codvarie, nomvarie1, codvarco, nomvarie2, codmarca, nommarca, codforfait, nomforfait, categori, pesobrut, pesonet
            tots = "N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux3(3)|T|Variedad Real|1800|;N||||0|;"
            tots = tots & "S|txtAux3(5)|T|Variedad Comercial|1800|;N||||0|;S|txtAux3(11)|T|Marca|1800|;N||||0|;S|txtAux3(12)|T|Forfait|1800|;S|txtAux3(8)|T|Categoria|1000|;S|txtAux3(13)|T|Cajas|675|;"
            tots = tots & "S|txtAux3(9)|T|Peso Bruto|1000|;S|txtAux3(10)|T|Peso Neto|1000|;"
            arregla tots, DataGrid2, Me
            
            DataGrid2.Columns(3).Alignment = dbgLeft
            DataGrid2.Columns(5).Alignment = dbgLeft
            DataGrid2.Columns(7).Alignment = dbgLeft
            DataGrid2.Columns(9).Alignment = dbgLeft
                     
'            DataGrid2_RowColChange 1, 1
    End Select
    
    vDataGrid.HoldFields
    Exit Sub
    
ECargaGRid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
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
            If Index = 7 Then PonerFoco Me.text2(16)
            
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
            MsgBox "El Palet no tiene lineas.", vbInformation
            Exit Sub
        End If
        TituloLinea = Cad
    End If
    ModificaLineas = 0
    PonerModo 5
    PonerBotonCabecera True
End Sub


Private Function Eliminar() As Boolean
Dim SQL As String, LEtra As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
    
    On Error GoTo FinEliminar

    b = False
    If Data1.Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        

    'Eliminar en tablas de factura de Ariges
    '------------------------------------------
    SQL = " " & ObtenerWhereCP(True)

    'Lineas de calibres (palets_calibre)
    conn.Execute "Delete from palets_calibre " & SQL

    'Lineas de variedades
    conn.Execute "Delete from palets_variedad " & SQL
    
    'Cabecera de palets (palets)
    conn.Execute "Delete from " & NombreTabla & SQL
    
    'Decrementar contador si borramos el ult. palet
    Set vTipoMov = New CTiposMov
    vTipoMov.DevolverContador "PAL", Val(text1(0).Text)
    Set vTipoMov = Nothing
    
    b = True
    
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Palet", Err.Description
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
Dim SQL As String, LEtra As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
    
    On Error GoTo FinEliminar

    b = False
    If Data3.Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        

    'Eliminar en tablas de paltes_variedad y palets_calibre
    '------------------------------------------
    SQL = " where numpalet = " & Data3.Recordset.Fields(0)
    SQL = SQL & " and numlinea = " & Data3.Recordset.Fields(1)

    'Lineas de calibres (palets_calibre)
    conn.Execute "Delete from palets_calibre " & SQL

    'Lineas de variedades
    conn.Execute "Delete from palets_variedad " & SQL
    
    b = True
    
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Variedad del Palet", Err.Description
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
Dim SQL As String

    On Error Resume Next
    
    SQL = " numpalet= " & text1(0).Text  ' Data1.Recordset!numpalet ' Text1(0).Text
    If conWhere Then SQL = " WHERE " & SQL
    ObtenerWhereCP = SQL
    
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
Dim SQL As String
    
    If Opcion = 1 Then
        SQL = "SELECT numpalet, numlinea, numline1, palets_calibre.codvarie, palets_calibre.codcalib, nomcalib, numcajas "
        SQL = SQL & " FROM palets_calibre, calibres WHERE palets_calibre.codvarie = calibres.codvarie and "
        SQL = SQL & " palets_calibre.codcalib = calibres.codcalib "
    ElseIf Opcion = 2 Then
        SQL = "SELECT palets_variedad.numpalet, numlinea, palets_variedad.codvarie, a.nomvarie as nomvarie1, palets_variedad.codvarco, "
        SQL = SQL & " b.nomvarie as nomvarie2, palets_variedad.codmarca, marcas.nommarca, palets_variedad.codforfait, forfaits.nomconfe, "
        SQL = SQL & " categori, numcajas, pesobrut, pesoneto "
        SQL = SQL & " FROM palets_variedad, variedades a, variedades b, marcas, forfaits " 'lineas de variedades del palet
        SQL = SQL & " WHERE palets_variedad.codvarie = a.codvarie "
        SQL = SQL & " and palets_variedad.codvarco = b.codvarie"
        SQL = SQL & " and palets_variedad.codmarca = marcas.codmarca "
        SQL = SQL & " and palets_variedad.codforfait = forfaits.codforfait "
    End If
    
    If enlaza Then
        SQL = SQL & " and " & ObtenerWhereCP(False)
        If Opcion = 1 Then SQL = SQL & " AND numlinea=" & Data3.Recordset.Fields!NumLinea
    Else
        SQL = SQL & " and numpalet = -1"
    End If
    SQL = SQL & " ORDER BY numpalet"
    If Opcion = 1 Then SQL = SQL & ", numlinea "
    MontaSQLCarga = SQL
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean, bAux As Boolean
Dim i As Integer

        b = (Modo = 2) Or (Modo = 0) 'Or (Modo = 5 And ModificaLineas = 0)
        'Buscar
        Toolbar1.Buttons(1).Enabled = b
        Me.mnBuscar.Enabled = b
        'Vore Tots
        Toolbar1.Buttons(2).Enabled = b
        Me.mnVerTodos.Enabled = b
        'Añadir
        Toolbar1.Buttons(4).Enabled = b
        Me.mnModificar.Enabled = b
        
        b = (Modo = 2 And Data1.Recordset.RecordCount > 0)
        'Modificar
        Toolbar1.Buttons(5).Enabled = b
        Me.mnModificar.Enabled = b
        'eliminar
        Toolbar1.Buttons(6).Enabled = (Modo = 2)
        Me.mnEliminar.Enabled = (Modo = 2)
            
        b = (Modo = 2)
        'Imprimir
        Toolbar1.Buttons(10).Enabled = b
        Me.mnImprimir.Enabled = b
        

    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
'++monica: si insertamos lo he quitado
'    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
    b = (Modo = 4 Or Modo = 2)
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.Data3.Recordset.RecordCount > 0)
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i


End Sub


Private Sub BotonImprimir()
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadselect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim SQL As String

    If text1(0).Text = "" Then
        MsgBox "Debe seleccionar un Palet para Imprimir.", vbInformation
        Exit Sub
    End If
    
    cadFormula = ""
    cadParam = ""
    cadselect = ""
    numParam = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    indRPT = 5 'Impresion de Palet
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
      
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de palet
    '---------------------------------------------------
    If text1(0).Text <> "" Then
        'Nº palet
        devuelve = "{" & NombreTabla & ".numpalet}=" & Val(text1(0).Text)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        devuelve = "numpalet = " & Val(text1(0).Text)
        If Not AnyadirAFormula(cadselect, devuelve) Then Exit Sub
    End If
    
    cadParam = cadParam & "|pImprimeBarras=""1""|"
    numParam = numParam + 1
    
    SQL = ""
    SQL = ClientePalet(text1(0).Text)
    
    cadParam = cadParam & "|pCliente=""" & Trim(SQL) & """|"
    numParam = numParam + 1
   
    If Not HayRegParaInforme(NombreTabla, cadselect) Then Exit Sub
     
     With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .ConSubInforme = True
            .Opcion = 0
            .Titulo = "Impresión de Palet"
            .Show vbModal
    End With
End Sub


Private Sub BotonImprimirTicket()
Dim MIPATH As String
Dim cadImpresion As String, SQL As String
Dim NomImpre As String
Dim NomImpTi As String
Dim bImpre As Boolean

    cadImpresion = "{scafac.codtipom}='" & text1(1).Text & "' and {scafac.numfactu}=" & text1(0).Text
    SQL = cadImpresion & " and {scafac.fecfactu}=" & DBSet(text1(2).Text, "F")
    cadImpresion = cadImpresion & " and {scafac.fecfactu}=Date(" & Year(CDate(text1(2).Text)) & "," & Month(CDate(text1(2).Text)) & "," & Day(CDate(text1(2).Text)) & ")"
    
    If Not HayRegParaInforme("scafac", SQL) Then Exit Sub
    
'    'Obtener que terminal es
'     'Terminal con el que trabajaremos, leemos el nombre del ordenador
'    SQL = ComputerName 'Nombre PC conectado por Terminal Server / local
'    SQL = DevuelveDesdeBDNew(conAri, "spatpvt", "numtermi", "nombrepc", SQL, "T")
'    If Not IsNumeric(SQL) Then
'        MsgBox "No se ha podido establecer la impresora de ticket." & vbCrLf & "Debe configurar primero los parámetros del TPV.", vbExclamation
'    Else
'        bImpre = True
'    End If
'
'    If bImpre Then
'         'Establecemos la impresora de ticket
'         NomImpTi = NombreImpresoraTicket(CInt(SQL))
'         If NomImpTi <> "" Then
'            If Printer.DeviceName <> NomImpTi Then
'                'guardamos la impresora que habia
'                NomImpre = Printer.DeviceName
'                'establecemos la de ticket
'                EstablecerImpresora NomImpTi
'            End If
'        End If
'    End If


    


    MIPATH = App.path & "\Informes\"
'    cadImpresion = cadImpresion & " and {scafac.fecfactu}=Date(" & Year(RSVenta!fecventa) & "," & Month(RSVenta!fecventa) & "," & Day(RSVenta!fecventa) & ")"
    With frmVisReport
        .FormulaSeleccion = cadImpresion
        .SoloImprimir = False
        .OtrosParametros = ""
        .NumeroParametros = 0
        .MostrarTree = False
        .Informe = MIPATH & "rTPVTicket.rpt"
        .ConSubInforme = False
        .Opcion = 93
        .ExportarPDF = False
        .Show vbModal
   End With
   
'   If bImpre Then
'        'volver la impresora a la predeterminada
'        EstablecerImpresora NomImpre
'   End If
   
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
Dim SQL As String
Dim i As Byte
    
    Combo1(0).Clear
    
    Combo1(0).AddItem "Cooperativa"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    
    Combo1(0).AddItem "Terceros"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    
    Combo1(0).AddItem "Mezclado"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    
    Combo1(0).AddItem "Otros"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 3

End Sub

Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim SQL As String

    On Error GoTo EInsertarCab
    
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
        SQL = CadenaInsertarDesdeForm(Me)
        If SQL <> "" Then
            If InsertarOferta(SQL, vTipoMov) Then
                CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                PonerCadenaBusqueda
                PonerModo 2
                'Ponerse en Modo Insertar Lineas
'                BotonMtoLineas 0, "Variedades"
'                BotonAnyadirLinea
                Set frmLPal = New frmVtasLinPalets
                
                frmLPal.ModoExt = 3
                frmLPal.Palet = text1(0).Text
                frmLPal.Show vbModal
                
                Set frmLPal = Nothing
            End If
        End If
        text1(0).Text = Format(text1(0).Text, "0000000")
        CalcularTaraEnvase 1
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
        devuelve = DevuelveDesdeBDNew(cAgro, NombreTabla, "numpalet", "numpalet", text1(0).Text, "N")
        If devuelve <> "" Then
            'Ya existe el contador incrementarlo
            Existe = True
            vTipoMov.IncrementarContador (CodTipoMov)
            text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
            cambiaSQL = True
        Else
            Existe = False
        End If
    Loop Until Not Existe
    If cambiaSQL Then vSQL = CadenaInsertarDesdeForm(Me)
    
    
    'Aqui empieza transaccion
    conn.BeginTrans
    MenError = "Error al insertar en la tabla Cabecera de Palets (" & NombreTabla & ")."
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
        
        MenError = "Error al actualizar el contador del Palets."
    '    bol = vTipoMov.IncrementarContador("REG")
        vTipoMov.IncrementarContador (CodTipoMov)
'    End If
    
EInsertarOferta:
        If Err.Number <> 0 Then
            MenError = "Insertando Albaran." & vbCrLf & "----------------------------" & vbCrLf & MenError
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

Private Sub CalcularTaraEnvase(NumLinea As String)
Dim Valor As Currency
Dim TotalCajas As Currency
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim TaraEnvase As String
Dim Forfaits As String
Dim PesoCaja As String

    If CCur(NumLinea) < 1 Then
        text2(1).Text = ""
        text2(2).Text = ""
        Exit Sub
    End If

'    'total importes de envases para ese forfait
'    Sql = "select sum(numcajas) "
'    Sql = Sql & " from palets_calibre where numpalet = " & Data1.Recordset.Fields(0)
'    Sql = Sql & " and palets_calibre.numlinea = " & numlinea
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    TotalCajas = 0
'    If Not Rs.EOF Then
'        If Rs.Fields(0).Value > 0 Then TotalCajas = Rs.Fields(0).Value
'    End If
'
'    Rs.Close
'    Set Rs = Nothing
    
    SQL = ""
    SQL = DevuelveDesdeBDNew(cAgro, "palets_variedad", "numcajas", "numpalet", Data1.Recordset.Fields(0), "N", , "numlinea", NumLinea, "N")
    If SQL = "" Then
        TotalCajas = 0
    Else
        TotalCajas = CLng(SQL)
    End If
    
    Forfaits = DevuelveDesdeBDNew(cAgro, "palets_variedad", "codforfait", "numpalet", Data1.Recordset.Fields(0), "N", , "numlinea", NumLinea, "N")
    
    SQL = ""
    SQL = DevuelveDesdeBDNew(cAgro, "forfaits", "pesocaja", "codforfait", Forfaits, "N")
    PesoCaja = ""
    If SQL <> "" Then
        PesoCaja = Format(TransformaPuntosComas(SQL), "###,###,##0.00")
    End If
        
    If PesoCaja <> "" Then
       text2(0).Text = PesoCaja
    Else
       text2(0).Text = ""
       PesoCaja = "0"
    End If
    
    text2(1).Text = Format(TotalCajas, "###,###,##0")
    Valor = Round(CCur(TransformaPuntosComas(PesoCaja)) * TotalCajas, 2)
    If Valor <> 0 Then
        text2(2).Text = Format(Valor, "###,###,##0.00")
    Else
        text2(2).Text = ""
    End If


    'Calculo de totales
    SQL = "select palets_variedad.numlinea, round(sum(palets_calibre.numcajas) * forfaits.pesocaja  ,2) "
    SQL = SQL & " from palets_variedad, forfaits, palets_calibre "
    SQL = SQL & " where palets_variedad.numpalet = " & Data1.Recordset.Fields(0) & " and "
    SQL = SQL & " palets_variedad.numpalet = palets_calibre.numpalet and "
    SQL = SQL & " palets_variedad.numlinea = palets_calibre.numlinea and "
    SQL = SQL & " palets_variedad.codforfait = forfaits.codforfait "
    SQL = SQL & " group by 1"
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    TaraEnvase = 0
    While Not Rs.EOF
        TaraEnvase = TaraEnvase + DBLet(Rs.Fields(1).Value, "N")
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    text2(5).Text = Format(TaraEnvase, "###,###,##0.00")
    
    text2(6).Text = Format(TaraEnvase + CCur(TransformaPuntosComas(ComprobarCero(text2(3).Text))), "###,###,##0.00")


End Sub


Private Function ClientePalet(Palet As String) As String
Dim Rs As ADODB.Recordset
Dim SQL As String

    On Error GoTo eClientePalet

    ClientePalet = ""
    SQL = "select pedidos.codclien, clientes.nomclien from palets, pedidos, clientes "
    SQL = SQL & " where palets.numpalet = " & DBSet(Palet, "N")
    SQL = SQL & " and palets.numpedid = pedidos.numpedid "
    SQL = SQL & " and pedidos.codclien = clientes.codclien "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    If Not Rs.EOF Then
        ClientePalet = "Cliente : " & Format(DBLet(Rs.Fields(0).Value, "N"), "000000") & " " & DBLet(Rs.Fields(1).Value, "T")
    End If
    
    Set Rs = Nothing
    Exit Function
    
eClientePalet:
    MuestraError Err.Number, "Cliente del pedido asociado"
End Function
