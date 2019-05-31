VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmComEntPedidos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedidos Proveedor"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   12825
   Icon            =   "frmComEntPedidos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   12825
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
      Left            =   10170
      TabIndex        =   136
      Top             =   330
      Width           =   1605
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   135
      TabIndex        =   134
      Top             =   90
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   135
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
      TabIndex        =   132
      Top             =   90
      Width           =   750
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   210
         TabIndex        =   133
         Top             =   180
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Generar Albarán"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   4620
      TabIndex        =   130
      Top             =   90
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   131
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
      Left            =   1485
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   127
      Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
      Top             =   6930
      Width           =   6900
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
      Left            =   9495
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   126
      Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
      Top             =   6930
      Width           =   3085
   End
   Begin VB.Frame Frame2 
      Height          =   870
      Left            =   120
      TabIndex        =   76
      Top             =   855
      Width           =   12505
      Begin VB.TextBox Text3 
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
         Index           =   50
         Left            =   10680
         MaxLength       =   15
         TabIndex        =   138
         Text            =   "Text1 7"
         Top             =   360
         Width           =   1530
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Height          =   285
         Index           =   0
         Left            =   10665
         MaxLength       =   15
         TabIndex        =   139
         Text            =   "TOTAL PEDIDO"
         Top             =   135
         Width           =   1515
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
         Index           =   5
         Left            =   7140
         MaxLength       =   40
         TabIndex        =   4
         Tag             =   "Nombre Proveedor|T|N|||scappr|nomprove||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   360
         Width           =   3405
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
         Left            =   6210
         MaxLength       =   30
         TabIndex        =   3
         Tag             =   "Cod. Proveedor|N|N|0|999999|scappr|codprove|000000|N|"
         Text            =   "Text1"
         Top             =   360
         Width           =   870
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
         Left            =   1380
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Pedido|F|N|||scappr|fecpedpr|dd/mm/yyyy|N|"
         Top             =   375
         Width           =   1305
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
         Left            =   240
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Nº Pedido|N|S|0||scappr|numpedpr|0000000|S|"
         Text            =   "Text1 7"
         Top             =   375
         Width           =   1080
      End
      Begin VB.CheckBox chkRestoPed 
         Caption         =   "Resto de Pedido"
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
         Height          =   255
         Left            =   2820
         TabIndex        =   2
         Tag             =   "Resto de Pedido|N|N|||scappr|restoped||N|"
         Top             =   375
         Width           =   1935
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   5925
         ToolTipText     =   "Buscar proveedor"
         Top             =   390
         Width           =   240
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
         Left            =   4860
         TabIndex        =   79
         Top             =   390
         Width           =   1050
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   14
         Left            =   1380
         TabIndex        =   78
         Top             =   135
         Width           =   990
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   2430
         Picture         =   "frmComEntPedidos.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   135
         Width           =   240
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
         Index           =   50
         Left            =   240
         TabIndex        =   77
         Top             =   135
         Width           =   1050
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   25
      Left            =   2025
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   123
      Text            =   "Text2"
      Top             =   975
      Width           =   3285
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   25
      Left            =   3270
      MaxLength       =   30
      TabIndex        =   122
      Text            =   "Text1"
      Top             =   990
      Width           =   780
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   23
      Left            =   1635
      MaxLength       =   30
      TabIndex        =   121
      Text            =   "Text1"
      Top             =   900
      Width           =   660
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   1
      Left            =   2340
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   120
      Text            =   "Text2"
      Top             =   915
      Width           =   3290
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      Left            =   2445
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   119
      Text            =   "Text2"
      Top             =   945
      Visible         =   0   'False
      Width           =   3405
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      Left            =   1605
      MaxLength       =   30
      TabIndex        =   118
      Text            =   "Text1"
      Top             =   945
      Visible         =   0   'False
      Width           =   780
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
      Index           =   16
      Left            =   1485
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   40
      Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
      Top             =   6570
      Width           =   6930
   End
   Begin VB.Frame Frame1 
      Height          =   525
      Index           =   0
      Left            =   135
      TabIndex        =   28
      Top             =   7335
      Width           =   2220
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
         TabIndex        =   29
         Top             =   180
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
      Left            =   11565
      TabIndex        =   17
      Top             =   7410
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
      Left            =   10395
      TabIndex        =   16
      Top             =   7410
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   4050
      Top             =   1125
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
      Left            =   5715
      Top             =   1305
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   120
      TabIndex        =   30
      Top             =   1800
      Width           =   12520
      _ExtentX        =   22093
      _ExtentY        =   8281
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos básicos"
      TabPicture(0)   =   "frmComEntPedidos.frx":0097
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FrameCliente"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdAux(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdAux(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtAux(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtAux(7)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtAux(6)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtAux(5)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtAux(4)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtAux(3)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtAux(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtAux(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "DataGrid1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "ToolAux(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Otros Datos"
      TabPicture(1)   =   "frmComEntPedidos.frx":00B3
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(45)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Text1(17)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Text1(18)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Text1(19)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Text1(20)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Text1(21)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "FrameDirMercancia"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "FrameDirFactura"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "FrameHco"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Totales"
      TabPicture(2)   =   "frmComEntPedidos.frx":00CF
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameFactura"
      Tab(2).ControlCount=   1
      Begin VB.Frame FrameHco 
         Caption         =   "Datos  Eliminación"
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
         Height          =   1740
         Left            =   8520
         TabIndex        =   114
         Top             =   2580
         Width           =   3765
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
            Index           =   24
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   25
            Top             =   420
            Width           =   1350
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
            Index           =   26
            Left            =   135
            MaxLength       =   30
            TabIndex        =   26
            Text            =   "Text1"
            Top             =   1245
            Width           =   540
         End
         Begin VB.TextBox Text2 
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
            Index           =   26
            Left            =   675
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   115
            Text            =   "Text2"
            Top             =   1245
            Width           =   2955
         End
         Begin VB.Label Label1 
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
            Height          =   255
            Index           =   37
            Left            =   120
            TabIndex        =   117
            Top             =   435
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Incidencia"
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
            Index           =   40
            Left            =   120
            TabIndex        =   116
            Top             =   960
            Width           =   1050
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   10
            Left            =   1170
            ToolTipText     =   "Buscar incidencia"
            Top             =   960
            Width           =   240
         End
      End
      Begin VB.Frame FrameFactura 
         Height          =   3300
         Left            =   -74640
         TabIndex        =   82
         Top             =   720
         Width           =   11835
         Begin VB.TextBox Text3 
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
            Index           =   33
            Left            =   240
            MaxLength       =   15
            TabIndex        =   99
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1485
         End
         Begin VB.TextBox Text3 
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
            Index           =   34
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   98
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1365
         End
         Begin VB.TextBox Text3 
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
            Index           =   35
            Left            =   3960
            MaxLength       =   15
            TabIndex        =   97
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1365
         End
         Begin VB.TextBox Text3 
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
            Index           =   36
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   96
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1485
         End
         Begin VB.TextBox Text3 
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
            Index           =   43
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   95
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1485
         End
         Begin VB.TextBox Text3 
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
            Index           =   37
            Left            =   3960
            MaxLength       =   4
            TabIndex        =   94
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   660
         End
         Begin VB.TextBox Text3 
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
            Index           =   40
            Left            =   5017
            MaxLength       =   5
            TabIndex        =   93
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   705
         End
         Begin VB.TextBox Text3 
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
            Index           =   46
            Left            =   7560
            MaxLength       =   15
            TabIndex        =   92
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1485
         End
         Begin VB.TextBox Text3 
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
            Index           =   44
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   91
            Text            =   "Text1 7"
            Top             =   1800
            Width           =   1485
         End
         Begin VB.TextBox Text3 
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
            Index           =   38
            Left            =   3960
            MaxLength       =   4
            TabIndex        =   90
            Text            =   "Text1 7"
            Top             =   1800
            Width           =   660
         End
         Begin VB.TextBox Text3 
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
            Index           =   41
            Left            =   5017
            MaxLength       =   5
            TabIndex        =   89
            Text            =   "Text1 7"
            Top             =   1800
            Width           =   705
         End
         Begin VB.TextBox Text3 
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
            Index           =   47
            Left            =   7560
            MaxLength       =   15
            TabIndex        =   88
            Text            =   "Text1 7"
            Top             =   1800
            Width           =   1485
         End
         Begin VB.TextBox Text3 
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
            Index           =   45
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   87
            Text            =   "Text1 7"
            Top             =   2175
            Width           =   1485
         End
         Begin VB.TextBox Text3 
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
            Index           =   39
            Left            =   3960
            MaxLength       =   4
            TabIndex        =   86
            Text            =   "Text1 7"
            Top             =   2175
            Width           =   660
         End
         Begin VB.TextBox Text3 
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
            Index           =   42
            Left            =   5017
            MaxLength       =   5
            TabIndex        =   85
            Text            =   "Text1 7"
            Top             =   2175
            Width           =   705
         End
         Begin VB.TextBox Text3 
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
            Index           =   48
            Left            =   7560
            MaxLength       =   15
            TabIndex        =   84
            Text            =   "Text1 7"
            Top             =   2175
            Width           =   1485
         End
         Begin VB.TextBox Text3 
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
            Index           =   49
            Left            =   7380
            MaxLength       =   15
            TabIndex        =   83
            Text            =   "Text1 7"
            Top             =   2730
            Width           =   1665
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
            Height          =   360
            Index           =   27
            Left            =   5760
            TabIndex        =   113
            Top             =   1185
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Importe Bruto"
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
            Left            =   240
            TabIndex        =   112
            Top             =   270
            Width           =   1620
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto PP"
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
            Left            =   2160
            TabIndex        =   111
            Top             =   270
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto Gn"
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
            Left            =   3960
            TabIndex        =   110
            Top             =   270
            Width           =   1575
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
            Height          =   360
            Index           =   18
            Left            =   5760
            TabIndex        =   109
            Top             =   270
            Width           =   1935
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
            Index           =   30
            Left            =   1920
            TabIndex        =   108
            Top             =   480
            Width           =   135
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
            Index           =   31
            Left            =   3720
            TabIndex        =   107
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "="
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
            Index           =   32
            Left            =   5520
            TabIndex        =   106
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Importe  IVA"
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
            Left            =   7560
            TabIndex        =   105
            Top             =   1185
            Width           =   1605
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
            Index           =   8
            Left            =   7320
            TabIndex        =   104
            Top             =   1320
            Width           =   135
         End
         Begin VB.Line Line1 
            X1              =   4005
            X2              =   7320
            Y1              =   1080
            Y2              =   1080
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
            TabIndex        =   103
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "TOTAL PEDIDO"
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
            Height          =   285
            Index           =   39
            Left            =   5775
            TabIndex        =   102
            Top             =   2745
            Width           =   1530
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
            Height          =   360
            Index           =   41
            Left            =   4995
            TabIndex        =   101
            Top             =   1185
            Width           =   720
         End
         Begin VB.Label Label1 
            Caption         =   "Cod. IVA"
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
            Index           =   42
            Left            =   3960
            TabIndex        =   100
            Top             =   1185
            Width           =   915
         End
      End
      Begin VB.Frame FrameDirFactura 
         Caption         =   "Dirección Factura"
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
         Height          =   1900
         Left            =   6435
         TabIndex        =   66
         Top             =   405
         Width           =   5895
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
            Left            =   1185
            MaxLength       =   30
            TabIndex        =   19
            Tag             =   "Direc. Factura|N|S|0|999|scappr|coddiref|000|N|"
            Text            =   "Text1"
            Top             =   360
            Width           =   540
         End
         Begin VB.TextBox Text2 
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
            Left            =   1740
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   75
            Text            =   "Text2"
            Top             =   360
            Width           =   4050
         End
         Begin VB.TextBox Text2 
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
            Index           =   24
            Left            =   1185
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   70
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   1425
            Width           =   3105
         End
         Begin VB.TextBox Text2 
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
            Index           =   22
            Left            =   1185
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   69
            Text            =   "Text15"
            Top             =   1065
            Width           =   855
         End
         Begin VB.TextBox Text2 
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
            Index           =   23
            Left            =   2055
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   68
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   1065
            Width           =   3720
         End
         Begin VB.TextBox Text2 
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
            Index           =   21
            Left            =   1185
            Locked          =   -1  'True
            MaxLength       =   35
            TabIndex        =   67
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   720
            Width           =   4590
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   5
            Left            =   900
            ToolTipText     =   "Buscar dirección"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
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
            Left            =   120
            TabIndex        =   74
            Top             =   1425
            Width           =   1005
         End
         Begin VB.Label Label1 
            Caption         =   "Población"
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
            Left            =   120
            TabIndex        =   73
            Top             =   1080
            Width           =   1050
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
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
            Left            =   120
            TabIndex        =   72
            Top             =   720
            Width           =   960
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
            Index           =   10
            Left            =   120
            TabIndex        =   71
            Top             =   360
            Width           =   750
         End
      End
      Begin VB.Frame FrameDirMercancia 
         Caption         =   "Dirección Mercancia"
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
         Height          =   1900
         Left            =   240
         TabIndex        =   56
         Top             =   420
         Width           =   6030
         Begin VB.TextBox Text2 
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
            Index           =   17
            Left            =   1230
            Locked          =   -1  'True
            MaxLength       =   35
            TabIndex        =   61
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   720
            Width           =   4725
         End
         Begin VB.TextBox Text2 
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
            Left            =   2190
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   60
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   1065
            Width           =   3765
         End
         Begin VB.TextBox Text2 
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
            Left            =   1230
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   59
            Text            =   "Text15"
            Top             =   1065
            Width           =   945
         End
         Begin VB.TextBox Text2 
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
            Index           =   20
            Left            =   1230
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   58
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   1425
            Width           =   3240
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
            Left            =   1230
            MaxLength       =   30
            TabIndex        =   18
            Tag             =   "Direc. Mercancia|N|S|0|999|scappr|coddirea|000|N|"
            Text            =   "Text1"
            Top             =   360
            Width           =   540
         End
         Begin VB.TextBox Text2 
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
            Left            =   1785
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   57
            Text            =   "Text2"
            Top             =   360
            Width           =   4185
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
            Index           =   9
            Left            =   120
            TabIndex        =   65
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
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
            Left            =   120
            TabIndex        =   64
            Top             =   720
            Width           =   1050
         End
         Begin VB.Label Label1 
            Caption         =   "Población"
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
            Left            =   120
            TabIndex        =   63
            Top             =   1080
            Width           =   1050
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
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
            Left            =   120
            TabIndex        =   62
            Top             =   1425
            Width           =   1050
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   4
            Left            =   945
            ToolTipText     =   "Buscar dirección"
            Top             =   360
            Width           =   240
         End
      End
      Begin VB.Frame FrameCliente 
         Height          =   1770
         Left            =   -74775
         TabIndex        =   45
         Top             =   315
         Width           =   12155
         Begin VB.TextBox Text2 
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
            Height          =   315
            Index           =   0
            Left            =   8580
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   80
            Text            =   "Text2"
            Top             =   240
            Width           =   3290
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
            Height          =   315
            Index           =   22
            Left            =   7875
            MaxLength       =   30
            TabIndex        =   11
            Tag             =   "Cliente|N|S|0|999999|scappr|codclien|000000|N|"
            Text            =   "Text1"
            Top             =   240
            Width           =   660
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
            Index           =   16
            Left            =   7590
            MaxLength       =   25
            TabIndex        =   13
            Tag             =   "Tipo Portes|T|S|||scappr|tipoporte||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwww"
            Top             =   975
            Width           =   2625
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
            Index           =   11
            Left            =   1395
            MaxLength       =   30
            TabIndex        =   10
            Tag             =   "Provincia|T|N|||scappr|proprove||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   1350
            Width           =   2790
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
            Index           =   9
            Left            =   1395
            MaxLength       =   6
            TabIndex        =   8
            Tag             =   "CPostal|T|N|||scappr|codpobla||N|"
            Text            =   "Text15"
            Top             =   960
            Width           =   630
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
            Index           =   10
            Left            =   2040
            MaxLength       =   30
            TabIndex        =   9
            Tag             =   "Población|T|N|||scappr|pobprove||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   960
            Width           =   4245
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
            Index           =   7
            Left            =   3735
            MaxLength       =   20
            TabIndex        =   6
            Tag             =   "teléfono Proveedor|T|S|||scappr|telprove||N|"
            Text            =   "12345678911234567899"
            Top             =   225
            Width           =   2550
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
            Index           =   6
            Left            =   1395
            MaxLength       =   15
            TabIndex        =   5
            Tag             =   "NIF Proveedor|T|N|||scappr|nifprove||N|"
            Text            =   "123456789"
            Top             =   240
            Width           =   1275
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
            Height          =   315
            Index           =   12
            Left            =   7875
            MaxLength       =   30
            TabIndex        =   12
            Tag             =   "Forma de Pago|N|N|0|999|scappr|codforpa|000|N|"
            Text            =   "Text1"
            Top             =   600
            Width           =   660
         End
         Begin VB.TextBox Text2 
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
            Height          =   315
            Index           =   12
            Left            =   8580
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   47
            Text            =   "Text2"
            Top             =   600
            Width           =   3290
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
            Left            =   7605
            MaxLength       =   7
            TabIndex        =   14
            Tag             =   "Descuento P.Pago|N|N|0|99.90|scaped|dtoppago|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1350
            Width           =   750
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
            Left            =   9495
            MaxLength       =   7
            TabIndex        =   15
            Tag             =   "Descuento General|N|N|0|99.90|scaped|dtognral|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1350
            Width           =   750
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
            Index           =   8
            Left            =   1395
            MaxLength       =   35
            TabIndex        =   7
            Tag             =   "Domicilio|T|N|||scappr|domprove||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   600
            Width           =   4890
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   7
            Left            =   7590
            ToolTipText     =   "Buscar cliente"
            Top             =   255
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Para Cliente"
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
            Left            =   6330
            TabIndex        =   81
            Top             =   240
            Width           =   1215
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   6
            Left            =   1125
            ToolTipText     =   "Buscar proveedor varios"
            Top             =   255
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Portes"
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
            Left            =   6315
            TabIndex        =   55
            Top             =   975
            Width           =   1125
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   1
            Left            =   1125
            ToolTipText     =   "Buscar población"
            Top             =   990
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
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
            Left            =   120
            TabIndex        =   54
            Top             =   1350
            Width           =   915
         End
         Begin VB.Label Label1 
            Caption         =   "Población"
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
            Left            =   120
            TabIndex        =   53
            Top             =   960
            Width           =   1005
         End
         Begin VB.Label Label1 
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
            Height          =   255
            Index           =   19
            Left            =   2790
            TabIndex        =   52
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label1 
            Caption         =   "NIF"
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
            Index           =   20
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Forma Pago"
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
            Left            =   6330
            TabIndex        =   50
            Top             =   600
            Width           =   1170
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. P.P"
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
            Index           =   25
            Left            =   6345
            TabIndex        =   49
            Top             =   1395
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. Gral"
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
            Index           =   26
            Left            =   8550
            TabIndex        =   48
            Top             =   1395
            Width           =   960
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   7590
            ToolTipText     =   "Buscar forma de pago"
            Top             =   615
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
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
            Left            =   120
            TabIndex        =   46
            Top             =   600
            Width           =   1005
         End
      End
      Begin VB.CommandButton cmdAux 
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
         Left            =   -72360
         TabIndex        =   44
         ToolTipText     =   "Buscar artículo"
         Top             =   3960
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton cmdAux 
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
         Left            =   -74040
         TabIndex        =   43
         ToolTipText     =   "Buscar almacen"
         Top             =   3960
         Visible         =   0   'False
         Width           =   195
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
         Height          =   330
         Index           =   2
         Left            =   -72120
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   35
         Tag             =   "Nombre Artículo"
         Text            =   "nomArtic"
         Top             =   3960
         Visible         =   0   'False
         Width           =   3285
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
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
         Index           =   7
         Left            =   -65400
         MaxLength       =   12
         TabIndex        =   41
         Tag             =   "Importe"
         Text            =   "Importe"
         Top             =   3960
         Visible         =   0   'False
         Width           =   1095
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
         Height          =   330
         Index           =   6
         Left            =   -66000
         MaxLength       =   30
         TabIndex        =   39
         Tag             =   "Descuento 2"
         Text            =   "Dto2"
         Top             =   3960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
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
         Index           =   5
         Left            =   -66600
         MaxLength       =   5
         TabIndex        =   38
         Tag             =   "Descuento 1"
         Text            =   "Dto1"
         Top             =   3960
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
         Height          =   330
         Index           =   4
         Left            =   -67560
         MaxLength       =   12
         TabIndex        =   37
         Tag             =   "Precio"
         Text            =   "123,456.7879"
         Top             =   3960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
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
         Left            =   -68760
         MaxLength       =   16
         TabIndex        =   36
         Tag             =   "Cantidad"
         Text            =   "1,234,567,891.25"
         Top             =   3960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
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
         Left            =   -73800
         MaxLength       =   18
         TabIndex        =   34
         Tag             =   "Código Artículo"
         Text            =   "Artic Artic Artic5"
         Top             =   3900
         Visible         =   0   'False
         Width           =   1455
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
         Height          =   330
         Index           =   0
         Left            =   -74640
         MaxLength       =   15
         TabIndex        =   33
         Tag             =   "Código Almacen"
         Text            =   "codalmac"
         Top             =   3900
         Visible         =   0   'False
         Width           =   615
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
         Index           =   21
         Left            =   225
         MaxLength       =   80
         TabIndex        =   24
         Tag             =   "Observación 5|T|S|||scappr|observa5||N|"
         Top             =   4005
         Width           =   8115
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
         Index           =   20
         Left            =   225
         MaxLength       =   80
         TabIndex        =   23
         Tag             =   "Observación 4|T|S|||scappr|observa4||N|"
         Top             =   3645
         Width           =   8115
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
         Index           =   19
         Left            =   225
         MaxLength       =   80
         TabIndex        =   22
         Tag             =   "Observación 3|T|S|||scappr|observa3||N|"
         Top             =   3285
         Width           =   8115
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
         Index           =   18
         Left            =   225
         MaxLength       =   80
         TabIndex        =   21
         Tag             =   "Observación 2|T|S|||scappr|observa2||N|"
         Top             =   2925
         Width           =   8115
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
         Index           =   17
         Left            =   225
         MaxLength       =   80
         TabIndex        =   20
         Tag             =   "Observación 1|T|S|||scappr|observa1||N|"
         Top             =   2565
         Width           =   8115
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmComEntPedidos.frx":00EB
         Height          =   2025
         Left            =   -74760
         TabIndex        =   42
         Top             =   2520
         Width           =   12160
         _ExtentX        =   21458
         _ExtentY        =   3572
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   0
         Left            =   -74775
         TabIndex        =   125
         Top             =   2115
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
      Begin VB.Label Label1 
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
         Index           =   45
         Left            =   270
         TabIndex        =   32
         Top             =   2295
         Width           =   1500
      End
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
      Left            =   11565
      TabIndex        =   27
      Top             =   7425
      Visible         =   0   'False
      Width           =   1065
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   12210
      TabIndex        =   137
      Top             =   270
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
      Height          =   255
      Index           =   21
      Left            =   405
      TabIndex        =   129
      Top             =   6990
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Unidad"
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
      Left            =   8550
      TabIndex        =   128
      Top             =   6990
      Width           =   885
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   9
      Left            =   2910
      Picture         =   "frmComEntPedidos.frx":0100
      ToolTipText     =   "Buscar trabajador"
      Top             =   960
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Trabajador"
      Height          =   255
      Index           =   38
      Left            =   2025
      TabIndex        =   124
      Top             =   975
      Width           =   825
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   8
      Left            =   1350
      Picture         =   "frmComEntPedidos.frx":0202
      ToolTipText     =   "Buscar trabajador"
      Top             =   915
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   2
      Left            =   1320
      Picture         =   "frmComEntPedidos.frx":0304
      ToolTipText     =   "Buscar trabajador"
      Top             =   945
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Ampliación Línea"
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
      Index           =   35
      Left            =   405
      TabIndex        =   31
      Top             =   6615
      Width           =   1335
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
      Begin VB.Menu mnLineas 
         Caption         =   "&Lineas"
         Enabled         =   0   'False
         HelpContextID   =   2
         Shortcut        =   ^L
         Visible         =   0   'False
      End
      Begin VB.Menu mnGenAlbaran 
         Caption         =   "&Generar Albaran"
         HelpContextID   =   2
         Shortcut        =   ^G
      End
      Begin VB.Menu mnImpPedido 
         Caption         =   "&Imprimir Pedido"
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
Attribute VB_Name = "frmComEntPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public EsHistorico As Boolean 'Si es true abrir el formulario con la tabla de
                              'de historico schppr, y solo en modo de consulta

Public Pedido As String ' si venimos de la ficha de proveedores consultando documentos

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'Form Calendario Fecha
Attribute frmC.VB_VarHelpID = -1

Private WithEvents frmProv As frmManProve  'Form Mto Proveedores
Attribute frmProv.VB_VarHelpID = -1
Private WithEvents frmProveV As frmComProveV  'Form Mto Proveedores Varios
Attribute frmProveV.VB_VarHelpID = -1
Private WithEvents frmDir As frmComDirecciones
Attribute frmDir.VB_VarHelpID = -1
Private WithEvents frmFP As frmManFpago 'Form Mto Formas de Pago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmAlm As frmManAlmProp   'Form Almacenes Propios
Attribute frmAlm.VB_VarHelpID = -1
Private WithEvents frmArt As frmManArtic   'Form Articulos
Attribute frmArt.VB_VarHelpID = -1
Private WithEvents frmCli As frmClientes 'form mantenimiento clientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmInc As frmManInciden  'form mantenimiento incidencias eliminacion
Attribute frmInc.VB_VarHelpID = -1

'Private WithEvents frmNSerie As frmRepCargarNSerie  'Form Cargar nº Series
'Private WithEvents frmNLote As frmAlmCargarNLote   'Form Cargar nº lote
Private WithEvents frmList As frmListadoOfer 'Listados
Attribute frmList.VB_VarHelpID = -1
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
'   6.- Cargar cantidad servidas al Generar Albaran no completo (Pedido --> Albaran)
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------


Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean


'Codigo tipo de movimiento en función del valor en la tabla de parámetros: stipom
Dim CodTipoMov As String


Dim EsDeVarios As Boolean 'Si el Proveedor mostrado es de Varios o No

'SQL de la tabla principal del formulario
Private CadenaConsulta As String
Private CadenaSQL As String 'Para crear consulta de Generar Albaran a partir del Pedido

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla de Cabecera
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

'Variable que indica el número del Boton  Anyadir en la Toolbar1
Dim btnAnyadir As Byte

'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1
Dim btnPrimero As Byte


Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos

Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal

Dim gridCargado As Boolean 'Saber si el grid esta cargado cuando se ejecuta DataGrid1_RowColChange

Dim AlbCompleto As Boolean 'Si se va a servir el Pedido Completo (slialb.cantidad=sliped.cantidad)
                            'o se va a servir una parte (slialb.cantidad=sliped.servidas)



'================================================================================

Private Sub cmdAceptar_Click()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim Sql As String
Dim PrimeraLin As Boolean 'Si se inserta la primera linea no esta creado el datagrid1 entonces llamar
                          ' a DataGrid, sino llamar solo a DataGrid2

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR Cabecera Pedido
            If DatosOk Then
                Set vTipoMov = New CTiposMov
                If vTipoMov.Leer(CodTipoMov) Then
                    text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
                    Sql = CadenaInsertarDesdeForm(Me)
                    If Sql <> "" Then
                        If InsertarPedido(Sql, vTipoMov) Then
                            CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                            PonerCadenaBusqueda
                            PonerModo 2
                            'Ponerse en Modo Insertar Lineas
                            BotonMtoLineas 1, "Pedidos"
                            BotonAnyadirLinea
                        End If
                    End If
                    FormateaCampo text1(0)
                End If
                Set vTipoMov = Nothing
            End If
            Me.SSTab1.Tab = 0
            
        Case 4  'MODIFICAR Cabecera Pedido
            If DatosOk Then
                If ModificaDesdeFormulario1(Me, 1) Then
                    'Actualizar los datos del Proveedor si es de varios
                    ActualizarProveVarios text1(4).Text, text1(6).Text
                    TerminaBloquear
                    PosicionarData
                End If
            End If
            
         Case 5 'InsertarModificar LINEA
            'Actualizar el registro en la tabla de lineas 'sliped'
            If ModificaLineas = 1 Then 'INSERTAR lineas Pedidos
                PrimeraLin = False
                If Data2.Recordset.EOF = True Then PrimeraLin = True
                If InsertarLinea Then
                    If PrimeraLin Then
                        CargaGrid DataGrid1, Data2, True
                    Else
                        CargaGrid2 DataGrid1, Data2
                    End If
                    BotonAnyadirLinea
                End If
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then
                    TerminaBloquear
                    ModificaLineas = 0
                    CargaTxtAux False, False
'--monica: rollo toolbar
'                    PonerBotonCabecera True
                    BloquearTxt Text2(16), True
                    BloquearTxt Text2(4), True
                    BloquearTxt Text2(5), True
                    CargaGrid2 DataGrid1, Data2
                End If
                Me.DataGrid1.Enabled = True
'++monica: rollo
                PonerModo 2
                PonerCampos
            End If
            CalcularDatosFactura
            
        Case 6 'Pasar Pedido a Albaran
            If BLOQUEADesdeFormulario(Me) Then GenerarAlbaran
            TerminaBloquear
    End Select
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdAux_Click(Index As Integer)
    Select Case Index
        Case 0 'Busqueda de Cod. Almacen
            Set frmAlm = New frmManAlmProp
            frmAlm.DatosADevolverBusqueda = "0|"
            frmAlm.Show vbModal
            Set frmAlm = Nothing
            
        Case 1 'Busqueda de Cod. Artic
            Set frmArt = New frmManArtic
            frmArt.DatosADevolverBusqueda = "0|1|" 'Poner en modo busqueda
            frmArt.Show vbModal
            Set frmArt = Nothing
    End Select
    PonerFoco txtAux(Index)
End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1, 3 'Busqueda, Insertar
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
            PonerFoco text1(0)
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco text1(0)
            
        Case 5 'Lineas Detalle
            TerminaBloquear
            CargaTxtAux False, False
            BloquearTxt Text2(16), True
            BloquearTxt Text2(4), True
            BloquearTxt Text2(5), True
            If ModificaLineas = 1 Then 'INSERTAR
                ModificaLineas = 0
                DataGrid1.AllowAddNew = False
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            Else
                ModificaLineas = 0
            End If
'--monica: rollo toolbar
'            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
            PonerModo 2
            PonerCampos
            
        Case 6 'Insertar servidas en Generar Albaran (Pedido --> Albaran)
            TerminaBloquear
            InicializarServidas
            PonerModo 2
            CargaTxtAuxServidas False, False
            CargaGrid DataGrid1, Data2, True, False
    End Select
'++monica: rollo toolbar
    If Not Data2.Recordset.EOF Then
        CargaForaGrid
    Else
        LimpiarCampos
    End If

End Sub


Private Sub BotonAnyadir()
'Añadir registro en tabla de cabecera de Pedidos: scaped (Cabecera)
Dim NomTraba As String

    LimpiarCampos 'Vacía los TextBox
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    
'--monica
'    'Poner el nombre del trabajador que esta conectado
'    Text1(3).Text = PonerTrabajadorConectado(NomTraba)
'    Text2(3).Text = NomTraba

    text1(1).Text = Format(Now, "dd/mm/yyyy") 'Fecha Oferta
    PonerFoco text1(1)
End Sub


Private Sub BotonAnyadirLinea()
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
    ModificaLineas = 1 'Ponemos Modo Añadir Linea
'--monica:rollo toolbar
'    'Añadiremos el boton de aceptar y demas objetos para insertar
'    PonerBotonCabecera False
'    lblIndicador.Caption = "INSERTAR"
    
    AnyadirLinea DataGrid1, Data2
    CargaTxtAux True, True
    
    'Poner el Almacen por defecto del Trabajador
    
'--monica
'    txtAux(0).Text = DevuelveDesdeBDNew(cAgro, "straba", "codalmac", "codtraba", Text1(3).Text, "N")
'++monica añadido a piñon
    txtAux(0).Text = vParamAplic.Almacen
    If txtAux(0).Text <> "" Then txtAux(0).Text = Format(txtAux(0).Text, "000")
    
    'Campo Ampliacion Linea
    Text2(16).Text = ""
    Text2(4).Text = ""
    Text2(5).Text = ""
    
    BloquearTxt Text2(16), False
    
    Text2(0).Text = ""
    Text2(1).Text = ""
    
    PonerFoco txtAux(1)
    Me.DataGrid1.Enabled = False
End Sub


Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        PonerModo 1
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
'    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        LimpiarCampos
        LimpiarDataGrids
        CadenaConsulta = "Select * from " & NombreTabla & " " & Ordenacion
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
'Prepara el Form para Modificar la cabecera de Pedidos (tabla: scaped)
Dim DeVarios As Boolean
Dim Sql As String
On Error GoTo EModificar

    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    PonerFoco text1(1)
            
    If EsDeVarios Then
        If Data1.Recordset.EOF Then Exit Sub
        Sql = " SELECT * FROM sprvar WHERE nifprove='" & Data1.Recordset!nifProve & "' FOR UPDATE "
        conn.Execute Sql
    End If
    
     'Si es Cliente de Varios no se pueden modificar sus datos
    DeVarios = EsProveedorVarios(text1(4).Text)
    BloquearDatosProve (DeVarios)
    
EModificar:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub BotonModificarLinea()
'Prepara el Form para Modificar una linea de Pedido (tabla: sliped)
Dim vWhere As String
On Error GoTo eModificarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    
    If Data2.Recordset.EOF Then Exit Sub
    vWhere = ObtenerWhereCP(False) & " and numlinea=" & Data2.Recordset!NumLinea
    vWhere = Replace(vWhere, NombreTabla, NomTablaLineas)
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
    
    CargaTxtAux True, False
    ModificaLineas = 2 'Modificar
'--monica:rollo toolbar
'    'Añadiremos el boton de aceptar y demas objetos para insertar
'    Me.lblIndicador.Caption = "MODIFICAR"
'    PonerBotonCabecera False
    BloquearTxt Text2(16), False 'Campo Ampliacion Linea
    BloquearTxt txtAux(2), True 'campo nombre articulo
    PonerFoco txtAux(0)
    Me.DataGrid1.Enabled = False
    
eModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Pedidos (scaped)
' y los registros correspondientes de las tablas de lineas (sliped)
Dim Cad As String
Dim vTipoMov As CTiposMov
Dim NumPedElim As Long 'Numero del Pedido que se ha Eliminado

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    Cad = "Cabecera de Pedidos Compras." & vbCrLf
    Cad = Cad & "--------------------------------------" & vbCrLf & vbCrLf
    Cad = Cad & "Va a eliminar el Pedido:            "
    Cad = Cad & vbCrLf & "Nº:  " & Format(text1(0).Text, "0000000")
    Cad = Cad & vbCrLf & "Proveedor:  " & Format(text1(4).Text, "000000") & " - " & text1(5).Text
    Cad = Cad & vbCrLf & vbCrLf & " ¿Desea Eliminarlo? "
       
    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        Screen.MousePointer = vbHourglass
        
        NumRegElim = Data1.Recordset.AbsolutePosition
        NumPedElim = Data1.Recordset.Fields(0).Value
        
        CadenaSQL = ""
        Set frmList = New frmListadoOfer
        frmList.Opcionlistado = 81
        frmList.Show vbModal
        Set frmList = Nothing
    
        If CadenaSQL = "" Then Exit Sub
        Cad = ""
        Cad = DBSet(RecuperaValor(CadenaSQL, 1), "F") & " as fechelim,"
'--monica
'        Cad = Cad & RecuperaValor(CadenaSQL, 2) & " as trabelim,"
        Cad = Cad & DBSet(RecuperaValor(CadenaSQL, 2), "T") & " as codincid"
        CadenaSQL = Cad
        
        
        If Not Eliminar() Then Exit Sub
        PosicionarDataTrasEliminar
        
        'Devolvemos contador, si no estamos actualizando
        Set vTipoMov = New CTiposMov
        vTipoMov.DevolverContador CodTipoMov, NumPedElim
        Set vTipoMov = Nothing
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
        Screen.MousePointer = vbDefault
        If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Pedido", Err.Description
End Sub


Private Sub BotonEliminarLinea()
'Eliminar una linea Del Pedido. (Tabla: sliped)
Dim Sql As String
On Error GoTo EEliminarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar

    If Data2.Recordset.EOF Then Exit Sub
            
    ModificaLineas = 3 'Eliminar
    Sql = "¿Seguro que desea eliminar la línea del Pedido?     "
    Sql = Sql & vbCrLf & "NumLinea:  " & Data2.Recordset!NumLinea & vbCrLf
    Sql = Sql & "Almacen:  " & Format(Data2.Recordset!codAlmac, "000")
    Sql = Sql & vbCrLf & "Artículo:  " & Data2.Recordset!codArtic & " - " & Data2.Recordset!NomArtic
    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Data2.Recordset.AbsolutePosition
        Sql = "Delete from " & NomTablaLineas & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
        
        Sql = Sql & " and numlinea=" & Data2.Recordset!NumLinea
        conn.Execute Sql
        
        ModificaLineas = 0
        CargaGrid2 DataGrid1, Data2
        SituarDataTrasEliminar Data2, NumRegElim
        CalcularDatosFactura
        '++monica: rollo
        PonerModo 2
'        CancelaADODC
    End If
    PonerFocoBtn Me.cmdRegresar
    
EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas Mantenimientos", Err.Description
End Sub


Private Sub BotonGenerarAlbaran()
    'Pasar una Pedido a Albaran
Dim Resp As Byte

    'Comprobar que hay un Pedido seleccionado
    If text1(0).Text = "" Then Exit Sub
        
    'Preguntar si se Recibe el pedido completo o no
    Resp = MsgBox("¿Recibir el pedido completo?", vbYesNoCancel + vbQuestion)
    If Resp = vbCancel Then Exit Sub
    
    If Resp = vbYes Then 'RECIBIR EL PEDIDO COMPLETO
        AlbCompleto = True
        Screen.MousePointer = vbHourglass

        GenerarAlbaran
        TerminaBloquear
        
    ElseIf Resp = vbNo Then 'RECIBIR PEDIDO INCOMPLETO
        AlbCompleto = False
        Me.SSTab1.Tab = 0
        TerminaBloquear
        'Si no se va a servir completo Mostrar lineas para que se indiquen las Servidas
        MsgBox "Introduzca la cantidad  a recibir para cada línea.", vbInformation
        Modo = 6
        gridCargado = False
        Me.cmdAceptar.visible = True
        Me.cmdCancelar.visible = True
        PonerModoOpcionesMenu Modo
        CargaGrid DataGrid1, Data2, True, True
        CargaTxtAuxServidas True, True
        PrimeraVez = True
    Else
        TerminaBloquear
    End If

End Sub





Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim Cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        'BloquearTabs False
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid DataGrid1
            DataGrid1.Bookmark = 1
        End If
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


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim i As Integer
On Error GoTo Error1
    If Modo = 6 And gridCargado Then '6: Pasar Pedido a Albaran no Completo (Introducir las servidas)
        CargaTxtAuxServidas True, True
        txtAux(3).Text = Format(Data2.Recordset!recibida, FormatoImporte)
    End If
'    If Modo = 5 Then 'Poner el valor al camp ampliacion linea '5: modo lineas
        If Not Data2.Recordset.EOF And ModificaLineas <> 1 Then '1: Insertar
            'Poner descripcion de ampliacion lineas
            Text2(16).Text = DevuelveDesdeBDNew(cAgro, NomTablaLineas, "ampliaci", "numpedpr", text1(0).Text, "N", , "numlinea", Data2.Recordset!NumLinea, "N")
            CargarDatosArticulo (Data2.Recordset!codArtic)
        Else
            Text2(16).Text = ""
            Text2(4).Text = ""
            Text2(5).Text = ""
        End If
'    End If
    
    
'++monica:rollo toolbar
    If Modo <> 4 Then 'Modificar
        CargaForaGrid
    Else
        For i = 0 To txtAux.Count - 1
            txtAux(i).Text = ""
        Next i
    End If
    
Error1:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PrimeraVez Then
        PrimeraVez = False
        'Viene de documentos de la ficha de proveedores
        If Pedido <> "" Then
            text1(0).Text = Pedido
            HacerBusqueda
        End If
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
Dim i As Integer
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    ' ICONITOS DE LA BARRA
    btnAnyadir = 5
'    btnPrimero = 19
'    With Me.Toolbar1
'        .HotImageList = frmPpal.imgListComun_OM
'        .DisabledImageList = frmPpal.imgListComun_BN
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 1   'Botón Buscar
'        .Buttons(2).Image = 2   'Botón Todos
'        .Buttons(5).Image = 3   'Insertar Nuevo
'        .Buttons(6).Image = 4   'Modificar
'        .Buttons(7).Image = 5   'Borrar
'        .Buttons(10).Image = 15 'Mto Lineas Ofertas
'        .Buttons(11).Image = 16 'Generar Albaran
'
'        .Buttons(14).Image = 10 'Imprimir Pedido
'        .Buttons(16).Image = 11  'Salir
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
        .Buttons(1).Image = 16 'Generar Albaran
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
    
    Me.SSTab1.Tab = 0
      
    LimpiarCampos   'Limpia los campos TextBox
    
    CodTipoMov = "PEC"
    VieneDeBuscar = False
          
    '## A mano
     Me.FrameHco.visible = EsHistorico
    
    If Not EsHistorico Then
        NombreTabla = "scappr"
        NomTablaLineas = "slippr" 'Tabla lineas de Pedido
        Me.Caption = "Pedidos Proveedores"
        Ordenacion = " ORDER BY numpedpr "
    Else
        NombreTabla = "schppr"
        NomTablaLineas = "slhppr"
        CargarTagsHco Me, "scappr", NombreTabla
        'Estos campos solo estan en la tabla del histórico
        text1(24).Tag = "Fecha Eliminación|F|N|||" & NombreTabla & "|fechelim|dd/mm/yyyy|N|"
'--monica
'        Text1(25).Tag = "Trabajador Eliminación|N|N|0|9999|" & NombreTabla & "|trabelim|0000|N|"
        text1(26).Tag = "Incidencia elim.|T|N|||" & NombreTabla & "|codincid||N|"
        Me.Caption = "Histórico Pedidos Proveedores"
        Ordenacion = " ORDER BY numpedpr,fecpedpr "
    End If
    
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    
    CadenaConsulta = "select * from " & NombreTabla
    If Pedido <> "" Then
        CadenaConsulta = CadenaConsulta & " where numpedpr = " & Pedido
    Else
        CadenaConsulta = CadenaConsulta & " where numpedpr = -1 "
    End If
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta '"Select * from " & NombreTabla & " where numpedpr=-1"
    Data1.Refresh
'    If DatosADevolverBusqueda = "" Then
'        PonerModo 0
'    Else
'        PonerModo 1
'        Text1(0).BackColor = vbYellow
'    End If
    
    
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
     
     'Poner los grid sin apuntar a nada
    LimpiarDataGrids
   
    If DatosADevolverBusqueda = "" Then
        If Pedido = "" Then
            PonerModo 0
'        Else
'            Text1(0).Text = Pedido
'            HacerBusqueda
        End If
    Else
        BotonBuscar
    End If

    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True


End Sub


Private Sub LimpiarCampos()
On Error Resume Next

    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.chkRestoPed.Value = 0
    
    Text3(0).Text = "TOTAL PEDIDO"
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    conn.Execute "DELETE FROM tmpnseries WHERE codusu=" & vUsu.Codigo
    If Err.Number <> 0 Then Err.Clear
    
    If Modo = 4 Or Modo = 5 Then TerminaBloquear
End Sub

Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Almacenes Propios
    txtAux(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Almacen
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Articulos
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
    txtAux(2).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Artic
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
            If EsHistorico Then
                Aux = ValorDevueltoFormGrid(text1(1), CadenaDevuelta, 2)
                CadB = CadB & " and " & Aux
            End If
            
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
            PonerCadenaBusqueda
            text1(0).Text = Format(RecuperaValor(CadenaDevuelta, 1), "0000000")
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    text1(22).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod cliente
    FormateaCampo text1(22)
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom clien
End Sub

'--monica
'Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
''Formulario Mantenimiento C. Postales
'Dim indice As Byte
'Dim devuelve As String
'
'    indice = 9
'    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
'    'poblacion
'    Text1(indice + 1).Text = ObtenerPoblacion(Text1(indice).Text, devuelve)
'    'provincia
'    Text1(indice + 2).Text = devuelve
'End Sub


Private Sub frmDir_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Direcciones
Dim indice As Byte
    indice = CByte(Me.imgBuscar(0).Tag)
    text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Direccion
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Direc

    CargarDatosDirec text1(indice).Text, indice
End Sub

Private Sub frmC_Selec(vFecha As Date) 'Calendario Fechas
Dim indice As Byte
    indice = CByte(Me.imgFec(0).Tag) + 1
    text1(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
Dim indice As Byte

    indice = 12
    text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Forma Pago
    FormateaCampo text1(indice)
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub


Private Sub frmInc_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de incidencias
    text1(26).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod incidencia
    Text2(26).Text = RecuperaValor(CadenaSeleccion, 2) 'nom incidencia
End Sub

Private Sub frmList_DatoSeleccionado(CadenaSeleccion As String)
'Aqui devuelve los valores que se introducen en el Listado
'para pasar de Pedido a Albaran, o para pasar al historico
    
    CadenaSQL = CadenaSeleccion
End Sub


Private Sub frmPed_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String
Dim Aux As String
      
    If CadenaSeleccion <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        CadB = "numpedpr = " & RecuperaValor(CadenaSeleccion, 1)
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
    Screen.MousePointer = vbDefault
End Sub

'--monica
'Private Sub frmNSerie_CargarNumSeries()
''Insertar un registro en la tabla "sserie" por cada uno de los
''Nº de Serie introducidos en la Tabla Temporal
'Dim RStmp As ADODB.Recordset
'Dim RsAlb As ADODB.Recordset
'Dim SQL As String
'Dim i As Byte
'Dim b As Boolean
'
'    On Error GoTo EInsertar
'
'
'    SQL = "SELECT slialp.codartic, numlinea, cantidad "
'    SQL = SQL & " FROM slialp INNER JOIN sartic on slialp.codartic=sartic.codartic "
'    SQL = SQL & " WHERE numalbar=" & DBSet(Me.cmdAux(1).Tag, "T") & " and fechaalb=" & DBSet(Me.cmdAux(0).Tag, "F") & " and "
'    SQL = SQL & "slialp.codprove=" & Text1(4).Text
'    SQL = SQL & " And nseriesn = 1 "
'    SQL = SQL & " ORDER BY codartic, numlinea "
'
'    Set RsAlb = New ADODB.Recordset
'    RsAlb.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    While Not RsAlb.EOF 'Para cada linea del ALbaran
'        'Recuperar los Nº Serie de ese articulo cargados en la Temporal
'        'Seleccionar los nº de serie cargados en la temporal: tmpnseries
'        SQL = "SELECT * FROM tmpnseries WHERE codusu=" & vUsu.Codigo
'        SQL = SQL & " AND codartic=" & DBSet(RsAlb!codArtic, "T")
'        SQL = SQL & " ORDER BY codartic, numlinea "
'        Set RStmp = New ADODB.Recordset
'        RStmp.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'        'If Not RStmp.EOF Then RStmp.MoveFirst
'        'Intentar asignar un Nº serie al total de cantidad del articulo
'
'        b = True
'        For i = 1 To RsAlb!Cantidad
'            If Not RStmp.EOF Then
'                InsertarNSerie RStmp!numSerie, RStmp!codArtic, RsAlb!numlinea
'                RStmp.MoveNext
'            End If
'        Next i
'        RStmp.Close
'        Set RStmp = Nothing
'        RsAlb.MoveNext
'    Wend
'    RsAlb.Close
'    Set RsAlb = Nothing
'EInsertar:
'    If Err.Number <> 0 Then MuestraError Err.Number, "Insertando Nº Serie", Err.Description
'End Sub


Private Sub frmProv_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Proveedores
Dim indice As Byte

    indice = 4
    text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Prove
    FormateaCampo text1(indice)
End Sub

Private Sub frmProveV_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento Proveedores varios
Dim indice As Byte

    indice = 6
    text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'nif Prove
    text1(indice - 1).Text = RecuperaValor(CadenaSeleccion, 2) 'nom Prove
    PonerDatosProveVario text1(indice).Text
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Trabajadores
Dim indice As Byte

    indice = Val(Me.imgBuscar(0).Tag)
    text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Trabajador
    FormateaCampo text1(indice)
    If indice = 23 Then indice = 1
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
End Sub


Private Sub imgBuscar_Click(Index As Integer)
Dim indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 0 'Cod. Proveedor
            indice = 4
            Set frmProv = New frmManProve
            frmProv.DatosADevolverBusqueda = "0|"
            frmProv.Show vbModal
            Set frmProv = Nothing
'--monica
'        Case 1 'Cod. Postal
'            Set frmCP = New frmCPostal
'            frmCP.DatosADevolverBusqueda = "0"
'            frmCP.Show vbModal
'            Set frmCP = Nothing
'            indice = 9
'            VieneDeBuscar = True
'
'        Case 2, 8 'Realizada Por Trabajador
'            If Index = 2 Then
'                indice = 3
'            Else
'                indice = 23
'            End If
'            Me.imgBuscar(0).Tag = indice
'            Set frmT = New frmAdmTrabajadores
'            frmT.DatosADevolverBusqueda = "0"
'            frmT.Show vbModal
'            Set frmT = Nothing
            
        Case 3 'Forma de Pago
            indice = 12
            Set frmFP = New frmManFpago
            frmFP.DatosADevolverBusqueda = "0|"
            frmFP.Show vbModal
            Set frmFP = Nothing
            
        Case 4, 5 'Direccion
            If Index = 4 Then indice = 15
            If Index = 5 Then indice = 2
            Me.imgBuscar(0).Tag = indice
            Set frmDir = New frmComDirecciones
            frmDir.DatosADevolverBusqueda = "0|"
            frmDir.Show vbModal
            Set frmDir = Nothing
            
        Case 6 'NIF de Proveedores VARIOS
            indice = 6
            Set frmProveV = New frmComProveV
            frmProveV.DatosADevolverBusqueda = "0|"
            frmProveV.Show vbModal
            Set frmProveV = Nothing
            
        Case 7 'Cliente
            indice = 22
            Set frmCli = New frmClientes
            frmCli.DatosADevolverBusqueda = "0|"
            frmCli.Show vbModal
            Set frmCli = Nothing
            
        Case 10 'Incidencias
            indice = 26
            Set frmInc = New frmManInciden
            frmInc.DatosADevolverBusqueda = "0|"
            frmInc.Show vbModal
            Set frmInc = Nothing
    End Select
    
    PonerFoco text1(indice)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFec_Click(Index As Integer) 'Abre calendario Fechas
Dim indice As Byte

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   '++monica
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
   
   frmC.NovaData = Now
   indice = Index + 1
   Me.imgFec(0).Tag = Index
   
   PonerFormatoFecha text1(indice)
   If text1(indice).Text <> "" Then frmC.NovaData = CDate(text1(indice).Text)

   Screen.MousePointer = vbDefault
   frmC.Show vbModal
   Set frmC = Nothing
   PonerFoco text1(indice)
End Sub



Private Sub mnBuscar_Click()
    Me.SSTab1.Tab = 0
    BotonBuscar
End Sub


Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Pedido
         BotonEliminarLinea
    Else   'Eliminar Pedido
         BotonEliminar
         Screen.MousePointer = vbDefault
    End If
End Sub


Private Sub mnGenAlbaran_Click()
    'bloqueamos el pedido y lo pasamos a Albaran
    If BLOQUEADesdeFormulario(Me) Then BotonGenerarAlbaran
End Sub


Private Sub mnImpPedido_Click()
'Imprime un Pedido
       frmListadoOfer.NumCod = text1(0).Text    'Nº de Pedido
       frmListadoOfer.CodClien = text1(4).Text 'Cod.Proveedor
       If EsHistorico Then
            AbrirListadoOfer (56) '59: Informe de Pedidos Compras (Historico)
            frmListadoOfer.FecEntre = text1(1).Text
       Else
            AbrirListadoOfer (55) '55: Informe de Pedidos Compras
       End If
End Sub

Private Sub mnLineas_Click()
    BotonMtoLineas 0, "Pedidos"
End Sub


Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
         BotonModificarLinea
    Else   'Modificar Pedido
         If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub


Private Sub mnNuevo_Click()
    If Modo = 5 Then 'Añadir lineas
         BotonAnyadirLinea
    Else 'Añadir Cabecera de Pedidos
         Me.SSTab1.Tab = 0
         BotonAnyadir
    End If
End Sub


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


Private Sub SSTab1_Click(PreviousTab As Integer)
    Text2(16).visible = (SSTab1.Tab = 0)
    Label1(35).visible = (SSTab1.Tab = 0)
End Sub

Private Sub SSTab1_DblClick()
    Text2(16).visible = (SSTab1.Tab = 0)
    Label1(35).visible = (SSTab1.Tab = 0)
End Sub

'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
Dim cadkey As Integer

    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
    If Index = 9 Then HaCambiadoCP = False 'CPostal
    ConseguirFoco text1(Index), Modo
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
Dim devuelve As String
Dim i As Byte
        
    If Not PerderFocoGnral(text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
       
    'Si queremos hacer algo ..
    Select Case Index
        Case 1 'Fecha Oferta, Fecha Entrega
            '[Monica]28/08/2013: controlamos que esté dentro de campaña
            PonerFormatoFecha text1(Index), True
            
'--monica
'        Case 3, 23 'Cod Trabajador
'            i = Index
'            If Index = 23 Then i = 1
'            If PonerFormatoEntero(Text1(Index)) Then
'                Text2(i).Text = PonerNombreDeCod(Text1(Index), cAgro, "straba", "nomtraba", "codtraba", "el Trabajador")
'            Else
'                Text2(i).Text = ""
'            End If
            
        Case 4 'Cod. Prove
            If PonerFormatoEntero(text1(Index)) Then
                If Modo = 1 Then 'Busqueda
                    text1(5).Text = PonerNombreDeCod(text1(Index), "proveedor", "nomprove")
                Else ' cargar datos de Tabla sprove
                    PonerDatosProveedor (text1(Index).Text)
                End If
            Else
                LimpiarDatosProve
            End If
            
         Case 6 'NIF
            If Not EsDeVarios Or Modo <> 3 Then Exit Sub
            If Modo = 4 Then 'Modificar
                'si no se ha modificado el nif del cliente no hacer nada
                If text1(6).Text = Data1.Recordset!nifProve Then
                    Exit Sub
                End If
            End If
            PonerDatosProveVario (text1(Index).Text)
            
'--monica
'        Case 9 'Cod. Postal
'            If Text1(Index).Locked Then Exit Sub
'            If Text1(Index).Text = "" Then
'                Text1(Index + 1).Text = ""
'                Text1(Index + 2).Text = ""
'                Exit Sub
'            End If
'            If (Not VieneDeBuscar) Or (VieneDeBuscar And HaCambiadoCP) Then
'                 Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, devuelve)
'                 Text1(Index + 2).Text = devuelve
'            End If
'            VieneDeBuscar = False
            
        Case 12 'Forma de Pago
            If PonerFormatoEntero(text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(text1(Index), "forpago", "nomforpa")
            Else
                Text2(Index).Text = ""
            End If
            
        Case 13, 14 'Descuentos
            If PonerFormatoDecimal(text1(Index), 4) Then   'Tipo 4: Decimal(4,2)
                 If Modo = 4 Then CalcularDatosFactura
            End If
            
        Case 15, 2 'Cod. Direccion
            If PonerFormatoEntero(text1(Index)) Then
                Me.imgBuscar(0).Tag = Index
                If Not CargarDatosDirec(text1(Index).Text, CByte(Index)) Then
                    PonerFoco text1(Index)
                End If
            Else
                LimpiarDatosDirec CByte(Index)
            End If
            
        Case 22 'cod.cliente
            If PonerFormatoEntero(text1(Index)) Then
                Text2(0).Text = PonerNombreDeCod(text1(Index), "clientes", "nomclien")
            Else
                Text2(0).Text = ""
            End If
            
        Case 21
            If Me.ActiveControl.Name = "SSTab1" Then PonerFocoBtn Me.cmdAceptar
            
        Case 26 'cod Incidencia de eliminacion
            If EsHistorico Then
                Text2(Index).Text = PonerNombreDeCod(text1(Index), "inciden", "nomincid")
                If Not (Text2(Index).Text = "" And text1(Index).Text <> "") Then
                    PonerFocoBtn Me.cmdAceptar
                Else
                    PonerFoco text1(Index)
                End If
            End If
    End Select
End Sub


Private Sub HacerBusqueda()
Dim CadB As String

    CadB = ObtenerBusqueda3(Me, False)
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)

    Set frmPed = New frmBasico2
    
    AyudaPedidosCompraPrev frmPed, , CadB
    
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
            PonerFoco text1(0)
        End If
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
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
'Carga las Pestañas con las tablas de lineas del Trabajador seleccionado para mostrar
On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass

    'Datos de la tabla slippr
    CargaGrid DataGrid1, Data2, True

    Screen.MousePointer = vbDefault
    Exit Sub
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    'Realizado por
'--monica
'    Text2(3).Text = PonerNombreDeCod(Text1(3), conAri, "straba", "nomtraba")
    Text2(12).Text = PonerNombreDeCod(text1(12), "forpago", "nomforpa")
    'Cliente para
    Text2(0).Text = PonerNombreDeCod(text1(22), "clientes", "nomclien")
'--monica
'    'Solicitado por
'    Text2(1).Text = PonerNombreDeCod(Text1(23), conAri, "straba", "nomtraba", "codtraba")
    
    'Poner las direcciones
    CargarDatosDirec text1(15).Text, 15
    CargarDatosDirec text1(2).Text, 2
    
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Pedidos
    
    If EsHistorico Then
        'poner datos de eliminacion
'--monica
'        Text2(25).Text = PonerNombreDeCod(Text1(25), conAri, "straba", "nomtraba", "codtraba")
        Text2(26).Text = PonerNombreDeCod(text1(26), "inciden", "nomincid", "codincid")
    End If
    
    CalcularDatosFactura 'rellenar campos pestaña de totales
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu Modo
    PonerOpcionesMenu
    
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

'    'Actualiza Iconos Insertar,Modificar,Eliminar
'--monica: rollo toolbar
'    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    If Modo = 6 Then Me.lblIndicador.Caption = "Insertar Cant. Servidas"
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Or Pedido <> "" Then
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
'    BloquearText1 Me, Modo
    b = (Modo = 3 Or Modo = 4 Or Modo = 1)
    For i = 0 To text1.Count - 1
        BloquearTxt text1(i), Not b
        text1(i).Enabled = b
    Next i
    
    'Campo Numero de Albaran siempre bloqueado, excepto si estamos en modo de busqueda
    b = (Modo <> 1)
    BloquearTxt text1(0), b, True
       
    'datos cliente siempre bloqueados hasta que sea de varios
    If Modo = 3 Then
        EsDeVarios = False
        BloquearDatosProve (EsDeVarios)
    End If
       
       
    '-----  Datos Totales de Factura siempre bloqueado
    For i = 33 To 50
        BloquearTxt Text3(i), True
    Next i
    'Campo B.Imp y Imp. IVA siempre en azul
    Text3(36).BackColor = &HFFFFC0
    Text3(46).BackColor = &HFFFFC0
    Text3(47).BackColor = &HFFFFC0
    Text3(48).BackColor = &HFFFFC0
    Text3(49).BackColor = &HC0C0FF    'Tatal factura
    Text3(50).BackColor = &HC0C0FF    'Tatal factura
    '---------------------------------------------------
       
       
    'Si no es modo lineas Boquear los TxtAux
    For i = 0 To txtAux.Count - 1
        BloquearTxt txtAux(i), (Modo <> 5)
    Next i
    txtAux(2).Enabled = False
    BloquearTxt Text2(16), (Modo <> 5)
    BloquearTxt Text2(4), True
    BloquearTxt Text2(5), True
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2) '--monica: rollo toolbar And Modo <> 5)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    For i = 0 To Me.imgFec.Count - 1
        Me.imgFec(i).Enabled = b
    Next i
    b = (Modo <> 0 And Modo <> 2 And Modo <> 5)
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = b
    Next i
    Me.imgBuscar(1).visible = False
           
    'Modo Linea de Ofertas. Poner el campo ampliacion linea
'--monica: rollo toolbar
'    Me.Label1(35).visible = (Modo = 5)
'    Me.Text2(16).visible = (Modo = 5)
    BloquearTxt Text2(16), True
    BloquearTxt Text2(4), True
    BloquearTxt Text2(5), True
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

Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Generar Albaran
            mnGenAlbaran_Click
    End Select
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
'Comprueba si los datos de la cabecera son correctos antes de Insertar o Modificar el
'Pedido
Dim b As Boolean
On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm(Me) 'Comprobar formato datos ok
    If Not b Then Exit Function
            
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
'Comprueba si los datos de una linea son correctos antes de Insertar o Modificar
'una linea del Pedido
Dim b As Boolean
'Dim devuelve As String
Dim i As Byte
Dim vArtic As CArticulo

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    b = True

    'Comprobar que los campos NOT NULL tienen valor
    For i = 0 To txtAux.Count - 1
        If txtAux(i).Text = "" Then
            MsgBox "El campo " & txtAux(i).Tag & " no puede ser nulo", vbExclamation
            b = False
            PonerFoco txtAux(i)
            Exit Function
        End If
    Next i
        
    'Comprobar que existe el articulo en el almacen seleccionado
    Set vArtic = New CArticulo
    vArtic.Codigo = txtAux(1).Text
    If Not vArtic.ExisteEnAlmacen(txtAux(0).Text) Then
        b = False
        PonerFoco txtAux(1)
    End If
    Set vArtic = Nothing
    
'    devuelve = DevuelveDesdeBDNew(conAri, "salmac", "codartic", "codartic", txtAux(1).Text, "T", , "codalmac", txtAux(0).Text, "N")
'    If devuelve = "" Then
'        MsgBox "No existen unidades del Artículo: " & txtAux(1).Text & "  en el Almacen: " & txtAux(0).Text, vbExclamation
'        b = False
'        PonerFoco txtAux(1)
'    End If
    
    DatosOkLinea = b
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 16 And KeyCode = 40 Then 'campo Ampliacion linea y Flecha hacia abajo
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 16 And KeyAscii = 13 Then 'campo Ampliación linea y ENTER
        PonerFocoBtn Me.cmdAceptar
    End If
End Sub

Private Sub Text2_LostFocus(Index As Integer)
    If Index = 16 And (Text2(Index).Locked = False) Then Text2(Index).Text = UCase(Text2(Index).Text)
End Sub

'++monica : rollo toolbar
Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
Dim vWhere As String

    BotonMtoLineas 0, "Pedidos"
'-- pon el bloqueo aqui
    'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
    Select Case Button.Index
        Case 1
            mnNuevo_Click 'BotonAnyadirLineas
        Case 2
            vWhere = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas) & " and numlinea=" & Me.Data2.Recordset.Fields(1)
            'vWhere = ObtenerWhereCP(False) & " and numlinea=" & Me.Data2.Recordset.Fields(1)
            If BloqueaRegistro(NomTablaLineas, vWhere) Then mnModificar_Click 'BotonModificarLinea
        Case 3
            mnEliminar_Click 'BotonEliminarLinea
        Case Else
    End Select
    'End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 5  'Buscar
            mnBuscar_Click
        Case 6  'Todos
            BotonVerTodos
        Case 1  'Nuevo
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
        Case 8 'Imprimir Pedido
             mnImpPedido_Click
    End Select
End Sub


Private Sub PonerOpcionesMenu()
Dim J As Byte

    PonerOpcionesMenuGeneral Me
       
    J = Val(Me.mnGenAlbaran.HelpContextID)
    If J < vUsu.Nivel Then Me.mnGenAlbaran.Enabled = False
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub
    
    
Private Function InsertarLinea() As Boolean
'Inserta un registro en la tabla de lineas de Pedido: slipre
Dim Sql As String
Dim NumLinea As String, vWhere As String
On Error GoTo EInsertarLinea

    InsertarLinea = False
    Sql = ""

    If DatosOkLinea() Then 'Lineas de Pedidos
        'Conseguir el siguiente numero de linea
        vWhere = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
        NumLinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
        
        Sql = "INSERT INTO " & NomTablaLineas
        Sql = Sql & "(numpedpr,numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, recibida, precioar, dtoline1, dtoline2, importel) "
        Sql = Sql & "VALUES (" & Val(text1(0).Text) & ", " & NumLinea & ", " & Val(txtAux(0).Text) & ","
        Sql = Sql & DBSet(txtAux(1).Text, "T") & ", " & DBSet(txtAux(2).Text, "T") & ", " & DBSet(Text2(16).Text, "T") & ", "
        Sql = Sql & DBSet(txtAux(3).Text, "N") & ", 0,"
        Sql = Sql & DBSet(txtAux(4).Text, "N") & "," & DBSet(txtAux(5).Text, "N") & ", "
        Sql = Sql & DBSet(txtAux(6).Text, "N") & ", " 'Dto 2
        Sql = Sql & DBSet(txtAux(7).Text, "N") & ")"
    End If
    
    If Sql <> "" Then
        conn.Execute Sql
        InsertarLinea = True
    End If
    Exit Function
EInsertarLinea:
    MuestraError Err.Number, "Insertar Lineas Pedido" & vbCrLf & Err.Description
End Function


Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Pedido: sliped
Dim Sql As String
On Error GoTo eModificarLinea

    ModificarLinea = False
    Sql = ""
    
    If DatosOkLinea() Then
        'Creamos la sentencia SQL
        Sql = "UPDATE " & NomTablaLineas & " Set codalmac = " & txtAux(0).Text & ", codartic=" & DBSet(txtAux(1).Text, "T") & ", "
        Sql = Sql & "nomartic=" & DBSet(txtAux(2).Text, "T") & ", ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
        Sql = Sql & "cantidad= " & DBSet(txtAux(3).Text, "N") & ", "
        Sql = Sql & "precioar= " & DBSet(txtAux(4).Text, "N") & ", "
        Sql = Sql & "dtoline1= " & DBSet(txtAux(5).Text, "N") & ", dtoline2= " & DBSet(txtAux(6).Text, "N") & ", "
        Sql = Sql & "importel= " & DBSet(txtAux(7).Text, "N")
        Sql = Sql & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea=" & Data2.Recordset!NumLinea
    End If
    
    If Sql <> "" Then
        conn.Execute Sql
        ModificarLinea = True
    End If
    Exit Function
eModificarLinea:
    MuestraError Err.Number, "Modificar Lineas Pedido" & vbCrLf & Err.Description
End Function


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
    
    'Habilitar las opciones correctas del menu según Modo
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu según Nivel de Acceso
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean, Optional conServidas As Boolean)
'IN: enlaza= si carga el grid con valores de la tabla o lo muestra vacio si no enlaza
'    conServidas=si enlaza, se muestra la columna de servidas solo cuando se va a generar el Albaran no completo
Dim b As Boolean
Dim Sql As String

On Error GoTo ECargaGRid

    b = DataGrid1.Enabled
    
    Sql = MontaSQLCarga(enlaza, conServidas)
    CargaGridGnral vDataGrid, vData, Sql, PrimeraVez

    If conServidas Then
        vDataGrid.ClearFields
        vDataGrid.ReBind
        vDataGrid.Refresh
    End If
    
    CargaGrid2 vDataGrid, vData, conServidas
    vDataGrid.ScrollBars = dbgAutomatic
    
    b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2) '5:Modo Mto Lineas (Insertando o Modificando linea)
    vDataGrid.Enabled = Not b
    PrimeraVez = False
    gridCargado = True
    Exit Sub
ECargaGRid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, Optional conServidas As Boolean)
Dim i As Byte
On Error GoTo ECargaGRid

    vData.Refresh
    
    vDataGrid.Columns(0).visible = False
    vDataGrid.Columns(1).visible = False
    i = 1
    Select Case vDataGrid.Name
        Case "DataGrid1" 'Cod. Almacen
                i = i + 1
                vDataGrid.Columns(i).Caption = "Alm."
                If conServidas Then
                    vDataGrid.Columns(i).Width = 550 '450
                Else
                    vDataGrid.Columns(i).Width = 600 '500
                End If
                vDataGrid.Columns(i).NumberFormat = "000"
                
                i = i + 1
                vDataGrid.Columns(i).Caption = "Artículo"
                If conServidas Then
                    vDataGrid.Columns(i).Width = 1800 '1600
                Else
                    vDataGrid.Columns(i).Width = 1900 '1700
                End If
                
                i = i + 1
                vDataGrid.Columns(i).Caption = "Descripción"
                If conServidas Then
                    vDataGrid.Columns(i).Width = 3300 '3100
                Else
                    vDataGrid.Columns(i).Width = 3600 '3400
                End If
                
                i = i + 1
                vDataGrid.Columns(i).Caption = "Ampl. Línea"
                vDataGrid.Columns(i).Width = 7980
                vDataGrid.Columns(i).visible = False
                
                i = i + 1
                vDataGrid.Columns(i).Caption = "Cantidad"
                vDataGrid.Columns(i).Width = 1250 '900
                vDataGrid.Columns(i).Alignment = dbgRight
                vDataGrid.Columns(i).NumberFormat = FormatoImporte
                
                i = i + 1
                If conServidas Then
                    'Cargar el grid con la columna de cantidad servida
                    vDataGrid.Columns(i).Caption = "Recibidas"
                    vDataGrid.Columns(i).Width = 800
                    vDataGrid.Columns(i).Alignment = dbgRight
                    vDataGrid.Columns(i).NumberFormat = FormatoImporte
                    i = i + 1
                End If
                vDataGrid.Columns(i).Caption = "Precio"
                vDataGrid.Columns(i).Width = 1300 '1100
                vDataGrid.Columns(i).Alignment = dbgRight
                vDataGrid.Columns(i).NumberFormat = FormatoPrecio
                
                
                i = i + 1
                vDataGrid.Columns(i).Caption = "Dto.1"
                If conServidas Then
                    vDataGrid.Columns(i).Width = 600 '550
                Else
                    vDataGrid.Columns(i).Width = 650 '600
                End If
                vDataGrid.Columns(i).Alignment = dbgRight
                vDataGrid.Columns(i).NumberFormat = FormatoDescuento
                
                i = i + 1
                vDataGrid.Columns(i).Caption = "Dto.2"
                If conServidas Then
                    vDataGrid.Columns(i).Width = 600 '550
                Else
                    vDataGrid.Columns(i).Width = 650 '600
                End If
                vDataGrid.Columns(i).Alignment = dbgRight
                vDataGrid.Columns(i).NumberFormat = FormatoDescuento
            
                i = i + 1
                vDataGrid.Columns(i).Caption = "Importe Línea"
                If conServidas Then
                    vDataGrid.Columns(i).Width = 1500 '1250
                Else
                    vDataGrid.Columns(i).Width = 1650 '1400
                End If
                vDataGrid.Columns(i).Alignment = dbgRight
                vDataGrid.Columns(i).NumberFormat = FormatoImporte
                
    End Select
    
'    ' *** Si n'hi han camps fora del grid ***
'    If Not Data2.Recordset.EOF Then
'        CargaForaGrid
'    Else
'        LimpiarCampos
'    End If
    ' **************************************
    
    vDataGrid.RowHeight = 350
    

    For i = 0 To vDataGrid.Columns.Count - 1
        vDataGrid.Columns(i).Locked = True
        vDataGrid.Columns(i).AllowSizing = False
    Next i
    vDataGrid.HoldFields
    Exit Sub
ECargaGRid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim i As Byte

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For i = 0 To txtAux.Count - 1 'TextBox
            txtAux(i).Top = 290
            txtAux(i).visible = visible
        Next i
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            For i = 0 To txtAux.Count - 1
                txtAux(i).Text = ""
                BloquearTxt txtAux(i), False
                txtAux(i).Enabled = True
            Next i
        Else 'Vamos a modificar
            For i = 0 To txtAux.Count - 1
                If i < 3 Then
                    txtAux(i).Text = DataGrid1.Columns(i + 2).Text
                Else
                    txtAux(i).Text = DataGrid1.Columns(i + 3).Text
                End If
                txtAux(i).Locked = False
                
            Next i
        End If
        
        txtAux(2).Enabled = False
        'El campo Importe es calculado y lo bloqueamos.
        BloquearTxt txtAux(7), True

        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 20)
        
        For i = 0 To txtAux.Count - 1
            txtAux(i).Top = alto
            txtAux(i).Height = DataGrid1.RowHeight
        Next i
        cmdAux(0).Top = alto
        cmdAux(1).Top = alto
        cmdAux(0).Height = DataGrid1.RowHeight
        cmdAux(1).Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Cod. Almac
        txtAux(0).Left = DataGrid1.Left + 330
        txtAux(0).Width = DataGrid1.Columns(2).Width - 160
        cmdAux(0).Left = txtAux(0).Left + txtAux(0).Width - 40
        'Cod Artic
        txtAux(1).Left = cmdAux(0).Left + cmdAux(0).Width + 20
        txtAux(1).Width = DataGrid1.Columns(3).Width - 160
        cmdAux(1).Left = txtAux(1).Left + txtAux(1).Width - 50
        'Nom Artic
        txtAux(2).Left = cmdAux(1).Left + cmdAux(1).Width
        txtAux(2).Width = DataGrid1.Columns(4).Width - 10
        'Cantidad
        txtAux(3).Left = txtAux(2).Left + txtAux(2).Width + 10
        txtAux(3).Width = DataGrid1.Columns(6).Width - 10
        'Precio, Dto1, Dto2, Precio
        For i = 4 To txtAux.Count - 1
            txtAux(i).Left = txtAux(i - 1).Left + txtAux(i - 1).Width + 10
            txtAux(i).Width = DataGrid1.Columns(i + 3).Width - 10
        Next i
        
        'Los ponemos Visibles o No
        '--------------------------
        For i = 0 To txtAux.Count - 1
            txtAux(i).visible = visible
        Next i
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
    End If
End Sub


Private Sub CargaTxtAuxServidas(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
'Carga el TxtAux(3) con el campo RECIBIDAS de la tabla slippr
Dim alto As Single
Dim i As Byte

    i = 3
    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        txtAux(i).Top = 290
        txtAux(i).visible = visible
        txtAux(i).BackColor = vbWhite
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            txtAux(i).Text = ""
            BloquearTxt txtAux(i), False
            txtAux(i).BackColor = &H80000013
        End If
      
        'Fijamos altura(Height) y posición Top
        '-------------------------------
        If DataGrid1.Row < 0 Then
            alto = DataGrid1.Top + 230
        Else
            alto = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + 20
        End If
        
        txtAux(i).Top = alto
        txtAux(i).Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Cantidad servida
        alto = DataGrid1.Left + 330 + DataGrid1.Columns(2).Width + DataGrid1.Columns(3).Width
        alto = alto + DataGrid1.Columns(4).Width + DataGrid1.Columns(6).Width
        txtAux(i).Left = alto
        txtAux(i).Width = DataGrid1.Columns(7).Width - 15
        
        'Los ponemos Visibles o No
        '--------------------------
        txtAux(i).visible = visible
        PonerFoco txtAux(i)
    End If
End Sub


Private Sub txtAux_GotFocus(Index As Integer)
Dim cadkey As Integer

    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
    ConseguirFocoLin txtAux(Index) '--monica , cadkey
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If Modo <> 6 Then
        KEYpress KeyAscii
    Else 'Pasar el Pedido a Albaran
        If KeyAscii = 13 Then 'ENTER
            PonerServidas
'            ConseguirFoco txtAux(3), Modo
        End If
    End If
End Sub




Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Modo <> 6 Then 'Pasar de Pedido a Albaran
        If Not (Index = 0 And KeyCode = 38) Then KEYdown KeyCode
    Else 'Modo lineas
        Select Case KeyCode
            Case 38 'Desplazamieto Fecha Hacia Arriba
                If DataGrid1.Row > 0 Then
                    DataGrid1.Row = DataGrid1.Row - 1
                    CargaTxtAuxServidas True, True
                Else
                    PonerFoco txtAux(3)
                End If
                txtAux(3).Text = Format(Data2.Recordset!recibida, FormatoImporte)
                ConseguirFoco txtAux(3), Modo
                
            Case 40 'Desplazamiento Flecha Hacia Abajo
'                If DataGrid1.Row < Data2.Recordset.RecordCount - 1 Then
'                    DataGrid1.Row = DataGrid1.Row + 1
'                    CargaTxtAuxServidas True, True
'                Else
'                    PonerFocoBtn Me.cmdAceptar
'                End If
'                txtAux(3).Text = Format(Data2.Recordset!recibida, FormatoImporte)
'                ConseguirFoco txtAux(3), Modo
                
                PonerServidas
        End Select
    End If
End Sub



Private Sub txtAux_LostFocus(Index As Integer)
Dim devuelve As String
'Dim vPrecio As CPreciosCom
Dim TipoDto As Byte
Dim b As Boolean

    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    Select Case Index
        Case 0 'Cod ALMACEN
            'Comprobar que existe el almacen
            devuelve = PonerAlmacen(txtAux(Index).Text)
            txtAux(Index).Text = devuelve
            If devuelve = "" Then PonerFoco txtAux(Index)

        Case 1 'Cod. ARTICULO
            If txtAux(1).Text = "" Then
                txtAux(2).Text = ""
                Exit Sub
            End If
            
            If txtAux(0).Text = "" Then
                MsgBox "Debe seleccionar un almacen.", vbInformation
                PonerFoco txtAux(0)
                Exit Sub
            End If
            
            If PonerArticulo(txtAux(1), txtAux(2), txtAux(0).Text, CodTipoMov, ModificaLineas) Then
                CargarDatosArticulo txtAux(1).Text
                
                b = (Me.ActiveControl.Name = "txtAux")
                If b Then b = (Me.ActiveControl.Index = 0)
                
                If Not b Then
'                    If txtAux(2).Locked Then PonerFoco txtAux(3)
                Else
                    PonerFoco txtAux(0)
                End If
            Else
                PonerFoco txtAux(Index)
            End If
            
            
'            If PonerArticulo(txtAux(1), txtAux(2), txtAux(0).Text, CodTipoMov) Then
'                If txtAux(2).Locked Then PonerFoco txtAux(3)
'                'Si es articulo de varios podemos modificar la descripción del articulo, sino bloqueamos.
''                If Not EsArticuloVarios(txtAux(Index).Text) Then
''                    BloquearTxt txtAux(2), True
''                Else
''                    BloquearTxt txtAux(2), False
''                    PonerFoco txtAux(2)
''                End If
'            Else
'                PonerFoco txtAux(Index)
'            End If
            
        Case 2 'Desc. Articulo
            If txtAux(Index).Locked = False Then txtAux(Index).Text = UCase(txtAux(Index).Text)
            
        Case 3 'CANTIDAD
            If PonerFormatoDecimal(txtAux(Index), 1) Then  'Tipo 1: Decimal(12,2)
                'Comprobar si hay suficiente stock
                If (Modo = 5) And (ModificaLineas = 1 Or (ModificaLineas = 2 And txtAux(4).Text = "")) Then 'Modo Insertar en Mto Lineas
                    'Obtener el precio correspondiente y los descuentos
                    ObtenerPrecioCompra
                    
'                    Set vPrecio = New CPreciosCom
'                    If vPrecio.Leer(txtAux(1).Text, Text1(4).Text) Then
'                        If vPrecio.ComprobarCantidad(CInt(txtAux(3).Text)) Then
'                            txtAux(4).Text = vPrecio.ObtenerPrecio(Text1(1).Text)
'                            PonerFormatoDecimal txtAux(4), 2
'                            txtAux(5).Text = vPrecio.Descuento1
'                            PonerFormatoDecimal txtAux(5), 4
'                            txtAux(6).Text = vPrecio.Descuento2
'                            PonerFormatoDecimal txtAux(6), 4
'                        Else
'                            PonerFoco txtAux(Index)
'                        End If
'                    End If
'                    Set vPrecio = Nothing
                End If
            End If
            
        Case 4 'Precio
            PonerFormatoDecimal txtAux(Index), 7 'Tipo 2: Decimal(10,4)
        Case 5, 6 'Descuentos
            PonerFormatoDecimal txtAux(Index), 4 'Tipo 4: Decimal(4,2)
        Case 7 'Importe Linea
            PonerFormatoDecimal txtAux(Index), 1 'Tipo 3: Decimal(12,2)
    End Select
    
    If Modo = 5 Then
         If (Index = 3 Or Index = 4 Or Index = 5 Or Index = 6) Then 'Cant., Precio, Dto1, Dto2
'            If Trim(TxtAux(3).Text) = "" Or Trim(TxtAux(4).Text) = "" Then Exit Sub
'            If Trim(TxtAux(6).Text) = "" Or Trim(TxtAux(7).Text) = "" Then Exit Sub
            If txtAux(1).Text = "" Then Exit Sub
            TipoDto = DevuelveDesdeBDNew(cAgro, "proveedor", "tipodtos", "codprove", text1(4).Text, "N")
            txtAux(7).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(5).Text, txtAux(6).Text, TipoDto, 0)
            PonerFormatoDecimal txtAux(7), 1
        End If
    End If
End Sub



Private Sub ObtenerPrecioCompra()
Dim vPrecio As CPreciosCom
Dim Cad As String

    On Error GoTo EPrecios
    
    Set vPrecio = New CPreciosCom
    If vPrecio.Leer(txtAux(1).Text, text1(4).Text) Then
        If vPrecio.ComprobarCantidad(CInt(txtAux(3).Text)) Then
            txtAux(4).Text = vPrecio.ObtenerPrecio(text1(1).Text)
'            PonerFormatoDecimal txtAux(4), 2
            txtAux(5).Text = vPrecio.Descuento1
'            PonerFormatoDecimal txtAux(5), 4
            txtAux(6).Text = vPrecio.Descuento2
'            PonerFormatoDecimal txtAux(6), 4
        Else
            PonerFoco txtAux(3)
            Exit Sub
        End If
    Else
        'Obtener el ult. precio de compra de ese articulo (sartic)
        Cad = DevuelveDesdeBDNew(cAgro, "sartic", "preciouc", "codartic", txtAux(1).Text, "T")
        If Cad <> "" Then
            txtAux(4).Text = Cad
            txtAux(5).Text = "0"
            txtAux(6).Text = "0"
        End If
    End If
    PonerFormatoDecimal txtAux(4), 7
    PonerFormatoDecimal txtAux(5), 4
    PonerFormatoDecimal txtAux(6), 4
    
    Set vPrecio = Nothing
    
EPrecios:
    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub BotonMtoLineas(numTab As Integer, Cad As String)
        Me.SSTab1.Tab = numTab
        TituloLinea = Cad
        ModificaLineas = 0
        PonerModo 5
'--monica:rollo toolbar
'        PonerBotonCabecera True
End Sub


Private Function Eliminar() As Boolean
Dim b As Boolean
Dim vWhere As String
On Error GoTo FinEliminar

        conn.BeginTrans
         vWhere = ObtenerWhereCP(False)

'        If opt = 1 Then 'ELIMINAR
'            b = EliminarPedido(Data1.Recordset!numpedpr)
'        Else 'Pasar al HISTORICO
            b = ActualizarElTraspaso("", vWhere, CodTipoMov, CadenaSQL)
'        End If
        
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Pedido"
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


Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next
    CargaGrid DataGrid1, Data2, False
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
'Despues de hacer refresh del Data, volver a situar el Data en el registro que estaba
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = ObtenerWhereCP(False)
         vWhere = Replace(vWhere, NombreTabla & ".", "")
         If SituarData(Data1, vWhere, Indicador) Then
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


Private Function ObtenerWhereCP(conW As Boolean) As String
'Obtiene la where de la Clave Primaria de la tabla de Cabecera: scaped
Dim Sql As String
On Error Resume Next
    Sql = ""
    If conW Then Sql = " WHERE "
    Sql = Sql & NombreTabla & ".numpedpr= " & Val(text1(0).Text)
    If EsHistorico Then Sql = Sql & " AND " & NomTablaLineas & ".fecpedpr=" & DBSet(text1(1).Text, "F")
    ObtenerWhereCP = Sql
End Function


Private Function MontaSQLCarga(enlaza As Boolean, Optional conServidas As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data2
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim Sql As String
    
    Sql = "SELECT numpedpr, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, "
    If conServidas Then Sql = Sql & "recibida, "
'    SQL = SQL & "precioar, origpre, dtoline1, dtoline2,importel "
    Sql = Sql & "precioar, dtoline1, dtoline2,importel "
    Sql = Sql & " FROM " & NomTablaLineas
    If enlaza Then
        Sql = Sql & " " & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
        If EsHistorico Then Sql = Sql & " and fecpedpr='" & Format(text1(1).Text, FormatoFecha) & "'"
    Else
        Sql = Sql & " WHERE numpedpr = -1"
    End If
    Sql = Sql & " Order by numpedpr, numlinea"
    MontaSQLCarga = Sql
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean, bAux As Boolean
Dim i As Integer


        b = (Modo = 2) And Pedido = "" '--monica: rollo toolbar--> Or (Modo = 5 And ModificaLineas = 0)
        'Insertar
        Toolbar1.Buttons(1).Enabled = (b Or Modo = 0) And Not EsHistorico
        Me.mnNuevo.Enabled = (b Or Modo = 0) And Not EsHistorico
        'Modificar
        Toolbar1.Buttons(2).Enabled = b And Not EsHistorico
        Me.mnModificar.Enabled = b And Not EsHistorico
        'eliminar
        Toolbar1.Buttons(3).Enabled = b And Not EsHistorico
        Me.mnEliminar.Enabled = b And Not EsHistorico
            
        b = (Modo = 2) And (Pedido = "") And Not EsHistorico
'--monica:rollo toolbar
'        'Mantenimiento lineas
'        Toolbar1.Buttons(10).Enabled = b
'        Me.mnLineas.Enabled = b
        'Generar Albaran desde Pedido
        Toolbar5.Buttons(1).Enabled = b
        Me.mnGenAlbaran.Enabled = b
        
        b = ((Modo >= 3) Or Modo = 1) And (Pedido = "")
        'Buscar
        Toolbar1.Buttons(5).Enabled = Not b
        Me.mnBuscar.Enabled = Not b
        'Ver Todos
        Toolbar1.Buttons(6).Enabled = Not b
        Me.mnVerTodos.Enabled = Not b



    '++monica: rollo toolbar
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not EsHistorico And Pedido = ""
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.Data2.Recordset.RecordCount > 0)
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i


End Sub


Private Function CargarDatosDirec(CodDirec As String, indice As Byte) As Boolean
'Direcciones Propias
Dim Rs As ADODB.Recordset
Dim devuelve As String
Dim b As Boolean
On Error GoTo ECargarProve

    b = False
    If CodDirec <> "" Then
        devuelve = "Select nomdirec, domdirec, codpobla, pobdirec, prodirec "
        devuelve = devuelve & " FROM sdirpr Where coddirec=" & Val(CodDirec)
        
        Set Rs = New ADODB.Recordset
        Rs.Open devuelve, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not Rs.EOF Then
            text1(indice).Text = Format(CodDirec, "000")
            Text2(indice).Text = Rs.Fields!nomdirec 'Nom Direccion
            If indice = 2 Then
                indice = 21
            Else
                indice = 17
            End If
            Text2(indice).Text = Rs.Fields!domdirec 'Domicilio
            Text2(indice + 1).Text = Rs.Fields!codpobla
            Text2(indice + 2).Text = Rs.Fields!pobdirec
            Text2(indice + 3).Text = Rs.Fields!prodirec
            b = True
        Else
            MsgBox "No existe la dirección: " & text1(indice).Text, vbInformation
            LimpiarDatosDirec (indice)
        End If
        Rs.Close
        Set Rs = Nothing
    Else
        LimpiarDatosDirec (indice)
        b = True
    End If
    
    CargarDatosDirec = b
    
ECargarProve:
    If Err.Number <> 0 Then CargarDatosDirec = False
End Function


Private Sub LimpiarDatosDirec(indice As Byte)
    Text2(indice).Text = ""
    If indice = 2 Then
        indice = 21
    Else
        indice = 17
    End If
    Text2(indice).Text = "" 'Domicilio
    Text2(indice + 1).Text = "" 'cpostal
    Text2(indice + 2).Text = "" 'poblacion
    Text2(indice + 3).Text = "" 'provincia
End Sub


Private Function InsertarPedido(vSQL As String, vTipoMov As CTiposMov) As Boolean
'Insertar la Cabecera de un Pedido, tabla: scaped
Dim MenError As String
Dim bol As Boolean, Existe As Boolean
Dim cambiaSQL As Boolean
Dim devuelve As String
On Error GoTo EInsertarOferta
    
    bol = True
    
    cambiaSQL = False
    'Comprobar si mientras tanto se incremento el contador de Pedidos
    'para ello vemos si existe un Pedido con ese contador y si existe lo incrementamos
    Do
        devuelve = DevuelveDesdeBDNew(cAgro, NombreTabla, "numpedpr", "numpedpr", text1(0).Text, "N")
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
    MenError = "Error al insertar en la tabla Cabecera de Pedidos (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
    
    'Actualizar los datos del proveedor si es de varios
    If EsDeVarios Then
        'Si es cliente de varios actualizar datos cliente en tabla:sclvar
        MenError = "Modificando datos proveedor varios."
        bol = ActualizarProveVarios(text1(4).Text, text1(6).Text)
    End If
    
    MenError = "Error al actualizar el contador del Pedido."
    vTipoMov.IncrementarContador (CodTipoMov)

EInsertarOferta:
        If Err.Number <> 0 Then
            MenError = "Insertando Pedido." & vbCrLf & "----------------------------" & vbCrLf & MenError
            MuestraError Err.Number, MenError, Err.Description
            bol = False
        End If
        If bol Then
            conn.CommitTrans
            InsertarPedido = True
        Else
            conn.RollbackTrans
            InsertarPedido = False
        End If
End Function


Private Sub LimpiarDatosProve()
'Limpia los campos del Form con datos del Proveedor
Dim i As Byte

    For i = 4 To 14
        text1(i).Text = ""
    Next i
End Sub
    


Private Function InicializarCStockAlbar(ByRef vCStock As CStock, TipoM As String, Optional NumLinea As String, Optional ByRef Rs As ADODB.Recordset) As Boolean
'Para comprobar stock al pasar de Pedido a Albaran de Venta
Dim TipoDto As Byte
On Error Resume Next

    vCStock.tipoMov = TipoM
    vCStock.DetaMov = "ALC"
    vCStock.Trabajador = CInt(text1(4).Text) 'En codigope ponemos el Proveedor
    vCStock.Documento = text1(0).Text
    vCStock.codArtic = Rs!codArtic
    vCStock.codAlmac = CInt(Rs!codAlmac)
    
    If AlbCompleto Then
        vCStock.Cantidad = CSng(Rs!Cantidad)
        If Rs.Fields.Count > 3 Then 'Si no se selecciona el campo importe de la tabla es que solo vamos a comprobar stock y no se necesita
            vCStock.Importe = CCur(Rs!ImporteL)
        End If
    Else
        vCStock.Cantidad = CSng(Rs!recibida)
        'Si se va a Insertar en alguna linea obtener el importe
        'Si solo vamos a comprobar stock no hace falta el importe
        If Rs.Fields.Count > 4 Then
            TipoDto = DevuelveDesdeBDNew(cAgro, "proveedor", "tipodtos", "codprove", Me.Data1.Recordset!codProve, "N")
            vCStock.Importe = CCur(CalcularImporte(Rs!recibida, Rs!precioar, Rs!dtoline1, Rs!dtoline2, TipoDto, 0))
        End If
    End If
    
    vCStock.LineaDocu = CInt(ComprobarCero(NumLinea))
    
    If Err.Number <> 0 Then
        MsgBox "No se han podido inicializar la clase para actualizar Stock", vbExclamation
        InicializarCStockAlbar = False
    Else
        InicializarCStockAlbar = True
    End If
End Function


Private Function PasarPedidoAAlbaran(NumAlb As String, FechaAlb As String) As Boolean
'OUT -> numalb: Devuelve el Nº de albaran asignado al pedido
Dim bol As Boolean
Dim MenError As String
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim vWhere As String
Dim cProve As CProveedor

    On Error GoTo EGenPedido

    bol = False
            
    
    'Aqui empieza transaccion
    conn.BeginTrans
    
    'Insertar en tablas de Albaranes Proveedor el Pedido  (scaalp, slialp)
    bol = InsertarAlbaran(MenError, NumAlb)
    
    
    
    
    
    
    'Para cada linea del pedido:
    ' Actualizar precio medio ponderado del articulo
    ' Actualizar precio y fecha ultima compra del articulo
    ' Actualizar Stock en salmac (entrada de stock), e introducir movimiento en smoval
    If bol Then
        MenError = "Actualizando Stocks"
        bol = InsertarMovStock(NumAlb, FechaAlb)
    End If
    
    If bol Then
        'Actualizar la ult.fecha de compra del Proveedor
        MenError = "Actualizando ultima fecha compra en Proveedor."
        Set cProve = New CProveedor
        bol = cProve.ActualizaFechaUltCompra(text1(4).Text, FechaAlb)
        Set cProve = Nothing
        
'        If bol Then
'            'Actualizar ult. fecha de compra y el precio ult compra de los articulos del Albaran
'            MenError = "Actualizando ultima fecha compra en Artículos."
'            SQL = "numalbar=" & DBSet(NumAlb, "T") & " and fechaalb=" & DBSet(FechaAlb, "F") & " and slialp.codprove=" & Text1(4).Text
'            bol = ActualizarUltFechaCom(SQL)
'        End If
    End If
    
    
    If bol Then
        If AlbCompleto Then  'Si se inserta Albaran
            'Borrar el Pedido de las tablas de Pedidos (scaped, sliped)
            MenError = "Eliminando cabecera y lineas del Pedido."
            bol = EliminarPedido(CLng(text1(0).Text))
        Else
            'Actualizar la cantidad=cantidad-recibida y recibida= 0 en slippr
            bol = ActualizarPedido()
            'Marcar Resto de pedido: restoped=1
            If bol Then bol = ActualizarCabPedido
        End If
    End If
    
    
'--monica
'    If bol Then
'        'si se ha generado correctamente el ALBARAN ver si hay alguna línea que tiene
'        'el artículo con control de nº de lote y pedir los nº de lotes.
'        ComprobarNumLotesLineas NumAlb, FechaAlb
'
'    End If
    
    
    
    
    If bol Then
        'Se ha generado correctamente el ALBARAN y vemos si tiene Nº Series
'        FechaAlb = RecuperaValor(CadenaSQL, 3)
        'Comprobar si Hay Nº SERIE en compras y Mostrar
        'ventana para pedir los Nº Serie de la cantidad introducida si lo requiere algun articulo
'--monica
'        ComprobarNSeriesLineas NumAlb, FechaAlb
        
        
        If Not AlbCompleto Then
            'Eliminar las filas del pedido que se servieron completas (slippr)
            MenError = "Eliminando lineas pedidido servidas completas."
            Sql = "DELETE FROM " & NomTablaLineas & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND cantidad=0"
            conn.Execute Sql
            
            'Comprobar que si no quedan lineas en el pedido se elimine la cabecera del pedido
            MenError = "Eliminando cabecera del pedido."
            Sql = "select codalmac, codartic FROM " & NomTablaLineas & " WHERE numpedpr=" & text1(0).Text
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Rs.EOF Then 'No hay lineas de pedido --> Eliminar la cabecera
                Sql = "DELETE FROM " & NombreTabla & " WHERE numpedpr=" & text1(0).Text
                conn.Execute Sql
            End If
            Rs.Close
            Set Rs = Nothing
        End If
        bol = True
    End If
    
    
EGenPedido:
    If Err.Number <> 0 Then
'        MenError = "Pasando Pedido a Albaran." & vbCrLf & "----------------------------" & vbCrLf & MenError
'        MuestraError Err.Number, MenError, Err.Description
        bol = False
    End If
    If bol Then
        conn.CommitTrans
        PasarPedidoAAlbaran = True
    Else
        conn.RollbackTrans
        PasarPedidoAAlbaran = False
        MenError = "Pasando Pedido a Albaran." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
    End If
End Function


Private Function InsertarAlbaran(MenError As String, NumAlb As String) As Boolean
'Devuelve el mensaje de error si se produce
Dim bol As Boolean
Dim vSQL As String
Dim FechaAlb As String
Dim TrabAlb As String

    On Error GoTo EInsertarAlbaran
    
    bol = False
    InsertarAlbaran = bol
    
    NumAlb = RecuperaValor(CadenaSQL, 1)
    FechaAlb = RecuperaValor(CadenaSQL, 2)
    
    vSQL = "INSERT INTO scaalp (numalbar, fechaalb, codprove, nomprove, domprove, codpobla, pobprove, proprove, nifprove, telprove, codforpa, dtoppago, dtognral, observa1, observa2, observa3, observa4, observa5, numpedpr, fecpedpr)"
    vSQL = vSQL & " SELECT " & DBSet(NumAlb, "T") & " as numalbar, " & DBSet(FechaAlb, "F") & " as fechaalb, "
    vSQL = vSQL & "codprove, nomprove, domprove, codpobla, pobprove, proprove, nifprove, telprove, codforpa, "
    vSQL = vSQL & "dtoppago, dtognral, observa1, observa2, observa3, observa4, observa5, numpedpr, fecpedpr "
    vSQL = vSQL & " FROM " & NombreTabla & " WHERE numpedpr=" & text1(0).Text

    'Insertar Cabecera
    MenError = "Error al insertar en la tabla Cabecera de Albaranes Proveedor (scaalp)."
    conn.Execute vSQL, , adCmdText
    
    'Insertar Lineas Albaran desde Pedido
    MenError = "Error al insertar en la tabla Lineas de Albaran (slialp)."
    If Not InsertarLineasAlbaran(NumAlb, FechaAlb) Then Exit Function
    
    bol = True
    
EInsertarAlbaran:
        If Err.Number <> 0 Then
            bol = False
            MenError = MenError & vbCrLf & Err.Description
        End If
        If bol Then
            InsertarAlbaran = True
        Else
            InsertarAlbaran = False
        End If
End Function


Private Function InsertarLineasAlbaran(NumAlb As String, FechaAlb As String) As Boolean
'Inserta en la tabla de lineas de albaran (slialb)
'IN -> TipoM, numAlb
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim ImpLinea As String
Dim TipoDto As Byte
On Error GoTo EInsertarLinAlb

    If AlbCompleto Then
        'Insertar en la tabla de Albaran, los registros seleccionados de la tabla de Pedidos
        Sql = ""
        Sql = "SELECT " & DBSet(NumAlb, "T") & " as numalbar, " & DBSet(FechaAlb, "F") & " as fechaalb, " & Val(text1(4).Text) & " as codprove, numlinea, codartic, codalmac, nomartic, ampliaci, "
        Sql = Sql & "cantidad, precioar, dtoline1, dtoline2, importel "
        Sql = Sql & " FROM " & NomTablaLineas & " WHERE numpedpr=" & Val(text1(0).Text)
        Sql = "INSERT INTO slialp (numalbar, fechaalb, codprove, numlinea, codartic, codalmac, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel) " & Sql
        conn.Execute Sql, , adCmdText
    Else
        Sql = "select * from " & NomTablaLineas
        Sql = Sql & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs.EOF 'Para cada linea de pedido insertar una de albaran si recibidas >0
            If Rs!recibida > 0 Then
                TipoDto = DevuelveDesdeBDNew(cAgro, "proveedor", "tipodtos", "codprove", text1(4).Text, "N")
                ImpLinea = CalcularImporte(Rs!recibida, Rs!precioar, Rs!dtoline1, Rs!dtoline2, TipoDto, 0)
                Sql = "INSERT INTO slialp (numalbar, fechaalb, codprove, numlinea,codartic, codalmac, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel) "
                Sql = Sql & " VALUES(" & DBSet(NumAlb, "T") & ", " & DBSet(FechaAlb, "F") & ", " & Val(text1(4).Text) & ", " & Rs!NumLinea & ", "
                Sql = Sql & DBSet(Rs!codArtic, "T") & "," & Rs!codAlmac & ", " & DBSet(Rs!NomArtic, "T") & ", " & DBSet(Rs!ampliaci, "T") & ", "
                Sql = Sql & DBSet(Rs!recibida, "N") & ", " & DBSet(Rs!precioar, "N") & ", " & DBSet(Rs!dtoline1, "N") & ", " & DBSet(Rs!dtoline2, "N") & ", "
                Sql = Sql & DBSet(ImpLinea, "N") & ")"
                conn.Execute Sql, , adCmdText
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
    End If
    
EInsertarLinAlb:
    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        InsertarLineasAlbaran = False
        MuestraError Err.Number, "Insertar lineas albaran.", Err.Description
    Else
        InsertarLineasAlbaran = True
    End If
End Function


Private Function EliminarPedido(numPed As Long) As Boolean
'Eliminar las lineas y la Cabecera de un Pedido. Tablas: scaped, sliped
Dim Sql As String
On Error GoTo EEliminarPed

     Sql = " WHERE  numpedpr=" & numPed

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


Private Function ActualizarPedido() As Boolean
'Actualiza la tabla de lineas de pedido (sliped)
'cantidad=cantidad-servidas y servidas=0
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim ImpLinea As String
Dim TipoDto As Byte

    On Error GoTo EActPedido

    Sql = "select numlinea, codalmac, codartic, cantidad, recibida, precioar, dtoline1, dtoline2 from " & NomTablaLineas
    Sql = Sql & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF 'Para cada linea
        TipoDto = DevuelveDesdeBDNew(cAgro, "proveedor", "tipodtos", "codprove", text1(4).Text, "N")
        ImpLinea = CalcularImporte(Rs!Cantidad - Rs!recibida, Rs!precioar, Rs!dtoline1, Rs!dtoline2, TipoDto, 0)
        Sql = "UPDATE " & NomTablaLineas & " SET cantidad=cantidad-recibida, recibida=0, importel=" & DBSet(ImpLinea, "N")
'        SQL = SQL & " WHERE codalmac=" & RS!codAlmac & " AND codartic='" & RS!codArtic & "'"
        Sql = Sql & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
        Sql = Sql & " AND numlinea=" & Rs!NumLinea
        conn.Execute Sql
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
EActPedido:
    If Err.Number <> 0 Then
        ActualizarPedido = False
    Else
        ActualizarPedido = True
    End If
End Function


Private Function ActualizarCabPedido() As Boolean
Dim Sql As String
On Error Resume Next

    Sql = "UPDATE " & NombreTabla & " SET restoped=1 " & ObtenerWhereCP(True)
    conn.Execute Sql
    If Err.Number <> 0 Then
        ActualizarCabPedido = False
    Else
        ActualizarCabPedido = True
    End If
End Function


Private Function InsertarMovStock(NumAlb As String, FechaAlb As String) As Boolean
Dim vCStock As CStock
Dim b As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim cArt As CArticulo

    On Error Resume Next

    InsertarMovStock = False
    
    Set vCStock = New CStock
    b = True
    
    Sql = Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
    Sql = "select * from " & NomTablaLineas & Sql
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    vCStock.Fechamov = FechaAlb
    
    'para cada linea del Pedido Insertar en smoval y Actualizar Stock en salmac
    While (Not Rs.EOF) And b
        If InicializarCStockAlbar(vCStock, "E", CStr(Rs!NumLinea), Rs) Then
            vCStock.Documento = NumAlb
            If vCStock.Cantidad <> 0 Then
                '==== Laura 22/09/2006
                '-- antes de actualizar el stock calculamos el precio medio ponderado del articulo
                Set cArt = New CArticulo
                If cArt.LeerDatos(vCStock.codArtic) Then
                    'Laura 19/12/2006: Calcular precio_med_pond con el precio con los descuentos,e.d. importe/cantidad
                    'If Not cArt.ActualizarPrecioMedPond(CCur(vCStock.Cantidad), CCur(RS!precioar)) Then b = False
                    If Not cArt.ActualizarPrecioMedPond(CCur(vCStock.Cantidad), Round2(CCur(vCStock.Importe) / CCur(vCStock.Cantidad), 4)) Then b = False
                        
                    '--actualizar fecha y precio ultima compra del articulo
                    'Laura 19/12/2006: actualizar precio_ult_compra con el precio con los descuentos,e.d. importe/cantidad
                    'If Not cArt.ActualizarUltFechaCompra(vCStock.Fechamov, CStr(RS!precioar)) Then b = False
                    'monica : 17/06/2009 añadida la condicion de que la cantidad sea positivo si se actualiza el UC
                    If CCur(vCStock.Cantidad) >= 0 Then
                        If Not cArt.ActualizarUltFechaCompra(vCStock.Fechamov, Round2(CCur(vCStock.Importe) / CCur(vCStock.Cantidad), 4)) Then b = False
                    End If
                End If
                Set cArt = Nothing
                '====
            
            
                'en actualizar stock comprobamos si el articulo tiene control de stock
                b = vCStock.ActualizarStock
            End If
        Else
            b = False
        End If
        Rs.MoveNext
    Wend
    Set vCStock = Nothing
    Rs.Close
    Set Rs = Nothing
    
    InsertarMovStock = b
    
End Function


Private Sub ImprimirAlbaran(Opcion As Integer, NumAlbar As String)
End Sub


Private Function ActualizarServidas() As Boolean
Dim Sql As String
On Error Resume Next

    Sql = "UPDATE " & NomTablaLineas & " SET recibida= " & DBSet(txtAux(3).Text, "N")
    Sql = Sql & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea=" & Data2.Recordset!NumLinea
    conn.Execute Sql
    
    If Err.Number <> 0 Then
        ActualizarServidas = False
    Else
        ActualizarServidas = True
    End If
End Function


Private Sub PonerServidas()
Dim NumFila As Integer
Dim cadMen As String

'    NumFila = DataGrid1.Row
    NumFila = Data2.Recordset.AbsolutePosition
    If PonerFormatoDecimal(txtAux(3), 1) Then  'Tipo 1: Decimal(12,2)
        If CCur(txtAux(3).Text) > Data2.Recordset!Cantidad Then
            cadMen = "La cantidad a Recibir no puede ser superior a la del Pedido."
            MsgBox cadMen, vbExclamation
            PonerFoco txtAux(3)
            Exit Sub
        End If
    End If
    ActualizarServidas
    CargaGrid2 DataGrid1, Data2, True
'    DataGrid1.Row = NumFila
    SituarDataPosicion Data2, CLng(NumFila), ""
    MoverSigRegistro
End Sub




Private Sub MoverSigRegistro()
    On Error GoTo EMover
    
    If Data2.Recordset.EOF Then Exit Sub
    If Data2.Recordset.AbsolutePosition <= Data2.Recordset.RecordCount - 1 Then
        DataGrid1.Row = DataGrid1.Row + 1
        CargaTxtAuxServidas True, True
    Else
        PonerFocoBtn Me.cmdAceptar
    End If
    txtAux(3).Text = Format(Data2.Recordset!recibida, FormatoImporte)
    ConseguirFocoLin txtAux(3)
EMover:
    If Err.Number <> 0 Then MuestraError Err.Description, "Mover registro.", Err.Description
End Sub



Private Sub GenerarAlbaran()
Dim numPed As Long 'Nº Pedido
Dim NumAlb As String 'Nº Albaran
Dim FechaAlb As String 'Fecha del Albaran
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim b As Boolean

    NumRegElim = Data1.Recordset.AbsolutePosition
    numPed = Data1.Recordset!numpedpr
    
    'pedir por pantalla:  Nº albaran y fecha albaran
    Set frmList = New frmListadoOfer
    frmList.Opcionlistado = 57
    CadenaSQL = ""
    frmList.Show vbModal
    Set frmList = Nothing
    
    If CadenaSQL = "" Then Exit Sub
    FechaAlb = RecuperaValor(CadenaSQL, 2)
    
    
    'Antes de pasar el pedido al albaran nos guardamos los articulos cuyo precio_compra
    'se han modificado para preguntar despues si se quiere actualizar precios_venta
    'hay q guardarlo antes de pasar pedido a albaran ya q aqui se actualiza el precio_ult_compra
    '-- Laura 19/12/2006: calcular precio_med_pond con el precio aplicados los descuentos, ed. importe/cantidad
    'SQL = "SELECT slippr.codartic,sartic.nomartic,slippr.precioar,sartic.preciouc,sum(cantidad) "
    Sql = "SELECT slippr.codartic,sartic.nomartic,round(slippr.importel/slippr.cantidad,4) as precioar,sartic.preciouc,sum(cantidad) "
    Sql = Sql & " FROM slippr INNER JOIN sartic ON slippr.codartic=sartic.codartic "
    'SQL = SQL & " WHERE numpedpr=" & numPed & " and (slippr.precioar<>sartic.preciouc)"
    Sql = Sql & " WHERE numpedpr=" & numPed & " and (round(slippr.importel/slippr.cantidad,4)<>sartic.preciouc)"
    'seleccionar solo de las que se vayan a recibir
    If Not AlbCompleto Then Sql = Sql & " and slippr.recibida>0 "
    Sql = Sql & " group by slippr.codartic,slippr.precioar,sartic.preciouc "
    Sql = Sql & " Having Sum(Cantidad) > 0"
    b = ObtenerRSprecios(Rs, Sql)
    
    
    
    If PasarPedidoAAlbaran(NumAlb, FechaAlb) Then
        MsgBox "Se ha generado correctamente el Albaran: " & NumAlb, vbInformation
                
'        FechaAlb = RecuperaValor(CadenaSQL, 3)
'        'Comprobar si Hay Nº SERIE en compras y Mostrar
'        'ventana para pedir los Nº Serie de la cantidad introducida si lo requiere algun articulo
'        ComprobarNSeriesLineas NumAlb, FechaAlb

        PonerModo 2
        
        
        'comprobar si hay lineas de artículos cuyo precio_ultima_compra
        'se ha modificado y preguntar si que quieren actualizar los precio_venta
        '--------------------------------------------------------
        If b Then
            While Not Rs.EOF
                Sql = "Se ha modificado el precio última compra del artículo:" & vbCrLf
                Sql = Sql & Rs!codArtic & ":  " & Rs!NomArtic & vbCrLf
                Sql = Sql & vbCrLf & "¿Desea actualizar los precios de venta?"
                If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                    'Comprobar que el artículo tiene margen comercial
'--monica
'                    If ArticuloTieneMargen(RS!codArtic) Then
'                        'Aplicar margen comercial a los precios
'                        'Modificar precios de venta en articulo y tarifas
'                        frmComActPrecios.parCodArtic = RS!codArtic
'                        frmComActPrecios.parNomArtic = RS!NomArtic
'                        frmComActPrecios.Show vbModal
'                    End If
                End If
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
        End If
        
        
        
        
        If AlbCompleto Then
            'Se habra eliminado el pedido de (scaped, sliped)
            PosicionarDataTrasEliminar
        Else
            Sql = DevuelveDesdeBDNew(cAgro, "scappr", "numpedpr", "numpedpr", text1(0).Text, "N")
            If Sql = "" Then 'Ya no existe le pedido lo hemos eliminado
                PosicionarDataTrasEliminar
            Else
                PosicionarData
                PonerCampos
                CargaGrid DataGrid1, Data2, True, False
            End If
            CargaTxtAuxServidas False, False
        
            'Eliminar las filas del pedido que se servieron completas (slippr)
'            SQL = "DELETE FROM " & NomTablaLineas & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND cantidad=0"
'            Conn.Execute SQL
            
            'Comprobar que si no quedan lineas en el pedido se elimine la cabecera del pedido
'            SQL = "select codalmac, codartic FROM " & NomTablaLineas & " WHERE numpedpr=" & numPed
'            Set RS = New ADODB.Recordset
'            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'            If RS.EOF Then 'No hay lineas de pedido --> Eliminar la cabecera
'                SQL = "DELETE FROM " & NombreTabla & " WHERE numpedcl=" & numPed
'                Conn.Execute SQL
'                PosicionarDataTrasEliminar
'            Else 'Quedan lineas en el pedido --> Actualizar las lineas
'                PosicionarData
'                PonerCampos
'                CargaGrid DataGrid1, Data2, True, False
'            End If
'            RS.Close
'            Set RS = Nothing
'            CargaTxtAuxServidas False, False
        End If
        Screen.MousePointer = vbDefault
        
'        Imprimer albaran si se solicitó
'        If ImprimeAlb Then
'            ImprimirAlbaran 45, NumAlb
'        End If
    Else 'Si no se ha pasado el Pedido a Albaran
        
    End If
End Sub


Private Sub InicializarServidas()
'Pone el campo servidas a 0 en la tabla lineas de pedido (sliped)
Dim Sql As String
    On Error Resume Next
    Sql = "UPDATE " & NomTablaLineas & " SET recibida= 0 "
    Sql = Sql & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
    conn.Execute Sql
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub ComprobarNumLotesLineas(NumAlb As String, FechaAlb As String)
'Al pasar de PEDIDO a ALBARAN
'control de Nº Lotes si hay algun articulo en las lineas de pedido que
'requiere Nº de lote en compras pedirlos
Dim Sql As String
Dim RSLineas As ADODB.Recordset
Dim cadwhere As String

    On Error GoTo ErrLotes

    cadwhere = " WHERE numalbar=" & DBSet(NumAlb, "T") & " AND "
    cadwhere = cadwhere & " fechaalb=" & DBSet(FechaAlb, "F") & " AND "
    cadwhere = cadwhere & " slialp.codprove=" & text1(4).Text

    'seleccionamos aquellas lineas del albaran insertado que tengan control de lote
    Sql = "SELECT slialp.* "
    Sql = Sql & " FROM (slialp INNER JOIN sartic ON slialp.codartic=sartic.codartic) "
    Sql = Sql & " LEFT OUTER JOIN scateg ON sartic.codcateg=scateg.codcateg "
    Sql = Sql & cadwhere
    Sql = Sql & " AND scateg.ctrlotes = 1"


    Set RSLineas = New ADODB.Recordset
    RSLineas.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Not RSLineas.EOF Then
        'Comprobar si NO Hay Nº SERIE en Compras y si no se realizo alli
        'Mostrar ahora ventana para pedir los Nº Serie de la cantidad introducida
'        Me.cmdAux(1).Tag = NumAlb
'        Me.cmdAux(0).Tag = FechaAlb
'--monica
'        PedirNLotes RSLineas
    
'        Set frmNLote = New frmAlmCargarNLote
'        frmNLote.parSQL = SQL
'        frmNLote.Show vbModal
'        Set frmNLote = Nothing

    End If
    
    RSLineas.Close
    Set RSLineas = Nothing
    Exit Sub

ErrLotes:
    MuestraError Err.Number, "Pedir Nº de lote.", Err.Description
End Sub



'--monica
'Private Sub ComprobarNSeriesLineas(NumAlb As String, FechaAlb As String)
''Al pasar de PEDIDO a ALBARAN
''control de Nº Series si hay algun articulo en las lineas de pedido que requiere Nº de serie
''y hay control de Nº de serie en compras pedirlos
'Dim SQL As String
'Dim RSLineas As ADODB.Recordset
'Dim cadwhere As String
'
'    If vParamAplic.NumSeries Then 'So control de Nº Series en COMPRAS
'        cadwhere = " WHERE numalbar=" & DBSet(NumAlb, "T") & " AND "
'        cadwhere = cadwhere & " fechaalb=" & DBSet(FechaAlb, "F") & " AND "
'        cadwhere = cadwhere & " slialp.codprove=" & Text1(4).Text
'
'        'Seleccionamos aquellas lineas de albaran que tienen Nº de Serie
'        SQL = "SELECT slialp.codartic, sum(cantidad) as cantidad, slialp.numlinea "
'        SQL = SQL & " FROM slialp INNER JOIN sartic on slialp.codartic=sartic.codartic "
'        SQL = SQL & cadwhere & " And nseriesn = 1 "
'        SQL = SQL & " GROUP BY codartic ORDER BY Codartic "
'
'        Set RSLineas = New ADODB.Recordset
'        RSLineas.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'        If Not RSLineas.EOF Then
'            'Comprobar si NO Hay Nº SERIE en Compras y si no se realizo alli
'            'Mostrar ahora ventana para pedir los Nº Serie de la cantidad introducida
'            Me.cmdAux(1).Tag = NumAlb
'            Me.cmdAux(0).Tag = FechaAlb
'            PedirNSeries RSLineas
'        End If
'        RSLineas.Close
'        Set RSLineas = Nothing
'    End If
'End Sub


'--monica
'Private Sub PedirNSeries(ByRef RS As ADODB.Recordset)
'On Error GoTo EPedirNSeries
'
'        'Visualizar en pantalla el Grid, y rellenar los Nº Serie
'        PedirNSeriesGnral RS, True
'
'        Set frmNSerie = New frmRepCargarNSerie
'        frmNSerie.DeVentas = False 'Se llama desde Alb. de Venta
'        frmNSerie.Show vbModal
'        Set frmNSerie = Nothing
'
'EPedirNSeries:
'    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
'End Sub


'--monica
'Private Sub PedirNLotes(ByRef RS As ADODB.Recordset)
'Dim cadSel As String
'
'    On Error GoTo EPedirNLotes
'
'    cadSel = "numalbar=" & DBSet(RS!numalbar, "T") & " AND fechaalb=" & DBSet(RS!FechaAlb, "F") & " AND codprove=" & DBSet(RS!codProve, "N")
'
'    'Visualizar en pantalla el Grid, y rellenar los Nº Serie
'    If Not PedirNLotesGnral(RS, True) Then
''             Visualizar en pantalla el Grid, y rellenar los Nº Serie
'        MsgBox "No se han podido mostrar todos los Artículos con Nº de Lote.", vbInformation
'    End If
'
'        Set frmNLote = New frmAlmCargarNLote
'        frmNLote.parSelSQL = cadSel
'        frmNLote.Show vbModal
'        Set frmNLote = Nothing
'
'
'     'Eliminar de la tabla temporal tmpnlotes los lotes introducidos
'    DescargarDatosTMPNumLotes "tmpnlotes", cadSel
'
'EPedirNLotes:
'    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
'End Sub


'--monica
'Private Function InsertarNSerie(numSerie As String, codArtic As String, numlinea As String) As Boolean
''Inserta o Actualiza en la tabla sserie, si al pasar Pedido -> Albaran
''existen lineas con control de Nº Serie
''Dim CadValues As String, cadValuesU As String
'Dim devuelve As String
'Dim numalbar As String
'Dim nSerie As CNumSerie
'Dim b As Boolean
'
'    On Error GoTo EInsertarNS
'
'    Set nSerie = New CNumSerie
'    nSerie.numSerie = numSerie
'    nSerie.Articulo = codArtic
'    nSerie.Proveedor = CInt(Text1(4).Text)
'    nSerie.NumAlbProve = Me.cmdAux(1).Tag
'    nSerie.fechacom = Me.cmdAux(0).Tag
'    nSerie.NumLinAlbPr = numlinea
'    'calculamos la fecha de fin garantia para el articulo comprado
'    nSerie.ObtenFechaFinGarantia codArtic, Me.cmdAux(0).Tag
'
'    'Comprobar si existe en la tabla sserie
'    numalbar = "numalbpr" 'Nº albaran de Compra
'    devuelve = DevuelveDesdeBDNew(conAri, "sserie", "numserie", "numserie", numSerie, "T", numalbar, "codartic", codArtic, "T")
'    If devuelve <> "" Then 'EXISTE en tabla sserie
'        If numalbar = "" Then
'            b = nSerie.ActualizarNumSerie(False)
'        End If
'    Else
'        b = nSerie.InsertarNumSerie
'    End If
'    Set nSerie = Nothing
'
'EInsertarNS:
'    If Err.Number <> 0 Then b = False
'    If Not b Then
'        InsertarNSerie = False
'    Else
'        InsertarNSerie = True
'    End If
'End Function



Private Sub PonerDatosProveedor(codProve As String, Optional nifProve As String)
'lee de la tabla de proveedores y pone los valores
Dim vProve As CProveedor
Dim Observaciones As String
    
    On Error GoTo EPonerDatos
    
    If codProve = "" Then
        LimpiarDatosProve
        Exit Sub
    End If

    Set vProve = New CProveedor
    'si se ha modificado el proveedor volver a cargar los datos
    If vProve.Existe(codProve) Then
        If vProve.LeerDatos(codProve) Then
           
            EsDeVarios = vProve.DeVarios
            BloquearDatosProve (EsDeVarios)
        
            If Modo = 4 And EsDeVarios Then 'Modificar
                'si no se ha modificado el proveedor no hacer nada
                If CLng(text1(4).Text) = CLng(Data1.Recordset!codProve) Then
                    Set vProve = Nothing
                    Exit Sub
                End If
            End If
        
            text1(4).Text = vProve.Codigo
            FormateaCampo text1(4)
            If (Modo = 3) Or (Modo = 4) Then
                text1(5).Text = vProve.Nombre  'Nom prove
                text1(8).Text = vProve.Domicilio
                text1(9).Text = vProve.CPostal
                text1(10).Text = vProve.Poblacion
                text1(11).Text = vProve.Provincia
                text1(6).Text = vProve.NIF
                text1(7).Text = DBLet(vProve.TfnoAdmon, "T")
            End If
            
            If Modo = 3 Then 'insertar
                text1(12).Text = vProve.ForPago
                Text2(12).Text = PonerNombreDeCod(text1(12), "forpago", "nomforpa")
                text1(13).Text = Format(vProve.DtoPPago, FormatoDescuento)
                text1(14).Text = Format(vProve.DtoGnral, FormatoDescuento)
            End If

            Observaciones = DBLet(vProve.Observaciones)
            If Observaciones <> "" Then
                MsgBox Observaciones, vbInformation, "Observaciones del proveedor"
            End If
        End If
    Else
        LimpiarDatosProve
        PonerFoco text1(4)
    End If
    Set vProve = Nothing

EPonerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Datos Proveedor", Err.Description
End Sub


Private Sub PonerDatosProveVario(nifProve As String)
'Poner el los campos Text el valor del cliente
Dim vProve As CProveedor
Dim b As Boolean
   
    If nifProve = "" Then Exit Sub
   
    Set vProve = New CProveedor
    b = vProve.LeerDatosProveVario(nifProve)
    
    If b Then
        text1(5).Text = vProve.Nombre   'Nom proveedor
        text1(8).Text = vProve.Domicilio
        text1(9).Text = vProve.CPostal
        text1(10).Text = vProve.Poblacion
        text1(11).Text = vProve.Provincia
        text1(7).Text = DBLet(vProve.TfnoAdmon, "T")
    End If
    Set vProve = Nothing
End Sub


Private Sub BloquearDatosProve(bol As Boolean)
Dim i As Byte

    'bloquear/desbloquear campos de datos segun sea proveedor de varios o no
    If Modo <> 5 Then
        Me.imgBuscar(6).visible = bol 'NIF
        Me.imgBuscar(6).Enabled = bol 'NIF
'--monica: no hay acceso a la tabla de poblaciones
'        Me.imgBuscar(1).Enabled = bol 'poblacion
        
        For i = 5 To 11 'si no es de varios no se pueden modificar los datos
            BloquearTxt text1(i), Not bol
            text1(i).Enabled = bol
        Next i
    End If
End Sub


Private Function ActualizarProveVarios(Prove As String, NIF As String) As Boolean
Dim vProve As CProveedor

    On Error GoTo EActualizarCV

    ActualizarProveVarios = False
    
    Set vProve = New CProveedor
    If EsProveedorVarios(Prove) Then
        vProve.NIF = NIF
        vProve.Nombre = text1(5).Text
        vProve.Domicilio = text1(8).Text
        vProve.CPostal = text1(9).Text
        vProve.Poblacion = text1(10).Text
        vProve.Provincia = text1(11).Text
        vProve.TfnoAdmon = text1(7).Text
        'Actualiza la tabla de proveedores varios con los datos que tenemos
        vProve.ActualizarProveV (NIF)
    End If
    Set vProve = Nothing
    
    ActualizarProveVarios = True
    
EActualizarCV:
    If Err.Number <> 0 Then
        ActualizarProveVarios = False
    Else
        ActualizarProveVarios = True
    End If
End Function


Private Sub CalcularDatosFactura()
Dim i As Byte
Dim cadwhere As String
Dim vFactu As CFacturaCom

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For i = 33 To 50
         Text3(i).Text = ""
    Next i
    
    cadwhere = ObtenerWhereCP(False)
    
    Set vFactu = New CFacturaCom
    vFactu.DtoPPago = CCur(ComprobarCero(text1(13).Text))
    vFactu.DtoGnral = CCur(ComprobarCero(text1(14).Text))
    If vFactu.CalcularDatosFactura(cadwhere, NombreTabla, NomTablaLineas) Then
        Text3(33).Text = vFactu.BrutoFac
        Text3(34).Text = vFactu.ImpPPago
        Text3(35).Text = vFactu.ImpGnral
        Text3(36).Text = vFactu.BaseImp
        Text3(37).Text = QuitarCero(vFactu.TipoIVA1)
        Text3(38).Text = QuitarCero(vFactu.TipoIVA2)
        Text3(39).Text = QuitarCero(vFactu.TipoIVA3)
        Text3(40).Text = vFactu.PorceIVA1
        Text3(41).Text = vFactu.PorceIVA2
        Text3(42).Text = vFactu.PorceIVA3
        Text3(43).Text = vFactu.BaseIVA1
        Text3(44).Text = vFactu.BaseIVA2
        Text3(45).Text = vFactu.BaseIVA3
        Text3(46).Text = vFactu.ImpIVA1
        Text3(47).Text = vFactu.ImpIVA2
        Text3(48).Text = vFactu.ImpIVA3
        Text3(49).Text = vFactu.TotalFac
        Text3(50).Text = vFactu.BaseImp
       
        FormatoDatosTotales
        
    Else
        MuestraError Err.Number, "Calculando Totales", Err.Description
    End If
    Set vFactu = Nothing
End Sub



Private Sub FormatoDatosTotales()
Dim i As Byte

    For i = 33 To 36
        If i = 34 Or i = 35 Then Text3(i).Text = QuitarCero(Text3(i).Text)
        Text3(i).Text = Format(Text3(i).Text, FormatoImporte)
    Next i
    
    'Desglose B.Imponible por IVA
    For i = 43 To 45
        If Text3(i).Text <> "" Then
             If CSng(Text3(i).Text) = 0 And Text3(i - 6).Text = "" Then
                Text3(i).Text = QuitarCero(Text3(i).Text)
                Text3(i - 3).Text = QuitarCero(Text3(i - 3).Text)
                Text3(i - 6).Text = QuitarCero(Text3(i - 6).Text)
                Text3(i + 3).Text = QuitarCero(Text3(i + 3).Text)
            Else
                Text3(i).Text = Format(Text3(i).Text, FormatoImporte)
                Text3(i - 3) = Format(Text3(i - 3).Text, FormatoDescuento)
    '            Text3(i - 6) = Format(Text3(i - 6).Text, "000")
                Text3(i + 3).Text = Format(Text3(i + 3).Text, FormatoImporte)
            End If
        End If
    Next i
    
    'Los Totales
    For i = 49 To 50
'        Text3(i).Text = QuitarCero(Text3(i).Text)
        Text3(i).Text = Format(Text3(i).Text, FormatoImporte)
    Next i
End Sub




Private Function ActualizarUltFechaCom(cadW As String) As Boolean
''Actualiza la ultima fecha de compra y el ult. precio de compra
''en el articulo, poniendo los valores del albaran de compra
'Dim SQL As String
'Dim RS As ADODB.Recordset
'
'    On Error GoTo EActualizaFecha
'
'    SQL = "select distinct numalbar,fechaalb,slialp.codartic,max(slialp.precioar) as precioar , sartic.ultfecco "
'    SQL = SQL & " from slialp INNER JOIN sartic ON slialp.codartic=sartic.codartic "
''    SQL = SQL & " where numalbar='K2500088' and fechaalb='2005-10-06' and slialp.codprove=21"
'    SQL = SQL & " WHERE " & cadW
'    SQL = SQL & " and (fechaalb>ultfecco or isnull(ultfecco))"
'    SQL = SQL & " group by numalbar,fechaalb,slialp.codartic "
'    SQL = SQL & " order by codartic "
'
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    While Not RS.EOF
'        SQL = "UPDATE sartic SET ultfecco=" & DBSet(RS!FechaAlb, "F") & ", preciouc=" & DBSet(RS!precioar, "N")
'        SQL = SQL & " WHERE codartic=" & DBSet(RS!codArtic, "T")
'        Conn.Execute SQL
'        RS.MoveNext
'    Wend
'    RS.Close
'    Set RS = Nothing
'
'EActualizaFecha:
'    If Err.Number <> 0 Then
'        ActualizarUltFechaCom = False
'    Else
'        ActualizarUltFechaCom = True
'    End If
End Function



Private Function ObtenerRSprecios(ByRef Rs As ADODB.Recordset, cadSQL As String) As Boolean
    On Error GoTo ErrRS
    Set Rs = New ADODB.Recordset
    Rs.Open cadSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ObtenerRSprecios = True
    Exit Function
    
ErrRS:
    ObtenerRSprecios = False
    If Not Rs Is Nothing Then Set Rs = Nothing
    MuestraError Err.Number, "Cargando RS precios ultima compra.", Err.Description
End Function


Private Sub CargaForaGrid()
        If DataGrid1.Columns.Count <= 2 Then Exit Sub
        ' *** posar als camps de fora del grid el valor de la columna corresponent ***
        Text2(16) = DataGrid1.Columns(5).Text
        ' **********************************************************************
 End Sub

Private Sub CargarDatosArticulo(codArtic As String)
Dim Rs As ADODB.Recordset
Dim Sql As String
        
    On Error GoTo eCargarDatosArticulo
        
    If Trim(codArtic) <> "" Then
        Sql = "select nomfamia, nomunida from sartic, sfamia, sunida "
        Sql = Sql & " where sartic.codartic = " & DBSet(codArtic, "T")
        Sql = Sql & " and sartic.codfamia = sfamia.codfamia and sartic.codunida = sunida.codunida"
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        Text2(4).Text = ""
        Text2(5).Text = ""
        If Not Rs.EOF Then
            Text2(5).Text = DBLet(Rs.Fields(0).Value, "T")
            Text2(4).Text = DBLet(Rs.Fields(1).Value, "T")
        End If
        
        Set Rs = Nothing
    End If
    Exit Sub
    
eCargarDatosArticulo:
    MuestraError Err.Number, "Error Cargar Datos Articulos"
End Sub

