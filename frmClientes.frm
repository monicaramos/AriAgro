VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmClientes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clientes"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   14355
   Icon            =   "frmClientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   14355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Cuenta Principal|N|N|0|1|cltebanc|ctaprpal|0||"
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   3915
      TabIndex        =   160
      Top             =   135
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   161
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
      Left            =   11655
      TabIndex        =   158
      Top             =   330
      Width           =   1605
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   225
      TabIndex        =   156
      Top             =   135
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   157
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
      Height          =   855
      Index           =   0
      Left            =   240
      TabIndex        =   50
      Top             =   975
      Width           =   13905
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
         Left            =   1035
         MaxLength       =   6
         TabIndex        =   0
         Tag             =   "Código de cliente|N|N|1|999999|clientes|codclien|000000|S|"
         Top             =   315
         Width           =   1080
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
         Index           =   2
         Left            =   3555
         MaxLength       =   40
         TabIndex        =   1
         Tag             =   "Nombre|T|N|||clientes|nomclien|||"
         Top             =   315
         Width           =   5220
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre "
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
         Left            =   2745
         TabIndex        =   52
         Top             =   315
         Width           =   810
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
         Left            =   270
         TabIndex        =   51
         Top             =   315
         Width           =   675
      End
   End
   Begin VB.TextBox text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   25
      Left            =   9600
      TabIndex        =   119
      Top             =   1170
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   240
      TabIndex        =   47
      Top             =   7230
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
         TabIndex        =   48
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
      Left            =   13020
      TabIndex        =   30
      Top             =   7335
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
      Left            =   11730
      TabIndex        =   29
      Top             =   7335
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   4185
      Top             =   6135
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
      Left            =   13005
      TabIndex        =   55
      Top             =   7335
      Visible         =   0   'False
      Width           =   1065
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5280
      Left            =   225
      TabIndex        =   49
      Top             =   1935
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   9313
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   3
      TabsPerRow      =   9
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos básicos"
      TabPicture(0)   =   "frmClientes.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(1)=   "Label6(0)"
      Tab(0).Control(2)=   "imgWeb(0)"
      Tab(0).Control(3)=   "Label9"
      Tab(0).Control(4)=   "Label29"
      Tab(0).Control(5)=   "imgZoom(0)"
      Tab(0).Control(6)=   "imgBuscar(0)"
      Tab(0).Control(7)=   "Label7"
      Tab(0).Control(8)=   "Label28"
      Tab(0).Control(9)=   "Label1(26)"
      Tab(0).Control(10)=   "imgBuscar(1)"
      Tab(0).Control(11)=   "imgBuscar(8)"
      Tab(0).Control(12)=   "Label40"
      Tab(0).Control(13)=   "Text1(7)"
      Tab(0).Control(14)=   "Text1(8)"
      Tab(0).Control(15)=   "Text1(3)"
      Tab(0).Control(16)=   "Text1(4)"
      Tab(0).Control(17)=   "Text1(9)"
      Tab(0).Control(18)=   "Text1(21)"
      Tab(0).Control(19)=   "Text1(5)"
      Tab(0).Control(20)=   "Text1(6)"
      Tab(0).Control(21)=   "text2(6)"
      Tab(0).Control(22)=   "Text1(18)"
      Tab(0).Control(23)=   "Text1(22)"
      Tab(0).Control(24)=   "FrameDatosDtoAdministracion"
      Tab(0).ControlCount=   25
      TabCaption(1)   =   "Direcciones"
      TabPicture(1)   =   "frmClientes.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(2)=   "FrameDatosContacto"
      Tab(1).Control(3)=   "Frame3"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Destinos"
      TabPicture(2)   =   "frmClientes.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameAux0"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Documentos"
      TabPicture(3)   =   "frmClientes.frx":0060
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "imgFec(0)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label44"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label45"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Toolbar3"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "lw1"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Toolbar2"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Text3(0)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "Datos Seguros"
      TabPicture(4)   =   "frmClientes.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Text1(36)"
      Tab(4).Control(1)=   "Text1(35)"
      Tab(4).Control(2)=   "Text1(34)"
      Tab(4).Control(3)=   "Label48"
      Tab(4).Control(4)=   "Label47"
      Tab(4).Control(5)=   "Label46"
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "Precios"
      TabPicture(5)   =   "frmClientes.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "FrameAux1"
      Tab(5).ControlCount=   1
      Begin VB.Frame Frame5 
         Caption         =   "Datos EDI"
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
         Height          =   1860
         Left            =   -68610
         TabIndex        =   152
         Top             =   2760
         Width           =   5685
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
            Index           =   31
            Left            =   1845
            MaxLength       =   17
            TabIndex        =   45
            Tag             =   "Código Edi|T|S|||clientes|codigoedi|||"
            Text            =   "12345678901234567"
            Top             =   360
            Width           =   2205
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
            Index           =   2
            Left            =   1830
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Tag             =   "Destinatario|N|N|0|1|clientes|destedi||N|"
            Top             =   720
            Width           =   1440
         End
         Begin VB.Label Label25 
            Caption         =   "Código EDI"
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
            Left            =   225
            TabIndex        =   154
            Top             =   405
            Width           =   1005
         End
         Begin VB.Label Label43 
            Caption         =   "Dest.Fact.EDI"
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
            Left            =   225
            TabIndex        =   153
            Top             =   765
            Width           =   1440
         End
      End
      Begin VB.Frame FrameAux1 
         BorderStyle     =   0  'None
         Height          =   4755
         Left            =   -74955
         TabIndex        =   144
         Top             =   405
         Width           =   12360
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
            Height          =   300
            Index           =   0
            Left            =   1230
            MaskColor       =   &H00000000&
            TabIndex        =   151
            ToolTipText     =   "Buscar Artículo"
            Top             =   3420
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtaux2 
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
            Index           =   6
            Left            =   1500
            TabIndex        =   150
            Top             =   3420
            Visible         =   0   'False
            Width           =   3015
         End
         Begin VB.TextBox txtaux1 
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
            Left            =   4650
            MaxLength       =   40
            TabIndex        =   146
            Tag             =   "Precio|N|N|||clientes_precio|precioar|###,##0.0000||"
            Text            =   "precio"
            Top             =   3420
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtaux1 
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
            Left            =   -120
            MaxLength       =   6
            TabIndex        =   147
            Tag             =   "Código de cliente|N|N|1|999999|clientes_precio|codclien|000000|S|"
            Text            =   "Client"
            Top             =   3405
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtaux1 
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
            Left            =   360
            MaxLength       =   16
            TabIndex        =   145
            Tag             =   "Articulo|T|N|||clientes_precio|codartic||S|"
            Text            =   "Art"
            Top             =   3405
            Visible         =   0   'False
            Width           =   825
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   1
            Left            =   45
            TabIndex        =   148
            Top             =   0
            Width           =   1290
            _ExtentX        =   2275
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
            Index           =   1
            Left            =   3720
            Top             =   480
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
            Bindings        =   "frmClientes.frx":00B4
            Height          =   4200
            Index           =   1
            Left            =   0
            TabIndex        =   149
            Top             =   495
            Width           =   9240
            _ExtentX        =   16298
            _ExtentY        =   7408
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
      Begin VB.Frame Frame4 
         Caption         =   "Dirección de Correo"
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
         Height          =   1860
         Left            =   -74820
         TabIndex        =   140
         Top             =   2760
         Width           =   5715
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
            Index           =   40
            Left            =   1215
            MaxLength       =   40
            TabIndex        =   41
            Tag             =   "Domicilio Correo|T|N|||clientes|domcliencorreo|||"
            Top             =   390
            Width           =   4305
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
            Index           =   39
            Left            =   1215
            MaxLength       =   6
            TabIndex        =   42
            Tag             =   "C.Postal Correo|T|N|||clientes|codpoblacorreo|||"
            Top             =   810
            Width           =   735
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
            Index           =   38
            Left            =   2040
            MaxLength       =   35
            TabIndex        =   43
            Tag             =   "Población Correo|T|N|||clientes|pobcliencorreo|||"
            Top             =   810
            Width           =   3465
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
            Index           =   37
            Left            =   1215
            MaxLength       =   35
            TabIndex        =   44
            Tag             =   "Provincia Correo|T|S|||clientes|procliencorreo|||"
            Top             =   1245
            Width           =   4290
         End
         Begin VB.Label Label6 
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
            Index           =   1
            Left            =   210
            TabIndex        =   143
            Top             =   390
            Width           =   1050
         End
         Begin VB.Label Label49 
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
            Left            =   210
            TabIndex        =   142
            Top             =   1305
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
            Index           =   1
            Left            =   210
            TabIndex        =   141
            Top             =   840
            Width           =   1050
         End
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
         Index           =   36
         Left            =   -72840
         MaxLength       =   15
         TabIndex        =   136
         Tag             =   "Limite Riesgos|N|S|||clientes|limiteriesgos|###,##0.00||"
         Top             =   1500
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
         Index           =   35
         Left            =   -72840
         MaxLength       =   4
         TabIndex        =   135
         Tag             =   "Días Asegurados|N|S|||clientes|diasasegurados|##0||"
         Top             =   1050
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
         Index           =   34
         Left            =   -72840
         MaxLength       =   15
         TabIndex        =   134
         Tag             =   "Nro.Seguro|T|S|||clientes|nroseguro|||"
         Text            =   "123456789012345"
         Top             =   600
         Width           =   1530
      End
      Begin VB.TextBox Text3 
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
         Left            =   11775
         TabIndex        =   128
         Text            =   "Text4"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Frame FrameDatosContacto 
         Caption         =   "Administración"
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
         Height          =   2010
         Left            =   -74820
         TabIndex        =   111
         Top             =   675
         Width           =   5730
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
            Left            =   1245
            MaxLength       =   40
            TabIndex        =   34
            Tag             =   "E-mail|T|S|||clientes|maiclie1|||"
            Top             =   1125
            Width           =   4320
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
            Index           =   13
            Left            =   1245
            MaxLength       =   15
            TabIndex        =   33
            Tag             =   "Fax|T|S|||clientes|faxclie1|||"
            Text            =   "123456789012345"
            Top             =   735
            Width           =   1680
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
            Left            =   4140
            MaxLength       =   10
            TabIndex        =   32
            Tag             =   "Móvil|T|S|||clientes|movclie1|||"
            Top             =   330
            Width           =   1455
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
            Left            =   1245
            MaxLength       =   15
            TabIndex        =   31
            Tag             =   "Teléfono|T|S|||clientes|telclie1|||"
            Text            =   "123456789012345"
            Top             =   330
            Width           =   1680
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
            Left            =   1245
            MaxLength       =   35
            TabIndex        =   35
            Tag             =   "Persona Contacto Admin.|T|S|||clientes|perclie1|||"
            Top             =   1530
            Width           =   4335
         End
         Begin VB.Label Label16 
            Caption         =   "E-mail"
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
            Left            =   225
            TabIndex        =   116
            Top             =   1155
            Width           =   675
         End
         Begin VB.Label Label14 
            Caption         =   "Fax"
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
            Left            =   225
            TabIndex        =   115
            Top             =   735
            Width           =   735
         End
         Begin VB.Label Label12 
            Caption         =   "Móvil"
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
            Left            =   3555
            TabIndex        =   114
            Top             =   330
            Width           =   495
         End
         Begin VB.Label Label10 
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
            Left            =   225
            TabIndex        =   113
            Top             =   330
            Width           =   960
         End
         Begin VB.Image imgMail 
            Height          =   240
            Index           =   0
            Left            =   945
            Top             =   1170
            Width           =   240
         End
         Begin VB.Label Label11 
            Caption         =   "Contacto"
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
            Left            =   225
            TabIndex        =   112
            Top             =   1530
            Width           =   1005
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Comercial"
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
         Height          =   2010
         Left            =   -68610
         TabIndex        =   105
         Top             =   675
         Width           =   5685
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
            Index           =   15
            Left            =   1200
            MaxLength       =   35
            TabIndex        =   40
            Tag             =   "Persona Contacto Comer.|T|S|||clientes|perclie2|||"
            Top             =   1530
            Width           =   4320
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
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   36
            Tag             =   "Teléfono|T|S|||clientes|telclie2|||"
            Text            =   "123456789012345"
            Top             =   330
            Width           =   1635
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
            Left            =   4095
            MaxLength       =   10
            TabIndex        =   37
            Tag             =   "Móvil|T|S|||clientes|movclie2|||"
            Top             =   315
            Width           =   1455
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
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   38
            Tag             =   "Fax|T|S|||clientes|faxclie2|||"
            Text            =   "123456789012345"
            Top             =   735
            Width           =   1635
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
            Left            =   1200
            MaxLength       =   40
            TabIndex        =   39
            Tag             =   "E-mail|T|S|||clientes|maiclie2|||"
            Top             =   1125
            Width           =   4320
         End
         Begin VB.Label Label13 
            Caption         =   "Contacto"
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
            Left            =   225
            TabIndex        =   110
            Top             =   1530
            Width           =   915
         End
         Begin VB.Image imgMail 
            Height          =   240
            Index           =   1
            Left            =   900
            Top             =   1200
            Width           =   240
         End
         Begin VB.Label Label15 
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
            Left            =   225
            TabIndex        =   109
            Top             =   330
            Width           =   960
         End
         Begin VB.Label Label17 
            Caption         =   "Móvil"
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
            Left            =   3510
            TabIndex        =   108
            Top             =   330
            Width           =   495
         End
         Begin VB.Label Label26 
            Caption         =   "Fax"
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
            Left            =   225
            TabIndex        =   107
            Top             =   735
            Width           =   735
         End
         Begin VB.Label Label27 
            Caption         =   "E-mail"
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
            Left            =   225
            TabIndex        =   106
            Top             =   1155
            Width           =   630
         End
      End
      Begin VB.Frame FrameAux0 
         BorderStyle     =   0  'None
         Height          =   4755
         Left            =   -74955
         TabIndex        =   71
         Top             =   360
         Width           =   13755
         Begin VB.TextBox txtaux 
            BeginProperty Font 
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
            Left            =   9630
            MaxLength       =   17
            TabIndex        =   95
            Tag             =   "Código Edicom|T|S|||destinos|codigoedi|||"
            Top             =   4140
            Width           =   1875
         End
         Begin VB.TextBox txtaux 
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
            Height          =   290
            Index           =   15
            Left            =   360
            MaxLength       =   4
            TabIndex        =   76
            Tag             =   "Código destino|N|N|1|9999|destinos|coddesti|0000|S|"
            Text            =   "co"
            Top             =   3405
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtaux 
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
            Height          =   290
            Index           =   14
            Left            =   -120
            MaxLength       =   6
            TabIndex        =   75
            Tag             =   "Código de cliente|N|N|1|999999|destinos|codclien|000000|S|"
            Text            =   "codc"
            Top             =   3405
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtaux 
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
            Height          =   290
            Index           =   0
            Left            =   990
            MaxLength       =   40
            TabIndex        =   78
            Tag             =   "Nombre|T|N|||destinos|nomdesti|||"
            Text            =   "nombre"
            Top             =   3420
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtaux 
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
            Height          =   290
            Index           =   1
            Left            =   2925
            MaxLength       =   6
            TabIndex        =   82
            Tag             =   "Cod.Pobla|T|S|||destinos|codpobla|||"
            Text            =   "C.P."
            Top             =   3420
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtaux 
            BeginProperty Font 
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
            Left            =   9630
            MaxLength       =   30
            TabIndex        =   83
            Tag             =   "Poblacion|T|S|||destinos|pobdesti|||"
            Top             =   90
            Width           =   4065
         End
         Begin VB.TextBox txtaux 
            BeginProperty Font 
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
            Left            =   9630
            MaxLength       =   30
            TabIndex        =   84
            Tag             =   "Provincia|T|S|||destinos|prodesti|||"
            Text            =   "Prov"
            Top             =   495
            Width           =   4065
         End
         Begin VB.TextBox txtaux2 
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
            Left            =   10170
            TabIndex        =   74
            Top             =   2925
            Width           =   3525
         End
         Begin VB.TextBox txtaux 
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
            Left            =   9630
            MaxLength       =   3
            TabIndex        =   91
            Tag             =   "Tipo de mercado|N|N|||destinos|codtimer|000||"
            Top             =   2925
            Width           =   495
         End
         Begin VB.TextBox txtaux2 
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
            Left            =   10170
            TabIndex        =   73
            Top             =   3330
            Width           =   3525
         End
         Begin VB.TextBox txtaux 
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
            Left            =   9630
            MaxLength       =   3
            TabIndex        =   92
            Tag             =   "Cadena|N|N|||destinos|codcaden|000||"
            Top             =   3330
            Width           =   495
         End
         Begin VB.TextBox txtaux 
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
            Left            =   9615
            MaxLength       =   10
            TabIndex        =   93
            Tag             =   "Cajas|T|S|||destinos|codcajas|||"
            Top             =   3765
            Width           =   1575
         End
         Begin VB.TextBox txtaux 
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
            Left            =   12285
            MaxLength       =   10
            TabIndex        =   94
            Tag             =   "Palets|T|S|||destinos|codpalet|||"
            Top             =   3780
            Width           =   1395
         End
         Begin VB.TextBox txtaux 
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
            Height          =   290
            Index           =   16
            Left            =   2250
            MaxLength       =   35
            TabIndex        =   80
            Tag             =   "Domicilio|T|S|||destinos|domdesti|||"
            Text            =   "domicilio"
            Top             =   3420
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtaux2 
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
            Left            =   10170
            TabIndex        =   72
            Top             =   2520
            Width           =   3510
         End
         Begin VB.TextBox txtaux 
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
            Left            =   9630
            MaxLength       =   3
            TabIndex        =   90
            Tag             =   "Cod.Pais|N|N|||destinos|codpaise|000||"
            Top             =   2520
            Width           =   510
         End
         Begin VB.TextBox txtaux 
            BeginProperty Font 
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
            Left            =   9630
            MaxLength       =   35
            TabIndex        =   89
            Tag             =   "Persona contacto|T|S|||destinos|perdesti|||"
            Top             =   2115
            Width           =   4035
         End
         Begin VB.TextBox txtaux 
            BeginProperty Font 
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
            Left            =   9630
            MaxLength       =   40
            TabIndex        =   88
            Tag             =   "Mail|T|S|||destinos|maidesti|||"
            Top             =   1710
            Width           =   4035
         End
         Begin VB.TextBox txtaux 
            BeginProperty Font 
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
            Left            =   9630
            MaxLength       =   10
            TabIndex        =   87
            Tag             =   "Fax|T|S|||destinos|faxdesti|||"
            Text            =   "1234567890"
            Top             =   1260
            Width           =   1380
         End
         Begin VB.TextBox txtaux 
            BeginProperty Font 
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
            Left            =   12240
            MaxLength       =   10
            TabIndex        =   86
            Tag             =   "Movil|T|S|||destinos|movdesti|||"
            Text            =   "1234567890"
            Top             =   900
            Width           =   1440
         End
         Begin VB.TextBox txtaux 
            BeginProperty Font 
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
            Left            =   9630
            MaxLength       =   10
            TabIndex        =   85
            Tag             =   "Teléfono|T|S|||destinos|teldesti|||"
            Text            =   "1234567890"
            Top             =   900
            Width           =   1380
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   0
            Left            =   45
            TabIndex        =   77
            Top             =   0
            Width           =   1290
            _ExtentX        =   2275
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
            Top             =   480
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
            Bindings        =   "frmClientes.frx":00CC
            Height          =   4200
            Index           =   0
            Left            =   45
            TabIndex        =   79
            Top             =   450
            Width           =   8355
            _ExtentX        =   14737
            _ExtentY        =   7408
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
         Begin VB.Image imgAyuda 
            Enabled         =   0   'False
            Height          =   240
            Index           =   0
            Left            =   1380
            ToolTipText     =   "Buscar Destinos"
            Top             =   60
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label39 
            Caption         =   "Código EDI"
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
            Left            =   8460
            TabIndex        =   124
            Top             =   4185
            Width           =   1140
         End
         Begin VB.Label Label24 
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
            Left            =   8460
            TabIndex        =   123
            Top             =   495
            Width           =   1230
         End
         Begin VB.Label Label2 
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
            Left            =   8460
            TabIndex        =   122
            Top             =   135
            Width           =   1230
         End
         Begin VB.Image imgMail 
            Height          =   240
            Index           =   2
            Left            =   9315
            Top             =   1710
            Width           =   240
         End
         Begin VB.Label Label34 
            Caption         =   "Mercado"
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
            Left            =   8460
            TabIndex        =   104
            Top             =   2925
            Width           =   825
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   9315
            ToolTipText     =   "Buscar Colectivo"
            Top             =   2970
            Width           =   240
         End
         Begin VB.Label Label35 
            Caption         =   "Cadena"
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
            Left            =   8460
            TabIndex        =   103
            Top             =   3330
            Width           =   780
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   9315
            ToolTipText     =   "Buscar Colectivo"
            Top             =   3330
            Width           =   240
         End
         Begin VB.Label Label36 
            Caption         =   "Cód.Cajas"
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
            Left            =   8460
            TabIndex        =   102
            Top             =   3780
            Width           =   1095
         End
         Begin VB.Label Label37 
            Caption         =   "Cód.Palets"
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
            Left            =   11205
            TabIndex        =   101
            Top             =   3780
            Width           =   1050
         End
         Begin VB.Label Label38 
            Caption         =   "País"
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
            Left            =   8460
            TabIndex        =   100
            Top             =   2520
            Width           =   420
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   7
            Left            =   9315
            ToolTipText     =   "Buscar Colectivo"
            Top             =   2520
            Width           =   240
         End
         Begin VB.Label Label33 
            Caption         =   "Contacto"
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
            Left            =   8460
            TabIndex        =   99
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label Label32 
            Caption         =   "Mail"
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
            Left            =   8460
            TabIndex        =   98
            Top             =   1710
            Width           =   600
         End
         Begin VB.Label Label31 
            Caption         =   "Fax"
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
            Left            =   8460
            TabIndex        =   97
            Top             =   1260
            Width           =   420
         End
         Begin VB.Label Label30 
            Caption         =   "Móvil"
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
            Left            =   11700
            TabIndex        =   96
            Top             =   930
            Width           =   555
         End
         Begin VB.Label Label3 
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
            Left            =   8460
            TabIndex        =   81
            Top             =   930
            Width           =   870
         End
      End
      Begin VB.Frame FrameDatosDtoAdministracion 
         Caption         =   "Datos Relacionados Dto.Administración"
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
         Height          =   4215
         Left            =   -68790
         TabIndex        =   62
         Top             =   405
         Width           =   7365
         Begin VB.CheckBox chkAbonos 
            Caption         =   "Envio CMR por eMail"
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
            Left            =   225
            TabIndex        =   21
            Tag             =   "Envio CMR por email|N|N|||clientes|envcmremail||N|"
            Top             =   2535
            Width           =   3180
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
            Index           =   42
            Left            =   1875
            MaxLength       =   4
            TabIndex        =   12
            Tag             =   "IBAN|T|S|||clientes|iban|||"
            Text            =   "Text"
            Top             =   1260
            Width           =   705
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
            Index           =   41
            Left            =   5715
            MaxLength       =   3
            TabIndex        =   22
            Tag             =   "Tipo Movimiento  Albaran|T|N|||clientes|codtipalb|||"
            Top             =   2520
            Width           =   1230
         End
         Begin VB.CheckBox chkAbonos 
            Caption         =   "Control Cobros por Albarán"
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
            Left            =   225
            TabIndex        =   27
            Tag             =   "Control Cobros Alb|N|N|0|1|clientes|ctrolcobroalb|0|N|"
            Top             =   3720
            Width           =   3315
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
            Index           =   33
            Left            =   5715
            MaxLength       =   1
            TabIndex        =   26
            Tag             =   "Decimales Redondeo Precio|N|N|2|4|clientes|nrodecprec|0||"
            Text            =   "1"
            Top             =   3300
            Width           =   405
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
            Index           =   32
            Left            =   5715
            MaxLength       =   17
            TabIndex        =   28
            Tag             =   "Copias Albarán|N|N|1|99|clientes|nrocopias|00||"
            Text            =   "12345678901234567"
            Top             =   3690
            Width           =   1230
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
            Index           =   30
            Left            =   4830
            MaxLength       =   10
            TabIndex        =   16
            Tag             =   "Cuenta Bancaria|T|S|||clientes|cuentaba|0000000000||"
            Text            =   "Text1"
            Top             =   1260
            Width           =   2070
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
            Index           =   29
            Left            =   4215
            MaxLength       =   2
            TabIndex        =   15
            Tag             =   "Digito Control|T|S|||clientes|digcontr|00||"
            Text            =   "Te"
            Top             =   1260
            Width           =   495
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
            Index           =   28
            Left            =   3435
            MaxLength       =   4
            TabIndex        =   14
            Tag             =   "Sucursal|N|S|0|9999|clientes|codsucur|0000||"
            Text            =   "Text"
            Top             =   1260
            Width           =   705
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
            Left            =   2670
            MaxLength       =   4
            TabIndex        =   13
            Tag             =   "Banco|N|S|0|9999|clientes|codbanco|0000||"
            Text            =   "Text"
            Top             =   1260
            Width           =   705
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
            Index           =   25
            Left            =   5715
            MaxLength       =   3
            TabIndex        =   24
            Tag             =   "Tipo Movimiento|T|S|||clientes|codtipom|||"
            Top             =   2910
            Width           =   1230
         End
         Begin VB.CheckBox chkAbonos 
            Caption         =   "Contador Manual Factura"
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
            Left            =   225
            TabIndex        =   23
            Tag             =   "Tipo Factura|N|N|||clientes|tipofact||N|"
            Top             =   2925
            Width           =   3180
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
            Left            =   1215
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Tag             =   "Tipo Iva|N|N|||clientes|tipoiva||N|"
            Top             =   1665
            Width           =   1575
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
            Left            =   1215
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Tag             =   "Tipo Descuento|N|N|||clientes|tipodtos||N|"
            Top             =   2070
            Width           =   1575
         End
         Begin VB.CheckBox chkAbonos 
            Caption         =   "Utiliza Cta.Ventas alternativa"
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
            Left            =   225
            TabIndex        =   25
            Tag             =   "Cancela abonos|N|N|||clientes|cliabono||N|"
            Top             =   3315
            Width           =   3315
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
            Index           =   27
            Left            =   1875
            MaxLength       =   3
            TabIndex        =   11
            Tag             =   "Código F.Pago|N|N|0|999|clientes|codforpa|000||"
            Top             =   855
            Width           =   555
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
            Index           =   27
            Left            =   2475
            TabIndex        =   64
            Top             =   855
            Width           =   4440
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
            Left            =   5715
            MaxLength       =   8
            TabIndex        =   20
            Tag             =   "Poc.Comis.2|N|S|0|100.00|clientes|porccom2|##0.00||"
            Top             =   2085
            Width           =   1230
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
            Index           =   24
            Left            =   3330
            TabIndex        =   63
            Top             =   450
            Width           =   3600
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
            Index           =   24
            Left            =   1875
            MaxLength       =   10
            TabIndex        =   10
            Tag             =   "Cta.Contable|T|S|||clientes|codmacta|||"
            Top             =   450
            Width           =   1395
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
            Index           =   23
            Left            =   4005
            MaxLength       =   8
            TabIndex        =   19
            Tag             =   "Porc.Comis.1|N|S|0|100.00|clientes|porccom1|#0.00||"
            Top             =   2085
            Width           =   1185
         End
         Begin VB.Label Label50 
            Caption         =   "Tipo Movimiento Alb."
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
            Left            =   4005
            TabIndex        =   155
            Top             =   2550
            Width           =   1875
         End
         Begin VB.Label Label42 
            Caption         =   "Decimales Precio"
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
            Left            =   4005
            TabIndex        =   127
            Top             =   3345
            Width           =   1680
         End
         Begin VB.Label Label41 
            Caption         =   "Nro.Copias Albarán"
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
            Left            =   4005
            TabIndex        =   126
            Top             =   3735
            Width           =   1680
         End
         Begin VB.Label Label1 
            Caption         =   "IBAN Cliente"
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
            Left            =   225
            TabIndex        =   121
            Top             =   1305
            Width           =   1005
         End
         Begin VB.Label Label19 
            Caption         =   "Tipo Movimiento Fra."
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
            Left            =   4005
            TabIndex        =   120
            Top             =   2940
            Width           =   1875
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   6
            Left            =   1620
            ToolTipText     =   "Buscar F.Pago"
            Top             =   900
            Width           =   240
         End
         Begin VB.Label Label23 
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
            Left            =   225
            TabIndex        =   70
            Top             =   915
            Width           =   1200
         End
         Begin VB.Label Label22 
            Caption         =   "%Comisión 2"
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
            Left            =   5715
            TabIndex        =   69
            Top             =   1830
            Width           =   1125
         End
         Begin VB.Label Label21 
            Caption         =   "Tipo Iva"
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
            Left            =   225
            TabIndex        =   68
            Top             =   1710
            Width           =   1065
         End
         Begin VB.Label Label20 
            Caption         =   "Cta.Contable"
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
            Left            =   225
            TabIndex        =   67
            Top             =   450
            Width           =   1320
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   2
            Left            =   1620
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   450
            Width           =   240
         End
         Begin VB.Label Label18 
            Caption         =   "%Comisión 1"
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
            Left            =   4005
            TabIndex        =   66
            Top             =   1815
            Width           =   1035
         End
         Begin VB.Label Label8 
            Caption         =   "Tipo Dto."
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
            Left            =   225
            TabIndex        =   65
            Top             =   2115
            Width           =   1305
         End
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
         Index           =   22
         Left            =   -73620
         MaxLength       =   35
         TabIndex        =   6
         Tag             =   "Provincia|T|S|||clientes|proclien|||"
         Top             =   1695
         Width           =   4560
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
         Left            =   -72795
         MaxLength       =   35
         TabIndex        =   5
         Tag             =   "Población|T|S|||clientes|pobclien|||"
         Top             =   1305
         Width           =   3735
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
         Index           =   6
         Left            =   -73080
         TabIndex        =   58
         Top             =   2100
         Width           =   4020
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
         Left            =   -73620
         MaxLength       =   3
         TabIndex        =   7
         Tag             =   "Pais|N|S|0|999|clientes|codpaise|000||"
         Top             =   2100
         Width           =   495
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
         Left            =   -73635
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "C.Postal|T|S|||clientes|codpobla|||"
         Top             =   1305
         Width           =   735
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
         Height          =   1365
         Index           =   21
         Left            =   -74805
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Tag             =   "Observaciones|T|S|||clientes|observac|||"
         Top             =   3255
         Width           =   5715
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
         Left            =   -73620
         MaxLength       =   40
         TabIndex        =   8
         Tag             =   "Web|T|S|||clientes|wwwclien|||"
         Top             =   2520
         Width           =   4530
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
         Index           =   4
         Left            =   -73620
         MaxLength       =   40
         TabIndex        =   3
         Tag             =   "Domicilio|T|S|||clientes|domclien|||"
         Top             =   900
         Width           =   4575
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
         Index           =   3
         Left            =   -73620
         MaxLength       =   15
         TabIndex        =   2
         Tag             =   "NIF / CIF|T|N|||clientes|cifclien|||"
         Text            =   "123456789012345"
         Top             =   480
         Width           =   1635
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   -67845
         MaxLength       =   3
         TabIndex        =   117
         Top             =   630
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   -68790
         MaxLength       =   15
         TabIndex        =   118
         Top             =   495
         Width           =   1200
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   2370
         Left            =   90
         TabIndex        =   129
         Top             =   450
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   4180
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Appearance      =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Pedidos"
               Object.Tag             =   "0"
               Style           =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Albaranes"
               Object.Tag             =   "1"
               Style           =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Albaranes Envases"
               Object.Tag             =   "2"
               Style           =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Facturas"
               Object.Tag             =   "3"
               Style           =   2
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lw1 
         Height          =   3855
         Left            =   930
         TabIndex        =   130
         Top             =   450
         Width           =   9335
         _ExtentX        =   16457
         _ExtentY        =   6800
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
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
         NumItems        =   0
      End
      Begin MSComctlLib.Toolbar Toolbar3 
         Height          =   2370
         Left            =   90
         TabIndex        =   133
         Top             =   450
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   4180
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Appearance      =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Pedidos"
               Object.Tag             =   "0"
               Style           =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Albaranes"
               Object.Tag             =   "1"
               Style           =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Albaranes Envases"
               Object.Tag             =   "2"
               Style           =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Facturas"
               Object.Tag             =   "3"
               Style           =   2
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin VB.Label Label48 
         Caption         =   "Límite Riesgos"
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
         Left            =   -74550
         TabIndex        =   139
         Top             =   1500
         Width           =   1500
      End
      Begin VB.Label Label47 
         Caption         =   "Días asegurados"
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
         Left            =   -74550
         TabIndex        =   138
         Top             =   1050
         Width           =   1695
      End
      Begin VB.Label Label46 
         Caption         =   "Nro.Seguro"
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
         Left            =   -74550
         TabIndex        =   137
         Top             =   630
         Width           =   1305
      End
      Begin VB.Label Label45 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   300
         Left            =   10515
         TabIndex        =   132
         Top             =   540
         Width           =   2865
      End
      Begin VB.Label Label44 
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
         Height          =   255
         Left            =   10575
         TabIndex        =   131
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   11325
         Picture         =   "frmClientes.frx":00E4
         ToolTipText     =   "Buscar fecha"
         Top             =   1080
         Width           =   240
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
         Left            =   -70410
         TabIndex        =   125
         Top             =   450
         Width           =   960
      End
      Begin VB.Image imgBuscar 
         Height          =   330
         Index           =   8
         Left            =   -69420
         ToolTipText     =   "Códigos EAN asociados"
         Top             =   405
         Width           =   375
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   -67935
         ToolTipText     =   "Buscar F.Pago"
         Top             =   1035
         Width           =   240
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
         Index           =   26
         Left            =   -74775
         TabIndex        =   61
         Top             =   1365
         Width           =   1050
      End
      Begin VB.Label Label28 
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
         Left            =   -74775
         TabIndex        =   60
         Top             =   1755
         Width           =   1005
      End
      Begin VB.Label Label7 
         Caption         =   "País"
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
         Left            =   -74775
         TabIndex        =   59
         Top             =   2130
         Width           =   735
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   -73935
         ToolTipText     =   "Buscar País"
         Top             =   2130
         Width           =   240
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   -73260
         Tag             =   "-1"
         ToolTipText     =   "Zoom descripción"
         Top             =   2985
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
         Left            =   -74790
         TabIndex        =   57
         Top             =   2985
         Width           =   1485
      End
      Begin VB.Label Label9 
         Caption         =   "Web"
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
         Left            =   -74760
         TabIndex        =   56
         Top             =   2520
         Width           =   495
      End
      Begin VB.Image imgWeb 
         Height          =   240
         Index           =   0
         Left            =   -73920
         Picture         =   "frmClientes.frx":016F
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label6 
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
         Index           =   0
         Left            =   -74760
         TabIndex        =   54
         Top             =   930
         Width           =   960
      End
      Begin VB.Label Label5 
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
         Left            =   -74760
         TabIndex        =   53
         Top             =   480
         Width           =   735
      End
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   13695
      TabIndex        =   159
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
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   5
      Left            =   6525
      ToolTipText     =   "Buscar T.Iva"
      Top             =   5220
      Width           =   240
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
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: CÈSAR                    -+-+
' +-+- Menú: General-Clientes-Clientes -+-+
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

' *** declarar els formularis als que vaig a cridar ***
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmC1 As frmCal 'calendario fecha
Attribute frmC1.VB_VarHelpID = -1

Private WithEvents frmPais As frmManPaises 'paises
Attribute frmPais.VB_VarHelpID = -1
Private WithEvents frmTMer As frmManTipMerc 'tipos de mercados
Attribute frmTMer.VB_VarHelpID = -1
Private WithEvents frmCad As frmManCadenas 'cadenas
Attribute frmCad.VB_VarHelpID = -1
Private WithEvents frmCtas As frmCtasConta 'cuenta contable
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmFPa As frmManFpago 'Formas de Pago
Attribute frmFPa.VB_VarHelpID = -1
Private WithEvents frmTIva As frmTipIVAConta 'Tipos de iva de conta
Attribute frmTIva.VB_VarHelpID = -1
Private WithEvents frmCEan As frmCodEAN 'Codigos Ean
Attribute frmCEan.VB_VarHelpID = -1
Private WithEvents frmFac As frmVtasFacturas 'Facturas de ventas
Attribute frmFac.VB_VarHelpID = -1
Private WithEvents frmAlb As frmVtasAlbaranes 'Albaranes de ventas
Attribute frmAlb.VB_VarHelpID = -1
Private WithEvents frmAlbEnv As frmVtasAlbEnvases 'Albaranes de envases
Attribute frmAlbEnv.VB_VarHelpID = -1
Private WithEvents frmPed As frmVtasPedidos 'Pedidos clientes
Attribute frmPed.VB_VarHelpID = -1
Private WithEvents frmArt As frmManArtic ' mantenimiento de articulos
Attribute frmArt.VB_VarHelpID = -1
Private WithEvents frmDest As frmDestCli 'Form Mto de destinos de clientes
Attribute frmDest.VB_VarHelpID = -1
' *****************************************************
Private WithEvents frmPais2 As frmManPaises 'paises
Attribute frmPais2.VB_VarHelpID = -1


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

Private BuscaChekc As String


'Cambio en cuentas de la contabilidad
Dim IbanAnt As String
Dim NombreAnt As String
Dim BancoAnt  As String
Dim SucurAnt As String
Dim DigitoAnt As String
Dim CuentaAnt As String

Dim DirecAnt As String
Dim cPostalAnt As String
Dim PoblaAnt As String
Dim ProviAnt As String
Dim NifAnt As String
Dim forpaant As String


Dim EMaiAnt As String
Dim WebAnt As String


Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 0 'Articulos
            Set frmArt = New frmManArtic
            frmArt.DatosADevolverBusqueda = "0|1|"
            frmArt.CodigoActual = txtAux1(1).Text
            frmArt.Show vbModal
            Set frmArt = Nothing
            PonerFoco txtAux1(1)
    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub

Private Sub chkAbonos_GotFocus(Index As Integer)
    PonerFocoChk Me.chkAbonos(Index)
End Sub

Private Sub chkAbonos_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "chkAbonos(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "chkAbonos(" & Index & ")|"
    End If
End Sub

Private Sub chkAbonos_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkAbonos_LostFocus(Index As Integer)
    If Index = 1 And (Modo = 3 Or Modo = 4) Then
        If chkAbonos(Index).Value = 1 Then Text1(25).Text = ""
    End If
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
                    Text2(24).Text = PonerNombreCuenta(Text1(24), Modo, Text1(0))
                    ' *** canviar o llevar el WHERE, repasar codEmpre ****
                    Data1.RecordSource = "Select * from " & NombreTabla & Ordenacion
                    'Data1.RecordSource = "Select * from " & NombreTabla & " where codempre = " & codEmpre & Ordenacion
                    ' ***************************************************************
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
                    
                    '[Monica]22/08/2013: Si han cambiado nombre o CCC pregunto si quieren cambiar los datos de la cuenta en la seccion de horto
                    ModificarDatosCuentaContable
                    
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
                    ModificarLinea
                    PosicionarData
            End Select
        ' **************************
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



' *** si n'hi han combos a la capçalera ***
Private Sub Combo1_GotFocus(Index As Integer)
    If Modo = 1 Then Combo1(Index).BackColor = vbYellow
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    If Combo1(Index).BackColor = vbYellow Then Combo1(Index).BackColor = vbWhite
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PrimeraVez Then PrimeraVez = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia(0).Value
    Screen.MousePointer = vbDefault
    If Modo = 4 Or Modo = 5 Then TerminaBloquear
End Sub

Private Sub Form_Load()
Dim i As Integer

    PrimeraVez = True
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    ' ICONETS DE LA BARRA
'    btnPrimero = 16 'index del botó "primero"
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
'        'el 10 i el 11 son separadors
'        .Buttons(12).Image = 10  'Imprimir
'        .Buttons(13).Image = 11  'Eixir
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
    
    ' La Ayuda
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
    
    ' ayuda por destinos
    For i = 0 To Me.imgAyuda.Count - 1
        Me.imgAyuda(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    
    
    Me.imgBuscar(8).Picture = frmPpal.imgListComun.ListImages(21).Picture
    'carga IMAGES de mail
    For i = 0 To Me.ImgMail.Count - 1
        Me.ImgMail(i).Picture = frmPpal.imgListImages16.ListImages(2).Picture
    Next i
    
    'IMAGES para zoom
    For i = 0 To Me.imgZoom.Count - 1
        Me.imgZoom(i).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    Next i
    
   'La nevegacion para entradas, facturas....
    ImagenesNavegacion
   'Ponemos los datos del listview
    imgFec(0).Tag = vParam.FecIniCam
    CargaColumnas 0

    
    ' *** si n'hi han tabs, per a que per defecte sempre es pose al 1r***
    Me.SSTab1.Tab = 0
    ' *******************************************************************
    
    LimpiarCampos   'Neteja els camps TextBox
    ' ******* si n'hi han llínies *******
    DataGridAux(0).ClearFields
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "clientes"
    Ordenacion = " ORDER BY clientes.codclien"
    '************************************************
    
    'Mirem com està guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    '***** canviar el nom de la PK de la capçalera; repasar codEmpre *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where codclien=-1"
    Data1.Refresh
       
    ' ******* si n'hi han llinies en datagrid *******
'    ReDim CadAncho(DataGridAux.Count) 'redimensione l'array a la quantitat de datagrids
'    CadAncho(0) = False
'    CadAncho(1) = False
'    CadAncho(2) = False
'    CadAncho(4) = False
    
    ModoLineas = 0
       
    ' **** si n'hi ha algun frame que no te datagrids ***
'    CargaFrame 3, False
    ' *************************************************
         
    ' *** si n'hi han combos (capçalera o llínies) ***
    CargaCombo
    ' ************************************************
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1 'búsqueda
        ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
        Text1(0).BackColor = vbYellow 'codclien
        ' ****************************************************************************
    End If
End Sub


Private Sub LimpiarCampos()
Dim i As Integer

    On Error Resume Next
    
    limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    
    For i = 0 To Combo1.Count - 1
        Combo1(i).ListIndex = -1
        Me.chkAbonos(i).Value = 0
    Next i
'    Me.chkAbonos(2).Value = 0
    
    ' *** si n'hi han combos a la capçalera ***
    ' *****************************************

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub LimpiarCamposLin(frameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, frameAux  'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO s'habiliten, o no, els diversos camps del
'   formulari en funció del modo en que anem a treballar
Private Sub PonerModo(Kmodo As Byte, Optional indFrame As Integer)
Dim i As Integer, Numreg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo
 
    'Actualisa Iconos Insertar,Modificar,Eliminar
    'ActualizarToolbar Modo, Kmodo
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo, ModoLineas
    
    BuscaChekc = ""
       
    'Modo 2. N'hi han datos i estam visualisant-los
    '=========================================
    'Posem visible, si es formulari de búsqueda, el botó "Regresar" quan n'hi han datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = (Modo = 2)
    Else
        cmdRegresar.visible = False
    End If
    '=======================================
    
    'El campo 3(0) NUNCA se puede escribir en el
    Text3(0).Enabled = False
    Text3(0).Text = Me.imgFec(0).Tag

    b = (Modo = 2)
    'Posar Fleches de desplasament visibles
    Numreg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then Numreg = 2 'Només es per a saber que n'hi ha + d'1 registre
    End If
    'DesplazamientoVisible Me.Toolbar1, btnPrimero, b, Numreg
    DesplazamientoVisible b And Data1.Recordset.RecordCount > 1
    '---------------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
       
    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
    'Si estem en Insertar a més neteja els camps Text1
    BloquearText1 Me, Modo
    BloquearCombo Me, Modo
    BloquearChk Me.chkAbonos(0), (Modo = 0 Or Modo = 2 Or Modo = 5)
    BloquearChk Me.chkAbonos(1), (Modo = 0 Or Modo = 2 Or Modo = 5)
    BloquearChk Me.chkAbonos(2), (Modo = 0 Or Modo = 2 Or Modo = 5)
    BloquearChk Me.chkAbonos(3), (Modo = 0 Or Modo = 2 Or Modo = 5)
    
    ' ***** bloquejar tots els controls visibles de la clau primaria de la capçalera ***
'    If Modo = 4 Then _
'        BloquearTxt Text1(0), True 'si estic en  modificar, bloqueja la clau primaria
    ' **********************************************************************************
    
    ' **** si n'hi han imagens de buscar en la capçalera *****
    BloquearImgBuscar Me, Modo, ModoLineas
'    BloquearImgFec Me, 25, Modo, ModoLineas
    BloquearImgZoom Me, Modo, ModoLineas
    ' ********************************************************
    Me.imgBuscar(8).Enabled = (Modo = 2)
    Me.imgBuscar(8).visible = (Modo = 2)
    Label40.visible = (Modo = 2)
    ' *** si n'hi han llínies i imagens de buscar que no estiguen als grids ******
    'Llínies Departaments
    b = (Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) And (NumTabMto = 0))
    BloquearImage imgBuscar(3), Not b
    BloquearImage imgBuscar(4), Not b
    BloquearImage imgBuscar(7), Not b
'    imgBuscar(3).Enabled = b
'    imgBuscar(3).visible = b
    ' ****************************************************************************
    b = (Modo = 5 And (ModoLineas = 2) And (NumTabMto = 1))
    BloquearBtn btnBuscar(0), b
    BloquearTxt txtAux1(1), b
'    imgBuscar(3).Enabled = b
            
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    PonerLongCampos

    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 0, False
        CargaGrid 1, False
'        CargaGrid 2, False
'        CargaGrid 4, False
    End If
    
    b = (Modo = 4) Or (Modo = 2)
    DataGridAux(0).Enabled = b
      
'    ' ****** si n'hi han combos a la capçalera ***********************
'    If (Modo = 0) Or (Modo = 2) Or (Modo = 4) Or (Modo = 5) Then
'        Combo1(0).Enabled = False
'        Combo1(0).BackColor = &H80000018 'groc
'    ElseIf (Modo = 1) Or (Modo = 3) Then
'        Combo1(0).Enabled = True
'        Combo1(0).BackColor = &H80000005 'blanc
'    End If
'    ' ****************************************************************
    
    ' *** si n'hi han llínies i algún tab que no te datagrid ***
'    BloquearFrameAux Me, "FrameAux3", Modo, NumTabMto
'    BloquearFrameAux2 Me, "FrameAux3", (Modo <> 5) Or (Modo = 5 And indFrame <> 3) 'frame datos viaje indiv.
    ' ***************************
        
    b = (Modo = 5)
    For i = 0 To txtAux.Count - 1
        txtAux(i).Enabled = b
    Next i
    
    '[Monica]22/10/2012: buscar por codigo destino
    For i = 0 To 17 '09/09/2010
        BloquearTxt txtAux(i), True
        txtAux(i).Enabled = False
    Next i
    For i = 0 To 17
        BloquearTxt txtAux(i), (Modo <> 1 And Modo <> 5)
        txtAux(i).Enabled = (Modo = 1) Or (Modo = 5)
    Next i
    'Hasta aqui
    
'    If Modo = 1 Then
'        BloquearTxt txtAux(15), False
'        txtAux(15).Enabled = True
'        Dim anc As Single
'        anc = DataGridAux(0).Top
'        If DataGridAux(0).Row < 0 Then
'            anc = anc + 210
'        Else
'            anc = anc + DataGridAux(0).RowTop(DataGridAux(0).Row) + 5
'        End If
'        txtAux(15).Top = anc
'        txtAux(15).visible = True
'
'    End If
    
    
     '-----------------------------
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
    Toolbar1.Buttons(1).Enabled = b
    Me.mnNuevo.Enabled = b
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0) 'And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnEliminar.Enabled = b
    
    'Imprimir
    'Toolbar1.Buttons(12).Enabled = (b Or Modo = 0)
    Toolbar1.Buttons(8).Enabled = b
       
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
    b = (Modo = 3 Or Modo = 4 Or Modo = 2)
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
Dim SQL As String
Dim tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
        Case 0 'DESTINOS
            tabla = "destinos"
            SQL = "SELECT destinos.codclien,destinos.coddesti,destinos.nomdesti,destinos.domdesti,destinos.codpobla,destinos.pobdesti,destinos.prodesti,destinos.codpaise,paises.nompaise,destinos.codtimer,tipomer.nomtimer,teldesti,faxdesti,movdesti,maidesti,perdesti,codcajas,codpalet,codcaden, codigoedi "
            SQL = SQL & " FROM " & tabla & " INNER JOIN paises ON " & tabla & ".codpaise=paises.codpaise INNER JOIN tipomer ON " & tabla & ".codtimer=tipomer.codtimer "
            If enlaza Then
                SQL = SQL & ObtenerWhereCab(True)
            Else
                SQL = SQL & " WHERE codclien = -1"
            End If
            SQL = SQL & " ORDER BY " & tabla & ".coddesti "
            
        Case 1 'Clientes
            tabla = "clientes_precio"
            SQL = "SELECT clientes_precio.codclien,clientes_precio.codartic,sartic.nomartic,clientes_precio.precioar"
            SQL = SQL & " FROM " & tabla & " INNER JOIN sartic ON " & tabla & ".codartic=sartic.codartic "
            If enlaza Then
                SQL = SQL & ObtenerWhereCab(True)
            Else
                SQL = SQL & " WHERE codclien = -1"
            End If
            SQL = SQL & " ORDER BY " & tabla & ".codartic "
            
    End Select
    ' ********************************************************************************
    
    MontaSQLCarga = SQL
End Function

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Articulos
    txtAux1(1).Text = RecuperaValor(CadenaSeleccion, 1) 'codartic
    txtAux2(6).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
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
        ' *** canviar o llevar el WHERE; repasar codEmpre ***
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        'CadenaConsulta = "select * from " & NombreTabla & " WHERE codempre = " & codEmpre & " AND " & CadB & " " & Ordenacion
        ' **********************************
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmC1_Selec(vFecha As Date)
Dim indice As Byte
    indice = CByte(Me.imgFec(0).Tag)
    Text3(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCad_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento Cadenas
    txtAux(11).Text = RecuperaValor(CadenaSeleccion, 1) 'codcadena
    FormateaCampo txtAux(11)
    txtAux2(1).Text = RecuperaValor(CadenaSeleccion, 2) 'nomcadena
End Sub

Private Sub frmDest_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        AdoAux(0).Recordset.Find (AdoAux(0).Recordset.Fields(1).Name & " =" & RecuperaValor(CadenaSeleccion, 1))
    End If
End Sub

Private Sub frmFpa_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento Formas de pago
    Text1(27).Text = RecuperaValor(CadenaSeleccion, 1) 'codforpa
    FormateaCampo Text1(27)
    Text2(27).Text = RecuperaValor(CadenaSeleccion, 2) 'nomforpa
End Sub

Private Sub frmPais2_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento Paises
    txtAux(4).Text = RecuperaValor(CadenaSeleccion, 1) 'codpaise
    FormateaCampo txtAux(4)
    txtAux2(2).Text = RecuperaValor(CadenaSeleccion, 2) 'nompaise
End Sub

Private Sub frmPais_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento paises
    Text1(6).Text = RecuperaValor(CadenaSeleccion, 1) 'codpais
    FormateaCampo Text1(6)
    Text2(6).Text = RecuperaValor(CadenaSeleccion, 2) 'nompais

End Sub

Private Sub frmTIva_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento Tipo de iva
    Text1(25).Text = RecuperaValor(CadenaSeleccion, 1) 'codiva
    FormateaCampo Text1(25)
    Text2(25).Text = RecuperaValor(CadenaSeleccion, 2) 'nomiva
End Sub

Private Sub frmTMer_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento tipos de mercado
    txtAux(10).Text = RecuperaValor(CadenaSeleccion, 1) 'tipo de mercado
    FormateaCampo txtAux(10)
    txtAux2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre de mercado
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(indice).Text = vCampo
End Sub

Private Sub imgAyuda_Click(Index As Integer)
    TerminaBloquear
    
    Select Case Index
        Case 0 'destinos del cliente
            Set frmDest = New frmDestCli
            frmDest.Cliente = Text1(0)
            frmDest.DatosADevolverBusqueda = "0|1|"
            frmDest.Show vbModal
            Set frmDest = Nothing
            PonerFoco Text1(indice)
    End Select
        
End Sub

Private Sub imgFec_Click(Index As Integer)
       
       Screen.MousePointer = vbHourglass
       
       Dim esq As Long
       Dim dalt As Long
       Dim menu As Long
       Dim obj As Object
    
       Set frmC1 = New frmCal
        
       esq = imgFec(Index).Left
       dalt = imgFec(Index).Top
        
       Set obj = imgFec(Index).Container
    
       While imgFec(Index).Parent.Name <> obj.Name
            esq = esq + obj.Left
            dalt = dalt + obj.Top
            Set obj = obj.Container
       Wend
        
       menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar
    
       frmC1.Left = esq + imgFec(Index).Parent.Left + 30
       frmC1.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40
    
       
       frmC1.NovaData = Now
       
       indice = Index
       
       Me.imgFec(0).Tag = indice
       
       PonerFormatoFecha Text3(indice)
       If Text3(indice).Text <> "" Then frmC1.NovaData = CDate(Text3(indice).Text)
    
       Screen.MousePointer = vbDefault
       frmC1.Show vbModal
       Set frmC1 = Nothing
       PonerFoco Text3(indice)
           
      'Para la fecha de la navegacion
       If Index = 0 And Text3(0).Text <> "" Then
            imgFec(0).Tag = Text3(0).Text
            CargaDatosLW
       End If
    
End Sub

Private Sub imgMail_Click(Index As Integer)
'Abrir Outlook para enviar e-mail
Dim dirMail As String

'    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 0: dirMail = Text1(14).Text
        Case 1: dirMail = Text1(20).Text
        Case 2: dirMail = txtAux(8).Text
    End Select

    If LanzaMailGnral(dirMail) Then espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        indice = 21
        frmZ.pTitulo = "Descripción de la venta"
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
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
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
        Case 1  'Nou
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
        Case 5  'Búscar
           mnBuscar_Click
        Case 6  'Tots
            mnVerTodos_Click
        Case 8 'Imprimir
            AbrirListado (10)
    End Select
End Sub

Private Sub BotonBuscar()
Dim i As Integer
Dim anc As Single

' ***** Si la clau primaria de la capçalera no es Text1(0), canviar-ho en <=== *****
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        
        'poner los txtaux para buscar por lineas de albaran
        anc = DataGridAux(0).Top
        If DataGridAux(0).Row < 0 Then
            anc = anc + 210 '440
        Else
            anc = anc + DataGridAux(0).RowTop(DataGridAux(0).Row) + 20
        End If
        LLamaLineas 0, Modo, anc
        
        
        PonerFoco Text1(0) ' <===
        Text1(0).BackColor = vbYellow ' <===
        ' *** si n'hi han combos a la capçalera ***
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
' ******************************************************************************
End Sub

Private Sub HacerBusqueda()

    CadB = ObtenerBusqueda2(Me, BuscaChekc, 0)
    
'    If txtAux(15).Text <> "" Then CadB = CadB & " and destinos.coddesti = " & DBSet(txtAux(15).Text, "N")
    
    If chkVistaPrevia(0) = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        CadenaConsulta = "select distinct clientes.* from " & NombreTabla & " left join destinos on clientes.codclien = destinos.codclien WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        ' *** foco al 1r camp visible de la capçalera que siga clau primaria ***
        PonerFoco Text1(0)
        ' **********************************************************************
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
    Dim Cad As String
        
    'Cridem al form
    ' **************** arreglar-ho per a vore lo que es desije ****************
    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
    Cad = ""  'Albaran|rhisfruta.numalbar|N||11·
    Cad = Cad & "Nombre|clientes.nomclien|N||45·" 'ParaGrid(Text1(2), 45, "Nombre")
    Cad = Cad & "Cód.|clientes.codclien|N||10·" 'ParaGrid(Text1(0), 10, "Cód.")
    Cad = Cad & "NIF.|clientes.cifclien|N||15·" 'ParaGrid(Text1(3), 15, "NIF")
    Cad = Cad & "Teléfono|clientes.telclie1|N||15·" 'ParaGrid(Text1(11), 15, "Teléfono")
    Cad = Cad & "Móvil|clientes.movclie1|N||15·" 'ParaGrid(Text1(12), 15, "Móvil")
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vtabla = NombreTabla & " left join destinos on clientes.codclien = destinos.codclien "
        If CadB = "" Then
            frmB.vSQL = CadB & "(1=1)  group by 1,2,3,4,5"
        Else
            frmB.vSQL = CadB & "  group by 1,2,3,4,5"
        End If
        HaDevueltoDatos = False
        frmB.vDevuelve = "1|" '*** els camps que volen que torne ***
        frmB.vTitulo = "Clientes" ' ***** repasa açò: títol de BuscaGrid *****
        frmB.vSelElem = 0

        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha posat valors i tenim que es formulari de búsqueda llavors
        'tindrem que tancar el form llançant l'event
        If HaDevueltoDatos Then
            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                cmdRegresar_Click
        Else   'de ha retornat datos, es a decir NO ha retornat datos
            PonerFoco Text1(kCampo)
        End If
    End If
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
        LLamaLineas 0, 0
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
        ' *** canviar o llevar, si cal, el WHERE; repasar codEmpre ***
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        'CadenaConsulta = "Select * from " & NombreTabla & " where codempre = " & codEmpre & Ordenacion
        ' ******************************************
        PonerCadenaBusqueda
        ' *** si n'hi han llínies sense grids ***
'        CargaFrame 0, True
        ' ************************************
    End If
End Sub


Private Sub BotonAnyadir()

    LimpiarCampos 'Huida els TextBox
    PonerModo 3
    
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la capçalera *******
    Text1(0).Text = SugerirCodigoSiguienteStr("clientes", "codclien")
    FormateaCampo Text1(0)
       
'    text2(3).Text = Format(Now, "dd/mm/yyyy") ' Quan afegixc pose en F.Alta i F.Modificación la data actual
'    text2(4).Text = Format(Now, "dd/mm/yyyy")
'yo
'    text2(9).Text = vEmpresa.codgrupo  ' Quan afegixc, pose en els camps de l'empleat els valors del que ha fet login
'    text2(11).Text = vSesion.Empresa
'    text2(13).Text = vSesion.Agencia
'    text2(15).Text = vSesion.Empleado
    
    PonerFoco Text1(0) '*** 1r camp visible que siga PK ***
    ' ***********************************************************
    
    ' *** si n'hi han camps de descripció a la capçalera ***
    PosarDescripcions
    ' ******************************************************

    '[Monica]04/07/2012: traemos el tipo de movimiento de albaranes de parametros
    Text1(41).Text = vParamAplic.CodTipomAlb

    ' *** si n'hi han tabs, em posicione al 1r ***
    Me.SSTab1.Tab = 0
    ' ********************************************
End Sub


Private Sub BotonModificar()

    PonerModo 4

    '[Monica]10/07/2013:me guardo los valores de nombre y CCC por si cambian
    NombreAnt = Text1(2).Text
    IbanAnt = Text1(42).Text
    BancoAnt = Text1(1).Text
    SucurAnt = Text1(28).Text
    DigitoAnt = Text1(29).Text
    CuentaAnt = Text1(30).Text
    
    DirecAnt = Text1(4).Text
    cPostalAnt = Text1(5).Text
    PoblaAnt = Text1(18).Text
    ProviAnt = Text1(22).Text
    NifAnt = Text1(3).Text
    EMaiAnt = Text1(14).Text
    WebAnt = Text1(9).Text
    
    '[Monica]26/03/2015: antes no se grababa la forma de pago en la cuenta de cliente
    forpaant = Text1(27).Text


    ' *** bloquejar els camps visibles de la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    ' *************************************************************************
    
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco Text1(2)
    ' *********************************************************
End Sub


Private Sub BotonEliminar()
Dim Cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************

    ' *************** canviar la pregunta ****************
    Cad = "¿Seguro que desea eliminar el Cliente?"
    Cad = Cad & vbCrLf & "Código: " & Format(Data1.Recordset.Fields(0), FormatoCampo(Text1(0)))
    Cad = Cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(1)
    ' **************************************************************************
    
    'borrem
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ' **** repasar si es diu Data1 l'adodc de la capçalera ***
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
        ' ********************************************************
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Cliente", Err.Description
End Sub


Private Sub PonerCampos()
Dim i As Integer
Dim codPobla As String, desPobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la capçalera
    
    ' *** si n'hi han llínies en datagrids ***
'    CargaGrid 0, True
'    CargaGrid 1, True
    
    For i = 0 To 1
        CargaGrid i, True
        
        If Not AdoAux(i).Recordset.EOF Then _
            PonerCamposForma2 Me, AdoAux(i), 2, "FrameAux" & i
            
        If i = 0 Then
            If Not AdoAux(i).Recordset.EOF Then
                Me.imgAyuda(0).visible = True
                Me.imgAyuda(0).Enabled = True
            Else
                Me.imgAyuda(0).visible = False
                Me.imgAyuda(0).Enabled = False
            End If
        
        End If
    Next i
    
    PosarDescripcions

    lblIndicador.Caption = "Datos navegacion"
    Me.Refresh
    DoEvents
    CargaDatosLW
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
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
                LLamaLineas 0, 0
                ' *** foco al primer camp visible de la capçalera ***
                PonerFoco Text1(0)
                ' ***************************************************

        Case 4  'Modificar
                TerminaBloquear
                PonerModo 2
                PonerCampos
                ' *** primer camp visible de la capçalera ***
                PonerFoco Text1(0)
                ' *******************************************
        
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 1 'afegir llínia
                    ModoLineas = 0
                    ' *** les llínies que tenen datagrid (en o sense tab) ***
                    If NumTabMto = 0 Or NumTabMto = 1 Or NumTabMto = 2 Or NumTabMto = 4 Then
                        DataGridAux(NumTabMto).AllowAddNew = False
                        ' **** repasar si es diu Data1 l'adodc de la capçalera ***
                        'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'Modificar
                        ' ********************************************************
                        LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                        DataGridAux(NumTabMto).Enabled = True
                        DataGridAux(NumTabMto).SetFocus

                        ' *** si n'hi han camps de descripció dins del grid, els neteje ***
                        'txtAux2(2).text = ""
                        ' *****************************************************************

                        ' ***  bloquejar i huidar els camps que estan fora del datagrid ***
                        Select Case NumTabMto
                            Case 0 'cuentas bancarias
                                'BotonModificar
'                                BloquearTxt txtaux(11), True
'                                BloquearTxt txtaux(12), True
                            Case 1 'departamentos
'                                For i = 21 To 24
'                                    txtAux(i).Text = ""
'                                    BloquearTxt txtAux(i), True
'                                Next i
'                                txtAux2(22).Text = ""
                            Case 2 'tarjetas
                                BloquearTxt txtAux(50), True
                                BloquearTxt txtAux(51), True
                        End Select
                    ' *** els tabs que no tenen datagrid ***
                    ElseIf NumTabMto = 3 Then
                        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'Modificar
                        CargaFrame 3, True
                    End If
                    
                    ' *** si n'hi han tabs ***
                    SituarTab (NumTabMto)
                    'SSTab1.Tab = 1
                    'SSTab2.Tab = NumTabMto
                    ' ************************
                    
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        AdoAux(NumTabMto).Recordset.MoveFirst
                    End If

                Case 2 'modificar llínies
                    ModoLineas = 0
                    
                    ' *** si n'hi han tabs ***
                    SituarTab (NumTabMto)
                    'SSTab1.Tab = 1
                    'SSTab2.Tab = NumTabMto
                    ' ***********************
                    
                    PonerModo 4
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        ' *** l'Index de Fields es el que canvie de la PK de llínies ***
                        V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
                        AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
                        ' ***************************************************************
                    End If

                    ' ***  bloquejar els camps fora dels grids ***
'                    Select Case NumTabMto
'                        Case 0 'Cuentas bancarias
'                            BloquearTxt txtAux(11), True
'                            BloquearTxt txtAux(12), True
'                        Case 1 'departamentos
'                            For i = 21 To 24
'                                BloquearTxt txtAux(i), True
'                            Next i
'                        Case 2 'Tarjetas
'                            BloquearTxt txtAux(50), True
'                            BloquearTxt txtAux(51), True
'                    ' *** si n'hi han tabs sense datagrid ********
'                        Case 3 'datos facturación
'                            CargaFrame NumTabMto, True 'he cancelat, recarregue els valors
'                    End Select
                    ' ***  bloquejar els camps fora dels grids ***
                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
            End Select
            
            PosicionarData
            
            ' *** si n'hi han llínies en grids i camps fora d'estos ***
            If Not AdoAux(NumTabMto).Recordset.EOF Then
                DataGridAux_RowColChange NumTabMto, 1, 1
            Else
                LimpiarCamposFrame NumTabMto
            End If
            ' *********************************************************
    End Select
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim SQL As String
'Dim Datos As String
Dim Nregs As Long

Dim cta As String
Dim cadMen As String


    On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm2(Me, 1)
    If Not b Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
    If (Modo = 3) Then 'insertar
        'comprobar si existe ya el cod. del campo clave primaria
        If ExisteCP(Text1(0)) Then b = False
    End If
    
    
    If b And (Modo = 3 Or Modo = 4) Then
        
        
        '[Monica]22/08/2013: añadida la comprobacion de que la cuenta contable sea correcta
        If Text1(1).Text = "" Or Text1(28).Text = "" Or Text1(29).Text = "" Or Text1(30).Text = "" Then
            '[Monica]20/11/2013: añadido el codigo de iban
            Text1(42).Text = ""
            Text1(1).Text = ""
            Text1(28).Text = ""
            Text1(29).Text = ""
            Text1(30).Text = ""
        Else
            cta = Format(Text1(1).Text, "0000") & Format(Text1(28).Text, "0000") & Format(Text1(29).Text, "00") & Format(Text1(30).Text, "0000000000")
            If Val(ComprobarCero(cta)) = 0 Then
                cadMen = "El cliente no tiene asignada cuenta bancaria."
                MsgBox cadMen, vbExclamation
            End If
            If Not Comprueba_CC(cta) Then
                cadMen = "La cuenta bancaria del cliente no es correcta. ¿ Desea continuar ?."
                If MsgBox(cadMen, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    b = True
                Else
                    PonerFoco Text1(1)
                    b = False
                End If
            Else
'                '[Monica]20/11/2013: añadimos el tema de la comprobacion del IBAN
'                If Not Comprueba_CC_IBAN(cta, Text1(42).Text) Then
'                    cadMen = "La cuenta IBAN del cliente no es correcta. ¿ Desea continuar ?."
'                    If MsgBox(cadMen, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
'                        b = True
'                    Else
'                        PonerFoco Text1(42)
'                        b = False
'                    End If
'                End If

'       sustituido por lo de David
                BuscaChekc = ""
                If Me.Text1(42).Text <> "" Then BuscaChekc = Mid(Text1(42).Text, 1, 2)
                    
                If DevuelveIBAN2(BuscaChekc, cta, cta) Then
                    If Me.Text1(42).Text = "" Then
                        If MsgBox("Poner IBAN ?", vbQuestion + vbYesNo) = vbYes Then Me.Text1(42).Text = BuscaChekc & cta
                    Else
                        If Mid(Text1(42).Text, 3) <> cta Then
                            cta = "Calculado : " & BuscaChekc & cta
                            cta = "Introducido: " & Me.Text1(42).Text & vbCrLf & cta & vbCrLf
                            cta = "Error en codigo IBAN" & vbCrLf & cta & "Continuar?"
                            If MsgBox(cta, vbQuestion + vbYesNo) = vbNo Then
                                PonerFoco Text1(42)
                                b = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        
        
        If chkAbonos(1).Value = 0 Then ' contador automatico de factura
            If Text1(25).Text = "" Then
                MsgBox "Si no tiene contador manual de factura debe introducir un valor en Tipo de Movimiento.", vbExclamation
                PonerFoco Text1(25)
                b = False
            Else
                ' comprobamos que existe el tipo de movimiento
                SQL = ""
'--monica:10/02/2009: cambiado por el stipom
'                SQL = DevuelveDesdeBDNew(cAgro, "stipom", "codtipom", "codtipom", Text1(25).Text, "T")
'                If SQL = "" Then
'++monica:10/02/2009
                Nregs = TotalRegistros("select count(*) from usuarios.stipom where codtipom = " & DBSet(Text1(25).Text, "T"))
                If Nregs = 0 Then
'++
                    MsgBox "No existe el Tipo de Movimiento. Reintroduzca.", vbExclamation
                    Text1(25).Text = ""
                    PonerFoco Text1(25)
                    b = False
                End If
            End If
        End If
        '[Monica]29/06/2012: comprobamos el tipo de movimiento de los albaranes de venta
        If b Then
            If Not (Text1(41).Text >= "AL1" And Text1(41).Text <= "AL9") And Text1(41).Text <> "ALV" Then
                MsgBox "Este Tipo de Movimiento no se corresponde con el de Albarán de Ventas. Reintroduzca.", vbExclamation
                PonerFoco Text1(41)
                b = False
            Else
                Nregs = TotalRegistros("select count(*) from usuarios.stipom where codtipom = " & DBSet(Text1(41).Text, "T"))
                If Nregs = 0 Then
                    If MsgBox("No existe el Tipo de Movimiento. ¿ Desea crearlo ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                        b = InsertarMovimientoTIPOM(Text1(41).Text, "Albarán Venta")
                    Else
                        MsgBox "No se ha creado el tipo de movimiento. Reintroduzca.", vbExclamation
                        PonerFoco Text1(41)
                        b = False
                    End If
                End If
            End If
        
        End If
    
    End If
    ' ************************************************************************************
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function InsertarMovimientoTIPOM(TipoM As String, Descripc As String) As Boolean
Dim SQL As String

    On Error GoTo eInsertarMovimientoTIPOM

    InsertarMovimientoTIPOM = False

    SQL = "insert into usuarios.stipom (codtipom,nomtipom,muevesto,contador,letraser,tipodocu) values ("
    SQL = SQL & DBSet(TipoM, "T") & "," & DBSet(Descripc, "T") & ",0,0," & ValorNulo & ",0)"
    
    conn.Execute SQL
    
    MsgBox "El tipo de Movimiento ha sido creado correctamente.", vbExclamation
    
    InsertarMovimientoTIPOM = True
    Exit Function
    
    
eInsertarMovimientoTIPOM:
    MuestraError Err.Number, "Insertar Movimiento Usuarios", Err.Description
End Function

                        
Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    Cad = "(codclien=" & Text1(0).Text & ")"
    ' ***************************************
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(Data1, cad, Indicador) Then
    If SituarData(Data1, Cad, Indicador) Then
        If ModoLineas <> 1 Then PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
    ' ***********************************************************************************
End Sub


Private Function Eliminar() As Boolean
Dim vWhere As String

    On Error GoTo FinEliminar

    conn.BeginTrans
    ' ***** canviar el nom de la PK de la capçalera, repasar codEmpre *******
    vWhere = " WHERE codclien=" & Data1.Recordset!CodClien
        ' ***********************************************************************
        
    ' ***** elimina les llínies ****
    conn.Execute "DELETE FROM destinos " & vWhere
        
    ' *******************************
        
    'Eliminar la CAPÇALERA
    vWhere = " WHERE codclien=" & Data1.Recordset!CodClien
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

Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
Dim SQL As String

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    Select Case Index
        Case 0 'cod cliente
            PonerFormatoEntero Text1(0)
            
            '[Monica]07/06/2013: Si estamos insertando comprobamos que el codigo que meten no exista
            If Modo = 3 Then
                If Text1(0).Text <> "" Then
                    SQL = "select count(*) from clientes where codclien = " & DBSet(Text1(0).Text, "N")
                    If TotalRegistros(SQL) <> 0 Then
                        MsgBox "Código ya existe. Reintroduzca.", vbExclamation
                        PonerFoco Text1(0)
                    End If
                End If
            End If

        Case 2 'NOMBRE
            Text1(Index).Text = UCase(Text1(Index).Text)
        
        Case 3 'NIF
            Text1(Index).Text = UCase(Text1(Index).Text)
            If Modo <> 1 Then ValidarNIF Text1(Index).Text
                
                
        Case 6 'PAIS
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "paises", "nompaise")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el País: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmPais = New frmManPaises
                        frmPais.DatosADevolverBusqueda = "0|1|"
                        frmPais.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmPais.Show vbModal
                        Set frmPais = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            

        
        Case 27 'FORMA DE PAGO
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "forpago", "nomforpa")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Forma de Pago: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmFPa = New frmManFpago
                        frmFPa.DatosADevolverBusqueda = "0|1|"
                        frmFPa.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmFPa.Show vbModal
                        Set frmFPa = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
'        Case 25 ' Tipo de Iva
'            If PonerFormatoEntero(Text1(Index)) Then
'                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "tiposiva", "porceiva", "codigiva", "N", cConta)
'                If Text2(Index).Text = "" Then
'                    MsgBox "No existe el Tipo de Iva. Reintroduzca.", vbExclamation
'                    Text1(Index).Text = ""
'                    PonerFoco Text1(Index)
'                End If
'            End If
            
        
        Case 999 'Fechas
            PonerFormatoFecha Text1(Index)
            
        Case 24 'cuenta contable
            If Text1(Index).Text = "" Then Exit Sub
            If Modo = 3 Then ' si estamos insertando puede que no tengamos todos los datos para
                             ' insertar la cuenta contable en contabilidad
                             ' cuando demos aceptar si no existe preguntamos si crear
                Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo, "") ' Text1(0).Text)
            Else
                Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo, Text1(0).Text)
            End If
            
        Case 23, 26 'porcentajes de comision
            cadMen = TransformaPuntosComas(Text1(Index).Text)
            Text1(Index).Text = Format(cadMen, "##0.00")
            
        Case 25 'tipo de movimiento
            If Text1(Index).Text <> "" Then Text1(Index).Text = UCase(Text1(Index).Text)
          
        Case 41 ' tipo de movimiento de albaranes
            If Modo = 1 Then Exit Sub
            If Text1(Index).Text <> "" Then Text1(Index).Text = UCase(Text1(Index).Text)
            If Not (Text1(Index).Text >= "AL1" And Text1(Index).Text <= "AL9") And Text1(Index).Text <> "ALV" Then
                MsgBox "El Tipo de Movimiento de albarán debe estar entre AL1 y AL9 o ser ALV." & vbCrLf & vbCrLf & "Revise.", vbExclamation
                PonerFoco Text1(Index)
            End If
          
        Case 1, 28 'ENTIDAD Y SUCURSAL BANCARIA
            PonerFormatoEntero Text1(Index)
          
        Case 32 'NRO DE COPIAS DE ALBARAN
            PonerFormatoEntero Text1(Index)
            
        Case 33 'NRO DE DECIMALES PRECIO
            PonerFormatoEntero Text1(Index)
    
        Case 35 ' Dias asegurados
            PonerFormatoEntero Text1(Index)
            
        Case 36 ' limite de riesgos
            PonerFormatoDecimal Text1(36), 9
    
    
        '[Monica]04/07/2012: Lo que hay en la direccion de cliente si estamos insertando que pase a la direccion de correo
        Case 4 ' domicilio
            If Modo = 3 Then Text1(40).Text = Text1(Index).Text
            
        Case 5 ' codigo postal
            If Modo = 3 Then Text1(39).Text = Text1(Index).Text
        
        Case 18 ' poblacion
            If Modo = 3 Then Text1(38).Text = Text1(Index).Text
        
        Case 22 ' provincia
            If Modo = 3 Then Text1(37).Text = Text1(Index).Text
    
        Case 42 ' codigo de iban
            Text1(Index).Text = UCase(Text1(Index).Text)
            
    End Select
    
    '[Monica]: calculo del iban si no lo ponen
    If Index = 1 Or Index = 28 Or Index = 29 Or Index = 30 Then
        Dim cta As String
        Dim CC As String
        If Text1(1).Text <> "" And Text1(28).Text <> "" And Text1(29).Text <> "" And Text1(30).Text <> "" Then
            
            cta = Format(Text1(1).Text, "0000") & Format(Text1(28).Text, "0000") & Format(Text1(29).Text, "00") & Format(Text1(30).Text, "0000000000")
            If Len(cta) = 20 Then
    '        Text1(42).Text = Calculo_CC_IBAN(cta, Text1(42).Text)
    
                If Text1(42).Text = "" Then
                    'NO ha puesto IBAN
                    If DevuelveIBAN2("ES", cta, cta) Then Text1(42).Text = "ES" & cta
                Else
                    CC = CStr(Mid(Text1(42).Text, 1, 2))
                    If DevuelveIBAN2(CStr(CC), cta, cta) Then
                        If Mid(Text1(42).Text, 3) <> cta Then
                            
                            MsgBox "Codigo IBAN distinto del calculado [" & CC & cta & "]", vbExclamation
                        End If
                    End If
                End If
                
                
            End If
        End If
    End If
    
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 5: KEYBusqueda KeyAscii, 0 'poblacion
                Case 7: KEYBusqueda KeyAscii, 1 'actividad
                Case 8: KEYBusqueda KeyAscii, 2 'grupo
            End Select
        End If
    Else
        If Index <> 21 Or (Index = 21 And Text1(21).Text = "") Then KEYpress KeyAscii
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

' **** si n'hi han camps de descripció a la capçalera ****
Private Sub PosarDescripcions()
Dim NomEmple As String

    On Error GoTo EPosarDescripcions

    Text2(6).Text = PonerNombreDeCod(Text1(6), "paises", "nompaise", "codpaise", "N")
'    Text2(8).Text = PonerNombreDeCod(Text1(8), "cadenas", "nomcaden")
    Text2(27).Text = PonerNombreDeCod(Text1(27), "forpago", "nomforpa", "codforpa", "N")
    If vParamAplic.NumeroConta <> 0 Then
        Text2(25).Text = PonerNombreDeCod(Text1(25), "tiposiva", "porceiva", "codigiva", "N", cConta)
        Text2(24).Text = PonerNombreCuenta(Text1(24), Modo)
    End If
    
EPosarDescripcions:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo descripciones", Err.Description
End Sub
' ************************************************************


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
Dim SQL As String
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
        Case 0 'destino
            SQL = "¿Seguro que desea eliminar el destino?"
            SQL = SQL & vbCrLf & "Destino: " & AdoAux(Index).Recordset!coddesti & " - " & AdoAux(Index).Recordset!nomdesti
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                SQL = "DELETE FROM destinos"
                SQL = SQL & vWhere & " AND coddesti= " & AdoAux(Index).Recordset!coddesti
            End If
        
        Case 1 'precios
            SQL = "¿Seguro que desea eliminar el articulo?"
            SQL = SQL & vbCrLf & "Artículo: " & AdoAux(Index).Recordset!codArtic & " - " & AdoAux(Index).Recordset!NomArtic
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                SQL = "DELETE FROM clientes_precio"
                SQL = SQL & vWhere & " AND codartic= '" & Trim(AdoAux(Index).Recordset!codArtic) & "'"
            End If
            
    End Select

    If Eliminar Then
        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        conn.Execute SQL
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        If Index <> 3 Then _
            CargaGrid Index, True
            
        ' ***************************************************
        If Not SituarDataTrasEliminar(AdoAux(Index), NumRegElim, True) Then
'            PonerCampos
            
        End If
        ' *** si n'hi han tabs sense datagrid ***
        If Index = 3 Then CargaFrame 3, True
        ' ***************************************
        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
        ' *** si n'hi han tabs ***
        SituarTab (NumTabMto)
        ' ************************
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
    ' **************************************************

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case Index
        Case 0: vtabla = "destinos"
        Case 1: vtabla = "clientes_precio"
    End Select
    ' ********************************************************
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 0, 1 ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            If Index <> 1 Then ' *** els index als que no volem sugerir-li un codi ***
                NumF = SugerirCodigoSiguienteStr(vtabla, "coddesti", vWhere)
            Else
                NumF = ""
            End If
            ' ***************************************************************

            AnyadirLinea DataGridAux(Index), AdoAux(Index)
    
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 240
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
            
            LLamaLineas Index, ModoLineas, anc
        
            Select Case Index
                ' *** valor per defecte a l'insertar i formateig de tots els camps ***
                Case 0 'cuentas
                    For i = 0 To txtAux.Count - 1
                        txtAux(i).Text = ""
                    Next i
                    txtAux(14).Text = Text1(0).Text 'codclien
                    txtAux(15).Text = NumF 'coddesti
                    For i = 0 To 2
                        txtAux2(i).Text = ""
                    Next i
                    
                    BloquearTxt txtAux(15), False
'                    BloquearTxt txtaux(12), False
                    PonerFoco txtAux(15)
                 
                 Case 1 ' clientes_precio
                    For i = 0 To txtAux1.Count - 1
                        txtAux1(i).Text = ""
                    Next i
                    txtAux1(0).Text = Text1(0).Text 'codclien
                    For i = 1 To 2
                        txtAux1(i).Text = ""
                    Next i
                    txtAux2(6).Text = ""
                    
                    PonerFoco txtAux1(1)
                    
            End Select
            
'        ' *** si n'hi han llínies sense datagrid ***
'        Case 3
'            LimpiarCamposLin "FrameAux3"
'            txtaux(42).Text = text1(0).Text 'codclien
'            txtaux(43).Text = vSesion.Empresa
'            Me.cmbAux(28).ListIndex = 0
'            Me.cmbAux(29).ListIndex = 1
'            PonerFoco txtaux(25)
'        ' ******************************************
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
    ' *********************************
  
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
        Case 0 'DESTINOS
            
            txtAux(0).Text = DataGridAux(Index).Columns(2).Text
            txtAux(15).Text = DataGridAux(Index).Columns(1).Text
            txtAux(16).Text = DataGridAux(Index).Columns(3).Text
            txtAux(1).Text = DataGridAux(Index).Columns(4).Text
            txtAux(2).Text = DataGridAux(Index).Columns(5).Text
            txtAux(3).Text = DataGridAux(Index).Columns(6).Text
            
            txtAux(5).Text = DataGridAux(Index).Columns(11).Text
            txtAux(6).Text = DataGridAux(Index).Columns(12).Text
            txtAux(7).Text = DataGridAux(Index).Columns(13).Text
            txtAux(8).Text = DataGridAux(Index).Columns(14).Text
            txtAux(9).Text = DataGridAux(Index).Columns(15).Text
            
            
            txtAux(4).Text = DataGridAux(Index).Columns(7).Text 'PAIS
            txtAux2(2).Text = DataGridAux(Index).Columns(8).Text 'NOMPAIS
            txtAux(10).Text = DataGridAux(Index).Columns(9).Text 'CODTIMER
            txtAux2(0).Text = DataGridAux(Index).Columns(10).Text 'NOMTIMER
            txtAux(11).Text = DataGridAux(Index).Columns(18).Text 'CODCADENA
            txtAux(12).Text = DataGridAux(Index).Columns(16).Text 'CAJAS
            txtAux(13).Text = DataGridAux(Index).Columns(17).Text 'PALETS
            txtAux(17).Text = DataGridAux(Index).Columns(19).Text 'codigo edi
            
            
            
'            SelComboBool AdoAux(Index).Recordset!codNacio, cmbAux(0)
'            For J = 3 To 7
'                txtAux(J).Text = DataGridAux(Index).Columns(J + 2).Text
'            Next J

                       
            BloquearTxt txtAux(11), False
            BloquearTxt txtAux(12), False
            
        Case 1 ' clientes_precios
            txtAux1(0).Text = DataGridAux(Index).Columns(0).Text
            txtAux1(1).Text = DataGridAux(Index).Columns(1).Text
            txtAux2(6).Text = DataGridAux(Index).Columns(2).Text
            txtAux1(2).Text = DataGridAux(Index).Columns(3).Text
            
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 0 'cuentas bancarias
            PonerFoco txtAux(0)
        Case 1 ' precios articulos
            PonerFoco txtAux1(2)
    End Select
    ' ***************************************************************************************
End Sub

' ***** Si n'hi han combos *****
' per a seleccionar la opcio del combo quan estem modificant; només per a "si" i "no"
'Private Sub SelComboBool(valor As Integer, combo As ComboBox)
'Private Sub SelComboBool(valor, combo As ComboBox)
'    Dim i As Integer
'    Dim j As Integer
'
'    i = valor
'    For j = 0 To combo.ListCount - 1
'        If combo.ItemData(j) = i Then
'            combo.ListIndex = j
'            Exit For
'        End If
'    Next j
'End Sub
' ********************************


Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim b As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    If Index <> 3 Then DeseleccionaGrid DataGridAux(Index)
    ' ***************************************************
       
    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
    Select Case Index
        Case 0 'destinos
             For jj = 0 To 1 '3
                txtAux(jj).visible = b
                txtAux(jj).Top = alto
            Next jj
            txtAux(15).visible = b
            txtAux(15).Top = alto
            txtAux(16).visible = b
            txtAux(16).Top = alto
            
'            txtaux2(0).visible = b
'            txtaux2(0).Top = alto
'
'            btnBuscar(0).visible = b
'            btnBuscar(0).Top = txtaux(3).Top
'            btnBuscar(0).Height = txtaux(3).Height

        Case 1 ' clientes_precio
             For jj = 1 To 2 '3
                txtAux1(jj).visible = b
                txtAux1(jj).Top = alto
            Next jj
            txtAux2(6).visible = b
            txtAux2(6).Top = alto
            
            btnBuscar(0).visible = b
            btnBuscar(0).Top = txtAux1(1).Top
            btnBuscar(0).Height = txtAux1(1).Height
            
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
        Case 0 ' nombre de direccion
            If txtAux(Index).Text <> "" Then txtAux(Index).Text = UCase(txtAux(Index).Text)
        
        Case 4 ' pais
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(Index - 2).Text = PonerNombreDeCod(txtAux(Index), "paises", "nompaise")
                If txtAux2(Index - 2).Text = "" Then
                    cadMen = "No existe el País: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmPais = New frmManPaises
                        frmPais.DatosADevolverBusqueda = "0|1|"
                        frmPais.NuevoCodigo = txtAux(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmPais.Show vbModal
                        Set frmPais = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                End If
            Else
                txtAux2(Index - 2).Text = ""
            End If
        
        Case 10 ' tipo de mercado
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(Index - 10).Text = PonerNombreDeCod(txtAux(Index), "tipomer", "nomtimer")
                If txtAux2(Index - 10).Text = "" Then
                    cadMen = "No existe el Tipo de Mercado: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmTMer = New frmManTipMerc
                        frmTMer.DatosADevolverBusqueda = "0|1|"
                        frmTMer.NuevoCodigo = txtAux(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmTMer.Show vbModal
                        Set frmTMer = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                End If
            Else
                txtAux2(Index - 10).Text = ""
            End If
        
        Case 11 ' cadena
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(Index - 10).Text = PonerNombreDeCod(txtAux(Index), "cadenas", "nomcaden")
                If txtAux2(Index - 10).Text = "" Then
                    cadMen = "No existe la Cadena: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCad = New frmManCadenas
                        frmCad.DatosADevolverBusqueda = "0|1|"
                        frmCad.NuevoCodigo = txtAux(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmCad.Show vbModal
                        Set frmCad = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                End If
            Else
                txtAux2(Index - 10).Text = ""
            End If
        
'        Case 12 'cajas
'            If txtAux(Index) <> "" Then PonerFormatoEntero txtAux(Index)
'        Case 13 'palets
'            If txtAux(Index) <> "" Then PonerFormatoEntero txtAux(Index)
            
        Case 15 'codigo de destino
            PonerFormatoEntero txtAux(Index)
        
        Case 17 'codigo edi
            cmdAceptar.SetFocus
        
            
        
        
    End Select
    
    ' ******************************************************************************
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
                    Case 4: KEYBusqueda KeyAscii, 7 'pais
                    Case 10: KEYBusqueda KeyAscii, 3 'mercado
                    Case 11: KEYBusqueda KeyAscii, 4 'cadena
                End Select
            End If
        Else
            KEYpress KeyAscii
        End If
    End If
End Sub


Private Function DatosOkLlin(nomFrame As String) As Boolean
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim b As Boolean
Dim Cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte

    DatosOkLlin = True
    
    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False

    b = CompForm2(Me, 2, nomFrame) 'Comprovar formato datos ok
    If Not b Then Exit Function
'
'    ' *** si cal fer atres comprovacions a les llínies (en o sense tab) ***
'    Select Case NumTabMto
'        Case 0  'CUENTAS BANCARIAS
'            SQL = "SELECT COUNT(ctaprpal) FROM cltebanc "
'            SQL = SQL & ObtenerWhereCab(True) & " AND ctaprpal=1"
'            If ModoLineas = 2 Then SQL = SQL & " AND numlinea<> " & AdoAux(NumTabMto).Recordset!numlinea
'            Set RS = New ADODB.Recordset
'            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'            Cant = IIf(Not RS.EOF, RS.Fields(0).Value, 0)
'
'            RS.Close
'            Set RS = Nothing
''yo
''            'no n'hi ha cap conter principal i ha seleccionat que no
''            If (Cant = 0) And (cmbAux(10).ItemData(cmbAux(10).ListIndex) = 0) Then
''                Mens = "Debe una haber una cuenta principal"
''            ElseIf (cmbAux(10).ItemData(cmbAux(10).ListIndex) = 1) And (cmbAux(9).ItemData(cmbAux(9).ListIndex) = 0) Then
''                Mens = "Debe seleccionar que esta cuenta está activa si desea que sea la principal"
''            End If
'
''            'No puede haber más de una cuenta principal
''            If cant > 0 And (cmbAux(10).ItemData(cmbAux(10).ListIndex) = 1) Then
''                Mens = "No puede haber más de una cuenta principal."
''            End If
''yo
''            'No pueden haber registros con el mismo: codbanco-codsucur-digcontr-ctabanc
''            If Mens = "" Then
''                SQL = "SELECT count(codclien) FROM cltebanc "
''                SQL = SQL & " WHERE codclien=" & text1(0).Text & " AND codempre= " & vSesion.Empresa
''                If ModoLineas = 2 Then SQL = SQL & " AND numlinea<> " & AdoAux(NumTabMto).Recordset!numlinea
''                SQL = SQL & " AND codnacio=" & cmbAux(0).ItemData(cmbAux(0).ListIndex)
''                SQL = SQL & " AND codbanco=" & txtaux(3).Text & " AND codsucur=" & txtaux(4).Text
''                SQL = SQL & " AND digcontr='" & txtaux(5).Text & "' AND ctabanco='" & txtaux(6).Text & "'"
''                Set RS = New ADODB.Recordset
''                RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
''                Cant = IIf(Not RS.EOF, RS.Fields(0).Value, 0)
''                If Cant > 0 Then
''                    Mens = "Ya Existe la cuenta bancaria: " & cmbAux(0).List(cmbAux(0).ListIndex) & " - " & txtaux(3).Text & "-" & txtaux(4).Text & "-" & txtaux(5).Text & "-" & txtaux(6).Text
''                End If
''                RS.Close
''                Set RS = Nothing
''            End If
''
''            If Mens <> "" Then
''                Screen.MousePointer = vbNormal
''                MsgBox Mens, vbExclamation
''                DatosOkLlin = False
''                'PonerFoco txtAux(3)
''                Exit Function
''            End If
''
'    End Select
'    ' ******************************************************************************
    DatosOkLlin = b

EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function ActualisaCtaprpal(ByRef NumLinea As Integer)
Dim SQL As String
'yo
'    On Error Resume Next
'    'tot lo que no siga un SELECT no fa falta un Record Set
'    SQL = "UPDATE cltebanc SET ctaprpal = 0"
'    SQL = SQL & " WHERE codclien=" & text1(0).Text & " AND codempre= " & vSesion.Empresa & " AND numlinea<> " & numlinea
'    Conn.Execute SQL
'
'    If Err.Number <> 0 Then Err.Clear
End Function

Private Function SepuedeBorrar(ByRef Index As Integer) As Boolean
    SepuedeBorrar = False
    
'    ' *** si cal comprovar alguna cosa abans de borrar ***
'    Select Case Index
'        Case 0 'cuentas bancarias
'            If AdoAux(Index).Recordset!ctaprpal = 1 Then
'                MsgBox "No puede borrar una Cuenta Principal. Seleccione antes otra cuenta como Principal", vbExclamation
'                Exit Function
'            End If
'    End Select
'    ' ****************************************************
    
    SepuedeBorrar = True
End Function

' *** si n'hi han formularis de buscar codi a les llínies ***
Private Sub imgBuscar_Click(Index As Integer)
    TerminaBloquear
    
    Select Case Index
        Case 0 'pais
            Set frmPais = New frmManPaises
            frmPais.DatosADevolverBusqueda = "0|1|"
            frmPais.CodigoActual = Text1(6).Text
            frmPais.Show vbModal
            Set frmPais = Nothing
            PonerFoco Text1(6)
        
        Case 7 'pais
            Set frmPais2 = New frmManPaises
            frmPais2.DatosADevolverBusqueda = "0|1|"
            frmPais2.CodigoActual = txtAux(4).Text
            frmPais2.Show vbModal
            Set frmPais2 = Nothing
            PonerFoco txtAux(4)
        
        Case 2 'Cuentas Contables (de contabilidad)
            If vParamAplic.NumeroConta = 0 Then Exit Sub
            
            indice = Index + 22
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(indice).Text
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(indice)
        Case 1 'cadenas
            Set frmCad = New frmManCadenas
            frmCad.DatosADevolverBusqueda = "0|1|"
            frmCad.CodigoActual = Text1(8).Text
            frmCad.Show vbModal
            Set frmCad = Nothing
            PonerFoco Text1(8)
       Case 6 'formas de pago
            Set frmFPa = New frmManFpago
            frmFPa.DatosADevolverBusqueda = "0|1|"
            frmFPa.CodigoActual = Text1(27).Text
            frmFPa.Show vbModal
            Set frmFPa = Nothing
            PonerFoco Text1(27)
       Case 5 'tipos de iva
            Set frmTIva = New frmTipIVAConta
            frmTIva.DeConsulta = True
            frmTIva.DatosADevolverBusqueda = "0|1|"
            frmTIva.CodigoActual = Text1(25).Text
            frmTIva.Show vbModal
            Set frmTIva = Nothing
            PonerFoco Text1(25)
       Case 3 'tipos de mercado
            Set frmTMer = New frmManTipMerc
            frmTMer.DatosADevolverBusqueda = "0|1|"
            frmTMer.CodigoActual = txtAux(10).Text
            frmTMer.Show vbModal
            Set frmTMer = Nothing
            PonerFoco txtAux(10)
            
       Case 4 'cadenas
            Set frmCad = New frmManCadenas
            frmCad.DatosADevolverBusqueda = "0|1|"
            frmCad.CodigoActual = txtAux(11).Text
            frmCad.Show vbModal
            Set frmCad = Nothing
            PonerFoco txtAux(11)
       
       Case 8 'codigos ean de ese cliente
            Set frmCEan = New frmCodEAN
            frmCEan.Tipo = 0
'            frmCEan.DatosADevolverBusqueda = "0|1|"
            frmCEan.CodigoActual = CStr(Me.Data1.Recordset!CodClien)
            frmCEan.Show vbModal
            Set frmCEan = Nothing
            
            
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub



Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    Text1(24).Text = RecuperaValor(CadenaSeleccion, 1) 'codiva
    FormateaCampo Text1(24)
    Text2(24).Text = RecuperaValor(CadenaSeleccion, 2) 'nomiva
End Sub


Private Sub imgWeb_Click(Index As Integer)
    'Abrimos el explorador de windows con la pagina Web del cliente
    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
'    If LanzaHome("websoporte") Then espera 2
    If LanzaHomeGnral(Text1(9).Text) Then espera 2
    Screen.MousePointer = vbDefault
End Sub

' *********************************************************************************

Private Sub DataGridAux_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
Dim i As Byte

    If ModoLineas <> 1 Then
        Select Case Index
            Case 0 'destinos
                If DataGridAux(Index).Columns.Count > 2 Then
                    txtAux(14).Text = DataGridAux(Index).Columns(0).Text
                    txtAux(15).Text = DataGridAux(Index).Columns(1).Text
                 
                    txtAux(2).Text = DataGridAux(Index).Columns(5).Text
                    txtAux(3).Text = DataGridAux(Index).Columns(6).Text
                
                    txtAux(4).Text = DataGridAux(Index).Columns(7).Text
                    txtAux(5).Text = DataGridAux(Index).Columns(11).Text
                    txtAux(6).Text = DataGridAux(Index).Columns(13).Text
                    txtAux(7).Text = DataGridAux(Index).Columns(12).Text
                    txtAux(8).Text = DataGridAux(Index).Columns(14).Text
                    txtAux(9).Text = DataGridAux(Index).Columns(15).Text
                    txtAux(10).Text = DataGridAux(Index).Columns(9).Text
                    txtAux(11).Text = DataGridAux(Index).Columns(18).Text
                    txtAux(12).Text = DataGridAux(Index).Columns(16).Text
                    txtAux(13).Text = DataGridAux(Index).Columns(17).Text
                    txtAux(17).Text = DataGridAux(Index).Columns(19).Text
                
                    txtAux2(0).Text = DataGridAux(Index).Columns(10).Text
                    txtAux2(1).Text = PonerNombreDeCod(txtAux(11), "cadenas", "nomcaden", "codcaden", "N")
                    txtAux2(2).Text = DataGridAux(Index).Columns(8).Text
                    
                End If
                
        End Select
        
    Else 'vamos a Insertar
        Select Case Index
            Case 0 'destinos
                    txtAux(5).Text = ""
                    txtAux(6).Text = ""
                    txtAux(7).Text = ""
                    txtAux(8).Text = ""
                    txtAux(9).Text = ""
                    txtAux(10).Text = ""
                    txtAux(11).Text = ""
                    txtAux(12).Text = ""
                    txtAux(13).Text = ""
                    txtAux(17).Text = ""
                
                    txtAux2(0).Text = ""
        End Select
    End If
End Sub

' ***** si n'hi han varios nivells de tabs *****
Private Sub SituarTab(numTab As Integer)
    On Error Resume Next
    
    If numTab = 0 Then
        SSTab1.Tab = 2
    ElseIf numTab = 1 Then
        SSTab1.Tab = 5
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub
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
        ' *** si n'hi han tabs sense datagrids, li pose els valors als camps ***
        If (Index = 3) Then 'datos facturacion
            tip = AdoAux(Index).Recordset!tipclien
            If (tip = 1) Then 'persona
                txtAux2(27).Text = AdoAux(Index).Recordset!ape_raso & "," & AdoAux(Index).Recordset!Nom_Come
            ElseIf (tip = 2) Then 'empresa
                txtAux2(27).Text = AdoAux(Index).Recordset!Nom_Come
            End If
            txtAux2(28).Text = DBLet(AdoAux(Index).Recordset!desforpa, "T")
            txtAux2(29).Text = DBLet(AdoAux(Index).Recordset!desrutas, "T")
            'txtAux2(31).Text = DBLet(AdoAux(Index).Recordset!comision, "T") & " %"
            txtAux2(32).Text = DBLet(AdoAux(Index).Recordset!nomrapel, "T")
            'Descripcion cuentas contables de la Contabilidad
            For i = 35 To 38
                txtAux2(i).Text = PonerNombreDeCod(txtAux(i), "cuentas", "nommacta", "codmacta", , cConta)
            Next i
        End If
        ' ************************************************************************
    Else
        ' *** si n'hi han tabs sense datagrids, li pose els valors als camps ***
        NetejaFrameAux "FrameAux3" 'neteja només lo que te TAG
        txtAux2(0).Text = ""
        txtAux2(1).Text = ""
        
'        txtaux2(27).Text = ""
'        txtaux2(28).Text = ""
'        txtaux2(29).Text = ""
        'txtAux2(31).Text = ""
'        txtaux2(32).Text = ""
'        For i = 35 To 38
'            txtaux2(i).Text = ""
'        Next i
        ' **********************************************************************
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
' ****************************************


Private Sub CargaGrid(Index As Integer, enlaza As Boolean)
Dim b As Boolean
Dim i As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)

    'b = DataGridAux(Index).Enabled
    'DataGridAux(Index).Enabled = False
    
'    AdoAux(Index).ConnectionString = Conn
'    AdoAux(Index).RecordSource = MontaSQLCarga(Index, enlaza)
'    AdoAux(Index).CursorType = adOpenDynamic
'    AdoAux(Index).LockType = adLockPessimistic
'    DataGridAux(Index).ScrollBars = dbgNone
'    AdoAux(Index).Refresh
'    Set DataGridAux(Index).DataSource = AdoAux(Index)
    
'    DataGridAux(Index).AllowRowSizing = False
'    DataGridAux(Index).RowHeight = 290
'    If PrimeraVez Then
'        DataGridAux(Index).ClearFields
'        DataGridAux(Index).ReBind
'        DataGridAux(Index).Refresh
'    End If
'
'    For i = 0 To DataGridAux(Index).Columns.Count - 1
'        DataGridAux(Index).Columns(i).AllowSizing = False
'    Next i
    
    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
    
    
    'DataGridAux(Index).Enabled = b
'    PrimeraVez = False
    
    Select Case Index
        Case 0 'destinos
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;S|txtaux(15)|T|Cód.|600|;" 'codclien,coddesti
            tots = tots & "S|txtaux(0)|T|Nombre|3100|;"
            tots = tots & "S|txtaux(16)|T|Domicilio|3150|;"
            tots = tots & "S|txtaux(1)|T|C.P|900|;"
'            tots = tots & "S|txtaux(2)|T|Poblacion|1100|;"
'            tots = tots & "S|txtaux(3)|T|Provincia|1100|;"
            tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            arregla tots, DataGridAux(Index), Me, 350
        
            DataGridAux(0).Columns(2).Alignment = dbgLeft
            DataGridAux(0).Columns(3).Alignment = dbgLeft
            DataGridAux(0).Columns(4).Alignment = dbgLeft
'            DataGridAux(0).Columns(5).Alignment = dbgLeft
'            DataGridAux(0).Columns(6).Alignment = dbgLeft
        
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            BloquearTxt txtAux(14), Not b
            BloquearTxt txtAux(15), Not b

            If (enlaza = True) And (Not AdoAux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
                txtAux(2).Text = DataGridAux(Index).Columns(5).Text
                txtAux(3).Text = DataGridAux(Index).Columns(6).Text
            
            
                txtAux(4).Text = DataGridAux(Index).Columns(7).Text
                txtAux(5).Text = DataGridAux(Index).Columns(11).Text
                txtAux(6).Text = DataGridAux(Index).Columns(13).Text
                txtAux(7).Text = DataGridAux(Index).Columns(12).Text
                txtAux(8).Text = DataGridAux(Index).Columns(14).Text
                txtAux(9).Text = DataGridAux(Index).Columns(15).Text
                txtAux(10).Text = DataGridAux(Index).Columns(9).Text
                txtAux(11).Text = DataGridAux(Index).Columns(18).Text
                txtAux(12).Text = DataGridAux(Index).Columns(16).Text
                txtAux(13).Text = DataGridAux(Index).Columns(17).Text
            
                txtAux2(0).Text = DataGridAux(Index).Columns(10).Text
                txtAux2(2).Text = DataGridAux(Index).Columns(8).Text
                txtAux2(1).Text = PonerNombreDeCod(txtAux(11), "cadenas", "nomcaden")
            Else
                txtAux(2).Text = ""
                txtAux(3).Text = ""
                For i = 5 To 13
                    txtAux(i).Text = ""
                Next i
                txtAux2(0).Text = ""
                txtAux2(1).Text = ""
            End If
            
            
        Case 1 'clientes_precio
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;S|txtaux1(1)|T|Artículo|1700|;S|btnBuscar(0)|B|||;" 'codclien,codartic
            tots = tots & "S|txtaux2(6)|T|Nombre|4700|;"
            tots = tots & "S|txtaux1(2)|T|Precio|2200|;"
            
            arregla tots, DataGridAux(Index), Me, 350
        
            DataGridAux(0).Columns(1).Alignment = dbgLeft
            DataGridAux(0).Columns(3).Alignment = dbgLeft
        
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            BloquearTxt txtAux(14), Not b
            BloquearTxt txtAux(15), Not b

            If (enlaza = True) And (Not AdoAux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
'                txtAux1(2).Text = DataGridAux(Index).Columns(5).Text
'                txtAux1(3).Text = DataGridAux(Index).Columns(6).Text
'
'
'                txtAux(4).Text = DataGridAux(Index).Columns(7).Text
'                txtAux(5).Text = DataGridAux(Index).Columns(11).Text
'                txtAux(6).Text = DataGridAux(Index).Columns(13).Text
'                txtAux(7).Text = DataGridAux(Index).Columns(12).Text
'                txtAux(8).Text = DataGridAux(Index).Columns(14).Text
'                txtAux(9).Text = DataGridAux(Index).Columns(15).Text
'                txtAux(10).Text = DataGridAux(Index).Columns(9).Text
'                txtAux(11).Text = DataGridAux(Index).Columns(18).Text
'                txtAux(12).Text = DataGridAux(Index).Columns(16).Text
'                txtAux(13).Text = DataGridAux(Index).Columns(17).Text
'
'                txtAux2(0).Text = DataGridAux(Index).Columns(10).Text
'                txtAux2(2).Text = DataGridAux(Index).Columns(8).Text
'                txtAux2(1).Text = PonerNombreDeCod(txtAux(11), "cadenas", "nomcaden")
            Else
'                txtAux(2).Text = ""
'                txtAux(3).Text = ""
'                For i = 5 To 13
'                    txtAux(i).Text = ""
'                Next i
'                txtAux2(0).Text = ""
'                txtAux2(1).Text = ""
            End If
            
            
    End Select
    DataGridAux(Index).ScrollBars = dbgAutomatic
      
    ' **** si n'hi han llínies en grids i camps fora d'estos ****
    If Not AdoAux(Index).Recordset.EOF Then
        DataGridAux_RowColChange Index, 1, 1
    Else
        LimpiarCamposFrame Index
    End If
    ' **********************************************************
      
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
        Case 0: nomFrame = "FrameAux0" 'destinos
        Case 1: nomFrame = "FrameAux1" 'precios
    End Select
    ' ***************************************************************
    
    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomFrame) Then
            ' *** si n'hi ha que fer alguna cosa abas d'insertar
            If NumTabMto = 0 Then
'yo                'si ha seleccionat "cuenta principal", actualise totes les atres a "no"
'                If (cmbAux(10).ItemData(cmbAux(10).ListIndex) = 1) Then
'                    ActualisaCtaprpal (txtaux(2).Text)
'                End If
            End If
            ' *************************************************
            b = BLOQUEADesdeFormulario2(Me, Data1, 1)
            Select Case NumTabMto
                Case 0, 1, 2, 4 ' *** els index de les llinies en grid (en o sense tab) ***
                     CargaGrid NumTabMto, True
                    If b Then BotonAnyadirLinea NumTabMto
                Case 3 ' *** els index dels tabs que NO tenen grid ***
                    CargaFrame 3, True
                    If b Then BotonModificar
                    ModoLineas = 0
'                LLamaLineas NumTabMto, 0
            End Select
           
            SituarTab (NumTabMto)
        End If
    End If
End Sub


Private Sub ModificarLinea()
'Modifica registre en les taules de Llínies
Dim nomFrame As String
Dim V As Integer
Dim Cad As String
    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomFrame = "FrameAux0" 'cuentas Bancarias
        Case 1: nomFrame = "FrameAux1" ' precios de articulos
    End Select
    ' **************************************************************

    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomFrame) Then
            ' *** si cal que fer alguna cosa abas d'insertar ***
            If NumTabMto = 0 Then
'yo                'si ha seleccionat "cuenta principal", actualise totes les atres a "no"
'                If (cmbAux(10).ItemData(cmbAux(10).ListIndex) = 1) Then
'                    ActualisaCtaprpal (txtaux(2).Text)
'                End If
            End If
            ' ******************************************************
            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
            ModoLineas = 0

            If NumTabMto <> 3 Then
                V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
                CargaGrid NumTabMto, True
            End If

            ' *** si n'hi han tabs ***
            SituarTab (NumTabMto)

            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
            If NumTabMto <> 3 Then
                DataGridAux(NumTabMto).SetFocus
                AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
            End If
            ' ***********************************************************

            LLamaLineas NumTabMto, 0
        End If
    End If
        
'    'Cridem al form
'    ' **************** arreglar-ho per a vore lo que es desije ****************
'    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
'    Cad = ""
'    Cad = Cad & ParaGrid(text1(0), 15, "Cód.")
'    Cad = Cad & ParaGrid(text1(2), 60, "Nombre")
'    Cad = Cad & ParaGrid(text1(3), 25, "N.I.F.")
'    If Cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = Cad
'        frmB.vtabla = NombreTabla
'        frmB.vSQL = CadB
'        HaDevueltoDatos = False
'        frmB.vDevuelve = "0|1|2|" '*** els camps que volen que torne ***
'        frmB.vTitulo = "Clientes" ' ***** repasa açò: títol de BuscaGrid *****
'        frmB.vSelElem = 1
'
'        frmB.Show vbModal
'        Set frmB = Nothing
'        'Si ha posat valors i tenim que es formulari de búsqueda llavors
'        'tindrem que tancar el form llançant l'event
'        If HaDevueltoDatos Then
'            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                cmdRegresar_Click
'        Else   'de ha retornat datos, es a decir NO ha retornat datos
'            PonerFoco text1(kCampo)
'        End If
'    End If
End Sub



Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " codclien=" & Val(Text1(0).Text)
    ' *******************************************************
    
    ObtenerWhereCab = vWhere
End Function

' *** neteja els camps dels tabs de grid que
'estan fora d'este, i els camps de descripció ***
Private Sub LimpiarCamposFrame(Index As Integer)
Dim i As Integer
    On Error Resume Next

    Select Case Index
        Case 0 'Destinos
            For i = 0 To txtAux.Count - 1
                txtAux(i).Text = ""
            Next i
        Case 1 ' precios articulos
            For i = 0 To txtAux1.Count - 1
                txtAux1(i).Text = ""
            Next i
    End Select
    
    If Err.Number <> 0 Then Err.Clear
End Sub
' ***********************************************

Private Sub printNou()
    With frmImprimir2
        .cadTabla2 = "clientes"
        .Informe2 = "rClientes.rpt"
        If CadB <> "" Then
            .cadRegSelec = Replace(SQL2SF(CadB), "clientes", "clientes_1")
        Else
            .cadRegSelec = ""
        End If
        .cadRegActua = Replace(POS2SF(Data1, Me), "clientes", "clientes_1")
        .cadTodosReg = ""
        '.OtrosParametros2 = "pEmpresa='" & vEmpresa.NomEmpre & "'|pOrden={clientes.ape_raso}|"
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomempre & "'|"
        .NumeroParametros2 = 1
        .MostrarTree2 = False
        .InfConta2 = False
        .ConSubInforme2 = False
        .SubInformeConta = ""
        .Show vbModal
    End With
End Sub


' ********* si n'hi han combos a la capçalera ************
Private Sub CargaCombo()
Dim Ini As Integer
Dim fin As Integer
Dim i As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For i = 0 To Combo1.Count - 1
        Combo1(i).Clear
    Next i
    
    Combo1(0).AddItem "Aditivo"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Resto"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    
    Combo1(1).AddItem "Normal"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Exento"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    Combo1(1).AddItem "Recargo Equiv."
    Combo1(1).ItemData(Combo1(1).NewIndex) = 2
   
    Combo1(2).AddItem "Cliente"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 0
    Combo1(2).AddItem "Destino"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 1
  
End Sub


'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'
'
'       El listview tendra los datos de albaranes, facturas... que tenga el cliente
'       Con lo cual, a partir de un click tendremos que ser capaces de situarnos en
'       el formulario correspondiente
'
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------


Private Sub ImagenesNavegacion()
    With Me.Toolbar2
        .ImageList = frmPpal.imgListPpal
        .Buttons(1).Image = 4 ' pedidos
        .Buttons(3).Image = 25 ' albaranes de venta
        .Buttons(5).Image = 27 ' albaranes de envases
        .Buttons(7).Image = 23 ' facturas
    End With
    'tenemos un toolbar oculto para que el icono sea de 16
    With Me.Toolbar3
        .ImageList = frmPpal.imgListImages16
        .Buttons(1).Image = 10 'pedidos
        .Buttons(3).Image = 4 ' albaranes de venta
        .Buttons(5).Image = 11 ' albaranes de envases
        .Buttons(7).Image = 3 ' facturas
    End With
    
'    Set lw1.SmallIcons = frmPpal.imgListPpal
    Set lw1.SmallIcons = frmPpal.imgListImages16
    
End Sub


Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    If Button.Tag = "" Then Exit Sub
    Label45.Caption = ""
    'Levantamos todos los botones y dejamos pulsado el de ahora
    For NumRegElim = 1 To Toolbar2.Buttons.Count
        If Toolbar2.Buttons(NumRegElim).Tag <> "" Then
            If Toolbar2.Buttons(NumRegElim).Index <> Button.Index Then Toolbar2.Buttons(NumRegElim).Value = tbrUnpressed
        End If
    Next NumRegElim
    CargaColumnas CByte(Button.Tag)
    
    'Hacemos las acciones
    If Modo = 2 Then CargaDatosLW
End Sub



Private Sub CargaColumnas(OpcionList As Byte)
Dim Columnas As String
Dim Ancho As String
Dim Alinea As String
Dim Formato As String
Dim Ncol As Integer
Dim C As ColumnHeader

    Select Case OpcionList
        
    Case 0
        ' pedidos
        Label45.Caption = "Pedidos"
        Columnas = "Pedido|Fecha|Destino|Matricula|M.Remolque|"
        Ancho = "1450|1450|2800|1600|1600|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|0|"
        'Formatos
        Formato = "|dd/mm/yyyy|0|0|0|"
        Ncol = 5
        
    Case 1
        'Albaranes
        Label45.Caption = "Albaranes Venta"
        Columnas = "Albarán|Fecha|Destino|Matrícula|M.Remolque|"
        Ancho = "1450|1450|2800|1600|1600|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|0|"
        'Formatos
        Formato = "|dd/mm/yyyy|0|0|0|"
        Ncol = 5
    
    Case 2
        'Albaranes Envases
        Label45.Caption = "Albaranes Envases"
        Columnas = "Albarán|Fecha|Destino|Forma Pago|"
        Ancho = "1450|1450|3000|3000|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|"
        'Formatos
        Formato = "|dd/mm/yyyy|0|0|"
        Ncol = 4
    
    Case 3
        'Facturas
        Label45.Caption = "Facturas"
        Columnas = "Tipo|Numero|Fecha|Importe|"
        Ancho = "1500|2500|1700|3200|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|1|"
        'Formatos
        Formato = "|00000000|dd/mm/yyyy|" & FormatoImporte & "|"
        Ncol = 4
    
    End Select
    
    
'    'Fecha incio busquedas
'    Text1(46).Text = Format(imgFecha(3).Tag, "dd/mm/yyyy")
    'Guardo la opcion en el tag
    lw1.Tag = OpcionList & "|" & Ncol & "|"
    
    lw1.ColumnHeaders.Clear
    
    For NumRegElim = 1 To Ncol
         Set C = lw1.ColumnHeaders.Add()
         C.Text = RecuperaValor(Columnas, CInt(NumRegElim))
         C.Width = RecuperaValor(Ancho, CInt(NumRegElim))
         C.Alignment = Val(RecuperaValor(Alinea, CInt(NumRegElim)))
         C.Tag = RecuperaValor(Formato, CInt(NumRegElim))
    Next NumRegElim
End Sub


Private Sub CargaDatosLW()
Dim C As String
Dim bs As Byte
    bs = Screen.MousePointer
    C = Me.lblIndicador.Caption
    lblIndicador.Caption = "Leyendo " & Label45.Caption
    lblIndicador.Refresh
    CargaDatosLW2
    Me.lblIndicador.Caption = C
    Screen.MousePointer = bs
End Sub

Private Sub CargaDatosLW2()
Dim Cad As String
Dim Rs As ADODB.Recordset
Dim IT As ListItem
Dim ElIcono As Integer
Dim GroupBy As String
Dim Orden As String


    On Error GoTo ECargaDatosLW
    
    If Modo <> 2 Then Exit Sub
    
    For NumRegElim = 1 To Toolbar2.Buttons.Count
        If Toolbar2.Buttons(NumRegElim).Value = tbrPressed Then
            ElIcono = Toolbar3.Buttons(NumRegElim).Image
            Exit For
        End If
    Next
    
    
    'Fecha incio busquedas
    Text3(0).Text = Format(imgFec(0).Tag, "dd/mm/yyyy")
    
    
    Select Case CByte(RecuperaValor(lw1.Tag, 1))
    Case 0
        'PEDIDOS DE VENTA
        Cad = "select h.numpedid,h.fechaped,d.nomdesti,h.matriveh,h.matrirem from pedidos h, destinos d where "
        Cad = Cad & " h.codclien = d.codclien and h.coddesti = d.coddesti "
        GroupBy = "1,2,3,4,5 "
        BuscaChekc = "h.fechaped"
        Orden = "h.numpedid"
    
    Case 1
        'ALBARANES DE VENTA
        Cad = "select h.numalbar,h.fechaalb,d.nomdesti,h.matriveh,h.matrirem from albaran h, destinos d where "
        Cad = Cad & " h.codclien = d.codclien and h.coddesti = d.coddesti "
        GroupBy = "1,2,3,4,5 "
        BuscaChekc = "h.fechaalb"
        Orden = "h.numalbar"
    
    Case 2
        'ALBARANES DE ENVASES
        Cad = "select h.numalbar,h.fechaalb,d.nomdesti,f.nomforpa from scaalb h, destinos d, forpago f where "
        Cad = Cad & " h.codclien = d.codclien and h.coddesti = d.coddesti "
        Cad = Cad & " and h.codforpa = f.codforpa "
        GroupBy = "1,2,3,4 "
        BuscaChekc = "h.fechaalb"
        Orden = "h.numalbar"
    
    Case 3
        'FACTURAS
        Cad = "select h.codtipom,h.numfactu,h.fecfactu,h.totalfac from facturas h WHERE 1=1"
        GroupBy = "1,2,3,4 "
        BuscaChekc = "h.fecfactu"
    
    End Select
    
    
    'La fecha
    
    'EL where del codclien
    Cad = Cad & " and h.codclien=" & Data1.Recordset!CodClien
    
    'La fecha
    If BuscaChekc <> "" Then Cad = Cad & " and " & BuscaChekc & " >='" & Format(imgFec(0).Tag, FormatoFecha) & "'"
    
    
    'El group by
    If GroupBy <> "" Then Cad = Cad & " GROUP BY " & GroupBy
    
    'El ORDER BY
'    If CByte(RecuperaValor(lw1.Tag, 1)) = 1 Then BuscaChekc = Orden
    
    'BuscaChekc="" si es la opcion de precios especiales
    Cad = Cad & " ORDER BY " & BuscaChekc & " DESC"
    BuscaChekc = ""
    
    lw1.ListItems.Clear
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Set IT = lw1.ListItems.Add()
        If lw1.ColumnHeaders(1).Tag <> "" Then
            IT.Text = Format(Rs.Fields(0), lw1.ColumnHeaders(1).Tag)
        Else
            IT.Text = Rs.Fields(0)
        End If
        'El resto de cmpos
        For NumRegElim = 2 To CInt(RecuperaValor(lw1.Tag, 2))
            If IsNull(Rs.Fields(NumRegElim - 1)) Then
                IT.SubItems(NumRegElim - 1) = " "
            Else
                If lw1.ColumnHeaders(NumRegElim).Tag <> "" Then
                    IT.SubItems(NumRegElim - 1) = Format(Rs.Fields(NumRegElim - 1), lw1.ColumnHeaders(NumRegElim).Tag)
                Else
                    IT.SubItems(NumRegElim - 1) = Rs.Fields(NumRegElim - 1)
                End If
            End If
        Next
        IT.SmallIcon = ElIcono
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
    Exit Sub
ECargaDatosLW:
    MuestraError Err.Number
    Set Rs = Nothing
    
End Sub


Private Sub lw1_DblClick()
Dim Seleccionado As Long
    
    If Modo <> 2 Then Exit Sub
    If lw1.ListItems.Count = 0 Then Exit Sub
    If lw1.SelectedItem Is Nothing Then Exit Sub

    If Me.DatosADevolverBusqueda <> "" Then
        'De momento NO dejo continuar
        MsgBox "Esta buscando un socio. No puede ver los documentos.", vbExclamation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Llegados aqui
    Select Case CByte(RecuperaValor(lw1.Tag, 1))
        
    Case 0
        'Pedidos
        Set frmPed = New frmVtasPedidos
        frmPed.hcoCodMovim = lw1.SelectedItem.Text
        frmPed.Show vbModal
        Set frmPed = Nothing
        
    Case 1
        'Albaranes de ventas
        Set frmAlb = New frmVtasAlbaranes
        frmAlb.NumAlbar = lw1.SelectedItem.Text
        frmAlb.hcoCodMovim = lw1.SelectedItem.Text
        frmAlb.Show vbModal
        Set frmAlb = Nothing
    
    Case 2
        'Albaranes de envases
        Set frmAlbEnv = New frmVtasAlbEnvases
        frmAlbEnv.Albaran = lw1.SelectedItem.Text
        frmAlbEnv.Show vbModal
        Set frmAlbEnv = Nothing
    
    Case 3
        'Facturas
        Set frmFac = New frmVtasFacturas
        frmFac.hcoCodMovim = lw1.SelectedItem.SubItems(1)
        frmFac.hcoCodTipoM = lw1.SelectedItem.Text
        frmFac.hcoFechaMov = lw1.SelectedItem.SubItems(2)
        frmFac.Show vbModal
        Set frmFac = Nothing
        
    End Select
        
    'Pase lo que pase, por si acaso, cargamos el lw
    lw1.SetFocus
    Seleccionado = lw1.SelectedItem.Index
    CargaDatosLW
    lw1.SelectedItem.Selected = False
    Set lw1.SelectedItem = Nothing
    If lw1.ListItems.Count >= Seleccionado Then
            lw1.ListItems(Seleccionado).Selected = True
            lw1.ListItems(Seleccionado).EnsureVisible
    End If
End Sub




'********************
Private Sub txtAux1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
    
    If Not PerderFocoGnral(txtAux1(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub


    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 1 ' articulo
            If txtAux1(Index) <> "" Then
                txtAux2(6).Text = PonerNombreDeCod(txtAux1(Index), "sartic", "nomartic")
                VisualizaPrecio
                If txtAux2(6).Text = "" Then
                    cadMen = "No existe el Envase: " & txtAux1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmArt = New frmManArtic
                        frmArt.DatosADevolverBusqueda = "0|1|"
                        frmArt.NuevoCodigo = txtAux1(Index).Text
                        txtAux1(Index).Text = ""
                        TerminaBloquear
                        frmArt.Show vbModal
                        Set frmArt = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        txtAux1(Index).Text = ""
                    End If
                    PonerFoco txtAux1(Index)
                End If
            Else
                txtAux2(6).Text = ""
            End If
            
            
        Case 2 'precio
            If PonerFormatoDecimal(txtAux1(Index), 7) Then cmdAceptar.SetFocus
        
    End Select
    
End Sub

Private Sub txtAux1_GotFocus(Index As Integer)
   If Not txtAux1(Index).MultiLine Then ConseguirFocoLin txtAux1(Index)
End Sub

Private Sub TxtAux1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not txtAux1(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub txtAux1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not txtAux1(Index).MultiLine Then
        If KeyAscii = teclaBuscar Then
            If Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) Then
                Select Case Index
                    Case 1: KEYBusqueda KeyAscii, 0 'articulo
                End Select
            End If
        Else
            KEYpress KeyAscii
        End If
    End If
End Sub

Private Sub VisualizaPrecio()
'    Select Case vParamAplic.TipoPrecio
'        Case 0
'            txtAux1(2).Text = DevuelveDesdeBDNew(cAgro, "sartic", "preciomp", "codartic", txtAux1(1), "T")
'        Case 1
'            txtAux1(2).Text = DevuelveDesdeBDNew(cAgro, "sartic", "preciouc", "codartic", txtAux1(1), "T")
'    End Select
'
' Precio de venta
    txtAux1(2).Text = DevuelveDesdeBDNew(cAgro, "sartic", "preciove", "codartic", txtAux1(1), "T")

End Sub




Private Sub ModificarDatosCuentaContable()
Dim SQL As String
Dim Cad As String

    On Error GoTo eModificarDatosCuentaContable



    If Text1(2).Text <> NombreAnt Or Text1(1).Text <> BancoAnt Or Text1(28).Text <> SucurAnt Or Text1(29).Text <> DigitoAnt Or Text1(30).Text <> CuentaAnt Or _
       DirecAnt <> Text1(4).Text Or cPostalAnt <> Text1(5).Text Or PoblaAnt <> Text1(18).Text Or ProviAnt <> Text1(22).Text Or NifAnt <> Text1(3).Text Or _
       forpaant <> Text1(27).Text Or _
       EMaiAnt <> Text1(14).Text Or WebAnt <> Text1(9).Text Or _
       IbanAnt <> Text1(42).Text Then
        
        Cad = "Se han producido cambios en datos del Cliente. " '& vbCrLf
        
'        If NombreAnt <> Text1(2).Text Then Cad = Cad & " Nombre,"
'        If DirecAnt <> Text1(4).Text Then Cad = Cad & " Direccion,"
'        If cPostalAnt <> Text1(5).Text Then Cad = Cad & " CPostal,"
'        If PoblaAnt <> Text1(18).Text Then Cad = Cad & " Población,"
'        If ProviAnt <> Text1(22).Text Then Cad = Cad & " Provincia,"
'        If NifAnt <> Text1(3).Text Then Cad = Cad & " NIF,"
''        If EMaiAnt <> Text1(12).Text Then Cad = Cad & " EMail,"
'        If BancoAnt <> Text1(1).Text Then Cad = Cad & " Banco,"
'        If SucurAnt <> Text1(28).Text Then Cad = Cad & " Sucursal,"
'        If DigitoAnt <> Text1(29).Text Then Cad = Cad & " Dig.Control,"
'        If CuentaAnt <> Text1(30).Text Then Cad = Cad & " Cuenta banco,"
'
'        Cad = Mid(Cad, 1, Len(Cad) - 1)
        
        Cad = Cad & vbCrLf & vbCrLf & "¿ Desea actualizarlos en la Contabilidad ?" & vbCrLf & vbCrLf
        
        If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                        
            SQL = "update cuentas set nommacta = " & DBSet(Trim(Text1(2).Text), "T")
            SQL = SQL & ", razosoci = " & DBSet(Trim(Text1(2).Text), "T")
            SQL = SQL & ", dirdatos = " & DBSet(Trim(Text1(4).Text), "T")
            SQL = SQL & ", codposta = " & DBSet(Trim(Text1(5).Text), "T")
            SQL = SQL & ", despobla = " & DBSet(Trim(Text1(18).Text), "T")
            SQL = SQL & ", desprovi = " & DBSet(Trim(Text1(22).Text), "T")
            SQL = SQL & ", nifdatos = " & DBSet(Trim(Text1(3).Text), "T")
            SQL = SQL & ", maidatos = " & DBSet(Trim(Text1(14).Text), "T")
            SQL = SQL & ", webdatos = " & DBSet(Trim(Text1(9).Text), "T")
            
            '[Monica]26/03/2015: antes no grababamos la forma de pago de la cuenta
            SQL = SQL & ", forpa = " & DBSet(Trim(Text1(27).Text), "N", "S")
            
            
            If vParamAplic.ContabilidadNueva Then
                Dim vIban As String
                Dim LetraPais As String
                
                vIban = MiFormat(Text1(42).Text, "") & MiFormat(Text1(1).Text, "0000") & MiFormat(Text1(28).Text, "0000") & MiFormat(Text1(29).Text, "00") & MiFormat(Text1(30).Text, "0000000000")
                
                '[Monica]08/06/2017: el pais es el del cliente
                If Text1(6).Text = "" Then
                    LetraPais = "ES"
                Else
                    LetraPais = DevuelveDesdeBDNew(cAgro, "paises", "letraspais", "codpaise", Text1(6).Text, "N")
                End If
            
                SQL = SQL & ", iban = " & DBSet(vIban, "T")
                SQL = SQL & ", codpais = " & DBSet(LetraPais, "T")
            Else
                SQL = SQL & ", entidad = " & DBSet(Trim(Text1(1).Text), "T", "S")
                SQL = SQL & ", oficina = " & DBSet(Trim(Text1(28).Text), "T", "S")
                SQL = SQL & ", cc = " & DBSet(Trim(Text1(29).Text), "T", "S")
                SQL = SQL & ", cuentaba = " & DBSet(Trim(Text1(30).Text), "T", "S")
            
                '[Monica]22/11/2013: tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    SQL = SQL & ", iban = " & DBSet(Trim(Text1(42).Text), "T", "S")
                End If
            End If
            SQL = SQL & " where codmacta = " & DBSet(Trim(Text1(24).Text), "T")
                        
            ConnConta.Execute SQL
                        
'            MsgBox "Datos de Cuenta modificados correctamente.", vbExclamation
                        
        End If
    End If
    
    
    '[Monica]30/08/2013: modificamos los datos de tesoreria sobre los cobros y pagos pendientes
    If Text1(1).Text <> BancoAnt Or Text1(28).Text <> SucurAnt Or Text1(29).Text <> DigitoAnt Or Text1(30).Text <> CuentaAnt _
        Or Text1(42).Text <> IbanAnt Or Text1(27).Text <> forpaant Then
        Cad = "Se han producido cambios en la Cta.Bancaria del cliente."
        Cad = Cad & vbCrLf & vbCrLf & "¿ Desea actualizar los Cobros y Pagos pendientes en Tesoreria ?" & vbCrLf & vbCrLf
        
        If HayCobrosPagosPendientes(Text1(24).Text) Then
            If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                If ActualizarCobrosPagosPdtes(Text1(24), Text1(1).Text, Text1(28).Text, Text1(29).Text, Text1(30).Text, Text1(42).Text, Text1(27).Text) Then
'                    MsgBox "Datos en Tesoreria modificados correctamente.", vbExclamation
                End If
            End If
        End If
    End If
    
    Exit Sub
    
eModificarDatosCuentaContable:
    MuestraError Err.Number, "Modificar Datos Cuenta Contable", Err.Description
End Sub


