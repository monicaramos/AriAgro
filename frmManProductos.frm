VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmManProductos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Productos"
   ClientHeight    =   9150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9555
   Icon            =   "frmManProductos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   6
      Left            =   8820
      MaxLength       =   2
      TabIndex        =   7
      Tag             =   "Globalgap|T|S|||productos|globalgap|||"
      Top             =   4920
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3195
      Left            =   90
      TabIndex        =   18
      Top             =   5070
      Width           =   9295
      _ExtentX        =   16404
      _ExtentY        =   5636
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Insecticidas"
      TabPicture(0)   =   "frmManProductos.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FrameAux0"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Herbicidas"
      TabPicture(1)   =   "frmManProductos.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "FrameAux1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame FrameAux1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2520
         Left            =   90
         TabIndex        =   30
         Top             =   390
         Width           =   9055
         Begin VB.CheckBox chkAux1 
            BackColor       =   &H80000005&
            Height          =   255
            Index           =   0
            Left            =   5460
            TabIndex        =   37
            Tag             =   "Es helicida|N|N|0|1|productos_herb|eshelicida|||"
            Top             =   1710
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   1
            Left            =   450
            MaxLength       =   4
            TabIndex        =   32
            Tag             =   "Linea|N|N|||productos_herb|numlinea|0000|S|"
            Text            =   "lin"
            Top             =   1725
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   4
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   35
            Tag             =   "Dosis|T|S|||productos_herb|dosis||N|"
            Text            =   "dosis"
            Top             =   1710
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.TextBox txtAux3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   3
            Left            =   2130
            MaxLength       =   40
            TabIndex        =   34
            Tag             =   "Nombre Comercial|T|N|||productos_herb|nombre||N|"
            Text            =   "nombre"
            Top             =   1710
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   0
            Left            =   45
            MaxLength       =   3
            TabIndex        =   31
            Tag             =   "Código|N|N|0|999|productos_herb|codprodu|000|S|"
            Text            =   "cod"
            Top             =   1725
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtAux3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   2
            Left            =   810
            MaxLength       =   40
            TabIndex        =   33
            Tag             =   "Materia Activa|T|S|||productos_herb|matactiva||N|"
            Text            =   "Materia activa"
            Top             =   1740
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.TextBox txtAux3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   5
            Left            =   4320
            MaxLength       =   40
            TabIndex        =   36
            Tag             =   "Plagas|T|S|||productos_herb|plagas||N|"
            Text            =   "Plagas"
            Top             =   1710
            Visible         =   0   'False
            Width           =   1080
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   1
            Left            =   135
            TabIndex        =   38
            Top             =   225
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   4
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
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Intercalar"
                  Object.Tag             =   "2"
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc AdoAux 
            Height          =   375
            Index           =   1
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
            Bindings        =   "frmManProductos.frx":0044
            Height          =   1700
            Index           =   1
            Left            =   135
            TabIndex        =   39
            Top             =   660
            Width           =   8800
            _ExtentX        =   15531
            _ExtentY        =   2990
            _Version        =   393216
            AllowUpdate     =   0   'False
            BorderStyle     =   0
            HeadLines       =   1
            RowHeight       =   15
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
      Begin VB.Frame FrameAux0 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2520
         Left            =   -74910
         TabIndex        =   19
         Top             =   390
         Width           =   9085
         Begin VB.TextBox txtAux1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   7
            Left            =   6720
            MaxLength       =   6
            TabIndex        =   28
            Tag             =   "CultAu|T|S|||productos_insec|cultivoaut||N|"
            Text            =   "CultAu"
            Top             =   1710
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.TextBox txtAux1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   6
            Left            =   5550
            MaxLength       =   40
            TabIndex        =   27
            Tag             =   "PSDias|N|S|||productos_insec|psdias|##0|N|"
            Text            =   "PSDias"
            Top             =   1710
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.TextBox txtAux1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   5
            Left            =   4320
            MaxLength       =   40
            TabIndex        =   26
            Tag             =   "Plagas|T|S|||productos_insec|plagas||N|"
            Text            =   "Plagas"
            Top             =   1710
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.TextBox txtAux1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   2
            Left            =   840
            MaxLength       =   40
            TabIndex        =   23
            Tag             =   "Materia Activa|T|S|||productos_insec|matactiva||N|"
            Text            =   "Materia activa"
            Top             =   1710
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.TextBox txtAux1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   0
            Left            =   45
            MaxLength       =   3
            TabIndex        =   20
            Tag             =   "Código|N|N|0|999|productos_insec|codprodu|000|S|"
            Text            =   "cod"
            Top             =   1725
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtAux1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   3
            Left            =   1980
            MaxLength       =   40
            TabIndex        =   24
            Tag             =   "Nombre Comercial|T|N|||productos_insec|nombre||N|"
            Text            =   "nombre"
            Top             =   1710
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.TextBox txtAux1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   4
            Left            =   3180
            MaxLength       =   15
            TabIndex        =   25
            Tag             =   "Dosis|T|S|||productos_insec|dosis||N|"
            Text            =   "dosis"
            Top             =   1710
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.TextBox txtAux1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   1
            Left            =   450
            MaxLength       =   4
            TabIndex        =   22
            Tag             =   "Linea|N|N|||productos_insec|numlinea|0000|S|"
            Text            =   "lin"
            Top             =   1725
            Visible         =   0   'False
            Width           =   240
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   0
            Left            =   135
            TabIndex        =   21
            Top             =   225
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   4
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
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Intercalar"
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
            Bindings        =   "frmManProductos.frx":005C
            Height          =   1700
            Index           =   0
            Left            =   135
            TabIndex        =   29
            Top             =   630
            Width           =   8800
            _ExtentX        =   15531
            _ExtentY        =   2990
            _Version        =   393216
            AllowUpdate     =   0   'False
            BorderStyle     =   0
            HeadLines       =   1
            RowHeight       =   15
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
   End
   Begin VB.CheckBox chkAux 
      BackColor       =   &H80000005&
      Height          =   255
      Index           =   0
      Left            =   8490
      TabIndex        =   6
      Tag             =   "Hay precioarroba|N|N|0|1|productos|precioarroba|||"
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   5
      Left            =   7380
      MaxLength       =   5
      TabIndex        =   5
      Tag             =   "Porc.Epig|N|S|||productos|porcepigrafe|##0.00||"
      Top             =   4950
      Width           =   1095
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   4
      Left            =   6270
      MaxLength       =   3
      TabIndex        =   4
      Tag             =   "Gr.Reten|N|S|||productos|gruporeten|000||"
      Top             =   4950
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   5520
      MaxLength       =   3
      TabIndex        =   3
      Tag             =   "CSIGPA|N|S|||productos|codsigpa|000||"
      Top             =   4950
      Width           =   675
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   4140
      TabIndex        =   17
      Top             =   4950
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   300
      Index           =   0
      Left            =   3930
      MaskColor       =   &H00000000&
      TabIndex        =   16
      ToolTipText     =   "Buscar grupo"
      Top             =   4950
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   2490
      MaxLength       =   3
      TabIndex        =   2
      Tag             =   "Codigo Grupo|N|S|||productos|codgrupo|000||"
      Top             =   4950
      Width           =   1395
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7200
      TabIndex        =   8
      Top             =   8475
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8310
      TabIndex        =   9
      Top             =   8475
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   900
      MaxLength       =   15
      TabIndex        =   1
      Tag             =   "Descripción|T|N|||productos|nomprodu|||"
      Top             =   4920
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   60
      MaxLength       =   3
      TabIndex        =   0
      Tag             =   "Código|N|N|0|999|productos|codprodu|000|S|"
      Top             =   4920
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmManProductos.frx":0074
      Height          =   4410
      Left            =   120
      TabIndex        =   12
      Top             =   540
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   7779
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
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
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   8310
      TabIndex        =   15
      Top             =   8490
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   8430
      Width           =   2385
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
         Height          =   255
         Left            =   40
         TabIndex        =   11
         Top             =   240
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   2790
      Top             =   0
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   7590
         TabIndex        =   14
         Top             =   90
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
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
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
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
Attribute VB_Name = "frmManProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: MANOLO  +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-

' **************** PER A QUE FUNCIONE EN UN ATRE MANTENIMENT ********************
' 0. Posar-li l'atribut Datasource a "adodc1" del Datagrid1. Canviar el Caption
'    del formulari
' 1. Canviar els TAGs i els Maxlength de TextAux(0) i TextAux(1)
' 2. En PonerModo(vModo) repasar els indexs del botons, per si es canvien
' 3. En la funció BotonAnyadir() canviar la taula i el camp per a SugerirCodigoSiguientStr
' 4. En la funció BotonBuscar() canviar el nom de la clau primaria
' 5. En la funció BotonEliminar() canviar la pregunta, les descripcions de la
'    variable SQL i el contingut del DELETE
' 6. En la funció PonerLongCampos() posar els camps als que volem canviar el MaxLength quan busquem
' 7. En Form_Load() repasar la barra d'iconos (per si es vol canviar algún) i
'    canviar la consulta per a vore tots els registres
' 8. En Toolbar1_ButtonClick repasar els indexs de cada botó per a que corresponguen
' 9. En la funció CargaGrid canviar l'ORDER BY (normalment per la clau primaria);
'    canviar ademés els noms dels camps, el format i si fa falta la cantitat;
'    repasar els index dels botons modificar i eliminar.
'    NOTA: si en Form_load ya li he posat clausula WHERE, canviar el `WHERE` de
'    `SQL = CadenaConsulta & " WHERE " & vSQL` per un `AND`
' 10. En txtAux_LostFocus canviar el mensage i el format del camp
' 11. En la funció DatosOk() canviar els arguments de DevuelveDesdeBD i el mensage
'    en cas d'error
' 12. En la funció SepuedeBorrar() canviar les comprovacions per a vore si es pot
'    borrar el registre
' *******************************SI N'HI HA COMBO*******************************
' 0. Comprovar que en el SQL de Form_Load() es faça referència a la taula del Combo
' 1. Pegar el Combo1 al  costat dels TextAux. Canviar-li el TAG
' 2. En BotonModificar() canviar el camp del Combo
' 3. En CargaCombo() canviar la consulta i els noms del camps, o posar els valor
'    a ma si no es llig de cap base de datos els valors del Combo

Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'codi per al registe que s'afegix al cridar des d'atre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean

Private CadenaConsulta As String
Private cadB As String

Private WithEvents frmGru As frmManGrupos 'grupos de productos
Attribute frmGru.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'para sacar los productos que vamos a modificar
Attribute frmMens.VB_VarHelpID = -1


Dim Modo As Byte
'----------- MODOS --------------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'--------------------------------------------------
Dim PrimeraVez As Boolean
Dim indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim I As Integer

Private BuscaChekc As String

Dim ModoLineas As Byte
Dim NumTabMto As Integer
Dim NombreTabla As String

Dim Grupo As Integer
Dim NomGrupo As String
Dim CadenaProductos As String
Dim Kaki As Long
Dim LineaIntercalada As Integer

Private Sub PonerModo(vModo As Byte, Optional indFrame As Integer)
Dim b As Boolean

    Modo = vModo
    
    b = (Modo = 2)
    If b Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    BuscaChekc = ""
    
    For I = 0 To txtAux.Count - 1
        txtAux(I).visible = Not b And (Modo <> 5)
    Next I
    
    txtAux2(2).visible = Not b And (Modo <> 5)
    btnBuscar(0).visible = Not b And (Modo <> 5)
    chkAux(0).visible = Not b And (Modo <> 5)

    cmdAceptar.visible = Not b
    cmdCancelar.visible = Not b
    DataGrid1.Enabled = b And (Modo <> 5)
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = b
    
    PonerLongCampos
    PonerModoOpcionesMenu Modo  'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu  'En funcion del usuario
    
    'Si estamos modo Modificar bloquear clave primaria
    BloquearTxt txtAux(0), (Modo = 4)
End Sub


Private Sub PonerModoOpcionesMenu(Modo)
'Activa/Desactiva botones del la toobar y del menu, segun el modo en que estemos
Dim b As Boolean
Dim bAux As Boolean

    b = (Modo = 2) Or (Modo = 0)
    'Busqueda
    Toolbar1.Buttons(2).Enabled = b
    Me.mnBuscar.Enabled = b
    'Ver Todos
    Toolbar1.Buttons(3).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    'Insertar
    Toolbar1.Buttons(6).Enabled = b And Not DeConsulta
    Me.mnNuevo.Enabled = b And Not DeConsulta
    
    b = (b And adodc1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnModificar.Enabled = b
    'Eliminar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnEliminar.Enabled = b
    'Imprimir
    Toolbar1.Buttons(11).Enabled = b
    Me.mnImprimir.Enabled = b
    
    b = (Modo = 4 Or Modo = 2) And Not DeConsulta
    For I = 0 To ToolAux.Count - 1
        If Not Me.AdoAux(I).Recordset Is Nothing Then
            ToolAux(I).Buttons(1).Enabled = b
            If b Then bAux = (b And Me.AdoAux(I).Recordset.RecordCount > 0)
            ToolAux(I).Buttons(2).Enabled = bAux
            ToolAux(I).Buttons(3).Enabled = bAux
            ToolAux(I).Buttons(4).Enabled = b
        End If
    Next I
    
    
    
    
End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    
    CargaGrid 'primer de tot carregue tot el grid
    cadB = ""
    '******************** canviar taula i camp **************************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        NumF = NuevoCodigo
    Else
        NumF = SugerirCodigoSiguienteStr("productos", "codprodu")
    End If
    '********************************************************************
    
    Modo = 3
    PonerModo Modo
    
    
    'Situamos el grid al final
    AnyadirLinea DataGrid1, adodc1
         
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 206
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
    End If
    txtAux(0).Text = NumF
    FormateaCampo txtAux(0)
    For I = 1 To txtAux.Count - 1
        txtAux(I).Text = ""
    Next I
    txtAux2(2).Text = ""
    chkAux(0).Value = 0

    '[Monica]09/12/2013: ponemos el porcentaje de epigrafe a 0, solo lo usa en ppio Bolbaite
    txtAux(5).Text = 0
    
    LLamaLineas anc, 3 'Pone el form en Modo=3, Insertar
       
    'Ponemos el foco
    PonerFoco txtAux(0)
End Sub

Private Sub BotonVerTodos()
    cadB = ""
    CargaGrid ""
    PonerModo 2
End Sub

Private Sub BotonBuscar()
    ' ***************** canviar per la clau primaria ********
    CargaGrid "productos.codprodu = -1"
    CargaGridAux 0, False
    CargaGridAux 1, False
    
    
    '*******************************************************************************
    'Buscar
    For I = 0 To txtAux.Count - 1
        txtAux(I).Text = ""
    Next I
    chkAux(0).Value = 0
    
    For I = 0 To txtAux1.Count - 1
        txtAux1(I).Text = ""
    Next I
    
    For I = 0 To txtAux3.Count - 1
        txtAux3(I).Text = ""
    Next I
    
    
    LLamaLineas DataGrid1.Top + 206, 1 'Pone el form en Modo=1, Buscar
    PonerFoco txtAux(0)
End Sub

Private Sub BotonModificar()
    Dim anc As Single
    Dim I As Integer
    
    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 545
    End If

    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    txtAux(2).Text = DataGrid1.Columns(2).Text
    ' ***** canviar-ho pel nom del camp del combo *********
'    SelComboBool DataGrid1.Columns(2).Text, Combo1(0)
    ' *****************************************************
    ' ### [Monica] 12/09/2006
    txtAux2(2).Text = DataGrid1.Columns(3).Text
    txtAux(3).Text = DataGrid1.Columns(4).Text
    txtAux(4).Text = DataGrid1.Columns(5).Text
    txtAux(5).Text = DataGrid1.Columns(6).Text
    txtAux(6).Text = DataGrid1.Columns(9).Text
    
    Me.chkAux(0).Value = Me.adodc1.Recordset!precioarroba

    LLamaLineas anc, 4 'Pone el form en Modo=4, Modificar
   
    'Como es modificar
    PonerFoco txtAux(1)
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    
    'Fijamos el ancho
    For I = 0 To txtAux.Count - 1
        txtAux(I).Top = alto
    Next I
    
    ' ### [Monica] 12/09/2006
    txtAux2(2).Top = alto
    btnBuscar(0).Top = alto - 15
    Me.chkAux(0).Top = alto

End Sub

Private Sub LLamaLineasAux(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim b As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    DeseleccionaGrid DataGridAux(Index)
       
    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
    Select Case Index
        Case 0 'insecticidas
             For jj = 2 To 7
                txtAux1(jj).visible = b
                txtAux1(jj).Top = alto
            Next jj
            
        Case 1 'herbicidas
            For jj = 2 To 5
                txtAux3(jj).visible = b
                txtAux3(jj).Top = alto
            Next jj
            
            chkAux1(0).visible = b
            chkAux1(0).Top = alto
            
    End Select
End Sub



Private Sub BotonEliminar()
Dim Sql As String
Dim temp As Boolean

    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    
'    If Not SepuedeBorrar Then Exit Sub
        
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCampo(txtAux(0))) Then Exit Sub
    ' ***************************************************************************
    
    '*************** canviar els noms i el DELETE **********************************
    Sql = "¿Seguro que desea eliminar el Producto?"
    Sql = Sql & vbCrLf & "Código: " & adodc1.Recordset.Fields(0)
    Sql = Sql & vbCrLf & "Descripción: " & adodc1.Recordset.Fields(1)
    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        
        CargarUnProducto CLng(adodc1.Recordset!codprodu), "D"
        
        Sql = "Delete from productos_herb where codprodu=" & adodc1.Recordset!codprodu
        conn.Execute Sql
        
        Sql = "Delete from productos_insec where codprodu=" & adodc1.Recordset!codprodu
        conn.Execute Sql
        
        Sql = "Delete from productos where codprodu=" & adodc1.Recordset!codprodu
        conn.Execute Sql
        
        
        CargaGrid cadB
'        If CadB <> "" Then
'            CargaGrid CadB
'            lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
'        Else
'            CargaGrid ""
'            lblIndicador.Caption = ""
'        End If
        
        temp = SituarDataTrasEliminar(adodc1, NumRegElim, True)
        PonerModoOpcionesMenu Modo
        adodc1.Recordset.Cancel
    End If
    
    Exit Sub
    
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los txtAux
    PonerLongCamposGnral Me, Modo, 3
End Sub

Private Sub btnBuscar_Click(Index As Integer)
 TerminaBloquear
    
    Select Case Index
        Case 0 'grupo de productos
            
            indice = Index + 2
            Set frmGru = New frmManGrupos
            frmGru.DatosADevolverBusqueda = "0|1|"
            frmGru.CodigoActual = txtAux(indice).Text
            frmGru.Show vbModal
            Set frmGru = Nothing
            PonerFoco txtAux(indice)
    
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Me.adodc1, 1
End Sub

Private Sub chkAux_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "chkAux(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "chkAux(" & Index & ")|"
    End If
End Sub

Private Sub chkAux1_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "chkAux1(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "chkAux1(" & Index & ")|"
    End If
End Sub


Private Sub cmdAceptar_Click()
    Dim I As Integer

    Select Case Modo
        Case 1 'BUSQUEDA
            cadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
            If cadB <> "" Then
                CargaGrid cadB
                PonerModo 2
'                lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
                PonerFocoGrid Me.DataGrid1
            End If
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    CargarUnProducto CLng(txtAux(0)), "I"
                    CargaGrid
                    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                        cmdCancelar_Click
'                        If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveLast
                        If Not adodc1.Recordset.EOF Then
                            adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & NuevoCodigo)
                        End If
                        cmdRegresar_Click
                    Else
                        BotonAnyadir
                    End If
                    cadB = ""
                End If
            End If
            
        Case 4 'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me) Then
                    CargarUnProducto CLng(txtAux(0)), "U"
                    TerminaBloquear
                    I = adodc1.Recordset.Fields(0)
                    PonerModo 2
                    CargaGrid cadB
'                    If CadB <> "" Then
'                        CargaGrid CadB
'                        lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
'                    Else
'                        CargaGrid
'                        lblIndicador.Caption = ""
'                    End If
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & I)
                    PonerFocoGrid Me.DataGrid1
                End If
            End If
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 1 'afegir llínia
                    InsertarLinea
                Case 2 'modificar llínies
                    If ModificarLinea Then
                        PonerModo 2
                    End If
            End Select
            'nuevo calculamos los totales de lineas
            
    End Select
End Sub

Private Sub cmdCancelar_Click()
Dim V

    On Error Resume Next
    
    Select Case Modo
        Case 1 'búsqueda
            CargaGrid cadB
        Case 3 'insertar
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        Case 4 'modificar
            TerminaBloquear
            
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 1 'afegir llínia
                    ModoLineas = 0
                    ' *** les llínies que tenen datagrid (en o sense tab) ***
                    If NumTabMto = 0 Or NumTabMto = 1 Or NumTabMto = 2 Or NumTabMto = 4 Then
                        DataGridAux(NumTabMto).AllowAddNew = False
                        LLamaLineasAux NumTabMto, ModoLineas 'ocultar txtAux
                        DataGridAux(NumTabMto).Enabled = True
                        DataGridAux(NumTabMto).SetFocus

                    End If
                    
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        AdoAux(NumTabMto).Recordset.MoveFirst
                    End If

                Case 2 'modificar llínies
                    ModoLineas = 0
                    
                    LLamaLineasAux NumTabMto, ModoLineas 'ocultar txtAux
                    PonerModo 4
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        ' *** l'Index de Fields es el que canvie de la PK de llínies ***
                        V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
                        AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
                        ' ***************************************************************
                    End If
            End Select
            
            TerminaBloquear
'            CargaGrid
'            PonerModo 2
            
            ' *** si n'hi han llínies en grids i camps fora d'estos ***
'            If Not AdoAux(NumTabMto).Recordset.EOF Then
'                DataGridAux_RowColChange NumTabMto, 1, 1
'            Else
'                LimpiarCamposFrame NumTabMto
'            End If
            
    End Select
    
    PonerModo 2
    
    PonerFocoGrid Me.DataGrid1
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub cmdRegresar_Click()
Dim cad As String
Dim I As Integer
Dim j As Integer
Dim Aux As String

    If adodc1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    cad = ""
    I = 0
    Do
        j = I + 1
        I = InStr(j, DatosADevolverBusqueda, "|")
        If I > 0 Then
            Aux = Mid(DatosADevolverBusqueda, j, I - j)
            j = Val(Aux)
            cad = cad & adodc1.Recordset.Fields(j) & "|"
        End If
    Loop Until I = 0
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim KK As Integer
    PonerContRegIndicador
    For KK = 0 To 1
        If Modo = 3 Then
            CargaGridAux KK, False
        Else
            CargaGridAux KK, True
        End If
    Next KK
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault

    If PrimeraVez Then
        PrimeraVez = False
        If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
            BotonAnyadir
        Else
            PonerModo 2
             If Me.CodigoActual <> "" Then
                SituarData Me.adodc1, "codprodu=" & CodigoActual, "", True
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    PrimeraVez = True

    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'el 1 es separadors
        .Buttons(2).Image = 1   'Buscar
        .Buttons(3).Image = 2   'Todos
        'el 4 i el 5 son separadors
        .Buttons(6).Image = 3   'Insertar
        .Buttons(7).Image = 4   'Modificar
        .Buttons(8).Image = 5   'Borrar
        'el 9 i el 10 son separadors
        .Buttons(11).Image = 10  'imprimir
        .Buttons(12).Image = 11  'Salir
    End With

    ' ******* si n'hi han llínies *******
    'ICONETS DE LES BARRES ALS TABS DE LLÍNIA
    For I = 0 To ToolAux.Count - 1
        With Me.ToolAux(I)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
            .Buttons(4).Image = 19   'Insertar linea
        End With
    Next I
    ' ***********************************


    
    '****************** canviar la consulta *********************************
    CadenaConsulta = "SELECT productos.codprodu, productos.nomprodu, productos.codgrupo,"
    CadenaConsulta = CadenaConsulta & "grupopro.nomgrupo, productos.codsigpa, productos.gruporeten, "
    CadenaConsulta = CadenaConsulta & " productos.porcepigrafe, precioarroba, IF(precioarroba=1,'*','') as dprecioarroba,  "
    CadenaConsulta = CadenaConsulta & " productos.globalgap "
    CadenaConsulta = CadenaConsulta & " FROM productos, grupopro"
    CadenaConsulta = CadenaConsulta & " WHERE productos.codgrupo = grupopro.codgrupo "
    '************************************************************************
    
    DataGridAux(0).ClearFields
    DataGridAux(1).ClearFields
    
    NombreTabla = "productos"
    
    cadB = ""
    CargaGrid
    PonerCampos

    Me.SSTab1.Tab = 0

       
    ModoLineas = 0
       

'    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
'        BotonAnyadir
'    Else
'        PonerModo 2
'    End If

    ' Para el chivato
    Set dbAriagro = New BaseDatos
    dbAriagro.abrir_MYSQL vConfig.SERVER, vUsu.CadenaConexion, vConfig.User, vConfig.password


End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    If Modo = 4 Then TerminaBloquear
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmgru_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    txtAux(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codgrupo
    txtAux2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre grupo
End Sub


Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
    CadenaProductos = CadenaSeleccion
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    printNou
End Sub

Private Sub mnModificar_Click()
    'Comprobaciones
    '--------------
    If adodc1.Recordset.EOF Then Exit Sub
    
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
    
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCampo(txtAux(0))) Then Exit Sub
    
    
    'Preparamos para modificar
    '-------------------------
    If BLOQUEADesdeFormulario2(Me, adodc1, 1) Then BotonModificar
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

Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
'-- pon el bloqueo aqui
    'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
    LineaIntercalada = -1
    
    Select Case Button.Index
        Case 1
            BotonAnyadirLinea Index
        Case 2
            BotonModificarLinea Index
        Case 3
            BotonEliminarLinea Index
            
        Case 4
            BotonIntercalarLinea Index
        
        Case Else
    End Select
    'End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 2
                mnBuscar_Click
        Case 3
                mnVerTodos_Click
        Case 6
                mnNuevo_Click
        Case 7
                mnModificar_Click
        Case 8
                mnEliminar_Click
        Case 11
                mnImprimir_Click
        Case 12
                mnSalir_Click
    End Select
End Sub

Private Sub CargaGrid(Optional vSQL As String)
    Dim Sql As String
    Dim tots As String
    
'    adodc1.ConnectionString = Conn
    If vSQL <> "" Then
        Sql = CadenaConsulta & " AND " & vSQL
    Else
        Sql = CadenaConsulta
    End If
    '********************* canviar el ORDER BY *********************++
    Sql = Sql & " ORDER BY productos.codprodu"
    '**************************************************************++
    
    CargaGridGnral Me.DataGrid1, Me.adodc1, Sql, PrimeraVez
    
    ' *******************canviar els noms i si fa falta la cantitat********************
    tots = "S|txtAux(0)|T|Cód.|500|;S|txtAux(1)|T|Descripción|2100|;"
    tots = tots & "S|txtAux(2)|T|Grupo|700|;"
    tots = tots & "S|btnBuscar(0)|B|||;S|txtAux2(2)|T|Nombre de Grupo|2200|;"
    tots = tots & "S|txtAux(3)|T|Sigpac|800|;"
    tots = tots & "S|txtAux(4)|T|Gr.Ret|800|;"
    tots = tots & "S|txtAux(5)|T|%Epigr|800|;"
    tots = tots & "N||||0|;S|chkAux(0)|CB|@|360|;"
    tots = tots & "S|txtAux(6)|T|GG|400|;"
    
    
    arregla tots, DataGrid1, Me
    
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.Columns(0).Alignment = dbgRight
'   DataGrid1.Columns(2).Alignment = dbgRight
End Sub


Private Sub CargaGridAux(Index As Integer, enlaza As Boolean)
Dim b As Boolean
Dim I As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)

    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
    
    Select Case Index
        Case 0 'insecticidas
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;" 'codprodu, numlinea
            tots = tots & "S|txtAux1(2)|T|Materia Activa|1700|;S|txtAux1(3)|T|Nombre Comercial|1800|;"
            tots = tots & "S|txtAux1(4)|T|Dosis|900|;"
            tots = tots & "S|txtAux1(5)|T|Plagas|1900|;S|txtAux1(6)|T|PS Dias|900|;S|txtAux1(7)|T|Cult.Aut|980|;"
            
            arregla tots, DataGridAux(Index), Me
        
            DataGridAux(0).Columns(4).Alignment = dbgCenter
            DataGridAux(0).Columns(6).Alignment = dbgCenter
        
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
        Case 1 'herbicidas
            'si es visible|control|tipo campo|nombre campo|ancho control|
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;" 'codprodu, numlinea
            tots = tots & "S|txtAux3(2)|T|Materia Activa|2250|;S|txtAux3(3)|T|Nombre Comercial|2200|;"
            tots = tots & "S|txtAux3(4)|T|Dosis|900|;"
            tots = tots & "S|txtAux3(5)|T|Plagas|1900|;N||||0|;S|chkAux1(0)|CB|Helicida|980|;"
            
            arregla tots, DataGridAux(Index), Me
        
            DataGridAux(1).Columns(4).Alignment = dbgCenter
        
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
    End Select
    
    DataGridAux(Index).ScrollBars = dbgAutomatic
    
    PonerModoOpcionesMenu Modo
      
    ' **** si n'hi han llínies en grids i camps fora d'estos ****
'    If Not AdoAux(Index).Recordset.EOF Then
'        DataGridAux_RowColChange Index, 1, 1
'    Else
''        LimpiarCamposFrame Index
'    End If
      
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub



Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 0
            PonerFormatoEntero txtAux(Index)
        Case 1
            txtAux(Index).Text = UCase(txtAux(Index).Text)
        
        Case 2 'codigo de grupo de productos
            If txtAux(Index).Text = "" Then Exit Sub
            txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), "grupopro", "nomgrupo", "codgrupo", "N")
            
        Case 3, 4 'sigpa y grupo retencion
            PonerFormatoEntero txtAux(Index)
            
        Case 5
            PonerFormatoDecimal txtAux(Index), 4
        
    End Select
    
End Sub


Private Sub txtAux1_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux1(Index)
End Sub


Private Sub txtAux1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux1_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux1(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 6 'PSdias
            PonerFormatoEntero txtAux1(Index)
            
    End Select
    
End Sub

Private Sub TxtAux3_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux3(Index)
End Sub

Private Sub TxtAux3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub TxtAux3_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux3(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 6 'sigpa y grupo retencion
            PonerFormatoEntero txtAux3(Index)
            
    End Select
    
End Sub



Private Function DatosOk() As Boolean
'Dim Datos As String
Dim b As Boolean
Dim Sql As String
Dim Mens As String


    b = CompForm(Me)
    If Not b Then Exit Function
    
    If Modo = 3 Then   'Estamos insertando
         If ExisteCP(txtAux(0)) Then b = False
    End If
    
    DatosOk = b
End Function


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub


Private Sub PonerContRegIndicador()
'si estamos en modo ver registros muestra el numero de registro en el que estamos
'situados del total de registros mostrados: 1 de 24
Dim cadReg As String

    If (Modo = 2 Or Modo = 0) Then
        cadReg = PonerContRegistros(Me.adodc1)
        If cadB = "" Then
            lblIndicador.Caption = cadReg
        Else
            lblIndicador.Caption = "BUSQUEDA: " & cadReg
        End If
    End If
End Sub

Private Sub printNou()
    With frmImprimir2
        .cadTabla2 = "sforpa"
        .Informe2 = "rManProductos.rpt"
        If cadB <> "" Then
            '.cadRegSelec = Replace(SQL2SF(CadB), "clientes", "clientes_1")
            .cadRegSelec = SQL2SF(cadB)
        Else
            .cadRegSelec = ""
        End If
        ' *** repasar el nom de l'adodc ***
        '.cadRegActua = Replace(POS2SF(Data1, Me), "clientes", "clientes_1")
        .cadRegActua = POS2SF(adodc1, Me)
        ' *** repasar codEmpre ***
        .cadTodosReg = ""
        '.cadTodosReg = "{itinerar.codempre} = " & codEmpre
        ' *** repasar si li pose ordre o no ****
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomempre & "'|pOrden={sforpa.codforpa}|"
        '.OtrosParametros2 = "pEmpresa='" & vEmpresa.NomEmpre & "'|"
        ' *** posar el nº de paràmetres que he posat en OtrosParametros2 ***
        '.NumeroParametros2 = 1
        .NumeroParametros2 = 2
        ' ******************************************************************
        .MostrarTree2 = False
        .InfConta2 = False
        .ConSubInforme2 = False
        .SubInformeConta = ""
        .Show vbModal
    End With
End Sub

'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGrid1_GotFocus()
'  WheelHook DataGrid1
'End Sub
'Private Sub DataGrid1_Lostfocus()
'  WheelUnHook
'End Sub
'
'Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
'    KEYpress KeyAscii
'End Sub
Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 2: KEYBusqueda KeyAscii, 0 'cuenta contable
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
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
    btnBuscar_Click (indice)
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
               
        Case 0 'insecticidas
            Sql = "SELECT productos_insec.codprodu, productos_insec.numlinea, productos_insec.matactiva, productos_insec.nombre,"
            Sql = Sql & " productos_insec.dosis, productos_insec.plagas, productos_insec.psdias, "
            Sql = Sql & " productos_insec.cultivoaut "
            Sql = Sql & " FROM productos_insec "
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE productos_insec.codprodu = '-1'"
            End If
            Sql = Sql & " ORDER BY productos_insec.numlinea"
               
        Case 1 'herbicidas
            Sql = "SELECT productos_herb.codprodu, productos_herb.numlinea, productos_herb.matactiva, productos_herb.nombre,"
            Sql = Sql & " productos_herb.dosis, productos_herb.plagas, "
            Sql = Sql & " productos_herb.eshelicida,IF(eshelicida=1,'*','') as deshelicida "
            Sql = Sql & " FROM productos_herb "
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE productos_herb.codprodu = '-1'"
            End If
            Sql = Sql & " ORDER BY productos_herb.numlinea"
            
    End Select
    
    MontaSQLCarga = Sql
End Function

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " codprodu=" & Me.adodc1.Recordset!codprodu
    
    ObtenerWhereCab = vWhere
End Function

Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vtabla As String
Dim anc As Single
Dim I As Integer
    
    ModoLineas = 1 'Posem Modo Afegir Llínia
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    CadenaProductos = ""
    Grupo = DevuelveValor("select codgrupo from productos where codprodu = " & Me.adodc1.Recordset!codprodu)
    
    
'    NomGrupo = DevuelveValor("select nomgrupo from grupopro where codgrupo = " & DBSet(Grupo, "N"))
'
'    Set frmMens = New frmMensajes
'
'    frmMens.OpcionMensaje = 28
'    frmMens.cadWHERE = " where codgrupo = " & Grupo & " and codprodu <> " & adodc1.Recordset!codprodu
'    frmMens.Show vbModal
'
'    Set frmMens = Nothing
    
    
    
    NumTabMto = Index
    PonerModo 5, Index
    
    ' *** bloquejar la clau primaria de la capçalera ***
    BloquearTxt txtAux(0), True

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case Index
        Case 0: vtabla = "productos_insec"
        Case 1: vtabla = "productos_herb"
    End Select
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 0, 1 ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            
            NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", vWhere)

            AnyadirLinea DataGridAux(Index), AdoAux(Index)
    
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
            
            LLamaLineasAux Index, ModoLineas, anc
        
            Select Case Index
                ' *** valor per defecte a l'insertar i formateig de tots els camps ***
                Case 0 'insecticidas
                    txtAux1(0).Text = Me.adodc1.Recordset!codprodu
                    txtAux1(1).Text = NumF 'numlinea
                    txtAux1(2).Text = ""
                    txtAux1(3).Text = ""
                    txtAux1(4).Text = ""
                    txtAux1(5).Text = ""
                    txtAux1(6).Text = ""
                    PonerFoco txtAux1(2)
                Case 1 'herbicidas
                    txtAux3(0).Text = Me.adodc1.Recordset!codprodu 'codprodu
                    txtAux3(1).Text = NumF 'numlinea
                    txtAux3(2).Text = ""
                    txtAux3(3).Text = ""
                    txtAux3(4).Text = ""
                    txtAux3(5).Text = ""
                    PonerFoco txtAux3(2)
            End Select
            
    End Select
End Sub


Private Sub BotonIntercalarLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vtabla As String
Dim anc As Single
Dim I As Integer
    
    ModoLineas = 1 'Posem Modo Afegir Llínia
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    CadenaProductos = ""
    Grupo = DevuelveValor("select codgrupo from productos where codprodu = " & Me.adodc1.Recordset!codprodu)
    
    
'    NomGrupo = DevuelveValor("select nomgrupo from grupopro where codgrupo = " & DBSet(Grupo, "N"))
'
'    Set frmMens = New frmMensajes
'
'    frmMens.OpcionMensaje = 28
'    frmMens.cadWHERE = " where codgrupo = " & Grupo & " and codprodu <> " & adodc1.Recordset!codprodu
'    frmMens.Show vbModal
'
'    Set frmMens = Nothing
    
    
    LineaIntercalada = CInt(Me.AdoAux(Index).Recordset!NumLinea)
       
    NumTabMto = Index
    PonerModo 5, Index
    
    ' *** bloquejar la clau primaria de la capçalera ***
    BloquearTxt txtAux(0), True

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case Index
        Case 0: vtabla = "productos_insec"
        Case 1: vtabla = "productos_herb"
    End Select
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 0, 1 ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            
            NumF = LineaIntercalada 'SugerirCodigoSiguienteStr(vTabla, "numlinea", vWhere)

            AnyadirLinea DataGridAux(Index), AdoAux(Index)
    
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
            
            LLamaLineasAux Index, ModoLineas, anc
        
            Select Case Index
                ' *** valor per defecte a l'insertar i formateig de tots els camps ***
                Case 0 'insecticidas
                    txtAux1(0).Text = Me.adodc1.Recordset!codprodu
                    txtAux1(1).Text = NumF 'numlinea
                    txtAux1(2).Text = ""
                    txtAux1(3).Text = ""
                    txtAux1(4).Text = ""
                    txtAux1(5).Text = ""
                    txtAux1(6).Text = ""
                    PonerFoco txtAux1(2)
                Case 1 'herbicidas
                    txtAux3(0).Text = Me.adodc1.Recordset!codprodu 'codprodu
                    txtAux3(1).Text = NumF 'numlinea
                    txtAux3(2).Text = ""
                    txtAux3(3).Text = ""
                    txtAux3(4).Text = ""
                    txtAux3(5).Text = ""
                    PonerFoco txtAux3(2)
            End Select
            
    End Select
End Sub





Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim I As Integer
    Dim j As Integer
    
    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub
    
    ModoLineas = 2 'Modificar llínia
       
    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    CadenaProductos = ""
    Grupo = DevuelveValor("select codgrupo from productos where codprodu = " & Me.adodc1.Recordset!codprodu)
'    NomGrupo = DevuelveValor("select nomgrupo from grupopro where codgrupo = " & DBSet(Grupo, "N"))
'
'    Set frmMens = New frmMensajes
'
'    frmMens.OpcionMensaje = 28
'    frmMens.cadWHERE = " where codgrupo = " & Grupo & " and codprodu <> " & adodc1.Recordset!codprodu
'    frmMens.Show vbModal
'
'    Set frmMens = Nothing
       
       
    NumTabMto = Index
    PonerModo 5, Index
    ' *** bloqueje la clau primaria de la capçalera ***
    BloquearTxt txtAux(0), True
  
    Select Case Index
        Case 0, 1 ' *** pose els index de llínies que tenen datagrid (en o sense tab) ***
            If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
                I = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
                DataGridAux(Index).Scroll 0, I
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
        Case 0 ' insecticidas
            For j = 0 To 6
                txtAux1(j).Text = DataGridAux(Index).Columns(j).Text
            Next j
            BloquearTxt txtAux(0), True
            BloquearTxt txtAux(3), True
            
            BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux0"
            
        Case 1 ' herbicidas
            For j = 0 To 5
                txtAux3(j).Text = DataGridAux(Index).Columns(j).Text
            Next j
            For I = 9 To 9
                BloquearTxt txtAux(I), True
            Next I
            BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux1"
            
    End Select
    
    LLamaLineasAux Index, ModoLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 0 'insecticidas
            PonerFoco txtAux1(2)
        Case 1 'herbicidas
            PonerFoco txtAux3(2)
    End Select
    ' ***************************************************************************************
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
    
    CadenaProductos = ""
    Grupo = DevuelveValor("select codgrupo from productos where codprodu = " & Me.adodc1.Recordset!codprodu)
'    NomGrupo = DevuelveValor("select nomgrupo from grupopro where codgrupo = " & DBSet(Grupo, "N"))
'
'    Set frmMens = New frmMensajes
'
'    frmMens.OpcionMensaje = 28
'    frmMens.cadWHERE = " where codgrupo = " & Grupo & " and codprodu <> " & adodc1.Recordset!codprodu
'    frmMens.Show vbModal
'
'    Set frmMens = Nothing
'
       
    NumTabMto = Index
    PonerModo 5, Index

    If AdoAux(Index).Recordset.EOF Then Exit Sub
    NumTabMto = Index
    Eliminar = False
   
    vWhere = ObtenerWhereCab(True)
    
    ' ***** independentment de si tenen datagrid o no,
    ' canviar els noms, els formats i el DELETE *****
    Select Case Index
        Case 0 'insecticidas
            Sql = "¿Seguro que desea el insecticida?"
            Sql = Sql & vbCrLf & "Nombre: " & AdoAux(Index).Recordset!Nombre
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM productos_insec "
                Sql = Sql & vWhere & " AND numlinea= " & AdoAux(Index).Recordset!NumLinea
            End If
            
        Case 1 'herbicidas
            Sql = "¿Seguro que desea el herbicida?"
            Sql = Sql & vbCrLf & "Nombre: " & AdoAux(Index).Recordset!Nombre
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM productos_herb "
                Sql = Sql & vWhere & " AND numlinea= " & AdoAux(Index).Recordset!NumLinea
            End If
            
    End Select

    If Eliminar Then
        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        conn.Execute Sql
        
        EliminarEnProductosGrupo CadenaProductos, CByte(Index)
        
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        CargaGridAux Index, True
        If Not SituarDataTrasEliminar(AdoAux(Index), NumRegElim, True) Then
'            PonerCampos
            
        End If
        If BLOQUEADesdeFormulario2(Me, Me.adodc1, 1) Then BotonModificar
        ' *** si n'hi han tabs ***
'        SituarTab (NumTabMto + 1)
    End If
    
    ModoLineas = 0
    PonerModo 2
    
    Exit Sub
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando linea", Err.Description
End Sub

Private Sub InsertarLinea()
'Inserta registre en les taules de Llínies
Dim nomFrame As String
Dim b As Boolean
Dim Sql2 As String
Dim Sql3 As String

    On Error GoTo EInsertarLinea

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomFrame = "FrameAux0" 'insecticidas
        Case 1: nomFrame = "FrameAux1" 'herbicidas
    End Select
    
    conn.BeginTrans
    b = False
    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        If LineaIntercalada <> -1 Then
            Select Case NumTabMto
                Case 0
                    Sql2 = "update productos_insec set numlinea = numlinea + 1 where codprodu = " & Me.adodc1.Recordset!codprodu & " and numlinea >= " & DBSet(LineaIntercalada, "N")
                    Sql2 = Sql2 & " ORDER BY CODPRODU, NUMLINEA DESC"

                    ' y en todos los productos del grupo
                    If Not EsKaki(Me.adodc1.Recordset!codprodu) Then
                        Sql3 = "update productos_insec set numlinea = numlinea + 1 where codprodu in (select codprodu from productos where codgrupo = " & DBSet(Grupo, "N") & " and codprodu <> " & Me.adodc1.Recordset!codprodu & " and codprodu <> " & DBSet(Kaki, "N") & ") and numlinea >= " & DBSet(LineaIntercalada, "N")
                        Sql3 = Sql3 & " ORDER BY CODPRODU, NUMLINEA DESC"
                        conn.Execute Sql3
                    End If
                Case 1
                    Sql2 = "update productos_herb set numlinea = numlinea + 1 where codprodu = " & Me.adodc1.Recordset!codprodu & " and numlinea >= " & DBSet(LineaIntercalada, "N")
                    Sql2 = Sql2 & " ORDER BY CODPRODU, NUMLINEA DESC"
                    
                    If Not EsKaki(Me.adodc1.Recordset!codprodu) Then
                        Sql3 = "update productos_herb set numlinea = numlinea + 1 where codprodu in (select codprodu from productos where codgrupo = " & DBSet(Grupo, "N") & " and codprodu <> " & Me.adodc1.Recordset!codprodu & " and codprodu <> " & DBSet(Kaki, "N") & ") and numlinea >= " & DBSet(LineaIntercalada, "N")
                        Sql3 = Sql3 & " ORDER BY CODPRODU, NUMLINEA DESC"
                        conn.Execute Sql3
                    End If
            End Select
            conn.Execute Sql2
        End If
        
        If InsertarDesdeForm2(Me, 2, nomFrame) Then
        
            b = InsertarEnProductosGrupo(CadenaProductos, CByte(NumTabMto))
        
        End If
    End If

EInsertarLinea:
    If Err.Number <> 0 Or Not b Then
        conn.RollbackTrans
    Else
        conn.CommitTrans
    
            If b Then
                b = BLOQUEADesdeFormulario2(Me, Me.adodc1, 1)
                Select Case NumTabMto
                    Case 0, 1 ' *** els index de les llinies en grid (en o sense tab) ***
                        CargaGridAux NumTabMto, True
                        
                        If b And LineaIntercalada = -1 Then
                            BotonAnyadirLinea NumTabMto
                        Else
'                            LLamaLineasAux NumTabMto, 0
'                            i = adodc1.Recordset.Fields(0)
'                            PonerModo 2
'                            CargaGrid CadB
'                            adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & i)
'                            PonerFocoGrid Me.DataGrid1

                            cmdCancelar_Click
                        End If
                End Select
            End If
    
    
    
    End If
End Sub


Private Function InsertarEnProductosGrupo(CADENA As String, tipo As Byte) As Boolean
' tipo = 0 : insecticidas
' tipo = 1 : herbicidas
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim SqlValues As String
Dim I As Integer
Dim SQLinsert As String


    On Error GoTo eInsertarEnProductosGrupo

    InsertarEnProductosGrupo = False

'   Sql = "select codprodu from productos where codprodu in (" & CADENA & ") and codprodu <> " & adodc1.Recordset!codprodu
    If EsKaki(CStr(Me.adodc1.Recordset!codprodu)) Then
        InsertarEnProductosGrupo = True
        Exit Function
    Else
        Sql = "select codprodu from productos where codgrupo = " & DBSet(Grupo, "N") & " and codprodu <> " & DBSet(Kaki, "N") & " and codprodu <> " & adodc1.Recordset!codprodu
    End If
    Set Rs = New ADODB.Recordset
    
    SqlValues = ""
    
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        SqlValues = SqlValues & "(" & DBSet(Rs!codprodu, "N") & ","
        If tipo = 0 Then
            For I = 1 To 6
                If I = 1 Or I = 6 Then
                    SqlValues = SqlValues & DBSet(txtAux1(I).Text, "N") & ","
                Else
                    SqlValues = SqlValues & DBSet(txtAux1(I).Text, "T") & ","
                End If
            Next I
        Else
            For I = 1 To 5
                If I = 1 Then
                    SqlValues = SqlValues & DBSet(txtAux3(I).Text, "N") & ","
                Else
                    SqlValues = SqlValues & DBSet(txtAux3(I).Text, "T") & ","
                End If
            Next I
            SqlValues = SqlValues & DBSet(Me.chkAux1(0).Value, "N") & ","
        End If
        SqlValues = Mid(SqlValues, 1, Len(SqlValues) - 1) & "),"
        Rs.MoveNext
    Wend
    Set Rs = Nothing

    If SqlValues <> "" Then
        If tipo = 0 Then
            SQLinsert = "insert into productos_insec (codprodu,numlinea,matactiva,nombre,dosis,plagas,psdias,cultivoaut) values "
        Else
            SQLinsert = "insert into productos_herb (codprodu,numlinea,matactiva,nombre,dosis,plagas,eshelicida) values "
        End If
        
        conn.Execute SQLinsert & Mid(SqlValues, 1, Len(SqlValues) - 1)
    End If
    InsertarEnProductosGrupo = True
    Exit Function
    
eInsertarEnProductosGrupo:
    MuestraError Err.Number, "Insertar en productos del grupo", Err.Description
End Function


Private Function ModificarEnProductosGrupo(CADENA As String, tipo As Byte) As Boolean
' tipo = 0 : insecticidas
' tipo = 1 : herbicidas
Dim Sql As String
Dim Rs As ADODB.Recordset

    On Error GoTo eModificarEnProductosGrupo


    ModificarEnProductosGrupo = True
    
    If EsKaki(Me.adodc1.Recordset!codprodu) Then Exit Function

'    If CADENA = "" Then Exit Function

    If tipo = 0 Then
        Sql = "update productos_insec set "
        Sql = Sql & "matactiva = " & DBSet(txtAux1(2).Text, "T")
        Sql = Sql & ",nombre = " & DBSet(txtAux1(3).Text, "T")
        Sql = Sql & ",dosis = " & DBSet(txtAux1(4).Text, "T")
        Sql = Sql & ",plagas = " & DBSet(txtAux1(5).Text, "T")
        Sql = Sql & ",psdias = " & DBSet(txtAux1(6).Text, "N")
        Sql = Sql & ",cultivoaut = " & DBSet(txtAux1(7).Text, "T")
        Sql = Sql & " where codprodu in (select codprodu from productos where codgrupo = " & DBSet(Grupo, "N") & " and codprodu <> " & DBSet(Kaki, "N") & " and codprodu <> " & Me.adodc1.Recordset!codprodu & ") and numlinea = " & DBSet(txtAux1(1).Text, "N")
    Else
        Sql = "update productos_herb set "
        Sql = Sql & "matactiva = " & DBSet(txtAux3(2).Text, "T")
        Sql = Sql & ",nombre = " & DBSet(txtAux3(3).Text, "T")
        Sql = Sql & ",dosis = " & DBSet(txtAux3(4).Text, "T")
        Sql = Sql & ",plagas = " & DBSet(txtAux3(5).Text, "T")
        Sql = Sql & ",eshelicida = " & DBSet(chkAux1(0).Value, "N")
        Sql = Sql & " where codprodu in (select codprodu from productos where codgrupo = " & DBSet(Grupo, "N") & " and codprodu <> " & DBSet(Kaki, "N") & " and codprodu <> " & Me.adodc1.Recordset!codprodu & ") and numlinea = " & DBSet(txtAux3(1).Text, "N")
    End If
            
    conn.Execute Sql
    Exit Function
    
eModificarEnProductosGrupo:
    ModificarEnProductosGrupo = False
    MuestraError Err.Number, "Modificar en productos del grupo", Err.Description
End Function


Private Function EliminarEnProductosGrupo(CADENA As String, tipo As Byte) As Boolean
' tipo = 0 : insecticidas
' tipo = 1 : herbicidas
Dim Sql As String
Dim Rs As ADODB.Recordset

    On Error GoTo eModificarEnProductosGrupo

    EliminarEnProductosGrupo = True
    
    If EsKaki(adodc1.Recordset!codprodu) Then Exit Function

    If tipo = 0 Then
        Sql = "delete from productos_insec "
        Sql = Sql & " where codprodu in (select codprodu from productos where codgrupo = " & DBSet(Grupo, "N") & ") and codprodu <> " & DBSet(Kaki, "N") & " and numlinea = " & Me.AdoAux(0).Recordset!NumLinea
    Else
        Sql = "delete from productos_herb "
        Sql = Sql & " where codprodu in (select codprodu from productos where codgrupo = " & DBSet(Grupo, "N") & ") and codprodu <> " & DBSet(Kaki, "N") & " and numlinea = " & Me.AdoAux(1).Recordset!NumLinea
    End If
            
    conn.Execute Sql
    Exit Function
    
eModificarEnProductosGrupo:
    EliminarEnProductosGrupo = False

    MuestraError Err.Number, "Modificar en productos del grupo", Err.Description
End Function

Private Function ModificarLinea() As Boolean
'Modifica registre en les taules de Llínies
Dim nomFrame As String
Dim V As Integer
Dim b As Boolean
    
    On Error GoTo eModificarLinea

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomFrame = "FrameAux0" 'productos_insectidas
        Case 1: nomFrame = "FrameAux1" 'productos_herbicidas
    End Select
    ModificarLinea = False
    
    b = False
    conn.BeginTrans
    
    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomFrame) Then
            
            b = ModificarEnProductosGrupo(CadenaProductos, CByte(NumTabMto))
            
            ModoLineas = 0
            
            Select Case NumTabMto
                Case 0
                    V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
                Case 1
                    V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
            End Select
            CargaGridAux NumTabMto, True
            
            ' *** si n'hi han tabs ***
'            SituarTab (NumTabMto + 1)

            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
            PonerFocoGrid Me.DataGridAux(NumTabMto)
            AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
            
            LLamaLineasAux NumTabMto, 0
            ModificarLinea = True
        End If
    End If
    
eModificarLinea:
    If Err.Number <> 0 Or Not b Then
        conn.RollbackTrans
    Else
        conn.CommitTrans
    End If
End Function


Private Sub LimpiarCampos()
    On Error Resume Next
    
    limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    
    ' *** si n'hi han combos a la capçalera ***
    ' *****************************************

    If Err.Number <> 0 Then Err.Clear
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
    ' ****************************************************
    
    SepuedeBorrar = True
End Function

Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    adodc1.RecordSource = CadenaConsulta
    adodc1.Refresh
    
    If adodc1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        PonerModo 2
        'Data1.Recordset.MoveLast
        adodc1.Recordset.MoveFirst
        PonerCampos
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub

EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub PonerCampos()
Dim I As Integer
Dim codpobla As String, despobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If adodc1.Recordset.EOF Then Exit Sub
    
    ' *** si n'hi han llínies en datagrids ***
    'For i = 0 To DataGridAux.Count - 1
    For I = 0 To 1
        If AdoAux(I).Recordset Is Nothing Then
            CargaGridAux I, False
        Else
            CargaGridAux I, True
        End If
        If Not AdoAux(I).Recordset.EOF Then _
            PonerCamposForma2 Me, AdoAux(I), 2, "FrameAux" & I
    Next I

    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Me.adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu
End Sub

Private Function EsKaki(Producto As String) As Boolean
Dim Sql As String

    EsKaki = False
    
    Sql = "select codprodukaki from rparam"
    Kaki = DevuelveValor(Sql)
    If Kaki <> 0 Then
        EsKaki = (Kaki = CLng(Producto))
    End If
    


End Function
