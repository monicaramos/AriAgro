VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmVtasLinPalets 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Variedades de Palets"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   8115
   Icon            =   "frmVtasLinPalets.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameAux0 
      Caption         =   "Calibres"
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
      Height          =   3360
      Left            =   180
      TabIndex        =   27
      Top             =   3510
      Width           =   7815
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   5
         Left            =   1755
         MaxLength       =   9
         TabIndex        =   13
         Tag             =   "Num.Linea 1|N|N|||palets_calibre|numline1|00|S|"
         Text            =   "Linea"
         Top             =   2655
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   4
         Left            =   2475
         MaxLength       =   6
         TabIndex        =   14
         Tag             =   "Variedad|N|N|||palets_calibre|codvarie|000000||"
         Text            =   "varied"
         Top             =   2655
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   3
         Left            =   1080
         MaxLength       =   9
         TabIndex        =   12
         Tag             =   "Num.Linea|N|N|||palets_calibre|numlinea|00|S|"
         Text            =   "Linea"
         Top             =   2655
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.CommandButton btnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   300
         Index           =   0
         Left            =   3735
         MaskColor       =   &H00000000&
         TabIndex        =   19
         ToolTipText     =   "Buscar calibre"
         Top             =   2655
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   3945
         TabIndex        =   36
         Top             =   2655
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   2
         Left            =   5760
         MaxLength       =   9
         TabIndex        =   16
         Tag             =   "Num.Cajas|N|N|||palets_calibre|numcajas|#,##0||"
         Text            =   "cajas"
         Top             =   2655
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   1
         Left            =   3285
         MaxLength       =   2
         TabIndex        =   15
         Tag             =   "Calibre|N|N|||palets_calibre|codcalib|00||"
         Text            =   "calibre"
         Top             =   2655
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   0
         Left            =   270
         MaxLength       =   16
         TabIndex        =   11
         Tag             =   "N�mero Palet|N|N|||palets_calibre|numpalet|000000|S|"
         Text            =   "numpalet"
         Top             =   2655
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   0
         Left            =   135
         TabIndex        =   28
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
         Bindings        =   "frmVtasLinPalets.frx":000C
         Height          =   2595
         Index           =   0
         Left            =   135
         TabIndex        =   29
         Top             =   630
         Width           =   7590
         _ExtentX        =   13388
         _ExtentY        =   4577
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
   Begin VB.Frame Frame2 
      Height          =   2940
      Index           =   0
      Left            =   135
      TabIndex        =   22
      Top             =   495
      Width           =   7875
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   7
         Tag             =   "Nro Lote|T|S|||palets_variedad|nrolote|||"
         Text            =   "123456789012345"
         Top             =   2430
         Width           =   1515
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   3420
         MaxLength       =   4
         TabIndex        =   8
         Tag             =   "Cajas|N|N|0|9999|palets_variedad|numcajas|#,##0||"
         Top             =   2430
         Width           =   1350
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1575
         MaxLength       =   16
         TabIndex        =   5
         Tag             =   "Forfait|T|N|||palets_variedad|codforfait|||"
         Text            =   "1234567890123456"
         Top             =   1710
         Width           =   1530
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   3150
         TabIndex        =   43
         Text            =   "12345678901234567890"
         Top             =   1710
         Width           =   4380
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1575
         MaxLength       =   3
         TabIndex        =   4
         Tag             =   "Marca|N|N|0|999|palets_variedad|codmarca|000||"
         Text            =   "123"
         Top             =   1350
         Width           =   720
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2385
         TabIndex        =   41
         Text            =   "12345678901234567890"
         Top             =   1350
         Width           =   5145
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1575
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "Variedad Comercial|N|N|0|999999|palets_variedad|codvarco|000000||"
         Text            =   "123456"
         Top             =   990
         Width           =   720
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2385
         TabIndex        =   39
         Text            =   "12345678901234567890"
         Top             =   990
         Width           =   5145
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2385
         TabIndex        =   34
         Text            =   "12345678901234567890"
         Top             =   630
         Width           =   5145
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1575
         MaxLength       =   6
         TabIndex        =   2
         Tag             =   "Variedad|N|N|0|999999|palets_variedad|codvarie|000000||"
         Text            =   "123456"
         Top             =   630
         Width           =   720
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   6090
         MaxLength       =   16
         TabIndex        =   10
         Tag             =   "Peso Neto|N|N|-999999|999999|palets_variedad|pesoneto|###,##0||"
         Top             =   2430
         Width           =   1395
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   4830
         MaxLength       =   16
         TabIndex        =   9
         Tag             =   "Peso Bruto|N|S|-999999|999999|palets_variedad|pesobrut|###,##0||"
         Top             =   2430
         Width           =   1215
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   6
         Tag             =   "Categoria|T|S|||palets_variedad|categori|||"
         Top             =   2070
         Width           =   735
      End
      Begin VB.TextBox text1 
         BackColor       =   &H80000013&
         Height          =   285
         Index           =   1
         Left            =   3645
         MaxLength       =   40
         TabIndex        =   1
         Tag             =   "Linea Palet|N|N|||palets_variedad|numlinea|00|S|"
         Text            =   "1234657890123456798012345678901234567890"
         Top             =   225
         Width           =   600
      End
      Begin VB.TextBox text1 
         BackColor       =   &H80000013&
         Height          =   285
         Index           =   0
         Left            =   1575
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "N�mero Palet|N|N|||palets_variedad|numpalet|0000000|S|"
         Text            =   "1234567"
         Top             =   225
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "N�mero Lote"
         Height          =   285
         Index           =   5
         Left            =   180
         TabIndex        =   46
         Top             =   2430
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Cajas"
         Height          =   255
         Index           =   4
         Left            =   3420
         TabIndex        =   45
         Top             =   2190
         Width           =   1185
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1305
         ToolTipText     =   "Buscar Forfait"
         Top             =   1710
         Width           =   240
      End
      Begin VB.Label Label13 
         Caption         =   "Forfait"
         Height          =   255
         Left            =   180
         TabIndex        =   44
         Top             =   1710
         Width           =   690
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1305
         ToolTipText     =   "Buscar Marca"
         Top             =   1350
         Width           =   240
      End
      Begin VB.Label Label12 
         Caption         =   "Marca"
         Height          =   255
         Left            =   180
         TabIndex        =   42
         Top             =   1350
         Width           =   690
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1305
         ToolTipText     =   "Buscar Variedad Comercial"
         Top             =   990
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Var.Comercial"
         Height          =   255
         Left            =   180
         TabIndex        =   40
         Top             =   990
         Width           =   1050
      End
      Begin VB.Label Label10 
         Caption         =   "Variedad"
         Height          =   255
         Left            =   180
         TabIndex        =   35
         Top             =   630
         Width           =   690
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1305
         ToolTipText     =   "Buscar Variedad"
         Top             =   630
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Peso Neto"
         Height          =   255
         Index           =   3
         Left            =   6120
         TabIndex        =   33
         Top             =   2190
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Peso Bruto"
         Height          =   255
         Index           =   2
         Left            =   4830
         TabIndex        =   32
         Top             =   2190
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "Categoria"
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   31
         Top             =   2070
         Width           =   1185
      End
      Begin VB.Label Label6 
         Caption         =   "Linea"
         Height          =   255
         Index           =   1
         Left            =   3105
         TabIndex        =   30
         Top             =   255
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "N�mero Palet"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   23
         Top             =   255
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   240
      TabIndex        =   20
      Top             =   6840
      Width           =   2865
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
         Left            =   120
         TabIndex        =   21
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6945
      TabIndex        =   18
      Top             =   6975
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5850
      TabIndex        =   17
      Top             =   6975
      Width           =   1035
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "�ltimo"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
      Enabled         =   0   'False
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Index           =   0
         Left            =   6525
         TabIndex        =   26
         Top             =   90
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   6930
      TabIndex        =   24
      Top             =   6960
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   6345
      TabIndex        =   37
      Text            =   "Text3"
      Top             =   6570
      Width           =   1300
   End
   Begin VB.Label Label11 
      Caption         =   "TOTAL CAJAS: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   0
      Left            =   3870
      TabIndex        =   38
      Top             =   6615
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
Attribute VB_Name = "frmVtasLinPalets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MANOLO                   -+-+
' +-+- Men�: CLIENTES                  -+-+
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

' +-+-+-+- DISSENY +-+-+-+-
' 1. Posar tots els controls al formulari
' 2. Posar els index correlativament
' 3. Si n'hi han botons de buscar repasar el ToolTipText
' 4. Alliniar els camps num�rics a la dreta i el resto a l'esquerra
' 5. Posar els TAGs
' (si es INTEGER: si PK => m�nim 1; si no PK => m�nim 0; m�xim => 99; format => 00)
' (si es DECIMAL; m�nim => 0; m�xim => 99.99; format => #,###,###,##0.00)
' (si es DATE; format => dd/mm/yyyy)
' 6. Posar els MAXLENGTHs
' 7. Posar els TABINDEXs

Option Explicit

'Dim T1 As Single

Public DatosADevolverBusqueda As String    'Tindr� el n� de text que vol que torne, empipat
Public Event DatoSeleccionado(CadenaSeleccion As String)
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean
Public Palet As Currency
Public Linea As Currency

Public ModoExt As Byte

' *** declarar els formularis als que vaig a cridar ***
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1

Private WithEvents frmMar As frmManMarcas 'marcas
Attribute frmMar.VB_VarHelpID = -1
Private WithEvents frmVar As frmManVariedad 'variedades
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmFor As frmManForfaits 'forfaits
Attribute frmFor.VB_VarHelpID = -1
Private WithEvents frmCali As frmManCalibres 'calibres
Attribute frmCali.VB_VarHelpID = -1


'*****************************************************
Private Modo As Byte
'*************** MODOS ********************
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la b�squeda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edici� del camp
'   3.-  Inserci� de nou registre
'   4.-  Modificar
'   5.-  Manteniment Llinies

'+-+-Variables comuns a tots els formularis+-+-+

Dim ModoLineas As Byte
'1.- Afegir,  2.- Modificar,  3.- Borrar,  0.-Passar control a Ll�nies

Dim NumTabMto As Integer 'Indica quin n� de Tab est� en modo Mantenimient
Dim TituloLinea As String 'Descripci� de la ll�nia que est� en Mantenimient
Dim PrimeraVez As Boolean

Private CadenaConsulta As String 'SQL de la taula principal del formulari
Private Ordenacion As String
Private NombreTabla As String  'Nom de la taula
Private NomTablaLineas As String 'Nom de la Taula de ll�nies del Mantenimient en que estem

Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

'Private VieneDeBuscar As Boolean
'Per a quan torna 2 poblacions en el mateix codi Postal. Si ve de pulsar prismatic
'de b�squeda posar el valor de poblaci� seleccionada i no tornar a recuperar de la Base de Datos

Dim btnPrimero As Byte 'Variable que indica el n� del Bot� PrimerRegistro en la Toolbar1
'Dim CadAncho() As Boolean  'array, per a quan cridem al form de ll�nies
Dim indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim CadB As String

Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    '++monica
'    BloqueaRegistro "palets", "numpalet = " & Text1(0).Text
    
    Select Case Index
        Case 0 'calibres
            Set frmCali = New frmManCalibres
            frmCali.DatosADevolverBusqueda = "0|2|3|"
            frmCali.CodigoActual = txtAux(1).Text
            frmCali.ParamVariedad = txtAux(4).Text
            frmCali.Show vbModal
            Set frmCali = Nothing
            PonerFoco txtAux(1)
    End Select
    If Modo = 4 Then BloqueaRegistro "palets", "numpalet = " & Text1(0).Text
    'BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub


Private Sub cmdAceptar_Click()
Dim b As Boolean
Dim V As Integer

    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'B�SQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm2(Me, 1) Then
'                    text2(9).Text = PonerNombreCuenta(text1(9), Modo, text1(0).Text)
        
'                    Data1.RecordSource = "Select * from " & NombreTabla & _
'                                        " where numpalet = " & DBSet(text1(0).Text, "N") & _
'                                        " and numlinea = " & DBSet(text1(1).Text, "N") & " " & Ordenacion
'                    PosicionarData

                    TerminaBloquear
                    BloqueaRegistro "palets", "numpalet = " & Text1(0).Text

                    CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                    PonerCadenaBusqueda
                    'Ponerse en Modo Insertar Lineas
                    BotonAnyadirLinea 0

                End If
            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario2(Me, 1) Then
                    TerminaBloquear
                    '++monica
                    BloqueaRegistro "palets", "numpalet = " & Text1(0).Text
                    
                    PosicionarData
                End If
            Else
                ModoLineas = 0
            End If
        ' *** si n'hi han ll�nies ***
        Case 5 'LL�NIES
            Select Case ModoLineas
                Case 1 'afegir ll�nia
'                    InsertarLinea
                    If InsertarLinea Then
                        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                        PonerCadenaBusqueda
                        b = BLOQUEADesdeFormulario2(Me, Data1, 1)
                        CargaGrid 0, True
                        If b Then BotonAnyadirLinea NumTabMto
            
                        
                    End If

                Case 2 'modificar ll�nies
                    If ModificarLinea Then
                        ModoLineas = 0
                        
                        V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el n� de llinia
                        
                        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                        PonerCadenaBusqueda
                        b = BLOQUEADesdeFormulario2(Me, Data1, 1)
                        
                        CargaGrid NumTabMto, True
                        
                        PonerFocoGrid Me.DataGridAux(NumTabMto)
                        AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
                        
                        LLamaLineas NumTabMto, 0
                        TerminaBloquear
                        '++monica
                        BloqueaRegistro "palets", "numpalet = " & Text1(0).Text
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


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PrimeraVez Then
        PrimeraVez = False
    
        PonerCampos
        ModoLineas = 0
           
        CalcularTotales
        
        Modo = ModoExt
        Select Case Modo
            Case 3
                mnNuevo_Click
            Case 4
                mnModificar_Click
        End Select
        
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim cad As String

'    If ComprobarNumCajas Then Exit Sub

    cad = ""
    If Text1(0).Text <> "" And Text1(1).Text <> "" Then
        cad = Text1(0).Text & "|" & Text1(1).Text & "|"
    End If
    RaiseEvent DatoSeleccionado(cad)

    CheckValueGuardar Me.Name, Me.chkVistaPrevia(0).Value
    Screen.MousePointer = vbDefault
    TerminaBloquear
    
End Sub

Private Sub Form_Load()
Dim I As Integer

    PrimeraVez = True
    
    ' ICONETS DE LA BARRA
    btnPrimero = 16 'index del bot� "primero"
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'l'1 i el 2 son separadors
        .Buttons(3).Image = 1   'Buscar
        .Buttons(4).Image = 2   'Totss
        'el 5 i el 6 son separadors
        .Buttons(7).Image = 3   'Insertar
        .Buttons(8).Image = 4   'Modificar
        .Buttons(9).Image = 5   'Borrar
        .Buttons(11).Image = 19   'Expandir A�adir, Borrar y Modificar
        'el 10 i el 11 son separadors
        .Buttons(12).Image = 10  'Imprimir
        .Buttons(13).Image = 11  'Eixir
        'el 13 i el 14 son separadors
        .Buttons(btnPrimero).Image = 6  'Primer
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Seg�ent
        .Buttons(btnPrimero + 3).Image = 9 '�ltim
    End With
    
    ' ******* si n'hi han ll�nies *******
    'ICONETS DE LES BARRES ALS TABS DE LL�NIA
    For I = 0 To ToolAux.Count - 1
        With Me.ToolAux(I)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
        End With
    Next I
    ' ***********************************
    
    'cargar IMAGES de busqueda
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
    
    'carga IMAGES de mail
'    For i = 0 To Me.imgMail.Count - 1
'        Me.imgMail(i).Picture = frmPpal.imgListImages16.ListImages(2).Picture
'    Next i
    
    
    LimpiarCampos   'Neteja els camps TextBox
    ' ******* si n'hi han ll�nies *******
    DataGridAux(0).ClearFields
    
    '*** canviar el nom de la taula i l'ordenaci� de la cap�alera ***
    NombreTabla = "palets_variedad"
    Ordenacion = " ORDER BY numpalet"
    
    'Mirem com est� guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    Data1.ConnectionString = conn
    '***** cambiar el nombre de la PK de la cabecera *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where numpalet=" & Palet & " and numlinea = " & Linea
    Data1.Refresh
    
'    If DatosADevolverBusqueda = "" Then
'        PonerModo 0
'    Else
'        PonerModo 1 'b�squeda
'        ' *** posar de groc els camps visibles de la clau primaria de la cap�alera ***
'        Text1(0).BackColor = vbYellow 'codforfait
'        ' ****************************************************************************
'    End If
End Sub

Private Sub LimpiarCampos()
    On Error Resume Next
    
    limpiar Me   'M�tode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub LimpiarCamposLin(FrameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, FrameAux  'M�tode general: Neteja els controls TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO s'habiliten, o no, els diversos camps del
'   formulari en funci� del modo en que anem a treballar
Private Sub PonerModo(Kmodo As Byte, Optional indFrame As Integer)
Dim I As Integer, NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo
 
    'Actualisa Iconos Insertar,Modificar,Eliminar
    'ActualizarToolbar Modo, Kmodo
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo, ModoLineas
       
    'Modo 2. N'hi han datos i estam visualisant-los
    '=========================================
    'Posem visible, si es formulari de b�squeda, el bot� "Regresar" quan n'hi han datos
'    If DatosADevolverBusqueda <> "" Then
'        cmdRegresar.visible = (Modo = 2)
'    Else
'        cmdRegresar.visible = False
'    End If
    
    Text1(5).Enabled = True
    
    
    '=======================================
    b = (Modo = 2)
    'Posar Fleches de desplasament visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Nom�s es per a saber que n'hi ha + d'1 registre
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    '---------------------------------------------
    
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    cmdRegresar.visible = Not b

    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
    'Si estem en Insertar a m�s neteja els camps Text1
    BloquearText1 Me, Modo
    
    '*** si n'hi han combos a la cap�alera ***
    '**************************
    
    ' ***** bloquejar tots els controls visibles de la clau primaria de la cap�alera ***
    If Modo = 4 Then
        BloquearTxt Text1(0), True 'si estic en  modificar, bloqueja la clau primaria
        BloquearTxt Text1(1), True 'si estic en  modificar, bloqueja la clau primaria
    End If
    ' **********************************************************************************
    
    ' **** si n'hi han imagens de buscar en la cap�alera *****
    BloquearImgBuscar Me, Modo, ModoLineas
    BloquearImgZoom Me, Modo, ModoLineas
    ' ********************************************************
    imgBuscar(0).visible = (Modo = 3)
    imgBuscar(0).Enabled = (Modo = 3)
    
        
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    PonerLongCampos

    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 0, False
    End If
    
    b = (Modo = 4) Or (Modo = 2)
    DataGridAux(0).Enabled = b
      
    ' ****** si n'hi han combos a la cap�alera ***********************
    ' ****************************************************************
    
    PonerModoOpcionesMenu (Modo) 'Activar opcions men� seg�n modo
    PonerOpcionesMenu   'Activar opcions de men� seg�n nivell
                        'de permisos de l'usuari

EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los TEXT1
    PonerLongCamposGnral Me, Modo, 1
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub

Private Sub PonerModoOpcionesMenu(Modo)
'Actives unes Opcions de Men� i Toolbar seg�n el modo en que estem
Dim b As Boolean, bAux As Boolean
Dim I As Byte
    
    'Barra de CAP�ALERA
    '------------------------------------------
    'b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    b = (Modo = 2 Or Modo = 0)
    'Buscar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnBuscar.Enabled = b
    'Vore Tots
    Toolbar1.Buttons(4).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    'Insertar
    Toolbar1.Buttons(7).Enabled = b And Not DeConsulta
    Me.mnNuevo.Enabled = b And Not DeConsulta
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(9).Enabled = b
    Me.mnEliminar.Enabled = b
    
    'Expandir operaciones
    Toolbar1.Buttons(11).Enabled = True And Not DeConsulta
    'Imprimir
    'Toolbar1.Buttons(12).Enabled = (b Or Modo = 0)
    Toolbar1.Buttons(12).Enabled = True And Not DeConsulta
       
    ' *** si n'hi han ll�nies que tenen grids (en o sense tab) ***
'++monica: si insertamos lo he quitado
'    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
    b = (Modo = 4 Or Modo = 2) And Not DeConsulta
    For I = 0 To ToolAux.Count - 1
        ToolAux(I).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.AdoAux(I).Recordset.RecordCount > 0)
        ToolAux(I).Buttons(2).Enabled = bAux
        ToolAux(I).Buttons(3).Enabled = bAux
    Next I
    
End Sub

Private Sub Desplazamiento(Index As Integer)
'Botons de Despla�ament; per a despla�ar-se pels registres de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index
    PonerCampos
End Sub

Private Function MontaSQLCarga(Index As Integer, enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basant-se en la informaci� proporcionada pel vector de camps
'   crea un SQl per a executar una consulta sobre la base de datos que els
'   torne.
' Si ENLAZA -> Enla�a en el data1
'           -> Si no el carreguem sense enlla�ar a cap camp
'--------------------------------------------------------------------
Dim Sql As String
Dim Tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
               
        Case 0 'CALIBRES
            Sql = "SELECT palets_calibre.numpalet, palets_calibre.numlinea, palets_calibre.numline1, "
            Sql = Sql & "palets_calibre.codvarie, palets_calibre.codcalib, calibres.nomcalib, palets_calibre.numcajas "
            
            Sql = Sql & " FROM palets_calibre, calibres "
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE palets_calibre.numpalet = '-1'"
            End If
            Sql = Sql & " and palets_calibre.codcalib = calibres.codcalib"
            Sql = Sql & " and palets_calibre.codvarie = calibres.codvarie"
            Sql = Sql & " ORDER BY palets_calibre.codcalib"
               
    End Select
    
    MontaSQLCarga = Sql
End Function

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
        '   Com la clau principal es �nica, en posar el sql apuntant
        '   al valor retornat sobre la clau ppal es suficient
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        ' **********************************
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmCali_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 2) 'codcalib
    txtAux2(2).Text = RecuperaValor(CadenaSeleccion, 3) 'descripcion
End Sub

Private Sub frmFor_DatoSeleccionado(CadenaSeleccion As String)
'Forfaits
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codforfait
    text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmMar_DatoSeleccionado(CadenaSeleccion As String)
'Marcas
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codmarca
    text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Variedades
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codvariedad
    text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(indice).Text = vCampo
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    Screen.MousePointer = vbHourglass
    frmListConfeccion.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnModificar_Click()
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************
'--monica
'    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
            BotonModificar
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
        Case 3  'B�scar
           mnBuscar_Click
        Case 4  'Tots
            mnVerTodos_Click
        Case 7  'Nou
            mnNuevo_Click
        Case 8  'Modificar
            mnModificar_Click
        Case 9  'Borrar
            mnEliminar_Click
        Case 12 'Imprimir
            mnImprimir_Click
        Case 13    'Eixir
            mnSalir_Click
            
        Case btnPrimero To btnPrimero + 3 'Fleches Despla�ament
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub BotonBuscar()
Dim I As Integer
' ***** Si la clau primaria de la cap�alera no es Text1(0), canviar-ho en <=== *****
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        PonerFoco Text1(0) ' <===
        Text1(0).BackColor = vbYellow ' <===
        ' *** si n'hi han combos a la cap�alera ***
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

    CadB = ObtenerBusqueda2(Me, 1)
    
    If chkVistaPrevia(0) = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        ' *** foco al 1r camp visible de la cap�alera que siga clau primaria ***
        PonerFoco Text1(0)
        ' **********************************************************************
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
    Dim cad As String
        
    'Cridem al form
    ' **************** arreglar-ho per a vore lo que es desije ****************
    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
    cad = ""
    cad = cad & ParaGrid(Text1(0), 20, "C�digo")
    cad = cad & ParaGrid(Text1(1), 20, "Confecci�n")
    cad = cad & ParaGrid(Text1(2), 60, "Descripci�n")
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vtabla = NombreTabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        frmB.vDevuelve = "0|1|2|" '*** els camps que volen que torne ***
        frmB.vTitulo = "Forfaits" ' ***** repasa a��: t�tol de BuscaGrid *****
        frmB.vSelElem = 1

        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha posat valors i tenim que es formulari de b�squeda llavors
        'tindrem que tancar el form llan�ant l'event
        If HaDevueltoDatos Then
            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                cmdRegresar_Click
        Else   'de ha retornat datos, es a decir NO ha retornat datos
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub

Private Sub cmdRegresar_Click()
Dim cad As String
Dim Aux As String
Dim I As Integer
Dim j As Integer

    
    Unload Me
End Sub

Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
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

    LimpiarCampos 'Huida els TextBox
    
    
    PonerModo 3
    
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la cap�alera *******
'    text1(0).Text = SugerirCodigoSiguienteStr("forfaits", "codforfait")
'    FormateaCampo text1(0)
    
    Text1(0).Text = Palet
    Text1(1).Text = SugerirCodigoSiguienteStr("palets_variedad", "numlinea", "numpalet = " & Text1(0).Text)
    Text1(0).BackColor = &H80000013
    Text1(1).BackColor = &H80000013
    Text1(0).Locked = True
    Text1(1).Locked = True
    
    PonerFoco Text1(2) '*** 1r camp visible que siga PK ***
    
    ' *** si n'hi han camps de descripci� a la cap�alera ***
    'PosarDescripcions

End Sub

Private Sub BotonModificar()

    PonerModo 4
    
    Text1(0).Text = Palet
    Text1(1).Text = Linea
    
    Text1(0).BackColor = &H80000013
    Text1(1).BackColor = &H80000013

    ' *** bloquejar els camps visibles de la clau primaria de la cap�alera ***
    BloquearTxt Text1(0), True
    BloquearTxt Text1(1), True
    BloquearTxt Text1(2), True
    
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco Text1(3)
End Sub

Private Sub BotonEliminar()
Dim cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
    ' ***************************************************************************

    ' *************** canviar la pregunta ****************
    cad = "�Seguro que desea eliminar el Forfait?"
    cad = cad & vbCrLf & "C�digo: " & Data1.Recordset.Fields(0)
    cad = cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(1)
    
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not eliminar Then
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
Dim I As Integer
Dim codPobla As String, desPobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la cap�alera
    
    ' *** si n'hi han ll�nies en datagrids ***
    'For i = 0 To DataGridAux.Count - 1
    For I = 0 To 0
            CargaGrid I, True
            If Not AdoAux(I).Recordset.EOF Then _
                PonerCamposForma2 Me, AdoAux(I), 2, "FrameAux" & I
    Next I

    
    ' ************* configurar els camps de les descripcions de la cap�alera *************
    text2(2).Text = PonerNombreDeCod(Text1(2), "variedades", "nomvarie")
    text2(3).Text = DevuelveDesdeBDNew(cAgro, "variedades", "nomvarie", "codvarie", Text1(3).Text, "N")
    text2(4).Text = PonerNombreDeCod(Text1(4), "marcas", "nommarca")
    text2(5).Text = PonerNombreDeCod(Text1(5), "forfaits", "nomconfe")
    ' ********************************************************************************
    
    CalcularTotales
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu
End Sub

Private Sub cmdCancelar_Click()
Dim I As Integer
Dim V


    Select Case Modo
        Case 1, 3 'B�squeda, Insertar
                LimpiarCampos
                If Data1.Recordset.EOF Then
                    PonerModo 0
                Else
                    PonerModo 2
                    PonerCampos
                End If
                ' *** foco al primer camp visible de la cap�alera ***
                PonerFoco Text1(0)

        Case 4  'Modificar
                TerminaBloquear
                '++monica
                BloqueaRegistro "palets", "numpalet = " & Text1(0).Text
                
                PonerModo 2
                PonerCampos
                ' *** primer camp visible de la cap�alera ***
                PonerFoco Text1(0)

        Case 5 'LL�NIES
            Select Case ModoLineas
                Case 1 'afegir ll�nia
                    ModoLineas = 0
                    ' *** les ll�nies que tenen datagrid (en o sense tab) ***
                    If NumTabMto = 0 Or NumTabMto = 1 Or NumTabMto = 2 Or NumTabMto = 4 Then
                        DataGridAux(NumTabMto).AllowAddNew = False
                        ' **** repasar si es diu Data1 l'adodc de la cap�alera ***
                        'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'Modificar
                        LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                        DataGridAux(NumTabMto).Enabled = True
                        DataGridAux(NumTabMto).SetFocus

                        ' *** si n'hi han camps de descripci� dins del grid, els neteje ***
                        'txtAux2(2).text = ""

                    End If

'                    ' *** si n'hi han tabs ***
'                    SituarTab (NumTabMto + 1)

                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        AdoAux(NumTabMto).Recordset.MoveFirst
                    End If

                Case 2 'modificar ll�nies
                    ModoLineas = 0

                    ' *** si n'hi han tabs ***
'                    SituarTab (NumTabMto + 1)
                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                    PonerModo 4
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        ' *** l'Index de Fields es el que canvie de la PK de ll�nies ***
                        V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el n� de llinia
                        AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
                        ' ***************************************************************
                    End If

            End Select

            PosicionarData

            ' *** si n'hi han ll�nies en grids i camps fora d'estos ***
            If Not AdoAux(NumTabMto).Recordset.EOF Then
                DataGridAux_RowColChange NumTabMto, 1, 1
            Else
                LimpiarCamposFrame NumTabMto
            End If
    End Select
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Sql As String
'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm2(Me, 1)
    If Not b Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
    If (Modo = 3) Then 'insertar
        'comprobar si existe ya el cod. del campo clave primaria
        Sql = ""
        Sql = DevuelveDesdeBDNew(cAgro, "palets_calibre", "numpalet", "numpalet", Text1(0).Text, "N", , "numlinea", Text1(1).Text, "N")
        If Sql <> "" Then
            MsgBox "Ya existe el numero de linea para este palet", vbExclamation
            b = False
        End If
    End If
    
    ' ************************************************************************************
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Sub PosicionarData()
Dim cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la cap�alera, no llevar els () ***
    cad = "(numpalet=" & DBSet(Text1(0).Text, "N") & ")"
    cad = cad & " and (numlinea = " & DBSet(Text1(1).Text, "N") & ")"
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    If SituarDataMULTI(Data1, cad, Indicador) Then
    'If SituarData(Data1, cad, Indicador) Then
        If ModoLineas <> 1 Then PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
End Sub

Private Function eliminar() As Boolean
Dim vWhere As String

    On Error GoTo FinEliminar

    conn.BeginTrans
    ' ***** canviar el nom de la PK de la cap�alera, repasar codEmpre *******
    vWhere = " WHERE codforfait=" & DBSet(Data1.Recordset!Codforfait, "T")
        
    ' ***** elimina les ll�nies ****
    conn.Execute "DELETE FROM forfaits_envases " & vWhere
        
    conn.Execute "DELETE FROM forfaits_costes " & vWhere
        
    'Eliminar la CAP�ALERA
    conn.Execute "Delete from " & NombreTabla & vWhere
       
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        conn.RollbackTrans
        eliminar = False
    Else
        conn.CommitTrans
        eliminar = True
    End If
End Function

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
Dim Variedad As String

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    
    ' ***************** configurar els LostFocus dels camps de la cap�alera *****************
    Select Case Index
        Case 0 'codigo de forfait
            Text1(Index).Text = UCase(Text1(Index).Text)
        
        Case 2, 3 'Variedad
            If PonerFormatoEntero(Text1(Index)) Then
                text2(Index).Text = DevuelveDesdeBDNew(cAgro, "variedades", "nomvarie", "codvarie", Text1(Index).Text, "N")
                If text2(Index).Text = "" Then
                    cadMen = "No existe la Variedad: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "�Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        indice = Index
                        Set frmVar = New frmManVariedad
                        frmVar.DatosADevolverBusqueda = "0|1|"
                        frmVar.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        
                        frmVar.Show vbModal
                        Set frmVar = Nothing
                        '++monica
                        BloqueaRegistro "palets", "numpalet = " & Text1(0).Text
'                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                text2(Index).Text = ""
            End If
            
        Case 4 'Marca
            If PonerFormatoEntero(Text1(Index)) Then
                text2(Index) = PonerNombreDeCod(Text1(Index), "marcas", "nommarca")
                If text2(Index).Text = "" Then
                    cadMen = "No existe la Marca: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "�Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        indice = Index
                        Set frmMar = New frmManMarcas
                        frmMar.DatosADevolverBusqueda = "0|1|"
                        frmMar.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        
                        frmMar.Show vbModal
                        Set frmMar = Nothing
                        '++monica
                        BloqueaRegistro "palets", "numpalet = " & Text1(0).Text
'                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                text2(Index).Text = ""
            End If
                
        Case 5 'Forfait
            If Text1(Index).Text <> "" Then
                text2(Index) = PonerNombreDeCod(Text1(Index), "forfaits", "nomconfe")
                If text2(Index).Text = "" Then
                    cadMen = "No existe el Forfait: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "�Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        indice = Index
                        Set frmFor = New frmManForfaits
                        frmFor.DatosADevolverBusqueda = "0|1|"
                        frmFor.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear

                        frmFor.Show vbModal
                        Set frmFor = Nothing
                        '++monica
                        BloqueaRegistro "palets", "numpalet = " & Text1(0).Text
'                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                '++monica:02/12/2008 control d que el forfait sea de la variedad introducida
                Else
                    Variedad = ""
                    Variedad = DevuelveDesdeBDNew(cAgro, "forfaits", "codvarie", "codforfait", Text1(Index).Text, "T")
                    If Variedad <> "" Then
                        If CInt(Variedad) <> CInt(Text1(2).Text) Then
                            MsgBox "El Forfait no es de la Variedad introducida.", vbExclamation
                        End If
                    End If
                '++
                End If
            Else
                text2(Index).Text = ""
            End If
        
        Case 6 ' categoria
            Text1(Index).Text = UCase(Text1(Index).Text)
            
        Case 7, 8 'peso bruto y peso neto
            PonerFormatoEntero Text1(Index)
            If Index = 7 Then CalculoPesoNeto False
            
        Case 9 ' numero de cajas
            PonerFormatoEntero Text1(Index)
            CalculoPesoNeto False
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
        KEYpress KeyAscii
    End If
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvan�ar/Retrocedir els camps en les fleches de despla�ament del teclat.
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
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
    Select Case Button.Index
        Case 1
            BotonAnyadirLinea Index
        Case 2
            BotonModificarLinea Index
        Case 3
            BotonEliminarLinea Index
        Case Else
    End Select
    End If
End Sub

Private Sub BotonEliminarLinea(Index As Integer)
Dim Sql As String
Dim vWhere As String
Dim eliminar As Boolean

    On Error GoTo Error2

    ModoLineas = 3 'Posem Modo Eliminar Ll�nia
    
    If Modo = 4 Then 'Modificar Cap�alera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index

    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar(Index) Then Exit Sub
    NumTabMto = Index
    eliminar = False
   
    vWhere = ObtenerWhereCab(True)
    
    ' ***** independentment de si tenen datagrid o no,
    ' canviar els noms, els formats i el DELETE *****
    Select Case Index
        Case 0 'calibres
            Sql = "�Seguro que desea eliminar el Calibre?"
            Sql = Sql & vbCrLf & "Calibre: " & AdoAux(Index).Recordset!codcalib
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
                Sql = "DELETE FROM palets_calibre "
                Sql = Sql & vWhere & " AND numline1= " & AdoAux(Index).Recordset!numline1
            End If
            
    End Select

    If eliminar Then
        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        '++monica
        BloqueaRegistro "palets", "numpalet = " & Text1(0).Text
        
        conn.BeginTrans
        
        conn.Execute Sql
        
        ActualizarVariedades Text1(0), Text1(1)

        conn.CommitTrans
        
        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
        PonerCadenaBusqueda
'--monica:02102008
'        ' *** si n'hi han tabs sense datagrid, posar l'If ***
'        CargaGrid Index, True
'        If Not SituarDataTrasEliminar(Adoaux(Index), NumRegElim, True) Then
''            PonerCampos
'
'        End If
'        CalcularTotales
'--monica
        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
'--monica:02102008
            'BotonModificar
'--monica
        End If
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
Dim I As Integer
    
    ModoLineas = 1 'Posem Modo Afegir Ll�nia
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Cap�alera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index
    
    ' *** bloquejar la clau primaria de la cap�alera ***
    BloquearTxt Text1(0), True
    BloquearTxt Text1(1), True
    

    ' *** posar el nom del les distintes taules de ll�nies ***
    Select Case Index
        Case 0: vtabla = "palets_calibre"
    End Select
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 0, 1 ' *** pose els index dels tabs de ll�nies que tenen datagrid ***
            ' *** canviar la clau primaria de les ll�nies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            
'            If Index = 1 Then NumF = SugerirCodigoSiguienteStr(vTabla, "codcoste", vWhere)

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
                Case 0 'calibres
                    txtAux(0).Text = Text1(0).Text 'numpalet
                    txtAux(3).Text = Text1(1).Text 'numlinea
                    txtAux(5).Text = SugerirCodigoSiguienteStr("palets_calibre", "numline1", "numpalet = " & Text1(0).Text & " and numlinea =  " & Text1(1).Text) 'numline1
                    txtAux(4).Text = Text1(2).Text
                    
                    txtAux(1).Text = ""
                    txtAux(2).Text = ""
                    txtAux2(2).Text = ""
'                    For i = 0 To 0
'                        BloquearTxt txtAux(i), False
'                    Next i
                    BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux0"
                    PonerFoco txtAux(1)
            End Select
            
    End Select
End Sub

Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim I As Integer
    Dim j As Integer
    
    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub
    
    ModoLineas = 2 'Modificar ll�nia
       
    If Modo = 4 Then 'Modificar Cap�alera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index
    ' *** bloqueje la clau primaria de la cap�alera ***
    BloquearTxt Text1(0), True
  
    Select Case Index
        Case 0, 1 ' *** pose els index de ll�nies que tenen datagrid (en o sense tab) ***
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
        Case 0 ' calibres
        
            txtAux(0).Text = DataGridAux(Index).Columns(0).Text
            txtAux(3).Text = DataGridAux(Index).Columns(1).Text
            txtAux(5).Text = DataGridAux(Index).Columns(2).Text
            txtAux(4).Text = DataGridAux(Index).Columns(3).Text
            txtAux(1).Text = DataGridAux(Index).Columns(4).Text
            txtAux(2).Text = DataGridAux(Index).Columns(6).Text
            
            txtAux2(2).Text = DataGridAux(Index).Columns(5).Text
'22-09-2008
'            For i = 1 To 1
'                BloquearTxt txtAux(i), True
'            Next i
'            BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux0"
            
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 0 'calibres
            PonerFoco txtAux(1)
    End Select
    ' ***************************************************************************************
End Sub

Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim b As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    DeseleccionaGrid DataGridAux(Index)
       
    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Ll�nies
    Select Case Index
        Case 0 'calibres
            txtAux(1).visible = b 'numpalet
            txtAux(1).Top = alto
            txtAux(2).visible = b 'numlinea
            txtAux(2).Top = alto
            txtAux2(2).visible = b
            txtAux2(2).Top = alto
            btnBuscar(0).visible = b
            btnBuscar(0).Top = alto
            
    End Select
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
    
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    

    ' ******* configurar el LostFocus dels camps de ll�nies (dins i fora grid) ********
    Select Case Index
        Case 1 ' codigo de calibre
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(2).Text = DevuelveDesdeBDNew(cAgro, "calibres", "nomcalib", "codvarie", txtAux(4).Text, "N", , "codcalib", txtAux(1).Text, "N")
                If txtAux2(2).Text = "" Then
                    cadMen = "No existe el Calibre: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "�Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCali = New frmManCalibres
                        frmCali.DatosADevolverBusqueda = "0|2|3|"
                        frmCali.NuevoCodigo = txtAux(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        '++monica
                        
                        frmCali.Show vbModal
                        Set frmCali = Nothing
'                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                        BloqueaRegistro "palets", "numpalet = " & Text1(0).Text
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                End If
            Else
                txtAux2(2).Text = ""
            End If
        
        
        Case 2 ' cajas
            If txtAux(Index).Text <> "" Then
                If PonerFormatoEntero(txtAux(Index)) Then
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
    '++monica
'    BloqueaRegistro "palets", "numpalet = " & Text1(0).Text
    
     indice = Index + 2
     Select Case Index
        Case 0, 1 'variedad y variedad comercial
            indice = Index + 2
            Set frmVar = New frmManVariedad
            frmVar.DatosADevolverBusqueda = "0|1|"
            frmVar.CodigoActual = Text1(indice).Text
            frmVar.Show vbModal
            Set frmVar = Nothing
            PonerFoco Text1(indice)
        Case 2 'Marca
            Set frmMar = New frmManMarcas
            frmMar.DatosADevolverBusqueda = "0|1|"
            frmMar.CodigoActual = Text1(4).Text
            frmMar.Show vbModal
            Set frmMar = Nothing
            PonerFoco Text1(4)
        Case 3 'forfait
            Set frmFor = New frmManForfaits
            frmFor.DatosADevolverBusqueda = "0|1|"
            frmFor.CodigoActual = Text1(5).Text
            frmFor.Show vbModal
            Set frmFor = Nothing
            PonerFoco Text1(5)
    End Select
    
    If Modo = 4 Then BloqueaRegistro "palets", "numpalet = " & Text1(0).Text
                'BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub


Private Sub DataGridAux_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
Dim I As Byte

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
                For I = 21 To 24
'                   txtAux(i).Text = ""
                Next I
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
Dim I As Byte

    AdoAux(Index).ConnectionString = conn
    AdoAux(Index).RecordSource = MontaSQLCarga(Index, enlaza)
    AdoAux(Index).CursorType = adOpenDynamic
    AdoAux(Index).LockType = adLockPessimistic
    AdoAux(Index).Refresh
    
    If Not AdoAux(Index).Recordset.EOF Then
        PonerCamposForma2 Me, AdoAux(Index), 2, "FrameAux" & Index
    Else
        ' *** si n'hi han tabs sense datagrids, li pose els valors als camps ***
        NetejaFrameAux "FrameAux3" 'neteja nom�s lo que te TAG
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
Dim I As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)

    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
    
    Select Case Index
        Case 0 'calibres
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;" 'numpalet,numlinea,numline1,codvarie
            tots = tots & "S|txtAux(1)|T|Calibre|1700|;S|btnBuscar(0)|B|||;"
            tots = tots & "S|txtAux2(2)|T|Denominaci�n|4000|;S|txtAux(2)|T|Cajas|1300|;"
            
            arregla tots, DataGridAux(Index), Me
        
'            DataGridAux(0).Columns(6).NumberFormat = "#,###0"
            DataGridAux(0).Columns(6).Alignment = dbgRight
        
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
            
    End Select
    
    DataGridAux(Index).ScrollBars = dbgAutomatic
      
    ' **** si n'hi han ll�nies en grids i camps fora d'estos ****
'    If Not AdoAux(Index).Recordset.EOF Then
'        DataGridAux_RowColChange Index, 1, 1
'    Else
''        LimpiarCamposFrame Index
'    End If
      
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub

Private Function InsertarLinea() As Boolean
'Inserta registre en les taules de Ll�nies
Dim nomFrame As String
Dim b As Boolean
Dim bol As Boolean
Dim MenError As String


    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomFrame = "FrameAux0" 'calibres
    End Select
    
    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        '++monica
        BloqueaRegistro "palets", "numpalet = " & Text1(0).Text
        
        'Aqui empieza transaccion
        conn.BeginTrans
        
        bol = InsertarDesdeForm2(Me, 2, nomFrame)
        If bol Then
            MenError = "Modificando variedades"
            bol = ActualizarVariedades(Text1(0), Text1(1))
        End If
        
'        If InsertarDesdeForm2(Me, 2, nomFrame) Then
'            b = BLOQUEADesdeFormulario2(Me, Data1, 1)
'            Select Case NumTabMto
'                Case 0, 1 ' *** els index de les llinies en grid (en o sense tab) ***
'                     CargaGrid NumTabMto, True
'                    If b Then BotonAnyadirLinea NumTabMto
'            End Select
'
''            SituarTab (NumTabMto + 1)
'        End If
'    End If
    Else
        InsertarLinea = False
        Exit Function
    End If

EInsertarLinea:
        If Err.Number <> 0 Then
            MenError = "Insertando Linea." & vbCrLf & "----------------------------" & vbCrLf & MenError
            MuestraError Err.Number, MenError, Err.Description
            bol = False
        End If
        If bol Then
            conn.CommitTrans
            InsertarLinea = True
        Else
            conn.RollbackTrans
            InsertarLinea = False
        End If

End Function

Private Function ModificarLinea() As Boolean
'Modifica registre en les taules de Ll�nies
Dim nomFrame As String
Dim V As Integer
Dim bol As Boolean
Dim MenError As String

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomFrame = "FrameAux0" 'calibres
    End Select
    ModificarLinea = False
    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        
        conn.BeginTrans
        
        bol = ModificaDesdeFormulario2(Me, 2, nomFrame)
        If bol Then
            MenError = "Modificando variedades"
            bol = ActualizarVariedades(Text1(0), Text1(1))
        End If
        
        
        
'        If ModificaDesdeFormulario2(Me, 2, nomFrame) Then
'            ModoLineas = 0
'
'            Select Case NumTabMto
'                Case 0
'                    V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el n� de llinia
'                Case 1
'                    V = AdoAux(NumTabMto).Recordset.Fields(2) 'el 2 es el n� de llinia
'            End Select
'            CargaGrid NumTabMto, True
'
'            ' *** si n'hi han tabs ***
''            SituarTab (NumTabMto + 1)
'
'            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
'            PonerFocoGrid Me.DataGridAux(NumTabMto)
'            AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
'
'            LLamaLineas NumTabMto, 0
'            ModificarLinea = True
'        End If
'
'        '++monica
'        BloqueaRegistro "palets", "numpalet = " & text1(0).Text
'
    End If
eModificarLinea:
    If Err.Number <> 0 Then
        MenError = "Modificando Linea." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        bol = False
    End If
    If bol Then
        conn.CommitTrans
        ModificarLinea = True
    Else
        conn.RollbackTrans
        ModificarLinea = False
    End If
    
End Function

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la cap�alera ***
    vWhere = vWhere & " numpalet=" & Me.Data1.Recordset!numpalet & " and numlinea = " & Me.Data1.Recordset!NumLinea
    
    ObtenerWhereCab = vWhere
End Function

'' *** neteja els camps dels tabs de grid que
''estan fora d'este, i els camps de descripci� ***
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

'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del rat�n.
'Private Sub DataGridAux_GotFocus(Index As Integer)
'  WheelHook DataGridAux(Index)
'End Sub
'Private Sub DataGridAux_LostFocus(Index As Integer)
'  WheelUnHook
'End Sub

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

    On Error Resume Next

    'total importes de envases para ese forfait
    Sql = "select sum(numcajas) "
    Sql = Sql & " from palets_calibre where numpalet = " & DBSet(Text1(0).Text, "N")
    Sql = Sql & " and numlinea = " & DBSet(Text1(1).Text, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    TotalEnvases = 0
    If Not Rs.EOF Then
        If Rs.Fields(0).Value > 0 Then TotalEnvases = Rs.Fields(0).Value
    End If
    Rs.Close
    Set Rs = Nothing
    
    Text3(0).Text = Format(TotalEnvases, "###,##0")
    If Err.Number <> 0 Then
        Err.Clear
    End If

End Sub

Private Function ObtenerWhereCP(conW As Boolean) As String
Dim Sql As String
On Error Resume Next
    
    Sql = ""
    If conW Then Sql = " WHERE "
    Sql = Sql & NombreTabla & ".numpalet= " & DBSet(Text1(0).Text, "N")
    Sql = Sql & " and " & NombreTabla & ".numlinea=" & Val(Text1(1).Text)
    ObtenerWhereCP = Sql
End Function

Private Sub CalculoPesoNeto(actTabla As Boolean)
Dim CodPalet As String
Dim PesoCaja As String
Dim PesoPalet As String

Dim nPesocaja As Currency
Dim nNumcajas As Long
Dim nPesopalet As Currency
Dim TaraCajas As Long
Dim PesoBruto As Long
Dim PesoNeto As Long

Dim Sql As String

    If Text1(5).Text <> "" Then ' si me han introducido el forfait
        PesoCaja = ""
        PesoCaja = DevuelveDesdeBDNew(cAgro, "forfaits", "pesocaja", "codforfait", Text1(5).Text, "T")
    End If
    
    CodPalet = ""
    CodPalet = DevuelveDesdeBDNew(cAgro, "palets", "codpalet", "numpalet", Text1(0).Text, "N")
    nPesopalet = 0
    If CodPalet <> "" Then
        PesoPalet = ""
        PesoPalet = DevuelveDesdeBDNew(cAgro, "confpale", "pesopale", "codpalet", CodPalet, "N")
        nPesopalet = 0
        If PesoPalet <> "" Then
            nPesopalet = CCur(ImporteSinFormato(PesoPalet))
        End If
    End If
    
'    PesoPalet = ""
'    PesoPalet = DevuelveDesdeBDNew(cAgro, "confpale", "pesopale", "codpalet", Text1(0).Text, "N")
'    nPesopalet = 0
'    If PesoPalet <> "" Then
'        nPesopalet = CCur(ImporteSinFormato(PesoPalet))
'    End If
    
    nPesocaja = 0
    nNumcajas = 0
    If PesoCaja <> "" And Text1(9).Text <> "" Then
        nPesocaja = CCur(ImporteSinFormato(PesoCaja))
        nNumcajas = CLng(ImporteSinFormato(Text1(9).Text))
    End If
    
    TaraCajas = Round2((nPesocaja * nNumcajas) + nPesopalet, 0)
    
    PesoBruto = 0
    If Text1(7).Text <> "" Then
        PesoBruto = CLng(Text1(7).Text)
    End If
    
    PesoNeto = PesoBruto - TaraCajas
    
    If actTabla Then
        Sql = "update palets_variedad set pesoneto = " & DBSet(PesoNeto, "N")
        Sql = Sql & " where numpalet = " & DBSet(Text1(0).Text, "N") & " and numlinea = " & DBSet(Text1(1).Text, "N")
        
        conn.Execute Sql
    Else
        Text1(8) = CStr(PesoNeto)
        PonerFormatoEntero Text1(8)
    End If

End Sub

Private Function ComprobarNumCajas() As Boolean
Dim Sql As String

    ComprobarNumCajas = True
    
'    Sql = "select sum(numcajas) from palets_calibre where codpalet = " & dbset(
    
End Function


Private Function ActualizarVariedades(Pedido As String, Linea As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Sql1 As String

    On Error GoTo eActualizarVariedades

    ActualizarVariedades = False

    Sql1 = "select sum(numcajas) from palets_calibre where numpalet = " & DBSet(Pedido, "N")
    Sql1 = Sql1 & " and numlinea = " & DBSet(Linea, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        If DBLet(Rs.Fields(0).Value, "N") = 0 Then
            Sql = "update palets_variedad set numcajas = 0 "
            Sql = Sql & " where numpalet = " & DBSet(Pedido, "N")
            Sql = Sql & " and numlinea = " & DBSet(Linea, "N")
    
            conn.Execute Sql
            
            Text1(9).Text = 0
        End If
        
        If DBLet(Rs.Fields(0).Value, "N") <> 0 Then
            Sql = "update palets_variedad set numcajas = " & DBSet(Rs.Fields(0).Value, "N")
            Sql = Sql & " where numpalet = " & DBSet(Pedido, "N")
            Sql = Sql & " and numlinea = " & DBSet(Linea, "N")
    
            conn.Execute Sql
            
            Text1(9).Text = Format(DBLet(Rs.Fields(0).Value, "N"), "#,##0")
        End If
    
    End If
    Rs.Close
    Set Rs = Nothing

    CalculoPesoNeto True

eActualizarVariedades:
    If Err.Number = 0 Then ActualizarVariedades = True
    
End Function


