VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCCManConcep 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conceptos de Costes"
   ClientHeight    =   9960
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   17640
   Icon            =   "frmCCManConcep.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9960
   ScaleWidth      =   17640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   90
      TabIndex        =   40
      Top             =   45
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   41
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
   Begin VB.Frame FrameAux1 
      Caption         =   "Líneas"
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
      Height          =   8940
      Left            =   13710
      TabIndex        =   22
      Top             =   105
      Width           =   1470
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   7605
         TabIndex        =   30
         Text            =   "Text3"
         Top             =   2160
         Width           =   1200
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
         Index           =   41
         Left            =   750
         MaxLength       =   16
         TabIndex        =   27
         Tag             =   "Linea|N|N|||ccconcostes_lin|numlinea|00|S|"
         Text            =   "lin"
         Top             =   1260
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.CommandButton btnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   285
         Index           =   1
         Left            =   1560
         MaskColor       =   &H00000000&
         TabIndex        =   26
         ToolTipText     =   "Buscar zona"
         Top             =   1230
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Index           =   42
         Left            =   1710
         TabIndex        =   25
         Text            =   "nombre"
         Top             =   720
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
         Index           =   42
         Left            =   1110
         MaxLength       =   4
         TabIndex        =   24
         Tag             =   "Codigo Linea|N|N|||ccconcostes_lin|codlinea|0000|N|"
         Text            =   "Cost"
         Top             =   1260
         Visible         =   0   'False
         Width           =   375
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
         Index           =   40
         Left            =   180
         MaxLength       =   16
         TabIndex        =   23
         Tag             =   "Código|N|N|0|9999|ccconcostes_lin|codcoste|0000|S|"
         Text            =   "codigo"
         Top             =   1260
         Visible         =   0   'False
         Width           =   555
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   1
         Left            =   120
         TabIndex        =   28
         Top             =   240
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
               Enabled         =   0   'False
               Object.Visible         =   0   'False
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
         Bindings        =   "frmCCManConcep.frx":000C
         Height          =   8145
         Index           =   1
         Left            =   135
         TabIndex        =   29
         Top             =   720
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   14367
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
         Top             =   660
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
   Begin VB.Frame FrameAux2 
      Caption         =   "Cuentas"
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
      Height          =   8940
      Left            =   15210
      TabIndex        =   31
      Top             =   105
      Width           =   2430
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
         Height          =   290
         Index           =   2
         Left            =   1530
         MaskColor       =   &H00000000&
         TabIndex        =   39
         ToolTipText     =   "Buscar Cta.Contable"
         Top             =   1260
         Visible         =   0   'False
         Width           =   195
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
         Height          =   285
         Index           =   50
         Left            =   120
         MaxLength       =   16
         TabIndex        =   36
         Tag             =   "Código|N|N|0|9999|ccconcostes_cta|codcoste|0000|S|"
         Text            =   "codigo"
         Top             =   1260
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
         Height          =   285
         Index           =   52
         Left            =   1140
         MaxLength       =   10
         TabIndex        =   35
         Tag             =   "Cta Contable|T|N|||ccconcostes_cta|ctacontable||N|"
         Text            =   "Cta Contab"
         Top             =   1290
         Visible         =   0   'False
         Width           =   375
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
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   34
         Text            =   "nombre"
         Top             =   1260
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
         Height          =   285
         Index           =   51
         Left            =   750
         MaxLength       =   16
         TabIndex        =   33
         Tag             =   "Linea|N|N|||ccconcostes_cta|numlinea|00|S|"
         Text            =   "lin"
         Top             =   1260
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   7605
         TabIndex        =   32
         Text            =   "Text3"
         Top             =   2160
         Width           =   1200
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   2
         Left            =   120
         TabIndex        =   37
         Top             =   240
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
               Enabled         =   0   'False
               Object.Visible         =   0   'False
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
         Bindings        =   "frmCCManConcep.frx":0024
         Height          =   8145
         Index           =   2
         Left            =   90
         TabIndex        =   38
         Top             =   720
         Width           =   2090
         _ExtentX        =   3678
         _ExtentY        =   14367
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
         Index           =   2
         Left            =   1350
         Top             =   660
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
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   8
      Left            =   12360
      MaxLength       =   14
      TabIndex        =   10
      Tag             =   "P3|N|S|||ccconcostes|p3|#0||"
      Top             =   4980
      Width           =   345
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   7
      Left            =   11940
      MaxLength       =   14
      TabIndex        =   9
      Tag             =   "P2|N|S|||ccconcostes|p2|#0||"
      Top             =   4980
      Width           =   345
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
      Index           =   6
      Left            =   11520
      MaxLength       =   14
      TabIndex        =   8
      Tag             =   "P1|N|S|||ccconcostes|p1|#0||"
      Top             =   4980
      Width           =   345
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BeginProperty Font 
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
      Left            =   10740
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Tag             =   "Tipo de Kilos|N|N|||ccconcostes|tipokilos|||"
      Top             =   4950
      Width           =   735
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
      Index           =   5
      Left            =   9270
      MaxLength       =   14
      TabIndex        =   6
      Tag             =   "Importe|N|S|||ccconcostes|importe|###,###,##0.00||"
      Top             =   4950
      Width           =   1395
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
      Index           =   4
      Left            =   8040
      MaxLength       =   10
      TabIndex        =   5
      Tag             =   "Cta.Contable|T|S|||ccconcostes|ctacontable|||"
      Top             =   4950
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   4590
      TabIndex        =   19
      Top             =   9045
      Width           =   9105
      Begin VB.TextBox txtAux2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   360
         Index           =   4
         Left            =   1980
         MaxLength       =   30
         TabIndex        =   20
         Top             =   225
         Width           =   6390
      End
      Begin VB.Label Label1 
         Caption         =   "Cta. Contable"
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
         Left            =   225
         TabIndex        =   21
         Top             =   225
         Width           =   2085
      End
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BeginProperty Font 
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
      Left            =   7260
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Tag             =   "Tipo de Coste|N|N|||ccconcostes|tipocoste|||"
      Top             =   4950
      Width           =   735
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
      Index           =   3
      Left            =   6030
      MaxLength       =   4
      TabIndex        =   3
      Tag             =   "Codigo Incidencia|N|S|||ccconcostes|codincid|0000||"
      Top             =   4950
      Width           =   1125
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
      Left            =   4140
      TabIndex        =   18
      Top             =   4950
      Visible         =   0   'False
      Width           =   1815
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
      Left            =   3930
      MaskColor       =   &H00000000&
      TabIndex        =   17
      ToolTipText     =   "Buscar área"
      Top             =   4950
      Visible         =   0   'False
      Width           =   195
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
      Index           =   2
      Left            =   2490
      MaxLength       =   4
      TabIndex        =   2
      Tag             =   "Codigo Área|N|N|||ccconcostes|codarea|0000||"
      Top             =   4950
      Width           =   1395
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
      Left            =   15210
      TabIndex        =   11
      Top             =   9375
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
      Left            =   16335
      TabIndex        =   12
      Top             =   9375
      Visible         =   0   'False
      Width           =   1065
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
      Left            =   900
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "Descripción|T|N|||ccconcostes|nomcoste|||"
      Top             =   4920
      Width           =   1395
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
      Left            =   60
      MaxLength       =   4
      TabIndex        =   0
      Tag             =   "Código|N|N|0|9999|ccconcostes|codcoste|0000|S|"
      Top             =   4920
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmCCManConcep.frx":003C
      Height          =   8145
      Left            =   90
      TabIndex        =   15
      Top             =   855
      Width           =   13580
      _ExtentX        =   23945
      _ExtentY        =   14367
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
      Left            =   16335
      TabIndex        =   16
      Top             =   9360
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   675
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   9030
      Width           =   3735
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
         Height          =   255
         Left            =   45
         TabIndex        =   14
         Top             =   240
         Width           =   3555
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
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
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
Attribute VB_Name = "frmCCManConcep"
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
' 3. En la funció BotonAnyadir() canviar la taula i el camp per a SugerirCodigoSiguienteStr
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
Private CadB As String

Private WithEvents frmGru As frmManGrupos 'grupos de ccconcostes
Attribute frmGru.VB_VarHelpID = -1
Private WithEvents frmBas As frmBasico ' Ayuda de Areas
Attribute frmBas.VB_VarHelpID = -1
Private WithEvents frmCtas As frmCtasConta 'cuentas contables
Attribute frmCtas.VB_VarHelpID = -1


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
Dim i As Integer
Dim indCodigo As Integer

Dim NumTabMto As Integer 'Indica quin nº de Tab està en modo Mantenimient

Dim ModoLineas As Byte
'1.- Afegir,  2.- Modificar,  3.- Borrar,  0.-Passar control a Llínies

Private Sub PonerModo(vModo)
Dim b As Boolean

    Modo = vModo

    b = (Modo = 2) Or (Modo = 5)
    If b Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo, ModoLineas
    End If
    
    For i = 0 To 8 'txtAux.Count - 1
        txtAux(i).visible = Not b
    Next i
    
    txtAux(3).visible = False
    txtAux(5).visible = False
    txtAux(4).visible = False
    
    txtAux2(2).visible = Not b
    btnBuscar(0).visible = Not b
    btnBuscar(2).visible = Not b
    Combo1(0).visible = Not b
    Combo1(1).visible = Not b
    
    txtAux(42).visible = (Modo = 5) And NumTabMto = 1
    txtAux(52).visible = (Modo = 5) And NumTabMto = 2
    btnBuscar(2).visible = (Modo = 5) And NumTabMto = 2
    btnBuscar(2).Enabled = (Modo = 5) And NumTabMto = 2
    
    cmdAceptar.visible = Not b Or Modo = 5
    cmdCancelar.visible = Not b Or Modo = 5
    DataGrid1.Enabled = b
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = b
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu  'En funcion del usuario
    
    Me.FrameAux1.Enabled = (Modo <> 1)
    If Not adodc1.Recordset.EOF Then
        Me.FrameAux2.Enabled = (Modo <> 1) And Me.adodc1.Recordset!TipoCoste
    Else
        Me.FrameAux2.Enabled = (Modo <> 1)
    End If
    
    
    'Si estamos modo Modificar bloquear clave primaria
    BloquearTxt txtAux(0), (Modo = 4)
End Sub


Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botones del la toobar y del menu, segun el modo en que estemos
Dim b As Boolean
Dim bAux As Boolean

    b = (Modo = 2)
    'Busqueda
    Toolbar1.Buttons(5).Enabled = b
    Me.mnBuscar.Enabled = b
    'Ver Todos
    Toolbar1.Buttons(6).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    'Insertar
    Toolbar1.Buttons(1).Enabled = b And Not DeConsulta
    Me.mnNuevo.Enabled = b And Not DeConsulta
    
    b = (b And adodc1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnModificar.Enabled = b
    'Eliminar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnEliminar.Enabled = b
    'Imprimir
    Toolbar1.Buttons(8).Enabled = b
    Me.mnImprimir.Enabled = b
    
    b = (Modo = 4 Or Modo = 2) And Not DeConsulta
    For i = 1 To 2 'ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.AdoAux(i).Recordset.RecordCount > 0)
'        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i
    
    
    
End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    
    CargaGrid 'primer de tot carregue tot el grid
    CadB = ""
    '******************** canviar taula i camp **************************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        NumF = NuevoCodigo
    Else
        NumF = SugerirCodigoSiguienteStr("ccconcostes", "codcoste")
    End If
    
    PonerModo 3
    
    '********************************************************************
    'Situamos el grid al final
    AnyadirLinea DataGrid1, adodc1
         
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 240
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
    End If
    txtAux(0).Text = NumF
    FormateaCampo txtAux(0)
    For i = 1 To 5
        txtAux(i).Text = ""
    Next i
    txtAux2(2).Text = ""

    LLamaLineas anc, 3 'Pone el form en Modo=3, Insertar
       
    Combo1(0).ListIndex = 0
    Combo1(1).ListIndex = 0
    
    'Ponemos el foco
    PonerFoco txtAux(0)
End Sub

Private Sub BotonVerTodos()
    CadB = ""
    CargaGrid ""
    PonerModo 2
End Sub

Private Sub BotonBuscar()
    ' ***************** canviar per la clau primaria ********
    CargaGrid "ccconcostes.codcoste = -1"
    CargaGridAux 1, False
    CargaGridAux 2, False

    '*******************************************************************************
    'Buscar
    For i = 0 To 8 'txtAux.Count - 1
        txtAux(i).Text = ""
    Next i
    Combo1(0).ListIndex = -1
    Combo1(1).ListIndex = -1
    txtAux2(2).Text = ""
    txtAux2(4).Text = ""
    
    LLamaLineas DataGrid1.Top + 240, 1 'Pone el form en Modo=1, Buscar
    PonerFoco txtAux(0)
End Sub

Private Sub BotonModificar()
    Dim anc As Single
    Dim i As Integer
    
    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.Top '+ 545
    End If

    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    txtAux(2).Text = DataGrid1.Columns(2).Text
    txtAux(3).Text = DataGrid1.Columns(4).Text
    'cta contable
    txtAux(4).Text = DataGrid1.Columns(7).Text
    'importe
    txtAux(5).Text = DataGrid1.Columns(8).Text
    'Puntos dentro del informe
    txtAux(6).Text = DataGrid1.Columns(11).Text
    txtAux(7).Text = DataGrid1.Columns(12).Text
    txtAux(8).Text = DataGrid1.Columns(13).Text
    
    ' ***** canviar-ho pel nom del camp del combo *********
'    SelComboBool DataGrid1.Columns(2).Text, Combo1(0)
    ' *****************************************************
    ' ### [Monica] 12/09/2006
    txtAux2(2).Text = DataGrid1.Columns(3).Text

    i = adodc1.Recordset!TipoCoste
    ' *****************************************************
    PosicionarCombo Me.Combo1(0), i

    i = adodc1.Recordset!tipokilos
    ' *****************************************************
    PosicionarCombo Me.Combo1(1), i


    LLamaLineas anc, 4 'Pone el form en Modo=4, Modificar
   
    'Como es modificar
    PonerFoco txtAux(1)
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    

    'Fijamos el ancho
    For i = 0 To 8
        txtAux(i).Top = alto
    Next i

    ' ### [Monica] 12/09/2006
    txtAux2(2).Top = alto
    btnBuscar(0).Top = alto - 15
    btnBuscar(2).Top = alto - 15
    Combo1(0).Top = alto
    Combo1(1).Top = alto

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
    Sql = "¿Seguro que desea eliminar el Concepto de Coste?"
    Sql = Sql & vbCrLf & "Código: " & adodc1.Recordset.Fields(0)
    Sql = Sql & vbCrLf & "Descripción: " & adodc1.Recordset.Fields(1)
    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        
        Sql = "Delete from ccconcostes_lin where codcoste=" & adodc1.Recordset!codCoste
        conn.Execute Sql
        
        Sql = "Delete from ccconcostes_cta where codcoste=" & adodc1.Recordset!codCoste
        conn.Execute Sql
        
        Sql = "Delete from ccconcostes where codcoste=" & adodc1.Recordset!codCoste
        conn.Execute Sql
        CargaGrid CadB
        
        temp = SituarDataTrasEliminar(adodc1, NumRegElim, True)
        PonerModoOpcionesMenu
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
        Case 0 'area
            
            indice = Index + 2
            Set frmBas = New frmBasico
            frmBas.DatosADevolverBusqueda = "0|1|"
            frmBas.DeConsulta = True
            frmBas.CodigoActual = txtAux(indice).Text
            frmBas.CadenaTots = "S|txtAux(0)|T|Código|800|;S|txtAux(1)|T|Descripción|3930|;"
            frmBas.CadenaConsulta = "SELECT ccareas.codarea, ccareas.nomarea "
            frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM ccareas "
            frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
            frmBas.Tag1 = "Código|N|N|0|9999|ccareas|codarea|0000|S|"
            frmBas.Tag2 = "Descripción|T|N|||ccareas|nomarea|||"
            frmBas.Maxlen1 = 4
            frmBas.Maxlen2 = 50
            frmBas.Tabla = "ccareas"
            frmBas.CampoCP = "codarea"
            frmBas.Report = "rManCCAreas.rpt"
            frmBas.Caption = "Áreas"
            
            frmBas.Show vbModal
            Set frmBas = Nothing
            PonerFoco txtAux(indice)
    
        Case 2 'Cuentas Contables (de contabilidad)
            AbrirFrmCuentas Index + 50
            
            PonerFoco txtAux(52)
            
    
    
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Me.adodc1, 1
End Sub


Private Sub cmdAceptar_Click()
    Dim i As Integer

    Select Case Modo
        Case 1 'BUSQUEDA
            CadB = ObtenerBusqueda2(Me, , 1)
            If CadB <> "" Then
                CargaGrid CadB
                PonerModo 2
'                lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
                PonerFocoGrid Me.DataGrid1
            End If
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    CargaGrid
                    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                        cmdCancelar_Click
'                        If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveLast
                        If Not adodc1.Recordset.EOF Then
                            adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & NuevoCodigo)
                        End If
                        cmdRegresar_Click
                    Else
                        PosicionarData
                        Me.DataGrid1.AllowAddNew = False
                        BotonAnyadirLinea 1
                    End If
                    CadB = ""
                End If
            End If
            
        Case 4 'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me) Then
                    TerminaBloquear
                    i = adodc1.Recordset.Fields(0)
                    PonerModo 2
                    CargaGrid CadB
'                    If CadB <> "" Then
'                        CargaGrid CadB
'                        lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
'                    Else
'                        CargaGrid
'                        lblIndicador.Caption = ""
'                    End If
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & i)
                    PonerFocoGrid Me.DataGrid1
                End If
            End If
        Case 5 'LLÍNIES
            
            Select Case ModoLineas
                Case 1 'afegir llínia
                    InsertarLinea
                Case 2 'modificar llínies
                    If ModificarLinea Then
                        PosicionarData
                    Else
                        PonerFoco txtAux(5)
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
            CargaGrid CadB
        Case 3 'insertar
            DataGrid1.AllowAddNew = False
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
            PonerModo 2
            DataGrid1_RowColChange 1, 0
        Case 4 'modificar
            TerminaBloquear
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 1 'afegir llínia
                
                    ModoLineas = 0
                    ' *** les llínies que tenen datagrid (en o sense tab) ***
                    DataGridAux(NumTabMto).AllowAddNew = False
                    ' **** repasar si es diu Data1 l'adodc de la capçalera ***
                    'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'Modificar
                    LLamaLineas2 NumTabMto, ModoLineas 'ocultar txtAux
                    DataGridAux(NumTabMto).Enabled = True
                    DataGridAux(NumTabMto).SetFocus

                    ' *** si n'hi han camps de descripció dins del grid, els neteje ***
                    'txtAux2(2).text = ""
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        AdoAux(NumTabMto).Recordset.MoveFirst
                    End If

                Case 2 'modificar llínies
                    ModoLineas = 0
                    
                    ' *** si n'hi han tabs ***
'                    SituarTab (NumTabMto + 1)
                    LLamaLineas 1, ModoLineas 'ocultar txtAux
                    PonerModo 4
                    If Not AdoAux(1).Recordset.EOF Then
                        ' *** l'Index de Fields es el que canvie de la PK de llínies ***
                        V = AdoAux(1).Recordset.Fields(1) 'el 2 es el nº de llinia
                        AdoAux(1).Recordset.Find (AdoAux(1).Recordset.Fields(1).Name & " =" & V)
                        ' ***************************************************************
                    End If
            End Select
            
            TerminaBloquear
    End Select
    
    PonerModo 2
    
    PonerFocoGrid Me.DataGrid1
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String
Dim i As Integer
Dim J As Integer
Dim Aux As String

    If adodc1.Recordset.EOF Then
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
            Cad = Cad & adodc1.Recordset.Fields(J) & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    PonerContRegIndicador
'    CargaForaGrid
 
    If Modo = 3 Then
        CargaGridAux 1, False
        CargaGridAux 2, False
    Else
        CargaGridAux 1, True
        CargaGridAux 2, True
    End If
    If Not adodc1.Recordset.EOF Then
        FrameAux2.Enabled = Me.adodc1.Recordset!TipoCoste
    End If
    
    
    PonerModoOpcionesMenu
    PonerOpcionesMenu

End Sub

Private Sub DataGridAux_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
    If Index = 2 Then
        CargaForaGrid
    End If
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
                SituarData Me.adodc1, "codcoste=" & CodigoActual, "", True
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
        .Buttons(5).Image = 1   'Buscar
        .Buttons(6).Image = 2   'Todos
        'el 4 i el 5 son separadors
        .Buttons(1).Image = 3   'Insertar
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Borrar
        'el 9 i el 10 son separadors
        .Buttons(8).Image = 10  'imprimir
    End With

    '## A mano
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    
    ' ******* si n'hi han llínies *******
    'ICONETS DE LES BARRES ALS TABS DE LLÍNIA
    For i = 1 To 2
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
'    LimpiarCampos   'Neteja els camps TextBox
    ' ******* si n'hi han llínies *******
    DataGridAux(1).ClearFields
    DataGridAux(2).ClearFields
    
    CargaCombo
    
    '****************** canviar la consulta *********************************+
    CadenaConsulta = "SELECT ccconcostes.codcoste, ccconcostes.nomcoste, ccconcostes.codarea,"
    CadenaConsulta = CadenaConsulta & "ccareas.nomarea, ccconcostes.codincid, ccconcostes.tipocoste, cctipocoste.descripcion, "
    CadenaConsulta = CadenaConsulta & "ccconcostes.ctacontable, ccconcostes.importe, ccconcostes.tipokilos, "
    CadenaConsulta = CadenaConsulta & " CASE ccconcostes.tipokilos WHEN 0 THEN ""Entrados"" WHEN 1 THEN ""Volcados"" WHEN 2 THEN ""Confeccionados"" WHEN 3 THEN ""No Procesa"" END, "
    CadenaConsulta = CadenaConsulta & " p1, p2, p3"
    CadenaConsulta = CadenaConsulta & " FROM ccconcostes, ccareas, cctipocoste"
    CadenaConsulta = CadenaConsulta & " WHERE ccconcostes.codarea = ccareas.codarea "
    CadenaConsulta = CadenaConsulta & " and ccconcostes.tipocoste = cctipocoste.tipocoste "
    '************************************************************************
    
    CadB = ""
    CargaGrid
    
    ' ****** Si n'hi han camps fora del grid ******
'    CargaForaGrid
    ' *********************************************
    If Me.adodc1.Recordset.RecordCount <> 0 Then
        CargaGridAux 1, True
        CargaGridAux 2, True
    Else
        CargaGridAux 1, False
        CargaGridAux 2, False
    End If
'    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
'        BotonAnyadir
'    Else
'        PonerModo 2
'    End If

    ModoLineas = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    If Modo = 4 Then TerminaBloquear
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmBas_DatoSeleccionado(CadenaSeleccion As String)
'Ayuda de Areas de Costes
    txtAux(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'codarea
    txtAux2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre de area
End Sub


Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    txtAux(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    txtAux2(4).Text = RecuperaValor(CadenaSeleccion, 2) 'des macta
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 5
                mnBuscar_Click
        Case 6
                mnVerTodos_Click
        Case 1
                mnNuevo_Click
        Case 2
                mnModificar_Click
        Case 3
                mnEliminar_Click
        Case 8
                'MsgBox "Imprimir...under construction"
                mnImprimir_Click
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
    Sql = Sql & " ORDER BY ccconcostes.codcoste"
    '**************************************************************++
    
    CargaGridGnral Me.DataGrid1, Me.adodc1, Sql, PrimeraVez
    
    ' *******************canviar els noms i si fa falta la cantitat********************
    tots = "S|txtAux(0)|T|Código|900|;S|txtAux(1)|T|Descripción|3875|;"
    tots = tots & "S|txtAux(2)|T|Area|900|;"
    tots = tots & "S|btnBuscar(0)|B|||;S|txtAux2(2)|T|Nombre de Area|2675|;N|txtAux(3)|T|Línea|900|;N||||0|;S|Combo1(0)|C|Tipo|1600|;"
    tots = tots & "N|txtAux(4)|T|Cta.Cble.|1100|;N|btnBuscar(2)|B|||;N|txtAux(5)|T|Importe|1100|;N||||0|;S|Combo1(1)|C|Tipo Kilos|1870|;"
    tots = tots & "S|txtAux(6)|T|P1|400|;S|txtAux(7)|T|P2|400|;S|txtAux(8)|T|P3|400|;"
    
    arregla tots, DataGrid1, Me, 350
    
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.Columns(0).Alignment = dbgRight
    DataGrid1.Columns(2).Alignment = dbgLeft
    DataGrid1.Columns(4).Alignment = dbgLeft
    
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
        
        Case 2 'codigo de area de coste
            If txtAux(Index).Text = "" Then Exit Sub
            txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), "ccareas", "nomarea", "codarea", "N")
        
        Case 3
            PonerFormatoEntero txtAux(Index)
            
        Case 52   ' cta contable
            If txtAux(Index).Text <> "" Then
                txtAux2(4).Text = PonerNombreCuenta(txtAux(Index), 2)
                If txtAux2(4).Text = "" Then
                    MsgBox "Número de Cuenta contable no existe en la contabilidad. Reintroduzca.", vbExclamation
                End If
            End If
        
        Case 5 ' importe
            PonerFormatoDecimal txtAux(Index), 1
            
    End Select
    
End Sub

Private Function DatosOk() As Boolean
'Dim Datos As String
Dim b As Boolean
Dim Sql As String
Dim Mens As String
Dim TipoCoste As Byte


    b = CompForm(Me)
    If Not b Then Exit Function
    
    If Modo = 3 Then   'Estamos insertando
         If ExisteCP(txtAux(0)) Then b = False
    End If
    
    If b And (Modo = 3 Or Modo = 4) Then
        Sql = DevuelveDesdeBDNew(cAgro, "ccareas", "nomarea", "codarea", txtAux(2).Text, "N")
        If Sql = "" Then
            MsgBox "No existe el código de Área introducida. Revise.", vbExclamation
            PonerFoco txtAux(2)
            b = False
        End If
    End If
    
    If b And (Modo = 3 Or Modo = 4) Then
'[Monica]11/06/2012: antes era codigo de incidencia, ahora es linea de coste
'        SQL = "select count(*) from ccconcostes where codcoste <> " & DBSet(txtAux(0).Text, "N")
'        SQL = SQL & " and codincid = " & DBSet(txtAux(3).Text, "N")
'
'        If TotalRegistros(SQL) <> 0 Then
'            MsgBox "El numero de incidencia del reloj está asignada a otro coste. Revise.", vbExclamation
'            PonerFoco txtAux(3)
'            b = False
'        End If
        
        '[Monica]10/04/2012: comprobamos segun sea el tipo de coste que me han introducido
        If b Then
            Select Case Combo1(0).ListIndex
                Case 0 ' coste directo
                       ' ni la cuenta ni el importe son requeridos, da lo mismo que los introduzcan
                        
'[Monica]21/11/2013: quitamos esto
'                Case 1 ' contabilidad
'
'                       ' pedimos obligatoriamente la cuenta contable
'                       If txtAux(4).Text = "" Then
'                            MsgBox "Debe introducir obligatoriamente la cta contable. Revise.", vbExclamation
'                            b = False
'                            PonerFoco txtAux(4)
'                       Else
'                            ' comprobamos que exista la cuenta
'                            txtAux2(4).Text = PonerNombreCuenta(txtAux(4), 2)
'                            If txtAux2(4).Text = "" Then
'                                MsgBox "Número de Cuenta contable no existe en la contabilidad. Reintroduzca.", vbExclamation
'                                b = False
'                            End If
'                       End If
'21/11/2013: hasta aqui
                
'                Case 2 ' periodificado
'                       ' pedimos obligatoriamente el importe
'                       If txtAux(5).Text = "" Then
'                            MsgBox "Debe introducir obligatoriamente el Importe. Revise.", vbExclamation
'                            b = False
'                            PonerFoco txtAux(5)
'                       End If
            
            End Select
        
        End If
        
        If b And Modo = 3 Then
            If (CLng(txtAux(0).Text) < 11 Or CLng(txtAux(0).Text) > 255) And Combo1(0).ListIndex = 0 Then
                MsgBox "Los códigos de concepto deben estar comprendidos entre 11 y 255. Reintroduzca.", vbExclamation
                b = False
                PonerFoco txtAux(0)
            End If
        End If
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
        If CadB = "" Then
            lblIndicador.Caption = cadReg
        Else
            lblIndicador.Caption = "BUSQUEDA: " & cadReg
        End If
    End If
End Sub

Private Sub printNou()
    With frmImprimir2
        .cadTabla2 = "ccconcostes"
        .Informe2 = "rManCCConceptos.rpt"
        If CadB <> "" Then
            '.cadRegSelec = Replace(SQL2SF(CadB), "clientes", "clientes_1")
            .cadRegSelec = SQL2SF(CadB)
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
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomempre & "'|pOrden={ccconcostes.codcoste}|"
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

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
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


Private Sub CargaCombo()
Dim Cad As String
Dim Rs As ADODB.Recordset
Dim i As Integer

    On Error GoTo ErrCarga
    Combo1(0).Clear
    'Conceptos
    
    Cad = "SELECT tipocoste, descripcion FROM cctipocoste ORDER BY tipocoste"
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    'Combo1.AddItem "" 'pose uno en blanc sinse valor
    i = 0
    While Not Rs.EOF
        Combo1(0).AddItem Rs!Descripcion
        Combo1(0).ItemData(Combo1(0).NewIndex) = Rs!TipoCoste
        Rs.MoveNext
        '.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    Combo1(1).Clear
    
    Combo1(1).AddItem "Entrados"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Volcados"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    Combo1(1).AddItem "Confeccionados"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 2
    Combo1(1).AddItem "No Procesa"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 3
    
    Exit Sub
    
ErrCarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar datos combo.", Err.Description
End Sub

Private Sub CargaForaGrid()
    
    If Me.AdoAux(2).Recordset.EOF Then Exit Sub
    
    txtAux2(4).Text = ""
    
    If IsNull(Me.AdoAux(2).Recordset.Fields(2).Value) Then Exit Sub
    
    txtAux2(4).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Me.AdoAux(2).Recordset.Fields(2).Value, "T")

End Sub

Private Sub AbrirFrmCuentas(indice As Integer)
    indCodigo = indice
    Set frmCtas = New frmCtasConta
    frmCtas.DatosADevolverBusqueda = "0|1|"
    frmCtas.CodigoActual = txtAux(indCodigo)
    frmCtas.Show vbModal
    Set frmCtas = Nothing
End Sub


'****************************************************
Private Sub CargaGridAux(Index As Integer, enlaza As Boolean)
Dim b As Boolean
Dim i As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)

    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
    
    Select Case Index
        Case 1 'lineas
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;" 'codorden,numlinea
            tots = tots & "S|txtAux(42)|T|Linea|620|;S|btnBuscar(1)|B||195|;"
'            tots = tots & "S|txtAux2(42)|T|Denominación|3050|;"
            
            arregla tots, DataGridAux(Index), Me, 350
            
            DataGridAux(1).Columns(2).Alignment = dbgLeft
            
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
        Case 2 'ctas contables
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;" 'codorden,numlinea
            tots = tots & "S|txtAux(52)|T|Cta.Cble|1400|;S|btnBuscar(2)|B||195|;"
'            tots = tots & "S|txtAux2(42)|T|Denominación|3050|;"
            
            arregla tots, DataGridAux(Index), Me, 350
            
            DataGridAux(2).Columns(2).Alignment = dbgLeft
            
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
    End Select
    
    DataGridAux(Index).ScrollBars = dbgAutomatic
      
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
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
        Case 1 ' LINEAS
            Sql = "SELECT ccconcostes_lin.codcoste, ccconcostes_lin.numlinea, ccconcostes_lin.codlinea  "
            Sql = Sql & " FROM ccconcostes_lin " '(cclinorden2 INNER JOIN ccconcostes ON cclinorden2.codcoste = ccconcostes.codcoste) "
'            SQL = SQL & " INNER JOIN salarios ON cclinorden2.codcateg = salarios.codcateg "
            
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE ccconcostes_lin.codcoste is null "
            End If
            Sql = Sql & " ORDER BY ccconcostes_lin.numlinea"
    
        Case 2 ' ctas contables
            Sql = "SELECT ccconcostes_cta.codcoste, ccconcostes_cta.numlinea, ccconcostes_cta.ctacontable  "
            Sql = Sql & " FROM ccconcostes_cta " '(cclinorden2 INNER JOIN ccconcostes ON cclinorden2.codcoste = ccconcostes.codcoste) "
'            SQL = SQL & " INNER JOIN salarios ON cclinorden2.codcateg = salarios.codcateg "
            
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE ccconcostes_cta.codcoste is null "
            End If
            Sql = Sql & " ORDER BY ccconcostes_cta.numlinea"
    
    End Select
    
    MontaSQLCarga = Sql
End Function

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " codcoste=" & Me.adodc1.Recordset!codCoste
    
    ObtenerWhereCab = vWhere
End Function




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
       
'    PonerModo 5

    If AdoAux(Index).Recordset.EOF Then Exit Sub
    Eliminar = False
   
    vWhere = ObtenerWhereCab(True)
    
    ' ***** independentment de si tenen datagrid o no,
    ' canviar els noms, els formats i el DELETE *****
    Select Case Index
        Case 1 'lineas de confeccion
            Sql = "¿ Seguro que desea eliminar la linea ?"
            Sql = Sql & vbCrLf & "Linea: " & AdoAux(Index).Recordset!codlinea '& " " & Adoaux(Index).Recordset!NomTraba
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM ccconcostes_lin "
                Sql = Sql & vWhere & " AND numlinea= " & AdoAux(Index).Recordset!NumLinea
            End If
            
        Case 2 'lineas de cuentas contables
            Sql = "¿ Seguro que desea eliminar la cuenta asociada ?"
            Sql = Sql & vbCrLf & "Cuenta Contable: " & AdoAux(Index).Recordset!ctacontable '& " " & Adoaux(Index).Recordset!NomTraba
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM ccconcostes_cta "
                Sql = Sql & vWhere & " AND numlinea= " & AdoAux(Index).Recordset!NumLinea
            End If
    End Select

    If Eliminar Then
        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        conn.Execute Sql
        
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        CargaGridAux Index, True
        If Not SituarDataTrasEliminar(AdoAux(Index), NumRegElim, True) Then
            
        End If
    End If
    
    ModoLineas = 0
    
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
    
    
    NumTabMto = Index
    
    ModoLineas = 1 'Posem Modo Afegir Llínia
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    PonerModo 5
    
    ' *** bloquejar la clau primaria de la capçalera ***
'    BloquearTxt text1(0), True

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case Index
        Case 1: vtabla = "ccconcostes_lin"
        Case 2: vtabla = "ccconcostes_cta"
    End Select
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 1, 2 ' *** pose els index dels tabs de llínies que tenen datagrid ***
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
            
            LLamaLineas2 Index, ModoLineas, anc
        
            Select Case Index
                ' *** valor per defecte a l'insertar i formateig de tots els camps ***
                Case 1 'lineas de confeccion
                    txtAux(40).Text = Me.adodc1.Recordset!codCoste  'codorden
                    txtAux(41).Text = NumF 'numlinea
                    txtAux(42).Text = ""
                    
                    BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux1"
'                    btnBuscar(1).visible = False
'                    btnBuscar(1).Enabled = False
                    
                    PonerFoco txtAux(42)
                    
                    
                Case 2 'cuentas contables
                    txtAux(50).Text = Me.adodc1.Recordset!codCoste  'codorden
                    txtAux(51).Text = NumF 'numlinea
                    txtAux(52).Text = ""
                    
                    BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux2"
                    btnBuscar(2).visible = True
                    btnBuscar(2).Enabled = True
                    
                    PonerFoco txtAux(52)
                    
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
       
    PonerModo 5
    ' *** bloqueje la clau primaria de la capçalera ***
    BloquearTxt txtAux(0), True
  
    Select Case Index
        Case 1, 2 ' *** pose els index de llínies que tenen datagrid (en o sense tab) ***
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
        Case 1 ' lineas de confeccion
            For J = 40 To 42
                txtAux(J).Text = DataGridAux(Index).Columns(J).Text
            Next J
            txtAux2(42).Text = DataGridAux(Index).Columns(3).Text
            
            BloquearbtnBuscar Me, Modo, 1, "FrameAux0"
            btnBuscar(1).visible = False
            btnBuscar(1).Enabled = False
            
        Case 2
            For J = 50 To 52
                txtAux(J).Text = DataGridAux(Index).Columns(J).Text
            Next J
            txtAux2(4).Text = DataGridAux(Index).Columns(3).Text
            
            BloquearbtnBuscar Me, Modo, 1, "FrameAux1"
            btnBuscar(2).visible = True
            btnBuscar(2).Enabled = True
        
            
    End Select
    
    LLamaLineas2 Index, ModoLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 1 'lineas de conceptos
            PonerFoco txtAux(42)
        Case 2 ' ctas contables
            PonerFoco txtAux(52)
    End Select
    ' ***************************************************************************************
End Sub

Private Sub LLamaLineas2(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim b As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    DeseleccionaGrid DataGridAux(Index)
       
    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
    Select Case Index
        Case 1 'lineas de confeccion
             For jj = 42 To 42
                txtAux(jj).visible = b
                txtAux(jj).Top = alto
            Next jj
'            For jj = 2 To 3
'                txtAux2(jj).visible = b
'                txtAux2(jj).Top = alto
'            Next jj
            btnBuscar(1).visible = b
            btnBuscar(1).Top = alto
        Case 2
             For jj = 52 To 52
                txtAux(jj).visible = b
                txtAux(jj).Top = alto
            Next jj
'            For jj = 2 To 3
'                txtAux2(jj).visible = b
'                txtAux2(jj).Top = alto
'            Next jj
            btnBuscar(2).visible = b
            btnBuscar(2).Top = alto
        
    End Select
End Sub

Private Sub InsertarLinea()
'Inserta registre en les taules de Llínies
Dim nomFrame As String
Dim b As Boolean

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    
    Select Case NumTabMto
        Case 1
            nomFrame = "FrameAux1" 'lineas de conceptos
        Case 2
            nomFrame = "FrameAux2" 'cuentas contables
    End Select
    
    
    If DatosOkLlin(nomFrame) Then
        
        TerminaBloquear
        
        If InsertarDesdeForm2(Me, 2, nomFrame) Then
            b = BLOQUEADesdeFormulario2(Me, Me.adodc1, 1)
            ' *** els index de les llinies en grid (en o sense tab) ***
                CargaGridAux NumTabMto, True
                If b Then BotonAnyadirLinea NumTabMto
        End If
    End If
    
End Sub

Private Function ModificarLinea() As Boolean
'Modifica registre en les taules de Llínies
Dim nomFrame As String
Dim V As Integer
    
    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    nomFrame = "FrameAux0" 'lineas de confeccion
    
    ModificarLinea = False
    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomFrame) Then
            ModoLineas = 0
            
            V = AdoAux(0).Recordset.Fields(1) 'el 2 es el nº de llinia
            CargaGridAux 0, True
            
            ' *** si n'hi han tabs ***
'            SituarTab (NumTabMto + 1)

            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
            PonerFocoGrid Me.DataGridAux(0)
            AdoAux(0).Recordset.Find (AdoAux(0).Recordset.Fields(1).Name & " =" & V)
            
            LLamaLineas 0, 0
            ModificarLinea = True
        End If
    End If
End Function

Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    Cad = "(codcoste=" & DBSet(txtAux(0).Text, "N") & ")"
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(Data1, cad, Indicador) Then
    If SituarData(adodc1, Cad, Indicador) Then
        If ModoLineas <> 1 Then PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
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
    
    If b And nomFrame = "FrameAux1" Then
        'comprobamos que no exista ese codigo de linea
        If ModoLineas = 1 Then
            Sql = "select count(*) from ccconcostes_lin where codcoste = " & Me.adodc1.Recordset!codCoste
            Sql = Sql & " and codlinea = " & DBSet(txtAux(42).Text, "N")
            If TotalRegistros(Sql) <> 0 Then
                MsgBox "Código de linea ya existe. Revise.", vbExclamation
                b = False
            End If
        End If
    End If
    
    If b And nomFrame = "FramAux2" Then
        'comprobamos que no exista ese codigo de linea
        If ModoLineas = 1 Then
            Sql = "select count(*) from ccconcostes_cta where codcoste = " & Me.adodc1.Recordset!codCoste
            Sql = Sql & " and ctacontable = " & DBSet(txtAux(52).Text, "T")
            If TotalRegistros(Sql) <> 0 Then
                MsgBox "Cuenta contable ya existe. Revise.", vbExclamation
                b = False
            End If
        End If
    
    End If
    
    ' ******************************************************************************
    DatosOkLlin = b
    
EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub LimpiarCampos()
    On Error Resume Next
    
    limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub LimpiarCamposLin(FrameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, FrameAux  'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub


