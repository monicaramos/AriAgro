VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmVtasLinFacturas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calibres de Factura"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   11295
   Icon            =   "frmVtasLinFacturas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameAux0 
      Caption         =   "Calibres"
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
      Height          =   2775
      Left            =   120
      TabIndex        =   19
      Top             =   2070
      Width           =   11080
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
         Index           =   14
         Left            =   1710
         MaxLength       =   9
         TabIndex        =   39
         Tag             =   "Num.Linea1 albaran|N|N|||facturas_calibre|numline1albar|00|N|"
         Text            =   "Linea1alb"
         Top             =   1800
         Visible         =   0   'False
         Width           =   630
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
         Left            =   3150
         MaskColor       =   &H00000000&
         TabIndex        =   38
         ToolTipText     =   "Buscar Línea de Calibre"
         Top             =   1770
         Visible         =   0   'False
         Width           =   195
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
         Index           =   9
         Left            =   7785
         MaxLength       =   11
         TabIndex        =   45
         Tag             =   "Imp.Bruto|N|N|||facturas_calibre|imporbru|##,###,##0.00|N|"
         Text            =   "Imp.Brut"
         Top             =   1950
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.TextBox txtAux 
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
         Index           =   10
         Left            =   8415
         MaxLength       =   11
         TabIndex        =   46
         Tag             =   "Imp.Neto|N|N|||facturas_calibre|impornet|##,###,##0.00|N|"
         Text            =   "Imp.Neto"
         Top             =   1950
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
         Height          =   330
         Index           =   11
         Left            =   5940
         MaxLength       =   11
         TabIndex        =   37
         Tag             =   "Dto1|N|S|||facturas_calibre|dtocom1|##0.00|N|"
         Text            =   "Dto1"
         Top             =   2220
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
         Height          =   330
         Index           =   12
         Left            =   6660
         MaxLength       =   11
         TabIndex        =   36
         Tag             =   "Dto2|N|S|||facturas_calibre|dtocom2|##0.00|N|"
         Text            =   "Dto2"
         Top             =   2220
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
         Height          =   330
         Index           =   13
         Left            =   7380
         MaxLength       =   11
         TabIndex        =   35
         Text            =   "Iva"
         Top             =   2220
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
         Height          =   330
         Index           =   17
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   30
         Tag             =   "Num.Linea|N|N|||facturas_calibre|numlinea|000|S|"
         Text            =   "linea"
         Top             =   2100
         Visible         =   0   'False
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
         Height          =   330
         Index           =   16
         Left            =   1020
         MaxLength       =   10
         TabIndex        =   29
         Tag             =   "Fecha Factura|F|N|||facturas_calibre|fecfactu|dd/mm/yyyy|S|"
         Text            =   "fecha"
         Top             =   2100
         Visible         =   0   'False
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
         Height          =   330
         Index           =   15
         Left            =   210
         MaxLength       =   7
         TabIndex        =   28
         Tag             =   "Nº Factura|N|S|||facturas_calibre|numfactu|0000000|S|"
         Text            =   "factura"
         Top             =   2100
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.TextBox txtAux 
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
         Index           =   8
         Left            =   6630
         MaxLength       =   6
         TabIndex        =   42
         Tag             =   "Unidades|N|S|||facturas_calibre|unidades|#,##0||"
         Text            =   "unida"
         Top             =   1800
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.TextBox txtAux 
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
         Index           =   7
         Left            =   8460
         MaxLength       =   11
         TabIndex        =   44
         Tag             =   "Precio Neto|N|N|||facturas_calibre|precinet|###,##0.0000||"
         Text            =   "pr.neto"
         Top             =   1800
         Visible         =   0   'False
         Width           =   840
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
         Left            =   7470
         MaxLength       =   11
         TabIndex        =   43
         Tag             =   "Precio Bruto|N|N|||facturas_calibre|precibru|###,##0.0000||"
         Text            =   "precbrut"
         Top             =   1800
         Visible         =   0   'False
         Width           =   1080
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
         Index           =   5
         Left            =   2610
         MaxLength       =   9
         TabIndex        =   11
         Tag             =   "Num.Linea 1|N|N|||facturas_calibre|numline1|00|S|"
         Text            =   "Linea1"
         Top             =   2100
         Visible         =   0   'False
         Width           =   630
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
         Left            =   2430
         MaxLength       =   6
         TabIndex        =   8
         Tag             =   "Tipo Movimiento|T|N|||facturas_calibre|codtipom||S|"
         Text            =   "Tipom"
         Top             =   1800
         Visible         =   0   'False
         Width           =   720
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
         Index           =   3
         Left            =   1035
         MaxLength       =   9
         TabIndex        =   10
         Tag             =   "Num.Linea|N|N|||facturas_calibre|numlinealbar|00|N|"
         Text            =   "Lineaalb"
         Top             =   1800
         Visible         =   0   'False
         Width           =   630
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
         Left            =   3240
         TabIndex        =   24
         Top             =   1800
         Visible         =   0   'False
         Width           =   1815
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
         Index           =   2
         Left            =   5820
         MaxLength       =   9
         TabIndex        =   41
         Tag             =   "Cant.Fact|N|N|||facturas_calibre|cantfact|###,##0||"
         Text            =   "cant.Fact"
         Top             =   1800
         Visible         =   0   'False
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
         Height          =   330
         Index           =   1
         Left            =   5100
         MaxLength       =   8
         TabIndex        =   40
         Tag             =   "Cant.Real|N|N|||facturas_calibre|cantreal|###,##0||"
         Text            =   "cant.rea"
         Top             =   1800
         Visible         =   0   'False
         Width           =   690
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
         Index           =   0
         Left            =   210
         MaxLength       =   16
         TabIndex        =   9
         Tag             =   "Número Albaran|N|N|||facturas_calibre|numalbar|000000|N|"
         Text            =   "albaran"
         Top             =   1800
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   0
         Left            =   135
         TabIndex        =   20
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
         Left            =   3735
         Top             =   720
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
         Bindings        =   "frmVtasLinFacturas.frx":000C
         Height          =   2040
         Index           =   0
         Left            =   135
         TabIndex        =   21
         Top             =   630
         Width           =   10710
         _ExtentX        =   18891
         _ExtentY        =   3598
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
   Begin VB.Frame Frame2 
      Height          =   1530
      Index           =   0
      Left            =   135
      TabIndex        =   14
      Top             =   495
      Width           =   11065
      Begin VB.TextBox Text3 
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
         Left            =   1245
         MaxLength       =   15
         TabIndex        =   34
         Text            =   "Text1 7"
         Top             =   990
         Width           =   2655
      End
      Begin VB.TextBox Text3 
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
         Left            =   5370
         MaxLength       =   30
         TabIndex        =   33
         Text            =   "Text1 7"
         Top             =   990
         Width           =   4965
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
         Index           =   5
         Left            =   4125
         MaxLength       =   4
         TabIndex        =   5
         Tag             =   "Num.Linea|N|N|||facturas_calibre|numlinea|000|S|"
         Text            =   "1234567890123456"
         Top             =   540
         Width           =   630
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
         Index           =   4
         Left            =   2625
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Fecha Factura|F|N|||facturas_calibre|fecfactu|dd/mm/yyyy|S|"
         Text            =   "123"
         Top             =   540
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
         Index           =   3
         Left            =   1245
         MaxLength       =   7
         TabIndex        =   3
         Tag             =   "Nº Factura|N|S|||facturas_calibre|numfactu|0000000|S|"
         Text            =   "123456"
         Top             =   540
         Width           =   1140
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
         Index           =   2
         Left            =   120
         MaxLength       =   6
         TabIndex        =   2
         Tag             =   "Tipo Movimiento|T|N|||facturas_calibre|codtipom||S|"
         Text            =   "123456"
         Top             =   540
         Width           =   945
      End
      Begin VB.TextBox Text1 
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
         Index           =   1
         Left            =   5925
         MaxLength       =   4
         TabIndex        =   1
         Tag             =   "Linea Albaran|N|N|||facturas_calibre|numlinealbar|00|S|"
         Text            =   "1234657890123456798012345678901234567890"
         Top             =   540
         Width           =   540
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
         Left            =   4905
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Número albaran|N|N|||facturas_calibre|numalbar|000000|S|"
         Text            =   "albaran"
         Top             =   540
         Width           =   855
      End
      Begin VB.Label Label6 
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
         Index           =   2
         Left            =   4110
         TabIndex        =   32
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label Label6 
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
         Index           =   0
         Left            =   180
         TabIndex        =   31
         Top             =   1020
         Width           =   960
      End
      Begin VB.Label Label13 
         Caption         =   "Línea"
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
         Left            =   4125
         TabIndex        =   27
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label12 
         Caption         =   "Fec. Factura"
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
         Left            =   2625
         TabIndex        =   26
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo Fact."
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
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label label 
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
         Left            =   1245
         TabIndex        =   23
         Top             =   240
         Width           =   1290
      End
      Begin VB.Label Label6 
         Caption         =   "L.Alb"
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
         Left            =   5925
         TabIndex        =   22
         Top             =   240
         Width           =   780
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
         Index           =   0
         Left            =   4935
         TabIndex        =   15
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   4860
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
         TabIndex        =   13
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
      Left            =   10155
      TabIndex        =   7
      Top             =   4995
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
      Left            =   9000
      TabIndex        =   6
      Top             =   4995
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   480
      Top             =   4320
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
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   11295
      _ExtentX        =   19923
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
            Object.ToolTipText     =   "Último"
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
         TabIndex        =   18
         Top             =   90
         Width           =   1215
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
      Left            =   10170
      TabIndex        =   16
      Top             =   5010
      Visible         =   0   'False
      Width           =   1035
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
Attribute VB_Name = "frmVtasLinFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MONICA                   -+-+
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

Public TipoM As String
Public NumFactu As String
Public FecFactu As String
Public NumLinea As String
Public Albaran As String
Public Linea As String
Public Variedad As String
Public Confeccion As String
Public ImpDtoc As String
Public Dto1 As String
Public Dto2 As String
Public TipoIva As String

Public ModoExt As Byte

' *** declarar els formularis als que vaig a cridar ***
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1

Private WithEvents frmCali As frmManCalibres 'calibres
Attribute frmCali.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes  ' mensajes
Attribute frmMens.VB_VarHelpID = -1

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
Dim cadB As String

Dim KilosAnt As Currency
Dim CajasAnt As Currency
Dim ForfaitAnt As String
Dim CodPaletAnt As String
Dim TotPaletAnt As String



Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    '++monica
'    BloqueaRegistro "palets", "numpalet = " & Text1(0).Text
    
    Select Case Index
        Case 0 'calibres
            
            Set frmMens = New frmMensajes
            frmMens.OpcionMensaje = 23
            frmMens.cadWHERE = "numalbar = " & DBSet(Text1(0).Text, "N") & " and numlinea = " & DBSet(Text1(1).Text, "N")
            frmMens.Show vbModal
            Set frmMens = Nothing
            PonerFoco txtAux(1)
    End Select
    If Modo = 4 Then BloqueaRegistro "albaran", "numalbar = " & Text1(0).Text
    'BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub


Private Sub cmdAceptar_Click()
Dim b As Boolean
Dim V As Integer
Dim Forfait As String

    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BÚSQUEDA
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
                    BloqueaRegistro "albaran", "numalbar = " & Text1(0).Text
                    
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
                Modificar
                TerminaBloquear
                '++monica
                BloqueaRegistro "albaran", "numalbar = " & Text1(0).Text
                
                PosicionarData
            Else
                ModoLineas = 0
            End If
        ' *** si n'hi han llínies ***
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 1 'afegir llínia
                    If InsertarLinea Then
                        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
'                        PonerCadenaBusqueda
'                        b = BLOQUEADesdeFormulario2(Me, Data1, 1)
                        CargaGrid 0, True
'                        If b Then
                        BotonAnyadirLinea NumTabMto
                    End If
                
                Case 2 'modificar llínies
                    If ModificarLinea Then
                        ModoLineas = 0
                        
                        V = AdoAux(NumTabMto).Recordset.Fields(4) 'el 2 es el nº de llinia
                        
                        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
'                        PonerCadenaBusqueda
'                        b = BLOQUEADesdeFormulario2(Me, Data1, 1)
                        
                        CargaGrid NumTabMto, True
                        
                        PonerFocoGrid Me.DataGridAux(NumTabMto)
                        AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
                        
                        LLamaLineas NumTabMto, 0
                        
                        TerminaBloquear
                        '++monica
'                        BloqueaRegistro "albaran", "numalbar = " & Text1(0).Text
                        PosicionarData
                    Else
                        PonerFoco txtAux(1)
                    End If
            End Select

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
           
        Modo = ModoExt
        
        DatosADevolverBusqueda = "ZZ"
        PonerModo Modo
        CargaGrid 0, True
        
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim cad As String

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
Dim i As Integer

    PrimeraVez = True
    
    ' ICONETS DE LA BARRA
    btnPrimero = 16 'index del botó "primero"
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
        .Buttons(11).Image = 19   'Expandir Añadir, Borrar y Modificar
        'el 10 i el 11 son separadors
        .Buttons(12).Image = 10  'Imprimir
        .Buttons(13).Image = 11  'Eixir
        'el 13 i el 14 son separadors
        .Buttons(btnPrimero).Image = 6  'Primer
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Següent
        .Buttons(btnPrimero + 3).Image = 9 'Últim
    End With
    
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
    LimpiarCampos   'Neteja els camps TextBox
    ' ******* si n'hi han llínies *******
    DataGridAux(0).ClearFields
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "facturas_calibre"
    Ordenacion = " ORDER BY numline1"
    
    'Mirem com està guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    Data1.ConnectionString = conn
    '***** cambiar el nombre de la PK de la cabecera *************
    
    CadenaConsulta = "Select * from " & NombreTabla & " where codtipom = " & DBSet(TipoM, "T")
    CadenaConsulta = CadenaConsulta & " and numfactu = " & DBSet(NumFactu, "N")
    CadenaConsulta = CadenaConsulta & " and fecfactu = " & DBSet(FecFactu, "F")
    CadenaConsulta = CadenaConsulta & " and numlinea = " & DBSet(NumLinea, "N")
    
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    Text1(2).Text = TipoM
    Text1(3).Text = NumFactu
    Text1(4).Text = FecFactu
    Text1(5).Text = NumLinea
    Text1(0).Text = Albaran
    Text1(1).Text = Linea
    Text3(3).Text = Variedad
    Text3(4).Text = Confeccion
    
    CargaGrid 0, False
End Sub

Private Sub LimpiarCampos()
    On Error Resume Next
    
    limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    

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
       
    'Modo 2. N'hi han datos i estam visualisant-los
    '=========================================
    'Posem visible, si es formulari de búsqueda, el botó "Regresar" quan n'hi han datos
'    If DatosADevolverBusqueda <> "" Then
'        cmdRegresar.visible = (Modo = 2)
'    Else
'        cmdRegresar.visible = False
'    End If
    
    Text1(5).Enabled = True
    
    
    '=======================================
    b = (Modo = 2)
    'Posar Fleches de desplasament visibles
    Numreg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then Numreg = 2 'Només es per a saber que n'hi ha + d'1 registre
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, Numreg
    '---------------------------------------------
    
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    cmdRegresar.visible = Not b

    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
    'Si estem en Insertar a més neteja els camps Text1
    BloquearText1 Me, Modo
    
    '*** si n'hi han combos a la capçalera ***
    '**************************
    
    ' ***** bloquejar tots els controls visibles de la clau primaria de la capçalera ***
    For i = 0 To 5
        BloquearTxt Text1(i), True, True 'si estic en  modificar, bloqueja la clau primaria
    Next i
    ' **********************************************************************************
    
    ' **** si n'hi han imagens de buscar en la capçalera *****
    BloquearImgBuscar Me, Modo, ModoLineas
    BloquearImgZoom Me, Modo, ModoLineas
    ' ********************************************************
'    imgBuscar(0).visible = (Modo = 3)
'    imgBuscar(0).Enabled = (Modo = 3)
    
        
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    PonerLongCampos

    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 0, False
    End If
    
    b = (Modo = 4) Or (Modo = 2)
    DataGridAux(0).Enabled = b
      
    For i = 0 To txtAux.Count - 1
        txtAux(i).visible = False
        BloquearTxt txtAux(i), True
    Next i
      
    ' ****** si n'hi han combos a la capçalera ***********************
    ' ****************************************************************
    
    PonerModoOpcionesMenu (Modo) 'Activar opcions menú según modo
    PonerOpcionesMenu   'Activar opcions de menú según nivell
                        'de permisos de l'usuari

EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
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
    DesplazamientoData Data1, Index
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
Dim Tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
               
        Case 0 'CALIBRES
            SQL = "SELECT facturas_calibre.codtipom, facturas_calibre.numfactu, facturas_calibre.fecfactu, facturas_calibre.numlinea , facturas_calibre.numline1 ,facturas_calibre.numalbar, facturas_calibre.numlinealbar, facturas_calibre.numline1albar, "
            SQL = SQL & " calibres.nomcalib, facturas_calibre.cantreal, facturas_calibre.cantfact, facturas_calibre.unidades, facturas_calibre.precibru, facturas_calibre.precinet, facturas_calibre.imporbru,"
            SQL = SQL & "facturas_calibre.impornet, facturas_calibre.dtocom1, facturas_calibre.dtocom2"
            SQL = SQL & " FROM facturas_calibre, albaran_calibre, calibres "
            If enlaza Then
                SQL = SQL & ObtenerWhereCab(True)
            Else
                SQL = SQL & " WHERE facturas_calibre.numfactu = '-1'"
            End If
            SQL = SQL & " and facturas_calibre.numalbar = albaran_calibre.numalbar"
            SQL = SQL & " and facturas_calibre.numlinealbar = albaran_calibre.numlinea"
            SQL = SQL & " and facturas_calibre.numline1albar = albaran_calibre.numline1"
            SQL = SQL & " and albaran_calibre.codvarie = calibres.codvarie"
            SQL = SQL & " and albaran_calibre.codcalib = calibres.codcalib"
            SQL = SQL & " ORDER BY facturas_calibre.numline1"
               
    End Select
    
    MontaSQLCarga = SQL
End Function

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim Aux As String
    
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabem quins camps son els que mos torna
        'Creem una cadena consulta i posem els datos
        cadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        cadB = Aux
        '   Com la clau principal es única, en posar el sql apuntant
        '   al valor retornat sobre la clau ppal es suficient
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        ' **********************************
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub


Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
Dim KiloCaja As Currency
Dim Cajas As Currency

    If CadenaSeleccion <> "" Then
        If ModoLineas = 1 Then ' insertar
            txtAux(14).Text = RecuperaValor(CadenaSeleccion, 1) ' linea de albaran
            txtAux2(2).Text = RecuperaValor(CadenaSeleccion, 3) ' nombre del calibre
            txtAux(1).Text = RecuperaValor(CadenaSeleccion, 7)  ' cantidad real
            
            Cajas = RecuperaValor(CadenaSeleccion, 4) ' cajas
            KiloCaja = DevuelveValor("select kiloscaj from albaran_variedad inner join forfaits on albaran_variedad.codforfait = forfaits.codforfait where numalbar=" & Text1(0).Text & " and numlinea = " & Text1(1).Text)
            txtAux(2).Text = Round2(KiloCaja * Cajas, 0) ' cantidad facturada
            
            txtAux(8).Text = RecuperaValor(CadenaSeleccion, 5) ' unidades
        Else ' modificar
            txtAux(14).Text = RecuperaValor(CadenaSeleccion, 1)
        End If
    End If
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
        Case 3  'Búscar
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
            
        Case btnPrimero To btnPrimero + 3 'Fleches Desplaçament
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub BotonBuscar()
Dim i As Integer
' ***** Si la clau primaria de la capçalera no es Text1(0), canviar-ho en <=== *****
'    If Modo <> 1 Then
'        LimpiarCampos
'        PonerModo 1
'        PonerFoco Text1(0) ' <===
'        Text1(0).BackColor = vbYellow ' <===
'        ' *** si n'hi han combos a la capçalera ***
'    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
'    End If
' ******************************************************************************
End Sub

Private Sub HacerBusqueda()

    cadB = ObtenerBusqueda2(Me, 1)
    
    If chkVistaPrevia(0) = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        ' *** foco al 1r camp visible de la capçalera que siga clau primaria ***
        PonerFoco Text1(0)
        ' **********************************************************************
    End If
End Sub

Private Sub MandaBusquedaPrevia(cadB As String)
    Dim cad As String
        
    'Cridem al form
    ' **************** arreglar-ho per a vore lo que es desije ****************
    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
    cad = ""
    cad = cad & ParaGrid(Text1(0), 20, "Código")
    cad = cad & ParaGrid(Text1(1), 20, "Confección")
    cad = cad & ParaGrid(Text1(2), 60, "Descripción")
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vtabla = NombreTabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        frmB.vDevuelve = "0|1|2|" '*** els camps que volen que torne ***
        frmB.vTitulo = "Forfaits" ' ***** repasa açò: títol de BuscaGrid *****
        frmB.vSelElem = 1

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
Dim cad As String
Dim Aux As String
Dim i As Integer
Dim J As Integer

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
    cadB = ""
    
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
    ' codEmpre i quins camps tenen la PK de la capçalera *******
'    text1(0).Text = SugerirCodigoSiguienteStr("forfaits", "codforfait")
'    FormateaCampo text1(0)
    
    Text1(0).Text = Albaran
    Text1(1).Text = SugerirCodigoSiguienteStr("albaran_variedad", "numlinea", "numalbar = " & Text1(0).Text)
    Text1(0).Locked = True
    Text1(1).Locked = True
    
    PonerFoco Text1(2) '*** 1r camp visible que siga PK ***
    
    ' *** si n'hi han camps de descripció a la capçalera ***
    'PosarDescripcions

End Sub

Private Sub BotonModificar()

    PonerModo 4
    
    Text1(0).Text = Albaran
    Text1(1).Text = Linea
    
    Text1(0).BackColor = &H80000013
    Text1(1).BackColor = &H80000013

    ' *** bloquejar els camps visibles de la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    BloquearTxt Text1(1), True
    BloquearTxt Text1(2), True
    
    'guardamos los kilos, cajas y forfaits
    KilosAnt = DBLet(Data1.Recordset!Pesoneto, "N")
    CajasAnt = DBLet(Data1.Recordset!NumCajas, "N")
    ForfaitAnt = DBLet(Data1.Recordset!codforfait, "T")
    CodPaletAnt = DBLet(Data1.Recordset!CodPalet, "N")
    TotPaletAnt = DBLet(Data1.Recordset!TotPalet, "N")
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
    cad = "¿Seguro que desea eliminar el Forfait?"
    cad = cad & vbCrLf & "Código: " & Data1.Recordset.Fields(0)
    cad = cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(1)
    
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
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
Dim codPobla As String, desPobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la capçalera
    
    ' *** si n'hi han llínies en datagrids ***
    'For i = 0 To DataGridAux.Count - 1
    For i = 0 To 0
            CargaGrid i, True
            If Not AdoAux(i).Recordset.EOF Then _
                PonerCamposForma2 Me, AdoAux(i), 2, "FrameAux" & i
    Next i

    
    ' ************* configurar els camps de les descripcions de la capçalera *************
    ' ********************************************************************************
    
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
                '++monica
                BloqueaRegistro "albaran", "numalbar= " & Text1(0).Text
                
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
Dim SQL As String
'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm2(Me, 1)
    If Not b Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
    If (Modo = 3) Then 'insertar
        'comprobar si existe ya el cod. del campo clave primaria
        SQL = ""
        SQL = DevuelveDesdeBDNew(cAgro, "facturas_calibre", "numalbar", "numalbar", Text1(0).Text, "N", , "numlinea", Text1(1).Text, "N")
        If SQL <> "" Then
            MsgBox "Ya existe el numero de linea para esta factura", vbExclamation
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

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    cad = "(codtipom=" & DBSet(Text1(2).Text, "T") & ")"
    cad = cad & " and numfactu = " & DBSet(Text1(3).Text, "N")
    cad = cad & " and fecfactu = " & DBSet(Text1(4).Text, "F")
    cad = cad & " and (numlinea = " & DBSet(Text1(5).Text, "N") & ")"
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    If SituarDataMULTI(Data1, cad, Indicador) Then
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

Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
Dim Variedad As String

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    
    ' ***************** configurar els LostFocus dels camps de la capçalera *****************
    Select Case Index
        Case 0 'codigo de forfait
            Text1(Index).Text = UCase(Text1(Index).Text)
        
    End Select
        ' ***************************************************************************
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 2: KEYBusqueda KeyAscii, 0 'VARIEDAD
                Case 3: KEYBusqueda KeyAscii, 1 'VARIEDAD COMERCIAL
                Case 4: KEYBusqueda KeyAscii, 2 'MARCA
                Case 5: KEYBusqueda KeyAscii, 3 'FORFAIT
                Case 13: KEYBusqueda KeyAscii, 4 'INCIDENCIA
            End Select
        End If
    Else
        KEYpress KeyAscii
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
End Sub



'************* LLINIES: ****************************
Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
'-- pon el bloqueo aqui
'    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
    
    Select Case Button.Index
        Case 1
            BotonAnyadirLinea Index
        Case 2
            BotonModificarLinea Index
        Case 3
            BotonEliminarLinea Index
        Case Else
    End Select
'    End If
End Sub

Private Sub BotonEliminarLinea(Index As Integer)
Dim SQL As String
Dim vWhere As String
Dim Eliminar As Boolean
Dim bol As Boolean
Dim MenError As String

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
        Case 0 'calibres
            SQL = "¿Seguro que desea eliminar la linea de calibre de la factura?"
            SQL = SQL & vbCrLf & "Línea: " & AdoAux(Index).Recordset!numline1
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                TerminaBloquear
                
                conn.BeginTrans
                
                SQL = "DELETE FROM facturas_calibre "
                SQL = SQL & vWhere & " AND numline1= " & AdoAux(Index).Recordset!numline1
                conn.Execute SQL
                
                bol = True
                If bol Then
                    bol = RecalcularDtos(Text1(2).Text, Text1(3).Text, Text1(4).Text, MenError)
                End If
                If bol Then
                    bol = ActualizarVariedades(Text1(2).Text, Text1(3).Text, Text1(4).Text, MenError)
                End If
            Else
                ' no hacemos nada en la transaccion
                bol = False
                conn.BeginTrans
            End If
    End Select

    
'    ModoLineas = 0
'    PosicionarData
    
Error2:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        MuestraError Err.Number, "Eliminando linea" & MenError, Err.Description
    
        bol = False
    End If
    If bol Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
    End If
    CargaGrid 0, True
    PonerModo 2
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
    BloquearTxt Text1(1), True
    

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case Index
        Case 0: vtabla = "facturas_calibre"
    End Select
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 0 ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            
            'NumF = SugerirCodigoSiguienteStr(vtabla, "numline1", vWhere)

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
                    txtAux(4).Text = Text1(2).Text 'codtipom
                    txtAux(15).Text = Text1(3).Text 'numfactu
                    txtAux(16).Text = Text1(4).Text 'fecha
                    txtAux(17).Text = Text1(5).Text 'linea
                    txtAux(0).Text = Text1(0).Text ' numero de albaran
                    txtAux(3).Text = Text1(1).Text ' numlinea del albaran
                    
                    BloquearTxt txtAux(14), False
                    BloquearTxt txtAux(1), False
                    BloquearTxt txtAux(2), False
                    BloquearTxt txtAux(6), False
                    BloquearTxt txtAux(9), False
                    
                    txtAux(5).Text = SugerirCodigoSiguienteStr("facturas_calibre", "numline1", "codtipom = '" & Text1(2).Text & "' and numfactu =  " & Text1(3).Text & " and fecfactu = " & DBSet(Text1(4).Text, "F") & " and numlinea = " & Text1(5).Text)
                    txtAux(1).Text = ""
                    txtAux(2).Text = ""
                    txtAux(6).Text = ""
                    txtAux(7).Text = ""
                    txtAux(8).Text = ""
                    txtAux(9).Text = ""
                    txtAux(10).Text = ""
                    txtAux2(2).Text = ""
                    txtAux(14).Text = ""
                    
                    txtAux(11).Text = Dto1
                    txtAux(12).Text = Dto2
                    
                    
                    BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux0"
                    PonerFoco txtAux(14)
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
        Case 0 ' calibres
'            Sql = "SELECT facturas_calibre.codtipom, facturas_calibre.numfactu, facturas_calibre.fecfactu, facturas_calibre.numlinea , facturas_calibre.numline1 ,facturas_calibre.numalbar, facturas_calibre.numlinealbar, facturas_calibre.numline1albar, "
'            Sql = Sql & " calibres.nomcalib, facturas_calibre.cantreal, facturas_calibre.cantfact, facturas_calibre.unidades, facturas_calibre.precibru, facturas_calibre.precinet, facturas_calibre.imporbru,"
'            Sql = Sql & "facturas_calibre.impornet, facturas_calibre.dtocom1, facturas_calibre.dtocom2"
'            Sql = Sql & " FROM facturas_calibre, albaran_calibre, calibres "
        
            txtAux(4).Text = DataGridAux(Index).Columns(0).Text
            txtAux(15).Text = DataGridAux(Index).Columns(1).Text
            txtAux(16).Text = DataGridAux(Index).Columns(2).Text
            txtAux(17).Text = DataGridAux(Index).Columns(3).Text
            txtAux(5).Text = DataGridAux(Index).Columns(4).Text
            txtAux(0).Text = DataGridAux(Index).Columns(5).Text
            txtAux(3).Text = DataGridAux(Index).Columns(6).Text
            txtAux(14).Text = DataGridAux(Index).Columns(7).Text
            txtAux2(2).Text = DataGridAux(Index).Columns(8).Text
            txtAux(1).Text = DataGridAux(Index).Columns(9).Text
            txtAux(2).Text = DataGridAux(Index).Columns(10).Text
            txtAux(8).Text = DataGridAux(Index).Columns(11).Text
            txtAux(6).Text = DataGridAux(Index).Columns(12).Text
            txtAux(7).Text = DataGridAux(Index).Columns(13).Text
            txtAux(9).Text = DataGridAux(Index).Columns(14).Text
            txtAux(10).Text = DataGridAux(Index).Columns(15).Text
            txtAux(11).Text = DataGridAux(Index).Columns(16).Text
            txtAux(12).Text = DataGridAux(Index).Columns(17).Text
            
            
            For i = 14 To 14
                BloquearTxt txtAux(i), True
            Next i
            
            For i = 1 To 2
                BloquearTxt txtAux(i), False
            Next i
            BloquearTxt txtAux(6), False
            BloquearTxt txtAux(9), False
            
            
            BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux0"
            
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
       
    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
    Select Case Index
        Case 0 'calibres
            txtAux(14).visible = b 'linea del albaran
            txtAux(14).Top = alto
            txtAux(1).visible = b 'cantidad real
            txtAux(1).Top = alto
            txtAux(2).visible = b 'cant.facturada
            txtAux(2).Top = alto
            txtAux(6).visible = b ' precio bruto
            txtAux(6).Top = alto
            txtAux(7).visible = b ' precio neto
            txtAux(7).Top = alto
            btnBuscar(0).visible = b
            btnBuscar(0).Top = alto
            txtAux(8).visible = b ' unidades
            txtAux(8).Top = alto
            txtAux2(2).visible = b ' nombre de calidad
            txtAux2(2).Top = alto
            txtAux(9).visible = b ' importe bruto
            txtAux(9).Top = alto
            txtAux(10).visible = b ' importe neto
            txtAux(10).Top = alto
            
    End Select
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
Dim Forfait As String
Dim SQL As String
Dim KilosUni As Currency

Dim TipoDto As Byte
Dim ImpDto As String
Dim Unidades As String
Dim Cantidad As String
Dim cad As String
Dim campo2 As String
Dim Cliente As String
Dim Rs As ADODB.Recordset
Dim Cajas As Currency
Dim KiloCaja As Currency


    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 14 ' linea de albaran
            If txtAux(14).Text <> "" Then
                SQL = "select albaran_calibre.*, calibres.nomcalib from albaran_calibre inner join calibres on albaran_calibre.codcalib = calibres.codcalib "
                SQL = SQL & " where numalbar = " & DBSet(Text1(0).Text, "N")
                SQL = SQL & " and numlinea = " & DBSet(Text1(1).Text, "N")
                SQL = SQL & " and numline1 = " & DBSet(txtAux(14).Text, "N")
                
                Set Rs = New ADODB.Recordset
                Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                If Not Rs.EOF Then
                    txtAux2(2).Text = DBLet(Rs!nomcalib, "T") ' nombre del calibre
                    txtAux(1).Text = DBLet(Rs!Pesoneto, "N") ' cantidad real
                    
                    Cajas = DBLet(Rs!NumCajas, "N") ' cajas
                    KiloCaja = DevuelveValor("select kiloscaj from albaran_variedad inner join forfaits on albaran_variedad.codforfait = forfaits.codforfait where numalbar=" & Text1(0).Text & " and numlinea = " & Text1(1).Text)
                    txtAux(2).Text = Round2(KiloCaja * Cajas, 0) ' cantidad facturada
                    
                    txtAux(8).Text = DBLet(Rs!Unidades, "N") ' unidades
                Else
                    MsgBox "No existe nro.línea de albarán. Reintroduzca.", vbExclamation
                    PonerFoco txtAux(14)
                End If
            End If
            
        Case 1 ' Cantidad real
            PonerFormatoEntero txtAux(Index)
        Case 2 ' Cantidad facturada
            PonerFormatoEntero txtAux(Index)
        
        Case 6 'precio bruto
            If txtAux(Index).Text <> "" Then
                If PonerFormatoDecimal(txtAux(Index), 7) Then
                    
                    Select Case TipoFacturarForfaits(txtAux(0).Text, txtAux(3).Text)
                        Case 0  'por unidades
                            txtAux(9).Text = Round2(CCur(ImporteSinFormato(ComprobarCero(txtAux(Index).Text))) * CCur(ImporteSinFormato(ComprobarCero(txtAux(8).Text))), 2)
                            PonerFormatoDecimal txtAux(9), 3
                        Case 1  'por kilos
                            txtAux(9).Text = Round2(CCur(ImporteSinFormato(ComprobarCero(txtAux(Index).Text))) * CCur(ImporteSinFormato(ComprobarCero(txtAux(1).Text))), 2)
                            PonerFormatoDecimal txtAux(9), 3
                        Case Else
                            
                    End Select
                    
                    cmdAceptar.SetFocus
                Else
                    Exit Sub
                End If
            End If
        
        Case 9 'importe bruto
            If txtAux(Index).Text <> "" Then
                If PonerFormatoDecimal(txtAux(Index), 3) Then
                
                    Select Case TipoFacturarForfaits(txtAux(0).Text, txtAux(3).Text)
                        Case 0
                            Unidades = ComprobarCero(txtAux(8).Text)
                            If CCur(Unidades) <> 0 Then
                                txtAux(6).Text = Round2(CCur(ImporteSinFormato(txtAux(Index).Text)) / CCur(Unidades), 4)
                            Else
                                txtAux(6).Text = 0
                            End If
                            PonerFormatoDecimal txtAux(8), 7
                        Case 1
                            Cantidad = ComprobarCero(txtAux(1).Text)
                            If CCur(Cantidad) <> 0 Then
                                txtAux(6).Text = Round2(CCur(ImporteSinFormato(txtAux(Index).Text)) / CCur(Cantidad), 4)
                            Else
                                txtAux(6).Text = 0
                            End If
                            PonerFormatoDecimal txtAux(6), 7
                        Case Else
                        
                    End Select
                        
                    cmdAceptar.SetFocus
               Else
                    Exit Sub
               End If
            End If
    End Select

If ((Index = 6 And txtAux(Index).Text <> "") Or (Index = 9 And txtAux(Index).Text <> "")) Then
        campo2 = "nrodecprec"
        Cliente = DevuelveValor("select codclien from albaran where numalbar = " & DBSet(Albaran, "N"))
        TipoDto = DevuelveDesdeBDNew(cAgro, "clientes", "tipodtos", "codclien", Cliente, "N", campo2)
        Select Case TipoFacturarForfaits(txtAux(0).Text, txtAux(3).Text)
            Case 0 ' unidades
                Unidades = ComprobarCero(txtAux(8).Text)
                '[Monica]24/11/2011: añadida condicion para evitar division por cero
                If CCur(Unidades) <> 0 Then
                    ImpDto = CalcularImporteDtoLinea(txtAux(8).Text, CStr(CCur(ImporteSinFormato(txtAux(9).Text)) / CCur(Unidades)), txtAux(4).Text, txtAux(15).Text, txtAux(16).Text, txtAux(17).Text, CStr(ImpDtoc), False)
                    txtAux(10).Text = CalcularImporteFClien(txtAux(8).Text, CStr(CCur(ImporteSinFormato(txtAux(9).Text)) / CCur(Unidades)), txtAux(11).Text, txtAux(12).Text, TipoDto, ImpDto, txtAux(9).Text)
                Else
                    ImpDto = CalcularImporteDtoLinea(txtAux(8).Text, CStr(0), txtAux(4).Text, txtAux(15).Text, txtAux(16).Text, txtAux(17).Text, CStr(ImpDtoc), False)
                    txtAux(10).Text = CalcularImporteFClien(txtAux(8).Text, CStr(0), txtAux(11).Text, txtAux(12).Text, TipoDto, ImpDto, txtAux(9).Text)
                End If
                PonerFormatoDecimal txtAux(10), 1
                
                'precio neto
                If CCur(ComprobarCero(txtAux(8).Text)) <> 0 Then
                    txtAux(7).Text = Round2(CCur(ImporteSinFormato(txtAux(10).Text)) / CCur(ImporteSinFormato(ComprobarCero(txtAux(8).Text))), CCur(campo2))
                Else
                    txtAux(7).Text = "0"
                End If
                PonerFormatoDecimal txtAux(7), 7
            
            Case 1 ' kilos
                Cantidad = ComprobarCero(txtAux(1).Text)
                If Cantidad <> "0" Then
                    ImpDto = CalcularImporteDtoLinea(txtAux(1).Text, CStr(CCur(ImporteSinFormato(txtAux(9).Text)) / CCur(Cantidad)), txtAux(4).Text, txtAux(15).Text, txtAux(16).Text, txtAux(17).Text, CStr(ImpDtoc), False)
                    txtAux(10).Text = CalcularImporteFClien(txtAux(1).Text, CStr(CCur(ImporteSinFormato(txtAux(9).Text)) / CCur(Cantidad)), txtAux(11).Text, txtAux(12).Text, TipoDto, ImpDto, txtAux(9).Text)
                Else
                    ImpDto = CalcularImporteDtoLinea(txtAux(1).Text, CStr(0), txtAux(4).Text, txtAux(15).Text, txtAux(16).Text, txtAux(17).Text, CStr(ImpDtoc), False)
                    txtAux(10).Text = CalcularImporteFClien(txtAux(1).Text, CStr(0), txtAux(11).Text, txtAux(12).Text, TipoDto, ImpDto, txtAux(9).Text)
                End If
                PonerFormatoDecimal txtAux(10), 1
                
                'precio neto
                If ComprobarCero(txtAux(1).Text) <> "0" Then
                    txtAux(7).Text = Round2(CCur(ImporteSinFormato(txtAux(10).Text)) / CCur(ImporteSinFormato(txtAux(1).Text)), CCur(campo2))
                Else
                    txtAux(7).Text = "0"
                End If
                PonerFormatoDecimal txtAux(7), 7
            
            Case Else
            
        End Select
        
    End If
    
    
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
   If Not txtAux(Index).MultiLine Then ConseguirFocoLin txtAux(Index)
End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
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
Dim SQL As String
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


Private Sub DataGridAux_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
Dim i As Byte

    If ModoLineas <> 1 Then
        Select Case Index
            Case 0 'cuentas bancarias
                If DataGridAux(Index).Columns.Count > 2 Then
                End If
                
            Case 1 'departamentos
                If DataGridAux(Index).Columns.Count > 2 Then
                End If
                
        End Select
        
    Else 'vamos a Insertar
        Select Case Index
            Case 0 'cuentas bancarias
            Case 1 'departamentos
                For i = 21 To 24
                Next i
            Case 2 'Tarjetas
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
        Case 0 'calibres
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;S|txtAux(14)|T|Linea|650|;S|btnBuscar(0)|B|||;" 'tipom,numfactu,fecfactu,numlinea,numline1,numalbar,numlinealbar,numline1albar de calibre
            tots = tots & "S|txtAux2(2)|T|Denominación|1700|;S|txtAux(1)|T|Cant.Real|1200|;S|txtAux(2)|T|Cant.Fact|1200|;S|txtAux(8)|T|Uds|800|;S|txtAux(6)|T|Pr.Bruto|1100|;S|txtAux(7)|T|Pr.Neto|1100|;"
            tots = tots & "S|txtAux(9)|T|Imp.Bruto|1200|;S|txtAux(10)|T|Imp.Neto|1200|;N||||0|;N||||0|;"
            
            arregla tots, DataGridAux(Index), Me, 350
        
            DataGridAux(0).Columns(6).Alignment = dbgRight
            DataGridAux(0).Columns(10).Alignment = dbgRight
            
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

Private Function InsertarLinea() As Boolean
'Inserta registre en les taules de Llínies
Dim nomFrame As String
Dim bol As Boolean
Dim MenError As String
Dim Pesoneto As String
Dim NumCajas As String

    On Error GoTo EInsertarLinea

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomFrame = "FrameAux0" 'calibres
    End Select
    
    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        '++monica
        BloqueaRegistro "facturas_variedad", "codtipom = " & DBSet(Text1(2).Text, "T") & " and numfactu = " & Text1(3).Text & " and fecfactu = " & DBSet(Text1(4).Text, "F") & " and numlinea = " & DBSet(Text1(5).Text, "N")
        
        'Aqui empieza transaccion
        conn.BeginTrans
        
        ' Si no existe la linea en factura_variedad metemos la insertamos
        Dim SQL As String
        Dim iva As Long
        
        SQL = "select count(*) from facturas_variedad where codtipom = " & DBSet(Text1(2).Text, "T")
        SQL = SQL & " and numfactu = " & DBSet(Text1(3).Text, "N")
        SQL = SQL & " and fecfactu = " & DBSet(Text1(4).Text, "F")
        SQL = SQL & " and numlinea = " & DBSet(Text1(5).Text, "N")
        If TotalRegistros(SQL) = 0 Then
            If TipoIva = "" Then
                iva = DevuelveValor("select codigiva from variedades inner join albaran_variedad On albaran_variedad.codvarie = variedades.codvarie where albaran_variedad.numalbar = " & DBSet(Text1(0).Text, "N") & " and albaran_variedad.numlinea = " & DBSet(Text1(1).Text, "N"))
            Else
                iva = TipoIva
            End If
        
            SQL = "insert into facturas_variedad (codtipom,numfactu,fecfactu,numlinea,numalbar,numlinealbar,cantreal,cantfact,precibru,precinet,dtocom1,dtocom2,imporbru,impornet,codigiva,unidades) values ("
            SQL = SQL & DBSet(Text1(2).Text, "T") & "," & DBSet(Text1(3).Text, "N") & "," & DBSet(Text1(4).Text, "F") & ","
            SQL = SQL & DBSet(Text1(5).Text, "N") & "," & DBSet(Text1(0).Text, "N") & "," & DBSet(Text1(1).Text, "N") & ","
            SQL = SQL & "0,0,0,0,0,0,0,0,"
            SQL = SQL & DBSet(iva, "N") & ",0)"
            
            conn.Execute SQL
        End If
        
        bol = InsertarDesdeForm2(Me, 2, nomFrame)
        
        If bol Then
            bol = RecalcularDtos(Text1(2).Text, Text1(3).Text, Text1(4).Text, MenError)
        End If
        
        If bol Then
            bol = ActualizarVariedades(Text1(2).Text, Text1(3).Text, Text1(4).Text, MenError)
        End If
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

Private Function ActualizarVariedades(CodTipoM As String, NumFactu As String, FecFactu As String, MensError As String) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim SQL1 As String
Dim Rdo As Integer
Dim Cliente As String


    On Error GoTo eActualizarVariedades

    ActualizarVariedades = False

    'Traemos el redondeo del precio
    SQL = ""
    SQL = DevuelveDesdeBDNew(cAgro, "facturas", "codclien", "codtipom", CodTipoM, "T", , "numfactu", NumFactu, "N", "fecfactu", FecFactu, "F")
    Cliente = ComprobarCero(SQL)
    SQL = ""
    SQL = DevuelveDesdeBDNew(cAgro, "clientes", "nrodecprec", "codclien", Cliente, "N")
    Rdo = ComprobarCero(SQL)

    SQL = "select numlinea from facturas_variedad where codtipom = " & DBSet(CodTipoM, "T")
    SQL = SQL & " and numfactu = " & DBSet(NumFactu, "N") & " and fecfactu = " & DBSet(FecFactu, "F")

    Set Rs2 = New ADODB.Recordset
    Rs2.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    While Not Rs2.EOF
        SQL1 = "select sum(if(cantreal is null,0,cantreal)), sum(if(cantfact is null,0,cantfact)), sum(if(imporbru is null,0,imporbru)), sum(if(impornet is null,0,impornet))"
        SQL1 = SQL1 & " from facturas_calibre "
        SQL1 = SQL1 & " where codtipom = " & DBSet(CodTipoM, "T")
        SQL1 = SQL1 & " and numfactu = " & DBSet(NumFactu, "N")
        SQL1 = SQL1 & " and fecfactu = " & DBSet(FecFactu, "F")
        SQL1 = SQL1 & " and numlinea = " & DBSet(Rs2!NumLinea, "N")
        conn.Execute SQL1
    
        Set Rs = New ADODB.Recordset
        Rs.Open SQL1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs.EOF Then
            SQL = "update facturas_variedad set cantreal = " & DBSet(Rs.Fields(0).Value, "N")
            SQL = SQL & ", cantfact = " & DBSet(Rs.Fields(1).Value, "N")
            SQL = SQL & ", imporbru = " & DBSet(Rs.Fields(2).Value, "N")
            SQL = SQL & ", impornet = " & DBSet(Rs.Fields(3).Value, "N")
            SQL = SQL & " where codtipom = " & DBSet(CodTipoM, "T")
            SQL = SQL & " and numfactu = " & DBSet(NumFactu, "N")
            SQL = SQL & " and fecfactu = " & DBSet(FecFactu, "F")
            SQL = SQL & " and numlinea = " & DBSet(Rs2!NumLinea, "N")
    
            conn.Execute SQL
                
            If TipoFacturarForfaits(txtAux(0).Text, txtAux(3).Text) = 1 Then
                SQL = "update facturas_variedad set precibru = round(imporbru / cantreal," & DBSet(Rdo, "N") & "), "
                SQL = SQL & " precinet = round(impornet / cantreal," & DBSet(Rdo, "N") & ") "
            Else
                SQL = "update facturas_variedad set precibru = round(imporbru / unidades," & DBSet(Rdo, "N") & "), "
                SQL = SQL & " precinet = round(impornet / unidades," & DBSet(Rdo, "N") & ") "
            End If
            SQL = SQL & " where codtipom = " & DBSet(CodTipoM, "T")
            SQL = SQL & " and numfactu = " & DBSet(NumFactu, "N")
            SQL = SQL & " and fecfactu = " & DBSet(FecFactu, "F")
            SQL = SQL & " and numlinea = " & DBSet(Rs2!NumLinea, "N")
    
            conn.Execute SQL
        
        End If
        Rs.Close
        Set Rs = Nothing
        
        Rs2.MoveNext
    Wend
    
    Set Rs2 = Nothing
    
eActualizarVariedades:
    If Err.Number = 0 Then ActualizarVariedades = True
End Function



Private Function ModificarLinea() As Boolean
'Modifica registre en les taules de Llínies
Dim nomFrame As String
Dim V As Integer
Dim bol As Boolean
Dim MenError As String
Dim Pesoneto As String
Dim NumCajas As String
    
    On Error GoTo eModificarLinea

    ' *** posa els noms del frames, tant si son de grid com si no ***
    nomFrame = "FrameAux0" 'calibres
    
    ModificarLinea = False
    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        
        conn.BeginTrans
        
        bol = ModificaDesdeFormulario2(Me, 2, nomFrame)
        
        If bol Then
            bol = RecalcularDtos(Text1(2).Text, Text1(3).Text, Text1(4).Text, MenError)
        End If
        
        If bol Then
            bol = ActualizarVariedades(Text1(2).Text, Text1(3).Text, Text1(4).Text, MenError)
        End If
    Else
        Exit Function
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
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " facturas_calibre.codtipom = '" & Trim(Text1(2).Text) & "' and facturas_calibre.numfactu=" & Text1(3).Text
    vWhere = vWhere & " and facturas_calibre.fecfactu = " & DBSet(Text1(4).Text, "F") & " and facturas_calibre.numlinea = " & Text1(5).Text
    
    
    ObtenerWhereCab = vWhere
End Function

'' *** neteja els camps dels tabs de grid que
''estan fora d'este, i els camps de descripció ***
Private Sub LimpiarCamposFrame(Index As Integer)
    On Error Resume Next
 
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Function ObtenerWhereCP(conW As Boolean) As String
Dim SQL As String
On Error Resume Next
    
    SQL = ""
    If conW Then SQL = " WHERE "
    SQL = SQL & NombreTabla & ".codtipom = " & DBSet(Text1(2).Text, "T")
    SQL = SQL & " and " & NombreTabla & ".numfactu=" & Val(Text1(3).Text)
    SQL = SQL & " and " & NombreTabla & ".fecfactu=" & DBSet(Text1(4).Text, "F")
    SQL = SQL & " and " & NombreTabla & ".numlinea=" & Val(Text1(5).Text)
    ObtenerWhereCP = SQL
End Function




Private Function Modificar() As Boolean
'Modifica registre en les taules de Llínies
Dim nomFrame As String
Dim V As Integer
Dim bol As Boolean
Dim MenError As String
Dim Forfait As String
Dim CodPalet As String
Dim TotPalet As String

    On Error GoTo EModificar

    TerminaBloquear
    
    conn.BeginTrans
    
    bol = ModificaDesdeFormulario2(Me, 1)
    
    If bol Then
        MenError = "Modificando variedades"
        bol = ActualizarVariedades(Text1(2).Text, Text1(3).Text, Text1(4).Text, MenError)
        
        If bol Then
            Forfait = ""
            Forfait = DevuelveDesdeBDNew(cAgro, "albaran_variedad", "codforfait", "numalbar", CStr(Albaran), "N", , "numlinea", CStr(Linea), "N")
            CodPalet = ""
            CodPalet = DevuelveDesdeBDNew(cAgro, "albaran_variedad", "codpalet", "numalbar", CStr(Albaran), "N", , "numlinea", CStr(Linea), "N")
            TotPalet = ""
            TotPalet = DevuelveDesdeBDNew(cAgro, "albaran_variedad", "codpalet", "numalbar", CStr(Albaran), "N", , "numlinea", CStr(Linea), "N")
            
            
        End If
    End If

EModificar:
    If Err.Number <> 0 Then
        MenError = "Modificando Registro Albarán Variedad." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        bol = False
    End If
    If bol Then
        conn.CommitTrans
        Modificar = True
    Else
        conn.RollbackTrans
        Modificar = False
    End If
End Function


Private Sub CargarDatosLineaAlbaran(Albaran As String, Linea As String, Sublinea As String)
Dim Forfait As String
Dim Pesoneto As String
Dim NumCajas As String
Dim KilosCaja As String
Dim i As Integer
Dim Rs As ADODB.Recordset
Dim SQL As String

'    If Albaran = "" Or Linea = "" Or Sublinea = "" Then
'        txtAux(6).Text = ""
'        txtAux(7).Text = ""
'        txtAux(15).Text = ""
'        For i = 0 To Text3.Count - 1
'            Text3(i).Text = ""
'        Next i
'    Else
'        Sql = "select albaran.fechaalb, albaran.matrirem, forfaits.kiloscaj, albaran_variedad.pesoneto, "
'        Sql = Sql & " albaran_variedad.numcajas, variedades.nomvarie, destinos.nomdesti, forfaits.nomconfe, albaran_variedad.unidades, albaran_variedad.codvarie "
'        Sql = Sql & " from albaran, albaran_variedad, variedades, destinos, forfaits "
'        Sql = Sql & " where albaran_variedad.numalbar = " & DBSet(Albaran, "N")
'        Sql = Sql & " and albaran_variedad.numlinea = " & DBSet(Linea, "N")
'        Sql = Sql & " and albaran.numalbar = albaran_variedad.numalbar "
'        Sql = Sql & " and albaran_variedad.codforfait = forfaitS.codforfait "
'        Sql = Sql & " and albaran_variedad.codvarie = variedades.codvarie "
'        Sql = Sql & " and albaran.codclien = destinos.codclien "
'        Sql = Sql & " and albaran.coddesti = destinos.coddesti "
'
'        Set Rs = New ADODB.Recordset
'
'        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'        If Not Rs.EOF Then
'            txtAux(6).Text = DBLet(Rs.Fields(3).Value, "N")
'            '[Monica]15/07/2011: añadido el round en la siguiente linea
'            txtAux(7).Text = Round2(DBLet(Rs.Fields(4).Value, "N") * DBLet(Rs.Fields(2).Value, "N"), 0)
'            txtAux(15).Text = DBLet(Rs.Fields(8).Value, "N")
'
'            '++monica:27/05/08: metemos el codigo iva de la variedad si el cliente es normal
'            If Me.Combo1(0).ListIndex = 0 Then
'                txtAux(14).Text = DevuelveDesdeBDNew(cAgro, "variedades", "codigiva", "codvarie", Rs.Fields(9).Value, "N")
'            End If
'            '++
'
'            Text3(0).Text = DBLet(Rs.Fields(0).Value, "F")
'            Text3(1).Text = DBLet(Rs.Fields(1).Value, "T")
'            Text3(2).Text = DBLet(Rs.Fields(6).Value, "T")
'            Text3(3).Text = DBLet(Rs.Fields(5).Value, "T")
'            Text3(4).Text = DBLet(Rs.Fields(7).Value, "T")
'        Else
'            txtAux3(6).Text = ""
'            txtAux3(7).Text = ""
'            txtAux3(15).Text = ""
'            For i = 0 To Text3.Count - 1
'                Text3(i).Text = ""
'            Next i
'
'            MsgBox "No existe la línea de albarán. Reintroduzca.", vbExclamation
'            txtAux3(5).Text = ""
''            PonerFoco txtAux3(5)
'
'        End If
'
'        Set Rs = Nothing
'    End If
'
End Sub


Private Function CalcularImporteDtoLinea(Cantidad As String, Precio As String, TipoM As String, Factura As String, FecFactu As String, Linea As String, ImpDto As String, Insertado As Boolean) As String
'Insertado: indica si ya hemos insertado el registro o no
'Calcula el Importe de una linea de Oferta, Pedido, Albaran, ...
'Importe=cantidad * precio - (descuentos)
Dim vCant As Currency
Dim vImp As Currency
Dim vDto As Currency
Dim vPre As Currency
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim SumaBruto As Currency

On Error Resume Next

    '[Monica]16/09/2011: antes estaba puesto que solo se hiciera para Castelduc, ahora lo he parametrizado
    '                    porque tambien lo van a hacer de esta manera en Alzira
    'If vParamAplic.Cooperativa = 5 Then
    If vParamAplic.TipoCalculoComision = 1 Then
        ' Para Castelduc y Alzira el importe de descuento se prorratea con respecto a los kilos
        
        SQL = "select sum(cantreal) from facturas_calibre where codtipom = " & DBSet(TipoM, "T")
        SQL = SQL & " and numfactu = " & DBSet(Factura, "N") & " and fecfactu = " & DBSet(FecFactu, "F")
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        SumaBruto = 0
        If Not Rs.EOF Then
            SumaBruto = DBLet(Rs.Fields(0).Value, "N")
        End If
        
        'Como son de tipo string comprobar que si vale "" lo ponemos a 0
        vCant = ComprobarCero(Cantidad)
        vDto = ComprobarCero(ImpDto)
        
        If Not Insertado Then '++monica 030608 añadido el round
            SumaBruto = SumaBruto + CCur(vCant)
        End If
        
        If SumaBruto <> 0 Then '++monica 030608 añadido el round
            vImp = Round2((CCur(Cantidad) * CCur(vDto)) / SumaBruto, 4)
        Else
            vImp = CCur(vDto)
        End If
        
        vImp = Round2(vImp, 6)
        
        CalcularImporteDtoLinea = CStr(vImp)
    Else
        ' como lo ha estado haciendo hasta ahora
        '(se prorratea el importe de descuento sobre el importe bruto de la linea)
        SQL = "select sum(imporbru) from facturas_calibre where codtipom = " & DBSet(TipoM, "T")
        SQL = SQL & " and numfactu = " & DBSet(Factura, "N") & " and fecfactu = " & DBSet(FecFactu, "F")
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        SumaBruto = 0
        If Not Rs.EOF Then
            SumaBruto = DBLet(Rs.Fields(0).Value, "N")
        End If
        
        'Como son de tipo string comprobar que si vale "" lo ponemos a 0
        vCant = ComprobarCero(Cantidad)
        vDto = ComprobarCero(ImpDto)
        
        If Not Insertado Then '++monica 030608 añadido el round
            SumaBruto = SumaBruto + Round2(CCur(vCant) * ComprobarCero(Precio), 2)
        End If
        
        If SumaBruto <> 0 Then '++monica 030608 añadido el round
            vImp = Round2((CCur(Cantidad) * ComprobarCero(Precio) * CCur(vDto)) / SumaBruto, 4)
        Else
            vImp = CCur(vDto)
        End If
        
        vImp = Round2(vImp, 6)
        
        CalcularImporteDtoLinea = CStr(vImp)
        
    End If

End Function


Private Function RecalcularDtos(TipoM As String, Factura As String, FecFactu As String, MenError As String) As Boolean
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim Sql2 As String
Dim vImpDto As Currency
Dim vDto1 As Currency
Dim vDto2 As Currency
Dim vImpDto1 As Currency
Dim vImpDto2 As Currency
Dim vImpNeto As Currency
Dim vPrecNeto As Currency
Dim TipoDto As String
Dim ImpDto As String
Dim Cliente As String
Dim Rdo As Long

Dim TotImpBru As Currency
Dim TotImpNet As Currency
Dim TotalDtos As Currency
Dim vAux As Currency
Dim Diferencia As Currency
Dim UltimaLinea As Currency
Dim UltimaLine1 As Currency
Dim TipoFactFor As Byte

Dim vHayReg As Byte

    On Error GoTo eRecalcularDtos

    SQL = "select * from facturas_calibre where codtipom = " & DBSet(TipoM, "T")
    SQL = SQL & " and numfactu = " & DBSet(Factura, "N") & " and fecfactu = " & DBSet(FecFactu, "F")
    SQL = SQL & " order by numlinea "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    SQL = ""
    SQL = DevuelveDesdeBDNew(cAgro, "facturas", "impdtoc", "codtipom", TipoM, "T", , "numfactu", Factura, "N", "fecfactu", FecFactu, "F")
    vImpDto = ComprobarCero(SQL)
    
    SQL = ""
    SQL = DevuelveDesdeBDNew(cAgro, "facturas", "dtocom1", "codtipom", TipoM, "T", , "numfactu", Factura, "N", "fecfactu", FecFactu, "F")
    vDto1 = ComprobarCero(SQL)
    
    SQL = ""
    SQL = DevuelveDesdeBDNew(cAgro, "facturas", "dtocom2", "codtipom", TipoM, "T", , "numfactu", Factura, "N", "fecfactu", FecFactu, "F")
    vDto2 = ComprobarCero(SQL)
    
    '++monica:030608:traemos el redondeo del precio
    SQL = ""
    SQL = DevuelveDesdeBDNew(cAgro, "facturas", "codclien", "codtipom", TipoM, "T", , "numfactu", Factura, "N", "fecfactu", FecFactu, "F")
    Cliente = ComprobarCero(SQL)
    SQL = ""
    SQL = DevuelveDesdeBDNew(cAgro, "clientes", "nrodecprec", "codclien", Cliente, "N")
    Rdo = ComprobarCero(SQL)
    
    vHayReg = 0
    
    While Not Rs.EOF
        vHayReg = 1
        
        TipoDto = DevuelveDesdeBDNew(cAgro, "clientes", "tipodtos", "codclien", Cliente, "N")
        If TipoFacturarForfaits(CStr(Rs!NumAlbar), CStr(Rs!numlinealbar)) = 1 Then 'kilos
            TipoFactFor = 1
            ImpDto = CalcularImporteDtoLinea(DBLet(Rs!cantreal, "N"), DBLet(Rs!precibru, "N"), TipoM, Factura, FecFactu, CStr(DBLet(Rs!NumLinea, "N")), CStr(vImpDto), True)
            vImpNeto = CalcularImporteFClien(DBLet(Rs!cantreal, "N"), DBLet(Rs!precibru, "N"), CStr(vDto1), CStr(vDto2), CByte(TipoDto), CStr(ImpDto), DBLet(Rs!imporbru, "N"))
            
            '[Monica]24/11/2011: si las unidades son 0 no hay division
            'precio neto
            vPrecNeto = 0
            If DBLet(Rs!cantreal, "N") <> 0 Then
                vPrecNeto = Round2(vImpNeto / DBLet(Rs!cantreal, "N"), Rdo)
            End If
            '++monica:040608 : solo si redondeo <> 4
            If Rdo = 2 Or Rdo = 3 Then
                vImpNeto = Round2(vPrecNeto * DBLet(Rs!cantreal, "N"), 2)
            End If
            
        Else 'unidades
            TipoFactFor = 0
            ImpDto = CalcularImporteDtoLinea(DBLet(Rs!Unidades, "N"), DBLet(Rs!precibru, "N"), TipoM, Factura, FecFactu, DBLet(Rs!NumLinea, "N"), CStr(vImpDto), True)
            vImpNeto = CalcularImporteFClien(DBLet(Rs!Unidades, "N"), DBLet(Rs!precibru, "N"), CStr(vDto1), CStr(vDto2), CByte(TipoDto), CStr(ImpDto), DBLet(Rs!imporbru, "N"))
            
            '[Monica]24/11/2011: si las unidades son 0 no hay division
            'precio neto
            vPrecNeto = 0
            If DBLet(Rs!Unidades, "N") <> 0 Then
                vPrecNeto = Round2(vImpNeto / DBLet(Rs!Unidades, "N"), Rdo)
            End If
            
            '++monica:040608
            If Rdo = 2 Or Rdo = 3 Then
                vImpNeto = Round2(vPrecNeto * DBLet(Rs!Unidades, "N"), 2)
            End If
        End If
        
        Sql2 = "update facturas_calibre set impornet = " & DBSet(vImpNeto, "N")
        Sql2 = Sql2 & ",precinet = " & DBSet(vPrecNeto, "N")
        Sql2 = Sql2 & ",dtocom1 = " & DBSet(vDto1, "N")
        Sql2 = Sql2 & ",dtocom2 = " & DBSet(vDto2, "N")
        Sql2 = Sql2 & " where codtipom = " & DBSet(TipoM, "T")
        Sql2 = Sql2 & " and numfactu = " & DBSet(Factura, "N")
        Sql2 = Sql2 & " and fecfactu = " & DBSet(FecFactu, "F")
        Sql2 = Sql2 & " and numlinea = " & DBSet(Rs!NumLinea, "N")
        Sql2 = Sql2 & " and numline1 = " & DBSet(Rs!numline1, "N")
    
        conn.Execute Sql2
    
        UltimaLinea = DBLet(Rs!NumLinea, "N")
        UltimaLine1 = DBLet(Rs!numline1, "N")
    
        Rs.MoveNext
    Wend
    
    Rs.Close
    
    '[Monica]16/09/2011: si no coincide la suma de dtos con el total descuento redondeamos en la ultima linea
    If vHayReg = 1 Then
        Sql2 = "select sum(imporbru) bruto, sum(impornet) neto from facturas_calibre "
        Sql2 = Sql2 & " where codtipom = " & DBSet(TipoM, "T")
        Sql2 = Sql2 & " and numfactu = " & DBSet(Factura, "N")
        Sql2 = Sql2 & " and fecfactu = " & DBSet(FecFactu, "F")
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql2, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
        Diferencia = 0
        If Not Rs.EOF Then
            TotImpBru = DBLet(Rs.Fields(0).Value, "N")
            TotImpNet = DBLet(Rs.Fields(1).Value, "N")
            If CByte(TipoDto) = 0 Then
                vImpDto1 = (CCur(vDto1) * TotImpBru) / 100
                vImpDto2 = (CCur(vDto2) * TotImpBru) / 100
            ElseIf CByte(TipoDto) = 1 Then 'Sobre Resto
                vImpDto1 = (CCur(vDto1) * TotImpBru) / 100
                vAux = TotImpBru - vImpDto1
                vImpDto2 = (CCur(vDto2) * vAux) / 100
            End If
            TotalDtos = vImpDto1 + vImpDto2
            
            TotalDtos = TotalDtos + vImpDto
        
            If TotImpBru - TotalDtos <> TotImpNet Then
                Diferencia = TotImpBru - TotalDtos - TotImpNet
                
                Sql2 = "update facturas_calibre set impornet = impornet + " & DBSet(Diferencia, "N")
                Sql2 = Sql2 & " where codtipom = " & DBSet(TipoM, "T")
                Sql2 = Sql2 & " and numfactu = " & DBSet(Factura, "N")
                Sql2 = Sql2 & " and fecfactu = " & DBSet(FecFactu, "F")
                Sql2 = Sql2 & " and numlinea = " & DBSet(UltimaLinea, "N")
                Sql2 = Sql2 & " and numline1 = " & DBSet(UltimaLine1, "N")
            
                conn.Execute Sql2
        
                If TipoFactFor = 1 Then 'kilos
                    Sql2 = "update facturas_calibre set precinet = round(impornet / cantreal, " & DBSet(Rdo, "N") & ") "
                    Sql2 = Sql2 & " where codtipom = " & DBSet(TipoM, "T")
                    Sql2 = Sql2 & " and numfactu = " & DBSet(Factura, "N")
                    Sql2 = Sql2 & " and fecfactu = " & DBSet(FecFactu, "F")
                    Sql2 = Sql2 & " and numlinea = " & DBSet(UltimaLinea, "N")
                    Sql2 = Sql2 & " and numline1 = " & DBSet(UltimaLine1, "N")
                
                    conn.Execute Sql2
                Else 'unidades
                    'precio neto
                    Sql2 = "update facturas_calibre set precinet = round(impornet / unidades, " & DBSet(Rdo, "N") & ") "
                    Sql2 = Sql2 & " where codtipom = " & DBSet(TipoM, "T")
                    Sql2 = Sql2 & " and numfactu = " & DBSet(Factura, "N")
                    Sql2 = Sql2 & " and fecfactu = " & DBSet(FecFactu, "F")
                    Sql2 = Sql2 & " and numlinea = " & DBSet(UltimaLinea, "N")
                    Sql2 = Sql2 & " and numline1 = " & DBSet(UltimaLine1, "N")
                
                    conn.Execute Sql2
                End If
            End If
        
        End If
    End If
    
    Set Rs = Nothing
    
    RecalcularDtos = True
    Exit Function

eRecalcularDtos:
    If Err.Number <> 0 Then
        MenError = MenError & vbCrLf & Err.Description
        RecalcularDtos = False
    End If
End Function

