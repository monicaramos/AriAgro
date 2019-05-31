VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmVtasAlbSocios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Albaranes Venta de Socios"
   ClientHeight    =   10080
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   18000
   Icon            =   "frmVtasAlbSocios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10080
   ScaleWidth      =   18000
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
      Left            =   15435
      TabIndex        =   202
      Top             =   270
      Width           =   1605
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   4800
      TabIndex        =   200
      Top             =   45
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   201
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
      TabIndex        =   198
      Top             =   45
      Width           =   930
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   210
         TabIndex        =   199
         Top             =   180
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Generar Factura"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   135
      TabIndex        =   196
      Top             =   45
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   197
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   4740
      Left            =   135
      TabIndex        =   47
      Top             =   4680
      Width           =   17730
      _ExtentX        =   31274
      _ExtentY        =   8361
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Variedades"
      TabPicture(0)   =   "frmVtasAlbSocios.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "imgFact(2)"
      Tab(0).Control(1)=   "Label9(2)"
      Tab(0).Control(2)=   "Label8(2)"
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(4)=   "DataGrid1"
      Tab(0).Control(5)=   "DataGrid2"
      Tab(0).Control(6)=   "ToolAux(0)"
      Tab(0).Control(7)=   "txtAux(7)"
      Tab(0).Control(8)=   "txtAux3(15)"
      Tab(0).Control(9)=   "txtAux3(14)"
      Tab(0).Control(10)=   "txtAux3(13)"
      Tab(0).Control(11)=   "txtAux3(12)"
      Tab(0).Control(12)=   "txtAux3(11)"
      Tab(0).Control(13)=   "txtAux3(6)"
      Tab(0).Control(14)=   "txtAux3(10)"
      Tab(0).Control(15)=   "txtAux3(9)"
      Tab(0).Control(16)=   "txtAux3(8)"
      Tab(0).Control(17)=   "txtAux3(7)"
      Tab(0).Control(18)=   "txtAux3(5)"
      Tab(0).Control(19)=   "txtAux3(4)"
      Tab(0).Control(20)=   "txtAux3(3)"
      Tab(0).Control(21)=   "txtAux(6)"
      Tab(0).Control(22)=   "txtAux(5)"
      Tab(0).Control(23)=   "txtAux(0)"
      Tab(0).Control(24)=   "txtAux(1)"
      Tab(0).Control(25)=   "txtAux(2)"
      Tab(0).Control(26)=   "txtAux(3)"
      Tab(0).Control(27)=   "txtAux(4)"
      Tab(0).Control(28)=   "txtAux3(0)"
      Tab(0).Control(29)=   "txtAux3(1)"
      Tab(0).Control(30)=   "txtAux3(2)"
      Tab(0).Control(31)=   "txtAux(22)"
      Tab(0).Control(32)=   "txtAux3(16)"
      Tab(0).ControlCount=   33
      TabCaption(1)   =   "Envases Paletización"
      TabPicture(1)   =   "frmVtasAlbSocios.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameAux0"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Costes"
      TabPicture(2)   =   "frmVtasAlbSocios.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameAux1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Palets"
      TabPicture(3)   =   "frmVtasAlbSocios.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FrameAux2"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Resultados"
      TabPicture(4)   =   "frmVtasAlbSocios.frx":007C
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "ListView1"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame4"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
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
         Left            =   -64245
         MaxLength       =   4
         TabIndex        =   181
         Tag             =   "Unidades|N|S|||albaran_variedad|unidades|##,##0|N|"
         Text            =   "Ud"
         Top             =   1935
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
         Index           =   22
         Left            =   -62940
         MaxLength       =   4
         TabIndex        =   180
         Tag             =   "Unidades|N|S|0|999|albaran_calibre|unidades|##,##0||"
         Text            =   "ud"
         Top             =   3915
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame Frame4 
         Height          =   3390
         Left            =   8775
         TabIndex        =   152
         Top             =   765
         Width           =   8835
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
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
            Index           =   39
            Left            =   1935
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   193
            Text            =   "Costes Comis"
            Top             =   1350
            Width           =   1700
         End
         Begin VB.TextBox Text2 
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
            Index           =   36
            Left            =   1935
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   163
            Text            =   "Costes Totales"
            Top             =   2115
            Width           =   1700
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
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
            Index           =   35
            Left            =   1935
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   162
            Text            =   "Costes Totales"
            Top             =   1710
            Width           =   1700
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
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
            Index           =   34
            Left            =   1935
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   161
            Text            =   "Costes Portes"
            Top             =   990
            Width           =   1700
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
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
            Index           =   33
            Left            =   1935
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   160
            Text            =   "Costes Envases"
            Top             =   630
            Width           =   1700
         End
         Begin VB.TextBox Text2 
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
            Index           =   32
            Left            =   1935
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   159
            Text            =   "Gastos/kg"
            Top             =   2835
            Width           =   1700
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
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
            Index           =   31
            Left            =   1935
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   158
            Text            =   "Gastos/caja"
            Top             =   2475
            Width           =   1700
         End
         Begin VB.TextBox Text2 
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
            Index           =   30
            Left            =   6300
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   157
            Text            =   "Importe Vta"
            Top             =   855
            Width           =   1700
         End
         Begin VB.TextBox Text2 
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
            Index           =   29
            Left            =   6300
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   156
            Text            =   "venta/kg"
            Top             =   1575
            Width           =   1700
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
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
            Index           =   28
            Left            =   6300
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   155
            Text            =   "venta/caja"
            Top             =   1215
            Width           =   1700
         End
         Begin VB.TextBox Text2 
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
            Index           =   27
            Left            =   6300
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   154
            Text            =   "Valorfruta"
            Top             =   2115
            Width           =   1700
         End
         Begin VB.TextBox Text2 
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
            Index           =   26
            Left            =   6300
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   153
            Text            =   "Neto/kg"
            Top             =   2475
            Width           =   1700
         End
         Begin VB.Label Label1 
            Caption         =   "Comision"
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
            Index           =   48
            Left            =   810
            TabIndex        =   194
            Top             =   1395
            Width           =   1080
         End
         Begin VB.Label Label9 
            Caption         =   "Facturado"
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
            Height          =   195
            Index           =   1
            Left            =   4320
            TabIndex        =   183
            Top             =   2925
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.Label Label8 
            Caption         =   "Cobrado"
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
            Height          =   195
            Index           =   1
            Left            =   6300
            TabIndex        =   182
            Top             =   2925
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Image imgFact 
            Height          =   330
            Index           =   1
            Left            =   7650
            ToolTipText     =   "Facturas asociadas"
            Top             =   2925
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.Label Label1 
            Caption         =   "Peso Neto "
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
            Index           =   43
            Left            =   4320
            TabIndex        =   179
            Top             =   270
            Width           =   1485
         End
         Begin VB.Label Label1 
            Caption         =   "Cajas"
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
            Index           =   42
            Left            =   495
            TabIndex        =   178
            Top             =   270
            Width           =   720
         End
         Begin VB.Line Line2 
            X1              =   450
            X2              =   3690
            Y1              =   2070
            Y2              =   2070
         End
         Begin VB.Label Label1 
            Caption         =   "Gastos/Kg."
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
            Left            =   495
            TabIndex        =   177
            Top             =   2895
            Width           =   1065
         End
         Begin VB.Label Label1 
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
            Height          =   180
            Index           =   40
            Left            =   495
            TabIndex        =   176
            Top             =   1710
            Width           =   165
         End
         Begin VB.Label Label1 
            Caption         =   "GASTOS"
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
            Left            =   495
            TabIndex        =   175
            Top             =   2175
            Width           =   1530
         End
         Begin VB.Label Label1 
            Caption         =   "Costes"
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
            Index           =   38
            Left            =   810
            TabIndex        =   174
            Top             =   1755
            Width           =   720
         End
         Begin VB.Label Label1 
            Caption         =   "Portes"
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
            Left            =   810
            TabIndex        =   173
            Top             =   1035
            Width           =   720
         End
         Begin VB.Label Label1 
            Caption         =   "Envases"
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
            Index           =   36
            Left            =   810
            TabIndex        =   172
            Top             =   675
            Width           =   990
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cajas"
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
            Left            =   1935
            TabIndex        =   171
            Top             =   225
            Width           =   1695
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Peso Neto"
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
            Left            =   6300
            TabIndex        =   170
            Top             =   225
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Gastos/Caja"
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
            Left            =   495
            TabIndex        =   169
            Top             =   2535
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "IMPORTE VENTA"
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
            Index           =   34
            Left            =   4320
            TabIndex        =   168
            Top             =   915
            Width           =   1965
         End
         Begin VB.Label Label1 
            Caption         =   "Importe Venta/Kg."
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
            Left            =   4320
            TabIndex        =   167
            Top             =   1620
            Width           =   1920
         End
         Begin VB.Label Label1 
            Caption         =   "Importe Venta/Caja"
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
            Index           =   32
            Left            =   4320
            TabIndex        =   166
            Top             =   1260
            Width           =   2010
         End
         Begin VB.Label Label1 
            Caption         =   "VALOR FRUTA"
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
            Index           =   31
            Left            =   4320
            TabIndex        =   165
            Top             =   2175
            Width           =   1470
         End
         Begin VB.Label Label1 
            Caption         =   "Neto/Kg."
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
            Left            =   4320
            TabIndex        =   164
            Top             =   2535
            Width           =   1380
         End
      End
      Begin VB.Frame FrameAux1 
         BorderStyle     =   0  'None
         Height          =   4110
         Left            =   -75000
         TabIndex        =   90
         Top             =   360
         Width           =   17565
         Begin VB.Frame Frame3 
            Height          =   3390
            Left            =   8775
            TabIndex        =   121
            Top             =   405
            Width           =   8775
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
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
               Index           =   38
               Left            =   1935
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   191
               Text            =   "Costes Comis"
               Top             =   1350
               Width           =   1695
            End
            Begin VB.TextBox Text2 
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
               Index           =   25
               Left            =   6300
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   148
               Text            =   "Neto/kg"
               Top             =   2475
               Width           =   1695
            End
            Begin VB.TextBox Text2 
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
               Index           =   24
               Left            =   6300
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   147
               Text            =   "Valorfruta"
               Top             =   2115
               Width           =   1695
            End
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
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
               Index           =   23
               Left            =   6300
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   146
               Text            =   "venta/caja"
               Top             =   1215
               Width           =   1695
            End
            Begin VB.TextBox Text2 
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
               Index           =   22
               Left            =   6300
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   145
               Text            =   "venta/kg"
               Top             =   1575
               Width           =   1695
            End
            Begin VB.TextBox Text2 
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
               Index           =   21
               Left            =   6300
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   144
               Text            =   "Importe Vta"
               Top             =   855
               Width           =   1695
            End
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
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
               Index           =   20
               Left            =   1935
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   143
               Text            =   "Gastos/caja"
               Top             =   2475
               Width           =   1695
            End
            Begin VB.TextBox Text2 
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
               Left            =   1935
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   142
               Text            =   "Gastos/kg"
               Top             =   2835
               Width           =   1695
            End
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
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
               Index           =   14
               Left            =   1935
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   125
               Text            =   "Costes Envases"
               Top             =   630
               Width           =   1695
            End
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
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
               Index           =   15
               Left            =   1935
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   124
               Text            =   "Costes Portes"
               Top             =   990
               Width           =   1695
            End
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
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
               Left            =   1935
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   123
               Text            =   "Costes Totales"
               Top             =   1710
               Width           =   1695
            End
            Begin VB.TextBox Text2 
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
               Index           =   17
               Left            =   1935
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   122
               Text            =   "Costes Totales"
               Top             =   2115
               Width           =   1695
            End
            Begin VB.Label Label1 
               Caption         =   "Comision"
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
               Index           =   47
               Left            =   810
               TabIndex        =   192
               Top             =   1395
               Width           =   945
            End
            Begin VB.Image imgFact 
               Height          =   330
               Index           =   0
               Left            =   7650
               ToolTipText     =   "Facturas asociadas"
               Top             =   2925
               Visible         =   0   'False
               Width           =   330
            End
            Begin VB.Label Label8 
               Caption         =   "Cobrado"
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
               Height          =   195
               Index           =   0
               Left            =   6300
               TabIndex        =   150
               Top             =   2925
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Label Label9 
               Caption         =   "Facturado"
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
               Height          =   195
               Index           =   0
               Left            =   4320
               TabIndex        =   149
               Top             =   2925
               Visible         =   0   'False
               Width           =   1275
            End
            Begin VB.Label Label1 
               Caption         =   "Neto/Kg."
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
               Index           =   30
               Left            =   4320
               TabIndex        =   141
               Top             =   2535
               Width           =   1380
            End
            Begin VB.Label Label1 
               Caption         =   "VALOR FRUTA"
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
               Index           =   27
               Left            =   4320
               TabIndex        =   140
               Top             =   2175
               Width           =   1470
            End
            Begin VB.Label Label1 
               Caption         =   "Importe Venta/Caja"
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
               Left            =   4320
               TabIndex        =   139
               Top             =   1260
               Width           =   2055
            End
            Begin VB.Label Label1 
               Caption         =   "Importe Venta/Kg."
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
               Left            =   4320
               TabIndex        =   138
               Top             =   1620
               Width           =   2010
            End
            Begin VB.Label Label1 
               Caption         =   "IMPORTE VENTA"
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
               Index           =   24
               Left            =   4320
               TabIndex        =   137
               Top             =   915
               Width           =   1605
            End
            Begin VB.Label Label1 
               Caption         =   "Gastos/Caja"
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
               Index           =   23
               Left            =   495
               TabIndex        =   136
               Top             =   2535
               Width           =   1290
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Peso Neto"
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
               Left            =   6300
               TabIndex        =   135
               Top             =   225
               Width           =   1695
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000018&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Cajas"
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
               Left            =   1935
               TabIndex        =   134
               Top             =   225
               Width           =   1695
            End
            Begin VB.Label Label1 
               Caption         =   "Envases"
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
               Left            =   810
               TabIndex        =   133
               Top             =   675
               Width           =   945
            End
            Begin VB.Label Label1 
               Caption         =   "Portes"
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
               Left            =   810
               TabIndex        =   132
               Top             =   1035
               Width           =   945
            End
            Begin VB.Label Label1 
               Caption         =   "Costes"
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
               Left            =   810
               TabIndex        =   131
               Top             =   1755
               Width           =   720
            End
            Begin VB.Label Label1 
               Caption         =   "GASTOS"
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
               Index           =   11
               Left            =   495
               TabIndex        =   130
               Top             =   2175
               Width           =   1530
            End
            Begin VB.Label Label1 
               Caption         =   "+"
               Height          =   180
               Index           =   19
               Left            =   495
               TabIndex        =   129
               Top             =   1800
               Width           =   165
            End
            Begin VB.Label Label1 
               Caption         =   "Gastos/Kg."
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
               Left            =   495
               TabIndex        =   128
               Top             =   2895
               Width           =   1065
            End
            Begin VB.Line Line1 
               X1              =   405
               X2              =   3870
               Y1              =   2070
               Y2              =   2070
            End
            Begin VB.Label Label1 
               Caption         =   "Cajas"
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
               Left            =   495
               TabIndex        =   127
               Top             =   270
               Width           =   720
            End
            Begin VB.Label Label1 
               Caption         =   "Peso Neto "
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
               Index           =   22
               Left            =   4320
               TabIndex        =   126
               Top             =   270
               Width           =   1485
            End
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
            Index           =   20
            Left            =   6255
            MaxLength       =   16
            TabIndex        =   103
            Tag             =   "Unidades|N|N|||albaran_costes|unidades|###,##0|N|"
            Text            =   "unidades"
            Top             =   3285
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
            Index           =   19
            Left            =   7065
            MaxLength       =   16
            TabIndex        =   102
            Tag             =   "Importe|N|N|||albaran_costes|importes|#,##0.0000|N|"
            Text            =   "importes"
            Top             =   3285
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
            Index           =   18
            Left            =   7920
            MaxLength       =   16
            TabIndex        =   101
            Tag             =   "Imp.Coste|N|N|||albaran_costes|impcoste|##,##,##0.0000|N|"
            Text            =   "impcoste"
            Top             =   3285
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox Text2 
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
            Left            =   4320
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   100
            Text            =   "Nomcoste"
            Top             =   3285
            Width           =   1740
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
            Index           =   17
            Left            =   3465
            MaxLength       =   2
            TabIndex        =   99
            Tag             =   "Cod.Coste|N|N|||albaran_costes|codcoste|00|N|"
            Text            =   "codcoste"
            Top             =   3285
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
            Index           =   16
            Left            =   1140
            MaxLength       =   12
            TabIndex        =   98
            Tag             =   "Num.Linea|N|N|||albaran_costes|numlinea|00|N|"
            Text            =   "numlinea"
            Top             =   3285
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
            Index           =   15
            Left            =   180
            MaxLength       =   12
            TabIndex        =   97
            Tag             =   "Num.Albaran|N|N|||albaran_costes|numalbar||S|"
            Text            =   "numalbar"
            Top             =   3285
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.ComboBox cmbAux 
            Appearance      =   0  'Flat
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
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   96
            Tag             =   "Tipo Movimiento|N|N|||albaran_costes|tipogasto|0||"
            Top             =   3285
            Width           =   1260
         End
         Begin MSDataGridLib.DataGrid DataGrid4 
            Bindings        =   "frmVtasAlbSocios.frx":0098
            Height          =   3285
            Left            =   45
            TabIndex        =   91
            Top             =   495
            Width           =   8345
            _ExtentX        =   14711
            _ExtentY        =   5794
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
         Begin MSAdodcLib.Adodc Adoaux 
            Height          =   330
            Index           =   1
            Left            =   135
            Top             =   675
            Visible         =   0   'False
            Width           =   1425
            _ExtentX        =   2514
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
         Begin VB.Label Label7 
            Caption         =   "Confección:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   9675
            TabIndex        =   120
            Top             =   135
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Marca:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6075
            TabIndex        =   119
            Top             =   135
            Width           =   465
         End
         Begin VB.Label Label5 
            Caption         =   "Comercial:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3105
            TabIndex        =   118
            Top             =   135
            Width           =   690
         End
         Begin VB.Label Label4 
            Caption         =   "Variedad:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   150
            Left            =   45
            TabIndex        =   117
            Top             =   135
            Width           =   645
         End
         Begin VB.Label Label2 
            Caption         =   "Forfait890123456789012345678901234567980"
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
            Height          =   285
            Index           =   3
            Left            =   10440
            TabIndex        =   114
            Top             =   90
            Width           =   4650
         End
         Begin VB.Label Label2 
            Caption         =   "Marca67890123456789012345"
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
            Height          =   285
            Index           =   2
            Left            =   6570
            TabIndex        =   113
            Top             =   90
            Width           =   2985
         End
         Begin VB.Label Label2 
            Caption         =   "Variedad Comercial90"
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
            Height          =   285
            Index           =   1
            Left            =   3825
            TabIndex        =   112
            Top             =   90
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "Variedad901234567890"
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
            Height          =   285
            Index           =   0
            Left            =   720
            TabIndex        =   111
            Top             =   90
            Width           =   2400
         End
      End
      Begin VB.Frame FrameAux2 
         BorderStyle     =   0  'None
         Height          =   4110
         Left            =   -74910
         TabIndex        =   85
         Top             =   405
         Width           =   13650
         Begin VB.TextBox Text2 
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
            Left            =   8775
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   110
            Tag             =   "Hora Inicio|FH|N|||palets|horafin|hh:mm:ss||"
            Text            =   "Hora Fin"
            Top             =   3240
            Width           =   1110
         End
         Begin VB.TextBox Text2 
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
            Left            =   7560
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   109
            Text            =   "Fecha Fin"
            Top             =   3240
            Width           =   1110
         End
         Begin VB.TextBox Text2 
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
            Left            =   6345
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   108
            Tag             =   "Hora Inicio|FH|N|||palets|horaini|hh:mm:ss||"
            Text            =   "Hora.Ini."
            Top             =   3240
            Width           =   1110
         End
         Begin VB.TextBox Text2 
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
            Left            =   5130
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   107
            Text            =   "Fec.Inicio"
            Top             =   3240
            Width           =   1110
         End
         Begin VB.TextBox Text2 
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
            Left            =   3960
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   106
            Text            =   "Tip.Mercancia"
            Top             =   3240
            Width           =   1110
         End
         Begin VB.TextBox Text2 
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
            Left            =   2790
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   105
            Text            =   "Lin.Confec"
            Top             =   3240
            Width           =   1110
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
            Index           =   14
            Left            =   2025
            MaxLength       =   6
            TabIndex        =   88
            Tag             =   "Num.Palets|N|N|||albaran_palets|numpalet|###,##0|N|"
            Text            =   "N.Palets"
            Top             =   3240
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
            Index           =   13
            Left            =   1215
            MaxLength       =   12
            TabIndex        =   87
            Tag             =   "Num.Linea|N|N|||albaran_palets|numlinea|00|S|"
            Text            =   "numlinea"
            Top             =   3240
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
            Index           =   12
            Left            =   315
            MaxLength       =   12
            TabIndex        =   86
            Tag             =   "Num.Albar|N|N|||albaran_palets|numalbar||S|"
            Text            =   "numalbar"
            Top             =   3240
            Visible         =   0   'False
            Width           =   855
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   2
            Left            =   45
            TabIndex        =   89
            Top             =   0
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
         Begin MSDataGridLib.DataGrid DataGrid5 
            Bindings        =   "frmVtasAlbSocios.frx":00AD
            Height          =   3285
            Left            =   0
            TabIndex        =   95
            Top             =   450
            Width           =   10955
            _ExtentX        =   19315
            _ExtentY        =   5794
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
         Begin MSAdodcLib.Adodc Adoaux 
            Height          =   330
            Index           =   2
            Left            =   1755
            Top             =   45
            Visible         =   0   'False
            Width           =   1425
            _ExtentX        =   2514
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
      End
      Begin VB.Frame FrameAux0 
         BorderStyle     =   0  'None
         Height          =   4110
         Left            =   -74955
         TabIndex        =   76
         Top             =   360
         Width           =   13650
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
            Index           =   23
            Left            =   10305
            MaxLength       =   10
            TabIndex        =   185
            Tag             =   "Cliente|N|N|||albaran_envase|codclien|000000||"
            Text            =   "cliente"
            Top             =   3240
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
            Index           =   21
            Left            =   9585
            MaxLength       =   10
            TabIndex        =   82
            Tag             =   "Fecha Movimiento|F|N|||albaran_envase|fechamov|dd/mm/yyyy||"
            Text            =   "fechamov"
            Top             =   3240
            Visible         =   0   'False
            Width           =   735
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
            Height          =   300
            Index           =   0
            Left            =   3195
            MaskColor       =   &H00000000&
            TabIndex        =   104
            ToolTipText     =   "Buscar Envase"
            Top             =   3240
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.ComboBox cmbAux 
            Appearance      =   0  'Flat
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
            Left            =   7515
            Style           =   2  'Dropdown List
            TabIndex        =   79
            Tag             =   "Tipo Movimiento|N|N|||albaran_envase|tipomovi|0||"
            Top             =   3240
            Width           =   1260
         End
         Begin VB.TextBox Text2 
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
            Left            =   5715
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   94
            Text            =   "Nomtipart"
            Top             =   3240
            Width           =   1740
         End
         Begin VB.TextBox Text2 
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
            Left            =   5085
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   93
            Text            =   "TipArt"
            Top             =   3240
            Width           =   570
         End
         Begin VB.TextBox Text2 
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
            Left            =   3330
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   92
            Text            =   "Nombre articulo"
            Top             =   3240
            Width           =   1740
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
            Left            =   525
            MaxLength       =   12
            TabIndex        =   83
            Tag             =   "Num.Albaran|N|N|||albaran_envase|numalbar||S|"
            Text            =   "numalbar"
            Top             =   3240
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
            Index           =   9
            Left            =   1485
            MaxLength       =   12
            TabIndex        =   80
            Tag             =   "Num.Linea|N|N|||albaran_envase|numlinea|00|S|"
            Text            =   "numlinea"
            Top             =   3240
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
            Index           =   10
            Left            =   2295
            MaxLength       =   16
            TabIndex        =   78
            Tag             =   "Artículo|T|N|||albaran_envase|codartic||N|"
            Text            =   "Articulo"
            Top             =   3240
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
            Index           =   11
            Left            =   8820
            MaxLength       =   7
            TabIndex        =   81
            Tag             =   "Cantidad|N|N|||albaran_envase|cantidad|###,##0||"
            Text            =   "cantidad"
            Top             =   3240
            Visible         =   0   'False
            Width           =   735
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   1
            Left            =   45
            TabIndex        =   77
            Top             =   45
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
         Begin MSDataGridLib.DataGrid DataGrid3 
            Bindings        =   "frmVtasAlbSocios.frx":00C2
            Height          =   3225
            Left            =   45
            TabIndex        =   84
            Top             =   540
            Width           =   11340
            _ExtentX        =   20003
            _ExtentY        =   5689
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
         Begin MSAdodcLib.Adodc Adoaux 
            Height          =   330
            Index           =   0
            Left            =   1395
            Top             =   135
            Visible         =   0   'False
            Width           =   1425
            _ExtentX        =   2514
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
         Left            =   -73515
         MaxLength       =   30
         TabIndex        =   72
         Tag             =   "Variedad|N|N|||albaran_variedad|codvarie||N|"
         Text            =   "variedad"
         Top             =   1935
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
         Index           =   1
         Left            =   -74190
         MaxLength       =   15
         TabIndex        =   71
         Tag             =   "Num.Linea|N|N|||albaran_variedad|numlinea|00|S|"
         Text            =   "numlinea"
         Top             =   1935
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
         Index           =   0
         Left            =   -74550
         MaxLength       =   7
         TabIndex        =   70
         Tag             =   "Num.Albaran|N|N|||albaran_variedad|numalbar||S|"
         Text            =   "numpedi"
         Top             =   1935
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
         Index           =   4
         Left            =   -65025
         MaxLength       =   5
         TabIndex        =   69
         Tag             =   "Calibre|N|N|||albaran_calibre|codcalib|00|N|"
         Text            =   "calib"
         Top             =   3915
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
         Index           =   3
         Left            =   -66090
         MaxLength       =   12
         TabIndex        =   68
         Tag             =   "Variedad|N|N|||albaran_calibre|codvarie|000000|N|"
         Text            =   "variedad"
         Top             =   3915
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
         Left            =   -67170
         MaxLength       =   12
         TabIndex        =   67
         Tag             =   "Num.Linea 1|N|N|||albaran_calibre|numline1||N|"
         Text            =   "numline1"
         Top             =   3915
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
         Index           =   1
         Left            =   -67980
         MaxLength       =   12
         TabIndex        =   66
         Tag             =   "Num.Linea|N|N|||albaran_calibre|numlinea|00|N|"
         Text            =   "numlinea"
         Top             =   3915
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
         Index           =   0
         Left            =   -68940
         MaxLength       =   12
         TabIndex        =   65
         Tag             =   "Num.Palet|N|N|||albaran_calibre|numpalet||S|"
         Text            =   "numpedid"
         Top             =   3915
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
         Index           =   5
         Left            =   -64335
         MaxLength       =   5
         TabIndex        =   64
         Text            =   "nomca"
         Top             =   3915
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
         Index           =   6
         Left            =   -63585
         MaxLength       =   30
         TabIndex        =   63
         Tag             =   "Num.Cajas|N|N|0||albaran_calibre|numcajas|#,##0||"
         Text            =   "numcajas"
         Top             =   3915
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
         Index           =   3
         Left            =   -72840
         MaxLength       =   30
         TabIndex        =   62
         Text            =   "nomvarie"
         Top             =   1935
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
         Left            =   -72165
         MaxLength       =   30
         TabIndex        =   61
         Tag             =   "Variedad Comercial|N|N|||albaran_variedad|codvarco|||"
         Text            =   "var.comer."
         Top             =   1935
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
         Index           =   5
         Left            =   -71400
         MaxLength       =   30
         TabIndex        =   60
         Text            =   "nom.var.comer"
         Top             =   1935
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
         Left            =   -69690
         MaxLength       =   30
         TabIndex        =   59
         Tag             =   "Forfait|N|N|||albaran_variedad|codforfait|||"
         Text            =   "forfait"
         Top             =   1935
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
         Index           =   8
         Left            =   -68880
         MaxLength       =   30
         TabIndex        =   58
         Tag             =   "Categoria|T|S|||albaran_variedad|categori|||"
         Text            =   "categ"
         Top             =   1935
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
         Index           =   9
         Left            =   -68070
         MaxLength       =   30
         TabIndex        =   57
         Tag             =   "Peso Bruto|N|N|||albaran_variedad|pesobrut|###,##0||"
         Text            =   "peso bruto"
         Top             =   1935
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
         Index           =   10
         Left            =   -67305
         MaxLength       =   30
         TabIndex        =   56
         Tag             =   "Peso Neto|N|S|||albaran_variedad|pesoneto|###,##0|N|"
         Text            =   "peso neto"
         Top             =   1935
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
         Index           =   6
         Left            =   -70635
         MaxLength       =   30
         TabIndex        =   54
         Tag             =   "Marca|N|N|||albaran_variedad|codmarca|000||"
         Text            =   "marca"
         Top             =   1935
         Visible         =   0   'False
         Width           =   405
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
         Left            =   -70410
         MaxLength       =   30
         TabIndex        =   53
         Text            =   "nom marca"
         Top             =   1935
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
         Index           =   12
         Left            =   -69150
         MaxLength       =   30
         TabIndex        =   52
         Text            =   "nom forf"
         Top             =   1935
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
         Index           =   13
         Left            =   -66540
         MaxLength       =   30
         TabIndex        =   51
         Tag             =   "Num.Cajas|N|S|||albaran_variedad|numcajas|#,##0|N|"
         Text            =   "num.caj"
         Top             =   1935
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
         Index           =   14
         Left            =   -65865
         MaxLength       =   30
         TabIndex        =   50
         Tag             =   "Total Palets|N|S|||albaran_variedad|totpalet|##0|N|"
         Text            =   "tot.palet"
         Top             =   1935
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
         Index           =   15
         Left            =   -65145
         MaxLength       =   30
         TabIndex        =   49
         Tag             =   "Prec.Profes.|N|S|||albaran_variedad|preciopro|#0.0000|N|"
         Text            =   "precio prof"
         Top             =   1935
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
         Index           =   7
         Left            =   -62310
         MaxLength       =   30
         TabIndex        =   48
         Tag             =   "Peso Neto|N|N|0||albaran_calibre|pesoneto|###,##0||"
         Text            =   "pesoneto"
         Top             =   3915
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   0
         Left            =   -74865
         TabIndex        =   55
         Top             =   405
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
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmVtasAlbSocios.frx":00D7
         Height          =   1725
         Left            =   -74910
         TabIndex        =   73
         Top             =   855
         Width           =   17670
         _ExtentX        =   31168
         _ExtentY        =   3043
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
      Begin MSComctlLib.ListView ListView1 
         Height          =   3285
         Left            =   90
         TabIndex        =   151
         Top             =   855
         Width           =   8345
         _ExtentX        =   14711
         _ExtentY        =   5794
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmVtasAlbSocios.frx":00EC
         Height          =   1710
         Left            =   -69105
         TabIndex        =   74
         Top             =   2610
         Width           =   11865
         _ExtentX        =   20929
         _ExtentY        =   3016
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
      Begin VB.Label Label3 
         Caption         =   "1234567890123456789012"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   645
         Left            =   -74880
         TabIndex        =   75
         Top             =   3120
         Width           =   6330
      End
      Begin VB.Label Label8 
         Caption         =   "Cobrado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   2
         Left            =   -72690
         TabIndex        =   187
         Top             =   3960
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label Label9 
         Caption         =   "Facturado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   2
         Left            =   -73950
         TabIndex        =   186
         Top             =   3960
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Image imgFact 
         Height          =   330
         Index           =   2
         Left            =   -74490
         ToolTipText     =   "Facturas asociadas"
         Top             =   3870
         Visible         =   0   'False
         Width           =   330
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3870
      Left            =   135
      TabIndex        =   26
      Top             =   765
      Width           =   17700
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
         Index           =   37
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   189
         Text            =   "Text2"
         Top             =   2745
         Width           =   4980
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
         Index           =   19
         Left            =   1665
         MaxLength       =   6
         TabIndex        =   6
         Tag             =   "Cod.Comsionista|N|S|0|999|albaran|codcomis|000||"
         Text            =   "Text1"
         Top             =   2745
         Width           =   825
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
         Left            =   15885
         MaxLength       =   10
         TabIndex        =   19
         Tag             =   "Comisiones Previstas|N|S|||albaran|comisionespre|###,##0.00||"
         Top             =   1485
         Width           =   1200
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
         Left            =   14175
         MaxLength       =   10
         TabIndex        =   12
         Tag             =   "Número CMR|N|S|||albaran|numerocmr|######0||"
         Top             =   765
         Width           =   1200
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
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   115
         Text            =   "Text2"
         Top             =   3240
         Width           =   4980
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
         MaxLength       =   6
         TabIndex        =   7
         Tag             =   "Cod.Almacen|N|N|0|999|albaran|codalmac|000||"
         Text            =   "Text1"
         Top             =   3240
         Width           =   825
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Pasa Aridoc"
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
         Left            =   15885
         TabIndex        =   13
         Tag             =   "Pasa Aridoc|N|N|||albaran|pasaridoc|0||"
         Top             =   765
         Width           =   1725
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
         Left            =   11355
         MaxLength       =   4
         TabIndex        =   16
         Tag             =   "Nro.Acta|N|S|||albaran|nroactas|##0||"
         Text            =   "Text3"
         Top             =   1485
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
         Index           =   12
         Left            =   9300
         MaxLength       =   15
         TabIndex        =   15
         Tag             =   "Nro.Contrato|T|S|||albaran|nrocontra|||"
         Text            =   "123456789012345"
         Top             =   1485
         Width           =   1950
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
         Left            =   12615
         MaxLength       =   3
         TabIndex        =   17
         Tag             =   "Nro.Palets|N|S|||albaran|totpalet|##0||"
         Text            =   "Text3"
         Top             =   1485
         Width           =   1365
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
         Left            =   1665
         MaxLength       =   6
         TabIndex        =   5
         Tag             =   "Agencia Transporte|N|N|0|999|albaran|codtrans|000||"
         Text            =   "Text1"
         Top             =   2250
         Width           =   825
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
         Index           =   6
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   42
         Text            =   "Text2"
         Top             =   2250
         Width           =   4980
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
         Left            =   1665
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "Tipo Mercado|N|N|0|999|albaran|codtimer|000||"
         Text            =   "Text1"
         Top             =   1755
         Width           =   825
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
         Index           =   5
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   40
         Text            =   "Text2"
         Top             =   1755
         Width           =   4980
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
         Left            =   12645
         MaxLength       =   10
         TabIndex        =   11
         Tag             =   "Fecha Pedido|F|S|||albaran|fechaped|dd/mm/yyyy||"
         Top             =   765
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
         Index           =   4
         Left            =   1665
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "Cod.Destino|N|N|0|999|albaran|coddesti|000||"
         Text            =   "Text1"
         Top             =   1260
         Width           =   825
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
         Index           =   4
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   37
         Text            =   "Text2"
         Top             =   1260
         Width           =   4980
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
         Left            =   14175
         MaxLength       =   10
         TabIndex        =   18
         Tag             =   "Portes Previstos|N|S|||albaran|portespre|###,##0.00||"
         Top             =   1485
         Width           =   1200
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
         Left            =   9330
         MaxLength       =   40
         TabIndex        =   9
         Tag             =   "Matricula Remolque|T|S|||albaran|matrirem|||"
         Text            =   "Text3"
         Top             =   765
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
         Left            =   4320
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Albarán|F|N|||albaran|fechaalb|dd/mm/yyyy||"
         Top             =   270
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
         Index           =   7
         Left            =   7605
         MaxLength       =   12
         TabIndex        =   8
         Tag             =   "Matricula Vehiculo|T|S|||albaran|matriveh|||"
         Top             =   765
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
         Left            =   11355
         MaxLength       =   7
         TabIndex        =   10
         Tag             =   "Nº Pedido|N|S|||albaran|numpedid|0000000||"
         Text            =   "Text3"
         Top             =   765
         Width           =   1140
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
         Index           =   3
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   31
         Text            =   "Text2"
         Top             =   765
         Width           =   4980
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
         Index           =   15
         Left            =   7605
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Tag             =   "Observaciones|T|S|||albaran|observac|||"
         Top             =   2250
         Width           =   9975
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
         Left            =   1665
         MaxLength       =   6
         TabIndex        =   2
         Tag             =   "Cod. Socio|N|N|0|999999|albaran|codsocio|000000||"
         Text            =   "Text1"
         Top             =   765
         Width           =   825
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
         Left            =   7605
         MaxLength       =   12
         TabIndex        =   14
         Tag             =   "Referencia Cl|T|S|||albaran|refclien|||"
         Text            =   "Text3"
         Top             =   1485
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
         Left            =   1440
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Nº Albarán|N|S|||albaran|numalbar|0000000|S|"
         Text            =   "Text1 7"
         Top             =   270
         Width           =   1080
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   20
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   195
         Tag             =   "Cod. Cliente|N|N|0|999999|albaran|codclien|000000||"
         Text            =   "Text1"
         Top             =   780
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Comisionista"
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
         Index           =   46
         Left            =   135
         TabIndex        =   190
         Top             =   2790
         Width           =   1215
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1395
         ToolTipText     =   "Buscar Comsionista"
         Top             =   2790
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Comis.Previstas"
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
         Left            =   15885
         TabIndex        =   188
         Top             =   1215
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "Número CMR"
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
         Index           =   44
         Left            =   14175
         TabIndex        =   184
         Top             =   495
         Width           =   1320
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
         Index           =   1
         Left            =   135
         TabIndex        =   116
         Top             =   3285
         Width           =   810
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1395
         ToolTipText     =   "Buscar Agencia"
         Top             =   3285
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
         Left            =   11340
         TabIndex        =   46
         Top             =   1215
         Width           =   975
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
         Left            =   9315
         TabIndex        =   45
         Top             =   1215
         Width           =   1515
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
         Left            =   12645
         TabIndex        =   44
         Top             =   1215
         Width           =   1020
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1395
         ToolTipText     =   "Buscar Agencia"
         Top             =   2295
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
         Left            =   135
         TabIndex        =   43
         Top             =   2295
         Width           =   810
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1395
         ToolTipText     =   "Buscar T.Mercado"
         Top             =   1800
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
         Left            =   135
         TabIndex        =   41
         Top             =   1800
         Width           =   1035
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   13725
         Picture         =   "frmVtasAlbSocios.frx":0101
         ToolTipText     =   "Buscar fecha"
         Top             =   495
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F.Pedido"
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
         Left            =   12645
         TabIndex        =   39
         Top             =   495
         Width           =   870
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1395
         ToolTipText     =   "Buscar Destino"
         Top             =   1305
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
         Left            =   135
         TabIndex        =   38
         Top             =   1305
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "Portes Previstos"
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
         Left            =   14175
         TabIndex        =   36
         Top             =   1215
         Width           =   1680
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
         Left            =   9315
         TabIndex        =   35
         Top             =   495
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Albaran"
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
         Left            =   2565
         TabIndex        =   34
         Top             =   315
         Width           =   1455
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
         Left            =   7605
         TabIndex        =   33
         Top             =   495
         Width           =   1635
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   4050
         Picture         =   "frmVtasAlbSocios.frx":018C
         ToolTipText     =   "Buscar fecha"
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Pedido"
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
         Left            =   11340
         TabIndex        =   32
         Top             =   495
         Width           =   750
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   9135
         ToolTipText     =   "Zoom descripción"
         Top             =   1980
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
         Left            =   7605
         TabIndex        =   30
         Top             =   1980
         Width           =   1485
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
         Left            =   7605
         TabIndex        =   29
         Top             =   1215
         Width           =   1560
      End
      Begin VB.Label Label1 
         Caption         =   "Socio"
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
         Left            =   135
         TabIndex        =   28
         Top             =   810
         Width           =   540
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1395
         ToolTipText     =   "Buscar Socio"
         Top             =   810
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Albarán"
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
         Left            =   135
         TabIndex        =   27
         Top             =   315
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   135
      TabIndex        =   24
      Top             =   9450
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
         TabIndex        =   25
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
      Left            =   16800
      TabIndex        =   22
      Top             =   9585
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
      Left            =   15630
      TabIndex        =   21
      Top             =   9570
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
      Left            =   16785
      TabIndex        =   23
      Top             =   9585
      Visible         =   0   'False
      Width           =   1065
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
      Left            =   855
      Top             =   8145
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
      Left            =   675
      Top             =   8145
      Visible         =   0   'False
      Width           =   1425
      _ExtentX        =   2514
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
      Left            =   17475
      TabIndex        =   203
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
         HelpContextID   =   2
         Shortcut        =   ^I
      End
      Begin VB.Menu mnOrdenCarga 
         Caption         =   "&Orden de Carga"
         Enabled         =   0   'False
         Shortcut        =   ^O
         Visible         =   0   'False
      End
      Begin VB.Menu mnCMR 
         Caption         =   "&CMR"
         Enabled         =   0   'False
         Shortcut        =   ^C
         Visible         =   0   'False
      End
      Begin VB.Menu mnGenerarFactura 
         Caption         =   "&Generar Factura"
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
Attribute VB_Name = "frmVtasAlbSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public NumAlbar As String  ' venimos de pedidos para insertar envases paletizacion

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
Private WithEvents frmLAlb As frmVtasLinAlbaranes 'Lineas de variedades de albaranes
Attribute frmLAlb.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmArt As frmManArtic 'Form Mto de Articulos
Attribute frmArt.VB_VarHelpID = -1
Private WithEvents frmPal As frmVtasPalets 'Form Mto de palets
Attribute frmPal.VB_VarHelpID = -1

Private WithEvents frmCli As frmClientes 'Form Mto de Clientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmTra As frmManAgencias 'Form Mto de Agencias de Transporte
Attribute frmTra.VB_VarHelpID = -1
Private WithEvents frmTra1 As frmManAgencias 'Form Mto de Comisionistas
Attribute frmTra1.VB_VarHelpID = -1
Private WithEvents frmMer As frmManTipMerc 'Form Mto de Tipos de Mercado
Attribute frmMer.VB_VarHelpID = -1
Private WithEvents frmDest As frmDestCli 'Form Mto de destinos de clientes
Attribute frmDest.VB_VarHelpID = -1
Private WithEvents frmAlm As frmManAlmProp 'Form Mto de almacenes propios
Attribute frmAlm.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'Form Mto de socios de recoleccion
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmAlb As frmBasico2 ' ayuda de albaranes
Attribute frmAlb.VB_VarHelpID = -1

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
Dim NumTabMto As Integer 'Indica quin nº de Tab està en modo Mantenimient


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
Dim Facturas As String

Dim Cliente As String

Dim CadB1 As String


Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 0 'Articulos
            Set frmArt = New frmManArtic
            frmArt.DatosADevolverBusqueda = "0|1|"
            frmArt.CodigoActual = txtAux(10).Text
            frmArt.Show vbModal
            Set frmArt = Nothing
            PonerFoco txtAux(10)
'        Case 1 'Costes
'            Set frmNCoste = New frmManNomCoste
'            frmNCoste.DatosADevolverBusqueda = "0|1|"
'            frmNCoste.CodigoActual = txtAux(9).Text
'            frmNCoste.Show vbModal
'            Set frmNCoste = Nothing
'            PonerFoco txtAux(9)
    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1

End Sub

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
'               If ModificaDesdeFormulario2(Me, 2, "Frame2") Then
                If ModificaCabecera Then
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
            Select Case ModificaLineas
                Case 1 'afegir llínia
                    InsertarLinea NumTabMto
                Case 2 'modificar llínies
                    ModificarLinea
                    PosicionarData
            End Select
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
            Select Case SSTab1.Tab
                Case 0
                    If ModificaLineas = 1 Then 'INSERTAR
                        ModificaLineas = 0
                        DataGrid1.AllowAddNew = False
                        If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
                    End If
                    ModificaLineas = 0
                    LLamaLineas Modo, 0, "DataGrid2"
                    PonerModo 2
                    DataGrid2.Enabled = True
                    If Not Data1.Recordset.EOF Then _
                        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            
                    'Habilitar las opciones correctas del menu segun Modo
                    PonerModoOpcionesMenu (Modo)
                    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
                    DataGrid2.Enabled = True
                Case 1
                    If ModificaLineas = 1 Then 'INSERTAR
                        ModificaLineas = 0
                        DataGrid3.AllowAddNew = False
                        If Not AdoAux(0).Recordset.EOF Then AdoAux(0).Recordset.MoveFirst
                    End If
                    ModificaLineas = 0
                    LLamaLineas Modo, 0, "DataGrid3"
                    PonerModo 2
                    DataGrid3.Enabled = True
                    If Not Data1.Recordset.EOF Then _
                        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            
                    'Habilitar las opciones correctas del menu segun Modo
                    PonerModoOpcionesMenu (Modo)
                    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
                    DataGrid3.Enabled = True
                    PonerFocoGrid DataGrid3
                Case 2
                    If ModificaLineas = 1 Then 'INSERTAR
                        ModificaLineas = 0
                        DataGrid4.AllowAddNew = False
                        If Not AdoAux(1).Recordset.EOF Then AdoAux(1).Recordset.MoveFirst
                    End If
                    ModificaLineas = 0
                    LLamaLineas Modo, 0, "DataGrid4"
                    PonerModo 2
                    DataGrid4.Enabled = True
                    If Not Data1.Recordset.EOF Then _
                        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            
                    'Habilitar las opciones correctas del menu segun Modo
                    PonerModoOpcionesMenu (Modo)
                    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
                    DataGrid4.Enabled = True
                    PonerFocoGrid DataGrid4
                Case 3
                    If ModificaLineas = 1 Then 'INSERTAR
                        ModificaLineas = 0
                        DataGrid5.AllowAddNew = False
                        If Not AdoAux(2).Recordset.EOF Then AdoAux(2).Recordset.MoveFirst
                    End If
                    ModificaLineas = 0
                    LLamaLineas Modo, 0, "DataGrid5"
                    PonerModo 2
                    DataGrid5.Enabled = True
                    If Not Data1.Recordset.EOF Then _
                        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            
                    'Habilitar las opciones correctas del menu segun Modo
                    PonerModoOpcionesMenu (Modo)
                    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
                    DataGrid5.Enabled = True
                    PonerFocoGrid DataGrid5
             End Select
            
'            PonerBotonCabecera True
    
            
            
            
'            Me.DataGrid1.Enabled = True
    End Select
End Sub
Private Sub BotonAnyadir()
Dim Sql As String

    LimpiarCampos 'Huida els TextBox
    
    text1(20).Text = Format(vParamAplic.ClienteVtas, "000000")
    text1(4).Text = DevuelveValor("select min(coddesti) from destinos where codclien = " & DBSet(vParamAplic.ClienteVtas, "N"))
    
    Sql = "select count(*) from destinos where codclien = " & DBSet(vParamAplic.ClienteVtas, "N") & " and coddesti = " & DBSet(text1(4).Text, "N")
    If TotalRegistros(Sql) = 0 Then
        MsgBox "No existe el destino del cliente de ventas asociado " & vParamAplic.ClienteVtas & ". Revise", vbExclamation
        Exit Sub
    End If
    PonerModo 3

    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la capçalera *******
'    Text1(0).Text = SugerirCodigoSiguienteStr("palets", "numpalet")
'    FormateaCampo Text1(0)
    cmbAux(0).ListIndex = -1
    
    text1(1).Text = Format(Now, "dd/mm/yyyy")
    text1(16).Text = Format(vParamAplic.Almacen, "000")
    Text2(18).Text = PonerNombreDeCod(text1(16), "salmpr", "nomalmac", "codalmac", "N")
    
    text1(20).Text = Format(vParamAplic.ClienteVtas, "000000")
    text1(4).Text = DevuelveValor("select min(coddesti) from destinos where codclien = " & DBSet(vParamAplic.ClienteVtas, "N"))
    Text1_LostFocus (20)
    
        
    LimpiarDataGrids
    
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
        MandaBusquedaPrevia "" & CadB1
    Else
        LimpiarCampos
        LimpiarDataGrids
        CadenaConsulta = "Select albaran.* "
        CadenaConsulta = CadenaConsulta & "from " & NombreTabla
        CadenaConsulta = CadenaConsulta & " where " & CadB1
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
Dim DeVarios As Boolean

'    'solo se puede modificar la factura si no esta contabilizada
'    If FactContabilizada Then
'        TerminaBloquear
'        Exit Sub
'    End If
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    PonerFoco text1(1) '*** 1r camp visible que siga PK ***
        
End Sub


Private Sub BotonModificarLinea(Index As Integer)
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
    
    NumTabMto = Index
    If Index = 2 Then NumTabMto = 3
    PonerModo 5, Index
 

    
    Select Case NumTabMto
        Case 1 ' envases
            vWhere = ObtenerWhereCP(False)
            vWhere = vWhere & " and numlinea=" & Data2.Recordset!NumLinea
            If Not BloqueaRegistro("albaran_envase", vWhere) Then
                TerminaBloquear
                Exit Sub
            End If
            If DataGrid3.Bookmark < DataGrid3.FirstRow Or DataGrid3.Bookmark > (DataGrid3.FirstRow + DataGrid3.VisibleRows - 1) Then
                J = DataGrid3.Bookmark - DataGrid3.FirstRow
                DataGrid3.Scroll 0, J
                DataGrid3.Refresh
            End If
            
        '    anc = ObtenerAlto(Me.DataGrid1)
            anc = DataGrid3.Top
            If DataGrid3.Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGrid3.RowTop(DataGrid3.Row) + 10
            End If
        
            For J = 8 To 10
                txtAux(J).Text = DataGrid3.Columns(J - 8).Text
            Next J
            Text2(0).Text = DataGrid3.Columns(3).Text
            Text2(1).Text = DataGrid3.Columns(4).Text
            Text2(2).Text = DataGrid3.Columns(5).Text
            
            cmbAux(0).Text = DataGrid3.Columns(7).Text
            txtAux(11).Text = DataGrid3.Columns(8).Text
            txtAux(21).Text = DataGrid3.Columns(9).Text ' fecha de movimiento
            txtAux(23).Text = DataGrid3.Columns(10).Text ' codigo de cliente
            
            ModificaLineas = 2 'Modificar
            LLamaLineas ModificaLineas, anc, "DataGrid3"
            
            'Añadiremos el boton de aceptar y demas objetos para insertar
            Me.lblIndicador.Caption = "MODIFICAR"
            PonerModoOpcionesMenu (Modo)
            PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
            DataGrid3.Enabled = True
            
'            PonerBotonCabecera False
            PonerFoco txtAux(10)
            Me.DataGrid3.Enabled = False


       Case 3
            vWhere = ObtenerWhereCP(False)
            vWhere = vWhere & " and numlinea=" & Data2.Recordset!NumLinea
            If Not BloqueaRegistro("albaran_palets", vWhere) Then
                TerminaBloquear
                Exit Sub
            End If
            If DataGrid5.Bookmark < DataGrid5.FirstRow Or DataGrid5.Bookmark > (DataGrid5.FirstRow + DataGrid5.VisibleRows - 1) Then
                J = DataGrid5.Bookmark - DataGrid5.FirstRow
                DataGrid5.Scroll 0, J
                DataGrid5.Refresh
            End If
            
        '    anc = ObtenerAlto(Me.DataGrid1)
            anc = DataGrid5.Top
            If DataGrid5.Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGrid5.RowTop(DataGrid5.Row) + 10
            End If
        
            For J = 12 To 14
                txtAux(J).Text = DataGrid5.Columns(J - 12).Text
            Next J
            
            ModificaLineas = 2 'Modificar
            LLamaLineas ModificaLineas, anc, "DataGrid5"
            
            'Añadiremos el boton de aceptar y demas objetos para insertar
            Me.lblIndicador.Caption = "MODIFICAR"
            PonerModoOpcionesMenu (Modo)
            PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
            DataGrid5.Enabled = True
            
'            PonerBotonCabecera False
            PonerFoco txtAux(14)
            Me.DataGrid5.Enabled = False
       
    End Select
    
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
            
        Case "DataGrid3"
            DeseleccionaGrid Me.DataGrid3
            b = (xModo = 1 Or xModo = 2)
             For jj = 8 To 11
                txtAux(jj).Height = DataGrid3.RowHeight - 10
                txtAux(jj).Top = alto + 5
                txtAux(jj).visible = b
            Next jj
            btnBuscar(0).Height = DataGrid3.RowHeight - 10
            btnBuscar(0).Top = alto + 5
            btnBuscar(0).visible = b
            For jj = 0 To 2
                Text2(jj).Height = DataGrid3.RowHeight - 10
                Text2(jj).Top = alto + 5
                Text2(jj).visible = b
            Next jj
            txtAux(8).visible = False
            txtAux(8).Enabled = False
            txtAux(9).visible = False
            txtAux(9).Enabled = False
            
            cmbAux(0).Top = alto + 5
            cmbAux(0).visible = b
            
        Case "DataGrid5"
            DeseleccionaGrid Me.DataGrid5
            txtAux(12).visible = False
            txtAux(12).Enabled = False
            txtAux(13).visible = False
            txtAux(13).Enabled = False
            
            b = (xModo = 1 Or xModo = 2)
            For jj = 14 To 14
                txtAux(jj).Height = DataGrid5.RowHeight - 10
                txtAux(jj).Top = alto + 5
                txtAux(jj).visible = b
            Next jj
            For jj = 8 To 13
                Text2(jj).Height = DataGrid5.RowHeight - 10
                Text2(jj).Top = alto + 5
                Text2(jj).visible = b
            Next jj
    End Select
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Facturas (scafac)
' y los registros correspondientes de las tablas cab. albaranes (scafac1)
' y las lineas de la factura (slifac)
Dim Cad As String
'Dim vTipoMov As CTiposMov

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    If Label8(0).visible Or Label9(0).visible Then
        If MsgBox("Este albarán está facturado y/o cobrado. ¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Sub
        End If
    End If
    
    '++monica
'    If CLng(Text1(0).Text) = 999999 Then
'        MsgBox "Este albarán no permite ser eliminado.", vbExclamation
'        Exit Sub
'    End If
    
    Cad = "Cabecera de Albaranes." & vbCrLf
    Cad = Cad & "-------------------------------------      " & vbCrLf & vbCrLf
    Cad = Cad & "Va a eliminar el Albarán:            "
    Cad = Cad & vbCrLf & "Nº Albarán:  " & Format(text1(0).Text, "0000000")
    Cad = Cad & vbCrLf & "Fecha:  " & Format(text1(1).Text, "dd/mm/yyyy")

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

Private Sub cmbAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub





Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim i As Byte

'    If LastCol = -1 Then Exit Sub

    'Datos de la tabla albaran_calibres
    If Not Data3.Recordset.EOF Then
        Label2(0).Caption = Data3.Recordset.Fields(3)
        Label2(1).Caption = Data3.Recordset.Fields(5)
        Label2(2).Caption = Data3.Recordset.Fields(7)
        Label2(3).Caption = Data3.Recordset.Fields(9)
        Label2(5).Caption = DBLet(Data3.Recordset.Fields(14), "N")
'        Label2(8).Caption = DBLet(Data3.Recordset.Fields(15), "N")
        Label2(4).Caption = DBLet(Data3.Recordset.Fields(16), "N")
        
        'Datos de la tabla albaran_calibres
        CargaGrid DataGrid1, Data2, True
        'Datos de la tabla albaran_costes
        CargaGrid DataGrid4, AdoAux(1), True
'        'Datos de gastos totales
'        CargarListView
        
        Text2(14).Text = TotalCostesEnvases(Data3.Recordset.Fields(0), Data3.Recordset.Fields(1), 1)
        Text2(14).Text = CCur(Text2(14).Text) + CCur(TotalCostesEnvases(Data3.Recordset.Fields(0), Data3.Recordset.Fields(1), 4))
        Text2(15).Text = TotalCostesEnvases(Data3.Recordset.Fields(0), Data3.Recordset.Fields(1), 2)
        Text2(16).Text = TotalCostesEnvases(Data3.Recordset.Fields(0), Data3.Recordset.Fields(1), 0)
        Text2(38).Text = TotalCostesEnvases(Data3.Recordset.Fields(0), Data3.Recordset.Fields(1), 3)
        'total gastos
        Text2(17).Text = CCur(ImporteSinFormato(DBLet(Text2(14), "N"))) + CCur(ImporteSinFormato(DBLet(Text2(15), "N"))) + CCur(ImporteSinFormato(DBLet(Text2(16), "N"))) + CCur(ImporteSinFormato(DBLet(Text2(38), "N")))
        'gastos/kilo
        If CCur(ImporteSinFormato(DBLet(Label2(4).Caption, "N"))) <> 0 Then
            Text2(19).Text = Round2(CCur(ImporteSinFormato(DBLet(Text2(17), "N"))) / CCur(ImporteSinFormato(DBLet(Label2(4).Caption, "N"))), 4)
        End If
        'gastos/caja
        If CCur(ImporteSinFormato(DBLet(Label2(5).Caption, "N"))) <> 0 Then
            Text2(20).Text = Round2(CCur(ImporteSinFormato(DBLet(Text2(17), "N"))) / CCur(ImporteSinFormato(DBLet(Label2(5).Caption, "N"))), 4)
        End If

        'albaran facturado
        For i = 0 To 2
            Label9(i).visible = (AlbaranSOCIO_Facturado(Data3.Recordset.Fields(0), Data3.Recordset.Fields(1)) = 1)
            If Label9(i).visible Then
                'factura cobrada
                '[Monica]16/04/2010:antes FacturaCobradaTesoreria
                'Label8(i).visible = (FacturaCobradaTesoreria(Data3.Recordset.Fields(0), Data3.Recordset.Fields(1)) = 1)
                Label8(i).visible = (AlbaranCobradoTesoreria(Data3.Recordset.Fields(0), Data3.Recordset.Fields(1)) = 1)
                'importe facturado: lo miramos de la factura
                Text2(21).Text = ImporteAlbaranSOCIO_Facturado(Data3.Recordset.Fields(0), Data3.Recordset.Fields(1))
            Else
                Label8(i).visible = False
                'importe facturado: precio provisional * kilos
                Text2(21).Text = Round2(CCur(ImporteSinFormato(DBLet(Label2(4).Caption, "N"))) * DBLet(Data3.Recordset.Fields(13).Value, "N"), 2)
            End If
        Next i
        
        'ventas / caja
        If CCur(ImporteSinFormato(DBLet(Label2(5).Caption, "N"))) <> 0 Then
            Text2(23).Text = Round2(CCur(ImporteSinFormato(DBLet(Text2(21), "N"))) / CCur(ImporteSinFormato(DBLet(Label2(5).Caption, "N"))), 4)
        End If
        'ventas / kilo
        If CCur(ImporteSinFormato(DBLet(Label2(4).Caption, "N"))) <> 0 Then
            Text2(22).Text = Round2(CCur(ImporteSinFormato(DBLet(Text2(21), "N"))) / CCur(ImporteSinFormato(DBLet(Label2(4).Caption, "N"))), 4)
        End If
        
        'valor fruta = importe venta - gastos
        Text2(24).Text = CCur(ImporteSinFormato(DBLet(Text2(21).Text, "N"))) - CCur(ImporteSinFormato(DBLet(Text2(17).Text, "N")))
        Text2(24).Text = Format(Text2(24).Text, "###,###,##0.00")
        
        'neto/kilo
        If CCur(ImporteSinFormato(DBLet(Label2(4).Caption, "N"))) <> 0 Then
            Text2(25).Text = Round2(CCur(ImporteSinFormato(DBLet(Text2(24), "N"))) / CCur(ImporteSinFormato(DBLet(Label2(4).Caption, "N"))), 4)
        End If
        For i = 0 To 2
            Me.imgFact(i).visible = Label9(0).visible
            Me.imgFact(i).Enabled = Label9(0).visible
        Next i
        Facturas = ""
        If Label9(0).visible Then
            Facturas = FacturasdeAlbaranSOCIO(Data3.Recordset.Fields(0), Data3.Recordset.Fields(1))
        End If
        
        Text2(14) = Format(Text2(14), "###,###,##0.00")
        Text2(15) = Format(Text2(15), "###,###,##0.00")
        Text2(16) = Format(Text2(16), "###,###,##0.00")
        Text2(38) = Format(Text2(38), "###,###,##0.00")
        Text2(17) = Format(Text2(17), "###,###,##0.00")
        Text2(19) = Format(Text2(19), "###,###,##0.0000")
        Text2(20) = Format(Text2(20), "###,###,##0.0000")
        Text2(21) = Format(Text2(21), "###,###,##0.00")
        Text2(22) = Format(Text2(22), "###,###,##0.0000")
        Text2(23) = Format(Text2(23), "###,###,##0.0000")
        Text2(25) = Format(Text2(25), "###,###,##0.0000")
        
    Else
        Label2(0).Caption = ""
        Label2(1).Caption = ""
        Label2(2).Caption = ""
        Label2(3).Caption = ""
        Label2(4).Caption = ""
        Label2(5).Caption = ""
        Label8(0).visible = False
        Label9(0).visible = False
        Label8(1).visible = False
        Label9(1).visible = False
        Label8(2).visible = False
        Label9(2).visible = False
        
        'Datos de la tabla palets_calibres
        CargaGrid DataGrid1, Data2, False
        'Datos de la tabla palets_costes
        CargaGrid DataGrid4, AdoAux(1), False
        
        For i = 14 To 17
            Text2(i) = ""
        Next i
        For i = 19 To 25
            Text2(i) = ""
        Next i
    End If
    
'    CargaForaGrid
End Sub

' doble click en el grid de variedades
Private Sub DataGrid2_DblClick()
    If Data3.Recordset.EOF Then Exit Sub

    Set frmLAlb = New frmVtasLinAlbaranes
    
    frmLAlb.ModoExt = 0
    frmLAlb.Albaran = Data3.Recordset.Fields(0).Value
    frmLAlb.Linea = Data3.Recordset.Fields(1).Value
    frmLAlb.Show vbModal
    
    Set frmLAlb = Nothing
End Sub


Private Sub DataGrid5_DblClick()

    If AdoAux(2).Recordset.EOF Then Exit Sub

    Set frmPal = New frmVtasPalets
    
    frmPal.DatosADevolverBusqueda = AdoAux(2).Recordset.Fields(2)
    frmPal.Show vbModal
    Set frmPal = Nothing

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
    Me.Icon = frmPpal.Icon
    
     'Icono de busqueda
    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next kCampo

    ' ICONITOS DE LA BARRA
    btnPrimero = 16
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 3   'Insertar
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Borrar
        .Buttons(5).Image = 1   'Botón Buscar
        .Buttons(6).Image = 2   'Botón Todos
        .Buttons(8).Image = 10  'Impresión de albaran
'        .Buttons(9).Image = 24  'Orden de Carga
'        .Buttons(10).Image = 23 'CMR
'        .Buttons(11).Image = 26 'Generar Factura
'        .Buttons(13).Image = 11  'Salir
'        .Buttons(btnPrimero).Image = 6  'Primero
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
'        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    
    With Me.Toolbar5
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26 'Generar Factura
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
    
    'icono de ver las facturas en donde aparece el albaran
    Me.imgFact(0).Picture = frmPpal.imgListComun.ListImages(25).Picture
    Me.imgFact(1).Picture = frmPpal.imgListComun.ListImages(25).Picture
    Me.imgFact(2).Picture = frmPpal.imgListComun.ListImages(25).Picture
    
    
    LimpiarCampos   'Limpia los campos TextBox
    CargaCombo
    
    CodTipoMov = "ALS" ' albaranes de venta de socio
    VieneDeBuscar = False
        
    '## A mano
    NombreTabla = "albaran"
    NomTablaLineas = "albaran_variedad" 'Tabla lineas de variedades
    Ordenacion = " ORDER BY albaran.numalbar"
    
    CadB1 = "albaran.codclien = " & vParamAplic.ClienteVtas
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    Data1.ConnectionString = conn
    Data1.RecordSource = "select * from albaran where numalbar = -1 "
    Data1.Refresh
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
'        CargaGrid DataGrid1, Data2, False
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    Label2(0).Caption = ""
    Label2(1).Caption = ""
    Label2(2).Caption = ""
    Label2(3).Caption = ""
    Label2(4).Caption = ""
    Label2(5).Caption = ""
   
    SSTab1.Tab = 0
    
    SSTab1.TabVisible(1) = False
    SSTab1.TabVisible(3) = False
    
'    If DatosADevolverBusqueda <> "" Then
'        Text1(0).Text = DatosADevolverBusqueda
'        HacerBusqueda
'        SSTab1.Tab = 1
'    Else
'        PonerModo 0
'    End If
    
    If DatosADevolverBusqueda = "" Then
        If NumAlbar = "" Then
            PonerModo 0
        Else
            text1(0).Text = NumAlbar
            HacerBusqueda
            If hcoCodMovim = "" Then
                SSTab1.Tab = 1
            Else
                SSTab1.Tab = 0
            End If
        End If
    Else
        BotonBuscar
    End If
    
    PrimeraVez = False
End Sub


Private Sub LimpiarCampos()
Dim i As Integer

    On Error Resume Next
    
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.cmbAux(0).ListIndex = -1
    For i = 0 To 7
        Label2(i).Caption = ""
    Next i
    Label8(0).visible = False
    Label9(0).visible = False
    Label8(1).visible = False
    Label9(1).visible = False
    Label8(2).visible = False
    Label9(2).visible = False
    imgFact(0).visible = False
    imgFact(0).Enabled = False
    imgFact(1).visible = False
    imgFact(1).Enabled = False
    imgFact(2).visible = False
    imgFact(2).Enabled = False
    CargarListView ""
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    If Modo = 4 Then TerminaBloquear
End Sub


Private Sub frmAlb_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String
Dim Aux As String
      
    If CadenaSeleccion <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        CadB = "numalbar = " & RecuperaValor(CadenaSeleccion, 1)
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
    text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod almacen
    Text2(indice + 2).Text = RecuperaValor(CadenaSeleccion, 2) 'Nombre del almacen
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Articulos
    txtAux(10).Text = RecuperaValor(CadenaSeleccion, 1) 'codartic
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
    If txtAux(10) <> "" Then
        Text2(1) = DevuelveDesdeBDNew(cAgro, "sartic", "codtipar", "codartic", txtAux(10), "N")
        Text2(2) = DevuelveDesdeBDNew(cAgro, "stipar", "nomtipar", "codtipar", Text2(1), "N")
    End If
'    VisualizaPrecio
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
        Screen.MousePointer = vbDefault
    End If
    Screen.MousePointer = vbDefault
End Sub



Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    If imgFec(0).Tag < 2 Then
        text1(CByte(imgFec(0).Tag) + 1).Text = Format(vFecha, "dd/mm/yyyy") '<===
    Else
        text1(CByte(imgFec(0).Tag) + 8).Text = Format(vFecha, "dd/mm/yyyy") '<===
    End If
    ' ********************************************
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000") 'Cod Cliente
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 3) 'Nombre del cliente
End Sub

Private Sub frmDest_DatoSeleccionado(CadenaSeleccion As String)
    text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Destino
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nombre del destino
    
    text1(5) = DevuelveDesdeBDNew(cAgro, "destinos", "codtimer", "codclien", text1(20).Text, "N", , "coddesti", text1(4), "N")
    text1(5) = Format(text1(5), "000")
    Text2(5) = DevuelveDesdeBDNew(cAgro, "tipomer", "nomtimer", "codtimer", text1(5).Text, "N")
    
    MostrarCadena text1(3), text1(4)
End Sub

' devolvemos la linea del datagrid en donde estabamos
Private Sub frmLAlb_DatoSeleccionado(CadenaSeleccion As String)
Dim vWhere As String
             
   PonerCamposLineas
   
   If CadenaSeleccion = "" Then Exit Sub
             
   vWhere = "(numalbar = " & RecuperaValor(CadenaSeleccion, 1) & " and numlinea = " & RecuperaValor(CadenaSeleccion, 2) & ")"
   SituarDataMULTI Data3, vWhere, "" ', Indicador
   
   PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
   PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
   

End Sub

Private Sub frmMer_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Tipos de Mercado
    text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Mercado
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Mercado
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de socios de recoleccion
    text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000") 'Cod de socios
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Descripcion
End Sub

Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Agencias de Transporte
    text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Agencias de Transporte
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Descripcion
End Sub

Private Sub frmTra1_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Agencias de Transporte (Comisionista)
    text1(19).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Codigo comisionista
    Text2(37).Text = RecuperaValor(CadenaSeleccion, 2) 'Descripcion
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     text1(indice).Text = vCampo
End Sub

Private Sub imgBuscar_Click(Index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. de Socio
            indice = 3
            PonerFoco text1(indice)
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
            frmSoc.DeConsulta = True
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            PonerFoco text1(indice)
        
        Case 1 'Cod. de Destino de Cliente
            If text1(3) = "" Then Exit Sub
            
            indice = 4
            PonerFoco text1(indice)
            Set frmDest = New frmDestCli
            frmDest.Cliente = text1(3)
            frmDest.DatosADevolverBusqueda = "0|1|"
            frmDest.Show vbModal
            Set frmDest = Nothing
            PonerFoco text1(indice)
            
        Case 2 ' Tipo de Mercado
            indice = 5
            PonerFoco text1(indice)
            Set frmMer = New frmManTipMerc
            frmMer.DatosADevolverBusqueda = "0|1|2|"
            frmMer.Show vbModal
            Set frmMer = Nothing
            PonerFoco text1(indice)
        
        Case 3 ' Agencia de Transporte
            indice = 6
            PonerFoco text1(indice)
            Set frmTra = New frmManAgencias
            frmTra.DatosADevolverBusqueda = "0|1|2|"
            frmTra.Show vbModal
            Set frmTra = Nothing
            PonerFoco text1(indice)
        
        Case 4 ' Almacen
            indice = 16
            PonerFoco text1(indice)
            Set frmAlm = New frmManAlmProp
            frmAlm.DatosADevolverBusqueda = "0|1|"
            frmAlm.Show vbModal
            Set frmAlm = Nothing
            PonerFoco text1(indice)
            
        Case 5 ' comisionista
            PonerFoco text1(19)
            Set frmTra1 = New frmManAgencias
            frmTra1.DatosADevolverBusqueda = "0|1|2|"
            frmTra1.Show vbModal
            Set frmTra1 = Nothing
            PonerFoco text1(19)
            
    End Select
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub imgFact_Click(Index As Integer)
Dim frmFac As frmVtasFactSocios
    If Facturas <> "" Then
        Set frmFac = New frmVtasFactSocios
        frmFac.Facturas = Facturas
        frmFac.Show vbModal
        Set frmFac = Nothing
    End If
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
        If text1(Index + 1).Text <> "" Then frmC.NovaData = text1(Index + 1).Text
    Else
        imgFec(0).Tag = Index '<===
        ' *** repasar si el camp es txtAux o Text1 ***
        If text1(Index + 8).Text <> "" Then frmC.NovaData = text1(Index + 8).Text
    End If
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    If Index < 2 Then
        PonerFoco text1(CByte(imgFec(0).Tag) + 1) '<===
    Else
        PonerFoco text1(CByte(imgFec(0).Tag) + 8) '<===
    End If
    ' ********************************************
End Sub


Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        indice = 15
        frmZ.pTitulo = "Observaciones del Albarán"
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

Private Sub mnGenerarAlb_Click()
'Generar Albarán
    
    If Data1.Recordset.EOF Then Exit Sub
    
'    BotonImprimir

End Sub

Private Sub mnImprimir_Click()
'Imprimir Factura
    
    If Data1.Recordset.EOF Then Exit Sub
    
    BotonImprimir
End Sub

Private Sub mnOrdenCarga_Click()
'Imprimir la Orden de Carga
    
    If Data1.Recordset.EOF Then Exit Sub
    
    If DBLet(Data1.Recordset!numpedid, "N") <> 0 Then BotonOrdenCarga
End Sub

Private Sub mnCMR_Click()
'Imprimir la Orden de Carga
    
    If Data1.Recordset.EOF Then Exit Sub
    
    BotonCMR
End Sub

Private Sub mnGenerarFactura_Click()
'Generacion de factura a partir del albaran aprovechando los precios provisionales
    
    If Data1.Recordset.EOF Then Exit Sub
    
    If BLOQUEADesdeFormulario(Me) Then
        BotonGenerarFactura Data1.Recordset.Fields(0).Value
        TerminaBloquear
    End If
End Sub


Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub


Private Sub mnModificar_Click()
    If Label8(0).visible Or Label9(0).visible Then
        If MsgBox("Este albarán está facturado y/o cobrado. ¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Sub
        End If
    End If

    If Modo = 5 Then 'Modificar lineas
        'bloquea la tabla cabecera de factura: scafac
        If BLOQUEADesdeFormulario(Me) Then
            'bloquear la tabla cabecera de albaranes de la factura: scafac1
            If BloqueaAlbxFac Then
                If BloqueaLineasFac Then BotonModificarLinea (NumTabMto)
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
    If Index = 3 Then 'codigo de cliente
        Cliente = text1(Index).Text
    End If
    ConseguirFoco text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 15 Or (Index = 15 And text1(15).Text = "") Then KEYpress KeyAscii
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

        
    If Not PerderFocoGnral(text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1, 10 'Fecha albaran y fecha de pedido
            '[Monica]28/08/2013: comprobamos q este dentro de campaña
            If text1(Index).Text <> "" Then PonerFormatoFecha text1(Index), True
        Case 2
            If text1(Index) <> "" Then PonerFormatoDecimal text1(Index), 9
        
        Case 3 'socio
            If PonerFormatoEntero(text1(Index)) Then Text2(Index).Text = PonerNombreDeCod(text1(Index), "rsocios", "nomsocio")
                
        Case 20 ' Destino del cliente
            If PonerFormatoEntero(text1(Index)) Then
                If vParamAplic.ClienteVtas <> "" Then
                    Text2(Index).Text = DevuelveDesdeBDNew(cAgro, "destinos", "nomdesti", "codclien", text1(20), "N", , "coddesti", text1(4), "N")
                    If Text2(Index).Text = "" Then
                        cadMen = "No existe el Destino: " & text1(Index).Text & vbCrLf
                        cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                        If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                            Set frmCli = New frmClientes
                            frmCli.DatosADevolverBusqueda = "0|1|"
                            text1(Index).Text = ""
                            TerminaBloquear
                            frmCli.Show vbModal
                            Set frmCli = Nothing
                            If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                        Else
                            text1(Index).Text = ""
                        End If
                        PonerFoco text1(Index)
                    Else
                        ' traemos el tipo de mercado del destino
                        text1(5).Text = DevuelveDesdeBDNew(cAgro, "destinos", "codtimer", "codclien", text1(20), "N", , "coddesti", text1(4), "N")
                        PonerFormatoEntero text1(5)
                        If text1(5) <> "" Then
                            Text2(5).Text = PonerNombreDeCod(text1(5), "tipomer", "nomtimer", "codtimer", "N")
                        End If
                        ' mostramos en el label3 la cadena
                        MostrarCadena text1(20), text1(4)
                    End If
                Else
                    MsgBox "Debe introducir previamente el cliente.", vbExclamation
                    PonerFoco text1(20)
                End If
            End If
            
        Case 5 'Tipo de Mercado
            If PonerFormatoEntero(text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(text1(Index), "tipomer", "nomtimer")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Tipo de Mercado: " & text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmMer = New frmManTipMerc
                        frmMer.DatosADevolverBusqueda = "0|1|"
                        frmMer.NuevoCodigo = text1(Index).Text
                        text1(Index).Text = ""
                        TerminaBloquear
                        frmMer.Show vbModal
                        Set frmMer = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        text1(Index).Text = ""
                    End If
                    PonerFoco text1(Index)
                End If
            End If
                
        Case 6 'Agencia de Transporte
            If PonerFormatoEntero(text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(text1(Index), "agencias", "nomtrans")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Agencia de Transporte: " & text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmTra = New frmManAgencias
                        frmTra.DatosADevolverBusqueda = "0|1|"
                        frmTra.NuevoCodigo = text1(Index).Text
                        text1(Index).Text = ""
                        TerminaBloquear
                        frmTra.Show vbModal
                        Set frmTra = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        text1(Index).Text = ""
                    End If
                    PonerFoco text1(Index)
                Else
                    If Not EsTransportista(text1(Index).Text) Then
                        MsgBox "Este código corresponde a un Comisionista." & vbCrLf & "No a una Agencia de Transporte. Revise.", vbExclamation
                                        
                    End If
                End If
            End If
    
    
        Case 19 'Comisionistas
            If PonerFormatoEntero(text1(Index)) Then
                Text2(37).Text = DevuelveDesdeBDNew(cAgro, "agencias", "nomtrans", "codtrans", text1(Index).Text, "N")
                If Text2(37).Text = "" Then
                    cadMen = "No existe el Comisionista: " & text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmTra1 = New frmManAgencias
                        frmTra1.DatosADevolverBusqueda = "0|1|"
                        frmTra1.NuevoCodigo = text1(Index).Text
                        text1(Index).Text = ""
                        TerminaBloquear
                        frmTra1.Show vbModal
                        Set frmTra1 = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        text1(Index).Text = ""
                    End If
                    PonerFoco text1(Index)
                Else
                    ' comprobamos que se trata de un comisionista
                    If EsTransportista(text1(Index)) Then
                        MsgBox "Este código corresponde a una Agencia de Transporte. " & vbCrLf & "No a un Comisionista. Revise.", vbExclamation
                        PonerFoco text1(Index)
                    End If
                
                End If
            End If
    
    
    
    
        Case 16 'Almacen
            If PonerFormatoEntero(text1(Index)) Then
                Text2(Index + 2).Text = PonerNombreDeCod(text1(Index), "salmpr", "nomalmac")
                If Text2(Index + 2).Text = "" Then
                    cadMen = "No existe el Almacén: " & text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmAlm = New frmManAlmProp
                        frmAlm.DatosADevolverBusqueda = "0|1|"
                        frmAlm.NuevoCodigo = text1(Index).Text
                        text1(Index).Text = ""
                        TerminaBloquear
                        frmAlm.Show vbModal
                        Set frmAlm = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        text1(Index).Text = ""
                    End If
                    PonerFoco text1(Index)
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
    CadB = ObtenerBusqueda(Me)
    If CadB = "" Then
        CadB = CadB1
    Else
        CadB = CadB & " and " & CadB1
    End If
    
    If chkVistaPrevia = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select albaran.* from " & NombreTabla & " LEFT JOIN albaran_variedad ON albaran.numalbar=albaran_variedad.numalbar "
        CadenaConsulta = CadenaConsulta & " WHERE " & CadB & " GROUP BY albaran.numalbar " & Ordenacion
'        CadenaConsulta = "select palets.* from " & NombreTabla
'        CadenaConsulta = CadenaConsulta & " WHERE " & CadB & " GROUP BY palets.numpalet " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)

    Set frmAlb = New frmBasico2

    AyudaAlbaranesSocio frmAlb, text1(0), CadB
    
    Set frmAlb = Nothing

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
Dim i As Integer


    On Error GoTo EPonerLineas
    
    If Data1.Recordset.EOF Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    For i = 1 To 5
        Select Case i
            Case 1
                CargaGrid DataGrid2, Data3, True
                '++monica
                If Data3.Recordset.RecordCount > 0 Then
                    CargaGrid DataGrid1, Data2, True
                Else
                    CargaGrid DataGrid1, Data2, False
                End If
                '++
            Case 2  ' envases
                If Data1.Recordset.RecordCount > 0 Then
                    CargaGrid DataGrid3, AdoAux(0), True
                Else
                    CargaGrid DataGrid3, AdoAux(0), False
                End If
            Case 3  ' costes
                If Data3.Recordset.RecordCount > 0 Then
                    CargaGrid DataGrid4, AdoAux(1), True
                Else
                    CargaGrid DataGrid4, AdoAux(1), False
                End If
            Case 4  ' palets
                If Data1.Recordset.RecordCount > 0 Then
                    CargaGrid DataGrid5, AdoAux(2), True
                Else
                    CargaGrid DataGrid5, AdoAux(2), False
                End If
            Case 5  ' totales
                If Data1.Recordset.RecordCount > 0 Then
                    CargarListView CStr(Data1.Recordset.Fields(0))
                    
                    CargarTotales
                    
                End If
        End Select
    Next i
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim BrutoFac As Single
Dim b As Boolean

    On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    
    b = PonerCamposForma2(Me, Data1, 2, "Frame2")
    
'    FormatoDatosTotales
    
    'poner descripcion campos
    Modo = 4
    
    Text2(3).Text = PonerNombreDeCod(text1(3), "rsocios", "nomsocio", "codsocio", "N") 'socio
    Text2(4).Text = DevuelveDesdeBDNew(cAgro, "destinos", "nomdesti", "codclien", text1(3), "N", , "coddesti", text1(4), "N") 'destino
    Text2(5).Text = PonerNombreDeCod(text1(5), "tipomer", "nomtimer", "codtimer", "N") 'tipo de mercado
    Text2(6).Text = PonerNombreDeCod(text1(6), "agencias", "nomtrans", "codtrans", "N") 'agencia
    Text2(18).Text = PonerNombreDeCod(text1(16), "salmpr", "nomalmac", "codalmac", "N") 'almacen
    Text2(37).Text = PonerNombreDeCod(text1(19), "agencias", "nomtrans", "codtrans", "N") 'comsionista
    
    
    MostrarCadena text1(3), text1(4)
    
    Modo = 2
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas
    
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
Private Sub PonerModo(Kmodo As Byte, Optional indFrame As Integer)
Dim i As Byte, NumReg As Byte
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
    If DatosADevolverBusqueda <> "" Or NumAlbar <> "" Then
        cmdRegresar.visible = b
    Else
        cmdRegresar.visible = False
    End If
        
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    DesplazamientoVisible b And Data1.Recordset.RecordCount > 1
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    'si estamos en modificar bloquea las compos que son clave primaria
    BloquearText1 Me, Modo
    BloquearCombo Me, Modo
    For i = 9 To 10
        BloquearTxt text1(i), Not (Modo = 1)
        text1(i).Enabled = (Modo = 1)
    Next i
    Me.Check1.Enabled = (Modo = 1)
    
    b = (Modo <> 1)
    'Campos Nº Pedido bloqueado y en azul
    BloquearTxt text1(0), b, True
'    BloquearTxt Text1(3), b 'referencia
    
    
    'Si no es modo lineas Boquear los TxtAux
    For i = 0 To txtAux.Count - 1
        txtAux(i).visible = False
        BloquearTxt txtAux(i), True
    Next i
    For i = 0 To cmbAux.Count - 1
        cmbAux(i).visible = False
        cmbAux(i).Enabled = True
    Next i
    For i = 0 To 2
        Text2(i).visible = False
        Text2(i).Enabled = True
    Next i
    For i = 7 To 13
        Text2(i).visible = False
        Text2(i).Enabled = True
    Next i
    
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
    
    ' destino y mercado solo se pueden buscar
    text1(4).Enabled = (Modo = 1)
    text1(5).Enabled = (Modo = 1)
    
    imgBuscar(1).Enabled = (Modo = 1)
    imgBuscar(1).visible = (Modo = 1)
    imgBuscar(2).Enabled = (Modo = 1)
    imgBuscar(2).visible = (Modo = 1)
    
    Label3.Caption = ""
    
'    Me.imgBuscar(1).visible = False
                    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
    ' *** si n'hi han llínies i algún tab que no te datagrid ***
    Select Case NumTabMto
        Case 1
            BloquearFrameAux Me, "FrameAux0", Modo, NumTabMto
        Case 3
            BloquearFrameAux Me, "FrameAux2", Modo, NumTabMto
    End Select
    
        
    ' ***************************
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
    
    If Label8(0).visible Or Label9(0).visible Then
        If MsgBox("Este albarán está facturado y/o cobrado. ¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Sub
        End If
    End If
    
    
    If BloqueaRegistro(NombreTabla, "numalbar = " & Data1.Recordset!NumAlbar) Then
'    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
        Select Case Index
            Case 0 'variedades
                Select Case Button.Index
                    Case 1 'añadir variedad
                        Set frmLAlb = New frmVtasLinAlbaranes
                        
                        frmLAlb.ModoExt = 3
                        frmLAlb.Albaran = Data1.Recordset.Fields(0).Value
                        frmLAlb.Show vbModal
                    
                        Set frmLAlb = Nothing
                    Case 2 'modificar variedad
                        Set frmLAlb = New frmVtasLinAlbaranes
                        
                        frmLAlb.ModoExt = 4
                        frmLAlb.Albaran = Data3.Recordset.Fields(0).Value
                        frmLAlb.Linea = Data3.Recordset.Fields(1).Value
                        frmLAlb.Show vbModal
                        
                        Set frmLAlb = Nothing
                        
                    Case 3 ' boton eliminar linea de variedades
                        BotonEliminarLinea 0
                    Case Else
                End Select
                PonerCampos
                TerminaBloquear
                
            Case Else 'envases o palets
                Select Case Button.Index
                    Case 1
                        BotonAnyadirLinea Index
                    Case 2
                        BotonModificarLinea Index
                    Case 3
                        BotonEliminarLinea Index
                    Case Else
                End Select
                
        End Select
        
    End If



End Sub


Private Sub BotonEliminarLinea(Index As Integer)
Dim Cad As String
Dim Sql As String

    On Error GoTo EEliminarLinea

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
    ' ***************************************************************************

    Select Case Index
        Case 0 'variedades
            ' *************** canviar la pregunta ****************
            Cad = "¿Seguro que desea eliminar la Variedad?"
            Cad = Cad & vbCrLf & "Albarán: " & Data3.Recordset.Fields(0)
            Cad = Cad & vbCrLf & "Variedad: " & Data3.Recordset.Fields(3)
            
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
                On Error GoTo EEliminarLinea
                Screen.MousePointer = vbHourglass
                NumRegElim = Data3.Recordset.AbsolutePosition
                
                If Not EliminarLinea Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                Else
                    If SituarDataTrasEliminar(Data3, NumRegElim) Then
                        PonerCampos
                    Else
                        LimpiarCampos
                        PonerModo 0
                    End If
                End If
            End If
            Screen.MousePointer = vbDefault
       Case 1 'envases
            ' *************** canviar la pregunta ****************
            Cad = "¿Seguro que desea eliminar el Envase?"
            Cad = Cad & vbCrLf & "Albarán: " & AdoAux(0).Recordset.Fields(0)
            Cad = Cad & vbCrLf & "Envase: " & AdoAux(0).Recordset.Fields(2) & "-" & AdoAux(0).Recordset.Fields(3)
            
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
                On Error GoTo EEliminarLinea
                Screen.MousePointer = vbHourglass
                NumRegElim = AdoAux(0).Recordset.AbsolutePosition
                TerminaBloquear
                Sql = "delete from albaran_envase where numalbar = " & AdoAux(0).Recordset.Fields(0)
                Sql = Sql & " and numlinea = " & AdoAux(0).Recordset.Fields(1)
                conn.Execute Sql
                
                SituarDataTrasEliminar AdoAux(0), NumRegElim
                
                CargaGrid DataGrid3, AdoAux(0), True
                SSTab1.Tab = 1

            End If
            Screen.MousePointer = vbDefault
       
       Case 2 'palets
            ' *************** canviar la pregunta ****************
            Cad = "¿Seguro que desea eliminar el Palet?"
            Cad = Cad & vbCrLf & "Albarán: " & AdoAux(2).Recordset.Fields(0)
            Cad = Cad & vbCrLf & "Palet: " & AdoAux(2).Recordset.Fields(1) & "-" & AdoAux(2).Recordset.Fields(2)
            
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
                On Error GoTo EEliminarLinea
                Screen.MousePointer = vbHourglass
                NumRegElim = AdoAux(2).Recordset.AbsolutePosition
                TerminaBloquear
               
                Sql = "delete from albaran_palets where numalbar = " & AdoAux(2).Recordset.Fields(0)
                Sql = Sql & " and numlinea = " & AdoAux(2).Recordset.Fields(1)
                conn.Execute Sql
                
                SituarDataTrasEliminar AdoAux(2), NumRegElim
                
                CargaGrid DataGrid5, AdoAux(2), True
                SSTab1.Tab = 3
            End If
            Screen.MousePointer = vbDefault
       
    End Select
       
EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Linea de Albarán", Err.Description

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
        Case 5  'Buscar
            mnBuscar_Click
        Case 6  'Todos
            BotonVerTodos
        
        Case 1  'Añadir
            mnNuevo_Click

        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
        Case 8  ' Impresion de albaran
            mnImprimir_Click
'        Case 9  ' Orden de Carga
'            mnOrdenCarga_Click
'        Case 10 ' CMR
'            mnCMR_Click
'        Case 11 ' Generar Factura
'            mnGenerarFactura_Click
'        Case 13    'Salir
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

    On Error GoTo ECargaGRid

    b = DataGrid1.Enabled
    Select Case vDataGrid.Name
        Case "DataGrid1"
            Opcion = 1
        Case "DataGrid2"
            Opcion = 2
        Case "DataGrid3" 'envases
            Opcion = 3
        Case "DataGrid4" 'costes
            Opcion = 4
        Case "DataGrid5" 'palets
            Opcion = 5
    End Select
    
    Sql = MontaSQLCarga(enlaza, Opcion)
    CargaGridGnral vDataGrid, vData, Sql, PrimeraVez
    
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
        Case "DataGrid1" 'albaran_calibres
'           SQL = "SELECT numalbar, numlinea, numline1, codvarie, codcalib, nomcalib, numcajas, pesoneto
            tots = "N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux(3)|T|Variedad|2000|;"
            tots = tots & "S|txtAux(4)|T|Calibre|1000|;S|txtAux(5)|T|Nombre Calibre|4000|;S|txtAux(6)|T|Cajas|1300|;S|txtAux(22)|T|Uds|1300|;S|txtAux(7)|T|Peso Neto|1600|;"
            arregla tots, DataGrid1, Me, 350
'            DataGrid1.Columns(11).Alignment = dbgCenter
'            DataGrid1.Columns(12).Alignment = dbgRight
'            DataGrid1.Columns(13).Alignment = dbgRight
'            DataGrid1.Columns(14).Alignment = dbgRight
                       
         Case "DataGrid2" 'albaran_variedad
'           SQL = "SELECT numalbar, numlinea, codvarie, nomvarie1, codvarco, nomvarie2, codmarca, nommarca, codforfait, nomforfait, categori, pesobrut, totpalet, preciopro, numcajas, pesoneto
            tots = "N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux3(3)|T|Variedad Real|2900|;N||||0|;"
            tots = tots & "S|txtAux3(5)|T|Var.Comercial|2900|;N||||0|;S|txtAux3(11)|T|Marca|2400|;N||||0|;S|txtAux3(12)|T|Forfait|2400|;S|txtAux3(8)|T|Cat.|500|;"
            tots = tots & "S|txtAux3(9)|T|Peso Bruto|1600|;S|txtAux3(14)|T|Palets|800|;S|txtAux3(15)|T|Pr.Prov.|800|;S|txtAux3(13)|T|Cajas|800|;S|txtAux3(16)|T|Uds|800|;S|txtAux3(10)|T|Peso Neto|1100|;"
            tots = tots & "N||||0|;"
            arregla tots, DataGrid2, Me, 350
            
            DataGrid2.Columns(3).Alignment = dbgLeft
            DataGrid2.Columns(5).Alignment = dbgLeft
            DataGrid2.Columns(7).Alignment = dbgLeft
            DataGrid2.Columns(9).Alignment = dbgLeft
                     
         Case "DataGrid3" 'albaran_envases
'       SQL = SELECT albaran_envase.numalbar, numlinea, albaran_envase.codartic, sartic.nomartic, sartic.codtipar, stipar.nomtipar, "
'             albaran_envase.tipomovi, CASE albaran_envase.tipomovi WHEN 0 THEN ""Salida"" WHEN 1 THEN ""Entrada"" END, albaran_envase.cantidad "
            tots = "N||||0|;N||||0|;"
            tots = tots & "S|txtAux(10)|T|Articulo|2000|;S|btnBuscar(0)|B|||;"
            tots = tots & "S|Text2(0)|T|Nombre|3500|;S|Text2(1)|T|Tipo|800|;S|Text2(2)|T|Denominacion|2900|;N||||0|;S|cmbAux(0)|C|Tipo Mov.|1900|;"
            tots = tots & "S|txtAux(11)|T|Cantidad|1600|;N||||0|;N||||0|;"
            arregla tots, DataGrid3, Me, 350
            
'            DataGrid3.Columns(3).Alignment = dbgLeft
'            DataGrid3.Columns(5).Alignment = dbgLeft
'            DataGrid3.Columns(7).Alignment = dbgLeft
'            DataGrid3.Columns(9).Alignment = dbgLeft
    
         Case "DataGrid4" 'albaran_costes
'SELECT albaran_costes.numalbar, numlinea, albaran_costes.tipogasto, CASE albaran_costes.tipogasto WHEN 0 THEN ""Costes"" WHEN 1 THEN ""Materiales"" WHEN 2 THEN ""Palets"" END, albaran_costes.codcoste, nombcoste.denominacion, albaran_costes.impcoste, albaran_costes.importes, albaran_costes.unidades "
            tots = "N||||0|;N||||0|;N||||0|;N|cmbAux(1)|C|Tipo Gasto|1900|;"
            tots = tots & "N|txtAux(17)|T|Coste|1000|;"
            tots = tots & "S|Text2(7)|T|Nombre|2400|;S|txtAux(20)|T|Cajas-Kg|1500|;S|txtAux(19)|T|Importe|2300|;S|txtAux(18)|T|Importe Coste|1500|;"
            arregla tots, DataGrid4, Me, 350
            
'            DataGrid3.Columns(3).Alignment = dbgLeft
'            DataGrid3.Columns(5).Alignment = dbgLeft
'            DataGrid3.Columns(7).Alignment = dbgLeft
'            DataGrid3.Columns(9).Alignment = dbgLeft
         Case "DataGrid5" 'albaran_palets
'       SQL = SELECT albaran_palets.numalbar, numlinea, numpalet, linconfe, CASE tipmercan WHEN 0 THEN ""Cooperativa"" WHEN 1 THEN ""Terceros"" WHEN 2 THEN ""Mezclado"" WHEN 3 THEN ""Otros"" END, fechaini, time(horaini), fechafin, time(horafin) "
            tots = "N||||0|;N||||0|;"
            tots = tots & "S|txtAux(14)|T|N.Palets|1500|;S|Text2(8)|T|Lin.Confec.|1700|;S|Text2(9)|T|Tipo Mercado|1900|;S|Text2(10)|T|Fec.Inicio|1400|;S|Text2(11)|T|Hora Inicio|1200|;S|Text2(12)|T|Fecha Fin|1300|;S|Text2(13)|T|Hora Fin|1200|;"
            arregla tots, DataGrid5, Me, 350
'            DataGrid5.Columns(3).Alignment = dbgCenter
'            DataGrid5.Columns(4).Alignment = dbgCenter
'            DataGrid5.Columns(5).Alignment = dbgCenter
'            DataGrid5.Columns(6).Alignment = dbgCenter
'            DataGrid5.Columns(7).Alignment = dbgCenter
'            DataGrid5.Columns(8).Alignment = dbgCenter
            
    End Select
    
    vDataGrid.HoldFields
    Exit Sub
    
ECargaGRid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 ' Generar Factura
            mnGenerarFactura_Click
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
Dim cadMen As String
Dim Sql As String

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
            
        Case 10 ' envase
            If txtAux(Index) <> "" Then
                Text2(0).Text = PonerNombreDeCod(txtAux(Index), "sartic", "nomartic")
                Text2(1) = DevuelveDesdeBDNew(cAgro, "sartic", "codtipar", "codartic", txtAux(10), "T")
                Text2(2) = ""
                If Text2(1) <> "" Then
                    If EsArticuloRetornable(Text2(1).Text) Then 'CByte(Text2(1)) = 1 Or CByte(Text2(1)) = 2 Or CByte(Text2(1)) = 3 Then
                        Text2(2) = DevuelveDesdeBDNew(cAgro, "stipar", "nomtipar", "codtipar", Text2(1), "T")
                    Else
                        MsgBox "El Tipo de Envase ha de ser retornable. Reintroduzca.", vbExclamation
                        txtAux(Index).Text = ""
                        Text2(0).Text = ""
                        Text2(1).Text = ""
                        Text2(2).Text = ""
                        PonerFoco txtAux(Index)
                        Exit Sub
                    End If
                End If
                    
                If Text2(0).Text = "" Then
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
                        Text2(0).Text = ""
                        Text2(1).Text = ""
                        Text2(2).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                End If
            Else
                Text2(0).Text = ""
                Text2(1).Text = ""
                Text2(2).Text = ""
            End If
        
        Case 11 'cantidad
            If PonerFormatoEntero(txtAux(Index)) Then
                cmdAceptar.SetFocus
            End If
            
        Case 14 'numero de palet
            If txtAux(Index) <> "" Then
                PonerFormatoEntero txtAux(Index)
                Sql = DevuelveDesdeBDNew(cAgro, "palets", "numpalet", "numpalet", txtAux(Index), "N")
                If Sql = "" Then
                    MsgBox "No existe el palet introducido. Reintroduzca.", vbExclamation
                    txtAux(Index) = ""
                    PonerFoco txtAux(Index)
                Else
                    ' el palet ha de ser del mismo cliente que el albaran
                    Sql = DevuelveDesdeBDNew(cAgro, "palets", "numpedid", "numpalet", Sql, "N")
                    If Sql = "" Then
                        MsgBox "El número de pedido asociado a este palet esta vacio. Reintroduzca.", vbExclamation
                        txtAux(Index) = ""
                        PonerFoco txtAux(Index)
                    Else
                        Sql = DevuelveDesdeBDNew(cAgro, "pedidos", "codclien", "numpedid", Sql, "N")
                        If CLng(Sql) <> CLng(DBLet(Data1.Recordset!CodClien, "N")) Then
                            MsgBox "El cliente del pedido asociado al palet no coincide con el cliente del albarán. Reintroduzca.", vbExclamation
                            txtAux(Index) = ""
                            PonerFoco txtAux(Index)
                        End If
                    End If
                End If
            End If
            cmdAceptar.SetFocus
    End Select
    
End Sub




Private Function Eliminar() As Boolean
Dim Sql As String, LEtra As String, Sql2 As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim Mens As String

    On Error GoTo FinEliminar

    If Data1.Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    Mens = "Eliminando Costes"
    b = EliminarCostes(Data1.Recordset.Fields(0))

    If b Then
        'Eliminar en tablas de cabecera de albaran
        '------------------------------------------
        Sql = " " & ObtenerWhereCP(True)
        
        'Lineas de envases (albaran_envase)
        conn.Execute "Delete from albaran_envase " & Sql
        
        'Lineas de coste (albaran_costes)
        conn.Execute "Delete from albaran_costes " & Sql
        
        'Lineas de palets (albaran_palets)
        conn.Execute "Delete from albaran_palets " & Sql
    
        'Lineas de calibres (albaran_calibre)
        conn.Execute "Delete from albaran_calibre " & Sql
    
        'Lineas de variedades
        conn.Execute "Delete from albaran_variedad " & Sql
        
        'Cabecera de albaran
        conn.Execute "Delete from " & NombreTabla & Sql
        
        'Decrementar contador si borramos el ult. palet
        Set vTipoMov = New CTiposMov
        vTipoMov.DevolverContador CodTipoMov, Val(text1(0).Text) ' "ALV", Val(Text1(0).Text)
        Set vTipoMov = Nothing
        
        'si este albarán esta asociado a pedidos actualizamos:
        'pedidos.numalbar=null
        'pedidos.fecalbar=null
        Sql2 = "update pedidos set numalbar = " & ValorNulo & ", fechaalb = " & ValorNulo
        Sql2 = Sql2 & Sql
        conn.Execute Sql2
        
        b = True
    End If
FinEliminar:
    If Err.Number <> 0 Or Not b Then
        MuestraError Err.Number, "Eliminar Albarán", Err.Description & " " & Mens
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
Dim Mens As String

    On Error GoTo FinEliminar

    b = False
    If Data3.Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    Mens = ""
    'Eliminar en tablas de paltes_variedad y albaran_calibre
    '------------------------------------------
    Sql = " where numalbar = " & Data3.Recordset.Fields(0)
    Sql = Sql & " and numlinea = " & Data3.Recordset.Fields(1)

    Mens = "Actualizando Costes"
    b = ActualizarCostes(Data3.Recordset.Fields(0), Data3.Recordset.Fields(1), False, Data3.Recordset!codforfait, DBLet(Data3.Recordset!CodPalet, "N"))

    '08/09/2009: si tuviera costes de portes en albaran_costes los eliminamos aquí
    ' o costes de comision
    conn.Execute "delete from albaran_costes " & Sql & " and (tipogasto = 2 or tipogasto = 3)"

    'Lineas de calibres (albaran_calibre)
    conn.Execute "Delete from albaran_calibre " & Sql

    'Lineas de variedades
    conn.Execute "Delete from albaran_variedad " & Sql
    
    
FinEliminar:
    If Err.Number <> 0 Or Not b Then
        MuestraError Err.Number, "Eliminar Variedad del Albarán ", Err.Description & " " & Mens
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

    CargaGrid DataGrid2, Data3, False 'variedades y calibres
    CargaGrid DataGrid1, Data2, False
    CargaGrid DataGrid3, Me.AdoAux(0), False 'envases
    CargaGrid DataGrid4, Me.AdoAux(1), False 'costes
    CargaGrid DataGrid5, Me.AdoAux(2), False 'palets
    
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
    
    Sql = " numalbar= " & text1(0).Text
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
    
    Select Case Opcion
    Case 1  ' calibres
        Sql = "SELECT numalbar, numlinea, numline1, albaran_calibre.codvarie, albaran_calibre.codcalib, nomcalib, numcajas, unidades, pesoneto" ', pesoneto "
        Sql = Sql & " FROM albaran_calibre, calibres WHERE albaran_calibre.codvarie = calibres.codvarie and "
        Sql = Sql & " albaran_calibre.codcalib = calibres.codcalib "
    Case 2  'variedades
        Sql = "SELECT albaran_variedad.numalbar, numlinea, albaran_variedad.codvarie, a.nomvarie as nomvarie1, albaran_variedad.codvarco, "
        Sql = Sql & " b.nomvarie as nomvarie2, albaran_variedad.codmarca, marcas.nommarca, albaran_variedad.codforfait, forfaits.nomconfe, "
        Sql = Sql & " categori, pesobrut, totpalet, preciopro, numcajas, unidades, pesoneto " ', preciodef, albaran_variedad.codincid, inciden.nomincid, "
        Sql = Sql & ", albaran_variedad.codpalet "
'        SQL = SQL & " impcomis, albaran_variedad.observac "
        Sql = Sql & " FROM albaran_variedad, variedades a, variedades b, marcas, forfaits, inciden " 'lineas de variedades del albaran
        Sql = Sql & " WHERE albaran_variedad.codvarie = a.codvarie "
        Sql = Sql & " and albaran_variedad.codvarco = b.codvarie"
        Sql = Sql & " and albaran_variedad.codmarca = marcas.codmarca "
        Sql = Sql & " and albaran_variedad.codforfait = forfaits.codforfait "
        Sql = Sql & " and albaran_variedad.codincid = inciden.codincid "
    Case 3  'envases
        Sql = "SELECT albaran_envase.numalbar, numlinea, albaran_envase.codartic, sartic.nomartic, sartic.codtipar, stipar.nomtipar, "
        Sql = Sql & " albaran_envase.tipomovi, CASE albaran_envase.tipomovi WHEN 0 THEN ""Salida"" WHEN 1 THEN ""Entrada"" END, albaran_envase.cantidad, albaran_envase.fechamov, albaran_envase.codclien "
        Sql = Sql & " FROM albaran_envase, sartic, stipar "
        Sql = Sql & " WHERE albaran_envase.codartic = sartic.codartic "
        Sql = Sql & " and sartic.codtipar = stipar.codtipar"
    Case 4  'costes  numalbar,numlinea,tipogasto,codcoste,impcoste,importes,unidades
        Sql = "SELECT albaran_costes.numalbar, numlinea, albaran_costes.tipogasto, CASE albaran_costes.tipogasto WHEN 0 THEN ""Costes"" WHEN 1 THEN ""Materiales"" WHEN 2 THEN ""Palets"" END, albaran_costes.codcoste, nombcoste.denominacion, albaran_costes.unidades, albaran_costes.importes, albaran_costes.impcoste "
        Sql = Sql & " FROM albaran_costes, nombcoste "
        Sql = Sql & " WHERE albaran_costes.tipogasto = 0 and albaran_costes.codcoste = nombcoste.codcoste "
    Case 5  'palets
        Sql = "SELECT albaran_palets.numalbar, numlinea, albaran_palets.numpalet, palets.linconfe, CASE palets.tipmercan WHEN 0 THEN ""Cooperativa"" WHEN 1 THEN ""Terceros"" WHEN 2 THEN ""Mezclado"" WHEN 3 THEN ""Otros"" END, palets.fechaini, time(palets.horaini), palets.fechafin, time(palets.horafin) "
        Sql = Sql & " FROM albaran_palets, palets " 'lineas de palets del albaran
        Sql = Sql & " WHERE albaran_palets.numpalet = palets.numpalet "
    End Select
    
    If enlaza Then
        Sql = Sql & " and " & ObtenerWhereCP(False)
        If Opcion = 1 Or Opcion = 4 Then Sql = Sql & " AND numlinea=" & Data3.Recordset.Fields!NumLinea
    Else
        Sql = Sql & " and numalbar = -1"
    End If
    Sql = Sql & " ORDER BY numalbar"
    If Opcion = 1 Or Opcion = 4 Then Sql = Sql & ", numlinea "
    MontaSQLCarga = Sql
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean, bAux As Boolean
Dim i As Integer

    b = ((Modo = 2) Or (Modo = 0)) And (NumAlbar = "") 'Or (Modo = 5 And ModificaLineas = 0)
    'Buscar
    Toolbar1.Buttons(5).Enabled = b
    Me.mnBuscar.Enabled = b
    'Vore Tots
    Toolbar1.Buttons(6).Enabled = b
    Me.mnVerTodos.Enabled = b
    'Añadir
    Toolbar1.Buttons(1).Enabled = b
    Me.mnModificar.Enabled = b
    
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0) And (NumAlbar = "") And Not (Check1.Value = 1)
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnEliminar.Enabled = b
    
    'Impresión de albaran
    Toolbar1.Buttons(8).Enabled = ((Modo = 2) And (NumAlbar = "")) Or (hcoCodMovim <> "")
    Me.mnImprimir.Enabled = ((Modo = 2) And (NumAlbar = "")) Or (hcoCodMovim <> "")
    
    'Generar Factura
    Toolbar5.Buttons(1).Enabled = (Modo = 2) And (NumAlbar = "")
    Me.mnGenerarFactura.Enabled = (Modo = 2) And (NumAlbar = "")
        

    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
'++monica: si insertamos lo he quitado
'    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
    b = (Modo = 2) And Not Check1.Value = 1
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = b And hcoCodMovim = ""
        
        If b Then
            Select Case i
              Case 0
                bAux = (b And Me.Data3.Recordset.RecordCount > 0) And (NumAlbar = "")
              Case 1
                bAux = (b And Me.AdoAux(0).Recordset.RecordCount > 0) And (NumAlbar = "")
              Case 2
                bAux = (b And Me.AdoAux(2).Recordset.RecordCount > 0) And (NumAlbar = "")
            End Select
        End If
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
Dim NroCopias As Integer

    If text1(0).Text = "" Then
        MsgBox "Debe seleccionar un Albarán para Imprimir.", vbInformation
        Exit Sub
    End If
    
    cadFormula = ""
    cadParam = ""
    cadselect = ""
    numParam = 0
    
    If MsgBox("¿Desea imprimir calibres?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
        cadParam = cadParam & "|pCalibre=1|"
        numParam = numParam + 1
    Else
        cadParam = cadParam & "|pCalibre=0|"
        numParam = numParam + 1
    End If
    
    '===================================================
    '============ PARAMETROS ===========================
    indRPT = 67 'Impresion de Albaran
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
      
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de albaran
    '---------------------------------------------------
    If text1(0).Text <> "" Then
        'Nº palet
        devuelve = "{" & NombreTabla & ".numalbar}=" & Val(text1(0).Text)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        devuelve = "numalbar = " & Val(text1(0).Text)
        If Not AnyadirAFormula(cadselect, devuelve) Then Exit Sub
    End If
    
    If Not HayRegParaInforme(NombreTabla, cadselect) Then Exit Sub
     
     NroCopias = 1 'DevuelveDesdeBDNew(cAgro, "clientes", "nrocopias", "codclien", Text1(3).Text, "N")
     
     With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 0
            .Titulo = "Impresión de Albarán"
            .ConSubInforme = True
            .NroCopias = NroCopias
            .Show vbModal
    End With

End Sub

Private Sub BotonOrdenCarga()
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadselect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String

    If text1(0).Text = "" Then
        MsgBox "Debe seleccionar un Albarán para Imprimir.", vbInformation
        Exit Sub
    End If
    
    cadFormula = ""
    cadParam = ""
    cadselect = ""
    numParam = 0
    
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
    If text1(0).Text <> "" Then
        'Nº pedido
        devuelve = "{palets.numpedid}=" & Val(text1(9).Text)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        devuelve = "numpedid = " & Val(text1(9).Text)
        If Not AnyadirAFormula(cadselect, devuelve) Then Exit Sub
    End If
    
    If Not HayRegParaInforme("palets", cadselect) Then Exit Sub
     
     With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .ConSubInforme = True
            .Opcion = 0
            .Titulo = "Orden de Carga"
            .Show vbModal
    End With
End Sub


Private Sub BotonCMR()
Dim frmCMR As frmVtasCMR

    Set frmCMR = New frmVtasCMR
    
    frmCMR.NumCod = Data1.Recordset.Fields(0).Value
    frmCMR.NomTrans = Text2(6).Text
    frmCMR.Show vbModal
    
    Set frmCMR = Nothing
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

Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim i As Byte
    
    cmbAux(1).Clear
    
    cmbAux(1).AddItem "Costes"
    cmbAux(1).ItemData(cmbAux(1).NewIndex) = 0
    
    cmbAux(1).AddItem "Materiales"
    cmbAux(1).ItemData(cmbAux(1).NewIndex) = 1
    
    cmbAux(1).AddItem "Portes"
    cmbAux(1).ItemData(cmbAux(1).NewIndex) = 2
    
    cmbAux(0).Clear
    
    cmbAux(0).AddItem "Salida"
    cmbAux(0).ItemData(cmbAux(0).NewIndex) = 0
    
    cmbAux(0).AddItem "Entrada"
    cmbAux(0).ItemData(cmbAux(0).NewIndex) = 1
    
    
End Sub

Private Function ModificaCabecera() As Boolean
Dim b As Boolean
Dim MenError As String

    On Error GoTo EModificarCab

    conn.BeginTrans
    b = True
    '++monica: añadido la fecha de movimiento a los envases retornables
    If CDate(Data1.Recordset!FechaAlb) <> CDate(text1(1).Text) Then
        MenError = "Actualizando Fecha de Movimiento de Envases Retornables"
        b = ModificarFechaMovimiento(Data1.Recordset.Fields(0), text1(1).Text)
    End If
    '++
    If b Then
        If CCur(Data1.Recordset!codAlmac) <> CCur(text1(16).Text) Then
        
            MenError = "Eliminando Costes"
            b = EliminarCostes(Data1.Recordset.Fields(0))
        
            If b Then b = ModificaDesdeFormulario2(Me, 2, "Frame2")
    
            If b Then
                MenError = "Insertando Costes"
                b = InsertarCostes(Data1.Recordset.Fields(0))
            End If
        Else
            ' solo actualizamos la tabla smoval
            MenError = "Actualizando Movimiento (smoval)"
            b = ActualizaMovimiento(MenError)
            
            If b Then b = ModificaDesdeFormulario2(Me, 2, "Frame2")
        End If
    End If
    '[Monica] 30/09/2010: modificamos el codigo de cliente en las lineas de envases retornables
    If b And CCur(Data1.Recordset!CodClien) <> CCur(ComprobarCero(text1(3).Text)) Then
        MenError = "Modificando Envases Retornables"
        b = ModificarClienteEnvasesRetornables(Data1.Recordset.Fields(0), text1(3).Text)
    End If


EModificarCab:
    If Err.Number <> 0 Or Not b Then
        MenError = "Modificando Albarán." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        b = False
    End If
    If b Then
        ModificaCabecera = True
        conn.CommitTrans
    Else
        ModificaCabecera = False
        conn.RollbackTrans
    End If
End Function


Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim Sql As String

    On Error GoTo EInsertarCab
    
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
        Sql = CadenaInsertarDesdeForm(Me)
        If Sql <> "" Then
            If InsertarOferta(Sql, vTipoMov) Then
                CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                PonerCadenaBusqueda
                PonerModo 2
                'Ponerse en Modo Insertar Lineas
'                BotonMtoLineas 0, "Variedades"
'                BotonAnyadirLinea
                Set frmLAlb = New frmVtasLinAlbaranes
                
                frmLAlb.ModoExt = 3
                frmLAlb.Albaran = CLng(text1(0).Text)
                frmLAlb.Show vbModal
                
                Set frmLAlb = Nothing
            End If
        End If
        text1(0).Text = Format(text1(0).Text, "0000000")
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
    'Comprobar si mientras tanto se incremento el contador de Albaranes
    'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
    Do
        devuelve = DevuelveDesdeBDNew(cAgro, NombreTabla, "numalbar", "numalbar", text1(0).Text, "N")
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
    MenError = "Error al insertar en la tabla Cabecera de Albaranes (" & NombreTabla & ")."
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
        
        MenError = "Error al actualizar el contador del Albarán."
    '    bol = vTipoMov.IncrementarContador("REG")
        vTipoMov.IncrementarContador (CodTipoMov)
'    End If
    
EInsertarOferta:
        If Err.Number <> 0 Then
            MenError = "Insertando Albarán." & vbCrLf & "----------------------------" & vbCrLf & MenError
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

    Sql = DevuelveDesdeBDNew(cAgro, "destinos", "codcaden", "codclien", vParamAplic.ClienteVtas, "N", , "coddesti", desti, "N")
    If Sql <> "" Then
        Label3.Caption = DevuelveDesdeBDNew(cAgro, "cadenas", "nomcaden", "codcaden", Sql, "N")
    Else
        Label3.Caption = ""
    End If

End Sub

'Private Sub CargaForaGrid()
'        If DataGrid2.Columns.Count <= 2 Then Exit Sub
'        ' *** posar als camps de fora del grid el valor de la columna corresponent ***
'        Text2(0) = DataGrid2.Columns(5).Text
'        Text2(1) = DataGrid2.Columns(7).Text
'        Text2(2) = DataGrid2.Columns(9).Text
'        Text2(7) = DataGrid2.Columns(10).Text
'
'        ' *** Si fora del grid n'hi han camps de descripció, posar-los valor ***
'        ' **********************************************************************
' End Sub

Private Sub InsertarLinea(Index As Integer)
'Inserta registre en les taules de Llínies
Dim nomFrame As String
Dim b As Boolean

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case Index
        Case 1: nomFrame = "FrameAux0" 'envases
        Case 2: nomFrame = "FrameAux1" 'costes
        Case 3: nomFrame = "FrameAux2" 'palets
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
            b = BloqueaRegistro("albaran", "numalbar = " & Data1.Recordset!NumAlbar)
            Select Case Index
                Case 1  ' *** els index de les llinies en grid (en o sense tab) ***
                    CargaGrid DataGrid3, AdoAux(0), True
                    If b Then BotonAnyadirLinea NumTabMto
                Case 3 ' *** els index dels tabs que NO tenen grid ***
                    CargaGrid DataGrid5, AdoAux(2), True
                    If b Then BotonAnyadirLinea NumTabMto
'                LLamaLineas NumTabMto, 0
            End Select
            SSTab1.Tab = NumTabMto
        End If
    End If
End Sub

Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vtabla As String
Dim anc As Single
Dim i As Integer
    
    ModificaLineas = 1 'Posem Modo Afegir Llínia
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Capçalera
        cmdAceptar_Click
        If ModificaLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    If Index = 2 Then NumTabMto = 3
    
    PonerModo 5, Index
    
    ' *** bloquejar la clau primaria de la capçalera ***
    BloquearTxt text1(0), True
    ' **************************************************

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case NumTabMto
        Case 1: vtabla = "albaran_envase"
        Case 3: vtabla = "albaran_palets"
    End Select
    ' ********************************************************
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case NumTabMto
        Case 1 ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", vWhere)
            ' ***************************************************************

            AnyadirLinea DataGrid3, AdoAux(0)
    
            anc = DataGrid3.Top
            If DataGrid3.Row < 0 Then
                anc = anc + 220
            Else
                anc = anc + DataGrid3.RowTop(DataGrid3.Row) + 5
            End If
            
            LLamaLineas ModificaLineas, anc, "DataGrid3"
        
            LimpiarCamposLin "FrameAux0"
            
            txtAux(8).Text = text1(0).Text 'numalbar
            txtAux(9).Text = NumF 'numlinea
            txtAux(21).Text = Format(Data1.Recordset!FechaAlb, "dd/mm/yyyy")
            txtAux(23).Text = Format(Data1.Recordset!CodClien, "000000")
            
            For i = 0 To 2
                Text2(i).Text = ""
            Next i
            
            BloquearTxt txtAux(10), False
'                    BloquearTxt txtaux(12), False
            PonerFoco txtAux(10)
                    
        ' *** si n'hi han llínies sense datagrid ***
        Case 3
            NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", vWhere)
            ' ***************************************************************

            AnyadirLinea DataGrid5, AdoAux(2)
    
            anc = DataGrid5.Top
            If DataGrid5.Row < 0 Then
                anc = anc + 220
            Else
                anc = anc + DataGrid5.RowTop(DataGrid5.Row) + 5
            End If
          
            LLamaLineas ModificaLineas, anc, "DataGrid5"
        
            LimpiarCamposLin "FrameAux2"
            txtAux(12).Text = text1(0).Text 'codclien
            txtAux(13).Text = NumF
            PonerFoco txtAux(14)
            For i = 8 To 13
                Text2(i).Text = ""
            Next i
        ' ******************************************
    End Select
End Sub

Private Sub LimpiarCamposLin(frameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, frameAux  'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub ModificarLinea()
'Modifica registre en les taules de Llínies
Dim nomFrame As String
Dim V As Integer
Dim Cad As String
    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 1: nomFrame = "FrameAux0" 'envases
        Case 3: nomFrame = "FrameAux2" 'palets
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
'            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
            ModificaLineas = 0

            Select Case NumTabMto
                Case 1

                    V = AdoAux(0).Recordset.Fields(1) 'el 2 es el nº de llinia
                    CargaGrid DataGrid3, AdoAux(0), True

                    ' *** si n'hi han tabs ***
                    SSTab1.Tab = 1

                    DataGrid3.SetFocus
                    AdoAux(0).Recordset.Find (AdoAux(0).Recordset.Fields(1).Name & " =" & V)

                    LLamaLineas ModificaLineas, 0, "DataGrid3"
                Case 3
                    V = AdoAux(2).Recordset.Fields(1) 'el 2 es el nº de llinia
                    CargaGrid DataGrid5, AdoAux(2), True

                    ' *** si n'hi han tabs ***
                    SSTab1.Tab = 3

                    DataGrid5.SetFocus
                    AdoAux(2).Recordset.Find (AdoAux(2).Recordset.Fields(1).Name & " =" & V)

                    LLamaLineas ModificaLineas, 0, "DataGrid5"
            End Select
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

Private Function DatosOkLlin(nomFrame As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
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

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " numalbar= " & Val(text1(0).Text)
    ' *******************************************************
    
    ObtenerWhereCab = vWhere
End Function

' ***** si n'hi han varios nivells de tabs *****
Private Sub SituarTab(numTab As Integer)
    On Error Resume Next
    
    If numTab = 0 Or numTab = 1 Or numTab = 2 Or numTab = 3 Then
        SSTab1.Tab = 2
    ElseIf numTab = 4 Then
        SSTab1.Tab = 2
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub
' **********************************************

Private Function ActualizaMovimiento(Mens As String) As Boolean
Dim Sql As String
    
    On Error GoTo eActualizaMovimiento
    
    
    Sql = "update smoval set fechamov = " & DBSet(text1(1).Text, "F") & ", codigope = " & DBSet(text1(3).Text, "N")
    Sql = Sql & " where tipomovi = 'ALV' and document = " & Data1.Recordset!NumAlbar
    Sql = Sql & " and codigope = " & Data1.Recordset!CodClien
    Sql = Sql & " and fechamov = " & DBSet(Data1.Recordset!FechaAlb, "F")
    
    conn.Execute Sql
    
eActualizaMovimiento:
    If Err.Number <> 0 Then
        Mens = Mens & vbCrLf & Err.Description
        ActualizaMovimiento = False
    Else
        ActualizaMovimiento = True
    End If
End Function

Private Sub CargarListView(Albaran As String)
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem

    On Error GoTo ECargar

    'Los encabezados
    ListView1.ColumnHeaders.Clear
    Me.ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Add , , "Nombre", 4770, dbgLeft
'    ListView1.ColumnHeaders.Add , , "Unidades", 1000, dbgRight
'    ListView1.ColumnHeaders.Add , , "Importe", 1300, dbgRight
    ListView1.ColumnHeaders.Add , , "Importe Coste", 2500, dbgRight
    
    If Albaran = "" Then Exit Sub
    
    Sql = "SELECT albaran_costes.codcoste, nombcoste.denominacion, sum(albaran_costes.unidades), round(sum(albaran_costes.impcoste)/sum(albaran_costes.unidades),4), sum(albaran_costes.impcoste) "
    Sql = Sql & " FROM albaran_costes, nombcoste "
    Sql = Sql & " WHERE albaran_costes.tipogasto = 0 and albaran_costes.codcoste = nombcoste.codcoste "
    Sql = Sql & " and albaran_costes.numalbar = " & Albaran
    Sql = Sql & " group by 1,2 "
    Sql = Sql & " order by 1,2 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Set ItmX = ListView1.ListItems.Add
        ItmX.Text = DBLet(Rs.Fields(1).Value, "T")
        
'        ItmX.SubItems(1) = Format(Rs.Fields(2).Value, "###,###0")
'        ItmX.SubItems(2) = Format(Rs.Fields(3).Value, "###,##0.0000")
        ItmX.SubItems(1) = Format(Rs.Fields(4).Value, "###,##0.0000")
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
ECargar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar gastos totales.", Err.Description
End Sub



Private Sub CargarTotales()
Dim Cajas As Long
Dim Kilos As Long
Dim ImpVentas As Currency
Dim Rs As ADODB.Recordset
Dim Sql As String

    Cajas = TotalRegistros("select sum(numcajas) from albaran_variedad where numalbar = " & Data1.Recordset.Fields(0))
    Kilos = TotalRegistros("select sum(pesoneto) from albaran_variedad where numalbar = " & Data1.Recordset.Fields(0))
    
    Text2(33).Text = TotalCostesEnvases(Data1.Recordset.Fields(0), -1, 1)
    Text2(33).Text = CCur(Text2(33).Text) + CCur(TotalCostesEnvases(Data1.Recordset.Fields(0), -1, 4))
    
    Text2(34).Text = TotalCostesEnvases(Data1.Recordset.Fields(0), -1, 2)
    Text2(35).Text = TotalCostesEnvases(Data1.Recordset.Fields(0), -1, 0)
    Text2(39).Text = TotalCostesEnvases(Data1.Recordset.Fields(0), -1, 3)
    
    'total gastos
    Text2(36).Text = CCur(ImporteSinFormato(DBLet(Text2(33), "N"))) + CCur(ImporteSinFormato(DBLet(Text2(34), "N"))) + CCur(ImporteSinFormato(DBLet(Text2(35), "N"))) + CCur(ImporteSinFormato(DBLet(Text2(39), "N")))
    'gastos/kilo
    If Kilos <> 0 Then
        Text2(32).Text = Round2(CCur(ImporteSinFormato(DBLet(Text2(36), "N"))) / Kilos, 4)
    End If
    'gastos/caja
    If Cajas <> 0 Then
        Text2(31).Text = Round2(CCur(ImporteSinFormato(DBLet(Text2(36), "N"))) / Cajas, 4)
    End If
    ImpVentas = 0
    
    Sql = "select numlinea, pesoneto, preciopro from albaran_variedad where numalbar = " & Data1.Recordset.Fields(0).Value
    Sql = Sql & " order by 1"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
        If AlbaranFacturado(Data1.Recordset.Fields(0).Value, DBLet(Rs.Fields(0).Value, "N")) = 1 Then
            'importe facturado: lo miramos de la factura
            ImpVentas = ImpVentas + ImporteAlbaranFacturado(Data1.Recordset.Fields(0).Value, DBLet(Rs.Fields(0).Value, "N"))
        Else
            'importe facturado: precio provisional * kilos
            ImpVentas = ImpVentas + Round2(DBLet(Rs.Fields(1).Value, "N") * DBLet(Rs.Fields(2).Value, "N"), 2)
        End If
    
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    Text2(30).Text = Format(ImpVentas, "###,###,##0.00")
    
    'ventas / caja
    If Cajas <> 0 Then
        Text2(28).Text = Round2(CCur(ImporteSinFormato(DBLet(Text2(30), "N"))) / Cajas, 4)
    End If
    'ventas / kilo
    If Kilos <> 0 Then
        Text2(29).Text = Round2(CCur(ImporteSinFormato(DBLet(Text2(30), "N"))) / Kilos, 4)
    End If
    
    'valor fruta = importe venta - gastos
    Text2(27).Text = CCur(ImporteSinFormato(DBLet(Text2(30).Text, "N"))) - CCur(ImporteSinFormato(DBLet(Text2(36).Text, "N")))
    Text2(27).Text = Format(Text2(27).Text, "###,###,##0.00")
    
    'neto/kilo
    If Kilos <> 0 Then
        Text2(26).Text = Round2(CCur(ImporteSinFormato(DBLet(Text2(27), "N"))) / Kilos, 4)
    End If
    
    
    Label2(7).Caption = Format(Cajas, "###,###,##0")
    Label2(6).Caption = Format(Kilos, "###,###,##0")
    Text2(33) = Format(Text2(33), "###,###,##0.00")
    Text2(34) = Format(Text2(34), "###,###,##0.00")
    Text2(35) = Format(Text2(35), "###,###,##0.00")
    Text2(36) = Format(Text2(36), "###,###,##0.00")
    Text2(39) = Format(Text2(39), "###,###,##0.00")
    Text2(32) = Format(Text2(32), "###,###,##0.0000")
    Text2(31) = Format(Text2(31), "###,###,##0.0000")
    Text2(29) = Format(Text2(29), "###,###,##0.0000")
    Text2(28) = Format(Text2(28), "###,###,##0.0000")
    Text2(26) = Format(Text2(26), "###,###,##0.0000")

End Sub

Private Sub BotonGenerarFactura(Albaran As String)
Dim Sql As String
Dim FecFactu As String
Dim vFacturaVta As CFacturaVta
Dim b As Boolean
Dim Observaciones As String

    Observaciones = DevuelveDesdeBDNew(cAgro, "rsocios", "observaciones", "codsocio", Data1.Recordset!CodSocio, "N")
    If Trim(Observaciones) <> "" Then
        MsgBox Observaciones, vbInformation, "Observaciones del socio"
    End If

    ' comprobamos si hay lineas con precio provisional = 0
    Sql = "select count(*) from albaransocio_variedad where numalbar = " & DBSet(Albaran, "N")
    Sql = Sql & " and (preciopro is null or preciopro = 0)"
    If TotalRegistros(Sql) <> 0 Then
        If MsgBox("Hay lineas de albaran sin precio provisional. ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
        End If
    End If
    
    FecFactu = InputBox("Fecha Factura:", "Fecha de Factura", Format(Now, "dd/mm/yyyy"))
    If EsFechaOK(FecFactu) Then
    
        '[Monica]20/06/2017: control de fechas que antes no estaba
        If vParamAplic.NumeroConta <> 0 Then
            ResultadoFechaContaOK = EsFechaOKConta(CDate(FecFactu))
            If ResultadoFechaContaOK > 0 Then
                If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
                Exit Sub
            End If
        End If
    
        Set vFacturaVta = New CFacturaVta
        b = vFacturaVta.PasarAlbaranAFacturaSOCIO("albaran.numalbar=" & Albaran, FecFactu)
        If b Then
            Data3.Refresh
            MsgBox "Proceso realizado correctamente.", vbExclamation
        End If
    Else
        MsgBox "Fecha de Factura incorrecta.", vbExclamation
    End If
End Sub



Private Function ModificarFechaMovimiento(Albaran As Long, Fechamov As String) As Boolean
Dim Sql As String
Dim b As Boolean
    
    On Error GoTo eModificarFechaMovimiento
        
    ModificarFechaMovimiento = False
    
    Sql = "update albaran_envase set fechamov = " & DBSet(Fechamov, "F")
    Sql = Sql & " where numalbar = " & DBSet(Albaran, "N")
    
    conn.Execute Sql
    
    ModificarFechaMovimiento = True
    Exit Function
    
eModificarFechaMovimiento:
    If Err.Number <> 0 Then
        ModificarFechaMovimiento = False
    End If
End Function


Private Function ModificarClienteEnvasesRetornables(Albaran As Long, ActCliente As String) As Boolean
Dim Sql As String
Dim b As Boolean
    
    On Error GoTo eModificarClienteEnvasesRetornables
        
    ModificarClienteEnvasesRetornables = False
    
    Sql = "update albaran_envase set codclien = " & DBSet(ActCliente, "N")
    Sql = Sql & " where numalbar = " & DBSet(Albaran, "N")
    
    conn.Execute Sql
    
    ModificarClienteEnvasesRetornables = True
    Exit Function
    
eModificarClienteEnvasesRetornables:
    If Err.Number <> 0 Then
        ModificarClienteEnvasesRetornables = False
    End If
End Function


