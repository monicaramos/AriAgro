VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmManArtic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Art�culos"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11715
   Icon            =   "frmManArtic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   765
      Index           =   0
      Left            =   240
      TabIndex        =   30
      Top             =   480
      Width           =   11295
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   990
         MaxLength       =   16
         TabIndex        =   0
         Tag             =   "C�digo de articulo|T|N|||sartic|codartic||S|"
         Text            =   "1234567890123456"
         Top             =   240
         Width           =   1530
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   3480
         MaxLength       =   40
         TabIndex        =   1
         Tag             =   "Nombre|T|N|||sartic|nomartic|||"
         Top             =   240
         Width           =   4140
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre "
         Height          =   255
         Left            =   2745
         TabIndex        =   32
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "C�digo"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   240
      TabIndex        =   18
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
         TabIndex        =   19
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10410
      TabIndex        =   17
      Top             =   6960
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9120
      TabIndex        =   16
      Top             =   6960
      Width           =   1035
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5460
      Left            =   270
      TabIndex        =   29
      Top             =   1350
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   9631
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   6
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos b�sicos"
      TabPicture(0)   =   "frmManArtic.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(11)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblSumaStocks"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "imgFec(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(16)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(9)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "imgBuscar(4)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "imgBuscar(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "imgBuscar(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "imgBuscar(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(5)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(6)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(8)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(17)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "imgBuscar(3)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text1(19)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text1(20)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "chkCtrStock(0)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtSumaStock"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text1(10)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text1(8)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text1(6)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text2(6)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text2(2)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text2(3)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text2(7)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text1(7)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text1(3)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text1(2)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text1(5)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text2(5)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "FrameDatosAlmacen"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text1(4)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text1(9)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text1(11)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Text1(12)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Text1(14)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Text1(16)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).ControlCount=   39
      TabCaption(1)   =   "Stocks Almacenes"
      TabPicture(1)   =   "frmManArtic.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameAux0"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Index           =   16
         Left            =   10080
         MaxLength       =   6
         TabIndex        =   65
         Text            =   "Text1"
         Top             =   4410
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Index           =   14
         Left            =   10080
         MaxLength       =   6
         TabIndex        =   64
         Text            =   "Text1"
         Top             =   4050
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Index           =   12
         Left            =   10080
         MaxLength       =   6
         TabIndex        =   63
         Text            =   "Text1"
         Top             =   3690
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Index           =   11
         Left            =   10080
         MaxLength       =   6
         TabIndex        =   62
         Text            =   "Text1"
         Top             =   3330
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Index           =   9
         Left            =   10080
         MaxLength       =   6
         TabIndex        =   61
         Text            =   "Text1"
         Top             =   2970
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   10080
         MaxLength       =   6
         TabIndex        =   60
         Text            =   "Text1"
         Top             =   2610
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Frame FrameDatosAlmacen 
         Caption         =   "Datos Relacionados con Almacen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   1860
         Left            =   6270
         TabIndex        =   45
         Top             =   1620
         Width           =   3630
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   18
            Left            =   2115
            MaxLength       =   10
            TabIndex        =   14
            Tag             =   "Fecha �ltima compra|F|S|||sartic|ultfecco|dd/mm/yyyy|N|"
            Text            =   "Text1"
            Top             =   1440
            Width           =   1320
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   17
            Left            =   2115
            MaxLength       =   12
            TabIndex        =   13
            Tag             =   "Precio Venta al p�blico|N|N|0|999999.0000|sartic|preciove|###,##0.0000|N|"
            Text            =   "Text1"
            Top             =   1080
            Width           =   1320
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   15
            Left            =   2115
            MaxLength       =   12
            TabIndex        =   12
            Tag             =   "Precio Ultima Compra|N|S|0|999999.0000|sartic|preciouc|###,##0.0000|N|"
            Text            =   "Text1"
            Top             =   720
            Width           =   1320
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   13
            Left            =   2100
            MaxLength       =   12
            TabIndex        =   11
            Tag             =   "Precio Medio Ponderado|N|S|0|999999.0000|sartic|preciomp|###,##0.0000|N|"
            Text            =   "Text1"
            Top             =   360
            Width           =   1320
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   1
            Left            =   1845
            Picture         =   "frmManArtic.frx":0044
            ToolTipText     =   "Buscar fecha"
            Top             =   1440
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "�lt. fecha compra"
            Height          =   255
            Index           =   15
            Left            =   270
            TabIndex        =   49
            Top             =   1485
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Precio Venta P�blico"
            Height          =   255
            Index           =   14
            Left            =   270
            TabIndex        =   48
            Top             =   1125
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Precio �ltima compra"
            Height          =   255
            Index           =   12
            Left            =   255
            TabIndex        =   47
            Top             =   780
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Precio Medio Ponderado"
            Height          =   255
            Index           =   10
            Left            =   255
            TabIndex        =   46
            Top             =   420
            Width           =   1815
         End
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   5
         Left            =   2790
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   44
         Text            =   "Text2"
         Top             =   1575
         Width           =   3285
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   1950
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "Cod. Tipo Unidad|N|N|0|99|sartic|codunida|00|N|"
         Text            =   "Text1"
         Top             =   1575
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   1950
         MaxLength       =   6
         TabIndex        =   2
         Tag             =   "Cod. Proveedor|N|N|0|999999|sartic|codprove|000000|N|"
         Text            =   "Text1"
         Top             =   855
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   1950
         MaxLength       =   4
         TabIndex        =   3
         Tag             =   "Cod. Familia|N|N|0|9999|sartic|codfamia|0000|N|"
         Text            =   "Text1"
         Top             =   1215
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   7
         Left            =   1950
         MaxLength       =   2
         TabIndex        =   6
         Tag             =   "Tipo de IVA|N|N|0|99|sartic|codigiva||N|"
         Text            =   "T"
         Top             =   2295
         Width           =   765
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   7
         Left            =   2790
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   43
         Text            =   "Text2"
         Top             =   2295
         Width           =   3285
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   2790
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   42
         Text            =   "Text2"
         Top             =   1215
         Width           =   3285
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   2
         Left            =   2790
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   41
         Text            =   "Text2"
         Top             =   855
         Width           =   3285
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   6
         Left            =   2790
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   40
         Text            =   "Text2"
         Top             =   1935
         Width           =   3285
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   1950
         MaxLength       =   2
         TabIndex        =   5
         Tag             =   "Cod. Tipo Art�culo|T|N|||sartic|codtipar||N|"
         Text            =   "Te"
         Top             =   1935
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   8
         Left            =   8340
         MaxLength       =   13
         TabIndex        =   9
         Tag             =   "C�digo de Barras|T|S|||sartic|codigoea||N|"
         Text            =   "1234567890123"
         Top             =   855
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   10
         Left            =   8340
         MaxLength       =   10
         TabIndex        =   10
         Tag             =   "Fecha de Alta|F|N|||sartic|fecaltas|dd/mm/yyyy|N|"
         Text            =   "Text1"
         Top             =   1260
         Width           =   1335
      End
      Begin VB.TextBox txtSumaStock 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         Height          =   315
         Left            =   8070
         Locked          =   -1  'True
         MaxLength       =   13
         TabIndex        =   39
         Text            =   "Text1"
         Top             =   4065
         Width           =   1590
      End
      Begin VB.CheckBox chkCtrStock 
         Caption         =   "�Control de stock?"
         Height          =   315
         Index           =   0
         Left            =   8160
         TabIndex        =   15
         Tag             =   "Control de stock|N|N|0|1|sartic|ctrstock||N|"
         Top             =   3600
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   735
         Index           =   20
         Left            =   285
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Tag             =   "Texto para compras|T|S|||sartic|textocom|||"
         Top             =   3870
         Width           =   5865
      End
      Begin VB.TextBox Text1 
         Height          =   735
         Index           =   19
         Left            =   285
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Tag             =   "Texto para Ventas|T|S|||sartic|textoven|||"
         Top             =   2850
         Width           =   5865
      End
      Begin VB.Frame FrameAux0 
         BorderStyle     =   0  'None
         Height          =   4440
         Left            =   -74760
         TabIndex        =   36
         Top             =   480
         Width           =   10920
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   9
            Left            =   9210
            MaxLength       =   8
            TabIndex        =   69
            Text            =   "Text3"
            Top             =   3930
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.TextBox Text3 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   0
            Left            =   90
            MaxLength       =   16
            TabIndex        =   68
            Tag             =   "C�digo Articulo|T|N|||salmac|codartic||S|"
            Text            =   "Text3"
            Top             =   3915
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.TextBox Text3 
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   1
            Left            =   810
            MaxLength       =   8
            TabIndex        =   20
            Tag             =   "C�digo Almacen|N|N|||salmac|codalmac|000|S|"
            Text            =   "Text3"
            Top             =   3915
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.ComboBox cmbAux 
            Height          =   315
            Index           =   0
            Left            =   10125
            TabIndex        =   27
            Tag             =   "Status inventario|N|N|0|1|salmac|statusin|0|S|"
            Text            =   "Combo2"
            Top             =   3915
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   5
            Left            =   6615
            MaxLength       =   16
            TabIndex        =   24
            Tag             =   "Stock M�ximo|N|S|||salmac|stockmax|#,###,###,##0.00|N|"
            Text            =   "Text3"
            Top             =   3915
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   7
            Left            =   8370
            MaxLength       =   10
            TabIndex        =   26
            Tag             =   "Fecha inventario|F|S|||salmac|fechainv|dd/mm/yyyy|N|"
            Text            =   "Text3"
            Top             =   3915
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   6
            Left            =   7470
            MaxLength       =   16
            TabIndex        =   25
            Tag             =   "Stock inventario|N|S|||salmac|stockinv|#,###,###,##0.00|N|"
            Text            =   "Text3"
            Top             =   3915
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   4
            Left            =   5760
            MaxLength       =   16
            TabIndex        =   23
            Tag             =   "Punto de Pedido|N|S|||salmac|puntoped|#,###,###,##0.00|N|"
            Text            =   "Text3"
            Top             =   3915
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   3
            Left            =   4815
            MaxLength       =   16
            TabIndex        =   22
            Tag             =   "Stock M�nimo|N|S|||salmac|stockmin|#,###,###,##0.00|N|"
            Text            =   "Text3"
            Top             =   3915
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   2
            Left            =   3735
            MaxLength       =   16
            TabIndex        =   21
            Tag             =   "Cantidad Stock|N|N|||salmac|canstock|#,###,###,##0.00|N|"
            Text            =   "Text3"
            Top             =   3915
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.TextBox txtAux2 
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   0
            Left            =   1710
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   67
            Text            =   "Text2"
            Top             =   3915
            Visible         =   0   'False
            Width           =   2010
         End
         Begin VB.CommandButton btnBuscar 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   300
            Index           =   0
            Left            =   1485
            MaskColor       =   &H00000000&
            TabIndex        =   28
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   3915
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.CommandButton btnBuscar 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   300
            Index           =   1
            Left            =   9000
            MaskColor       =   &H00000000&
            TabIndex        =   66
            ToolTipText     =   "Buscar cuenta contable"
            Top             =   3915
            Visible         =   0   'False
            Width           =   195
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   0
            Left            =   0
            TabIndex        =   37
            Top             =   0
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
            Bindings        =   "frmManArtic.frx":00CF
            Height          =   3825
            Index           =   0
            Left            =   0
            TabIndex        =   38
            Top             =   480
            Width           =   10830
            _ExtentX        =   19103
            _ExtentY        =   6747
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
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   8
            Left            =   9210
            MaxLength       =   40
            TabIndex        =   70
            Tag             =   "Hora Inventario|FH|S|||salmac|horainve|yyyy-mm-dd hh:mm:ss|N|"
            Text            =   "Text3"
            Top             =   3660
            Visible         =   0   'False
            Width           =   900
         End
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1650
         ToolTipText     =   "Buscar tipo unidad"
         Top             =   1575
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Tipo Unidad"
         Height          =   255
         Index           =   17
         Left            =   270
         TabIndex        =   59
         Top             =   1575
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de I.V.A."
         Height          =   255
         Index           =   8
         Left            =   270
         TabIndex        =   58
         Top             =   2295
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Familia"
         Height          =   255
         Index           =   6
         Left            =   270
         TabIndex        =   57
         Top             =   1215
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Cod.  Proveedor"
         Height          =   255
         Index           =   5
         Left            =   270
         TabIndex        =   56
         Top             =   855
         Width           =   1215
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1650
         Tag             =   "-1"
         ToolTipText     =   "Buscar proveedor"
         Top             =   855
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1650
         ToolTipText     =   "Buscar familia"
         Top             =   1215
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1650
         ToolTipText     =   "Buscar tipo IVA"
         Top             =   2295
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1650
         ToolTipText     =   "Buscar tipo art�culo"
         Top             =   1935
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Tipo Art�culo"
         Height          =   255
         Index           =   9
         Left            =   270
         TabIndex        =   55
         Top             =   1935
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo Asociado"
         Height          =   255
         Index           =   2
         Left            =   6510
         TabIndex        =   54
         Top             =   855
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de Alta"
         Height          =   255
         Index           =   16
         Left            =   6510
         TabIndex        =   53
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   8070
         Picture         =   "frmManArtic.frx":00E7
         ToolTipText     =   "Buscar fecha"
         Top             =   1260
         Width           =   240
      End
      Begin VB.Label lblSumaStocks 
         Caption         =   "Suma Stock Almacenes"
         Height          =   375
         Left            =   6990
         TabIndex        =   52
         Top             =   4035
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Texto para Compras"
         Height          =   240
         Index           =   2
         Left            =   285
         TabIndex        =   51
         Top             =   3675
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Texto para Ventas"
         Height          =   240
         Index           =   11
         Left            =   285
         TabIndex        =   50
         Top             =   2655
         Width           =   1575
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   4200
      Top             =   6960
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
      TabIndex        =   34
      Top             =   0
      Width           =   11715
      _ExtentX        =   20664
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
            Style           =   3
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
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Index           =   0
         Left            =   8520
         TabIndex        =   35
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10440
      TabIndex        =   33
      Top             =   6960
      Visible         =   0   'False
      Width           =   1035
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
Attribute VB_Name = "frmManArtic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MANOLO                   -+-+
' +-+- Men�: ARTICULOS                 -+-+
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

' *** declarar els formularis als que vaig a cridar ***
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1

Private WithEvents frmA As frmManAlmProp  'almacenes
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmFam As frmManFamilias 'Familias
Attribute frmFam.VB_VarHelpID = -1
Private WithEvents frmTUn As frmManTipUnid 'Tipo de Unidad
Attribute frmTUn.VB_VarHelpID = -1
Private WithEvents frmPro As frmManProve 'Proveedores
Attribute frmPro.VB_VarHelpID = -1
Private WithEvents frmTAr As frmManTipArtic 'Tipo de articulos
Attribute frmTAr.VB_VarHelpID = -1

Private WithEvents frmCtas As frmCtasConta 'cuentas contables
Attribute frmCtas.VB_VarHelpID = -1
'VRS:4.0.1
Private WithEvents frmTipIva As frmTipIVAConta  'Tipos de IVA de la contabilidad
Attribute frmTipIva.VB_VarHelpID = -1
' *****************************************************


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
    Select Case Index
        Case 0 'C�digo de Almacen
            Set frmA = New frmManAlmProp
            frmA.DatosADevolverBusqueda = "0|1|"
            frmA.Show vbModal
            Set frmA = Nothing
            
        Case 1
            Dim esq As Long
            Dim dalt As Long
            Dim menu As Long
            Dim obj As Object
        
            Set frmC = New frmCal
            
            esq = btnBuscar(Index).Left
            dalt = btnBuscar(Index).Top
                
            Set obj = btnBuscar(Index).Container
              
              While btnBuscar(Index).Parent.Name <> obj.Name
                    esq = esq + obj.Left
                    dalt = dalt + obj.Top
                    Set obj = obj.Container
              Wend
            
            menu = Me.Height - Me.ScaleHeight 'ac� tinc el heigth del men� i de la toolbar
        
            frmC.Left = esq + btnBuscar(Index).Parent.Left + 30
            frmC.Top = dalt + btnBuscar(Index).Parent.Top + btnBuscar(Index).Height + menu - 40
        
            imgFec(Index).Tag = 7 '<===
            ' *** repasar si el camp es txtAux o Text1 ***
            If Text3(7) <> "" Then frmC.NovaData = Text3(7).Text
            ' ********************************************
        
            frmC.Show vbModal
            Set frmC = Nothing
            ' *** repasar si el camp es txtAux o Text1 ***
            PonerFoco Text3(7) '<===
            ' ********************************************
                
    End Select

End Sub

Private Sub cmbAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim cad As String, Indicador As String
Dim bol As Boolean
Dim codigo As String

    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'B�SQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm2(Me, 1) Then
                    codigo = Text1(0).Text
                    InsetarArticulosPorAlmacen
                    PosicionarData codigo
                    CargaGrid 0, True
                    PonerModo 2
                End If
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario2(Me, 1) Then
                    TerminaBloquear
                    PosicionarData Text1(0).Text
                End If
            Else
                ModoLineas = 0
            End If
        ' *** si n'hi han ll�nies ***
        Case 5 'LL�NIES
'            If InsertarModificarLinea Then
''                DesBloqueaRegistroForm Text1(0)
'                TerminaBloquear
'                cad = "codalmac = " & Text3(0).Text & ""
'                If SituarData(Data4, cad, Indicador) Then
'                    ModificaLineas = 0
'                    lblIndicador.Caption = Indicador
'                    PonerModoFrame 0
'                    PonerSumaStocks
'                End If
'            End If
            Select Case ModoLineas
                Case 1 'afegir ll�nia
                    InsertarLinea
                Case 2 'modificar ll�nies
                    ModificarLinea
                    PosicionarData Data1.Recordset!codArtic
            End Select
            PonerSumaStocks
        ' **************************

    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
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
                PonerModo 1 'b�squeda
                ' *** posar de groc els camps visibles de la clau primaria de la cap�alera ***
                Text1(0).BackColor = vbYellow 'codartic
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
    
    'cargar IMAGES de busqueda
    For i = 0 To imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    
    ' *** si n'hi han tabs, per a que per defecte sempre es pose al 1r***
    Me.SSTab1.Tab = 0
    ' *******************************************************************
    
    LimpiarCampos   'Neteja els camps TextBox
    ' ******* si n'hi han ll�nies *******
    DataGridAux(0).ClearFields
    
    '*** canviar el nom de la taula i l'ordenaci� de la cap�alera ***
    NombreTabla = "sartic"
    Ordenacion = " ORDER BY codartic"
    
    'Mirem com est� guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    Data1.ConnectionString = conn
    '***** cambiar el nombre de la PK de la cabecera *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where codartic=-1"
    Data1.Refresh
    
    ModoLineas = 0
    CargaCombo 0
    
'    If DatosADevolverBusqueda = "" Then
'        PonerModo 0
'    Else
'        PonerModo 1 'b�squeda
'        ' *** posar de groc els camps visibles de la clau primaria de la cap�alera ***
'        Text1(0).BackColor = vbYellow 'codartic
'        ' ****************************************************************************
'    End If
End Sub

Private Sub LimpiarCampos()
    On Error Resume Next
    
    limpiar Me   'M�tode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    chkCtrStock(0).Value = 0
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub LimpiarCamposLin(frameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, frameAux  'M�tode general: Neteja els controls TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO s'habiliten, o no, els diversos camps del
'   formulari en funci� del modo en que anem a treballar
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
    'Posem visible, si es formulari de b�squeda, el bot� "Regresar" quan n'hi han datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = (Modo = 2)
    Else
        cmdRegresar.visible = False
    End If
    
    '=======================================
    b = (Modo = 2)
    'Posar Fleches de desplasament visibles
    Numreg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then Numreg = 2 'Nom�s es per a saber que n'hi ha + d'1 registre
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, Numreg
    
    '---------------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
       
    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
    'Si estem en Insertar a m�s neteja els camps Text1
    BloquearText1 Me, Modo
    BloquearCombo Me, Modo
    
    
    BloquearTxt Text1(13), Modo >= 2
    

    ' ***** bloquejar tots els controls visibles de la clau primaria de la cap�alera ***
'    If Modo = 4 Then _
'        BloquearTxt Text1(0), True 'si estic en  modificar, bloqueja la clau primaria
    ' **********************************************************************************
    
    ' **** si n'hi han imagens de buscar en la cap�alera *****
    BloquearImgBuscar Me, Modo, ModoLineas
    BloquearImgFec Me, 0, Modo
    BloquearImgFec Me, 1, Modo
    BloquearChk chkCtrStock(0), (Modo = 0 Or Modo = 2 Or Modo = 5)
    
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
Dim i As Byte
    
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
    Me.mnNuevo.Enabled = b
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(9).Enabled = b
    Me.mnEliminar.Enabled = b
    
    'Imprimir
    'Toolbar1.Buttons(12).Enabled = (b Or Modo = 0)
    Toolbar1.Buttons(12).Enabled = b
    Me.mnImprimir.Enabled = b
       
    ' *** si n'hi han ll�nies que tenen grids (en o sense tab) ***
    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.AdoAux(i).Recordset.RecordCount > 0)
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i
    
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
               
        Case 0 'stocks en almacenes
            Sql = "SELECT codartic,salmac.codalmac,salmpr.nomalmac,canstock,stockmin,puntoped,stockmax,stockinv,fechainv,horainve,statusin,CASE statusin WHEN 0 THEN ""No"" WHEN 1 THEN ""S�"" END "
            Sql = Sql & " FROM salmac, salmpr "
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE salmac.codartic = '-1'"
            End If
            Sql = Sql & " and salmac.codalmac = salmpr.codalmac "
            Sql = Sql & " ORDER BY salmac.codalmac"
            
    End Select
    
    MontaSQLCarga = Sql
End Function

Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Colectivos
    Text3(1).Text = RecuperaValor(CadenaSeleccion, 1) 'codalmacen
    FormateaCampo Text3(1)
    txtAux2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'nomalmacen
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
        '   Com la clau principal es �nica, en posar el sql apuntant
        '   al valor retornat sobre la clau ppal es suficient
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        ' **********************************
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmPro_DatoSeleccionado(CadenaSeleccion As String)
'proveedores
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo
    FormateaCampo Text1(2)
    Text2(2).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion

End Sub

Private Sub frmTAr_DatoSeleccionado(CadenaSeleccion As String)
'Tipos de articxulos
    Text1(6).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo
    FormateaCampo Text1(6)
    Text2(6).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

'VRS:4.0.1
Private Sub frmTipIva_DatoSeleccionado(CadenaSeleccion As String)
'Tipos de IVA (de la Contabilidad)
    Text1(7).Text = RecuperaValor(CadenaSeleccion, 1) 'codigiva
    FormateaCampo Text1(7)
    Text2(7).Text = RecuperaValor(CadenaSeleccion, 2) '% iva

End Sub

Private Sub frmTUn_DatoSeleccionado(CadenaSeleccion As String)
'Tipos de unidad
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo
    FormateaCampo Text1(5)
    Text2(5).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

' *** si n'hi ha buscar data, posar a les <=== el menor index de les imagens de buscar data ***
' NOTA: ha de coincidir l'index de la image en el del camp a on va a parar el valor
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
    
    menu = Me.Height - Me.ScaleHeight 'ac� tinc el heigth del men� i de la toolbar

    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40

    Select Case Index
        Case 0
            imgFec(1).Tag = 10 '<===
        
            ' *** repasar si el camp es Text3 o Text1 ***
            If Text1(10).Text <> "" Then frmC.NovaData = Text1(10).Text
            ' ********************************************
        
        Case 1
            imgFec(1).Tag = 18 '<===
            
            ' *** repasar si el camp es Text3 o Text1 ***
            If Text1(18).Text <> "" Then frmC.NovaData = Text1(18).Text
            ' ********************************************
    End Select
    

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es Text3 o Text1 ***
    PonerFoco Text1(CByte(imgFec(1).Tag)) '<===
    ' ********************************************
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es Text3 o Text1 ***
    Select Case imgFec(1).Tag
        Case 10
            Text1(10).Text = Format(vFecha, "dd/mm/yyyy") '<===
        Case 18
            Text1(18).Text = Format(vFecha, "dd/mm/yyyy") '<===
        Case 7
            Text3(7).Text = Format(vFecha, "dd/mm/yyyy") '<===
    End Select
    ' ********************************************
End Sub
' *****************************************************

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    Screen.MousePointer = vbHourglass
    frmListArticulos.Show vbModal
    Screen.MousePointer = vbDefault
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


Private Sub Text3_LostFocus(Index As Integer)
Dim cadMen As String

    
     If Screen.ActiveForm.ActiveControl.Name = "cmdCancelar" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 1 'Codigo Almacen
             txtAux2(0).Text = PonerNombreDeCod(Text3(Index), "salmpr", "nomalmac")
             If txtAux2(0).Text = "" Then
                cadMen = "No existe el Almac�n: " & Text3(Index).Text & vbCrLf
                cadMen = cadMen & "�Desea crearlo?" & vbCrLf
                If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                    Set frmA = New frmManAlmProp
            
                    frmA.DatosADevolverBusqueda = "0|1|"
                    frmA.NuevoCodigo = Text3(Index).Text
                    Text3(Index).Text = ""
                    TerminaBloquear
                    frmA.Show vbModal
                    Set frmA = Nothing
                    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                Else
                    Text3(Index).Text = ""
                End If
                PonerFoco Text3(Index)
             End If
             
        Case 2, 3, 4, 5, 6 'Stocks, Punto Pedido
                'Formato tipo 1: Decimal(12,2)
                If Trim(Text3(Index)) <> "" Then PonerFormatoDecimal Text3(Index), 1
        
        Case 7  'Fecha Inventario
            If Text3(Index).Text <> "" Then PonerFormatoFecha Text3(Index)

        Case 9  'Hora Inventario
            If Trim(Text3(Index).Text) <> "" Then PonerFormatoHora Text3(Index)
    End Select
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
'            printNou
        Case 13    'Eixir
            mnSalir_Click
            
        Case btnPrimero To btnPrimero + 3 'Fleches Despla�ament
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub BotonBuscar()
Dim i As Integer
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

    CadB = ObtenerBusqueda2(Me, , 1)
    
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
    cad = cad & ParaGrid(Text1(0), 25, "C�digo")
    cad = cad & ParaGrid(Text1(1), 50, "Nombre")
    cad = cad & ParaGrid(Text1(2), 25, "EAN")
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = NombreTabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        frmB.vDevuelve = "0|1|2|" '*** els camps que volen que torne ***
        frmB.vTitulo = "Articulos" ' ***** repasa a��: t�tol de BuscaGrid *****
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
Dim i As Integer
Dim J As Integer

    If Data1.Recordset.EOF Then
        MsgBox "Ning�n registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    cad = ""
    i = 0
    Do
        J = i + 1
        i = InStr(J, DatosADevolverBusqueda, "|")
        If i > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, i - J)
            J = Val(Aux)
            cad = cad & Text1(J).Text & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(cad)
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
'    Text1(0).Text = SugerirCodigoSiguienteStr("sartic", "codartic")
'    FormateaCampo Text1(0)
       
    PonerFoco Text1(0) '*** 1r camp visible que siga PK ***
    
    ' *** si n'hi han camps de descripci� a la cap�alera ***
    'PosarDescripcions
    Text1(17).Text = 0

    ' *** si n'hi han tabs, em posicione al 1r ***
    Me.SSTab1.Tab = 0
End Sub

Private Sub BotonModificar()

    PonerModo 4

    ' *** bloquejar els camps visibles de la clau primaria de la cap�alera ***
    BloquearTxt Text1(0), True
    
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco Text1(1)
End Sub

Private Sub BotonEliminar()
Dim cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    If Not SepuedeBorrarArticulo Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************

    ' *************** canviar la pregunta ****************
    cad = "�Seguro que desea eliminar el Articulo?"
    cad = cad & vbCrLf & "C�digo: " & Format(Data1.Recordset.Fields(0), FormatoCampo(Text1(0)))
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
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Articulo", Err.Description
End Sub

Private Sub PonerCampos()
Dim i As Integer
Dim codPobla As String, desPobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la cap�alera
    
    ' *** si n'hi han ll�nies en datagrids ***
    'For i = 0 To DataGridAux.Count - 1
    For i = 0 To 0
            CargaGrid i, True
            If Not AdoAux(i).Recordset.EOF Then _
                PonerCamposForma2 Me, AdoAux(i), 2, "FrameAux" & i
    Next i

    
    ' ************* configurar els camps de les descripcions de la cap�alera *************
    Text2(2).Text = PonerNombreDeCod(Text1(2), "proveedor", "nomprove")
    Text2(3).Text = PonerNombreDeCod(Text1(3), "sfamia", "nomfamia")
    Text2(5).Text = PonerNombreDeCod(Text1(5), "sunida", "nomunida")
    Text2(6).Text = PonerNombreDeCod(Text1(6), "stipar", "nomtipar")
    If vParamAplic.NumeroConta <> 0 Then
        Text2(7).Text = PonerNombreDeCod(Text1(7), "tiposiva", "nombriva", , , cConta)
    Else
        Text2(7).Text = PonerNombreDeCod(Text1(7), "tiposiva", "nombriva", , , cAgro)
    End If
    
    PonerSumaStocks 'Poner la suma total de stocks de los almacenes donde esta el artic

    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
End Sub

Private Sub PonerSumaStocks()
Dim rst As ADODB.Recordset
Dim Sql As String
    
    Sql = DevuelveDesdeBDNew(cAgro, "salmac", "codartic", "codartic", Text1(0).Text, "T")
    If Sql <> "" Then
        Sql = "select sum(canstock) from salmac where codartic=" & DBSet(Text1(0).Text, "T")
        Set rst = New ADODB.Recordset
        rst.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not rst.EOF Then
            Me.txtSumaStock.Text = rst.Fields(0).Value
        End If
        rst.Close
        Set rst = Nothing
    Else
        Me.txtSumaStock.Text = 0
    End If
End Sub

Private Sub cmdCancelar_Click()
Dim i As Integer
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
                        LLamaLineas NumTabMto, ModoLineas 'ocultar Text3
                        DataGridAux(NumTabMto).Enabled = True
                        DataGridAux(NumTabMto).SetFocus

                        ' *** si n'hi han camps de descripci� dins del grid, els neteje ***
                        'Text32(2).text = ""

                    End If
                    
                    ' *** si n'hi han tabs ***
                    SituarTab (NumTabMto + 1)
                    
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        AdoAux(NumTabMto).Recordset.MoveFirst
                    End If

                Case 2 'modificar ll�nies
                    ModoLineas = 0
                    
                    ' *** si n'hi han tabs ***
                    SituarTab (NumTabMto + 1)
                    LLamaLineas NumTabMto, ModoLineas 'ocultar Text3
                    PonerModo 4
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        ' *** l'Index de Fields es el que canvie de la PK de ll�nies ***
                        V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el n� de llinia
                        AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
                        ' ***************************************************************
                    End If

            End Select
            
            PosicionarData Data1.Recordset!codArtic
             
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
'Dim Datos As String
Dim Mens As String


    On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm2(Me, 1)
    If Not b Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
    If (Modo = 3) Then 'insertar
        'comprobar si existe ya el cod. del campo clave primaria
        If ExisteCP(Text1(0)) Then b = False
    End If
    If Not b Then Exit Function
    ' ************************************************************************************
    
    ' comprobamos que si hay contabilidad las cuentas contables existan
'    If Modo = 3 Or Modo = 4 Then
'        If vParamAplic.NumeroConta <> 0 Then
'                 Text2(4).Text = PonerNombreCuenta(Text1(4), Modo)
'                 Text2(17).Text = PonerNombreCuenta(Text1(17), Modo)
                 
                 
'            If text1(4).Text <> "" Then
'                text2(4).Text = PonerNombreCuenta(text1(4), Modo)
'                If text2(4).Text = "" Then b = False
'            Else
'                MsgBox "Debe poner una Cuenta Contable Socio existente. Reintroduzca.", vbExclamation
'                b = False
'            End If
'            If text1(17).Text <> "" Then
'                text2(17).Text = PonerNombreCuenta(text1(17), Modo)
'                If text2(17).Text = "" Then b = False
'            Else
'                MsgBox "Debe poner una Cuenta Contable Cliente existente. Reintroduzca.", vbExclamation
'                b = False
'            End If
'        End If
'    End If
    
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Sub PosicionarData(codigo As String)
Dim cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la cap�alera, no llevar els () ***
    cad = "(codartic='" & Trim(codigo) & "')" 'DBSet(Text1(0).Text, "T") & ")"
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(Data1, cad, Indicador) Then
    If SituarData(Data1, cad, Indicador, False) Then
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
    ' ***** canviar el nom de la PK de la cap�alera, repasar codEmpre *******
    vWhere = " WHERE codartic=" & DBSet(Data1.Recordset!codArtic, "T")
        
    ' ***** elimina les ll�nies ****
    conn.Execute "DELETE FROM salmac " & vWhere
        
        
    'Eliminar la CAP�ALERA
    vWhere = " WHERE codartic=" & DBSet(Data1.Recordset!codArtic, "T")
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



    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    ' ***************** configurar els LostFocus dels camps de la cap�alera *****************

    
    'Si queremos hacer algo ..
    Select Case Index
        Case 0 'Codigo Art�culo
            'Comprobar si ya existe el cod de articulo en la tabla
            If Modo = 3 Then 'Insertar
                If ExisteCP(Text1(Index)) Then PonerFoco Text1(Index)
            End If

        Case 2 'Codigo de Proveedor
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "proveedor", "nomprove")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Proveedor: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "�Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmPro = New frmManProve
                        frmPro.DatosADevolverBusqueda = "0|1|"
                        frmPro.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmPro.Show vbModal
                        Set frmPro = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 3 'C�digo de Familia
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "sfamia", "nomfamia")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Familia: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "�Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmFam = New frmManFamilias
                        frmFam.DatosADevolverBusqueda = "0|1|"
                        frmFam.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmFam.Show vbModal
                        Set frmFam = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
            
        Case 5 'C�digo Tipo Unidad
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "sunida", "nomunida")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Tipo de Unidad: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "�Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmTUn = New frmManTipUnid
                        frmTUn.DatosADevolverBusqueda = "0|1|"
                        frmTUn.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmTUn.Show vbModal
                        Set frmTUn = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 6 'Codigo Tipo Art�culo
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "stipar", "nomtipar")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Tipo de Art�culo: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "�Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmTAr = New frmManTipArtic
                        frmTAr.DatosADevolverBusqueda = "0|1|"
                        frmTAr.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmTAr.Show vbModal
                        Set frmTAr = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 7 'Tipo de IVA
            'conConta: BD Contabilidad
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "tiposiva", "nombriva", , , cConta)
            Else
                Text2(Index).Text = ""
            End If
            
        Case 10, 18 'Fecha alta, Fecha �ltima compra
            PonerFormatoFecha Text1(Index)

'        Case 11, 12 'numericos
'            PonerFormatoEntero Text1(Index)
'
        Case 13, 15, 17 'Precios
            'Formato tipo 2: Decimal(10,4)
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 7
           
    End Select



End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 2: KEYBusqueda KeyAscii, 0 'proveedor
                Case 3: KEYBusqueda KeyAscii, 1 'familia
                Case 5: KEYBusqueda KeyAscii, 3 'tipo de unidad
                Case 6: KEYBusqueda KeyAscii, 4 'tipo de articulo
                Case 7: KEYBusqueda KeyAscii, 2 'tipo de iva
                Case 10: KEYFecha KeyAscii, 0 'fecha de alta
                Case 18: KEYFecha KeyAscii, 1 'fecha de ultima compra
            End Select
        End If
    Else
        If (Index <> 19 Or (Index = 19 And Text1(19).Text = "")) And _
           (Index <> 20 Or (Index = 20 And Text1(20).Text = "")) Then KEYpress KeyAscii
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

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub KEYBusquedaLin(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
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
    Eliminar = False
   
    vWhere = ObtenerWhereCab(True)
    
    ' ***** independentment de si tenen datagrid o no,
    ' canviar els noms, els formats i el DELETE *****
    Select Case Index
        Case 0 'articulo en almacen
            Sql = "Seguro que desea eliminar de la BD el registro:"
            Sql = Sql & vbCrLf & "Cod. Art�culo: " & AdoAux(Index).Recordset.Fields(0)
            Sql = Sql & vbCrLf & "Cod. Almacen: " & AdoAux(Index).Recordset.Fields(1)

            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM salmac"
                Sql = Sql & vWhere & " AND codalmac= " & AdoAux(Index).Recordset!codAlmac
            End If
    End Select

    If Eliminar Then
        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        conn.Execute Sql
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        CargaGrid Index, True
        If Not SituarDataTrasEliminar(AdoAux(Index), NumRegElim, True) Then
            PonerCampos
            
        End If
        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
        ' *** si n'hi han tabs ***
        SituarTab (NumTabMto + 1)
    End If
    
    ModoLineas = 0
    PosicionarData Data1.Recordset!codArtic
    
    Exit Sub
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando linea", Err.Description
End Sub

Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vTabla As String
Dim anc As Single
Dim i As Integer
    
    ModoLineas = 1 'Posem Modo Afegir Ll�nia
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Cap�alera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index
    
    ' *** bloquejar la clau primaria de la cap�alera ***
    BloquearTxt Text1(0), True
    BloquearTxt Text3(1), False
    BloquearBtn btnBuscar(0), False

    ' *** posar el nom del les distintes taules de ll�nies ***
    Select Case Index
        Case 0: vTabla = "salmac"
    End Select
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 0 ' *** pose els index dels tabs de ll�nies que tenen datagrid ***
            ' *** canviar la clau primaria de les ll�nies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***1
'            If Index = 1 Then
'                NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", vWhere)
'            End If

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
                Case 0 'stocks en almacenes
                    Text3(0).Text = Text1(0).Text 'codartic
                    Text3(1).Text = ""
                    For i = 1 To Text3.Count - 1
                        Text3(i).Text = ""
                    Next i
                    txtAux2(0).Text = ""
                    PonerFoco Text3(1)
            End Select
    End Select
End Sub

Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim i As Integer
    Dim J As Integer
    
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
    BloquearTxt Text3(0), True
    BloquearTxt Text3(1), True
    BloquearTxt txtAux2(0), True
    BloquearBtn btnBuscar(0), True
    Select Case Index
        Case 0 ' *** pose els index de ll�nies que tenen datagrid (en o sense tab) ***
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
        Case 0 'stocks en almacenes
        
            For J = 0 To 1
                Text3(J).Text = DataGridAux(Index).Columns(J).Text
            Next J
            txtAux2(0).Text = DataGridAux(Index).Columns(2).Text
            For J = 3 To 8
                Text3(J - 1).Text = DataGridAux(Index).Columns(J).Text
            Next J
            
            Text3(9).Text = Mid(Text3(8).Text, 12, 8)
            
            For i = 0 To 0
                BloquearTxt Text3(i), False
            Next i
            
            PosicionarCombo cmbAux(0), AdoAux(Index).Recordset!statusin
            
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 0 'stocks en almacenes
            PonerFoco Text3(2)
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
        Case 0 'stocks
             For jj = 1 To 9
                Text3(jj).visible = b
                Text3(jj).Top = alto
             Next jj
             Text3(9).Left = Text3(8).Left
             txtAux2(0).visible = b
             txtAux2(0).Top = alto
             For jj = 0 To 1
                btnBuscar(jj).visible = b
                btnBuscar(jj).Top = alto
             Next jj
             Me.cmbAux(0).visible = b
             Me.cmbAux(0).Top = alto
            
    End Select
End Sub

' ********* si n'hi han combos a la cap�alera ************
Private Sub CargaCombo(Index As Integer)
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    cmbAux(0).Clear

    cmbAux(0).AddItem "No"
    cmbAux(0).ItemData(cmbAux(0).NewIndex) = 0
    cmbAux(0).AddItem "S�"
    cmbAux(0).ItemData(cmbAux(0).NewIndex) = 1

End Sub


Private Sub Text3_GotFocus(Index As Integer)
   If Not Text3(Index).MultiLine Then ConseguirFocoLin Text3(Index)
End Sub

Private Sub Text3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not Text3(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not Text3(Index).MultiLine Then
        If KeyAscii = teclaBuscar Then
            If Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) Then
            End If
        Else
            KEYpress KeyAscii
        End If
    End If
End Sub

Private Function DatosOkLlin(nomframe As String) As Boolean
Dim RS As ADODB.Recordset
Dim Sql As String
Dim b As Boolean
Dim Cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte
Dim devuelve As String

    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False
        
    b = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not b Then Exit Function
    
    'Campo de cantidad de Stock (Son decimales)
    If Trim(Text3(2).Text) = "" Or IsNull(Text3(2).Text) Then
        MsgBox "El campo Cantidad Stock no puede ser nulo", vbExclamation, "Art�culos"
        b = False
    End If
    If Not b Then Exit Function
    
    ' ******************************************************************************
    'Comprobamos  si existe
    devuelve = DevuelveDesdeBDNew(cAgro, "salmac", "codartic", "codartic", Text1(0).Text, "T", , "codalmac", Text3(1).Text, "N")
    If ModoLineas = 1 And devuelve <> "" Then
        b = False
        devuelve = "Ya existe el Art�culo en el Almacen: " & vbCrLf
        devuelve = devuelve & "Codigo: " & Text3(0).Text & vbCrLf
        devuelve = devuelve & "Descripci�n: " & txtAux2(0).Text
        MsgBox devuelve, vbExclamation, "Art�culos"
    End If
    DatosOkLlin = b
    
EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function SepuedeBorrar(ByRef Index As Integer) As Boolean
    SepuedeBorrar = False
    
    ' *** si cal comprovar alguna cosa abans de borrar ***
    Select Case Index
        Case 0 'cuentas bancarias
            If AdoAux(Index).Recordset!ctaprpal = 1 Then
                MsgBox "No puede borrar una Cuenta Principal. Seleccione antes otra cuenta como Principal", vbExclamation
                Exit Function
            End If
    End Select
    ' ****************************************************
    
    SepuedeBorrar = True
End Function

Private Function SepuedeBorrarArticulo() As Boolean
Dim Sql As String

    SepuedeBorrarArticulo = False
    
    ' *** si cal comprovar alguna cosa abans de borrar ***
    Sql = "select count(*) from forfaits_envases where codartic = " & DBSet(Text1(0).Text, "T")
    If TotalRegistros(Sql) <> 0 Then
        MsgBox "Este art�culo est� en una confecci�n, no se puede eliminar. Revise", vbExclamation
        Exit Function
    End If
    ' ****************************************************
    
    SepuedeBorrarArticulo = True
End Function



Private Sub imgBuscar_Click(Index As Integer)
    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Codigo Proveedor
            Set frmPro = New frmManProve
            frmPro.DatosADevolverBusqueda = "0|1|"
            frmPro.Show vbModal
            Set frmPro = Nothing
        Case 1  'Cod. Familia
            Set frmFam = New frmManFamilias
            frmFam.DatosADevolverBusqueda = "0|1|"
            frmFam.Show vbModal
            Set frmFam = Nothing
        Case 3  'Cod. Tipo Unidad
            Set frmTUn = New frmManTipUnid
            frmTUn.DatosADevolverBusqueda = "0|1|"
            frmTUn.Show vbModal
            Set frmTUn = Nothing
        Case 4  'Cod. Tipo Articulo
            Set frmTAr = New frmManTipArtic
            frmTAr.DatosADevolverBusqueda = "0|1|"
            frmTAr.Show vbModal
            Set frmTAr = Nothing
            
        Case 2  'Tipos de IVA. Tabla de la BD Contabilidad
            Set frmTipIva = New frmTipIVAConta
            frmTipIva.DeConsulta = True
            frmTipIva.DatosADevolverBusqueda = "0|1|2|"
            frmTipIva.CodigoActual = Text1(5).Text
            frmTipIva.Show vbModal
            Set frmTipIva = Nothing
    End Select
    PonerFoco Text1(Index + 2)
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'des macta
End Sub

Private Sub frmFam_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Colectivos
    Text1(3).Text = RecuperaValor(CadenaSeleccion, 1) 'codfamia
    FormateaCampo Text1(3)
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'nomfamia
End Sub

Private Sub DataGridAux_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
Dim i As Byte

    If ModoLineas <> 1 Then
        Select Case Index
            Case 0 'tarifas
                If DataGridAux(Index).Columns.Count > 2 Then
'                    Text3(11).Text = DataGridAux(Index).Columns("direccio").Text
'                    Text3(12).Text = DataGridAux(Index).Columns("observac").Text
                End If
                
            Case 1 'bonificaiones
                If DataGridAux(Index).Columns.Count > 2 Then
'                    Text3(21).Text = DataGridAux(Index).Columns(5).Text
'                    Text3(22).Text = DataGridAux(Index).Columns(6).Text
'                    Text3(23).Text = DataGridAux(Index).Columns(8).Text
'                    Text3(24).Text = DataGridAux(Index).Columns(15).Text
'                    Text32(22).Text = DataGridAux(Index).Columns(7).Text
                End If
                
        End Select
        
    Else 'vamos a Insertar
        Select Case Index
            Case 0 'cuentas bancarias
'                Text3(11).Text = ""
'                Text3(12).Text = ""
            Case 1 'departamentos
                For i = 21 To 24
'                   Text3(i).Text = ""
                Next i
'               Text32(22).Text = ""
            Case 2 'Tarjetas
'               Text3(50).Text = ""
'               Text3(51).Text = ""
        End Select
    End If
End Sub

' ***** si n'hi han varios nivells de tabs *****
Private Sub SituarTab(numTab As Integer)
    On Error Resume Next
    
    SSTab1.Tab = numTab
    
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
Dim i As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)

    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
    
    Select Case Index
        Case 0 'stocks en almacenes
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;" 'codartic
            tots = tots & "S|Text3(1)|T|Alm.|650|;" 'almacen
            tots = tots & "S|btnBuscar(0)|B|||;S|txtAux2(0)|T|Denominaci�n|1700|;"
            tots = tots & "S|Text3(2)|T|Cant.Stock|1000|;"
            tots = tots & "S|Text3(3)|T|Stock Min.|1000|;"
            tots = tots & "S|Text3(4)|T|Punto Ped.|1000|;"
            tots = tots & "S|Text3(5)|T|Stock Max.|1000|;"
            tots = tots & "S|Text3(6)|T|Stock Inv.|1000|;"
            tots = tots & "S|Text3(7)|T|Fecha Inv.|1200|;"
            tots = tots & "S|btnBuscar(1)|B|||;"
            tots = tots & "S|Text3(8)|T|Hora Inv.|1000|;"
            tots = tots & "N||||0|;" 'inventariandose
            tots = tots & "S|cmbAux(0)|C|Inv.|600|;"
            
            Text3(8).Tag = "Hora Inventario|FH|S|||salmac|horainve|hh:mm:ss|N|"
            arregla tots, DataGridAux(Index), Me
            Text3(8).Tag = "Hora Inventario|FH|S|||salmac|horainve|yyyy-mm-dd hh:mm:ss|N|"
        
            DataGridAux(0).Columns(1).Alignment = dbgLeft
'            DataGridAux(0).Columns(6).Alignment = dbgRight
'            DataGridAux(0).Columns(7).Alignment = dbgRight
'            DataGridAux(0).Columns(8).Alignment = dbgRight
        
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
    End Select
    
    DataGridAux(Index).ScrollBars = dbgAutomatic
      
    ' **** si n'hi han ll�nies en grids i camps fora d'estos ****
    If Not AdoAux(Index).Recordset.EOF Then
        DataGridAux_RowColChange Index, 1, 1
    Else
'        LimpiarCamposFrame Index
    End If
      
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub

Private Sub InsertarLinea()
'Inserta registre en les taules de Ll�nies
Dim nomframe As String
Dim b As Boolean

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'stocks en almacenes
    End Select
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        
        If Text3(9).Text = "" Then Text3(9).Text = "00:00:00"
        Text3(8).Text = Format(Text3(7).Text, FormatoFecha) & " " & Text3(9).Text

        If InsertarDesdeForm2(Me, 2, nomframe) Then
            b = BLOQUEADesdeFormulario2(Me, Data1, 1)
            Select Case NumTabMto
                Case 0 ' *** els index de les llinies en grid (en o sense tab) ***
                    CargaGrid NumTabMto, True
                    If b Then BotonAnyadirLinea NumTabMto
            End Select
           
            SituarTab (NumTabMto + 1)
        End If
    End If
End Sub

Private Sub ModificarLinea()
'Modifica registre en les taules de Ll�nies
Dim nomframe As String
Dim V As Integer
    
    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'stocks en almacenes
    End Select
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        
        If Text3(9).Text = "" Then Text3(9).Text = "00:00:00"
        Text3(8).Text = Format(Text3(7).Text, FormatoFecha) & " " & Text3(9).Text
        
        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
            ModoLineas = 0
            
            V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el n� de llinia
            CargaGrid NumTabMto, True
            
            ' *** si n'hi han tabs ***
            SituarTab (NumTabMto + 1)

            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
            PonerFocoGrid Me.DataGridAux(NumTabMto)
            AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
            
            LLamaLineas NumTabMto, 0
        End If
    End If
End Sub

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la cap�alera ***
    vWhere = vWhere & " codartic='" & Trim(Text1(0).Text) & "'"
    
    ObtenerWhereCab = vWhere
End Function

'' *** neteja els camps dels tabs de grid que
''estan fora d'este, i els camps de descripci� ***
Private Sub LimpiarCamposFrame(Index As Integer)
    On Error Resume Next
 
'    Select Case Index
'        Case 0 'Cuentas Bancarias
'            Text3(11).Text = ""
'            Text3(12).Text = ""
'        Case 1 'Departamentos
'            Text3(21).Text = ""
'            Text3(22).Text = ""
'            Text32(22).Text = ""
'            Text3(23).Text = ""
'            Text3(24).Text = ""
'        Case 2 'Tarjetas
'            Text3(50).Text = ""
'            Text3(51).Text = ""
'        Case 4 'comisiones
'            Text32(2).Text = ""
'    End Select
'
    If Err.Number <> 0 Then Err.Clear
End Sub

' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del rat�n.
'Private Sub DataGridAux_GotFocus(Index As Integer)
'  WheelHook DataGridAux(Index)
'End Sub
'Private Sub DataGridAux_LostFocus(Index As Integer)
'  WheelUnHook
'End Sub

Private Sub InsetarArticulosPorAlmacen()
'Inserta en la tabla salmac una fila del art�culo que se esta insertando
'para cada uno de los almacenes que existen en la tabla salmpr
Dim vCodArtic As String, vcodalmac As Integer
Dim rsAlmPr As ADODB.Recordset
Dim cad As String
    
    On Error GoTo EInsEnAlm

    vCodArtic = Text1(0).Text
    Set rsAlmPr = New ADODB.Recordset
    cad = "Select codalmac from salmpr order by codalmac;"
    rsAlmPr.Open cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    While Not rsAlmPr.EOF
        vcodalmac = rsAlmPr.Fields(0).Value
        cad = "INSERT INTO salmac (codartic,codalmac,canstock,stockmin,puntoped,stockmax,stockinv,fechainv,horainve,statusin)"
        cad = cad & " VALUES (" & DBSet(vCodArtic, "T") & "," & vcodalmac & ",0,0,0,0,0,NULL,NULL,0)"
        conn.Execute cad
        rsAlmPr.MoveNext
    Wend
        
    rsAlmPr.Close
    Set rsAlmPr = Nothing
EInsEnAlm:
    If Err.Number <> 0 Then MuestraError Err.Number, "Insertando Art�culo en Almacenes.", Err.Description
End Sub

