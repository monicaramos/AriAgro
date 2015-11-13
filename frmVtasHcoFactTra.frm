VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmVtasHcoFactTra 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Histórico de Facturas Transportistas/Comisionistas"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   11745
   Icon            =   "frmVtasHcoFactTra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmVtasHcoFactTra.frx":000C
   ScaleHeight     =   7155
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   45
      TabIndex        =   42
      Tag             =   "Fecha Oferta|F|N|||scapre|fecentre|dd/mm/yyyy|N|"
      Top             =   1485
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   8705
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos básicos"
      TabPicture(0)   =   "frmVtasHcoFactTra.frx":0A0E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2(1)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Albaranes"
      TabPicture(1)   =   "frmVtasHcoFactTra.frx":0A2A
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(6)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "FrameObserva"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "DataGrid2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "DataGrid1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtAux(7)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtAux(6)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtAux(5)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtAux(4)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdObserva"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtAux(0)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtAux(1)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtAux(2)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtAux(3)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtAux3(0)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtAux3(1)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Text3(2)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "Portes Vuelta"
      TabPicture(2)   =   "frmVtasHcoFactTra.frx":0A46
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1(13)"
      Tab(2).Control(1)=   "DataGrid3"
      Tab(2).Control(2)=   "Text4(0)"
      Tab(2).Control(3)=   "Text4(1)"
      Tab(2).Control(4)=   "Text4(2)"
      Tab(2).Control(5)=   "Text5"
      Tab(2).ControlCount=   6
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   -69600
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   116
         Text            =   "Text5"
         Top             =   3015
         Width           =   1800
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   2
         Left            =   -73065
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   112
         Text            =   "Text4"
         Top             =   2700
         Visible         =   0   'False
         Width           =   5085
      End
      Begin VB.TextBox Text4 
         Height          =   330
         Index           =   1
         Left            =   -67890
         TabIndex        =   111
         Tag             =   "Importe|N|N|||tcafpv|importel|###,###,##0.00|N|"
         Text            =   "Text4"
         Top             =   2700
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.TextBox Text4 
         Height          =   330
         Index           =   0
         Left            =   -74460
         TabIndex        =   110
         Tag             =   "Cuenta Contable|T|N|||tcafpv|codmacta||N|"
         Text            =   "Text4"
         Top             =   2700
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Frame Frame2 
         Height          =   4560
         Index           =   1
         Left            =   -74865
         TabIndex        =   82
         Top             =   315
         Width           =   11175
         Begin VB.Frame FrameCliente 
            Caption         =   "Datos Agencia de Transporte/Comisionista"
            ForeColor       =   &H00972E0B&
            Height          =   1740
            Left            =   45
            TabIndex        =   99
            Top             =   135
            Width           =   11055
            Begin VB.TextBox Text1 
               Height          =   315
               Index           =   6
               Left            =   1125
               MaxLength       =   35
               TabIndex        =   8
               Tag             =   "Domicilio|T|N|||tcafpc|domtrans||N|"
               Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
               Top             =   645
               Width           =   4030
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Index           =   12
               Left            =   8925
               MaxLength       =   5
               TabIndex        =   14
               Tag             =   "Descuento General|N|N|0|99.90|tcafpc|dtognral|#0.00|N|"
               Text            =   "Text1 7"
               Top             =   1110
               Width           =   525
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Index           =   11
               Left            =   7560
               MaxLength       =   5
               TabIndex        =   13
               Tag             =   "Descuento P.Pago|N|N|0|99.90|tcafpc|dtoppago|#0.00|N|"
               Text            =   "Text1 7"
               Top             =   1110
               Width           =   525
            End
            Begin VB.TextBox Text2 
               BackColor       =   &H80000018&
               Height          =   315
               Index           =   10
               Left            =   7530
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   100
               Text            =   "Text2"
               Top             =   645
               Width           =   3285
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Index           =   10
               Left            =   6930
               MaxLength       =   3
               TabIndex        =   12
               Tag             =   "Forma de Pago|N|N|0|999|tcafpc|codforpa|000|N|"
               Text            =   "Text1"
               Top             =   645
               Width           =   540
            End
            Begin VB.TextBox Text1 
               Height          =   315
               Index           =   4
               Left            =   1125
               MaxLength       =   15
               TabIndex        =   6
               Tag             =   "NIF transportista|T|N|||tcafpc|niftrans||N|"
               Text            =   "123456789"
               Top             =   285
               Width           =   1110
            End
            Begin VB.TextBox Text1 
               Height          =   315
               Index           =   5
               Left            =   3195
               MaxLength       =   20
               TabIndex        =   7
               Tag             =   "teléfono transportista|T|S|||tcafpc|teltrans||N|"
               Text            =   "12345678911234567899"
               Top             =   285
               Width           =   1965
            End
            Begin VB.TextBox Text1 
               Height          =   315
               Index           =   8
               Left            =   1755
               MaxLength       =   30
               TabIndex        =   10
               Tag             =   "Población|T|N|||tcafpc|pobtrans||N|"
               Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
               Top             =   990
               Width           =   3405
            End
            Begin VB.TextBox Text1 
               Height          =   315
               Index           =   7
               Left            =   1125
               MaxLength       =   6
               TabIndex        =   9
               Tag             =   "CPostal|T|N|||tcafpc|codpobla||N|"
               Text            =   "Text15"
               Top             =   960
               Width           =   630
            End
            Begin VB.TextBox Text1 
               Height          =   315
               Index           =   9
               Left            =   1125
               MaxLength       =   30
               TabIndex        =   11
               Tag             =   "Provincia|T|N|||tcafpc|protrans||N|"
               Text            =   "Text1 Text1 Text1 Text1 Text22"
               Top             =   1350
               Width           =   2445
            End
            Begin VB.Label Label1 
               Caption         =   "Domicilio"
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   108
               Top             =   645
               Width           =   735
            End
            Begin VB.Image imgBuscar 
               Height          =   240
               Index           =   3
               Left            =   6660
               ToolTipText     =   "Buscar forma de pago"
               Top             =   675
               Width           =   240
            End
            Begin VB.Label Label1 
               Caption         =   "Dto. Gral"
               Height          =   255
               Index           =   26
               Left            =   8235
               TabIndex        =   107
               Top             =   1110
               Width           =   735
            End
            Begin VB.Label Label1 
               Caption         =   "Dto. P.P"
               Height          =   255
               Index           =   25
               Left            =   6900
               TabIndex        =   106
               Top             =   1110
               Width           =   615
            End
            Begin VB.Label Label1 
               Caption         =   "Forma Pago"
               Height          =   255
               Index           =   15
               Left            =   5730
               TabIndex        =   105
               Top             =   645
               Width           =   855
            End
            Begin VB.Image imgBuscar 
               Height          =   240
               Index           =   1
               Left            =   855
               ToolTipText     =   "Buscar proveedor varios"
               Top             =   300
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.Label Label1 
               Caption         =   "NIF"
               Height          =   255
               Index           =   20
               Left            =   120
               TabIndex        =   104
               Top             =   285
               Width           =   615
            End
            Begin VB.Label Label1 
               Caption         =   "Teléfono"
               Height          =   255
               Index           =   19
               Left            =   2445
               TabIndex        =   103
               Top             =   285
               Width           =   735
            End
            Begin VB.Label Label1 
               Caption         =   "Población"
               Height          =   255
               Index           =   16
               Left            =   120
               TabIndex        =   102
               Top             =   990
               Width           =   735
            End
            Begin VB.Label Label1 
               Caption         =   "Provincia"
               Height          =   255
               Index           =   17
               Left            =   120
               TabIndex        =   101
               Top             =   1350
               Width           =   735
            End
         End
         Begin VB.Frame FrameFactura 
            Height          =   2670
            Left            =   45
            TabIndex        =   83
            Top             =   1845
            Width           =   11055
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   34
               Left            =   7560
               MaxLength       =   15
               TabIndex        =   32
               Tag             =   "Importe Retencion|N|S|||tcafpc|trefacpr|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   2295
               Width           =   1485
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   33
               Left            =   5040
               MaxLength       =   5
               TabIndex        =   30
               Tag             =   "% Ret|N|S|0|99.90|tcafpc|retfacpr|#0.00|N|"
               Text            =   "Text1 7"
               Top             =   2295
               Width           =   525
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   32
               Left            =   5760
               MaxLength       =   15
               TabIndex        =   31
               Tag             =   "Base Imponible 3|N|S|||tcafpc|baseiva3|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   2295
               Width           =   1485
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   14
               Left            =   240
               MaxLength       =   15
               TabIndex        =   15
               Tag             =   "Imp.Bruto|N|N|||tcafpc|brutofac|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   435
               Width           =   1485
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   15
               Left            =   2160
               MaxLength       =   15
               TabIndex        =   16
               Tag             =   "Imp. Dto PP|N|N|||tcafpc|impppago|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   435
               Width           =   1365
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   16
               Left            =   3960
               MaxLength       =   15
               TabIndex        =   17
               Tag             =   "Imp. Dto Gn|N|N|||tcafpc|impgnral|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   450
               Width           =   1365
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   17
               Left            =   5760
               MaxLength       =   15
               TabIndex        =   84
               Text            =   "Text1 7"
               Top             =   450
               Width           =   1485
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   24
               Left            =   5760
               MaxLength       =   15
               TabIndex        =   20
               Tag             =   "Base Imponible 1|N|N|||tcafpc|baseiva1|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   1080
               Width           =   1485
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   18
               Left            =   4320
               MaxLength       =   3
               TabIndex        =   18
               Tag             =   "Cod. IVA 1|N|S|0|999|tcafpc|tipoiva1|000|N|"
               Text            =   "Text1 7"
               Top             =   1080
               Width           =   525
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   21
               Left            =   5040
               MaxLength       =   5
               TabIndex        =   19
               Tag             =   "% IVA 1|N|S|0|99.90|tcafpc|porciva1|#0.00|N|"
               Text            =   "Text1 7"
               Top             =   1080
               Width           =   525
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   27
               Left            =   7560
               MaxLength       =   15
               TabIndex        =   21
               Tag             =   "Importe IVA 1|N|N|||tcafpc|impoiva1|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   1080
               Width           =   1485
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   25
               Left            =   5760
               MaxLength       =   15
               TabIndex        =   24
               Tag             =   "Base Imponible 2 |N|S|||tcafpc|baseiva2|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   1395
               Width           =   1485
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   19
               Left            =   4320
               MaxLength       =   3
               TabIndex        =   22
               Tag             =   "Cod. IVA 2|N|S|0|999|tcafpc|tipoiva2|000|N|"
               Text            =   "Text1 7"
               Top             =   1395
               Width           =   525
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   22
               Left            =   5040
               MaxLength       =   5
               TabIndex        =   23
               Tag             =   "& IVA 2|N|S|0|99.90|tcafpc|porciva2|#0.00|N|"
               Text            =   "Text1 7"
               Top             =   1395
               Width           =   525
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   28
               Left            =   7560
               MaxLength       =   15
               TabIndex        =   25
               Tag             =   "Importe IVA 2|N|S|||tcafpc|impoiva2|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   1395
               Width           =   1485
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   26
               Left            =   5760
               MaxLength       =   15
               TabIndex        =   28
               Tag             =   "Base Imponible 3|N|S|||tcafpc|baseiva3|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   1725
               Width           =   1485
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   20
               Left            =   4320
               MaxLength       =   3
               TabIndex        =   26
               Tag             =   "Cod. IVA 3|N|S|0|999|tcafpc|tipoiva3|000|N|"
               Text            =   "Text1 7"
               Top             =   1725
               Width           =   525
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   23
               Left            =   5040
               MaxLength       =   5
               TabIndex        =   27
               Tag             =   "% IVA 3|N|S|0|99.90|tcafpc|porciva3|#0.00|N|"
               Text            =   "Text1 7"
               Top             =   1725
               Width           =   525
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   29
               Left            =   7560
               MaxLength       =   15
               TabIndex        =   29
               Tag             =   "Importe IVA 3|N|S|||tcafpc|impoiva3|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   1725
               Width           =   1485
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0FF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   30
               Left            =   9360
               MaxLength       =   15
               TabIndex        =   33
               Tag             =   "Total Factura|N|N|||tcafpc|totalfac|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   2310
               Width           =   1605
            End
            Begin VB.Label Label1 
               Caption         =   "Imp. Retención"
               Height          =   255
               Index           =   8
               Left            =   7560
               TabIndex        =   115
               Top             =   2070
               Width           =   1335
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
               Index           =   5
               Left            =   7335
               TabIndex        =   114
               Top             =   2160
               Width           =   135
            End
            Begin VB.Label Label1 
               Caption         =   "% RET"
               Height          =   255
               Index           =   4
               Left            =   5040
               TabIndex        =   113
               Top             =   2070
               Width           =   495
            End
            Begin VB.Label Label1 
               Caption         =   "Imp. Bruto"
               Height          =   255
               Index           =   10
               Left            =   240
               TabIndex        =   98
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "Imp. Dto PP"
               Height          =   255
               Index           =   11
               Left            =   2280
               TabIndex        =   97
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "Imp. Dto Gn"
               Height          =   255
               Index           =   12
               Left            =   4080
               TabIndex        =   96
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "Base Imponible"
               Height          =   255
               Index           =   14
               Left            =   5880
               TabIndex        =   95
               Top             =   240
               Width           =   1215
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
               TabIndex        =   94
               Top             =   360
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
               TabIndex        =   93
               Top             =   360
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
               TabIndex        =   92
               Top             =   360
               Width           =   135
            End
            Begin VB.Label Label1 
               Caption         =   "Imp. IVA"
               Height          =   255
               Index           =   33
               Left            =   7605
               TabIndex        =   91
               Top             =   855
               Width           =   1335
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
               Index           =   37
               Left            =   7320
               TabIndex        =   90
               Top             =   960
               Width           =   135
            End
            Begin VB.Line Line1 
               X1              =   4320
               X2              =   7320
               Y1              =   825
               Y2              =   825
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
               TabIndex        =   89
               Top             =   2160
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
               Index           =   38
               Left            =   9120
               TabIndex        =   88
               Top             =   2310
               Width           =   135
            End
            Begin VB.Label Label1 
               Caption         =   "TOTAL FACTURA"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   255
               Index           =   39
               Left            =   9330
               TabIndex        =   87
               Top             =   2070
               Width           =   1530
            End
            Begin VB.Label Label1 
               Caption         =   "% IVA"
               Height          =   255
               Index           =   41
               Left            =   5040
               TabIndex        =   86
               Top             =   870
               Width           =   495
            End
            Begin VB.Label Label1 
               Caption         =   "Cod. IVA"
               Height          =   255
               Index           =   42
               Left            =   4320
               TabIndex        =   85
               Top             =   840
               Width           =   735
            End
         End
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   9360
         MaxLength       =   14
         TabIndex        =   80
         Tag             =   "Importe|N|S|||tcafpa|importe|###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2025
         Width           =   1350
      End
      Begin VB.TextBox txtAux3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   2070
         MaxLength       =   30
         TabIndex        =   64
         Tag             =   "Fecha Albaran|F|N|||tcafpa|fechaalb|dd/mm/yyyy|N|"
         Text            =   "fecalbar"
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   63
         Tag             =   "Nº Albaran|N|N|||tcafpa|numalbar|0|N|"
         Text            =   "numalbar"
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   3
         Left            =   3240
         MaxLength       =   12
         TabIndex        =   56
         Tag             =   "Forfait|T|N|0||slifac|cantidad||N|"
         Text            =   "nom.forfait"
         Top             =   4320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   2160
         MaxLength       =   12
         TabIndex        =   55
         Tag             =   "Nombre Marca|T|N|||slifac|nommarca||N|"
         Text            =   "nommarca"
         Top             =   4320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   54
         Tag             =   "Variedad comercial|T|N|||slifac|codartic||N|"
         Text            =   "comercial"
         Top             =   4320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   360
         MaxLength       =   12
         TabIndex        =   53
         Tag             =   "Variedad|T|N|||slifac|codalmac||N|"
         Text            =   "variedad"
         Top             =   4320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdObserva 
         Height          =   375
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   520
         Width           =   375
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   4
         Left            =   4080
         MaxLength       =   12
         TabIndex        =   57
         Tag             =   "Precio|N|N|0|999999.0000|slifac|precioar|###,##0.0000|N|"
         Text            =   "Precio"
         Top             =   4320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   5
         Left            =   5760
         MaxLength       =   5
         TabIndex        =   58
         Tag             =   "Dto 1|N|N|0|99.90|slifac|dtoline1|#0.00|N|"
         Text            =   "Dto1"
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   6
         Left            =   6480
         MaxLength       =   30
         TabIndex        =   59
         Tag             =   "Dto 2|N|N|0|99.90|slifac|dtolinea|#0.00|N|"
         Text            =   "Dto2"
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   7
         Left            =   7080
         MaxLength       =   12
         TabIndex        =   62
         Tag             =   "Importe|N|N|0||slifac|importel|#,###,###,##0.00|N|"
         Text            =   "Importe"
         Top             =   4320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmVtasHcoFactTra.frx":0A62
         Height          =   2025
         Left            =   225
         TabIndex        =   44
         Top             =   2625
         Width           =   10935
         _ExtentX        =   19288
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
            AllowFocus      =   0   'False
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmVtasHcoFactTra.frx":0A77
         Height          =   1940
         Left            =   240
         TabIndex        =   45
         Top             =   525
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   3413
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
         Caption         =   "Albaranes de la Factura"
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
      Begin VB.Frame FrameObserva 
         Caption         =   "Observaciones"
         ForeColor       =   &H00972E0B&
         Height          =   2055
         Left            =   225
         TabIndex        =   46
         Tag             =   "Observación 4|T|S|||scafac1|observa4||N|"
         Top             =   2610
         Width           =   10575
         Begin VB.TextBox Text3 
            Height          =   300
            Index           =   8
            Left            =   720
            MaxLength       =   80
            TabIndex        =   51
            Tag             =   "Observación 5|T|S|||tcafpa|observa5||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   1560
            Width           =   8940
         End
         Begin VB.TextBox Text3 
            Height          =   300
            Index           =   7
            Left            =   720
            MaxLength       =   80
            TabIndex        =   50
            Tag             =   "Observación 4|T|S|||tcafpa|observa4||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   1230
            Width           =   8940
         End
         Begin VB.TextBox Text3 
            Height          =   300
            Index           =   6
            Left            =   720
            MaxLength       =   80
            TabIndex        =   49
            Tag             =   "Observación 3|T|S|||tcafpa|observa3||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   900
            Width           =   8940
         End
         Begin VB.TextBox Text3 
            Height          =   300
            Index           =   5
            Left            =   720
            MaxLength       =   80
            TabIndex        =   48
            Tag             =   "Observación 2|T|S|||tcafpa|observa2||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   570
            Width           =   8940
         End
         Begin VB.TextBox Text3 
            Height          =   300
            Index           =   4
            Left            =   720
            MaxLength       =   80
            TabIndex        =   47
            Tag             =   "Observación 1|T|S|||tcafpa|observa1||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   240
            Width           =   8940
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "frmVtasHcoFactTra.frx":0A8C
         Height          =   1995
         Left            =   -74775
         TabIndex        =   109
         Top             =   540
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   3519
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
         Caption         =   "Cuentas Contables"
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
      Begin VB.Label Label1 
         Caption         =   "Importe Total"
         Height          =   255
         Index           =   13
         Left            =   -70680
         TabIndex        =   117
         Top             =   3060
         Width           =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Total"
         Height          =   255
         Index           =   6
         Left            =   9360
         TabIndex        =   81
         Top             =   1755
         Width           =   1320
      End
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   4485
      MaxLength       =   30
      TabIndex        =   77
      Text            =   "Text1"
      Top             =   2205
      Width           =   660
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   1
      Left            =   5145
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   76
      Text            =   "Text2"
      Top             =   2205
      Width           =   3525
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   4485
      MaxLength       =   30
      TabIndex        =   75
      Text            =   "Text1"
      Top             =   1845
      Width           =   660
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   0
      Left            =   5145
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   74
      Text            =   "Text2"
      Top             =   1845
      Width           =   3525
   End
   Begin VB.Frame Frame2 
      Height          =   840
      Index           =   0
      Left            =   45
      TabIndex        =   65
      Top             =   540
      Width           =   11415
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Tag             =   "Tipo|N|N|0|1|tcafpc|tipo|||"
         Text            =   "Combo1"
         Top             =   450
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   31
         Left            =   3990
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Fecha Recepción|F|N|||tcafpc|fecrecep|dd/mm/yyyy|N|"
         Top             =   450
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   6150
         MaxLength       =   40
         TabIndex        =   5
         Tag             =   "Nombre Transportista|T|N|||tcafpc|nomtrans||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   450
         Width           =   3750
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   5265
         MaxLength       =   3
         TabIndex        =   4
         Tag             =   "Cod.Agencia|N|N|0|999|tcafpc|codtrans|000|S|"
         Text            =   "Text1"
         Top             =   450
         Width           =   825
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Factura|F|N|||tcafpc|fecfactu|dd/mm/yyyy|S|"
         Top             =   450
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   0
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Nº Factura|T|N|||tcafpc|numfactu||S|"
         Text            =   "Text1 7"
         Top             =   450
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Contabilizado"
         Height          =   255
         Index           =   0
         Left            =   9990
         TabIndex        =   36
         Tag             =   "Contabilizado|N|N|0|1|tcafpc|intconta||N|"
         Top             =   495
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Factura"
         Height          =   255
         Index           =   18
         Left            =   180
         TabIndex        =   118
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "F. Recepción"
         Height          =   255
         Index           =   2
         Left            =   4005
         TabIndex        =   69
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Agencia"
         Height          =   255
         Index           =   0
         Left            =   5310
         TabIndex        =   68
         Top             =   180
         Width           =   735
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   6210
         ToolTipText     =   "Buscar proveedor"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F. Factura"
         Height          =   255
         Index           =   29
         Left            =   2880
         TabIndex        =   67
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Factura"
         Height          =   255
         Index           =   28
         Left            =   1710
         TabIndex        =   66
         Top             =   180
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   13
      Left            =   3555
      MaxLength       =   4
      TabIndex        =   72
      Text            =   "Text1"
      Top             =   720
      Width           =   540
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   13
      Left            =   4140
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   71
      Text            =   "Text2"
      Top             =   720
      Width           =   3285
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   17
      Left            =   7500
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   61
      Text            =   "ABCDKFJADKSFJAK"
      Top             =   5430
      Visible         =   0   'False
      Width           =   1725
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   480
      Top             =   5160
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
      Left            =   1920
      Top             =   5160
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
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   16
      Left            =   2355
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   60
      Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
      Top             =   6690
      Visible         =   0   'False
      Width           =   5085
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   75
      TabIndex        =   38
      Top             =   6525
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
         TabIndex        =   39
         Top             =   180
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10485
      TabIndex        =   35
      Top             =   6615
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9315
      TabIndex        =   34
      Top             =   6615
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
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
            Style           =   3
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
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Lineas Factura"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir Factura"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8400
         TabIndex        =   41
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10485
      TabIndex        =   37
      Top             =   6615
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
   Begin MSAdodcLib.Adodc Data4 
      Height          =   465
      Left            =   6840
      Top             =   765
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   820
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
   Begin VB.Image imgBuscar 
      Enabled         =   0   'False
      Height          =   240
      Index           =   2
      Left            =   4455
      Picture         =   "frmVtasHcoFactTra.frx":0AA1
      ToolTipText     =   "Buscar población"
      Top             =   900
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   6
      Left            =   4125
      Picture         =   "frmVtasHcoFactTra.frx":0BA3
      ToolTipText     =   "Buscar trabajador"
      Top             =   2220
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Trabajador Pedido"
      Height          =   255
      Index           =   9
      Left            =   2565
      TabIndex        =   79
      Top             =   2220
      Width           =   1425
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   5
      Left            =   4140
      Picture         =   "frmVtasHcoFactTra.frx":0CA5
      ToolTipText     =   "Buscar trabajador"
      Top             =   1845
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Trabajador Albaran"
      Height          =   255
      Index           =   21
      Left            =   2565
      TabIndex        =   78
      Top             =   1890
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Trabajador"
      Height          =   255
      Index           =   1
      Left            =   2340
      TabIndex        =   73
      Top             =   720
      Width           =   855
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   4
      Left            =   3270
      Picture         =   "frmVtasHcoFactTra.frx":0DA7
      ToolTipText     =   "Buscar trabajador"
      Top             =   750
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Lote"
      Height          =   255
      Index           =   3
      Left            =   7500
      TabIndex        =   70
      Top             =   5250
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Ampliación Línea"
      Height          =   255
      Index           =   35
      Left            =   2355
      TabIndex        =   43
      Top             =   6510
      Visible         =   0   'False
      Width           =   1335
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
      Begin VB.Menu mnLineas 
         Caption         =   "&Lineas"
         Enabled         =   0   'False
         HelpContextID   =   2
         Shortcut        =   ^L
         Visible         =   0   'False
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Shortcut        =   ^I
         Visible         =   0   'False
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmVtasHcoFactTra"
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
Public hcoFechaMovim As Date 'Fecha del Movim
Public hcoCodProve As Integer 'Codigo de Proveedor

'========== VBLES PRIVADAS ====================
Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
'--monica
'Private WithEvents frmCP As frmCPostal 'Codigos Postales

Private WithEvents frmTra As frmManAgencias  'Form Mto Transportistas
Attribute frmTra.VB_VarHelpID = -1
Private WithEvents frmPV As frmComProveV  'Form Mto Proveedores Varios
Attribute frmPV.VB_VarHelpID = -1
Private WithEvents frmFP As frmManFpago 'Form Mto Formas de Pago
Attribute frmFP.VB_VarHelpID = -1
'--monica
'Private WithEvents frmT As frmAdmTrabajadores  'Form Mto Trabajadores


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

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean

'Si el cliente mostrado es de Varios o No
Dim EsDeVarios As Boolean


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


' utilizado para buscar por checks
Private BuscaChekc As String



Private Sub Check1_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "Check1(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "Check1(" & Index & ")|"
    End If
End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpress KeyAscii
End Sub


Private Sub cmdAceptar_Click()

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 4  'MODIFICAR
            If DatosOk Then
               If ModificarFactura Then
                    TerminaBloquear
'                    PosicionarData
               Else
                    '---- Laura 24/10/2006
                    'como no hemos modificado dejamos la fecha como estaba ya que ahora se puede modificar
                    Text1(1).Text = Me.Data1.Recordset!FecFactu
               End If
               PosicionarData
            End If
            
         Case 5 'InsertarModificar LINEAS
            'Actualizar el registro en la tabla de lineas 'slialb'
            If ModificaLineas = 1 Then 'INSERTAR lineas Albaran
'                PrimeraLin = False
'                If Data2.Recordset.EOF = True Then PrimeraLin = True
'                If InsertarLinea(NumLinea) Then
'                    'Comprobar si el Articulo tiene control de Nº de Serie
'                    ComprobarNSeriesLineas NumLinea
'                    If PrimeraLin Then
'                        CargaGrid DataGrid1, Data2, True
'                    Else
'                        CargaGrid2 DataGrid1, Data2
'                    End If
'                    BotonAnyadirLinea
'                End If
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then
                    TerminaBloquear
                    NumRegElim = Data2.Recordset.AbsolutePosition
                    
                    CargaGrid2 DataGrid1, Data2
                    ModificaLineas = 0
                    PonerBotonCabecera True
                    BloquearTxt Text2(16), True
                    BloquearTxt Text2(17), True
           
                    LLamaLineas Modo, 0, "DataGrid1"
                    PosicionarData
                    If (Not Data2.Recordset.EOF) And (Not Data2.Recordset.BOF) Then
                        SituarDataPosicion Data2, NumRegElim, ""
                    End If
                End If
                Me.DataGrid1.Enabled = True
                Me.DataGrid2.Enabled = True
                Me.DataGrid3.Enabled = True
            End If
    End Select
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 0, 1 'Busqueda, Insertar
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
            LLamaLineas Modo, 0, "DataGrid2"
            LLamaLineas Modo, 0, "DataGrid3"
            PonerFoco Text1(0)
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
            
        Case 5 'Lineas Detalle
            TerminaBloquear
            BloquearTxt Text2(16), True
            BloquearTxt Text2(17), True
'            If ModificaLineas = 1 Then 'INSERTAR
'                ModificaLineas = 0
'                DataGrid1.AllowAddNew = False
'                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
'            End If
            ModificaLineas = 0
            LLamaLineas Modo, 0, "DataGrid1"
            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
    End Select
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
        LLamaLineas Modo, anc, "DataGrid3"
        
        'Si pasamos el control aqui lo ponemos en amarillo
        
'        PonerFoco Text1(0)
'        Text1(0).BackColor = vbYellow

        PonerFocoCmb Combo1(0)
        Combo1(0).BackColor = vbYellow
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
'            Text1(kCampo).BackColor = vbYellow
'            PonerFoco Text1(kCampo)
            PonerFocoCmb Combo1(0)
            Combo1(0).BackColor = vbYellow
        End If
    End If
End Sub


Private Sub BotonVerTodos()

    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        LimpiarCampos

        LimpiarDataGrids
        CadenaConsulta = "Select tcafpc.* "
        CadenaConsulta = CadenaConsulta & "from " & NombreTabla & Ordenacion
        

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
Dim DeVarios As Boolean

    'solo se puede modificar la factura si no esta contabilizada
    '++monica:añadida la condicion de solo si hay contabilidad
    If vParamAplic.NumeroConta <> 0 Then
        If FactContabilizada Then
            TerminaBloquear
            Exit Sub
        End If
    End If
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    PonerFocoChk Me.Check1(0)
        
    'Si es proveedor de Varios no se pueden modificar sus datos
'--monica
'    DeVarios = EsProveedorVarios(Text1(2).Text)
    DeVarios = False
    BloquearDatosTrans (DeVarios)
End Sub


Private Sub BotonModificarLinea()
'Modificar una linea
Dim vWhere As String
Dim anc As Single
Dim J As Byte
On Error GoTo eModificarLinea




    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then
        TerminaBloquear
        Exit Sub '1= Insertar
    End If
    
    If Data2.Recordset.EOF Then
        TerminaBloquear
        Exit Sub
    End If
    
    vWhere = ObtenerWhereCP(False)
    vWhere = vWhere & " AND numalbar=" & Data3.Recordset.Fields!NumAlbar & ""
    vWhere = vWhere & " and numlinea=" & Data2.Recordset!numlinea
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then
        TerminaBloquear
        Exit Sub
    End If

    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        J = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, J
        DataGrid1.Refresh
    End If
    
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 210
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 20
    End If

    For J = 0 To 2
        txtAux(J).Text = DataGrid1.Columns(J + 5).Text
    Next J
    Text2(16).Text = DataGrid1.Columns(J + 5).Text
    For J = J + 1 To 8
        txtAux(J - 1).Text = DataGrid1.Columns(J + 5).Text
    Next J
'--monica
'    Text2(17).Text = DataGrid1.Columns(14).Text
    
    ModificaLineas = 2 'Modificar
    LLamaLineas ModificaLineas, anc, "DataGrid1"
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR LINEAS"
    PonerBotonCabecera False
    BloquearTxt Text2(16), False 'Campo Ampliacion Linea
    BloquearTxt Text2(17), False 'Campo Ampliacion Linea
'    PonerFoco txtAux(4)
    PonerFoco Text2(16)
    Me.DataGrid1.Enabled = False

eModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub LLamaLineas(xModo As Byte, Optional alto As Single, Optional grid As String)
Dim jj As Integer
Dim b As Boolean
'
'    Select Case grid
'        Case "DataGrid1"
'            DeseleccionaGrid Me.DataGrid1
'            'PonerModo xModo + 1
'
'            b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Lineas
'
'            For jj = 0 To txtAux.Count - 1
'                If jj = 4 Or jj = 5 Or jj = 6 Or jj = 7 Then
'                    txtAux(jj).Height = DataGrid1.RowHeight
'                    txtAux(jj).Top = alto
'                    txtAux(jj).visible = b
'                End If
'            Next jj
'
        If grid = "DataGrid2" Then
            DeseleccionaGrid Me.DataGrid2
            b = (xModo = 1)
             For jj = 0 To txtAux3.Count - 1
                txtAux3(jj).Height = DataGrid2.RowHeight
                txtAux3(jj).Top = alto
                txtAux3(jj).visible = b
            Next jj
        End If
        
        If grid = "DataGrid3" Then
            DeseleccionaGrid Me.DataGrid3
            b = (xModo = 1)
            For jj = 0 To Text4.Count - 1
                Text4(jj).Height = DataGrid3.RowHeight
                Text4(jj).Top = alto
                Text4(jj).visible = b
            Next jj
        End If
'    End Select
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Facturas (scafac)
' y los registros correspondientes de las tablas cab. albaranes (scafac1)
' y las lineas de la factura (slifac)
Dim cad As String
Dim NumPedElim As Long
On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    
    'solo se puede eliminar si no esta en la contabilidad
    If Me.Check1(0).Value = 1 Then Exit Sub
    
'    'solo se puede modificar la factura si no esta contabilizada
'    If FactContabilizada Then Exit Sub
    
    cad = "Cabecera de Facturas." & vbCrLf
    cad = cad & "-----------------------------------" & vbCrLf & vbCrLf
    cad = cad & "Va a eliminar la Factura:            "
    cad = cad & vbCrLf & "Agencia:  " & Text1(2).Text & " - " & Text1(3).Text
    cad = cad & vbCrLf & "Nº Fact.:  " & Text1(0).Text
    cad = cad & vbCrLf & "Fecha:  " & Format(Text1(1).Text, "dd/mm/yyyy")

    cad = cad & vbCrLf & vbCrLf & " ¿Desea Eliminarla? "
    
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
'        NumPedElim = Data1.Recordset.Fields(1).Value
        
        If Not eliminar Then
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
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
        Screen.MousePointer = vbDefault
        If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Factura", Err.Description
End Sub


Private Sub cmdObserva_Click()
    If Modo <> 2 And Modo <> 4 Then Exit Sub
    If Me.FrameObserva.visible = False Then
        Me.DataGrid1.visible = False
        Me.FrameObserva.visible = True
        Me.cmdObserva.Picture = frmPpal.imgListComun.ListImages(18).Picture
'        CargarICO Me.cmdObserva, "volver.ico"
        Me.cmdObserva.ToolTipText = "volver lineas albaran"
        BloqueaText3
    Else
        Me.DataGrid1.visible = True
        Me.FrameObserva.visible = False
'        CargarICO Me.cmdObserva, "message.ico"
        Me.cmdObserva.Picture = frmPpal.imgListComun.ListImages(12).Picture
        Me.cmdObserva.ToolTipText = "ver observaciones albaran"
    End If
End Sub


Private Sub BloqueaText3()
Dim i As Byte
    'bloquear los Text3 que son las lineas de scafpa
    For i = 0 To 1
        BloquearTxt Text3(i), (Modo <> 4)
    Next i
    If Me.FrameObserva.visible Then
        For i = 4 To 8
            BloquearTxt Text3(i), (Modo <> 4)
        Next i
    End If
    'numpedpr, fecpedpr siempre bloqueados
    For i = 2 To 3
        BloquearTxt Text3(i), True
    Next i
End Sub


Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim cad As String

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
        cad = Data1.Recordset.Fields(0) & "|"
        cad = cad & Data1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(cad)
        Unload Me
    End If
End Sub


Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
        If Combo1(0).BackColor = vbYellow Then
            If Combo1(0).Locked Then
                Combo1(0).BackColor = &H80000018
            Else
                Combo1(0).BackColor = vbWhite
            End If
        End If
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo Error1

    If Not Data2.Recordset.EOF Then
        If ModificaLineas <> 1 Then
'            Text2(16).Text = DBLet(Data2.Recordset.Fields!ampliaci)
'            Text2(17).Text = DBLet(Data2.Recordset.Fields!numlotes)
        End If
    Else
        Text2(16).Text = ""
        Text2(17).Text = ""
    End If

Error1:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim i As Byte

    If Not Data3.Recordset.EOF Then
'        Text3(0).Text = DBLet(Data3.Recordset.Fields!codtrab2, "T")
'        Text3_LostFocus (0)
'        Text3(1).Text = DBLet(Data3.Recordset.Fields!codtrab1, "T")
'        Text3_LostFocus (1)

        Text3(2).Text = DBLet(Data3.Recordset.Fields!ImporteL, "N")
        If Text3(2).Text <> "0" Then
            FormateaCampo Text3(2)
        Else
            Text3(2).Text = ""
        End If
'--monica
'        Text3(3).Text = DBLet(Data3.Recordset.Fields!fecpedpr, "F")
        
        'Observaciones
        Text3(4).Text = DBLet(Data3.Recordset.Fields!observa1, "T")
        Text3(5).Text = DBLet(Data3.Recordset.Fields!observa2, "T")
        Text3(6).Text = DBLet(Data3.Recordset.Fields!observa3, "T")
        Text3(7).Text = DBLet(Data3.Recordset.Fields!observa4, "T")
        Text3(8).Text = DBLet(Data3.Recordset.Fields!observa5, "T")
        
        'Datos de la tabla
        CargaGrid DataGrid1, Data2, True
    Else
        For i = 0 To Text3.Count - 1
            If i <> 3 Then Text3(i).Text = ""
        Next i
        Text2(0).Text = ""
        Text2(1).Text = ""
        'Datos de la tabla slipre
        CargaGrid DataGrid1, Data2, False
    End If
    
End Sub



Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
    'Viene de DblClick en frmAlmMovimArticulos y carga el form con los valores
    If hcoCodMovim <> "" And Not Data1.Recordset.EOF Then PonerCadenaBusqueda
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
Dim i As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon

    ' ICONITOS DE LA BARRA
    btnPrimero = 15
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(5).Image = 4   'Modificar
        .Buttons(6).Image = 5   'Borrar
        .Buttons(9).Image = 15 'Mto Lineas Ofertas
        .Buttons(10).Image = 10 'Imprimir
        .Buttons(12).Image = 11  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    Me.SSTab1.Tab = 0
      
    'cargar IMAGES de busqueda
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
      
      
      
    LimpiarCampos   'Limpia los campos TextBox
     
    'cargar icono de observaciones de los albaranes de factura
'    CargarICO Me.cmdObserva, "message.ico"
    Me.cmdObserva.Picture = frmPpal.imgListComun.ListImages(12).Picture '--monica antes 41
    Me.FrameObserva.visible = False
    Me.cmdObserva.ToolTipText = "ver observaciones albaran"
    
    VieneDeBuscar = False
            
    '## A mano
    NombreTabla = "tcafpc"
    NomTablaLineas = "tlifpc" 'Tabla lineas de Facturacion
    Ordenacion = " ORDER BY tcafpc.fecrecep desc ,tcafpc.codtrans, tcafpc.numfactu "
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    CadenaConsulta = "Select * from " & NombreTabla
    
    If hcoCodMovim <> "" Then
        'Se llama desde Dobleclick en frmAlmMovimArticulos
'        CadenaConsulta = CadenaConsulta & " WHERE numalbar='" & hcoCodMovim & "' AND fechaalb= """ & Format(hcoFechaMovim, "yyyy-mm-dd") & """"
'        CadenaConsulta = CadenaConsulta & " AND codprove=" & hcoCodProve
        CadenaConsulta = CadenaConsulta & ObtenerSelFactura
    Else
        CadenaConsulta = CadenaConsulta & " where numfactu=-1"
    End If
    
    Data1.ConnectionString = Conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
    CargaCombo
    
    If hcoCodMovim = "" Then
        If DatosADevolverBusqueda = "" Then
            PonerModo 0
        Else
            PonerModo 1
        End If
        'Poner los grid sin apuntar a nada
    Else
         PonerModo 0
    End If
    LimpiarDataGrids
    PrimeraVez = False

End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.Check1(0).Value = 0
    Me.Combo1(0).ListIndex = -1
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    If Modo = 4 Or Modo = 5 Then TerminaBloquear
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadB = Aux
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
        CadB = CadB & " and " & Aux
        Aux = ValorDevueltoFormGrid(Text1(2), CadenaDevuelta, 3)
        CadB = CadB & " and " & Aux
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB
        CadenaConsulta = CadenaConsulta & " GROUP BY tcafpc.codtrans, tcafpc.numfactu, tcafpc.fecfactu "
        CadenaConsulta = CadenaConsulta & " " & Ordenacion
        PonerCadenaBusqueda
'        Text1(0).Text = RecuperaValor(CadenaDevuelta, 2)
    End If
    Screen.MousePointer = vbDefault
End Sub

'--monica
'Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
''Formulario Mantenimiento C. Postales
'Dim Indice As Byte
'Dim devuelve As String
'
'        Indice = 7
'        Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
'        Text1(Indice + 1).Text = ObtenerPoblacion(Text1(Indice).Text, devuelve)  'Poblacion
'        'provincia
'        Text1(Indice + 2).Text = devuelve
'End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
Dim Indice As Byte
    Indice = 10
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Forma Pago
    Text2(10).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub


Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Agencias
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1)  'Cod Trans
End Sub

Private Sub frmPV_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Proveedores Varios
Dim Indice As Byte

    Indice = 4
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'NIF
    Text1(Indice - 1).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Clien
    PonerDatosProveVario (Text1(Indice).Text)
End Sub

'--monica
'Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
''Form Mantenimiento de Trabajadores
'Dim Indice As Byte
'
'    Indice = Val(Me.imgBuscar(4).Tag)
'    If Indice = 4 Then
'        Indice = Indice + 9
'        Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'Cod Trabajador
'        Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
'    Else
'        Text3(Indice - 5).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'Cod Trabajador
'        Text2(Indice - 5).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
'    End If
'End Sub


Private Sub imgBuscar_Click(Index As Integer)
Dim Indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. Transportista
            PonerFoco Text1(2)
            Set frmTra = New frmManAgencias
            frmTra.DatosADevolverBusqueda = "0|"
            frmTra.Show vbModal
            Set frmTra = Nothing
            Indice = 2
            PonerFoco Text1(Indice)
            
'--monica
'        Case 1 'NIF para proveedor de Varios
'            Set frmPV = New frmComProveV
'            frmPV.DatosADevolverBusqueda = "0|"
'            frmPV.Show vbModal
'            Set frmPV = Nothing
'            Indice = 7
'            PonerFoco Text1(Indice)
            
'--monica
'        Case 2 'Cod. Postal
'            Set frmCP = New frmCPostal
'            frmCP.DatosADevolverBusqueda = "0"
'            frmCP.Show vbModal
'            Set frmCP = Nothing
'            Indice = 7
'            VieneDeBuscar = True
'            PonerFoco Text1(Indice)
      
         Case 3 'Forma de Pago
            Indice = 10
            PonerFoco Text1(Indice)
            Set frmFP = New frmManFpago
            frmFP.DatosADevolverBusqueda = "0|"
            frmFP.Show vbModal
            Set frmFP = Nothing
            
'--monica
'        Case 4, 5, 6 'Realizada Por Trabajador (Pedido, Albaran, Preparador Material
'            Me.imgBuscar(4).Tag = Index
'            Set frmT = New frmAdmTrabajadores
'            frmT.DatosADevolverBusqueda = "0"
'            frmT.Show vbModal
'            Set frmT = Nothing
'            If Index = 4 Then
'                PonerFoco Text1(13)
'            Else
'                PonerFoco Text3(Index - 5)
'            End If
    End Select
    
    Screen.MousePointer = vbDefault
End Sub




Private Sub mnBuscar_Click()
    Me.SSTab1.Tab = 0
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
'    BotonImprimir (53) '53: Informe de Facturas
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
        'bloquea la tabla cabecera de factura: scafpc
        If BLOQUEADesdeFormulario(Me) Then
            'bloquear la tabla cabecera de albaranes de la factura: scafpa
            If BloqueaAlbxFac Then BotonModificar
        End If
    End If
End Sub


Private Function BloqueaAlbxFac() As Boolean
'bloquea todos los albaranes de la factura
Dim sql As String
On Error GoTo EBloqueaAlb

    BloqueaAlbxFac = False
    'bloquear cabecera albaranes x factura
    sql = "select * FROM tcafpa "
    sql = sql & ObtenerWhereCP(True) & " FOR UPDATE"
    Conn.Execute sql, , adCmdText
    BloqueaAlbxFac = True

EBloqueaAlb:
    If Err.Number <> 0 Then BloqueaAlbxFac = False
End Function


Private Function BloqueaLineasFac() As Boolean
'bloquea TODAS las lineas de la factura
Dim sql As String
    
    On Error GoTo EBloqueaLin

    BloqueaLineasFac = False
    
    'bloquear cabecera albaranes x factura
    sql = "select * FROM tlifpc "
    sql = sql & ObtenerWhereCP(True) & " FOR UPDATE"
    Conn.Execute sql, , adCmdText
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
    ConseguirFoco Text1(Index), Modo
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
        
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1, 31 'Fecha factura,fecha recepcion
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
                
'--monica
'        Case 13 'Cod trabajador
'            Text2(Index).Text = PonerNombreDeCod(Text1(Index), cAgro, "straba", "nomtraba", "codtraba")

        Case 2 'Cod. transportista
            If Modo = 1 Then 'Modo=1 Busqueda
                Text1(Index + 1).Text = PonerNombreDeCod(Text1(Index), "agencias", "nomtrans")
            Else
                PonerDatosTransportista (Text1(Index).Text)
            End If
        
        Case 4 'NIF
            If Not EsDeVarios Then Exit Sub
            If Modo = 4 Then 'Modificar
                'si no se ha modificado el nif del cliente no hacer nada
                If Text1(4).Text = DBLet(Data1.Recordset!nifProve, "T") Then
                    Exit Sub
                End If
            End If
            PonerDatosProveVario (Text1(Index).Text)
            
'--monica
'        Case 7 'Cod. Postal
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
        
        
        Case 10 'Forma de Pago
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "forpago", "nomforpa")
            Else
                Text2(Index).Text = ""
            End If
            
        Case 11, 12 'Descuentos
            If Modo = 4 Then 'comprobar que el dato a cambiado
                If Index = 11 Then
                    If CCur(Text1(Index).Text) = CCur(Data1.Recordset!DtoPPago) Then Exit Sub
                ElseIf Index = 12 Then
                    If CCur(Text1(Index).Text) = CCur(Data1.Recordset!DtoGnral) Then Exit Sub
                End If
            End If
            
            If Modo = 3 Or Modo = 4 Then
                If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 4 'Tipo 4: Decimal(4,2)
                If Not ActualizarDatosFactura Then
                   If Index = 11 Then Text1(Index).Text = Data1.Recordset!DtoPPago
                   If Index = 12 Then Text1(Index).Text = Data1.Recordset!DtoGnral
                   FormateaCampo Text1(Index)
                End If
            End If
    End Select
End Sub


Private Sub HacerBusqueda()
Dim CadB As String

    CadB = ObtenerBusqueda(Me, BuscaChekc)
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select " & NombreTabla & ".* from (" & NombreTabla & " LEFT OUTER JOIN tcafpa ON " & NombreTabla & ".codtrans=tcafpa.codtrans AND " & NombreTabla & ".numfactu=tcafpa.numfactu AND " & NombreTabla & ".fecfactu=tcafpa.fecfactu) LEFT OUTER JOIN tcafpv ON " & NombreTabla & ".codtrans=tcafpv.codtrans AND " & NombreTabla & ".numfactu=tcafpv.numfactu AND " & NombreTabla & ".fecfactu=tcafpv.fecfactu"
        CadenaConsulta = CadenaConsulta & " WHERE " & CadB
        CadenaConsulta = CadenaConsulta & " GROUP BY tcafpc.codtrans, tcafpc.numfactu, tcafpc.fecfactu "
        CadenaConsulta = CadenaConsulta & " " & Ordenacion
        
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim Tabla As String
Dim Titulo As String
Dim devuelve As String
    
    'Llamamos a al form
    '##A mano
    cad = ""
'        cad = cad & ParaGrid(Text1(1), 10, "Tipo Fac.")
        cad = cad & ParaGrid(Text1(0), 18, "Nº Factura")
        cad = cad & ParaGrid(Text1(1), 15, "Fecha Fac.")
        cad = cad & ParaGrid(Text1(2), 12, "Prov.")
        cad = cad & ParaGrid(Text1(3), 55, "Nombre Prov")
        Tabla = NombreTabla
        Titulo = "Facturas"
        devuelve = "0|1|2|"
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = Tabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        '###A mano
'        frmB.vDevuelve = "0|1|"
        frmB.vDevuelve = devuelve
        frmB.vTitulo = Titulo
        frmB.vSelElem = 0
'--monica
'        frmB.vConexionGrid = conAri  'Conexión a BD: Ariges

'        If Not EsCabecera Then frmB.Label1.FontSize = 11
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
'''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
        'End If
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
        If Modo = 1 Then PonerFoco Text1(0)
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
        LLamaLineas Modo, 0, "DataGrid2"
        LLamaLineas Modo, 0, "DataGrid3"
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
'Carga el grid de los AlbaranesxFactura, es decir, la tabla scafpc de la factura seleccionada
On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass
    
    'Datos de la tabla albaranes x factura: scafpa
    CargaGrid DataGrid2, Data3, True
    CargaGrid DataGrid3, Data4, True
   
    Screen.MousePointer = vbDefault
    Exit Sub
    
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim BrutoFac As Single
Dim ImpTotal As Currency

    On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    'Poner la base imponible (impbruto - dtoppago - dtognral
    BrutoFac = CSng(Text1(14).Text) - CSng(Text1(15).Text) - CSng(Text1(16).Text)
    Text1(17).Text = Format(BrutoFac, FormatoImporte)
    Text1(32).Text = Format(BrutoFac, FormatoImporte)
    
    'poner descripcion campos
    Text2(10).Text = PonerNombreDeCod(Text1(10), "forpago", "nomforpa")
'--monica
'    Text2(13).Text = PonerNombreDeCod(Text1(13), conAri, "straba", "nomtraba", "codtraba")
    
    ImpTotal = DevuelveValor("select sum(importel) from tcafpv where " & ObtenerWhereCP(False))
    Text5.Text = Format(ImpTotal, FormatoImporte)
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Ofertas
    
'++monica
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
'++
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
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
    ActualizarToolbar Modo, Kmodo
    
    BuscaChekc = ""
        
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
    '---- laura 24/10/2006: si ponemos las claves de la tabla con ON UPDATE CASCADE
    'podemos permitir modificar la fecha de la factura que es clave primaria
'    If Modo = 4 Then BloquearTxt Text1(1), False
    
    BloquearCombo Me, Modo
    If Modo = 4 Then BloquearCombo Me, 5
    
    Me.Check1(0).Enabled = (Modo = 1 Or Modo = 3)
    
    b = (Modo <> 1)
    'Campos Nº Factura bloqueado y en azul
    BloquearTxt Text1(0), b, True
    BloquearTxt Text1(3), b 'referencia
    
    'Importes siempre bloqueados
    For i = 14 To 30
        BloquearTxt Text1(i), (Modo <> 1)
    Next i
    For i = 32 To 34
        BloquearTxt Text1(i), (Modo <> 1)
    Next i

    'Campo B.Imp y Imp. IVA siempre en azul
    Text1(17).BackColor = &HFFFFC0
    Text1(27).BackColor = &HFFFFC0
    Text1(28).BackColor = &HFFFFC0
    Text1(29).BackColor = &HFFFFC0
    Text1(30).BackColor = &HC0C0FF
    
    'bloquear los Text3 que son las lineas de scafac1
    BloqueaText3
    
    'Si no es modo lineas Boquear los TxtAux
    For i = 0 To txtAux.Count - 1
        BloquearTxt txtAux(i), (Modo <> 5)
    Next i
    BloquearTxt txtAux(8), True
    
    'Si no es modo Busqueda Bloquear los TxtAux3 (son los txtaux de los albaranes de factura)
    For i = 0 To txtAux3.Count - 1
        BloquearTxt txtAux3(i), (Modo <> 1)
    Next i
    For i = 0 To Text4.Count - 1
        BloquearTxt Text4(i), (Modo <> 1)
    Next i
    
    'ampliacion linea
    b = (Modo = 5) And Me.DataGrid1.visible
    'Modo Linea de Albaranes
    Me.Label1(35).visible = b
    Me.Label1(3).visible = b
    Me.Text2(16).visible = b
    Me.Text2(17).visible = b
    BloquearTxt Text2(16), (Modo <> 5) Or (Modo = 5 And ModificaLineas <> 1)
    BloquearTxt Text2(17), (Modo <> 5) Or (Modo = 5 And ModificaLineas <> 1)

    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2 And Modo <> 5)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = b
    Next i
    Me.imgBuscar(0).Enabled = (Modo = 1)
    Me.imgBuscar(1).visible = False
                    
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
On Error GoTo EDatosOK

    DatosOk = False
    
    'Para que no den errores los 0's de los importes de dtos
    ComprobarDatosTotales
        
    'comprobamos datos OK de la tabla scafac
    b = CompForm(Me) 'Comprobar formato datos ok de la cabecera: opcion=1
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
        If i = 4 Or i = 5 Or i = 6 Then
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
    If Index = 17 And KeyAscii = 13 Then 'campo nº de lote y ENTER
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub Text3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Text3_LostFocus(Index As Integer)
    Select Case Index
'--monica
'        Case 0, 1 'trabajador
'            Text2(Index).Text = PonerNombreDeCod(Text3(Index), conAri, "straba", "nomtraba", "codtraba", "Cod. Trabajador", "N")
        Case 8 'observa 5
            PonerFocoBtn Me.cmdAceptar
    End Select
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Buscar
            mnBuscar_Click
        Case 2  'Todos
            BotonVerTodos

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
    
    
Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Albaran: slialb
Dim sql As String
Dim vWhere As String
Dim b As Boolean

    On Error GoTo eModificarLinea

    ModificarLinea = False
    If Data2.Recordset.EOF Then Exit Function
    
    vWhere = ObtenerWhereCP(True)
    vWhere = vWhere & " AND numalbar='" & Data3.Recordset.Fields!NumAlbar & "'"
    vWhere = vWhere & " AND numlinea=" & Data2.Recordset.Fields!numlinea
    
    If DatosOkLinea() Then
        sql = "UPDATE " & NomTablaLineas & " SET "
        sql = sql & " ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
        sql = sql & "precioar= " & DBSet(txtAux(4).Text, "N") & ", "
        sql = sql & "dtoline1= " & DBSet(txtAux(5).Text, "N") & ", dtoline2= " & DBSet(txtAux(6).Text, "N") & ", "
        sql = sql & "importel= " & DBSet(txtAux(7).Text, "N")
        sql = sql & ", numlotes=" & DBSet(Text2(17).Text, "T")
        sql = sql & vWhere
    End If
    
    If sql <> "" Then
        'actualizar la factura y vencimientos
        b = ModificarFactura(sql)
        ModificarLinea = b
    End If
    
eModificarLinea:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Modificar Lineas Factura" & vbCrLf & Err.Description
        b = False
    End If
    ModificarLinea = b
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
    'Habilitar las opciones correctas del menu segun Modo
'    PonerModoOpcionesMenu (Modo)
'    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
    DataGrid2.Enabled = Not b
    DataGrid3.Enabled = Not b
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim b As Boolean
Dim Opcion As Byte
Dim sql As String

    On Error GoTo ECargaGRid

'    b = DataGrid1.Enabled
 
    Select Case vDataGrid.Name
        Case "DataGrid1"
            Opcion = 1
        Case "DataGrid2"
            Opcion = 2
        Case "DataGrid3"
            Opcion = 3
    End Select
    
    sql = MontaSQLCarga(enlaza, Opcion)
    CargaGridGnral vDataGrid, vData, sql, PrimeraVez
    
    CargaGrid2 vDataGrid, vData
    vDataGrid.ScrollBars = dbgAutomatic
    
    b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
    vDataGrid.Enabled = Not b
'    PrimeraVez = False
    Exit Sub
    
ECargaGRid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim tots As String

    On Error GoTo ECargaGRid
    
    vData.Refresh
    Select Case vDataGrid.Name
        Case "DataGrid1" 'Lineas de Albaran
'        SQL = "SELECT codtrans, numfactu, fecfactu, numalbar, numlinea, importel "
        
            'SQL = "SELECT codtipom, numfactu, fecfactu, numalbar, numlinea,
            'codalmac, codartic, nomartic, ampliaci, cantidad, precioar, origpre, dtoline1, dtoline2, importel "
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux(0)|T|Variedad|1900|;S|txtAux(1)|T|Var.Comercial|1900|;S|txtAux(2)|T|Marca|2500|;S|txtAux(3)|T|Forfait|2500|;S|txtAux(7)|T|Importe|1350|;"     '       S|txtAux(1)|T|Artículo|1750|;S|txtAux(2)|T|Nombre Art.|3150|;"
'            tots = tots & "N||||0|;S|txtAux(3)|T|Cantidad|950|;S|txtAux(4)|T|Precio|1300|;S|txtAux(5)|T|Dto 1|600|;S|txtAux(6)|T|Dto 2|600|;S|txtAux(7)|T|Importe|1350|;" 'N||||0|;"
            arregla tots, DataGrid1, Me
'            DataGrid1.Columns(9).Alignment = dbgRight
'            DataGrid1.Columns(10).Alignment = dbgRight
'            DataGrid1.Columns(12).Alignment = dbgRight
'            DataGrid1.Columns(13).Alignment = dbgRight
                       
         Case "DataGrid2" 'albaranes x articulo
'        SQL = "SELECT codtrans,numfactu,fecfactu,numalbar, fechaalb, importel,observa1,observa2,observa3,observa4,observa5  "
            
            
            'SQL = "SELECT codtipom,numfactu,fecfactu,codtipoa,numalbar, fechaalb,"
            'numpedcl,fecpedcl,sementre,numofert,fecofert, referenc, codenvio,codtraba, codtrab1, codtrab2,observa1,observa2,observa3,observa4,observa5  "
            tots = "N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux3(0)|T|Albaran|1400|;S|txtAux3(1)|T|Fecha|1300|;"
            tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;" 'N||||0|;N||||0|;"
'            tots = tots & "N||||0|;"
            arregla tots, DataGrid2, Me
        
            DataGrid2_RowColChange 1, 1
    
         Case "DataGrid3" 'cuentas contables por portes vuelta
'        SQL = "SELECT codtrans,numfactu,fecfactu,codmacta, nommacta, importel "
            tots = "N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|Text4(0)|T|Cta.Contable|1400|;S|Text4(2)|T|Descripción|3400|;"
            tots = tots & "S|Text4(1)|T|Importe|1800|;"
            arregla tots, DataGrid3, Me
        
    End Select
    
    vDataGrid.HoldFields
    Exit Sub
ECargaGRid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)

    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    Select Case Index
        Case 4 'Precio
            If txtAux(Index).Text <> "" Then
                PonerFormatoDecimal txtAux(Index), 2 'Tipo 2: Decimal(10,4)
            End If
            
        Case 5, 6 'Descuentos
            PonerFormatoDecimal txtAux(Index), 4 'Tipo 4: Decimal(4,2)
            If Index = 6 Then PonerFoco Me.Text2(16)
            
        Case 7 'Importe Linea
            PonerFormatoDecimal txtAux(Index), 1 'Tipo 3: Decimal(12,2)
    End Select
    
    If (Index = 3 Or Index = 4 Or Index = 6 Or Index = 7) Then 'Cant., Precio, Dto1, Dto2
        If txtAux(1).Text = "" Then Exit Sub
'        txtAux(7).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(5).Text, txtAux(6).Text, vParamAplic.TipoDtos)
        txtAux(7).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(5).Text, txtAux(6).Text, 0, 0)
        PonerFormatoDecimal txtAux(7), 1
    End If
End Sub


Private Sub BotonMtoLineas(numTab As Integer, cad As String)
    Me.SSTab1.Tab = numTab
    
    If Me.DataGrid1.visible Then 'Lineas de Albaranes
        If Me.Data2.Recordset.RecordCount < 1 Then
            MsgBox "La factura no tiene lineas.", vbInformation
            Exit Sub
        End If
        TituloLinea = cad
        
        ModificaLineas = 0
        PonerModo 5
        PonerBotonCabecera True
    End If
End Sub


Private Function eliminar() As Boolean
Dim sql As String
Dim Sql2 As String
Dim cta As String
Dim b As Boolean

    On Error GoTo FinEliminar

        b = False
        eliminar = False
        If Data1.Recordset.EOF Then Exit Function
        
        Conn.BeginTrans
        ConnConta.BeginTrans
        
        'Eliminar en la tabla pagos de la Contabilidad: spagop
        '------------------------------------------------
        cta = DevuelveDesdeBDNew(cAgro, "agencias", "codmacta", "codtrans", Text1(2).Text, "N")
        sql = " ctaprove='" & cta & "' AND numfactu='" & Data1.Recordset.Fields!NumFactu & "'"
        sql = sql & " AND fecfactu='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
        ConnConta.Execute "Delete from spagop WHERE " & sql
        b = True
        
        
        'Eliminar en tablas de factura de Ariagro: tcafpc, tcafpa, tlifpc
        '---------------------------------------------------------------
        If b Then
            sql = " " & ObtenerWhereCP(True)
        
'[Monica] 11/01/2010: ahora se decrementa el impportes de coste bien de comision o de portes
'            ' borramos los costes (albaran_costes)
'            Sql2 = "delete from albaran_costes where tipogasto = 2 "
'            Sql2 = Sql2 & " and numalbar in (select numalbar from tcafpa  " & SQL & ")"
'            conn.Execute Sql2
            
            Select Case Data1.Recordset.Fields!Tipo
                Case 0 ' factura de transportista
                    Sql2 = "update albaran_costes, tlifpc  set impcoste = impcoste - tlifpc.importel, importes = importes - tlifpc.importel "
                    Sql2 = Sql2 & " where albaran_costes.numalbar = tlifpc.numalbar and "
                    Sql2 = Sql2 & " albaran_costes.numlinea = tlifpc.numlinea and "
                    Sql2 = Sql2 & " albaran_costes.tipogasto = 2 and "
                    Sql2 = Sql2 & " tlifpc.codtrans = " & Data1.Recordset!codTrans & " and "
                    Sql2 = Sql2 & " tlifpc.numfactu = '" & Trim(Data1.Recordset!NumFactu) & "' and "
                    Sql2 = Sql2 & " tlifpc.fecfactu = '" & Format(Data1.Recordset!FecFactu, "yyyy-mm-dd") & "'"
                 
                Case 1 ' factura de comisionista
                    Sql2 = "update albaran_costes, tlifpc  set impcoste = impcoste - tlifpc.importel, importes = importes - tlifpc.importel "
                    Sql2 = Sql2 & " where albaran_costes.numalbar = tlifpc.numalbar and "
                    Sql2 = Sql2 & " albaran_costes.numlinea = tlifpc.numlinea and "
                    Sql2 = Sql2 & " albaran_costes.tipogasto = 3 and "
                    Sql2 = Sql2 & " tlifpc.codtrans = " & Data1.Recordset!codTrans & " and "
                    Sql2 = Sql2 & " tlifpc.numfactu = '" & Trim(Data1.Recordset!NumFactu) & "' and "
                    Sql2 = Sql2 & " tlifpc.fecfactu = '" & Format(Data1.Recordset!FecFactu, "yyyy-mm-dd") & "'"
            End Select
            Conn.Execute Sql2
            
            
            ' actualizamos el importe asignado a cada albaran a cero
            Sql2 = "update albaran set portespag = 0 "
            Sql2 = Sql2 & " where numalbar in (select numalbar from tcafpa " & sql & ")"
            Conn.Execute Sql2
        
            'Lineas de facturas (tlifpc)
            Conn.Execute "Delete from " & NomTablaLineas & sql
        
            'Lineas de cabeceras de albaranes de la factura
            Conn.Execute "Delete from tcafpa " & sql
            
            'Lineas de cabeceras de portes de vuelta de la factura
            Conn.Execute "Delete from tcafpv " & sql
            
            
            'Cabecera de facturas (tcafpc)
            Conn.Execute "Delete from " & NombreTabla & sql
        End If
        
        'Eliminar los movimientos generados por el albaran que genero la factura
        '-----------------------------------------------------------------------
        If b Then
        
        End If
        
'        b = True
        
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Factura", Err.Description
        b = False
    End If
    If Not b Then
        Conn.RollbackTrans
        ConnConta.RollbackTrans
    Else
        Conn.CommitTrans
        ConnConta.CommitTrans
    End If
    eliminar = b
End Function


Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next
    CargaGrid DataGrid2, Data3, False
    CargaGrid DataGrid1, Data2, False
    CargaGrid DataGrid3, Data4, False
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP(False) & ")"
         If SituarDataMULTI(Data1, vWhere, Indicador) Then
             If Modo <> 5 Then
                PonerModo 2
                PonerCampos
             End If
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
Dim sql As String
On Error Resume Next
    sql = "codtrans= " & Text1(2).Text & " and numfactu= '" & Text1(0).Text & "' and fecfactu='" & Format(Text1(1).Text, FormatoFecha) & "' "
    If conWhere Then sql = " WHERE " & sql
    ObtenerWhereCP = sql
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
Dim sql As String
    
    Select Case Opcion
        Case 1
            sql = "SELECT codtrans, numfactu, fecfactu, tlifpc.numalbar, tlifpc.numlinea, a.nomvarie, b.nomvarie, nommarca, nomconfe, importel "
            sql = sql & " FROM tlifpc, albaran_variedad, variedades a, variedades b, marcas, forfaits  " 'lineas de factura
        Case 2
            sql = "SELECT codtrans,numfactu,fecfactu,numalbar, fechaalb, importel,observa1,observa2,observa3,observa4,observa5  "
            sql = sql & " FROM tcafpa " 'cabeceras albaranes de la factura
        Case 3
            sql = "SELECT codtrans,numfactu,fecfactu,tcafpv.codmacta,conta" & vParamAplic.NumeroConta & ".cuentas.nommacta, importel"
            sql = sql & " FROM tcafpv, conta" & vParamAplic.NumeroConta & ".cuentas " 'cabeceras cuentas de portes de vuelta
    End Select
    
    If enlaza Then
        sql = sql & " " & ObtenerWhereCP(True)
        'lineas factura proveedor
        If Opcion = 1 Then
            sql = sql & " AND tlifpc.numalbar=" & DBLet(Data3.Recordset.Fields!NumAlbar, "N")
            sql = sql & " AND albaran_variedad.numalbar = tlifpc.numalbar "
            sql = sql & " AND albaran_variedad.numlinea = tlifpc.numlinea "
            sql = sql & " AND albaran_variedad.codvarie = a.codvarie "
            sql = sql & " AND albaran_variedad.codvarco = b.codvarie "
            sql = sql & " AND albaran_variedad.codmarca = marcas.codmarca "
            sql = sql & " AND albaran_variedad.codforfait = forfaits.codforfait "
        End If
        If Opcion = 3 Then
            sql = sql & " AND tcafpv.codmacta = conta" & vParamAplic.NumeroConta & ".cuentas.codmacta "
        End If
    Else
        sql = sql & " WHERE numfactu = -1"
        If Opcion = 1 Then
            sql = sql & " AND albaran_variedad.numalbar = tlifpc.numalbar "
            sql = sql & " AND albaran_variedad.numlinea = tlifpc.numlinea "
            sql = sql & " AND albaran_variedad.codvarie = a.codvarie "
            sql = sql & " AND albaran_variedad.codvarco = b.codvarie "
            sql = sql & " AND albaran_variedad.codmarca = marcas.codmarca "
            sql = sql & " AND albaran_variedad.codforfait = forfaits.codforfait "
        End If
        If Opcion = 3 Then
            sql = sql & " AND tcafpv.codmacta = conta" & vParamAplic.NumeroConta & ".cuentas.codmacta "
        End If
    End If
    
    Select Case Opcion
        Case 1
            sql = sql & " ORDER BY codtrans, numfactu, fecfactu,tlifpc.numalbar,tlifpc.numlinea "
        Case 2
            sql = sql & " ORDER BY codtrans, numfactu, fecfactu,numalbar "
        Case 3
            sql = sql & " ORDER BY codtrans, numfactu, fecfactu,tcafpv.codmacta "
    End Select
    MontaSQLCarga = sql
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean

        b = ((Modo = 2) Or (Modo = 5 And ModificaLineas = 0)) And Check1(0).Value = 0
        'Modificar
        Toolbar1.Buttons(5).Enabled = b
        Me.mnModificar.Enabled = b
        'eliminar
        Toolbar1.Buttons(6).Enabled = b
        Me.mnEliminar.Enabled = b
            
'        b = (Modo = 2)
'        'Mantenimiento lineas
'        Toolbar1.Buttons(9).Enabled = b
'        Me.mnLineas.Enabled = b
        'Imprimir
'        Toolbar1.Buttons(10).Enabled = b
'        Me.mnImprimir.Enabled = b
        
        b = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(1).Enabled = Not b
        Me.mnBuscar.Enabled = Not b
        'Ver Todos
        Toolbar1.Buttons(2).Enabled = Not b
        Me.mnVerTodos.Enabled = Not b
End Sub


Private Sub PonerDatosTransportista(codTrans As String, Optional NIFTrans As String)
Dim vTrans As CTransportista
Dim observaciones As String
    
    On Error GoTo EPonerDatos
    
    If codTrans = "" Then
        LimpiarDatosTrans
        Exit Sub
    End If

    Set vTrans = New CTransportista
    'si se ha modificado el proveedor volver a cargar los datos
    If vTrans.Existe(codTrans) Then
        If vTrans.LeerDatos(codTrans) Then
        
'--monica
'            EsDeVarios = vProve.DeVarios
'            BloquearDatosProve (EsDeVarios)
'++monica
            EsDeVarios = False
            BloquearDatosTrans (EsDeVarios)
            
            If Modo = 4 And EsDeVarios Then 'Modificar
                'si no se ha modificado el proveedor no hacer nada
                If CLng(Text1(2).Text) = CLng(Data1.Recordset!codTrans) Then
                    Set vTrans = Nothing
                    Exit Sub
                End If
            End If
        
            Text1(2).Text = vTrans.codigo
            FormateaCampo Text1(2)
            
            If (Modo = 3) Or (Modo = 4) Then
                Text1(3).Text = vTrans.Nombre  'Nom prove
                Text1(6).Text = vTrans.Domicilio
                Text1(7).Text = vTrans.CPostal
                Text1(8).Text = vTrans.Poblacion
                Text1(9).Text = vTrans.Provincia
                Text1(4).Text = vTrans.NIF
                Text1(5).Text = DBLet(vTrans.TfnoAdmon, "T")
            End If
            
            observaciones = DBLet(vTrans.observaciones)
            If observaciones <> "" Then
                MsgBox observaciones, vbInformation, "Observaciones del transportista"
            End If
        End If
    Else
        LimpiarDatosTrans
    End If
    Set vTrans = Nothing

EPonerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Datos Transportista", Err.Description
End Sub


Private Sub PonerDatosProveVario(nifProve As String)
'Poner el los campos Text el valor del proveedor
Dim vProve As CProveedor
Dim b As Boolean
   
    If nifProve = "" Then Exit Sub
   
    Set vProve = New CProveedor
    b = vProve.LeerDatosProveVario(nifProve)
    If b Then
        Text1(3).Text = vProve.Nombre   'Nom proveedor
        Text1(6).Text = vProve.Domicilio
        Text1(7).Text = vProve.CPostal
        Text1(8).Text = vProve.Poblacion
        Text1(9).Text = vProve.Provincia
        Text1(5).Text = DBLet(vProve.TfnoAdmon, "T")
    End If
    Set vProve = Nothing
End Sub


Private Sub LimpiarDatosTrans()
Dim i As Byte

    For i = 3 To 9
        Text1(i).Text = ""
    Next i
End Sub
   

Private Function ModificaAlbxFac() As Boolean
Dim sql As String
Dim b As Boolean
On Error GoTo EModificaAlb
    
    ModificaAlbxFac = False
    If Data3.Recordset.EOF Then
        ModificaAlbxFac = True
        Exit Function
    End If
    'comprobar datos OK de la scafac1
     b = CompForm(Me) 'Comprobar formato datos ok de la cabecera alb: opcion=2
    If Not b Then Exit Function

'--monica
'    SQL = "UPDATE scafpa SET codtrab2=" & DBSet(Text3(0).Text, "N", "S") & ", "
'    SQL = SQL & "codtrab1=" & DBSet(Text3(1).Text, "N", "S")
    If Me.FrameObserva.visible Then
        sql = "UPDATE tcafpa SET "
        sql = sql & " observa1=" & DBSet(Text3(4).Text, "T")
        sql = sql & ", observa2=" & DBSet(Text3(5).Text, "T")
        sql = sql & ", observa3=" & DBSet(Text3(6).Text, "T")
        sql = sql & ", observa4=" & DBSet(Text3(7).Text, "T")
        sql = sql & ", observa5=" & DBSet(Text3(8).Text, "T")
        sql = sql & ObtenerWhereCP(True)
        sql = sql & " AND numalbar=" & Data3.Recordset.Fields!NumAlbar
        Conn.Execute sql
    End If
'--monica
'    SQL = SQL & ObtenerWhereCP(True)
'    SQL = SQL & " AND numalbar=" & Data3.Recordset.Fields!numalbar
'    Conn.Execute SQL
    ModificaAlbxFac = True
    
EModificaAlb:
If Err.Number <> 0 Then MuestraError Err.Number, "Modificar Albaranes de factura", Err.Description
End Function



Private Function ModificarFactura(Optional sqlLineas As String) As Boolean
'si se ha modificado la linea de slifac, añadir a la transaccion la modificación de la linea y recalcular
Dim bol As Boolean
Dim MenError As String
Dim sql As String
Dim vFactu As CFacturaTra
On Error GoTo EModFact

    bol = False
    Conn.BeginTrans
    '++monica:añadida la condicion de solo si hay contabilidad
    If vParamAplic.NumeroConta <> 0 Then ConnConta.BeginTrans
    
    If sqlLineas <> "" Then
        'actualizar el importe de la linea modificada
        MenError = "Modificando lineas de Factura."
        Conn.Execute sqlLineas
    End If
    
    'recalcular las bases imponibles x IVA
    MenError = "Recalcular importes IVA"
    bol = ActualizarDatosFactura
    
    If bol Then
        'modificamos la scafpc
        MenError = "Modificando cabecera de factura"
        bol = ModificaDesdeFormulario1(Me, 1)
        
'--monica : no hay agencias de transporte de varios
'        If bol Then
'            'Si es proveedor de varios actualizar datos proveedor en tabla:sprvar
'            MenError = "Modificando datos proveedor varios"
'            bol = ActualizarProveVarios(Text1(2).Text, Text1(4).Text)
'        End If
        
        If bol Then
            MenError = "Modificando albaranes de factura"
            'modificar la tabla: scafpa
            bol = ModificaAlbxFac
            '++monica:añadida la condicion de solo si hay contabilidad
            If vParamAplic.NumeroConta <> 0 Then
                If bol Then 'si se ha modificado la factura
                    MenError = "Actualizando en Tesoreria"
                    'y eliminar de tesoreria conta.spagop los registros de la factura
                    
                    'antes de Eliminar en las tablas de la Contabilidad
                    Set vFactu = New CFacturaTra
                    bol = vFactu.LeerDatos(Text1(2).Text, Text1(0).Text, Text1(1).Text)
                    
                    If bol Then
                        'Eliminar de la spagop
                        sql = " ctaprove='" & vFactu.CtaTrans & "' AND numfactu='" & Data1.Recordset.Fields!NumFactu & "'"
                        sql = sql & " AND fecfactu='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
                        ConnConta.Execute "Delete from spagop WHERE " & sql
                        
                        'Volvemos a grabar en TESORERIA. Tabla de Contabilidad: sconta.spagop
                        If bol Then
                            bol = vFactu.InsertarEnTesoreria(MenError)
                        End If
                    End If
                    Set vFactu = Nothing
                End If
            End If
        End If
    End If

EModFact:
     If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        bol = False
    End If
    If bol Then
        Conn.CommitTrans
        '++monica:añadida la condicion de solo si hay contabilidad
        If vParamAplic.NumeroConta <> 0 Then ConnConta.CommitTrans
        ModificarFactura = True
    Else
        Conn.RollbackTrans
        '++monica:añadida la condicion de solo si hay contabilidad
        If vParamAplic.NumeroConta <> 0 Then ConnConta.RollbackTrans
        ModificarFactura = False
        MenError = "Actualizando Factura." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
    End If
End Function



Private Function FactContabilizada() As Boolean
Dim cta As String, numasien As String
On Error GoTo EContab

    'comprabar que se puede modificar/eliminar la factura
    If Me.Check1(0).Value = 1 Then 'si esta contabilizada
        'comprobar en la contabilidad si esta contabilizada
        cta = DevuelveDesdeBDNew(cAgro, "agencias", "codmacta", "codtrans", Text1(2).Text, "N")
        If cta <> "" Then
            numasien = DevuelveDesdeBDNew(cConta, "cabfactprov", "numasien", "codmacta", cta, "T", , "numfacpr", Text1(0).Text, "T", "fecfacpr", Text1(1).Text, "F")
            If numasien <> "" Then
                FactContabilizada = True
                MsgBox "La factura esta contabilizada y no se puede modificar ni eliminar.", vbInformation
                Exit Function
            Else
                FactContabilizada = False
            End If
        Else
            FactContabilizada = True
            Exit Function
        End If
    Else
        FactContabilizada = False
    End If
EContab:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Factura contabilizada", Err.Description
End Function


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


Private Sub Text4_GotFocus(Index As Integer)
    ConseguirFoco Text4(Index), Modo
End Sub

Private Sub Text4_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index <> 0 And KeyCode <> 38 Then KEYdown KeyCode
End Sub

Private Sub Text4_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text4_LostFocus(Index As Integer)
    If Not PerderFocoGnral(Text4(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 0 'Cta Contable
            If Text4(Index).Text = "" Then Exit Sub
            Text4(2).Text = PonerNombreCuenta(Text4(Index), Modo)
            If Text4(2).Text = "" Then
                PonerFoco Text4(Index)
            End If
            
        Case 1 ' Importe
            If Text4(Index).Text <> "" Then PonerFormatoDecimal Text4(Index), 1
            
    End Select
End Sub



Private Sub BloquearDatosTrans(bol As Boolean)
Dim i As Byte

    'bloquear/desbloquear campos de datos segun sea de varios o no
    If Modo <> 5 Then
        Me.imgBuscar(1).visible = bol 'NIF
        Me.imgBuscar(1).Enabled = bol 'NIF
        Me.imgBuscar(2).Enabled = bol 'poblacion
        
        For i = 3 To 9 'si no es de varios no se pueden modificar los datos
            BloquearTxt Text1(i), Not bol
        Next i
    
    End If
End Sub

'--monica
'Private Function ActualizarProveVarios(Prove As String, NIF As String) As Boolean
''Modifica los datos de la tabla de Proveedores Varios
'Dim vProve As CProveedor
'On Error GoTo EActualizarCV
'
'    ActualizarProveVarios = False
'
'    Set vProve = New CProveedor
'    If EsProveedorVarios(Prove) Then
'        vProve.NIF = NIF
'        vProve.Nombre = Text1(3).Text
'        vProve.Domicilio = Text1(6).Text
'        vProve.CPostal = Text1(7).Text
'        vProve.Poblacion = Text1(8).Text
'        vProve.Provincia = Text1(9).Text
'        vProve.TfnoAdmon = Text1(5).Text
'        vProve.ActualizarProveV (NIF)
'    End If
'    Set vProve = Nothing
'
'    ActualizarProveVarios = True
'
'EActualizarCV:
'    If Err.Number <> 0 Then
'        ActualizarProveVarios = False
'    Else
'        ActualizarProveVarios = True
'    End If
'End Function


Private Function ObtenerSelFactura() As String
'Cuando venimos desde dobleClick en Movimientos de Articulos para Albaranes ya
'Facturados, abrimos este form pero cargando los datos de la factura
'correspendiente al albaran que se selecciono
Dim cad As String
Dim RS As ADODB.Recordset
On Error Resume Next

    cad = "SELECT codprove,numfactu,fecfactu FROM scafpa "
    cad = cad & " WHERE codprove=" & DBSet(hcoCodProve, "N") & " AND numalbar=" & DBSet(hcoCodMovim, "T")
    cad = cad & " AND fechaalb=" & DBSet(hcoFechaMovim, "F")

    Set RS = New ADODB.Recordset
    RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then 'where para la factura
        cad = " WHERE codprove=" & RS!codProve & " AND numfactu= '" & RS!NumFactu & "' AND fecfactu=" & DBSet(RS!FecFactu, "F")
    Else
        cad = " where numfactu=-1"
    End If
    RS.Close
    Set RS = Nothing

    ObtenerSelFactura = cad
End Function



Private Function ActualizarDatosFactura() As Boolean
Dim vFactu As CFacturaTra
Dim cadSel As String

    Set vFactu = New CFacturaTra
    cadSel = ObtenerWhereCP(False)
    cadSel = "tcafpa." & cadSel
    vFactu.DtoPPago = CCur(Text1(11).Text)
    vFactu.DtoGnral = CCur(Text1(12).Text)
    If vFactu.CalcularDatosFactura(cadSel, Text1(2).Text, True) Then '"tcafpa") Then
        Text1(14).Text = vFactu.BrutoFac
        Text1(15).Text = vFactu.ImpPPago
        Text1(16).Text = vFactu.ImpGnral
        Text1(17).Text = vFactu.BaseImp
        Text1(18).Text = vFactu.TipoIVA1
        Text1(19).Text = vFactu.TipoIVA2
        Text1(20).Text = vFactu.TipoIVA3
        Text1(21).Text = vFactu.PorceIVA1
        Text1(22).Text = vFactu.PorceIVA2
        Text1(23).Text = vFactu.PorceIVA3
        Text1(24).Text = vFactu.BaseIVA1
        Text1(25).Text = vFactu.BaseIVA2
        Text1(26).Text = vFactu.BaseIVA3
        Text1(27).Text = vFactu.ImpIVA1
        Text1(28).Text = vFactu.ImpIVA2
        Text1(29).Text = vFactu.ImpIVA3
        Text1(30).Text = vFactu.TotalFac
        
        FormatoDatosTotales
        
        ActualizarDatosFactura = True
    Else
        ActualizarDatosFactura = False
        MuestraError Err.Number, "Recalculando Factura", Err.Description
    End If
    Set vFactu = Nothing
End Function


Private Sub FormatoDatosTotales()
Dim i As Byte

    For i = 14 To 17
'        Text1(I).Text = QuitarCero(Text1(I).Text)
        FormateaCampo Text1(i)
    Next i
    
    For i = 24 To 26
        If Text1(i).Text <> "" Then
            'Si la Base Imp. es 0
            If CSng(Text1(i).Text) = 0 Then
                Text1(i).Text = QuitarCero(Text1(i).Text)
                Text1(i - 3).Text = QuitarCero(Text1(i - 3).Text)
                Text1(i - 6).Text = QuitarCero(Text1(i - 6).Text)
                Text1(i + 3).Text = QuitarCero(Text1(i + 3).Text)
            Else
                FormateaCampo Text1(i)
                FormateaCampo Text1(i - 3)
                FormateaCampo Text1(i - 6)
                FormateaCampo Text1(i + 3)
            End If
        Else 'No hay Base Imponible
            Text1(i - 3).Text = QuitarCero(Text1(i - 3).Text)
            Text1(i - 6).Text = QuitarCero(Text1(i - 6).Text)
            Text1(i + 3).Text = ""
        End If
    Next i
End Sub



Private Sub ComprobarDatosTotales()
Dim i As Byte

    For i = 14 To 17
        Text1(i).Text = ComprobarCero(Text1(i).Text)
    Next i
End Sub



Private Sub CargaCombo()
Dim cad As String
Dim i As Byte
Dim RS As ADODB.Recordset

    On Error GoTo ErrCarga
    
    Combo1(0).Clear
    
    Combo1(0).AddItem "Transportista"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Comisionista"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    
    Exit Sub
    
ErrCarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar datos combo.", Err.Description
End Sub


