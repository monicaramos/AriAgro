VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmComHcoFacturas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hist�rico de Facturas Proveedores"
   ClientHeight    =   9105
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   13590
   Icon            =   "frmComHcoFacturas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   13590
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
      Left            =   11025
      TabIndex        =   111
      Top             =   285
      Width           =   1605
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   45
      TabIndex        =   109
      Top             =   90
      Width           =   2595
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   110
         Top             =   180
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
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
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
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
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   2700
      TabIndex        =   107
      Top             =   90
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   108
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
               Object.ToolTipText     =   "�ltimo"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6150
      Left            =   45
      TabIndex        =   38
      Tag             =   "Fecha Oferta|F|N|||scapre|fecentre|dd/mm/yyyy|N|"
      Top             =   2025
      Width           =   13425
      _ExtentX        =   23680
      _ExtentY        =   10848
      _Version        =   393216
      Style           =   1
      Tabs            =   2
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
      TabCaption(0)   =   "Datos b�sicos"
      TabPicture(0)   =   "frmComHcoFacturas.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Albaranes"
      TabPicture(1)   =   "frmComHcoFacturas.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(6)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(18)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "FrameObserva"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "DataGrid2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "DataGrid1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtAux(7)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtAux(6)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtAux(5)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtAux(4)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdObserva"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtAux(0)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtAux(1)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtAux(2)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtAux(3)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtAux3(0)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtAux3(1)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Text3(2)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Text3(3)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).ControlCount=   18
      Begin VB.Frame Frame2 
         Height          =   5640
         Index           =   1
         Left            =   135
         TabIndex        =   80
         Top             =   315
         Width           =   13125
         Begin VB.Frame FrameCliente 
            Caption         =   "Datos Proveedor"
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
            Height          =   2460
            Left            =   45
            TabIndex        =   96
            Top             =   135
            Width           =   12960
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
               Left            =   8160
               MaxLength       =   15
               TabIndex        =   13
               Tag             =   "Concepto|T|S|||scafpc|confacpr|||"
               Text            =   "Text1"
               Top             =   360
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
               Index           =   6
               Left            =   1125
               MaxLength       =   35
               TabIndex        =   9
               Tag             =   "Domicilio|T|N|||scafpc|domprove||N|"
               Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
               Top             =   870
               Width           =   4260
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
               Left            =   10710
               MaxLength       =   5
               TabIndex        =   16
               Tag             =   "Descuento General|N|N|0|99.90|scafpc|dtognral|#0.00|N|"
               Text            =   "Text1 7"
               Top             =   1425
               Width           =   525
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
               Left            =   9000
               MaxLength       =   5
               TabIndex        =   15
               Tag             =   "Descuento P.Pago|N|N|0|99.90|scafpc|dtoppago|#0.00|N|"
               Text            =   "Text1 7"
               Top             =   1425
               Width           =   525
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
               Index           =   10
               Left            =   8745
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   97
               Text            =   "Text2"
               Top             =   870
               Width           =   4005
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
               Left            =   8160
               MaxLength       =   3
               TabIndex        =   14
               Tag             =   "Forma de Pago|N|N|0|999|scafpc|codforpa|000|N|"
               Text            =   "Text1"
               Top             =   870
               Width           =   540
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
               Left            =   1125
               MaxLength       =   15
               TabIndex        =   7
               Tag             =   "NIF proveedor|T|N|||scafpc|nifprove||N|"
               Text            =   "123456789"
               Top             =   375
               Width           =   1245
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
               Left            =   3420
               MaxLength       =   20
               TabIndex        =   8
               Tag             =   "tel�fono proveedor|T|S|||scafpc|telprove||N|"
               Text            =   "12345678911234567899"
               Top             =   375
               Width           =   1965
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
               Left            =   1935
               MaxLength       =   30
               TabIndex        =   11
               Tag             =   "Poblaci�n|T|N|||scafpc|pobprove||N|"
               Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
               Top             =   1395
               Width           =   3450
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
               Left            =   1125
               MaxLength       =   6
               TabIndex        =   10
               Tag             =   "CPostal|T|N|||scafpc|codpobla||N|"
               Text            =   "Text15"
               Top             =   1395
               Width           =   810
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
               Left            =   1125
               MaxLength       =   30
               TabIndex        =   12
               Tag             =   "Provincia|T|N|||scafpc|proprove||N|"
               Text            =   "Text1 Text1 Text1 Text1 Text22"
               Top             =   1935
               Width           =   2670
            End
            Begin VB.Label Label1 
               Caption         =   "Concepto"
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
               Left            =   6675
               TabIndex        =   106
               Top             =   360
               Width           =   1095
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
               TabIndex        =   105
               Top             =   915
               Width           =   915
            End
            Begin VB.Image imgBuscar 
               Height          =   240
               Index           =   3
               Left            =   7875
               ToolTipText     =   "Buscar forma de pago"
               Top             =   900
               Width           =   240
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
               Left            =   9765
               TabIndex        =   104
               Top             =   1440
               Width           =   1140
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
               Left            =   8160
               TabIndex        =   103
               Top             =   1440
               Width           =   930
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
               Left            =   6675
               TabIndex        =   102
               Top             =   870
               Width           =   1215
            End
            Begin VB.Image imgBuscar 
               Height          =   240
               Index           =   1
               Left            =   855
               ToolTipText     =   "Buscar proveedor varios"
               Top             =   390
               Visible         =   0   'False
               Width           =   240
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
               TabIndex        =   101
               Top             =   375
               Width           =   615
            End
            Begin VB.Label Label1 
               Caption         =   "Tel�fono"
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
               Left            =   2445
               TabIndex        =   100
               Top             =   375
               Width           =   870
            End
            Begin VB.Label Label1 
               Caption         =   "Poblaci�n"
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
               TabIndex        =   99
               Top             =   1440
               Width           =   1005
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
               TabIndex        =   98
               Top             =   1980
               Width           =   915
            End
         End
         Begin VB.Frame FrameFactura 
            Height          =   2850
            Left            =   45
            TabIndex        =   81
            Top             =   2655
            Width           =   12960
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
               Left            =   240
               MaxLength       =   15
               TabIndex        =   17
               Tag             =   "Imp.Bruto|N|N|||scafpc|brutofac|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   480
               Width           =   1485
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
               Left            =   2160
               MaxLength       =   15
               TabIndex        =   18
               Tag             =   "Imp. Dto PP|N|N|||scafpc|impppago|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   480
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
               Index           =   16
               Left            =   3960
               MaxLength       =   15
               TabIndex        =   19
               Tag             =   "Imp. Dto Gn|N|N|||scafpc|impgnral|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   480
               Width           =   1365
            End
            Begin VB.TextBox Text1 
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
               Index           =   17
               Left            =   5760
               MaxLength       =   15
               TabIndex        =   82
               Text            =   "Text1 7"
               Top             =   495
               Width           =   1485
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
               Index           =   24
               Left            =   5760
               MaxLength       =   15
               TabIndex        =   22
               Tag             =   "Base Imponible 1|N|N|||scafpc|baseiva1|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   1350
               Width           =   1485
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
               Left            =   3960
               MaxLength       =   3
               TabIndex        =   20
               Tag             =   "Cod. IVA 1|N|S|0|999|scafpc|tipoiva1|000|N|"
               Text            =   "Text1 7"
               Top             =   1350
               Width           =   705
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
               Index           =   21
               Left            =   4950
               MaxLength       =   5
               TabIndex        =   21
               Tag             =   "% IVA 1|N|S|0|99.90|scafpc|porciva1|#0.00|N|"
               Text            =   "Text1 7"
               Top             =   1350
               Width           =   705
            End
            Begin VB.TextBox Text1 
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
               Index           =   27
               Left            =   7560
               MaxLength       =   15
               TabIndex        =   23
               Tag             =   "Importe IVA 1|N|N|||scafpc|impoiva1|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   1350
               Width           =   1485
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
               Index           =   25
               Left            =   5760
               MaxLength       =   15
               TabIndex        =   26
               Tag             =   "Base Imponible 2 |N|S|||scafpc|baseiva2|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   1800
               Width           =   1485
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
               Left            =   3960
               MaxLength       =   3
               TabIndex        =   24
               Tag             =   "Cod. IVA 2|N|S|0|999|scafpc|tipoiva2|000|N|"
               Text            =   "Text1 7"
               Top             =   1800
               Width           =   705
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
               Index           =   22
               Left            =   4950
               MaxLength       =   5
               TabIndex        =   25
               Tag             =   "& IVA 2|N|S|0|99.90|scafpc|porciva2|#0.00|N|"
               Text            =   "Text1 7"
               Top             =   1800
               Width           =   705
            End
            Begin VB.TextBox Text1 
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
               Index           =   28
               Left            =   7560
               MaxLength       =   15
               TabIndex        =   27
               Tag             =   "Importe IVA 2|N|S|||scafpc|impoiva2|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   1800
               Width           =   1485
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
               Left            =   5760
               MaxLength       =   15
               TabIndex        =   30
               Tag             =   "Base Imponible 3|N|S|||scafpc|baseiva3|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   2265
               Width           =   1485
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
               Index           =   20
               Left            =   3960
               MaxLength       =   3
               TabIndex        =   28
               Tag             =   "Cod. IVA 3|N|S|0|999|scafpc|tipoiva3|000|N|"
               Text            =   "Text1 7"
               Top             =   2265
               Width           =   705
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
               Left            =   4950
               MaxLength       =   5
               TabIndex        =   29
               Tag             =   "% IVA 3|N|S|0|99.90|scafpc|porciva3|#0.00|N|"
               Text            =   "Text1 7"
               Top             =   2265
               Width           =   705
            End
            Begin VB.TextBox Text1 
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
               Left            =   7560
               MaxLength       =   15
               TabIndex        =   31
               Tag             =   "Importe IVA 3|N|S|||scafpc|impoiva3|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   2265
               Width           =   1485
            End
            Begin VB.TextBox Text1 
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
               Left            =   9360
               MaxLength       =   15
               TabIndex        =   32
               Tag             =   "Total Factura|N|N|||scafpc|totalfac|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   2280
               Width           =   1605
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
               Height          =   255
               Index           =   10
               Left            =   240
               TabIndex        =   95
               Top             =   195
               Width           =   1485
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
               Height          =   255
               Index           =   11
               Left            =   2145
               TabIndex        =   94
               Top             =   195
               Width           =   1215
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
               Height          =   255
               Index           =   12
               Left            =   3945
               TabIndex        =   93
               Top             =   195
               Width           =   1215
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
               Height          =   255
               Index           =   14
               Left            =   5790
               TabIndex        =   92
               Top             =   195
               Width           =   1575
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
               TabIndex        =   91
               Top             =   495
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
               TabIndex        =   90
               Top             =   495
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
               TabIndex        =   89
               Top             =   495
               Width           =   135
            End
            Begin VB.Label Label1 
               Caption         =   "Importe IVA"
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
               Left            =   7590
               TabIndex        =   88
               Top             =   1050
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
               TabIndex        =   87
               Top             =   1050
               Width           =   135
            End
            Begin VB.Line Line1 
               X1              =   3960
               X2              =   7320
               Y1              =   945
               Y2              =   945
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
               TabIndex        =   86
               Top             =   2265
               Width           =   135
            End
            Begin VB.Label Label1 
               Caption         =   "TOTAL FACTURA"
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
               Left            =   9330
               TabIndex        =   85
               Top             =   1980
               Width           =   1800
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
               Height          =   255
               Index           =   41
               Left            =   4950
               TabIndex        =   84
               Top             =   1050
               Width           =   675
            End
            Begin VB.Label Label1 
               Caption         =   "Cod.IVA"
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
               Left            =   3960
               TabIndex        =   83
               Top             =   1050
               Width           =   915
            End
         End
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
         Index           =   3
         Left            =   -68805
         MaxLength       =   10
         TabIndex        =   77
         Tag             =   "Fecha Pedido|F|S|||scafpc|fecpedpr|dd/mm/yyyy|N|"
         Top             =   2040
         Width           =   1500
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
         Index           =   2
         Left            =   -70095
         MaxLength       =   7
         TabIndex        =   76
         Tag             =   "N� Pedido|N|S|||scafpa|numpedpr|0000000|N|"
         Text            =   "Text1 7"
         Top             =   2040
         Width           =   1065
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
         Height          =   330
         Index           =   1
         Left            =   -72960
         MaxLength       =   30
         TabIndex        =   60
         Tag             =   "Fecha Albaran|F|N|||scafpa|fechaalb|dd/mm/yyyy|N|"
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
         Left            =   -73920
         MaxLength       =   15
         TabIndex        =   59
         Tag             =   "N� Albaran|T|N|||scafpa|numalbar|0|N|"
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
         Left            =   -71760
         MaxLength       =   12
         TabIndex        =   52
         Tag             =   "Cantidad|N|N|0||slifac|cantidad|#,###,###,##0.00|N|"
         Text            =   "cantidad"
         Top             =   4320
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
         Height          =   330
         Index           =   2
         Left            =   -72840
         MaxLength       =   12
         TabIndex        =   51
         Tag             =   "Nombre Art.|T|N|||slifac|nomartic||N|"
         Text            =   "nomartic"
         Top             =   4320
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
         Height          =   330
         Index           =   1
         Left            =   -73680
         MaxLength       =   12
         TabIndex        =   50
         Tag             =   "Art.|T|N|||slifac|codartic||N|"
         Text            =   "codartic"
         Top             =   4320
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
         Height          =   330
         Index           =   0
         Left            =   -74640
         MaxLength       =   12
         TabIndex        =   49
         Tag             =   "Almacen|N|N|0|999|slifac|codalmac|000|N|"
         Text            =   "almacen"
         Top             =   4320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdObserva 
         Height          =   375
         Left            =   -70920
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   520
         Width           =   375
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
         Left            =   -70920
         MaxLength       =   12
         TabIndex        =   53
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
         Left            =   -69240
         MaxLength       =   5
         TabIndex        =   54
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
         Left            =   -68520
         MaxLength       =   30
         TabIndex        =   55
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
         Left            =   -67920
         MaxLength       =   12
         TabIndex        =   58
         Tag             =   "Importe|N|N|0||slifac|importel|#,###,###,##0.00|N|"
         Text            =   "Importe"
         Top             =   4320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmComHcoFacturas.frx":0044
         Height          =   3150
         Left            =   -74775
         TabIndex        =   40
         Top             =   2715
         Width           =   12930
         _ExtentX        =   22807
         _ExtentY        =   5556
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
         Bindings        =   "frmComHcoFacturas.frx":0059
         Height          =   1905
         Left            =   -74760
         TabIndex        =   41
         Top             =   525
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   3360
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
         Height          =   3270
         Left            =   -74760
         TabIndex        =   42
         Tag             =   "Observaci�n 4|T|S|||scafac1|observa4||N|"
         Top             =   2610
         Width           =   12745
         Begin VB.TextBox Text3 
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
            Left            =   720
            MaxLength       =   80
            TabIndex        =   47
            Tag             =   "Observaci�n 5|T|S|||scafpa|observa5||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   2325
            Width           =   10695
         End
         Begin VB.TextBox Text3 
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
            Left            =   720
            MaxLength       =   80
            TabIndex        =   46
            Tag             =   "Observaci�n 4|T|S|||scafpa|observa4||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   1860
            Width           =   10695
         End
         Begin VB.TextBox Text3 
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
            Left            =   720
            MaxLength       =   80
            TabIndex        =   45
            Tag             =   "Observaci�n 3|T|S|||scafpa|observa3||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   1395
            Width           =   10695
         End
         Begin VB.TextBox Text3 
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
            Left            =   720
            MaxLength       =   80
            TabIndex        =   44
            Tag             =   "Observaci�n 2|T|S|||scafpa|observa2||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   930
            Width           =   10695
         End
         Begin VB.TextBox Text3 
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
            Left            =   720
            MaxLength       =   80
            TabIndex        =   43
            Tag             =   "Observaci�n 1|T|S|||scafpa|observa1||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Text1 Te"
            Top             =   465
            Width           =   10695
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Pedido"
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
         Left            =   -68805
         TabIndex        =   79
         Top             =   1755
         Width           =   1410
      End
      Begin VB.Label Label1 
         Caption         =   "N� Pedido"
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
         Left            =   -70095
         TabIndex        =   78
         Top             =   1755
         Width           =   1050
      End
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   4485
      MaxLength       =   30
      TabIndex        =   73
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
      TabIndex        =   72
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
      TabIndex        =   71
      Text            =   "Text1"
      Top             =   2025
      Width           =   660
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   0
      Left            =   5145
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   70
      Text            =   "Text2"
      Top             =   2025
      Width           =   3525
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Index           =   0
      Left            =   45
      TabIndex        =   61
      Top             =   900
      Width           =   13455
      Begin VB.CheckBox Check1 
         Caption         =   "Es Rectificativa"
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
         Left            =   4905
         TabIndex        =   4
         Tag             =   "Es rectificativa|N|N|0|1|scafpc|esrectificativa||N|"
         Top             =   540
         Width           =   2145
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
         Index           =   31
         Left            =   3300
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Recepci�n|F|N|||scafpc|fecrecep|dd/mm/yyyy|N|"
         Top             =   360
         Width           =   1350
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
         Left            =   9255
         MaxLength       =   40
         TabIndex        =   6
         Tag             =   "Nombre Proveedor|T|N|||scafpc|nomprove||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   330
         Width           =   3885
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
         Left            =   8370
         MaxLength       =   6
         TabIndex        =   5
         Tag             =   "Cod. Proveedor|N|N|0|999999|scafpc|codprove|000000|S|"
         Text            =   "Text1"
         Top             =   330
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
         Left            =   1830
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Factura|F|N|||scafpc|fecfactu|dd/mm/yyyy|S|"
         Top             =   360
         Width           =   1350
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
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "N� Factura|T|N|||scafpc|numfactu||S|"
         Text            =   "Text1 7"
         Top             =   360
         Width           =   1485
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Contabilizado"
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
         Left            =   4905
         TabIndex        =   3
         Tag             =   "Contabilizado|N|N|0|1|scafpc|intconta||N|"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "F. Recepci�n"
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
         Left            =   3315
         TabIndex        =   65
         Top             =   120
         Width           =   1380
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
         Left            =   7020
         TabIndex        =   64
         Top             =   330
         Width           =   1050
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   8100
         ToolTipText     =   "Buscar proveedor"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F. Factura"
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
         Left            =   1830
         TabIndex        =   63
         Top             =   120
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "N� Factura"
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
         Left            =   270
         TabIndex        =   62
         Top             =   135
         Width           =   1260
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   13
      Left            =   3555
      MaxLength       =   4
      TabIndex        =   68
      Text            =   "Text1"
      Top             =   1035
      Width           =   540
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   13
      Left            =   4140
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   67
      Text            =   "Text2"
      Top             =   1035
      Width           =   3285
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   17
      Left            =   7500
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   57
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
      Left            =   2355
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   56
      Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
      Top             =   8535
      Visible         =   0   'False
      Width           =   7155
   End
   Begin VB.Frame Frame1 
      Height          =   570
      Index           =   0
      Left            =   75
      TabIndex        =   36
      Top             =   8325
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
         TabIndex        =   37
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
      Left            =   12420
      TabIndex        =   34
      Top             =   8460
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
      Left            =   11250
      TabIndex        =   33
      Top             =   8460
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
      Left            =   12420
      TabIndex        =   35
      Top             =   8460
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
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   13065
      TabIndex        =   112
      Top             =   225
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
      Enabled         =   0   'False
      Height          =   240
      Index           =   2
      Left            =   4455
      Picture         =   "frmComHcoFacturas.frx":006E
      ToolTipText     =   "Buscar poblaci�n"
      Top             =   1080
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   6
      Left            =   4125
      Picture         =   "frmComHcoFacturas.frx":0170
      ToolTipText     =   "Buscar trabajador"
      Top             =   2220
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Trabajador Pedido"
      Height          =   255
      Index           =   9
      Left            =   2565
      TabIndex        =   75
      Top             =   2220
      Width           =   1425
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   5
      Left            =   4140
      Picture         =   "frmComHcoFacturas.frx":0272
      ToolTipText     =   "Buscar trabajador"
      Top             =   2025
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Trabajador Albaran"
      Height          =   255
      Index           =   21
      Left            =   2565
      TabIndex        =   74
      Top             =   2115
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Trabajador"
      Height          =   255
      Index           =   1
      Left            =   2340
      TabIndex        =   69
      Top             =   1035
      Width           =   855
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   4
      Left            =   3270
      Picture         =   "frmComHcoFacturas.frx":0374
      ToolTipText     =   "Buscar trabajador"
      Top             =   1065
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "N� Lote"
      Height          =   255
      Index           =   3
      Left            =   7500
      TabIndex        =   66
      Top             =   5250
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Ampliaci�n L�nea"
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
      Left            =   2355
      TabIndex        =   39
      Top             =   8265
      Visible         =   0   'False
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
         HelpContextID   =   2
         Shortcut        =   ^L
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Shortcut        =   ^I
         Visible         =   0   'False
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
Attribute VB_Name = "frmComHcoFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)


Public Factura As String ' cuando venimos de documentos de proveedores


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

Private WithEvents frmProv As frmManProve  'Form Mto Proveedores
Attribute frmProv.VB_VarHelpID = -1
Private WithEvents frmPV As frmComProveV  'Form Mto Proveedores Varios
Attribute frmPV.VB_VarHelpID = -1
Private WithEvents frmFP As frmManFpago 'Form Mto Formas de Pago
Attribute frmFP.VB_VarHelpID = -1
'--monica
'Private WithEvents frmT As frmAdmTrabajadores  'Form Mto Trabajadores
Private WithEvents frmFac As frmBasico2
Attribute frmFac.VB_VarHelpID = -1



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
'1.- A�adir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

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
'Variable que indica el n�mero del Boton  PrimerRegistro en la Toolbar1

Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos
Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal

Private BuscaChekc As String


Private Sub Check1_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "check1(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "check1(" & Index & ")|"
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
'                    'Comprobar si el Articulo tiene control de N� de Serie
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
        
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()

    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        LimpiarCampos

        LimpiarDataGrids
        CadenaConsulta = "Select scafpc.* "
        CadenaConsulta = CadenaConsulta & "from " & NombreTabla & Ordenacion
        

        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index, 350
    PonerCampos
End Sub


Private Sub BotonModificar()
Dim DeVarios As Boolean

    'solo se puede modificar la factura si no esta contabilizada
    '++monica:a�adida la condicion de solo si hay contabilidad
    If vParamAplic.NumeroConta <> 0 Then
'        If Me.Check1.Value = 1 Then
'            TerminaBloquear
'            Exit Sub
'        End If
        
        If FactContabilizada Then
            TerminaBloquear
            Exit Sub
        End If
    End If
    
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    PonerFocoChk Me.Check1(0)
        
    'Si es proveedor de Varios no se pueden modificar sus datos
    DeVarios = EsProveedorVarios(Text1(2).Text)
    BloquearDatosProve (DeVarios)
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
    vWhere = vWhere & " AND numalbar='" & Data3.Recordset.Fields!NumAlbar & "'"
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
    
    'A�adiremos el boton de aceptar y demas objetos para insertar
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
'    End Select
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Facturas (scafac)
' y los registros correspondientes de las tablas cab. albaranes (scafac1)
' y las lineas de la factura (slifac)
Dim Cad As String
Dim NumPedElim As Long
On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    'solo se puede eliminar si no esta en la contabilidad
    If Me.Check1(0).Value = 1 Then Exit Sub
    
'    'solo se puede modificar la factura si no esta contabilizada
'    If FactContabilizada Then Exit Sub
    
    Cad = "Cabecera de Facturas." & vbCrLf
    Cad = Cad & "-----------------------------------" & vbCrLf & vbCrLf
    Cad = Cad & "Va a eliminar la Factura:            "
    Cad = Cad & vbCrLf & "Proveedor:  " & Text1(2).Text & " - " & Text1(3).Text
    Cad = Cad & vbCrLf & "N� Fact.:  " & Text1(0).Text
    Cad = Cad & vbCrLf & "Fecha:  " & Format(Text1(1).Text, "dd/mm/yyyy")

    Cad = Cad & vbCrLf & vbCrLf & " �Desea Eliminarla? "
    
    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        NumPedElim = Data1.Recordset.Fields(1).Value
        
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
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
        Screen.MousePointer = vbDefault
        If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Albaran", Err.Description
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
Dim Cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        DataGrid2.Enabled = True
        If Not Data1.Recordset.EOF Then _
            Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    Else 'Se llama desde alg�n Prismatico de otro Form al Mantenimiento de Trabajadores
        If Data1.Recordset.EOF Then
            MsgBox "Ning�n registro devuelto.", vbExclamation
            Exit Sub
        End If
        Cad = Data1.Recordset.Fields(0) & "|"
        Cad = Cad & Data1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(Cad)
        Unload Me
    End If
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo Error1

    If Not Data2.Recordset.EOF Then
        If ModificaLineas <> 1 Then
            Text2(16).Text = DBLet(Data2.Recordset.Fields!ampliaci)
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

        Text3(2).Text = DBLet(Data3.Recordset.Fields!numpedpr, "N")
        If Text3(2).Text <> "0" Then
            FormateaCampo Text3(2)
        Else
            Text3(2).Text = ""
        End If
        Text3(3).Text = DBLet(Data3.Recordset.Fields!fecpedpr, "F")
        
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
            Text3(i).Text = ""
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
'    btnPrimero = 15
'    With Me.Toolbar1
'        .HotImageList = frmPpal.imgListComun_OM
'        .DisabledImageList = frmPpal.imgListComun_BN
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 1   'Bot�n Buscar
'        .Buttons(2).Image = 2   'Bot�n Todos
'        .Buttons(5).Image = 4   'Modificar
'        .Buttons(6).Image = 5   'Borrar
'        .Buttons(9).Image = 15 'Mto Lineas Ofertas
'        .Buttons(10).Image = 10 'Imprimir
'        .Buttons(12).Image = 11  'Salir
'        .Buttons(btnPrimero).Image = 6  'Primero
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
'        .Buttons(btnPrimero + 3).Image = 9 '�ltimo
'    End With

    Me.SSTab1.Tab = 0
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'l'1 i el 2 son separadors
        .Buttons(5).Image = 1   'Buscar
        .Buttons(6).Image = 2   'Totss
        'el 5 i el 6 son separadors
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Borrar
        'el 10  son separadors
'        .Buttons(8).Image = 10  'Imprimir
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
    NombreTabla = "scafpc"
    NomTablaLineas = "slifpc" 'Tabla lineas de Facturacion
    Ordenacion = " ORDER BY scafpc.fecrecep desc ,scafpc.codprove, scafpc.numfactu "
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    CadenaConsulta = "Select * from " & NombreTabla
    
    If hcoCodMovim <> "" Then
        'Se llama desde Dobleclick en frmAlmMovimArticulos
'        CadenaConsulta = CadenaConsulta & " WHERE numalbar='" & hcoCodMovim & "' AND fechaalb= """ & Format(hcoFechaMovim, "yyyy-mm-dd") & """"
'        CadenaConsulta = CadenaConsulta & " AND codprove=" & hcoCodProve
        If Factura <> "" Then
            CadenaConsulta = CadenaConsulta & " WHERE codprove = " & hcoCodProve & " and numfactu = '" & hcoCodMovim & "' and "
            CadenaConsulta = CadenaConsulta & " fecfactu = '" & Format(hcoFechaMovim, "yyyy-mm-dd") & "'"
        Else
            CadenaConsulta = CadenaConsulta & ObtenerSelFactura
        End If
    Else
        CadenaConsulta = CadenaConsulta & " where numfactu=-1"
    End If
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
    If hcoCodMovim = "" Then
        If DatosADevolverBusqueda = "" Then
            PonerModo 0
        Else
            PonerModo 1
        End If
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        PrimeraVez = False
    Else
         PonerModo 0
    End If
End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.Check1(0).Value = 0
    Me.Check1(1).Value = 0
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
        CadenaConsulta = CadenaConsulta & " GROUP BY scafpc.codprove, scafpc.numfactu, scafpc.fecfactu "
        CadenaConsulta = CadenaConsulta & " " & Ordenacion
        PonerCadenaBusqueda
'        Text1(0).Text = RecuperaValor(CadenaDevuelta, 2)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmFac_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String
Dim Aux As String
      
    If CadenaSeleccion <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        CadB = "numfactu = '" & RecuperaValor(CadenaSeleccion, 1) & "' and fecfactu = " & DBSet(RecuperaValor(CadenaSeleccion, 2), "F") & " and codprove = " & RecuperaValor(CadenaSeleccion, 3)
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
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
Dim indice As Byte
    indice = 10
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Forma Pago
    Text2(10).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub


Private Sub frmProv_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Proveedores
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1)  'Cod Prove
End Sub

Private Sub frmPV_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Proveedores Varios
Dim indice As Byte

    indice = 4
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'NIF
    Text1(indice - 1).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Clien
    PonerDatosProveVario (Text1(indice).Text)
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
Dim indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. Proveedor
            PonerFoco Text1(2)
            Set frmProv = New frmManProve
            frmProv.DatosADevolverBusqueda = "0|"
            frmProv.Show vbModal
            Set frmProv = Nothing
            indice = 2
            PonerFoco Text1(indice)
            
        Case 1 'NIF para proveedor de Varios
            Set frmPV = New frmComProveV
            frmPV.DatosADevolverBusqueda = "0|"
            frmPV.Show vbModal
            Set frmPV = Nothing
            indice = 7
            PonerFoco Text1(indice)
            
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
            indice = 10
            PonerFoco Text1(indice)
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
Dim Sql As String
On Error GoTo EBloqueaAlb

    BloqueaAlbxFac = False
    'bloquear cabecera albaranes x factura
    Sql = "select * FROM scafpa "
    Sql = Sql & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute Sql, , adCmdText
    BloqueaAlbxFac = True

EBloqueaAlb:
    If Err.Number <> 0 Then BloqueaAlbxFac = False
End Function


Private Function BloqueaLineasFac() As Boolean
'bloquea TODAS las lineas de la factura
Dim Sql As String
    
    On Error GoTo EBloqueaLin

    BloqueaLineasFac = False
    
    'bloquear cabecera albaranes x factura
    Sql = "select * FROM slifpc "
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

        Case 2 'Cod. prove
            If Modo = 1 Then 'Modo=1 Busqueda
                Text1(Index + 1).Text = PonerNombreDeCod(Text1(Index), "proveedor", "nomprove")
            Else
                PonerDatosProveedor (Text1(Index).Text)
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

    CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select " & NombreTabla & ".* from " & NombreTabla & " LEFT OUTER JOIN scafpa ON " & NombreTabla & ".codprove=scafpa.codprove AND " & NombreTabla & ".numfactu=scafpa.numfactu AND " & NombreTabla & ".fecfactu=scafpa.fecfactu "
        CadenaConsulta = CadenaConsulta & " WHERE " & CadB
        CadenaConsulta = CadenaConsulta & " GROUP BY scafpc.codprove, scafpc.numfactu, scafpc.fecfactu "
        CadenaConsulta = CadenaConsulta & " " & Ordenacion
        
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)

    Set frmFac = New frmBasico2
    
    AyudaFacturasCompraPrev frmFac, , CadB
    
    Set frmFac = Nothing


End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass

    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        If Modo = 1 Then PonerFoco Text1(0)
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
'Carga el grid de los AlbaranesxFactura, es decir, la tabla scafpc de la factura seleccionada
On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass
    
    'Datos de la tabla albaranes x factura: scafpa
    CargaGrid DataGrid2, Data3, True
   
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
    PonerCamposForma Me, Data1
    
    'Poner la base imponible (impbruto - dtoppago - dtognral
    BrutoFac = CSng(Text1(14).Text) - CSng(Text1(15).Text) - CSng(Text1(16).Text)
    Text1(17).Text = Format(BrutoFac, FormatoImporte)
    
    'poner descripcion campos
    Text2(10).Text = PonerNombreDeCod(Text1(10), "forpago", "nomforpa")
'--monica
'    Text2(13).Text = PonerNombreDeCod(Text1(13), conAri, "straba", "nomtraba", "codtraba")
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Ofertas
    
'++monica
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu seg�n modo
    PonerOpcionesMenu 'Activar opciones de menu seg�n nivel de permisos del usuario
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
Dim i As Byte, NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo


    BuscaChekc = ""

    'Actualiza Iconos Insertar,Modificar,Eliminar
    '## No tiene el boton modificar y no utiliza la funcion general
    ActualizarToolbar Modo, Kmodo
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Or (Factura <> "") Then
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
    'Si estamos en Insertar adem�s limpia los campos Text1
    'si estamos en modificar bloquea las compos que son clave primaria
    BloquearText1 Me, Modo
    '---- laura 24/10/2006: si ponemos las claves de la tabla con ON UPDATE CASCADE
    'podemos permitir modificar la fecha de la factura que es clave primaria
'    If Modo = 4 Then BloquearTxt Text1(1), False
    
    
    Me.Check1(0).Enabled = (Modo = 1 Or Modo = 3 Or Modo = 4)
    Me.Check1(1).Enabled = (Modo = 1 Or Modo = 3 Or Modo = 4)
    
    b = (Modo <> 1)
    'Campos N� Factura bloqueado y en azul
    BloquearTxt Text1(0), b, True
    BloquearTxt Text1(3), b 'referencia
    
    'Importes siempre bloqueados
    For i = 14 To 30
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
       
    'Poner el tama�o de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu seg�n modo
    PonerOpcionesMenu 'Activar opciones de menu seg�n nivel de permisos del usuario
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub DesplazamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
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
    If Index = 17 And KeyAscii = 13 Then 'campo n� de lote y ENTER
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
        Case 5  'Buscar
            mnBuscar_Click
        Case 6  'Todos
            BotonVerTodos

        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
'        Case 9  'Lineas
'            mnLineas_Click
'        Case 10 'Imprimir Albaran
'            mnImprimir_Click
'        Case 12    'Salir
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
    
    
Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Albaran: slialb
Dim Sql As String
Dim vWhere As String
Dim b As Boolean

    On Error GoTo eModificarLinea

    ModificarLinea = False
    If Data2.Recordset.EOF Then Exit Function
    
    vWhere = ObtenerWhereCP(True)
    vWhere = vWhere & " AND numalbar='" & Data3.Recordset.Fields!NumAlbar & "'"
    vWhere = vWhere & " AND numlinea=" & Data2.Recordset.Fields!NumLinea
    
    If DatosOkLinea() Then
        Sql = "UPDATE " & NomTablaLineas & " SET "
        Sql = Sql & " ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
        Sql = Sql & "precioar= " & DBSet(txtAux(4).Text, "N") & ", "
        Sql = Sql & "dtoline1= " & DBSet(txtAux(5).Text, "N") & ", dtoline2= " & DBSet(txtAux(6).Text, "N") & ", "
        Sql = Sql & "importel= " & DBSet(txtAux(7).Text, "N")
        Sql = Sql & ", numlotes=" & DBSet(Text2(17).Text, "T")
        Sql = Sql & vWhere
    End If
    
    If Sql <> "" Then
        'actualizar la factura y vencimientos
        b = ModificarFactura(Sql)
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
        Me.lblIndicador.Caption = "L�neas " & TituloLinea
        PonerFocoBtn Me.cmdRegresar
    End If
    'Habilitar las opciones correctas del menu segun Modo
'    PonerModoOpcionesMenu (Modo)
'    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
    DataGrid2.Enabled = Not b
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim b As Boolean
Dim Opcion As Byte
Dim Sql As String

    On Error GoTo ECargaGRid

'    b = DataGrid1.Enabled

    If vDataGrid.Name = "DataGrid1" Then
        Opcion = 1
    Else
        Opcion = 2
    End If
    
    Sql = MontaSQLCarga(enlaza, Opcion)
    CargaGridGnral vDataGrid, vData, Sql, PrimeraVez
    
    CargaGrid2 vDataGrid, vData
    vDataGrid.ScrollBars = dbgAutomatic
    
    b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
    vDataGrid.Enabled = Not b
    PrimeraVez = False
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
            'SQL = "SELECT codtipom, numfactu, fecfactu, numalbar, numlinea,
            'codalmac, codartic, nomartic, ampliaci, cantidad, precioar, origpre, dtoline1, dtoline2, importel "
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux(0)|T|Alm.|520|;S|txtAux(1)|T|Art�culo|1750|;S|txtAux(2)|T|Nombre|4150|;"
            tots = tots & "N||||0|;S|txtAux(3)|T|Cantidad|1450|;S|txtAux(4)|T|Precio|1500|;S|txtAux(5)|T|Dto 1|650|;S|txtAux(6)|T|Dto 2|650|;S|txtAux(7)|T|Importe|1650|;" 'N||||0|;"
            arregla tots, DataGrid1, Me, 350
            DataGrid1.Columns(9).Alignment = dbgRight
            DataGrid1.Columns(10).Alignment = dbgRight
            DataGrid1.Columns(12).Alignment = dbgRight
            DataGrid1.Columns(13).Alignment = dbgRight
                       
         Case "DataGrid2" 'albaranes x articulo
            'SQL = "SELECT codtipom,numfactu,fecfactu,codtipoa,numalbar, fechaalb,"
            'numpedcl,fecpedcl,sementre,numofert,fecofert, referenc, codenvio,codtraba, codtrab1, codtrab2,observa1,observa2,observa3,observa4,observa5  "
            tots = "N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux3(0)|T|Albaran|1500|;S|txtAux3(1)|T|Fecha|1500|;"
            tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;" 'N||||0|;N||||0|;"
            tots = tots & "N||||0|;"
            arregla tots, DataGrid2, Me, 350
        
            DataGrid2_RowColChange 1, 1
    End Select
    
    vDataGrid.HoldFields
    Exit Sub
ECargaGRid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
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


Private Sub BotonMtoLineas(numTab As Integer, Cad As String)
    Me.SSTab1.Tab = numTab
    
    If Me.DataGrid1.visible Then 'Lineas de Albaranes
        If Me.Data2.Recordset.RecordCount < 1 Then
            MsgBox "La factura no tiene lineas.", vbInformation
            Exit Sub
        End If
        TituloLinea = Cad
        
        ModificaLineas = 0
        PonerModo 5
        PonerBotonCabecera True
    End If
End Sub


Private Function Eliminar() As Boolean
Dim Sql As String
Dim cta As String
Dim b As Boolean

    On Error GoTo FinEliminar

        b = False
        Eliminar = False
        If Data1.Recordset.EOF Then Exit Function
        
        conn.BeginTrans
        ConnConta.BeginTrans
        
        'Eliminar en la tabla pagos de la Contabilidad: spagop
        '------------------------------------------------
        
        If vParamAplic.ContabilidadNueva Then
            cta = DevuelveDesdeBDNew(cAgro, "proveedor", "codmacta", "codprove", Text1(2).Text, "N")
            Sql = " codmacta='" & cta & "' AND numfactu='" & Data1.Recordset.Fields!NumFactu & "'"
            Sql = Sql & " AND fecfactu='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
            ConnConta.Execute "Delete from pagos WHERE " & Sql
        
        Else
            cta = DevuelveDesdeBDNew(cAgro, "proveedor", "codmacta", "codprove", Text1(2).Text, "N")
            Sql = " ctaprove='" & cta & "' AND numfactu='" & Data1.Recordset.Fields!NumFactu & "'"
            Sql = Sql & " AND fecfactu='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
            ConnConta.Execute "Delete from spagop WHERE " & Sql
        End If
        b = True
        
        
        'Eliminar en tablas de factura de Ariges: scafpc, scafpa, slifpc
        '---------------------------------------------------------------
        If b Then
            Sql = " " & ObtenerWhereCP(True)
        
            'Lineas de facturas (slifpc)
            conn.Execute "Delete from " & NomTablaLineas & Sql
        
            'Lineas de cabeceras de albaranes de la factura
            conn.Execute "Delete from scafpa " & Sql
            
            'Cabecera de facturas (scafpc)
            conn.Execute "Delete from " & NombreTabla & Sql
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
        conn.RollbackTrans
        ConnConta.RollbackTrans
    Else
        conn.CommitTrans
        ConnConta.CommitTrans
    End If
    Eliminar = b
End Function


Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ning�n registro
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
Dim Sql As String
On Error Resume Next
    Sql = "codprove= " & Text1(2).Text & " and numfactu= '" & Text1(0).Text & "' and fecfactu='" & Format(Text1(1).Text, FormatoFecha) & "' "
    If conWhere Then Sql = " WHERE " & Sql
    ObtenerWhereCP = Sql
End Function


Private Function MontaSQLCarga(enlaza As Boolean, Opcion As Byte) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Bas�ndose en la informaci�n proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim Sql As String
    
    If Opcion = 1 Then
        Sql = "SELECT codprove, numfactu, fecfactu, numalbar, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel "
        Sql = Sql & " FROM slifpc " 'lineas de factura
    ElseIf Opcion = 2 Then
        Sql = "SELECT codprove,numfactu,fecfactu,numalbar, fechaalb, numpedpr,fecpedpr,observa1,observa2,observa3,observa4,observa5  "
        Sql = Sql & " FROM scafpa " 'cabeceras albaranes de la factura
    End If
    
    If enlaza Then
        Sql = Sql & " " & ObtenerWhereCP(True)
        'lineas factura proveedor
        If Opcion = 1 Then Sql = Sql & " AND numalbar=" & DBSet(Data3.Recordset.Fields!NumAlbar, "T")
    Else
        Sql = Sql & " WHERE numfactu = -1"
    End If
    Sql = Sql & " ORDER BY codprove, numfactu, fecfactu,numalbar "
    If Opcion = 1 Then Sql = Sql & ", numlinea "
    MontaSQLCarga = Sql
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar seg�n el modo en que estemos
Dim b As Boolean

        b = ((Modo = 2) Or (Modo = 5 And ModificaLineas = 0)) And Me.Check1(0).Value = 0 And (Factura = "")
        'Modificar
        Toolbar1.Buttons(2).Enabled = b
        Me.mnModificar.Enabled = b
        'eliminar
        Toolbar1.Buttons(3).Enabled = b '(Modo = 2)
        Me.mnEliminar.Enabled = b '(Modo = 2)
            
'        b = (Modo = 2)
'        'Mantenimiento lineas
'        Toolbar1.Buttons(9).Enabled = b
'        Me.mnLineas.Enabled = b
        'Imprimir
'        Toolbar1.Buttons(10).Enabled = b
'        Me.mnImprimir.Enabled = b
        
        b = ((Modo >= 3) Or Modo = 1)
        'Buscar
        Toolbar1.Buttons(5).Enabled = Not b And (Factura = "")
        Me.mnBuscar.Enabled = Not b And (Factura = "")
        'Ver Todos
        Toolbar1.Buttons(6).Enabled = Not b And (Factura = "")
        Me.mnVerTodos.Enabled = Not b And (Factura = "")
End Sub


Private Sub PonerDatosProveedor(codProve As String, Optional nifProve As String)
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
                If CLng(Text1(2).Text) = CLng(Data1.Recordset!codProve) Then
                    Set vProve = Nothing
                    Exit Sub
                End If
            End If
        
            Text1(2).Text = vProve.Codigo
            FormateaCampo Text1(2)
            
            If (Modo = 3) Or (Modo = 4) Then
                Text1(3).Text = vProve.Nombre  'Nom prove
                Text1(6).Text = vProve.Domicilio
                Text1(7).Text = vProve.CPostal
                Text1(8).Text = vProve.Poblacion
                Text1(9).Text = vProve.Provincia
                Text1(4).Text = vProve.NIF
                Text1(5).Text = DBLet(vProve.TfnoAdmon, "T")
            End If
            
            Observaciones = DBLet(vProve.Observaciones)
            If Observaciones <> "" Then
                MsgBox Observaciones, vbInformation, "Observaciones del proveedor"
            End If
        End If
    Else
        LimpiarDatosProve
    End If
    Set vProve = Nothing

EPonerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Datos Proveedor", Err.Description
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


Private Sub LimpiarDatosProve()
Dim i As Byte

    For i = 3 To 9
        Text1(i).Text = ""
    Next i
End Sub
   

Private Function ModificaAlbxFac() As Boolean
Dim Sql As String
Dim b As Boolean
On Error GoTo EModificaAlb
    
    ModificaAlbxFac = False
    If Data3.Recordset.EOF Then Exit Function
    
    'comprobar datos OK de la scafac1
     b = CompForm(Me) 'Comprobar formato datos ok de la cabecera alb: opcion=2
    If Not b Then Exit Function
'--monica
'    SQL = "UPDATE scafpa SET codtrab2=" & DBSet(Text3(0).Text, "N", "S") & ", "
'    SQL = SQL & "codtrab1=" & DBSet(Text3(1).Text, "N", "S")
    If Me.FrameObserva.visible Then
        Sql = "UPDATE scafpa SET "
        Sql = Sql & " observa1=" & DBSet(Text3(4).Text, "T")
        Sql = Sql & ", observa2=" & DBSet(Text3(5).Text, "T")
        Sql = Sql & ", observa3=" & DBSet(Text3(6).Text, "T")
        Sql = Sql & ", observa4=" & DBSet(Text3(7).Text, "T")
        Sql = Sql & ", observa5=" & DBSet(Text3(8).Text, "T")
        Sql = Sql & ObtenerWhereCP(True)
        Sql = Sql & " AND numalbar=" & Data3.Recordset.Fields!NumAlbar
        conn.Execute Sql
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
'si se ha modificado la linea de slifac, a�adir a la transaccion la modificaci�n de la linea y recalcular
Dim bol As Boolean
Dim MenError As String
Dim Sql As String
Dim vFactu As CFacturaCom
On Error GoTo EModFact

    bol = False
    conn.BeginTrans
    '++monica:a�adida la condicion de solo si hay contabilidad
    If vParamAplic.NumeroConta <> 0 Then ConnConta.BeginTrans
    
    If sqlLineas <> "" Then
        'actualizar el importe de la linea modificada
        MenError = "Modificando lineas de Factura."
        conn.Execute sqlLineas
    End If
    
    'recalcular las bases imponibles x IVA
    MenError = "Recalcular importes IVA"
    bol = ActualizarDatosFactura
    
    If bol Then
        'modificamos la scafpc
        MenError = "Modificando cabecera de factura"
        bol = ModificaDesdeFormulario1(Me, 1)
        
        If bol Then
            'Si es proveedor de varios actualizar datos proveedor en tabla:sprvar
            MenError = "Modificando datos proveedor varios"
            bol = ActualizarProveVarios(Text1(2).Text, Text1(4).Text)
        End If
        
        If bol Then
            MenError = "Modificando albaranes de factura"
            'modificar la tabla: scafpa
            bol = ModificaAlbxFac
            '++monica:a�adida la condicion de solo si hay contabilidad
            If vParamAplic.NumeroConta <> 0 Then
                If bol Then 'si se ha modificado la factura
                    MenError = "Actualizando en Tesoreria"
                    'y eliminar de tesoreria conta.spagop los registros de la factura
                    
                    'antes de Eliminar en las tablas de la Contabilidad
                    Set vFactu = New CFacturaCom
                    bol = vFactu.LeerDatos(Text1(2).Text, Text1(0).Text, Text1(1).Text)
                    
                    If bol Then
                        'Eliminar de la spagop
                        If vParamAplic.ContabilidadNueva Then
                            Sql = " codmacta='" & vFactu.CtaProve & "' AND numfactu='" & Data1.Recordset.Fields!NumFactu & "'"
                            Sql = Sql & " AND fecfactu='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
                            ConnConta.Execute "Delete from pagos WHERE " & Sql
                        
                        Else
                            Sql = " ctaprove='" & vFactu.CtaProve & "' AND numfactu='" & Data1.Recordset.Fields!NumFactu & "'"
                            Sql = Sql & " AND fecfactu='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
                            ConnConta.Execute "Delete from spagop WHERE " & Sql
                        End If
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
        conn.CommitTrans
        '++monica:a�adida la condicion de solo si hay contabilidad
        If vParamAplic.NumeroConta <> 0 Then ConnConta.CommitTrans
        ModificarFactura = True
    Else
        conn.RollbackTrans
        '++monica:a�adida la condicion de solo si hay contabilidad
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
        cta = DevuelveDesdeBDNew(cAgro, "proveedor", "codmacta", "codprove", Text1(2).Text, "N")
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


Private Sub BloquearDatosProve(bol As Boolean)
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


Private Function ActualizarProveVarios(Prove As String, NIF As String) As Boolean
'Modifica los datos de la tabla de Proveedores Varios
Dim vProve As CProveedor
On Error GoTo EActualizarCV

    ActualizarProveVarios = False
    
    Set vProve = New CProveedor
    If EsProveedorVarios(Prove) Then
        vProve.NIF = NIF
        vProve.Nombre = Text1(3).Text
        vProve.Domicilio = Text1(6).Text
        vProve.CPostal = Text1(7).Text
        vProve.Poblacion = Text1(8).Text
        vProve.Provincia = Text1(9).Text
        vProve.TfnoAdmon = Text1(5).Text
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


Private Function ObtenerSelFactura() As String
'Cuando venimos desde dobleClick en Movimientos de Articulos para Albaranes ya
'Facturados, abrimos este form pero cargando los datos de la factura
'correspendiente al albaran que se selecciono
Dim Cad As String
Dim Rs As ADODB.Recordset
On Error Resume Next

    Cad = "SELECT codprove,numfactu,fecfactu FROM scafpa "
    Cad = Cad & " WHERE codprove=" & DBSet(hcoCodProve, "N") & " AND numalbar=" & DBSet(hcoCodMovim, "T")
    Cad = Cad & " AND fechaalb=" & DBSet(hcoFechaMovim, "F")

    Set Rs = New ADODB.Recordset
    Rs.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then 'where para la factura
        Cad = " WHERE codprove=" & Rs!codProve & " AND numfactu= '" & Rs!NumFactu & "' AND fecfactu=" & DBSet(Rs!FecFactu, "F")
    Else
        Cad = " where numfactu=-1"
    End If
    Rs.Close
    Set Rs = Nothing

    ObtenerSelFactura = Cad
End Function



Private Function ActualizarDatosFactura() As Boolean
Dim vFactu As CFacturaCom
Dim cadSel As String

    Set vFactu = New CFacturaCom
    cadSel = ObtenerWhereCP(False)
    cadSel = "slifpc." & cadSel
    vFactu.DtoPPago = CCur(Text1(11).Text)
    vFactu.DtoGnral = CCur(Text1(12).Text)
    If vFactu.CalcularDatosFactura(cadSel, "scafpa", "slifpc") Then
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
