VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmVtasFactSocios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturas Ventas Socio"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   13575
   Icon            =   "frmVtasFactSocios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmVtasFactSocios.frx":000C
   ScaleHeight     =   8700
   ScaleWidth      =   13575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   3375
      Left            =   90
      TabIndex        =   35
      Top             =   540
      Width           =   13305
      Begin VB.CheckBox Check1 
         Caption         =   "Pasa EDI"
         Height          =   195
         Index           =   2
         Left            =   5040
         TabIndex        =   120
         Tag             =   "Pasa Edicom|N|N|||facturassocio|pasedicom|0||"
         Top             =   2160
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   5220
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Fecha Factura|F|N|||facturassocio|fecfactu|dd/mm/yyyy|S|"
         Top             =   720
         Width           =   1065
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   225
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "Tipo Iva|N|N|||facturassocio|tipoivac||N|"
         Top             =   1890
         Width           =   1440
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   1305
         MaxLength       =   12
         TabIndex        =   1
         Tag             =   "Tipo Movimiento|T|N|||facturassocio|codtipom||S|"
         Top             =   720
         Width           =   765
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Contabilizado"
         Height          =   195
         Index           =   1
         Left            =   5040
         TabIndex        =   9
         Tag             =   "Contabilizado|N|N|||facturassocio|intconta|0||"
         Top             =   1620
         Width           =   1275
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Pasa Aridoc"
         Height          =   195
         Index           =   0
         Left            =   5040
         TabIndex        =   10
         Tag             =   "Pasa Aridoc|N|N|||facturassocio|pasaridoc|0||"
         Top             =   1890
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   1305
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "Forma Pago|N|N|0|999|facturassocio|codforpa|000||"
         Text            =   "Text1"
         Top             =   1170
         Width           =   780
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   44
         Text            =   "Text2"
         Top             =   1170
         Width           =   4125
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   8
         Left            =   2700
         MaxLength       =   5
         TabIndex        =   7
         Tag             =   "%Dto 2|N|S|0|100|facturassocio|dtocom2|##0.00||"
         Top             =   1890
         Width           =   945
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   7
         Left            =   1710
         MaxLength       =   5
         TabIndex        =   6
         Tag             =   "%Dto 1|N|S|0|100|facturassocio|dtocom1|##0.00||"
         Top             =   1890
         Width           =   945
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   3690
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "Importe Dto|N|S|||facturassocio|impdtoc|###,##0.00||"
         Text            =   "Text3"
         Top             =   1890
         Width           =   1140
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   39
         Text            =   "Text2"
         Top             =   270
         Width           =   4125
      End
      Begin VB.TextBox Text1 
         Height          =   690
         Index           =   2
         Left            =   225
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Tag             =   "Observaciones|T|S|||facturassocio|observac|||"
         Top             =   2520
         Width           =   6105
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   1305
         MaxLength       =   6
         TabIndex        =   0
         Tag             =   "Cod. Socio|N|N|0|999999|facturassocio|codsocio|000000||"
         Text            =   "Text1"
         Top             =   270
         Width           =   780
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Height          =   315
         Index           =   0
         Left            =   3060
         MaxLength       =   7
         TabIndex        =   2
         Tag             =   "Nº Factura|N|S|||facturassocio|numfactu|0000000|S|"
         Text            =   "Text1 7"
         Top             =   720
         Width           =   980
      End
      Begin VB.Frame FrameFactura 
         Height          =   3165
         Left            =   6705
         TabIndex        =   49
         Top             =   135
         Width           =   6450
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   31
            Left            =   4230
            MaxLength       =   5
            TabIndex        =   115
            Tag             =   "% REC 1|N|S|0|99.90|facturassocio|porcrec1|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   30
            Left            =   4815
            MaxLength       =   15
            TabIndex        =   114
            Tag             =   "Importe REC 1|N|S|0||facturassocio|imporec1|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   29
            Left            =   4230
            MaxLength       =   5
            TabIndex        =   113
            Tag             =   "% REC 2|N|S|0|99.90|facturassocio|porcrec2|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1875
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   28
            Left            =   4815
            MaxLength       =   15
            TabIndex        =   112
            Tag             =   "Importe REC 2|N|S|0||facturassocio|imporec2|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1875
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   27
            Left            =   4245
            MaxLength       =   5
            TabIndex        =   111
            Tag             =   "% REC 3|N|S|0|99.90|facturassocio|porcrec3|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   2205
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   26
            Left            =   4815
            MaxLength       =   15
            TabIndex        =   110
            Tag             =   "Importe REC 3|N|S|0||facturassocio|imporec3|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   2205
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
            Height          =   350
            Index           =   25
            Left            =   3960
            MaxLength       =   15
            TabIndex        =   27
            Tag             =   "Total Factura|N|S|0||facturassocio|totalfac|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   2685
            Width           =   2325
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   24
            Left            =   2730
            MaxLength       =   15
            TabIndex        =   26
            Tag             =   "Importe IVA 3|N|S|0||facturassocio|impoiva3|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   2205
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   22
            Left            =   2160
            MaxLength       =   5
            TabIndex        =   24
            Tag             =   "% IVA 3|N|S|0|99.90|facturassocio|porciva3|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   2205
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   23
            Left            =   630
            MaxLength       =   15
            TabIndex        =   25
            Tag             =   "Base Imponible 3|N|S|0||facturassocio|baseimp3|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   2205
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   20
            Left            =   2730
            MaxLength       =   15
            TabIndex        =   22
            Tag             =   "Importe IVA 2|N|S|0||facturassocio|impoiva2|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1875
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   18
            Left            =   2145
            MaxLength       =   5
            TabIndex        =   20
            Tag             =   "& IVA 2|N|S|0|99.90|facturassocio|porciva2|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1875
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   19
            Left            =   630
            MaxLength       =   15
            TabIndex        =   21
            Tag             =   "Base Imponible 2 |N|S|0||facturassocio|baseimp2|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1890
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   16
            Left            =   2730
            MaxLength       =   15
            TabIndex        =   18
            Tag             =   "Importe IVA 1|N|S|0||facturassocio|impoiva1|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   14
            Left            =   2145
            MaxLength       =   5
            TabIndex        =   16
            Tag             =   "% IVA 1|N|S|0|99.90|facturassocio|porciva1|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   15
            Left            =   630
            MaxLength       =   15
            TabIndex        =   17
            Tag             =   "Base Imponible 1|N|S|0||facturassocio|baseimp1|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   13
            Left            =   120
            MaxLength       =   5
            TabIndex        =   15
            Tag             =   "IVA 1|N|S|0|9|facturassocio|codiiva1|0|N|"
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   500
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   17
            Left            =   120
            MaxLength       =   5
            TabIndex        =   19
            Tag             =   "IVA 2|N|S|0|9|facturassocio|codiiva2|0|N|"
            Text            =   "Text1 7"
            Top             =   1875
            Width           =   500
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   21
            Left            =   120
            MaxLength       =   5
            TabIndex        =   23
            Tag             =   "IVA 3|N|S|0|9|facturassocio|codiiva3|0|N|"
            Text            =   "Text1 7"
            Top             =   2205
            Width           =   500
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   9
            Left            =   630
            MaxLength       =   15
            TabIndex        =   12
            Tag             =   "Bruto Factura|N|S|0||facturassocio|brutofac|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   540
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   10
            Left            =   2745
            MaxLength       =   15
            TabIndex        =   13
            Tag             =   "Impòrte Descuento|N|S|0||facturassocio|impordto|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   540
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   12
            Left            =   4815
            MaxLength       =   15
            TabIndex        =   14
            Text            =   "Text1 7"
            Top             =   540
            Width           =   1485
         End
         Begin VB.Label Label1 
            Caption         =   "Importe Recargo"
            Height          =   255
            Index           =   16
            Left            =   4815
            TabIndex        =   117
            Top             =   1350
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "% Rec"
            Height          =   255
            Index           =   15
            Left            =   4230
            TabIndex        =   116
            Top             =   1350
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "% IVA"
            Height          =   255
            Index           =   41
            Left            =   2145
            TabIndex        =   60
            Top             =   1350
            Width           =   495
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
            Left            =   2160
            TabIndex        =   59
            Top             =   2745
            Width           =   1530
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
            TabIndex        =   58
            Top             =   2160
            Width           =   135
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
            Left            =   5535
            TabIndex        =   57
            Top             =   1035
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Importe IVA"
            Height          =   255
            Index           =   33
            Left            =   2730
            TabIndex        =   56
            Top             =   1350
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   13
            Left            =   645
            TabIndex        =   55
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Cod."
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   54
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Bruto Factura"
            Height          =   255
            Index           =   5
            Left            =   630
            TabIndex        =   53
            Top             =   270
            Width           =   1110
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   7
            Left            =   4815
            TabIndex        =   52
            Top             =   270
            Width           =   1215
         End
         Begin VB.Line Line1 
            X1              =   4815
            X2              =   6315
            Y1              =   945
            Y2              =   945
         End
         Begin VB.Label Label1 
            Caption         =   "Imp.Descuento"
            Height          =   255
            Index           =   1
            Left            =   2745
            TabIndex        =   51
            Top             =   270
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
            Index           =   9
            Left            =   2430
            TabIndex        =   50
            Top             =   540
            Width           =   135
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Fact."
         Height          =   255
         Index           =   3
         Left            =   225
         TabIndex        =   48
         Top             =   765
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Iva"
         Height          =   255
         Index           =   14
         Left            =   225
         TabIndex        =   46
         Top             =   1620
         Width           =   810
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1035
         ToolTipText     =   "Buscar Destino"
         Top             =   1215
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F.Pago"
         Height          =   255
         Index           =   8
         Left            =   225
         TabIndex        =   45
         Top             =   1215
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "%Dto 2"
         Height          =   255
         Index           =   4
         Left            =   2700
         TabIndex        =   43
         Top             =   1665
         Width           =   1050
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Fac."
         Height          =   255
         Index           =   29
         Left            =   4095
         TabIndex        =   42
         Top             =   780
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "%Dto 1"
         Height          =   255
         Index           =   2
         Left            =   1710
         TabIndex        =   41
         Top             =   1665
         Width           =   1140
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   4950
         Picture         =   "frmVtasFactSocios.frx":0A0E
         ToolTipText     =   "Buscar fecha"
         Top             =   765
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Dtos"
         Height          =   255
         Index           =   6
         Left            =   3735
         TabIndex        =   40
         Top             =   1665
         Width           =   1155
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   1350
         ToolTipText     =   "Zoom descripción"
         Top             =   2250
         Width           =   240
      End
      Begin VB.Label Label29 
         Caption         =   "Observaciones"
         Height          =   255
         Left            =   225
         TabIndex        =   38
         Top             =   2250
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Socio"
         Height          =   255
         Index           =   0
         Left            =   225
         TabIndex        =   37
         Top             =   315
         Width           =   540
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1035
         ToolTipText     =   "Buscar Socio"
         Top             =   315
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Factura"
         Height          =   255
         Index           =   28
         Left            =   2205
         TabIndex        =   36
         Top             =   780
         Width           =   855
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4065
      Left            =   90
      TabIndex        =   47
      Top             =   3960
      Width           =   13320
      _ExtentX        =   23495
      _ExtentY        =   7170
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Variedades"
      TabPicture(0)   =   "frmVtasFactSocios.frx":0A99
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameAux0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Envases"
      TabPicture(1)   =   "frmVtasFactSocios.frx":0AB5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameAux1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame FrameAux1 
         BorderStyle     =   0  'None
         Height          =   3660
         Left            =   -74880
         TabIndex        =   77
         Top             =   330
         Width           =   11625
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   4
            Left            =   3105
            MaxLength       =   3
            TabIndex        =   89
            Tag             =   "Almacen|N|N|||facturassocio_envase|codalmac|000||"
            Text            =   "Alm"
            Top             =   2250
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   16
            Left            =   1575
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   94
            Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
            Top             =   3240
            Width           =   8430
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   10
            Left            =   9630
            MaxLength       =   2
            TabIndex        =   87
            Tag             =   "CodIva|N|N|||facturassocio_envase|codigiva|00||"
            Text            =   "Codiva"
            Top             =   2250
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   9
            Left            =   8820
            MaxLength       =   12
            TabIndex        =   95
            Tag             =   "Importe|N|N|||facturassocio_envase|importel|##,###,##0.00||"
            Text            =   "Importe"
            Top             =   2250
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   8
            Left            =   8010
            MaxLength       =   5
            TabIndex        =   93
            Tag             =   "Dto.Linea|N|N|||facturassocio_envase|dtolinea|#0.00||"
            Text            =   "Dto.Linea"
            Top             =   2250
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   7
            Left            =   7155
            MaxLength       =   12
            TabIndex        =   92
            Tag             =   "Precio|N|N|||facturassocio_envase|precioar|###,##0.0000||"
            Text            =   "Precio"
            Top             =   2250
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   3
            Left            =   2430
            MaxLength       =   12
            TabIndex        =   86
            Tag             =   "Num.Linea|N|N|||facturassocio_variedad|numlinea|000|S|"
            Text            =   "Linea"
            Top             =   2250
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   2
            Left            =   1710
            MaxLength       =   12
            TabIndex        =   85
            Tag             =   "Fec.Factu|F|N|||facturassocio_variedad|fecfactu||S|"
            Text            =   "FecFact"
            Top             =   2250
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   6
            Left            =   6255
            MaxLength       =   12
            TabIndex        =   91
            Tag             =   "Cantidad|N|N|||facturassocio_envase|cantidad|###,##0.00||"
            Text            =   "cantidad"
            Top             =   2250
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtAux 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   5
            Left            =   3645
            MaxLength       =   16
            TabIndex        =   90
            Tag             =   "Artículo|T|N|||facturassocio_envase|codartic||N|"
            Text            =   "Articulo"
            Top             =   2250
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
            Left            =   990
            MaxLength       =   12
            TabIndex        =   81
            Tag             =   "Num.Factura|N|N|||facturassocio_variedad|numfactu|0000000|S|"
            Text            =   "NumFact"
            Top             =   2250
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   0
            Left            =   270
            MaxLength       =   12
            TabIndex        =   80
            Tag             =   "Tipo Movim.|T|N|||facturassocio_variedad|codtipom||S|"
            Text            =   "TipoMov"
            Top             =   2250
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   0
            Left            =   4905
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   79
            Text            =   "Nombre articulo"
            Top             =   2250
            Width           =   1200
         End
         Begin VB.CommandButton btnBuscar 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   300
            Index           =   0
            Left            =   4680
            MaskColor       =   &H00000000&
            TabIndex        =   78
            ToolTipText     =   "Buscar Envase"
            Top             =   2250
            Visible         =   0   'False
            Width           =   195
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   1
            Left            =   90
            TabIndex        =   82
            Top             =   135
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
            Bindings        =   "frmVtasFactSocios.frx":0AD1
            Height          =   2550
            Left            =   90
            TabIndex        =   83
            Top             =   600
            Width           =   11340
            _ExtentX        =   20003
            _ExtentY        =   4498
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
            Index           =   1
            Left            =   1305
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
         Begin VB.Label Label1 
            Caption         =   "Ampliación Línea"
            Height          =   255
            Index           =   35
            Left            =   225
            TabIndex        =   88
            Top             =   3285
            Width           =   1335
         End
      End
      Begin VB.Frame FrameAux0 
         BorderStyle     =   0  'None
         Height          =   3525
         Left            =   90
         TabIndex        =   61
         Top             =   420
         Width           =   13200
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   15
            Left            =   5670
            MaxLength       =   11
            TabIndex        =   70
            Tag             =   "Unidades|N|S|||facturassocio_variedad|unidades|#,##0||"
            Text            =   "Unidades"
            Top             =   1800
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   14
            Left            =   7380
            MaxLength       =   11
            TabIndex        =   109
            Tag             =   "Iva|N|N|||facturassocio_variedad|codigiva|00|N|"
            Text            =   "Iva"
            Top             =   2250
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   13
            Left            =   6660
            MaxLength       =   11
            TabIndex        =   108
            Tag             =   "Dto2|N|S|||facturassocio_variedad|dtocom2|##0.00|N|"
            Text            =   "Dto2"
            Top             =   2250
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   12
            Left            =   5940
            MaxLength       =   11
            TabIndex        =   107
            Tag             =   "Dto1|N|S|||facturassocio_variedad|dtocom1|##0.00|N|"
            Text            =   "Dto1"
            Top             =   2250
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Frame Frame3 
            Caption         =   "Albarán"
            Height          =   3030
            Left            =   9585
            TabIndex        =   96
            Top             =   405
            Width           =   3570
            Begin VB.TextBox Text3 
               BackColor       =   &H80000018&
               Enabled         =   0   'False
               Height          =   285
               Index           =   4
               Left            =   180
               MaxLength       =   30
               TabIndex        =   106
               Text            =   "Text1 7"
               Top             =   2115
               Width           =   3150
            End
            Begin VB.TextBox Text3 
               BackColor       =   &H80000018&
               Enabled         =   0   'False
               Height          =   285
               Index           =   3
               Left            =   180
               MaxLength       =   15
               TabIndex        =   104
               Text            =   "Text1 7"
               Top             =   1575
               Width           =   3150
            End
            Begin VB.TextBox Text3 
               BackColor       =   &H80000018&
               Enabled         =   0   'False
               Height          =   285
               Index           =   2
               Left            =   180
               MaxLength       =   30
               TabIndex        =   102
               Text            =   "Text1 7"
               Top             =   1035
               Width           =   3150
            End
            Begin VB.TextBox Text3 
               BackColor       =   &H80000018&
               Enabled         =   0   'False
               Height          =   285
               Index           =   1
               Left            =   1800
               MaxLength       =   12
               TabIndex        =   100
               Text            =   "Text1 7"
               Top             =   495
               Width           =   1530
            End
            Begin VB.TextBox Text3 
               BackColor       =   &H80000018&
               Enabled         =   0   'False
               Height          =   285
               Index           =   0
               Left            =   180
               MaxLength       =   10
               TabIndex        =   99
               Text            =   "Text1 7"
               Top             =   495
               Width           =   1305
            End
            Begin VB.Label Label6 
               Caption         =   "Confección"
               Height          =   195
               Left            =   180
               TabIndex        =   105
               Top             =   1890
               Width           =   915
            End
            Begin VB.Label Label5 
               Caption         =   "Variedad"
               Height          =   195
               Left            =   180
               TabIndex        =   103
               Top             =   1350
               Width           =   915
            End
            Begin VB.Label Label4 
               Caption         =   "Destino"
               Height          =   195
               Left            =   180
               TabIndex        =   101
               Top             =   810
               Width           =   915
            End
            Begin VB.Label Label3 
               Caption         =   "Mat.Remolque"
               Height          =   195
               Left            =   1800
               TabIndex        =   98
               Top             =   270
               Width           =   1095
            End
            Begin VB.Label Label2 
               Caption         =   "Fecha Alb."
               Height          =   195
               Left            =   180
               TabIndex        =   97
               Top             =   270
               Width           =   915
            End
         End
         Begin VB.CommandButton btnBuscar 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   300
            Index           =   1
            Left            =   4005
            MaskColor       =   &H00000000&
            TabIndex        =   84
            ToolTipText     =   "Buscar Línea Albarán"
            Top             =   1800
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   5
            Left            =   3420
            MaxLength       =   3
            TabIndex        =   67
            Tag             =   "Lin.Albaran|N|N|||facturassocio_variedad|numlinealbar|000||"
            Text            =   "Lin.Alb"
            Top             =   1800
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   11
            Left            =   8415
            MaxLength       =   11
            TabIndex        =   74
            Tag             =   "Imp.Neto|N|N|||facturassocio_variedad|impornet|##,###,##0.00|N|"
            Text            =   "Imp.Neto"
            Top             =   1800
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   10
            Left            =   7785
            MaxLength       =   11
            TabIndex        =   73
            Tag             =   "Imp.Bruto|N|N|||facturassocio_variedad|imporbru|##,###,##0.00|N|"
            Text            =   "Imp.Brut"
            Top             =   1800
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   6
            Left            =   4185
            MaxLength       =   6
            TabIndex        =   68
            Tag             =   "Cant.Real|N|N|||facturassocio_variedad|cantreal|###,##0||"
            Text            =   "Cant.R"
            Top             =   1800
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   9
            Left            =   7200
            MaxLength       =   11
            TabIndex        =   72
            Tag             =   "Precio Neto|N|N|||facturassocio_variedad|precinet|###,##0.0000||"
            Text            =   "Prec.Neto"
            Top             =   1800
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   8
            Left            =   6435
            MaxLength       =   11
            TabIndex        =   71
            Tag             =   "Precio Bruto|N|N|||facturassocio_variedad|precibru|###,##0.0000||"
            Text            =   "Prec.Bruto"
            Top             =   1800
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   7
            Left            =   4950
            MaxLength       =   6
            TabIndex        =   69
            Tag             =   "Cant.Fact|N|N|||facturassocio_variedad|cantfact|###,##0||"
            Text            =   "Cant.Fact"
            Top             =   1800
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   4
            Left            =   2745
            MaxLength       =   7
            TabIndex        =   66
            Tag             =   "Num.Albaran|N|N|||facturassocio_variedad|numalbar|000000||"
            Text            =   "Albaran"
            Top             =   1800
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   3
            Left            =   2205
            MaxLength       =   30
            TabIndex        =   65
            Tag             =   "Num.Linea|N|N|||facturassocio_variedad|numlinea|000|S|"
            Text            =   "Linea"
            Top             =   1800
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   0
            Left            =   225
            MaxLength       =   7
            TabIndex        =   64
            Tag             =   "Tipo Movim.|T|N|||facturassocio_variedad|codtipom||S|"
            Text            =   "Tipomov"
            Top             =   1800
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   1
            Left            =   855
            MaxLength       =   15
            TabIndex        =   63
            Tag             =   "Num.Factura|N|N|||facturassocio_variedad|numfactu|0000000|S|"
            Text            =   "Factura"
            Top             =   1800
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   2
            Left            =   1440
            MaxLength       =   30
            TabIndex        =   62
            Tag             =   "Fec.Factu|F|N|||facturassocio_variedad|fecfactu||S|"
            Text            =   "Fec.Factu"
            Top             =   1800
            Visible         =   0   'False
            Width           =   765
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   0
            Left            =   135
            TabIndex        =   75
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
         Begin MSDataGridLib.DataGrid DataGrid2 
            Bindings        =   "frmVtasFactSocios.frx":0AE6
            Height          =   2940
            Left            =   135
            TabIndex        =   76
            Top             =   495
            Width           =   9360
            _ExtentX        =   16510
            _ExtentY        =   5186
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
            Index           =   0
            Left            =   1395
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
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   90
      TabIndex        =   31
      Top             =   8100
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
         TabIndex        =   32
         Top             =   180
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   11925
      TabIndex        =   29
      Top             =   8100
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   10710
      TabIndex        =   28
      Top             =   8100
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
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
            Object.ToolTipText     =   "Añadir"
            Object.Tag             =   "2"
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
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8400
         TabIndex        =   34
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   11925
      TabIndex        =   30
      Top             =   8100
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   270
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
      Left            =   270
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
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   11
      Left            =   6705
      MaxLength       =   15
      TabIndex        =   118
      Text            =   "Text1 7"
      Top             =   1710
      Width           =   1485
   End
   Begin VB.Label Label1 
      Caption         =   "Imp.Descuento 2"
      Height          =   255
      Index           =   10
      Left            =   6705
      TabIndex        =   119
      Top             =   1440
      Width           =   1215
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
         Enabled         =   0   'False
         HelpContextID   =   2
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
Attribute VB_Name = "frmVtasFactSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public Facturas As String  ' venimos de albaranes para ver las facturassocio donde aparece el albaran

'Si se llama del mantenimiento de clientes desde la solapa de documentos
Public hcoCodMovim As String 'cod. movim
Public hcoCodTipoM As String 'Codigo detalle de Movimiento(ALC)
Public hcoFechaMov As String 'fecha del movimiento

'========== VBLES PRIVADAS ====================
Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmArt As frmManArtic 'Form Mto de Articulos
Attribute frmArt.VB_VarHelpID = -1
Private WithEvents frmAlb As frmAlbVtasSocio 'Form Mto de Albaranes
Attribute frmAlb.VB_VarHelpID = -1

Private WithEvents frmSoc As frmManSocios 'Form Mto de socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmFPag As frmManFpago 'Form Mto de Formas de Pago
Attribute frmFPag.VB_VarHelpID = -1
Private WithEvents frmAlm As frmManAlmProp 'Form Mto de almacenes propios
Attribute frmAlm.VB_VarHelpID = -1

Private WithEvents frmMens As frmMensajes ' devolvemos las facturassocio a cuenta que vamos a descontar
Attribute frmMens.VB_VarHelpID = -1




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
'Para saber en MandaBusquedaPrevia si busca en la tabla scapla o en la tabla sdirec
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


Dim FacturasADescontar As String

Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos
Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal
Dim indice As Byte

Dim TipoFactura As Byte
Private BuscaChekc As String

Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 0 'Articulos
            Set frmArt = New frmManArtic
            frmArt.DatosADevolverBusqueda = "0|1|"
            frmArt.CodigoActual = txtAux(4).Text
            frmArt.Show vbModal
            Set frmArt = Nothing
            PonerFoco txtAux(4)
        Case 1 'Albaranes
            Set frmAlb = New frmAlbVtasSocio
            frmAlb.DatosADevolverBusqueda = "0|1|4|"
            frmAlb.CodigoActual = Text1(3).Text
            frmAlb.Show vbModal
            Set frmAlb = Nothing
            PonerFoco txtAux3(4)
    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1

End Sub

Private Sub Check1_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "check1(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "check1(" & Index & ")|"
    End If
End Sub

Private Sub chkAbonos_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Check1_GotFocus(Index As Integer)
    PonerFocoChk Me.Check1(Index)
End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
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
                    If ModificarLinea Then PosicionarData
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
            PonerFoco Text1(3)
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(3)
            
        Case 5 'Lineas Detalle
            TerminaBloquear
            Select Case SSTab1.Tab
                Case 0
                    If ModificaLineas = 1 Then 'INSERTAR
                        ModificaLineas = 0
                        DataGrid2.AllowAddNew = False
                        If Not Adoaux(0).Recordset.EOF Then Adoaux(0).Recordset.MoveFirst
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
                        If Not Adoaux(1).Recordset.EOF Then Adoaux(1).Recordset.MoveFirst
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
                    
             End Select
            
'            PonerBotonCabecera True
    
            
            
            
'            Me.DataGrid1.Enabled = True
    End Select
End Sub
Private Sub BotonAnyadir()

    LimpiarCampos 'Huida els TextBox
    
'    TipoFactura = 1
    
    PonerModo 3
    
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la capçalera *******
'    Text1(0).Text = SugerirCodigoSiguienteStr("palets", "numpalet")
'    FormateaCampo Text1(0)
'    cmbAux(0).ListIndex = -1
    
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
        
    LimpiarDataGrids
    
    PonerFoco Text1(3) '*** 1r camp visible que siga PK ***
    
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
        
'        'poner los txtaux para buscar por lineas de albaran
'        anc = DataGrid2.Top
'        If DataGrid2.Row < 0 Then
'            anc = anc + 440
'        Else
'            anc = anc + DataGrid2.RowTop(DataGrid2.Row) + 20
'        End If
'        LLamaLineas Modo, anc, "DataGrid2"
        
        
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(3)
        Text1(3).BackColor = vbYellow
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
        EsCabecera = True
        MandaBusquedaPrevia "facturassocio.codtipom <> 'EAC'"
    Else
        LimpiarCampos
        LimpiarDataGrids
        CadenaConsulta = "Select facturassocio.* "
        CadenaConsulta = CadenaConsulta & "from " & NombreTabla & " where codtipom <> 'EAC' "
'        CadenaConsulta = CadenaConsulta & " WHERE scafac.codtipom='" & CodTipoMov & "'"
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

'    'solo se puede modificar la factura si no esta contabilizada
'    If Check1(0).Value = 1 Or Check1(1).Value = 1 Or Check1(2).Value = 1 Then
'        MsgBox "Esta factura no podemos modificarla", vbExclamation
'        TerminaBloquear
'        Exit Sub
'    End If
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    PonerFoco Text1(1) '*** 1r camp visible que siga PK ***
        
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
    
'--monica
'    If Data2.Recordset.EOF Then
'        TerminaBloquear
'        Exit Sub
'    End If
    
    
    NumTabMto = Index
    If Index = 2 Then NumTabMto = 3
    PonerModo 5, Index
 

    
    Select Case NumTabMto
        Case 0 ' variedades
            vWhere = Replace(ObtenerWhereCP(False), "facturassocio", "facturassocio_variedad")
            vWhere = vWhere & " and numlinea=" & Adoaux(0).Recordset!numlinea
            If Not BloqueaRegistro("facturassocio_variedad", vWhere) Then
                TerminaBloquear
                Exit Sub
            End If
            If DataGrid2.Bookmark < DataGrid2.FirstRow Or DataGrid2.Bookmark > (DataGrid2.FirstRow + DataGrid2.VisibleRows - 1) Then
                J = DataGrid2.Bookmark - DataGrid2.FirstRow
                DataGrid2.Scroll 0, J
                DataGrid2.Refresh
            End If
            
        '    anc = ObtenerAlto(Me.DataGrid1)
            anc = DataGrid2.Top
            If DataGrid2.Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGrid2.RowTop(DataGrid2.Row) + 10
            End If
        
            For J = 0 To 7
                txtAux3(J).Text = DataGrid2.Columns(J).Text
            Next J
            txtAux3(15).Text = DataGrid2.Columns(8).Text
            txtAux3(8).Text = DataGrid2.Columns(9).Text
            txtAux3(9).Text = DataGrid2.Columns(10).Text
            
            txtAux3(10).Text = DataGrid2.Columns(11).Text
            txtAux3(11).Text = DataGrid2.Columns(12).Text
            
            txtAux3(12).Text = DataGrid2.Columns(18).Text '17
            txtAux3(13).Text = DataGrid2.Columns(19).Text '18
            txtAux3(14).Text = DataGrid2.Columns(20).Text '19
            
            For J = 4 To 11
                txtAux3(J).Enabled = True
            Next J
            txtAux3(15).Enabled = True
            
            BloquearTxt txtAux3(8), False
            BloquearTxt txtAux3(10), False
            
            ModificaLineas = 2 'Modificar
            LLamaLineas ModificaLineas, anc, "DataGrid2"
            
            'Añadiremos el boton de aceptar y demas objetos para insertar
            Me.lblIndicador.Caption = "MODIFICAR"
            PonerModoOpcionesMenu (Modo)
            PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
            DataGrid2.Enabled = True
            
'            PonerBotonCabecera False
            Me.DataGrid3.Enabled = False
            PonerFoco txtAux3(4)
        
            BloquearBtn Me.btnBuscar(1), True
        Case 1 ' envases
            vWhere = Replace(ObtenerWhereCP(False), "facturassocio", "facturassocio_envases")
            vWhere = vWhere & " and numlinea=" & Adoaux(1).Recordset!numlinea
            If Not BloqueaRegistro("facturassocio_envases", vWhere) Then
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
        
            For J = 0 To 5
                txtAux(J).Text = DataGrid3.Columns(J).Text
            Next J
            For J = 7 To 10
                txtAux(J - 1).Text = DataGrid3.Columns(J).Text
            Next J
            
            Text2(0).Text = DataGrid3.Columns(6).Text  ' nombre de articulo
            Text2(16).Text = DataGrid3.Columns(11).Text ' ampliacion
            
'            cmbAux(0).Text = DataGrid3.Columns(7).Text
            txtAux(10).Text = DataGrid3.Columns(12).Text

            BloquearTxt txtAux(4), True
            BloquearTxt txtAux(5), True
            BloquearTxt txtAux(7), True
            BloquearTxt txtAux(9), True
            txtAux(4).Enabled = False
            txtAux(5).Enabled = False
            txtAux(7).Enabled = False
            txtAux(9).Enabled = False
            
            BloquearTxt txtAux(6), False
            BloquearTxt txtAux(8), False
            
            BloquearBtn Me.btnBuscar(0), True
            
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


       
    End Select
    
eModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub LLamaLineas(xModo As Byte, Optional alto As Single, Optional grid As String)
Dim jj As Integer
Dim b As Boolean
    
    Select Case grid
            
        Case "DataGrid2"
            DeseleccionaGrid Me.DataGrid2
            b = (xModo = 1 Or xModo = 2)
            For jj = 4 To 11
                txtAux3(jj).Height = DataGrid2.RowHeight
                txtAux3(jj).Top = alto + 5
                txtAux3(jj).visible = b
            Next jj
            txtAux3(15).Height = DataGrid2.RowHeight
            txtAux3(15).Top = alto + 5
            txtAux3(15).visible = b
            
            btnBuscar(1).Height = DataGrid2.RowHeight - 10
            btnBuscar(1).Top = alto + 5
            btnBuscar(1).visible = b
        
        Case "DataGrid3"
            DeseleccionaGrid Me.DataGrid3
            b = (xModo = 1 Or xModo = 2)
            For jj = 4 To 9
                txtAux(jj).Height = DataGrid3.RowHeight - 10
                txtAux(jj).Top = alto + 5
                txtAux(jj).visible = b
            Next jj
            Text2(0).Height = DataGrid3.RowHeight - 10
            Text2(0).Top = alto + 5
            Text2(0).visible = b
           
            btnBuscar(0).Height = DataGrid3.RowHeight - 10
            btnBuscar(0).Top = alto + 5
            btnBuscar(0).visible = b
            
            
    End Select
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Facturas (scafac)
' y los registros correspondientes de las tablas cab. albaranes (scafac1)
' y las lineas de la factura (slifac)
Dim cad As String
'Dim vTipoMov As CTiposMov

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
'    'solo se puede modificar la factura si no esta contabilizada
'    If FactContabilizada Then Exit Sub
    
    cad = "Cabecera de Facturas." & vbCrLf
    cad = cad & "-------------------------------------      " & vbCrLf & vbCrLf
    cad = cad & "Va a eliminar la Factura:            "
    cad = cad & vbCrLf & "Tipo:  " & Text1(6).Text
    cad = cad & vbCrLf & "Nº Factura:  " & Format(Text1(0).Text, "0000000")
    cad = cad & vbCrLf & "Fecha:  " & Format(Text1(1).Text, "dd/mm/yyyy")

    cad = cad & vbCrLf & vbCrLf & " ¿Desea Eliminarla? "

    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
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

Private Sub Combo1_GotFocus(Index As Integer)
    If Modo = 1 Then Combo1(Index).BackColor = vbYellow
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    If Combo1(Index).BackColor = vbYellow Then Combo1(Index).BackColor = vbWhite
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not Me.Adoaux(0).Recordset.EOF And ModificaLineas <> 1 Then
        CargarDatosAlbaran Me.Adoaux(0).Recordset!NumAlbar, Adoaux(0).Recordset!numlinealbar
    End If
End Sub

Private Sub DataGrid3_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not Adoaux(1).Recordset.EOF Then
        Text2(16).Text = DBLet(Adoaux(1).Recordset!ampliaci, "T")
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
    'Viene de DblClick en documentos en la ficha de cliente
    If hcoCodMovim <> "" And Not Data1.Recordset.EOF Then PonerCadenaBusqueda
    
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
    btnPrimero = 13
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(4).Image = 3   'Insertar
        .Buttons(5).Image = 4   'Modificar
        .Buttons(6).Image = 5   'Borrar
        .Buttons(8).Image = 10  'Impresión de albaran
        .Buttons(10).Image = 11  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
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
    
    LimpiarCampos   'Limpia los campos TextBox
    CargaCombo

'--monica
'    CodTipoMov = "ALV" 'hcoCodTipoM
    VieneDeBuscar = False
    
        
    '## A mano
    NombreTabla = "facturassocio"
    NomTablaLineas = "facturassocio_variedad" 'Tabla lineas de variedades

    Ordenacion = " ORDER BY facturassocio.codtipom, facturassocio.numfactu, facturassocio.fecfactu"
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    CadenaConsulta = "select * from facturassocio "
    If hcoCodMovim <> "" Then
        CadenaConsulta = CadenaConsulta & " where codtipom = '" & hcoCodTipoM & "' and numfactu = " & DBSet(hcoCodMovim, "N") & " and fecfactu = " & DBSet(hcoFechaMov, "F") & " and codtipom <> 'EAC'"
    Else
        CadenaConsulta = CadenaConsulta & " where numfactu = -1 and codtipom <> 'EAC'"
    End If
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
'        CargaGrid DataGrid1, Data2, False
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    SSTab1.Tab = 0
    
'    If DatosADevolverBusqueda <> "" Then
'        Text1(0).Text = DatosADevolverBusqueda
'        HacerBusqueda
'        SSTab1.Tab = 1
'    Else
'        PonerModo 0
'    End If
    
    If DatosADevolverBusqueda = "" Then
        If Facturas = "" Then
            PonerModo 0
        Else
            HacerBusqueda
            SSTab1.Tab = 0
        End If
    Else
        BotonBuscar
    End If
    
    PrimeraVez = False
End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.Combo1(0).ListIndex = -1
    Me.Check1(0).Value = 0
    Me.Check1(1).Value = 0
    Me.Check1(2).Value = 0
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    If Modo = 4 Then TerminaBloquear
End Sub

Private Sub frmAlb_DatoSeleccionado(CadenaSeleccion As String)
    txtAux3(4).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000") 'Numero albaran
    txtAux3(5).Text = Format(RecuperaValor(CadenaSeleccion, 2), "00") 'Numero linea
    CargarDatosAlbaran txtAux3(4).Text, txtAux3(5).Text
End Sub

Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod almacen
    Text2(indice + 2).Text = RecuperaValor(CadenaSeleccion, 2) 'Nombre del almacen
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Articulos
    txtAux(5).Text = RecuperaValor(CadenaSeleccion, 1) 'codartic
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
    If txtAux(5) <> "" Then
        txtAux(7) = DevuelveDesdeBDNew(cAgro, "sartic", "preciove", "codartic", txtAux(5), "T")
        If Combo1(0).ListIndex = 1 Then
            txtAux(10).Text = vParamAplic.CodIvaExento
        Else
            txtAux(10) = DevuelveDesdeBDNew(cAgro, "sartic", "codigiva", "codartic", txtAux(5), "T")
        End If
    End If
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim cadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        cadB = ""
        Aux = ValorDevueltoFormGrid(Text1(6), CadenaDevuelta, 1)
        cadB = Aux
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 2)
        cadB = cadB & " and  " & Aux
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 3)
        cadB = cadB & " and " & Aux
        
        
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
    Screen.MousePointer = vbDefault
End Sub



Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    If imgFec(0).Tag < 2 Then
        Text1(CByte(imgFec(0).Tag) + 1).Text = Format(vFecha, "dd/mm/yyyy") '<===
    Else
        Text1(CByte(imgFec(0).Tag) + 8).Text = Format(vFecha, "dd/mm/yyyy") '<===
    End If
    ' ********************************************
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000") 'Cod Cliente
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 3) 'Nombre del cliente
    PonerFoco Text1(indice)
End Sub

Private Sub frmFPag_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00") 'Codigo
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nombre
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
Dim sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        FacturasADescontar = " (codtipom, numfactu, fecfactu) in (" & CadenaSeleccion & ")"
    Else
        FacturasADescontar = ""
    End If
  
    If FacturasADescontar <> "" Then
        sql = "insert into facturassocio_acuenta (codtipom,numfactu,fecfactu,codtipomcta,numfactucta,fecfactucta,totalfaccta)"
        sql = sql & " select " & DBSet(Text1(6).Text, "T") & "," & DBSet(Text1(0).Text, "N") & "," & DBSet(Text1(1).Text, "F") & ", codtipom, numfactu, fecfactu, totalfac from facturassocio "
        sql = sql & " where " & FacturasADescontar
        
        conn.Execute sql
        
    End If
  
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(indice).Text = vCampo
End Sub

Private Sub imgBuscar_Click(Index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. de Socio
            indice = 3
            PonerFoco Text1(indice)
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            PonerFoco Text1(indice)
        
        Case 1 'Forma de Pago
            indice = 4
            PonerFoco Text1(indice)
            Set frmFPag = New frmManFpago
            frmFPag.DatosADevolverBusqueda = "0|1|"
            frmFPag.Show vbModal
            Set frmFPag = Nothing
            PonerFoco Text1(indice)
            
'        Case 4 ' Almacen
'            Indice = 16
'            PonerFoco Text1(Indice)
'            Set frmAlm = New frmManAlmProp
'            frmAlm.DatosADevolverBusqueda = "0|1|"
'            frmAlm.Show vbModal
'            Set frmAlm = Nothing
'            PonerFoco Text1(Indice)
    End Select
    
    Screen.MousePointer = vbDefault
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
        If Text1(Index + 1).Text <> "" Then frmC.NovaData = Text1(Index + 1).Text
    Else
        imgFec(0).Tag = Index '<===
        ' *** repasar si el camp es txtAux o Text1 ***
        If Text1(Index + 8).Text <> "" Then frmC.NovaData = Text1(Index + 8).Text
    End If
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    If Index < 2 Then
        PonerFoco Text1(CByte(imgFec(0).Tag) + 1) '<===
    Else
        PonerFoco Text1(CByte(imgFec(0).Tag) + 8) '<===
    End If
    ' ********************************************
End Sub


Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        indice = 2
        frmZ.pTitulo = "Observaciones de la Factura"
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
    If Modo = 5 Then 'Eliminar lineas de Pedido
'         BotonEliminarLinea
    Else   'Eliminar Pedido
         BotonEliminar
    End If
End Sub

Private Sub mnImprimir_Click()
'Imprimir Factura
    
    If Data1.Recordset.EOF Then Exit Sub
    
    BotonImprimir
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub


Private Sub mnModificar_Click()
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
Dim sql As String

    On Error GoTo EBloqueaAlb
    
    BloqueaAlbxFac = False
    'bloquear cabecera albaranes x factura
    sql = "select * FROM scafac1 "
    sql = sql & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute sql, , adCmdText
    BloqueaAlbxFac = True

EBloqueaAlb:
    If Err.Number <> 0 Then BloqueaAlbxFac = False
End Function


Private Function BloqueaLineasFac() As Boolean
'bloquea todas las lineas de la factura
Dim sql As String

    On Error GoTo EBloqueaLin

    BloqueaLineasFac = False
    'bloquear cabecera albaranes x factura
    sql = "select * FROM slifac "
    sql = sql & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute sql, , adCmdText
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
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 2 Or (Index = 2 And Text1(2).Text = "") Then KEYpress KeyAscii
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
Dim sql As String
Dim Nregs As Long
Dim vTipoMov As CTiposMov
        
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1 'Fecha factura
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
            
        Case 3 'Socios
            If PonerFormatoEntero(Text1(Index)) Then
                If Modo = 1 Then
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rsocios", "nomsocio")
                End If
                PonerDatosSocio (Text1(Index).Text)
            End If
            
        Case 4 ' Forma de Pago
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "forpago", "nomforpa")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Forma de Pago: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmFPag = New frmManFpago
                        frmFPag.DatosADevolverBusqueda = "0|1|"
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmFPag.Show vbModal
                        Set frmFPag = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            End If
            
         Case 6 ' tipo de movimiento
            If Text1(Index).Text <> "" Then
                Text1(Index).Text = UCase(Text1(Index).Text)
'--monica:10/02/2009 stipom
'                SQL = ""
'                SQL = DevuelveDesdeBDNew(cAgro, "stipom", "contador", "codtipom", Text1(Index).Text, "T")
'                If SQL = "" Then
'++monica:10/02/2009 stipom
                Nregs = TotalRegistros("select count(*) from usuarios.stipom stipom where codtipom = " & DBSet(Text1(Index).Text, "T"))
                If Nregs = 0 Then
'++
                    MsgBox "No existe el tipo de movimiento. Reintroduzca.", vbExclamation
                    PonerFoco Text1(Index)
                Else
                    If TipoFactura = 1 Then
                        Set vTipoMov = New CTiposMov
                        If vTipoMov.leer(Text1(6).Text) Then
                            Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
                            PonerFoco Text1(0)
                        End If
                        Set vTipoMov = Nothing
                        
                    End If
                    
                End If
            End If
            
         Case 7, 8 'descuentos
            If Modo = 1 Then Exit Sub
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 4
            
         Case 5 ' importe de descuento
            If Modo = 1 Then Exit Sub
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 3
    

    End Select
End Sub


Private Sub HacerBusqueda()
Dim cadB As String
Dim cadAux As String
    
'    '--- Laura 12/01/2007
'    cadAux = Text1(5).Text
'    If Text1(4).Text <> "" Then Text1(5).Text = ""
'    '---
    
'    '--- Laura 12/01/2007
'    Text1(5).Text = cadAux
'    '---
    
'--monica
'    CadB = ObtenerBusqueda(Me)
'++monica
    If Facturas = "" Then
        cadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    Else
        cadB = Facturas
    End If
    
    '[Monica]22/06/2010 seleccionamos unicamente las facturassocio que no sean de tipo a cuenta
    If cadB <> "" Then
        cadB = cadB & " and facturassocio.codtipom <> 'EAC' "
    Else
        cadB = "facturassocio.codtipom <> 'EAC'"
    End If

    If chkVistaPrevia = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select facturassocio.* from " & NombreTabla
        CadenaConsulta = CadenaConsulta & " WHERE " & cadB & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim tabla As String
Dim Titulo As String
Dim Desc As String, devuelve As String
    'Llamamos a al form
    '##A mano
    cad = ""
    cad = cad & "Tipo|facturas.codtipom|N||10·"
    cad = cad & "Nº.Factura|facturas.numfactu|N||15·"
    cad = cad & "Cliente|facturas.codclien|N||10·" 'ParaGrid(Text1(3), 10, "Cliente")
    cad = cad & "Nombre Cliente|clientes.nomclien|N||45·"
    cad = cad & ParaGrid(Text1(1), 15, "F.Factura")
    tabla = NombreTabla & " INNER JOIN clientes ON facturassocio.codclien=clientes.codclien "
    
    Titulo = "Facturas"
    devuelve = "0|1|4|"
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|1|4|"
        frmB.vDevuelve = devuelve
        frmB.vTitulo = Titulo
        frmB.vSelElem = 0
'        frmB.vConexionGrid = cAgro  'Conexión a BD: Ariagro
        If Not EsCabecera Then frmB.Label1.FontSize = 11
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
'        If EsCabecera Then
'            PonerCadenaBusqueda
'            Text1(0).Text = Format(Text1(0).Text, "0000000")
'        End If
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If HaDevueltoDatos Then
            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                cmdRegresar_Click
        Else   'de ha devuelto datos, es decir NO ha devuelto datos
            PonerFoco Text1(kCampo)
        End If
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
        If Modo = 1 Then
            PonerFoco Text1(kCampo)
'            Text1(0).BackColor = vbYellow
        End If
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
        '--monica
        'LLamaLineas Modo, 0, "DataGrid2"
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
    
    For i = 0 To 2
        Select Case i
            Case 0 'variedades
                If Data1.Recordset.RecordCount > 0 Then
                    CargaGrid DataGrid2, Adoaux(0), True
                Else
                    CargaGrid DataGrid2, Adoaux(0), False
                End If
                If Not Adoaux(0).Recordset.EOF Then CargarDatosAlbaran Adoaux(0).Recordset!NumAlbar, Adoaux(0).Recordset!numlinealbar
            Case 1  ' envases
                If Data1.Recordset.RecordCount > 0 Then
                    CargaGrid DataGrid3, Adoaux(1), True
                Else
                    CargaGrid DataGrid3, Adoaux(1), False
                End If
                If Not Adoaux(1).Recordset.EOF Then
                    Text2(16).Text = DBLet(Adoaux(1).Recordset!ampliaci, "T")
                Else
                    Text2(16).Text = ""
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
    b = PonerCamposForma2(Me, Data1, 2, "FrameFactura")

    Text1(12).Text = Format(Round2(DBLet(Data1.Recordset!BrutoFac, "N") - DBLet(Data1.Recordset!impordto, "N"), 2), FormatoImporte)

    CodTipoMov = Text1(6).Text
    
'    FormatoDatosTotales
    
    'poner descripcion campos
    Modo = 4
    
    TipoFactura = 0 'DevuelveDesdeBDNew(cAgro, "clientes", "tipofact", "codclien", Text1(3).Text, "N")
    
    Text2(3).Text = PonerNombreDeCod(Text1(3), "rsocios", "nomsocio", "codsocio", "N") 'codsocio
    Text2(4).Text = DevuelveDesdeBDNew(cAgro, "forpago", "nomforpa", "codforpa", Text1(4), "N") 'forma de pago
'    Text2(18).Text = PonerNombreDeCod(Text1(16), "salmpr", "nomalmac", "codalmac", "N") 'almacen
    
    Modo = 2
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Ofertas
    
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
Dim i As Byte, Numreg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo

    'Actualiza Iconos Insertar,Modificar,Eliminar
    '## No tiene el boton modificar y no utiliza la funcion general
'    ActualizarToolbar Modo, Kmodo
    BuscaChekc = ""
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Or Facturas <> "" Or hcoCodMovim <> "" Then
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
    For i = 9 To 31
        BloquearTxt Text1(i), Not (Modo = 1)
        Text1(i).Enabled = (Modo = 1)
    Next i
'    Me.Check1.Enabled = (Modo = 1)
    
    b = (Modo <> 1) And (Modo <> 3)
    
    'Campos Nº Factura bloqueado y en azul
    BloquearTxt Text1(0), b, True 'And (TipoFactura = 0)  'numero factura
    BloquearTxt Text1(6), b 'And (TipoFactura = 0) 'codtipom
    BloquearTxt Text1(1), b 'fechafactura
    BloquearTxt Text1(3), b 'cliente
    BloquearCmb Me.Combo1(0), (Modo <> 1)
    BloquearChk Me.Check1(0), (Modo <> 1)
    BloquearChk Me.Check1(1), (Modo <> 1)
    BloquearChk Me.Check1(2), (Modo <> 1)
    
    Me.imgZoom(0).Enabled = Not (Modo = 0)
    
    If Modo = 3 Then
        Text1(0).Enabled = (TipoFactura = 1)
        Text1(6).Enabled = (TipoFactura = 1)
    End If
    
    'Si no es modo lineas Boquear los TxtAux
    For i = 0 To txtAux.Count - 1
        txtAux(i).visible = False
        BloquearTxt txtAux(i), True
    Next i
    
    For i = 0 To 0
        Text2(i).visible = ((Modo = 5) And (indFrame = 1))
        Text2(i).Enabled = False
    Next i
    
    BloquearTxt Text2(16), (Modo <> 5)
    
    BloquearBtn Me.btnBuscar(0), True
    BloquearBtn Me.btnBuscar(1), True
    
    For i = 0 To Text3.Count - 1
        Text3(i).Enabled = False
    Next i
    
'    For I = 4 To 11
    For i = 0 To txtAux3.Count - 1
        txtAux3(i).visible = False
        BloquearTxt txtAux3(i), True
'        txtAux3(I).visible = ((Modo = 5) And (indFrame = 0))
'        If I <> 6 And I <> 7 And I <> 9 And I <> 11 Then
'            BloquearTxt txtAux3(I), Not ((Modo = 5) And (indFrame = 0))
'        Else
'            BloquearTxt txtAux3(I), True
'        End If
'
    Next i
    
    
    
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    BloquearImgBuscar Me, Modo, ModificaLineas
    BloquearImgFec Me, 0, Modo
                    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
    ' *** si n'hi han llínies i algún tab que no te datagrid ***
    Select Case NumTabMto
        Case 0
            BloquearFrameAux Me, "FrameAux0", Modo, NumTabMto
        Case 1
            BloquearFrameAux Me, "FrameAux1", Modo, NumTabMto
        Case 2
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


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim b As Boolean
Dim Serie As String
Dim sql As String

    On Error GoTo EDatosOK

    DatosOk = False
    
'    ComprobarDatosTotales

    'comprobamos datos OK de la tabla scafac
    b = CompForm2(Me, 2, "Frame2") ' , 1) 'Comprobar formato datos ok de la cabecera: opcion=1
    If Not b Then Exit Function
    
    If Modo = 3 Then
        'si el tipo de factura es manual y no hemos introducido valor en numero de factura
        If TipoFactura = 1 And Text1(0).Text = "" Then
            MsgBox "El Número de Factura no puede estar vacio. Reintroduzca.", vbExclamation
            PonerFoco Text1(0)
            b = False
        End If

        'comprobamos que no exista ya la factura en la tabla facturassocio de ariagro
        sql = ""
        sql = DevuelveDesdeBDNew(cAgro, "facturassocio", "numfactu", "codtipom", Text1(6).Text, "T", , "numfactu", Text1(0).Text, "N") ', "fecfactu", Text1(1).Text, "F")
        If sql <> "" Then
            MsgBox "Factura ya existente. Reintroduzca.", vbExclamation
            PonerFoco Text1(0)
            b = False
        End If
        If Not b Then Exit Function
        'comprobamos que no exista ya en la tabla facturassocio de contabilidad
        Serie = ""
'--monica:10/02/2009 stipom
'        Serie = DevuelveDesdeBDNew(cAgro, "stipom", "letraser", "codtipom", Text1(6).Text, "T")
'++monica
        Serie = ObtenerLetraSerie(Text1(6).Text)
'++
        If Serie <> "" Then
            sql = ""
            If vParamAplic.ContabilidadNueva Then
                sql = DevuelveDesdeBDNew(cConta, "factcli", "numfactu", "numserie", Serie, "T", , "numfactu", Text1(0).Text, "N", "fecfactu", Text1(1).Text, "F")
            Else
                sql = DevuelveDesdeBDNew(cConta, "cabfact", "codfaccl", "numserie", Serie, "T", , "codfaccl", Text1(0).Text, "N", "fecfaccl", Text1(1).Text, "F")
            End If
            If sql <> "" Then
                MsgBox "Factura existente en contabilidad. Reintroduzca.", vbExclamation
                PonerFoco Text1(0)
                b = False
            End If
        Else
            MsgBox "El tipo de factura no tiene serie asociada. Revise.", vbExclamation
            b = False
        End If
        If Not b Then Exit Function
    End If
    
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

    If BloqueaRegistro(NombreTabla, "numfactu = " & Data1.Recordset!NumFactu) Then
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
    End If

End Sub


Private Sub BotonEliminarLinea(Index As Integer)
Dim cad As String
Dim sql As String
Dim Mens As String
Dim b As Boolean

    On Error GoTo EEliminarLinea

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
    ' ***************************************************************************

    b = True

    Select Case Index
        Case 0 'variedades
            ' *************** canviar la pregunta ****************
            cad = "¿Seguro que desea eliminar la Variedad?"
            cad = cad & vbCrLf & "Factura: " & Adoaux(0).Recordset.Fields(0)
            cad = cad & vbCrLf & "Albarán: " & Adoaux(0).Recordset.Fields(4) & "-" & Adoaux(0).Recordset.Fields(5)
            
            If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
                On Error GoTo EEliminarLinea
                Screen.MousePointer = vbHourglass
                NumRegElim = Adoaux(0).Recordset.AbsolutePosition
                TerminaBloquear
                sql = "delete from facturassocio_variedad where codtipom = " & DBSet(Adoaux(0).Recordset.Fields(0), "T")
                sql = sql & " and numfactu = " & Adoaux(0).Recordset.Fields(1)
                sql = sql & " and fecfactu = " & DBSet(Adoaux(0).Recordset.Fields(2), "F")
                sql = sql & " and numlinea = " & DBSet(Adoaux(0).Recordset.Fields(3), "N")
                conn.Execute sql
                               
                Mens = "Recalcular Dtos lineas"
                b = RecalcularDtos(Adoaux(0).Recordset.Fields(0), Adoaux(0).Recordset.Fields(1), Adoaux(0).Recordset.Fields(2), Mens)
                
                CalcularDatosFactura
                
                If b Then
                    SituarDataTrasEliminar Adoaux(0), NumRegElim
                    
                    If Me.Adoaux(0).Recordset.EOF Then
                        CargarDatosAlbaran "", ""
                    End If
                    
                    CargaGrid DataGrid2, Adoaux(0), True
                    SSTab1.Tab = 0
                End If
            End If
            Screen.MousePointer = vbDefault
       Case 1 'envases
            ' *************** canviar la pregunta ****************
            cad = "¿Seguro que desea eliminar el Envase?"
            cad = cad & vbCrLf & "Factura: " & Adoaux(1).Recordset.Fields(1)
            cad = cad & vbCrLf & "Artículo: " & Adoaux(1).Recordset.Fields(5) & " - " & Adoaux(1).Recordset.Fields(6)
            
            If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
                On Error GoTo EEliminarLinea
                Screen.MousePointer = vbHourglass
                NumRegElim = Adoaux(1).Recordset.AbsolutePosition
                
                If Not EliminarLinea Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                Else
                    CalcularDatosFactura
                    If SituarDataTrasEliminar(Adoaux(1), NumRegElim) Then
                        PonerCampos
                    Else
                        PonerCampos
'                        LimpiarCampos
'                        PonerModo 0
                    End If
                End If
            End If
            Screen.MousePointer = vbDefault
       
       Case 2 'facturassocio a cuenta
            ' *************** canviar la pregunta ****************
            cad = "¿Seguro que desea eliminar la factura a cuenta?"
            cad = cad & vbCrLf & "Factura: " & Adoaux(2).Recordset.Fields(4)
            
            If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
                On Error GoTo EEliminarLinea
                Screen.MousePointer = vbHourglass
                NumRegElim = Adoaux(2).Recordset.AbsolutePosition
                
                If Not EliminarLineaFacCta Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                Else
                
                    CalcularDatosFactura
                    
                    If SituarDataTrasEliminar(Adoaux(2), NumRegElim) Then
                        PonerCampos
                    Else
                        PonerCampos
                    End If
                End If
            End If
            Screen.MousePointer = vbDefault
       
    End Select
       
EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Or Not b Then MuestraError Err.Number, "Eliminar Linea de Factura", Err.Description
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Buscar
            mnBuscar_Click
        Case 2  'Todos
            BotonVerTodos
        
        Case 4  'Añadir
            mnNuevo_Click

        Case 5  'Modificar
            mnModificar_Click
        Case 6  'Borrar
            mnEliminar_Click
        Case 8  ' Impresion de albaran
            mnImprimir_Click
        Case 10    'Salir
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
Dim sql As String

    On Error GoTo ECargaGRid

    b = DataGrid3.Enabled
    Select Case vDataGrid.Name
        Case "DataGrid2" 'variedades
            Opcion = 0
        Case "DataGrid3" 'envases
            Opcion = 1
        Case "DataGrid1"  'facturassocio a cuenta
            Opcion = 2
    End Select
    
    sql = MontaSQLCarga(enlaza, Opcion)
    CargaGridGnral vDataGrid, vData, sql, PrimeraVez
    
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
                       
         Case "DataGrid2" 'facturassocio_variedad
'select codtipom,numfactu,fecfactu,numlinea,numalbar,numlinealbar,cantreal,cantfact,precibru,precinet,
'dtocom1,dtocom2,imporbru,impornet,codigiva
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux3(4)|T|Albarán|800|;S|txtAux3(5)|T|Linea|600|;"
            tots = tots & "S|btnBuscar(1)|B|||;S|txtAux3(6)|T|Cant.Real|1000|;S|txtAux3(7)|T|Cant.Fact.|1000|;S|txtAux3(15)|T|Uds|800|;S|txtAux3(8)|T|Prec.Bruto|1000|;"
            tots = tots & "S|txtAux3(9)|T|Prec.Neto|1000|;S|txtAux3(10)|T|Imp.Bruto|1200|;"
            tots = tots & "S|txtAux3(11)|T|Imp.Neto|1200|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            arregla tots, DataGrid2, Me
            
            
                     
         Case "DataGrid3" 'facturassocio_envases
'select codtipom,numfactu,fecfactu,numlinea,codartic,sartic.nomartic,cantidad,precioar,dtolinea,importel,ampliaci,codigiva
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;S|txtAux(4)|T|Alm|900|;"
            tots = tots & "S|txtAux(5)|T|Articulo|1500|;S|btnBuscar(0)|B|||;"
            tots = tots & "S|Text2(0)|T|Nombre|3500|;S|txtAux(6)|T|Cantidad|1200|;"
            tots = tots & "S|txtAux(7)|T|Precio|1200|;S|txtAux(8)|T|Dto|800|;S|txtAux(9)|T|Importe|1200|;N||||0|;N||||0|;"
            arregla tots, DataGrid3, Me
            
            
            If Adoaux(0).Recordset.EOF Then CargarDatosAlbaran "", ""
            
    End Select
    
    vDataGrid.HoldFields
    Exit Sub
    
ECargaGRid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub

'
'Private Sub TxtAux_Change(Index As Integer)
'    If Index = 6 And ModificaLineas = 2 Then 'Precio y Modo Borrar Lineas
'        txtAux(5).Text = "M"
'    End If
'End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    
    ConseguirFocoLin txtAux(Index)
End Sub


Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim sql As String
Dim devuelve As String
Dim b As Boolean
Dim TipoDto As Byte
Dim vCStock As CStock


    'Quitar espacios en blanco
    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    Select Case Index
        Case 4 'almacen
            If txtAux(Index).Text <> "" Then PonerFormatoEntero txtAux(Index)
        
'        Case 7 'Precio
'             'Tipo 2: Decimal(10,4)
'             If txtAux(Index).Text <> "" Then PonerFormatoDecimal txtAux(Index), 2
'
            
        Case 5 'articulo
            If txtAux(Index).Text = "" Then
                Exit Sub
            End If
        
            If txtAux(4).Text = "" Then
                MsgBox "Debe seleccionar un almacen.", vbInformation
                PonerFoco txtAux(4)
                Exit Sub
            End If
        
            devuelve = ""
            If ModificaLineas = 2 Then
                If Not Adoaux(1).Recordset.EOF Then devuelve = Adoaux(1).Recordset!codArtic
            End If
        
            If Not PonerArticulo(txtAux(5), Text2(0), txtAux(4).Text, CodTipoMov, ModificaLineas, devuelve) Then
                PonerFoco txtAux(Index)
            Else
                txtAux(7) = DevuelveDesdeBDNew(cAgro, "sartic", "preciove", "codartic", txtAux(5), "T")
                If Combo1(0).ListIndex = 1 Then
                    txtAux(10).Text = vParamAplic.CodIvaExento
                Else
                    txtAux(10).Text = DevuelveDesdeBDNew(cAgro, "sartic", "codigiva", "codartic", txtAux(5), "T")
                End If
'--monica
'                b = (Me.ActiveControl.Name = "txtAux")
'                If b Then b = (Me.ActiveControl.Index = 4)
'                If Not b Then
''                    If txtAux(2).Locked Then PonerFoco txtAux(3)
'                Else
'                    PonerFoco txtAux(4)
'                End If
            End If
        
        Case 6 ' Cantidad
            If PonerFormatoDecimal(txtAux(Index), 1) Then  'Tipo 1: Decimal(12,2)
            
                'Comprobar si hay suficiente stock
                Set vCStock = New CStock
                If Not InicializarCStock(vCStock, "S") Then Exit Sub
                If vCStock.MueveStock Then 'Comprobar si el articulo mueve stock: tiene control de stock y no es instalacion
                  If Not vCStock.MoverStock Then
                    PonerFoco txtAux(Index)
                    Set vCStock = Nothing
                    Exit Sub
                  End If
                End If
                
                Set vCStock = Nothing
            End If
            
        Case 7 ' Precio
            PonerFormatoDecimal txtAux(Index), 7
            
        Case 8  'Descuentos
            PonerFormatoDecimal txtAux(Index), 4 'Tipo 4: Decimal(4,2)
            
        Case 9 'Importe Linea
            PonerFormatoDecimal txtAux(Index), 3 'Tipo 3: Decimal(10,2)
            
    End Select
     If (Index = 6 Or Index = 7 Or Index = 8 Or Index = 9) Then 'Cant., Precio, Dto1, Dto2
'        If Trim(TxtAux(3).Text) = "" Or Trim(TxtAux(4).Text) = "" Then Exit Sub
'        If Trim(TxtAux(6).Text) = "" Or Trim(TxtAux(7).Text) = "" Then Exit Sub
        If txtAux(8).Text = "" Then txtAux(8).Text = 0
        TipoDto = 0 'DevuelveDesdeBDNew(cAgro, "clientes", "tipodtos", "codclien", Text1(3).Text, "N")
        
        txtAux(9).Text = CalcularImporteFClien(txtAux(6).Text, txtAux(7).Text, txtAux(8).Text, 0, TipoDto, 0)
        PonerFormatoDecimal txtAux(9), 3
    End If
    
End Sub




Private Function Eliminar() As Boolean
Dim sql As String, LEtra As String, Sql2 As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim Mens As String

    On Error GoTo FinEliminar

    If Data1.Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    b = EliminarStock

    If b Then
        'Eliminar en tablas de cabecera de albaran
        '------------------------------------------
        sql = " " & ObtenerWhereCP(True)
        
        'Lineas de envases (facturassocio_variedad)
        conn.Execute "Delete from facturassocio_variedad " & Replace(sql, "facturassocio", "facturassocio_variedad")
        
        'Lineas de coste (facturassocio_envases)
        conn.Execute "Delete from facturassocio_envases " & Replace(sql, "facturassocio", "facturassocio_envases")
        
        'Cabecera de factura
        conn.Execute "Delete from " & NombreTabla & sql
        
        'Decrementar contador si borramos el ult. palet
        Set vTipoMov = New CTiposMov
        vTipoMov.DevolverContador Text1(6).Text, Val(Text1(0).Text)
        Set vTipoMov = Nothing
        
        b = True
    End If
FinEliminar:
    If Err.Number <> 0 Or Not b Then
        MuestraError Err.Number, "Eliminar Factura", Err.Description & " " & Mens
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
Dim sql As String, LEtra As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim Mens As String
Dim vCStock As CStock

    On Error GoTo FinEliminar

    b = False
    If Adoaux(1).Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    Mens = ""
    
    'Eliminar en tablas de facturassocio_envases
    '------------------------------------------
    sql = " where codtipom = " & DBSet(Adoaux(1).Recordset.Fields(0), "T")
    sql = sql & " and numfactu = " & Adoaux(1).Recordset.Fields(1)
    sql = sql & " and fecfactu = " & DBSet(Adoaux(1).Recordset.Fields(2), "F")
    sql = sql & " and numlinea = " & DBSet(Adoaux(1).Recordset.Fields(3), "N")


     ' borramos el movimiento y aumentamos el stock
    Set vCStock = New CStock
    If Not InicializarCStock(vCStock, "E") Then Exit Function

     'en actualizar stock comprobamos si el articulo tiene control de stock
     b = vCStock.DevolverStock
     Set vCStock = Nothing

    'Lineas de variedades
    conn.Execute "Delete from facturassocio_envases " & sql
    
FinEliminar:
    If Err.Number <> 0 Or Not b Then
        MuestraError Err.Number, "Eliminar Envases de la Factura ", Err.Description & " " & Mens
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

Private Function EliminarLineaFacCta() As Boolean
Dim sql As String, LEtra As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim Mens As String
Dim vCStock As CStock

    On Error GoTo FinEliminar

    b = False
    If Adoaux(2).Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    Mens = ""
    
    b = True
    
    'Eliminar en tablas de facturassocio_acuenta
    '------------------------------------------
    sql = " where codtipom = " & DBSet(Adoaux(2).Recordset.Fields(0), "T")
    sql = sql & " and numfactu = " & Adoaux(2).Recordset.Fields(1)
    sql = sql & " and fecfactu = " & DBSet(Adoaux(2).Recordset.Fields(2), "F")
    sql = sql & " and codtipomcta = " & DBSet(Adoaux(2).Recordset.Fields(3), "T")
    sql = sql & " and numfactucta = " & Adoaux(2).Recordset.Fields(4)
    sql = sql & " and fecfactucta = " & DBSet(Adoaux(2).Recordset.Fields(5), "F")


    'Lineas de variedades
    conn.Execute "Delete from facturassocio_acuenta " & sql
    
FinEliminar:
    If Err.Number <> 0 Or Not b Then
        MuestraError Err.Number, "Eliminar Facturas a Cuenta ", Err.Description & " " & Mens
        b = False
    End If
    If Not b Then
        conn.RollbackTrans
        EliminarLineaFacCta = False
    Else
        conn.CommitTrans
        EliminarLineaFacCta = True
    End If
End Function




Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next

    CargaGrid DataGrid2, Me.Adoaux(1), False 'envases
    CargaGrid DataGrid3, Me.Adoaux(0), False 'variedades
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & Replace(ObtenerWhereCP(False), "facturassocio.", "") & ")"
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
Dim sql As String

    On Error Resume Next
    
    sql = "facturassocio.codtipom = " & DBSet(Text1(6).Text, "T") & " and facturassocio.numfactu= " & DBSet(Text1(0).Text, "N") & " and facturassocio.fecfactu= " & DBSet(Text1(1).Text, "F")
    If conWhere Then sql = " WHERE " & sql
    ObtenerWhereCP = sql
    
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
Dim sql As String
    
    Select Case Opcion
        Case 0  'variedades
'select codtipom,numfactu,fecfactu,numlinea,numalbar,numlinealbar,cantreal,cantfact,precibru,precinet,
'dtocom1,dtocom2,imporbru,impornet,codigiva
            sql = "SELECT facturassocio_variedad.codtipom,numfactu,fecfactu,facturassocio_variedad.numlinea,"
            sql = sql & " facturassocio_variedad.numalbar,numlinealbar,cantreal,"
            sql = sql & " cantfact,facturassocio_variedad.unidades, precibru, precinet, imporbru,impornet, fechaalb, matrirem, "
            sql = sql & " destinos.nomdesti,variedades.nomvarie, forfaits.nomconfe, dtocom1, dtocom2, facturassocio_variedad.codigiva  "
            sql = sql & " FROM facturassocio_variedad, albaran, albaran_variedad, variedades, forfaits, destinos " 'lineas de variedades de la factura
            sql = sql & " WHERE facturassocio_variedad.numalbar = albaran.numalbar "
            sql = sql & " and albaran.numalbar = albaran_variedad.numalbar "
            sql = sql & " and facturassocio_variedad.numlinealbar = albaran_variedad.numlinea "
            sql = sql & " and albaran_variedad.codvarie = variedades.codvarie "
            sql = sql & " and albaran_variedad.codforfait = forfaits.codforfait "
            sql = sql & " and albaran.codclien = destinos.codclien "
            sql = sql & " and albaran.coddesti = destinos.coddesti "
            
            If enlaza Then
                sql = sql & " and " & Replace(ObtenerWhereCP(False), "facturassocio", "facturassocio_variedad")
            Else
                sql = sql & " and numfactu = -1"
            End If
            sql = sql & " ORDER BY codtipom,numfactu,fecfactu,numlinea"
                    
        Case 1  'envases
'select codtipom,numfactu,fecfactu,numlinea,codalmac,codartic,sartic.nomartic,cantidad,precioar,dtolinea,importel,ampliaci,codigiva
            sql = "SELECT codtipom,numfactu,fecfactu,numlinea,codalmac,facturassocio_envases.codartic,sartic.nomartic,cantidad,"
            sql = sql & "precioar,dtolinea,importel,ampliaci,facturassocio_envases.codigiva"
            sql = sql & " FROM facturassocio_envases, sartic "
            sql = sql & " WHERE facturassocio_envases.codartic = sartic.codartic "
    
            If enlaza Then
                sql = sql & " and " & Replace(ObtenerWhereCP(False), "facturassocio", "facturassocio_envases")
            Else
                sql = sql & " and numfactu = -1"
            End If
            sql = sql & " ORDER BY codtipom,numfactu,fecfactu,numlinea"
    
        Case 2 ' facturassocio a cuenta
            sql = "SELECT facturassocio_acuenta.codtipom, facturassocio_acuenta.numfactu, facturassocio_acuenta.fecfactu, facturassocio_acuenta.codtipomcta, "
            sql = sql & " facturassocio_acuenta.numfactucta, facturassocio_acuenta.fecfactucta, facturassocio_acuenta.totalfaccta, facturassocio.baseimp1, facturassocio.porciva1, facturassocio.impoiva1,  "
            sql = sql & " facturassocio.porcrec1, facturassocio.imporec1 "
            sql = sql & " FROM facturassocio_acuenta, facturassocio "
            sql = sql & " WHERE facturassocio_acuenta.codtipomcta = facturassocio.codtipom  "
            sql = sql & " and facturassocio_acuenta.numfactucta = facturassocio.numfactu "
            sql = sql & " and facturassocio_acuenta.fecfactucta = facturassocio.fecfactu "
    
            If enlaza Then
                sql = sql & " and " & Replace(ObtenerWhereCP(False), "facturassocio", "facturassocio_acuenta")
            Else
                sql = sql & " and facturassocio_acuenta.numfactu = -1"
            End If
            sql = sql & " ORDER BY facturassocio_acuenta.codtipom,facturassocio_acuenta.numfactu,facturassocio_acuenta.fecfactu,facturassocio_acuenta.codtipomcta,facturassocio_acuenta.numfactucta,facturassocio_acuenta.fecfactucta"
    
    End Select
    
    MontaSQLCarga = sql
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean, bAux As Boolean
Dim i As Integer

        b = ((Modo = 2) Or (Modo = 0)) And (Facturas = "") And (hcoCodMovim = "") 'Or (Modo = 5 And ModificaLineas = 0)
        'Buscar
        Toolbar1.Buttons(1).Enabled = b
        Me.mnBuscar.Enabled = b
        'Vore Tots
        Toolbar1.Buttons(2).Enabled = b
        Me.mnVerTodos.Enabled = b
        'Añadir
        Toolbar1.Buttons(4).Enabled = b
        Me.mnModificar.Enabled = b
        
        
        b = (Modo = 2 And Data1.Recordset.RecordCount > 0) And (Facturas = "") And (hcoCodMovim = "")
        'Modificar
        Toolbar1.Buttons(5).Enabled = b And Not (Check1(0).Value = 1 Or Check1(1).Value = 1 Or Check1(2).Value = 1)
        Me.mnModificar.Enabled = b And Not (Check1(0).Value = 1 Or Check1(1).Value = 1 Or Check1(2).Value = 1)
        'eliminar
        Toolbar1.Buttons(6).Enabled = b And Not (Check1(0).Value = 1 Or Check1(1).Value = 1 Or Check1(2).Value = 1)
        Me.mnEliminar.Enabled = b And Not (Check1(0).Value = 1 Or Check1(1).Value = 1 Or Check1(2).Value = 1)
        'Impresión de factura
        Toolbar1.Buttons(8).Enabled = ((Modo = 2) And (Facturas = "")) Or (hcoCodMovim <> "")
        Me.mnImprimir.Enabled = ((Modo = 2) And (Facturas = "")) Or (hcoCodMovim <> "")
'        'Orden de Carga
'        Toolbar1.Buttons(9).Enabled = b
'        Me.mnOrdenCarga.Enabled = b
'        'Generar CMR
'        Toolbar1.Buttons(10).Enabled = b
'        Me.mnCMR.Enabled = b
        

    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
'++monica: si insertamos lo he quitado
'    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
    b = (Modo = 2) And (Facturas = "") And (hcoCodMovim = "") And Not (Check1(0).Value = 1 Or Check1(1).Value = 1 Or Check1(2).Value = 1)
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = b
        
        If b Then
            Select Case i
              Case 0
                bAux = (b And Me.Adoaux(0).Recordset.RecordCount > 0)
              Case 1
                bAux = (b And Me.Adoaux(1).Recordset.RecordCount > 0)
              Case 2
                bAux = (b And Me.Adoaux(2).Recordset.RecordCount > 0)
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

    If Text1(0).Text = "" Then
        MsgBox "Debe seleccionar una Factura para Imprimir.", vbInformation
        Exit Sub
    End If
    
    cadFormula = ""
    cadParam = ""
    cadselect = ""
    numParam = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    indRPT = 68 'Impresion de Factura
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
      
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de factura
    '---------------------------------------------------
    If Text1(0).Text <> "" Then
        'Tipo de factura
        devuelve = "{" & NombreTabla & ".codtipom}='" & Text1(6).Text & "'"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        devuelve = "codtipom = '" & Text1(6).Text & "'"
        If Not AnyadirAFormula(cadselect, devuelve) Then Exit Sub
        'Nº Factura
        devuelve = "{" & NombreTabla & ".numfactu}=" & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        devuelve = "numfactu = " & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadselect, devuelve) Then Exit Sub
        'Fecha Factura
        devuelve = "{" & NombreTabla & ".fecfactu}=Date(" & Year(Text1(1).Text) & "," & Month(Text1(1).Text) & "," & Day(Text1(1).Text) & ")"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        devuelve = "fecfactu = " & DBSet(Text1(1).Text, "F")
        If Not AnyadirAFormula(cadselect, devuelve) Then Exit Sub
    End If
    
    If Not HayRegParaInforme(NombreTabla, cadselect) Then Exit Sub
     
     With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 0
            .Titulo = "Impresión de Factura Venta Socio"
            .ConSubInforme = True
            .Show vbModal
    End With
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
Dim TipoDto As Byte
Dim ImpDto As String
Dim Unidades As String
Dim Cantidad As String
Dim cad As String

    'Quitar espacios en blanco
    If Not PerderFocoGnralLineas(txtAux3(Index), ModificaLineas) Then Exit Sub
    
    Select Case Index
        Case 4 'Albaran
            If txtAux3(Index) <> "" Then PonerFormatoEntero txtAux3(Index)
            
            CargarDatosAlbaran txtAux3(4).Text, txtAux3(5).Text
                
        Case 5 'Linea de albaran
            If txtAux3(Index) <> "" Then PonerFormatoEntero txtAux3(Index)
            
            If txtAux3(4).Text <> "" And txtAux3(5).Text <> "" Then
                If AlbaranFacturado(txtAux3(4).Text, txtAux3(5).Text) Then
                    cad = "Esta línea de Albarán está facturada. " & vbCrLf & vbCrLf & "    ¿ Desea continuar ? "
                    If MsgBox(cad, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                        CargarDatosAlbaran txtAux3(4).Text, txtAux3(5).Text
                    Else
                        txtAux3(4).Text = ""
                        txtAux3(5).Text = ""
                    End If
                Else
                    CargarDatosAlbaran txtAux3(4).Text, txtAux3(5).Text
                End If
            End If
            
            If txtAux3(4).Text = "" Or txtAux3(5).Text = "" Then
                PonerFoco txtAux3(4)
            Else
                PonerFoco txtAux3(8)
            End If
        
        Case 8 'precio bruto
            If txtAux3(Index).Text <> "" Then
                If PonerFormatoDecimal(txtAux3(Index), 7) Then
                    
                    Select Case TipoFacturarForfaits(txtAux3(4).Text, txtAux3(5).Text)
                        Case 0  'por unidades
                            txtAux3(10).Text = Round2(CCur(ImporteSinFormato(txtAux3(Index).Text)) * CCur(ImporteSinFormato(txtAux3(15).Text)), 2)
                            PonerFormatoDecimal txtAux3(10), 3
                        Case 1  'por kilos
                            txtAux3(10).Text = Round2(CCur(ImporteSinFormato(txtAux3(Index).Text)) * CCur(ImporteSinFormato(txtAux3(6).Text)), 2)
                            PonerFormatoDecimal txtAux3(10), 3
                        Case Else
                            
                    End Select
                    
                    cmdAceptar.SetFocus
                Else
                    Exit Sub
                End If
            End If
        
        Case 10 'importe bruto
            If txtAux3(Index).Text <> "" Then
                If PonerFormatoDecimal(txtAux3(Index), 3) Then
                
                    Select Case TipoFacturarForfaits(txtAux3(4).Text, txtAux3(5).Text)
                        Case 0
                            Unidades = ComprobarCero(txtAux3(15).Text)
                            If CCur(Unidades) <> 0 Then
                                txtAux3(8).Text = Round2(CCur(ImporteSinFormato(txtAux3(Index).Text)) / CCur(Unidades), 4)
                            Else
                                txtAux3(8).Text = 0
                            End If
                            PonerFormatoDecimal txtAux3(8), 7
                        Case 1
                            Cantidad = ComprobarCero(txtAux3(6).Text)
                            If CCur(Cantidad) <> 0 Then
                                txtAux3(8).Text = Round2(CCur(ImporteSinFormato(txtAux3(Index).Text)) / CCur(Cantidad), 4)
                            Else
                                txtAux3(8).Text = 0
                            End If
                            PonerFormatoDecimal txtAux3(8), 7
                        Case Else
                        
                    End Select
                        
                    cmdAceptar.SetFocus
               Else
                    Exit Sub
               End If
            End If
    End Select

If ((Index = 8 And txtAux3(Index).Text <> "") Or (Index = 10 And txtAux3(Index).Text <> "")) Then
        Dim campo2 As String
        campo2 = "nrodecprec"
        TipoDto = 0 'DevuelveDesdeBDNew(cAgro, "clientes", "tipodtos", "codclien", Text1(3).Text, "N", campo2)
        campo2 = 2
        Select Case TipoFacturarForfaits(txtAux3(4).Text, txtAux3(5).Text)
            Case 0 ' unidades
'                ImpDto = CalcularImporteDto(txtAux3(15).Text, txtAux3(8).Text, txtAux3(0).Text, txtAux3(1).Text, txtAux3(2).Text, CStr(DBLet(Data1.Recordset!impdtoc, "N")), False)
'                txtAux3(11).Text = CalcularImporte(txtAux3(15).Text, txtAux3(8).Text, txtAux3(12).Text, txtAux3(13).Text, TipoDto, ImpDto)
                Unidades = ComprobarCero(txtAux3(15).Text)
                ImpDto = CalcularImporteDto(txtAux3(15).Text, CStr(CCur(ImporteSinFormato(txtAux3(10).Text)) / CCur(Unidades)), txtAux3(0).Text, txtAux3(1).Text, txtAux3(2).Text, CStr(DBLet(Data1.Recordset!ImpDtoc, "N")), False)
                txtAux3(11).Text = CalcularImporteFClien(txtAux3(15).Text, CStr(CCur(ImporteSinFormato(txtAux3(10).Text)) / CCur(Unidades)), txtAux3(12).Text, txtAux3(13).Text, TipoDto, ImpDto, txtAux3(10).Text)
                PonerFormatoDecimal txtAux3(11), 1
                
                'precio neto
                If ComprobarCero(txtAux3(15).Text) <> "0" Then
                    txtAux3(9).Text = Round2(CCur(ImporteSinFormato(txtAux3(11).Text)) / CCur(ImporteSinFormato(txtAux3(15).Text)), CCur(campo2))
                End If
                PonerFormatoDecimal txtAux3(9), 7
            
            Case 1 ' kilos
'                ImpDto = CalcularImporteDto(txtAux3(6).Text, txtAux3(8).Text, txtAux3(0).Text, txtAux3(1).Text, txtAux3(2).Text, CStr(DBLet(Data1.Recordset!impdtoc, "N")), False)
'                txtAux3(11).Text = CalcularImporte(txtAux3(6).Text, txtAux3(8).Text, txtAux3(12).Text, txtAux3(13).Text, TipoDto, ImpDto)
                Cantidad = ComprobarCero(txtAux3(6).Text)
                ImpDto = CalcularImporteDto(txtAux3(6).Text, CStr(CCur(ImporteSinFormato(txtAux3(10).Text)) / CCur(Cantidad)), txtAux3(0).Text, txtAux3(1).Text, txtAux3(2).Text, CStr(DBLet(Data1.Recordset!ImpDtoc, "N")), False)
                txtAux3(11).Text = CalcularImporteFClien(txtAux3(6).Text, CStr(CCur(ImporteSinFormato(txtAux3(10).Text)) / CCur(Cantidad)), txtAux3(12).Text, txtAux3(13).Text, TipoDto, ImpDto, txtAux3(10).Text)
                PonerFormatoDecimal txtAux3(11), 1
                
                'precio neto
                If ComprobarCero(txtAux3(6).Text) <> "0" Then
                    txtAux3(9).Text = Round2(CCur(ImporteSinFormato(txtAux3(11).Text)) / CCur(ImporteSinFormato(txtAux3(6).Text)), CCur(campo2))
                End If
                PonerFormatoDecimal txtAux3(9), 7
            
            Case Else
            
        End Select
        
    End If
    
End Sub

Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
Dim sql As String
Dim i As Byte
    
    Combo1(0).Clear
    
    Combo1(0).AddItem "Normal"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    
    Combo1(0).AddItem "Exento"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    
    Combo1(0).AddItem "Recargo Equiv."
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    
End Sub

Private Function ModificaCabecera() As Boolean
Dim b As Boolean
Dim MenError As String

    On Error GoTo EModificarCab

    conn.BeginTrans
    
'    MenError = "Eliminando Costes"
'    b = EliminarCostes(Data1.Recordset.Fields(0))
    
    b = ModificaDesdeFormulario2(Me, 2, "Frame2")
    
    
    MenError = "Recalculando Importes Netos de lineas"
    If b Then b = RecalcularDtos(Text1(6).Text, Text1(0).Text, Text1(1).Text, MenError)
    CalcularDatosFactura
'    MenError = "Insertando Costes"
'    b = InsertarCostes(Data1.Recordset.Fields(0))

EModificarCab:
    If Err.Number <> 0 Then
        MenError = "Modificando Factura." & vbCrLf & "----------------------------" & vbCrLf & MenError
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
Dim sql As String

    On Error GoTo EInsertarCab
    
    CodTipoMov = Text1(6).Text
    
   If TipoFactura = 0 Then
        Set vTipoMov = New CTiposMov
        If vTipoMov.leer(CodTipoMov) Then
            Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
            sql = CadenaInsertarDesdeForm(Me)
            If sql <> "" Then
                If InsertarOferta(sql, vTipoMov) Then
                    CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                    PonerCadenaBusqueda
                    PonerModo 2
                    'Ponerse en Modo Insertar Lineas
    '                BotonMtoLineas 0, "Variedades"
    
'                    DescontarFacturasACuenta Text1(6).Text, Text1(0).Text, Text1(1).Text, Text1(3).Text
                    
                    BotonAnyadirLinea 0
                End If
            End If
            Text1(0).Text = Format(Text1(0).Text, "0000000")
        End If
        Set vTipoMov = Nothing
    Else
        sql = CadenaInsertarDesdeForm(Me)
        conn.Execute sql

        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
        PonerCadenaBusqueda
        PonerModo 2
        'Ponerse en Modo Insertar Lineas
'                BotonMtoLineas 0, "Variedades"

'        DescontarFacturasACuenta Text1(6), Text1(0).Text, Text1(1).Text, Text1(3).Text
        
        BotonAnyadirLinea 0
        Text1(0).Text = Format(Text1(0).Text, "0000000")

    End If
    
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
    'para ello vemos si existe una factura con ese contador y si existe la incrementamos
    Do
        devuelve = DevuelveDesdeBDNew(cAgro, NombreTabla, "numfactu", "codtipom", Text1(6).Text, "T", , "numfactu", Text1(0), "N", "fecfactu", Text1(1), "F")
        If devuelve <> "" Then
            'Ya existe el contador incrementarlo
            Existe = True
            vTipoMov.IncrementarContador (CodTipoMov)
            Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
            cambiaSQL = True
        Else
            Existe = False
        End If
    Loop Until Not Existe
    If cambiaSQL Then vSQL = CadenaInsertarDesdeForm(Me)
    
    
    'Aqui empieza transaccion
    conn.BeginTrans
    MenError = "Error al insertar en la tabla Cabecera de Facturas (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
    MenError = "Error al actualizar el contador de la Factura."
    vTipoMov.IncrementarContador (CodTipoMov)
    
EInsertarOferta:
        If Err.Number <> 0 Then
            MenError = "Insertando Factura." & vbCrLf & "----------------------------" & vbCrLf & MenError
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


Private Sub CargaForaGrid()
    If DataGrid2.Columns.Count <= 2 Then Exit Sub
    ' *** posar als camps de fora del grid el valor de la columna corresponent ***
    Text3(0) = DataGrid2.Columns(12).Text    'Fecha
    Text3(1) = DataGrid2.Columns(13).Text    'Matricula
    Text3(2) = DataGrid2.Columns(14).Text    'Destino
    Text3(3) = DataGrid2.Columns(15).Text   'Variedad
    Text3(4) = DataGrid2.Columns(16).Text   'Confeccion
    ' *** Si fora del grid n'hi han camps de descripció, posar-los valor ***
    ' **********************************************************************
End Sub

Private Sub InsertarLinea(Index As Integer)
'Inserta registre en les taules de Llínies
Dim nomframe As String
Dim b As Boolean
Dim Mens As String

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case Index
        Case 0: nomframe = "FrameAux0" 'variedades
        Case 1: nomframe = "FrameAux1" 'envases
        Case 2: nomframe = "FrameAux2" 'facturassocio_acuenta
    End Select
    ' ***************************************************************
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        
        Select Case Index
            Case 0
                If InsertarDesdeForm2(Me, 2, nomframe) Then
                    ' *** si n'hi ha que fer alguna cosa abas d'insertar
                    If NumTabMto = 0 Then
        'yo                'si ha seleccionat "cuenta principal", actualise totes les atres a "no"
        '                If (cmbAux(10).ItemData(cmbAux(10).ListIndex) = 1) Then
        '                    ActualisaCtaprpal (txtaux(2).Text)
        '                End If
                    End If
                    
                    Mens = "Recalcular Dtos lineas"
                    b = RecalcularDtos(txtAux3(0).Text, txtAux3(1).Text, txtAux3(2).Text, Mens)
                    
                    If b Then
                        CalcularDatosFactura
                        ' *************************************************
                        b = BloqueaRegistro("facturassocio", "codtipom = " & DBSet(Data1.Recordset!codTipoM, "T") & " and numfactu = " & DBSet(Data1.Recordset!NumFactu, "N") & " and fecfactu = " & DBSet(Data1.Recordset!FecFactu, "F"))
                        CargaGrid DataGrid2, Adoaux(0), True
                        If b Then BotonAnyadirLinea NumTabMto
                        SSTab1.Tab = NumTabMto
                    End If
                End If
            Case 1
                If InsertarLineaEnv(txtAux(3).Text) Then
                    CalcularDatosFactura
                    b = BloqueaRegistro("facturassocio", "numfactu = " & Data1.Recordset!NumFactu)
                    CargaGrid DataGrid3, Adoaux(1), True
                    If b Then BotonAnyadirLinea NumTabMto
                    SSTab1.Tab = NumTabMto
                End If
'            Case 2
'                If InsertarLineaFacCta(txtAux(3).Text) Then
'                    CalcularDatosFactura
'                    b = BloqueaRegistro("facturassocio", "numfactu = " & Data1.Recordset!NumFactu)
'                    CargaGrid DataGrid1, Adoaux(2), True
'                    If b Then BotonAnyadirLinea NumTabMto
'                    SSTab1.Tab = NumTabMto
'                End If
            
        End Select
    End If
End Sub

Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vTabla As String
Dim anc As Single
Dim i As Integer
    
    ModificaLineas = 1 'Posem Modo Afegir Llínia
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Capçalera
        cmdAceptar_Click
        If ModificaLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    
    PonerModo 5, Index
    
    ' *** bloquejar la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    BloquearTxt Text1(6), True
    BloquearTxt Text1(1), True
    
    
    ' **************************************************

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case NumTabMto
        Case 0: vTabla = "facturassocio_variedad"
        Case 1: vTabla = "facturassocio_envases"
        Case 2: vTabla = "facturassocio_acuenta"
    End Select
    ' ********************************************************
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case NumTabMto
        Case 0 ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", vWhere)
            ' ***************************************************************

            AnyadirLinea DataGrid2, Adoaux(0)
    
            anc = DataGrid2.Top
            If DataGrid2.Row < 0 Then
                anc = anc + 220 '210
            Else
                anc = anc + DataGrid2.RowTop(DataGrid2.Row) + 5
            End If
            
            LLamaLineas ModificaLineas, anc, "DataGrid2"
        
            LimpiarCamposLin "FrameAux0"
            LimpiarCamposLin "Frame3"
            
            txtAux3(0).Text = Text1(6).Text 'codtipom
            txtAux3(1).Text = Text1(0).Text 'numfactu
            txtAux3(2).Text = Text1(1).Text 'fecfactu
            txtAux3(3).Text = NumF 'numlinea
            
            
            BloquearTxt txtAux3(4), False
            BloquearTxt txtAux3(5), False
            BloquearTxt txtAux3(8), False
            BloquearTxt txtAux3(10), False
            txtAux3(12).Enabled = False
            txtAux3(12).visible = False
            txtAux3(13).Enabled = False
            txtAux3(13).visible = False
            txtAux3(14).Enabled = False
            txtAux3(14).visible = False
            txtAux3(12).Text = Text1(7).Text
            txtAux3(13).Text = Text1(8).Text
            
            Select Case Me.Combo1(0).ListIndex
                Case 0  'normal
                    '++monica:27/05/08 el iva se cogerá de la variedad cuando carguemos datos de albaran
                    txtAux3(14).Text = "" 'vParamAplic.CodIvaNormal
                Case 1  'exento
                    txtAux3(14).Text = vParamAplic.CodIvaExento
                Case 2  'recargo de equivalencia
                    txtAux3(14).Text = vParamAplic.CodIvaRecargo
            End Select
            
            BloquearBtn Me.btnBuscar(1), False
            
'                    BloquearTxt txtaux(12), False
            PonerFoco txtAux3(4)
                    
        ' *** si n'hi han llínies sense datagrid ***
        Case 1
            NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", vWhere)
            ' ***************************************************************

            AnyadirLinea DataGrid3, Adoaux(1)
    
            anc = DataGrid3.Top
            If DataGrid3.Row < 0 Then
                anc = anc + 215 '210
            Else
                anc = anc + DataGrid3.RowTop(DataGrid3.Row) + 5
            End If
          
            LLamaLineas ModificaLineas, anc, "DataGrid3"
        
            LimpiarCamposLin "FrameAux1"
            txtAux(0).Text = Text1(6).Text 'codtipom
            txtAux(1).Text = Text1(0).Text 'numfactu
            txtAux(2).Text = Text1(1).Text 'fecfactu
            txtAux(3).Text = NumF
            PonerFoco txtAux(4)
            For i = 0 To 0
                Text2(i).Text = ""
            Next i
            txtAux(10).Enabled = False
            txtAux(10).visible = False
            BloquearTxt txtAux(9), True
            BloquearTxt Text2(16), False
            BloquearBtn Me.btnBuscar(0), False
        ' ******************************************
        
        ' *** si n'hi han llínies sense datagrid ***
        
    
    End Select
End Sub

Private Sub LimpiarCamposLin(frameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, frameAux  'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Function ModificarLinea() As Boolean
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim V As Integer
Dim cad As String
Dim sql As String
Dim vCStock As CStock
Dim b As Boolean
Dim Mens As String
    
    On Error GoTo eModificarLinea

    ModificarLinea = False
    sql = ""

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'variedades
        Case 1: nomframe = "FrameAux1" 'envases
        Case 2: nomframe = "FrameAux2" 'facturassocio a cuenta
    End Select
    ' **************************************************************

    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        
        Select Case NumTabMto
        Case 0
            If ModificaDesdeFormulario2(Me, 2, nomframe) Then
                ' *** si cal que fer alguna cosa abas d'insertar ***
                If NumTabMto = 0 Then
    'yo                'si ha seleccionat "cuenta principal", actualise totes les atres a "no"
    '                If (cmbAux(10).ItemData(cmbAux(10).ListIndex) = 1) Then
    '                    ActualisaCtaprpal (txtaux(2).Text)
    '                End If
                End If
                ' ******************************************************
                Mens = "Recalcular Dtos lineas"
                b = RecalcularDtos(txtAux3(0).Text, txtAux3(1).Text, txtAux3(2).Text, Mens)
    
    '            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
                If b Then
                    CalcularDatosFactura
                    ModificaLineas = 0
        
                    V = Adoaux(0).Recordset.Fields(3) 'el 2 es el nº de llinia
                    CargaGrid DataGrid2, Adoaux(0), True
    
                    ' *** si n'hi han tabs ***
                    SSTab1.Tab = 0
    
                    DataGrid2.SetFocus
                    Adoaux(0).Recordset.Find (Adoaux(0).Recordset.Fields(3).Name & " =" & V)
    
                    LLamaLineas ModificaLineas, 0, "DataGrid2"
               End If
            End If
    
        Case 1
            Set vCStock = New CStock
            If Not InicializarCStock(vCStock, "S") Then Exit Function
            
            If DatosOkLineaEnv(vCStock) Then
                '#### LAURA 15/11/2006
                conn.BeginTrans
                
        '        Set vCStock = New CStock
                'iniciamos la clase con los valores anteriores para deshacer lo q insertamos antes
                b = InicializarCStock(vCStock, "E")
                If b Then
                    b = vCStock.DevolverStock 'eliminamos de smoval y devolvemos stock valores anteriores
                    'ahora leemos los valores nuevos
                    If b Then b = InicializarCStock(vCStock, "S")
                    'insertamos en smoval y actualizamos stock a los valores nuevos
                    vCStock.Cantidad = CSng(ComprobarCero(txtAux(6).Text))
                    If b Then b = vCStock.ActualizarStock
            
                    'actualizar la linea de Albaran
                    If b Then
                        sql = "UPDATE facturassocio_envases Set codalmac = " & txtAux(4).Text & ", codartic=" & DBSet(txtAux(5).Text, "T") & ", "
                        sql = sql & "ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
                        sql = sql & "cantidad= " & DBSet(txtAux(6).Text, "N") & ", "
                        sql = sql & "precioar= " & DBSet(txtAux(7).Text, "N") & ", " 'precio
                        sql = sql & "dtolinea= " & DBSet(txtAux(8).Text, "N") & ", "
                        sql = sql & "importel= " & DBSet(txtAux(9).Text, "N") & ", " 'Importe
                        sql = sql & "codigiva= " & DBSet(txtAux(10).Text, "N") & " " 'codigo de iva
                        sql = sql & Replace(ObtenerWhereCP(True), NombreTabla, "facturassocio_envases") & " AND numlinea=" & Adoaux(1).Recordset!numlinea
                        conn.Execute sql
                    End If
                End If
            End If
            Set vCStock = Nothing
                
            ModificaLineas = 0
            
            CalcularDatosFactura
            
            V = Adoaux(1).Recordset.Fields(3) 'el 2 es el nº de llinia
            CargaGrid DataGrid3, Adoaux(1), True

            ' *** si n'hi han tabs ***
            SSTab1.Tab = 1

            DataGrid3.SetFocus
            Adoaux(1).Recordset.Find (Adoaux(1).Recordset.Fields(3).Name & " =" & V)

            LLamaLineas ModificaLineas, 0, "DataGrid3"
        
        End Select
    
    End If
        
eModificarLinea:
    If Err.Number <> 0 Or Not b Then
        MuestraError Err.Number, "Modificar Lineas Factura" & vbCrLf & Err.Description & vbCrLf & Mens
        b = False
    End If
    
    If NumTabMto = 1 Then
        If b Then
            conn.CommitTrans
            ModificarLinea = True
        Else
            conn.RollbackTrans
            ModificarLinea = False
        End If
    Else
        ModificarLinea = b
    End If
End Function
        

Private Function DatosOkLlin(nomframe As String) As Boolean
Dim Rs As ADODB.Recordset
Dim sql As String
Dim b As Boolean
Dim Cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte
Dim Socio As String

    DatosOkLlin = True
    
    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False

    b = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not b Then Exit Function
    
    'en variedades comprobamos que el albaran introducido corresponde al cliente
    Select Case nomframe
        Case "FrameAux0"
            Socio = ""
            Socio = DevuelveDesdeBDNew(cAgro, "albaran", "codsocio", "numalbar", txtAux3(4).Text, "N")
            
            If CLng(Socio) <> CLng(Data1.Recordset!CodSocio) Then
                MsgBox "El albarán introducido no es del socio del la factura. Revise.", vbExclamation
                b = False
            End If
            
'            '++
'            '[Monica]15/02/2011: Problema con el Alt+A
'            TxtAux3_LostFocus (8)
'            TxtAux3_LostFocus (10)
'            '++
        
        Case "FrameAux1"
'        '++
'        '[Monica]15/02/2011: Problema con el Alt+A
'            txtAux_LostFocus (6)
'            txtAux_LostFocus (7)
'            txtAux_LostFocus (8)
'            txtAux_LostFocus (9)
'        '++
        
    End Select
    
    
    DatosOkLlin = b

EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " codtipom = " & DBSet(Text1(6).Text, "T") & " and numfactu= " & Val(Text1(0).Text) & " and fecfactu = " & DBSet(Text1(1).Text, "F")
    ' *******************************************************
    
    ObtenerWhereCab = vWhere
End Function

' ***** si n'hi han varios nivells de tabs *****
Private Sub SituarTab(numTab As Integer)
    On Error Resume Next
    
    SSTab1.Tab = numTab
    
    If Err.Number <> 0 Then Err.Clear
End Sub

' **********************************************
Private Sub PonerDatosSocio(CodSocio As String, Optional nifSocio As String)
Dim vSocio As CSocio
Dim Observaciones As String
    
    On Error GoTo EPonerDatos
    
    If CodSocio = "" Then
        LimpiarDatosSocio
        Exit Sub
    End If

    Set vSocio = New CSocio
    
    'si se ha modificado el cliente volver a cargar los datos
    If vSocio.Existe(CodSocio) Then
        If vSocio.LeerDatos(CodSocio) Then
            Text1(3).Text = vSocio.Codigo
            FormateaCampo Text1(3)
            If (Modo = 3) Or (Modo = 4) Then
                Text2(3).Text = vSocio.Nombre  'Nom clien
                Text1(4).Text = vSocio.ForPago
                Text2(4).Text = PonerNombreDeCod(Text1(4), "forpago", "nomforpa")
                Text1(7).Text = 0 'Format(vSocio.Dto1, FormatoDescuento)
                Text1(8).Text = 0 'Format(vSocio.Dto2, FormatoDescuento)
                Me.Combo1(0).ListIndex = 0 'vSocio.TipoIva  he puesto a todos normal
                
                TipoFactura = 0 'vSocio.TipoFactu
                Text1(6).Text = "FAS"
            End If

            Observaciones = DBLet(vSocio.Observaciones)
            If Trim(Observaciones) <> "" Then
                MsgBox Observaciones, vbInformation, "Observaciones del socio"
            End If
        End If
    Else
        LimpiarDatosSocio
    End If
    Set vSocio = Nothing

EPonerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Datos Cliente", Err.Description
End Sub


Private Sub LimpiarDatosSocio()
Dim i As Byte

    Text1(2).Text = ""
    Text1(4).Text = ""
    Text1(7).Text = ""
    Text1(8).Text = ""

    Text2(3).Text = ""
    Text2(4).Text = ""
    Me.Combo1(0).ListIndex = -1
End Sub
    
Private Sub CargarDatosAlbaran(Albaran As String, Linea As String)
Dim Forfait As String
Dim Pesoneto As String
Dim NumCajas As String
Dim KilosCaja As String
Dim i As Integer
Dim Rs As ADODB.Recordset
Dim sql As String

    If Albaran = "" Or Linea = "" Then
        txtAux3(6).Text = ""
        txtAux3(7).Text = ""
        txtAux3(15).Text = ""
        For i = 0 To Text3.Count - 1
            Text3(i).Text = ""
        Next i
    Else
        sql = "select albaran.fechaalb, albaran.matrirem, forfaits.kiloscaj, albaran_variedad.pesoneto, "
        sql = sql & " albaran_variedad.numcajas, variedades.nomvarie, destinos.nomdesti, forfaits.nomconfe, albaran_variedad.unidades, albaran_variedad.codvarie "
        sql = sql & " from albaran, albaran_variedad, variedades, destinos, forfaits "
        sql = sql & " where albaran_variedad.numalbar = " & DBSet(Albaran, "N")
        sql = sql & " and albaran_variedad.numlinea = " & DBSet(Linea, "N")
        sql = sql & " and albaran.numalbar = albaran_variedad.numalbar "
        sql = sql & " and albaran_variedad.codforfait = forfaitS.codforfait "
        sql = sql & " and albaran_variedad.codvarie = variedades.codvarie "
        sql = sql & " and albaran.codclien = destinos.codclien "
        sql = sql & " and albaran.coddesti = destinos.coddesti "
        
        Set Rs = New ADODB.Recordset
        
        Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs.EOF Then
            txtAux3(6).Text = DBLet(Rs.Fields(3).Value, "N")
            txtAux3(7).Text = DBLet(Rs.Fields(4).Value, "N") * DBLet(Rs.Fields(2).Value, "N")
            txtAux3(15).Text = DBLet(Rs.Fields(8).Value, "N")
            
            '++monica:27/05/08: metemos el codigo iva de la variedad si el cliente es normal
            If Me.Combo1(0).ListIndex = 0 Then
                txtAux3(14).Text = DevuelveDesdeBDNew(cAgro, "variedades", "codigiva", "codvarie", Rs.Fields(9).Value, "N")
            End If
            '++
            
            Text3(0).Text = DBLet(Rs.Fields(0).Value, "F")
            Text3(1).Text = DBLet(Rs.Fields(1).Value, "T")
            Text3(2).Text = DBLet(Rs.Fields(6).Value, "T")
            Text3(3).Text = DBLet(Rs.Fields(5).Value, "T")
            Text3(4).Text = DBLet(Rs.Fields(7).Value, "T")
        Else
            txtAux3(6).Text = ""
            txtAux3(7).Text = ""
            txtAux3(15).Text = ""
            For i = 0 To Text3.Count - 1
                Text3(i).Text = ""
            Next i
            
            MsgBox "No existe la línea de albarán. Reintroduzca.", vbExclamation
            txtAux3(5).Text = ""
'            PonerFoco txtAux3(5)
            
        End If
        
        Set Rs = Nothing
    End If
    
End Sub



Private Function InsertarLineaEnv(numlinea As String) As Boolean
'Inserta un registro en la tabla de lineas de Albaranes: slialb
Dim sql As String
Dim vWhere As String
Dim b As Boolean
Dim vCStock As CStock
Dim DentroTRANS As Boolean

    InsertarLineaEnv = False
    sql = ""
    DentroTRANS = False
    
    'Conseguir el siguiente numero de linea
    vWhere = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
'    numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
    
    Set vCStock = New CStock
    If Not InicializarCStock(vCStock, "S", numlinea) Then Exit Function
    
    If DatosOkLineaEnv(vCStock) Then 'Lineas de factura
        'Inserta en tabla "facturassocio_envases"
        sql = "INSERT INTO facturassocio_envases "
        sql = sql & "(codtipom,numfactu,fecfactu,numlinea,codalmac,codartic,cantidad,precioar,dtolinea,importel,ampliaci,codigiva) "
        sql = sql & "VALUES ('" & txtAux(0).Text & "', " & DBSet(txtAux(1).Text, "N") & ", " & DBSet(txtAux(2).Text, "F") & ", " & numlinea & ", " & DBSet(txtAux(4).Text, "N") & ","
        sql = sql & DBSet(txtAux(5).Text, "T") & ", "
        sql = sql & DBSet(txtAux(6).Text, "N") & ", "
        sql = sql & DBSet(txtAux(7).Text, "N") & ", " & DBSet(txtAux(8).Text, "N") & ", "
        sql = sql & DBSet(txtAux(9).Text, "N") & ","
        sql = sql & DBSet(Text2(16).Text, "T") & ","
        sql = sql & DBSet(txtAux(10).Text, "N") & ")"
     Else
        Exit Function
     End If
    
    If sql <> "" Then
        On Error GoTo EInsertarLineaEnv
        conn.BeginTrans
        DentroTRANS = True
        
        'insertar la linea
        conn.Execute sql
        
        'si hay control de stock para el articulo actualizar en salmac e insertar en smoval
        'en actualizar stock comprobamos si el articulo tiene control de stock
        b = vCStock.ActualizarStock
        
    
    End If
    Set vCStock = Nothing
    
    If b Then
        conn.CommitTrans
        InsertarLineaEnv = True
    Else
        conn.RollbackTrans
         InsertarLineaEnv = False
    End If
    Exit Function
    
EInsertarLineaEnv:
    If Err.Number <> 0 Then
        InsertarLineaEnv = False
        If DentroTRANS Then conn.RollbackTrans
        MuestraError Err.Number, "Insertar Lineas Facturas" & vbCrLf & Err.Description
'        b = False
    End If
'    If b Then
'        Conn.CommitTrans
'        InsertarLinea = True
'    Else
'        Conn.RollbackTrans
'         InsertarLinea = False
'    End If
End Function


'Private Function InsertarLineaFacCta(numlinea As String) As Boolean
''Inserta un registro en la tabla de lineas de Albaranes: slialb
'Dim Sql As String
'Dim vWhere As String
'Dim b As Boolean
'Dim vCStock As CStock
'Dim DentroTRANS As Boolean
'
'    InsertarLineaFacCta = False
'    Sql = ""
'    DentroTRANS = False
'
'    'Conseguir el siguiente numero de linea
'    vWhere = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
''    numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
'
'    If DatosOkLineaFacCta() Then 'Lineas de factura
'        'Inserta en tabla "facturassocio_envases"
'        Sql = "INSERT INTO facturassocio_acuenta "
'        Sql = Sql & "(codtipom,numfactu,fecfactu,codtipomcta,numfactucta, fecfactucta, totalfaccta) "
'        Sql = Sql & "VALUES ('" & txtAux1(0).Text & "', " & DBSet(txtAux1(1).Text, "N") & ", " & DBSet(txtAux1(2).Text, "F") & ", " & DBSet(txtAux(3).Text, "T") & ","
'        Sql = Sql & DBSet(txtAux1(4).Text, "N") & ", "
'        Sql = Sql & DBSet(txtAux1(5).Text, "F") & ", " & DBSet(txtAux1(6).Text, "N") & ") "
'     Else
'        Exit Function
'     End If
'
'    If Sql <> "" Then
'        On Error GoTo EInsertarLineaFacCta
'        conn.BeginTrans
'        DentroTRANS = True
'
'        'insertar la linea
'        conn.Execute Sql
'
'
'    End If
'    Set vCStock = Nothing
'
'    If b Then
'        conn.CommitTrans
'        InsertarLineaFacCta = True
'    Else
'        conn.RollbackTrans
'        InsertarLineaFacCta = False
'    End If
'    Exit Function
'
'EInsertarLineaFacCta:
'    If Err.Number <> 0 Then
'        InsertarLineaFacCta = False
'        If DentroTRANS Then conn.RollbackTrans
'        MuestraError Err.Number, "Insertar Lineas Facturas a Cuenta " & vbCrLf & Err.Description
'    End If
'End Function



Private Function InicializarCStock(ByRef vCStock As CStock, TipoM As String, Optional numlinea As String) As Boolean
    On Error Resume Next

    vCStock.tipoMov = TipoM
    vCStock.DetaMov = Text1(6).Text
    vCStock.Trabajador = CInt(Text1(3).Text) 'guardamos el cliente de la factura
    vCStock.Documento = Text1(0).Text 'Nº Factura
    vCStock.Fechamov = Text1(1).Text 'Fecha de la Factura
    
    '1=Insertar, 2=Modificar
    If ModificaLineas = 1 Or (ModificaLineas = 2 And TipoM = "S") Then
        vCStock.codArtic = txtAux(5).Text
        vCStock.codAlmac = CInt(txtAux(4).Text)
        If ModificaLineas = 1 Then '1=Insertar
            vCStock.Cantidad = CSng(ComprobarCero(txtAux(6).Text))
        Else '2=Modificar(Debe haber en stock la diferencia)
            If Adoaux(1).Recordset!codArtic = txtAux(5).Text Then
                vCStock.Cantidad = CSng(ComprobarCero(txtAux(6).Text)) - Adoaux(1).Recordset!Cantidad
            Else
                vCStock.Cantidad = CSng(ComprobarCero(txtAux(6).Text))
            End If
        End If
        vCStock.Importe = CCur(ComprobarCero(txtAux(9).Text))
    Else
        vCStock.codArtic = Adoaux(1).Recordset!codArtic
        vCStock.codAlmac = CInt(Adoaux(1).Recordset!codAlmac)
        vCStock.Cantidad = CSng(Adoaux(1).Recordset!Cantidad)
        vCStock.Importe = CCur(Adoaux(1).Recordset!ImporteL)
    End If
    If ModificaLineas = 1 Then
        vCStock.LineaDocu = CInt(ComprobarCero(numlinea))
    Else
        vCStock.LineaDocu = CInt(Adoaux(1).Recordset!numlinea)
    End If
    If Err.Number <> 0 Then
        MsgBox "No se han podido inicializar la clase para actualizar Stock", vbExclamation
        InicializarCStock = False
    Else
        InicializarCStock = True
    End If
End Function

Private Function DatosOkLineaEnv(ByRef vCStock As CStock) As Boolean
Dim b As Boolean
Dim i As Byte
    
    On Error GoTo EDatosOkLineaEnv

    DatosOkLineaEnv = False
    b = True

    'Comprobar si se mueve mover stock (hay stock, o si no hay pero no control de stock)
    If vCStock.MueveStock Then
        b = vCStock.MoverStock
    End If
    DatosOkLineaEnv = b
    
EDatosOkLineaEnv:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function



Private Function EliminarStock() As Boolean
Dim sql As String
Dim Rs As ADODB.Recordset
Dim vCStock As CStock
Dim b As Boolean

    On Error GoTo eEliminarStock
    
    sql = "select * from facturassocio_envases where " & Replace(ObtenerWhereCP(False), "facturassocio", "facturassocio_envases")
    Set Rs = New ADODB.Recordset
    
    Rs.Open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    b = True
    While Not Rs.EOF And b
        Set vCStock = New CStock
        
        vCStock.Cantidad = DBLet(Rs!Cantidad, "N")
        vCStock.codAlmac = DBLet(Rs!codAlmac, "N")
        vCStock.codArtic = DBLet(Rs!codArtic, "T")
        vCStock.Documento = Format(DBLet(Rs!NumFactu, "N"), "0000000")
        vCStock.DetaMov = DBLet(Rs!codTipoM, "T")
        vCStock.Fechamov = DBLet(Rs!FecFactu, "F")
        vCStock.Importe = DBLet(Rs!ImporteL, "N")
        vCStock.LineaDocu = DBLet(Rs!numlinea, "N")
        vCStock.tipoMov = "E"
        
        b = vCStock.DevolverStock
        
        Rs.MoveNext
        
        Set vCStock = Nothing
    Wend

    Set Rs = Nothing

eEliminarStock:
    If Err.Number <> 0 Or Not b Then
        EliminarStock = False
    Else
        EliminarStock = True
    End If

End Function


Private Sub CalcularDatosFactura()
Dim i As Integer
Dim cadWHERE As String, sql As String

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For i = 9 To 31
         Text1(i).Text = ""
    Next i
    
    'Comprobar que hay lineas de facturassocio_variedad para calcular totales
    cadWHERE = ObtenerWhereCP(False)
    sql = "Select count(*) from " & NomTablaLineas & " Where " & Replace(cadWHERE, NombreTabla, NomTablaLineas)
    If RegistrosAListar(sql) = 0 Then
        'Comprobar que hay lineas de facturassocio_envases para calcular totales
        sql = "Select count(*) from facturassocio_envases Where " & Replace(cadWHERE, NombreTabla, "facturassocio_envases")
        If RegistrosAListar(sql) = 0 Then
            '[Monica]22/06/2010 añadido por facturassocio_acuenta -- antes sólo: If RegistrosAListar(Sql) = 0 Then exit sub
            'Comprobar que hay lineas de facturassocio_acuenta para calcular totales
'            sql = "Select count(*) from facturassocio_acuenta Where " & Replace(cadwhere, NombreTabla, "facturassocio_acuenta")
'            If RegistrosAListar(sql) = 0 Then Exit Sub
        Else
'            Exit Sub
        End If
    End If
    
    
    If CalcularDatosFacturaVenta(cadWHERE, NombreTabla, NomTablaLineas) Then
        PosicionarData
        PonerCampos
    Else
        MuestraError Err.Number, "Calculando Totales", Err.Description
    End If
'    Set vFactu = Nothing
End Sub

'
'##Monica
'
Private Function CalcularDatosFacturaVenta(cadWHERE As String, NomTabla As String, NomTablaLin As String) As Boolean
'cadWhere: cad para la where de la SQL que selecciona las lineas del albaran o la factura
'nomTabla: nombre de la tabla de albaranes(scaalp) o de AlbaranesXFactura(scafpa)
'           segun llamemos desde recepcion de facturassocio o desde Hco de Facturas
Dim Rs As ADODB.Recordset
Dim i As Integer

Dim sql As String
Dim cadAux As String
Dim cadAux1 As String

'Aqui vamos acumulando los totales
Dim TotBruto As Currency
Dim TotNeto As Currency
Dim TotImpIVA As Currency

Dim ImpAux As Currency
Dim impiva As Currency
Dim ImpREC As Currency
Dim ImpBImIVA As Currency 'Importe Base imponible a la que hay q aplicar el IVA

Dim vBruto As Currency
Dim vNeto As Currency

Dim exentoIVA As Boolean
Dim conDesplaz As Boolean
    
Dim BaseImp As Currency
Dim BaseIVA1 As Currency
Dim BaseIVA2 As Currency
Dim BaseIVA3 As Currency
    
Dim BrutoFac As Currency
    
Dim ImpIVA1 As Currency
Dim ImpIVA2 As Currency
Dim ImpIVA3 As Currency
    
Dim PorceIVA1 As Currency
Dim PorceIVA2 As Currency
Dim PorceIVA3 As Currency
    
Dim ImpREC1 As Currency
Dim ImpREC2 As Currency
Dim ImpREC3 As Currency
    
Dim PorceREC1 As Currency
Dim PorceREC2 As Currency
Dim PorceREC3 As Currency
    
Dim TipoIVA1 As Currency
Dim TipoIVA2 As Currency
Dim TipoIVA3 As Currency
    
Dim ImpDto1 As Currency
Dim ImpDto2 As Currency
Dim TotalFac As Currency

Dim IvaAnt As Integer
Dim cadwhere1 As String
    
Dim Nulo2 As String
Dim Nulo3 As String

    CalcularDatosFacturaVenta = False
    On Error GoTo ECalcular

    BaseImp = 0
    BaseIVA1 = 0
    BaseIVA2 = 0
    BaseIVA3 = 0
    
    BrutoFac = 0
    
    ImpIVA1 = 0
    ImpIVA2 = 0
    ImpIVA3 = 0
    
    PorceIVA1 = 0
    PorceIVA2 = 0
    PorceIVA3 = 0
    
    ImpREC1 = 0
    ImpREC2 = 0
    ImpREC3 = 0
    
    PorceREC1 = 0
    PorceREC2 = 0
    PorceREC3 = 0
    
    TipoIVA1 = 0
    TipoIVA2 = 0
    TipoIVA3 = 0
    
    ImpDto1 = 0
    ImpDto2 = 0
    TotalFac = 0

    'Agrupar el importe bruto por tipos de iva
    cadwhere1 = Replace(cadWHERE, "facturassocio", "facturassocio_variedad")
    sql = "SELECT facturassocio_variedad.codigiva, sum(imporbru) as bruto, sum(impornet) as neto"
    sql = sql & " FROM facturassocio_variedad "
    sql = sql & " WHERE " & cadwhere1
    sql = sql & " GROUP BY 1 "
    sql = sql & " UNION "
    cadwhere1 = Replace(cadWHERE, "facturassocio", "facturassocio_envases")
    sql = sql & "SELECT facturassocio_envases.codigiva, sum(importel) as bruto, sum(importel) as neto"
    sql = sql & " FROM facturassocio_envases "
    sql = sql & " WHERE " & cadwhere1
    sql = sql & " GROUP BY 1 "
'    sql = sql & " UNION "
'    cadWhere1 = Replace(cadwhere, "facturassocio", "facturassocio_acuenta")
'    sql = sql & "SELECT facturassocio.codiiva1 as codigiva, sum(brutofac * (-1)) as bruto, sum(brutofac * (-1)) as neto"
'    sql = sql & " FROM facturassocio_acuenta, facturassocio "
'    sql = sql & " WHERE " & cadWhere1
'    sql = sql & " and facturassocio.codtipom = facturassocio_acuenta.codtipomcta "
'    sql = sql & " and facturassocio.numfactu = facturassocio_acuenta.numfactucta "
'    sql = sql & " and facturassocio.fecfactu = facturassocio_acuenta.fecfactucta "
'    sql = sql & " GROUP BY 1 "
    sql = sql & " ORDER BY 1 "

    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    TotBruto = 0
    TotNeto = 0
    TotImpIVA = 0
    vBruto = 0
    vNeto = 0
    i = 1

    If Not Rs.EOF Then Rs.MoveFirst
    IvaAnt = Rs.Fields(0).Value
    While Not Rs.EOF
        
        If IvaAnt <> Rs.Fields(0).Value Then
            TotBruto = TotBruto + vBruto
            TotNeto = TotNeto + vNeto
            ImpBImIVA = vNeto
        

            'Obtener el % de IVA
            cadAux = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(IvaAnt), "N")

            'aplicar el IVA a la base imponible de ese tipo
            impiva = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
            
            'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
            'los vamos acumulando
            TotImpIVA = TotImpIVA + impiva

            If CInt(Data1.Recordset!TipoIvac) = 2 Then
                'Obtener el % de RECARGO
                cadAux1 = DevuelveDesdeBDNew(cConta, "tiposiva", "porcerec", "codigiva", CStr(IvaAnt), "N")
    
                'aplicar el RECARGO a la base imponible de ese tipo
                ImpREC = CalcularPorcentaje(ImpBImIVA, CCur(cadAux1), 2)
                
                'sumamos todos los RECARGOS para sumarselo a la base imponible total de la factura
                'los vamos acumulando
                TotImpIVA = TotImpIVA + ImpREC
            Else
                cadAux1 = "0"
                ImpREC = 0
            End If


            Select Case i
                Case 1  'IVA 1
                    TipoIVA1 = IvaAnt 'RS!codigiva

                    BaseIVA1 = ImpBImIVA 'BASE IMPONIBLE

                    PorceIVA1 = cadAux '% de IVA

                    'Importe total con IVA
                    ImpIVA1 = impiva
                    
                    PorceREC1 = cadAux1 '% de REC

                    'Importe total con RECARGO
                    ImpREC1 = ImpREC

                Case 2  'IVA 2
                    TipoIVA2 = IvaAnt 'RS!codigiva

                    BaseIVA2 = ImpBImIVA 'BASE IMPONIBLE

                    PorceIVA2 = cadAux '% de IVA

                    'Importe total con IVA
                    ImpIVA2 = impiva

                    PorceREC2 = cadAux1 '% de REC

                    'Importe total con RECARGO
                    ImpREC2 = ImpREC
                Case 3  'IVA 3
                    TipoIVA3 = IvaAnt 'RS!codigiva

                    BaseIVA3 = ImpBImIVA 'BASE IMPONIBLE

                    PorceIVA3 = cadAux '% de IVA

                    'Importe total con IVA
                    ImpIVA3 = impiva
                    
                    PorceREC3 = cadAux1 '% de REC

                    'Importe total con RECARGO
                    ImpREC3 = ImpREC
            End Select
            
            
            i = i + 1
            IvaAnt = Rs.Fields(0).Value
            vBruto = DBLet(Rs.Fields(1).Value, "N")
            vNeto = DBLet(Rs.Fields(2).Value, "N")
        Else
            vBruto = vBruto + DBLet(Rs.Fields(1).Value, "N")
            vNeto = vNeto + DBLet(Rs.Fields(2).Value, "N")
        End If
        
        
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing

    ' ULTIMO REGISTRO
    TotBruto = TotBruto + vBruto
    TotNeto = TotNeto + vNeto
    ImpBImIVA = vNeto


    'Obtener el % de IVA
    cadAux = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(IvaAnt), "N")

    'aplicar el IVA a la base imponible de ese tipo
    impiva = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
    
    'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
    'los vamos acumulando
    TotImpIVA = TotImpIVA + impiva
    
    If CInt(Data1.Recordset!TipoIvac) = 2 Then
        'Obtener el % de RECARGO
        cadAux1 = DevuelveDesdeBDNew(cConta, "tiposiva", "porcerec", "codigiva", CStr(IvaAnt), "N")
    
        'aplicar el RECARGO a la base imponible de ese tipo
        ImpREC = CalcularPorcentaje(ImpBImIVA, CCur(cadAux1), 2)
    Else
        cadAux1 = "0"
        ImpREC = 0
    End If
    'sumamos todos los RECARGOS para sumarselo a la base imponible total de la factura
    'los vamos acumulando
    TotImpIVA = TotImpIVA + ImpREC



    Select Case i
        Case 1  'IVA 1
            TipoIVA1 = IvaAnt

            BaseIVA1 = ImpBImIVA 'BASE IMPONIBLE

            PorceIVA1 = cadAux '% de IVA

            'Importe total con IVA
            ImpIVA1 = impiva
            
            PorceREC1 = cadAux1 '% de REC

            'Importe total con RECARGO
            ImpREC1 = ImpREC

        Case 2  'IVA 2
            TipoIVA2 = IvaAnt

            BaseIVA2 = ImpBImIVA 'BASE IMPONIBLE

            PorceIVA2 = cadAux '% de IVA

            'Importe total con IVA
            ImpIVA2 = impiva

            PorceREC2 = cadAux1 '% de REC

            'Importe total con RECARGO
            ImpREC2 = ImpREC
        Case 3  'IVA 3
            TipoIVA3 = IvaAnt

            BaseIVA3 = ImpBImIVA 'BASE IMPONIBLE

            PorceIVA3 = cadAux '% de IVA

            'Importe total con IVA
            ImpIVA3 = impiva
            
            PorceREC3 = cadAux1 '% de REC

            'Importe total con RECARGO
            ImpREC3 = ImpREC
    End Select

    'Base Imponible
    BaseImp = TotNeto

    'TOTAL de la factura
    TotalFac = BaseImp + TotImpIVA

    'ACTUALIZAMOS LA FACTURA (tabla facturassocio)
    sql = "update facturassocio "
    sql = sql & "set baseimp1 = " & DBSet(BaseIVA1, "N")
    sql = sql & ",impoiva1 = " & DBSet(ImpIVA1, "N")
    sql = sql & ",imporec1 = " & DBSet(ImpREC1, "N")
    sql = sql & ",porciva1 = " & DBSet(PorceIVA1, "N")
    sql = sql & ",porcrec1 = " & DBSet(PorceREC1, "N")
    sql = sql & ",codiiva1 = " & DBSet(TipoIVA1, "N")
    Nulo2 = "N"
    Nulo3 = "N"
    If DBSet(TipoIVA2, "N", "S") = ValorNulo Then Nulo2 = "S"
    If DBSet(TipoIVA3, "N", "S") = ValorNulo Then Nulo3 = "S"
    sql = sql & ",baseimp2 = " & DBSet(BaseIVA2, "N", Nulo2)
    sql = sql & ",impoiva2 = " & DBSet(ImpIVA2, "N", Nulo2)
    sql = sql & ",imporec2 = " & DBSet(ImpREC2, "N", Nulo2)
    sql = sql & ",porciva2 = " & DBSet(PorceIVA2, "N", Nulo2)
    sql = sql & ",porcrec2 = " & DBSet(PorceREC2, "N", Nulo2)
    sql = sql & ",codiiva2 = " & DBSet(TipoIVA2, "N", Nulo2)
    sql = sql & ",baseimp3 = " & DBSet(BaseIVA3, "N", Nulo3)
    sql = sql & ",impoiva3 = " & DBSet(ImpIVA3, "N", Nulo3)
    sql = sql & ",imporec3 = " & DBSet(ImpREC3, "N", Nulo3)
    sql = sql & ",porciva3 = " & DBSet(PorceIVA3, "N", Nulo3)
    sql = sql & ",porcrec3 = " & DBSet(PorceREC3, "N", Nulo3)
    sql = sql & ",codiiva3 = " & DBSet(TipoIVA3, "N", Nulo3)
    sql = sql & ",brutofac = " & DBSet(TotBruto, "N")
    sql = sql & ",impordto = " & DBSet(Round2(TotBruto - TotNeto, 2), "N")
    sql = sql & ",totalfac = " & DBSet(TotalFac, "N")
    sql = sql & " where " & cadWHERE
    
    conn.Execute sql

    CalcularDatosFacturaVenta = True

ECalcular:
    If Err.Number <> 0 Then
        CalcularDatosFacturaVenta = False
    Else
        CalcularDatosFacturaVenta = True
    End If
End Function


Private Function CalcularImporteDto(Cantidad As String, Precio As String, TipoM As String, Factura As String, FecFactu As String, ImpDto As String, Insertado As Boolean) As String
'Insertado: indica si ya hemos insertado el registro o no
'Calcula el Importe de una linea de Oferta, Pedido, Albaran, ...
'Importe=cantidad * precio - (descuentos)
Dim vCant As Currency
Dim vImp As Currency
Dim vDto As Currency
Dim vPre As Currency
Dim Rs As ADODB.Recordset
Dim sql As String
Dim SumaBruto As Currency

On Error Resume Next


    If vParamAplic.Cooperativa = 5 Then
        ' Para Castelduc el importe de descuento se prorratea con respecto a los kilos
        
        sql = "select sum(cantreal) from facturassocio_variedad where codtipom = " & DBSet(TipoM, "T")
        sql = sql & " and numfactu = " & DBSet(Factura, "N") & " and fecfactu = " & DBSet(FecFactu, "F")
        
        Set Rs = New ADODB.Recordset
        Rs.Open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        SumaBruto = 0
        If Not Rs.EOF Then
            SumaBruto = DBLet(Rs.Fields(0).Value, "N")
        End If
        
        'Como son de tipo string comprobar que si vale "" lo ponemos a 0
        vCant = ComprobarCero(Cantidad)
    '    vPre = ComprobarCero(precio)
        vDto = ComprobarCero(ImpDto)
        
        If Not Insertado Then '++monica 030608 añadido el round
    '        SumaBruto = SumaBruto + Round2(CCur(vCant) * CCur(vPre), 2)
            SumaBruto = SumaBruto + CCur(vCant)
        End If
        
        If SumaBruto <> 0 Then '++monica 030608 añadido el round
    '        vImp = Round2((CCur(cantidad) * CCur(vPre) * CCur(vDto)) / SumaBruto, 4)
            vImp = Round2((CCur(Cantidad) * CCur(vDto)) / SumaBruto, 4)
        Else
            vImp = CCur(vDto)
        End If
        
        vImp = Round2(vImp, 6)
        
        CalcularImporteDto = CStr(vImp)
    
    
    Else
        ' como lo ha estado haciendo hasta ahora
        '(se prorratea el importe de descuento sobre el importe bruto de la linea)
    
        sql = "select sum(imporbru) from facturassocio_variedad where codtipom = " & DBSet(TipoM, "T")
        sql = sql & " and numfactu = " & DBSet(Factura, "N") & " and fecfactu = " & DBSet(FecFactu, "F")
        
        Set Rs = New ADODB.Recordset
        Rs.Open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        SumaBruto = 0
        If Not Rs.EOF Then
            SumaBruto = DBLet(Rs.Fields(0).Value, "N")
        End If
        
        'Como son de tipo string comprobar que si vale "" lo ponemos a 0
        vCant = ComprobarCero(Cantidad)
    '    vPre = ComprobarCero(precio)
        vDto = ComprobarCero(ImpDto)
        
        If Not Insertado Then '++monica 030608 añadido el round
    '        SumaBruto = SumaBruto + Round2(CCur(vCant) * CCur(vPre), 2)
            SumaBruto = SumaBruto + Round2(CCur(vCant) * ComprobarCero(Precio), 2)
        End If
        
        If SumaBruto <> 0 Then '++monica 030608 añadido el round
    '        vImp = Round2((CCur(cantidad) * CCur(vPre) * CCur(vDto)) / SumaBruto, 4)
            vImp = Round2((CCur(Cantidad) * ComprobarCero(Precio) * CCur(vDto)) / SumaBruto, 4)
        Else
            vImp = CCur(vDto)
        End If
        
        vImp = Round2(vImp, 6)
        
        CalcularImporteDto = CStr(vImp)

    End If

End Function


Private Function RecalcularDtos(TipoM As String, Factura As String, FecFactu As String, MenError As String) As Boolean
Dim Rs As ADODB.Recordset
Dim sql As String
Dim Sql2 As String
Dim vImpDto As Currency
Dim vDto1 As Currency
Dim vDto2 As Currency
Dim vImpNeto As Currency
Dim vPrecNeto As Currency
Dim TipoDto As String
Dim ImpDto As String
Dim Cliente As String
Dim Rdo As Long

    On Error GoTo eRecalcularDtos

    sql = "select * from facturassocio_variedad where codtipom = " & DBSet(TipoM, "T")
    sql = sql & " and numfactu = " & DBSet(Factura, "N") & " and fecfactu = " & DBSet(FecFactu, "F")
    sql = sql & " order by numlinea "
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    sql = ""
    sql = DevuelveDesdeBDNew(cAgro, "facturassocio", "impdtoc", "codtipom", TipoM, "T", , "numfactu", Factura, "N", "fecfactu", FecFactu, "F")
    vImpDto = ComprobarCero(sql)
    
    sql = ""
    sql = DevuelveDesdeBDNew(cAgro, "facturassocio", "dtocom1", "codtipom", TipoM, "T", , "numfactu", Factura, "N", "fecfactu", FecFactu, "F")
    vDto1 = ComprobarCero(sql)
    
    sql = ""
    sql = DevuelveDesdeBDNew(cAgro, "facturassocio", "dtocom2", "codtipom", TipoM, "T", , "numfactu", Factura, "N", "fecfactu", FecFactu, "F")
    vDto2 = ComprobarCero(sql)
    
    '++monica:030608:traemos el redondeo del precio
'    sql = ""
'    sql = DevuelveDesdeBDNew(cAgro, "facturassocio", "codclien", "codtipom", TipoM, "T", , "numfactu", Factura, "N", "fecfactu", FecFactu, "F")
'    Cliente = ComprobarCero(sql)
'    sql = ""
'    sql = DevuelveDesdeBDNew(cAgro, "clientes", "nrodecprec", "codclien", Cliente, "N")
    Rdo = 4 'ComprobarCero(sql)
    
    While Not Rs.EOF
'        TipoDto = DevuelveDesdeBDNew(cAgro, "clientes", "tipodtos", "codclien", Text1(3).Text, "N")
        TipoDto = 0
        If TipoFacturarForfaits(CStr(Rs!NumAlbar), CStr(Rs!numlinealbar)) = 1 Then 'kilos
            ImpDto = CalcularImporteDto(DBLet(Rs!cantreal, "N"), DBLet(Rs!precibru, "N"), TipoM, Factura, FecFactu, CStr(vImpDto), True)
            vImpNeto = CalcularImporteFClien(DBLet(Rs!cantreal, "N"), DBLet(Rs!precibru, "N"), CStr(vDto1), CStr(vDto2), CByte(TipoDto), CStr(ImpDto), DBLet(Rs!imporbru, "N"))
            
            'precio neto
            vPrecNeto = Round2(vImpNeto / DBLet(Rs!cantreal, "N"), Rdo)
            
            '++monica:040608 : solo si redondeo <> 4
            If Rdo = 2 Or Rdo = 3 Then
                vImpNeto = Round2(vPrecNeto * DBLet(Rs!cantreal, "N"), 2)
            End If
        Else 'unidades
            ImpDto = CalcularImporteDto(DBLet(Rs!Unidades, "N"), DBLet(Rs!precibru, "N"), TipoM, Factura, FecFactu, CStr(vImpDto), True)
            vImpNeto = CalcularImporteFClien(DBLet(Rs!Unidades, "N"), DBLet(Rs!precibru, "N"), CStr(vDto1), CStr(vDto2), CByte(TipoDto), CStr(ImpDto), DBLet(Rs!imporbru, "N"))
            
            'precio neto
            vPrecNeto = Round2(vImpNeto / DBLet(Rs!Unidades, "N"), Rdo)
            
            '++monica:040608
            If Rdo = 2 Or Rdo = 3 Then
                vImpNeto = Round2(vPrecNeto * DBLet(Rs!Unidades, "N"), 2)
            End If
        End If
        
        Sql2 = "update facturassocio_variedad set impornet = " & DBSet(vImpNeto, "N")
        Sql2 = Sql2 & ",precinet = " & DBSet(vPrecNeto, "N")
        Sql2 = Sql2 & ",dtocom1 = " & DBSet(vDto1, "N")
        Sql2 = Sql2 & ",dtocom2 = " & DBSet(vDto2, "N")
        Sql2 = Sql2 & " where codtipom = " & DBSet(TipoM, "T")
        Sql2 = Sql2 & " and numfactu = " & DBSet(Factura, "N")
        Sql2 = Sql2 & " and fecfactu = " & DBSet(FecFactu, "F")
        Sql2 = Sql2 & " and numlinea = " & DBSet(Rs!numlinea, "N")
    
        conn.Execute Sql2
    
        Rs.MoveNext
    Wend

    Set Rs = Nothing
    
    RecalcularDtos = True
    Exit Function

eRecalcularDtos:
    If Err.Number <> 0 Then
        MenError = MenError & vbCrLf & Err.Description
        RecalcularDtos = False
    End If
End Function




'Private Sub DescontarFacturasACuenta(TipoM As String, Factu As String, fecFac As String, Cliente As String)
'Dim sql As String
'Dim cadwhere As String
'
'    sql = "select codtipom, numfactu, fecfactu, totalfac from facturassocio "
'    cadwhere = "where codtipom = 'EAC' and codclien = " & DBSet(Cliente, "N")
'    cadwhere = cadwhere & " and (codtipom, numfactu, fecfactu) not in (select codtipomcta, numfactucta, fecfactucta from facturassocio_acuenta) "
''    cadWHERE = cadWHERE & " where codclien = " & DBSet(Cliente, "N") & " and codtipom = 'EAC')"
'
'    sql = sql & cadwhere
'
'    If TotalRegistrosConsulta(sql) <> 0 Then
'
'        Set frmMens = New frmMensajes
'
'        frmMens.OpcionMensaje = 22
'        frmMens.cadwhere = cadwhere
'
'        frmMens.Show vbModal
'
'        Set frmMens = Nothing
'
''    Else
''        MsgBox "No hay facturassocio a cuenta para descontar.", vbExclamation
'    End If
'
'End Sub
