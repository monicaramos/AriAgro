VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmVtasFacturasCta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturas a Cuenta"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   7155
   Icon            =   "frmVtasFacturasCta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmVtasFacturasCta.frx":000C
   ScaleHeight     =   7605
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   6525
      Left            =   150
      TabIndex        =   29
      Top             =   480
      Width           =   6795
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   20
         Left            =   1305
         MaxLength       =   6
         TabIndex        =   5
         Tag             =   "T.Mercado|N|S|0|999|facturas|codtimer|000||"
         Text            =   "Text1"
         Top             =   1830
         Width           =   780
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   20
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   59
         Text            =   "Text2"
         Top             =   1830
         Width           =   4335
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   19
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   57
         Text            =   "Text2"
         Top             =   1440
         Width           =   4335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   19
         Left            =   1305
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "Variedad|N|S|0|999999|facturas|codvarie|000000||"
         Text            =   "Text1"
         Top             =   1440
         Width           =   780
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   0
         Left            =   1320
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Text            =   "Text1"
         Top             =   270
         Width           =   915
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   56
         Tag             =   "Tipo Movimiento|T|N|||facturas|codtipom||S|"
         Text            =   "EAC"
         Top             =   270
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   5430
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Factura|F|N|||facturas|fecfactu|dd/mm/yyyy|S|"
         Top             =   300
         Width           =   1065
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   225
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "Tipo Iva|N|N|||facturas|tipoivac||N|"
         Top             =   2460
         Width           =   1440
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Contabilizado"
         Height          =   195
         Index           =   1
         Left            =   5040
         TabIndex        =   10
         Tag             =   "Contabilizado|N|N|||facturas|intconta|0||"
         Top             =   2310
         Width           =   1275
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Pasa Aridoc"
         Height          =   195
         Index           =   0
         Left            =   5040
         TabIndex        =   11
         Tag             =   "Pasa Aridoc|N|N|||facturas|pasaridoc|0||"
         Top             =   2580
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   1305
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "Forma Pago|N|N|0|999|facturas|codforpa|000||"
         Text            =   "Text1"
         Top             =   1050
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
         TabIndex        =   37
         Text            =   "Text2"
         Top             =   1050
         Width           =   4335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   8
         Left            =   2700
         MaxLength       =   5
         TabIndex        =   8
         Tag             =   "%Dto 2|N|S|0|100|facturas|dtocom2|##0.00||"
         Top             =   2460
         Width           =   945
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   7
         Left            =   1710
         MaxLength       =   5
         TabIndex        =   7
         Tag             =   "%Dto 1|N|S|0|100|facturas|dtocom1|##0.00||"
         Top             =   2460
         Width           =   945
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   33
         Text            =   "Text2"
         Top             =   660
         Width           =   4335
      End
      Begin VB.TextBox Text1 
         Height          =   630
         Index           =   2
         Left            =   225
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Tag             =   "Observaciones|T|S|||facturas|observac|||"
         Top             =   3090
         Width           =   6255
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   1305
         MaxLength       =   6
         TabIndex        =   2
         Tag             =   "Cod. Cliente|N|N|0|999999|facturas|codclien|000000||"
         Text            =   "Text1"
         Top             =   660
         Width           =   780
      End
      Begin VB.Frame FrameFactura 
         Height          =   2655
         Left            =   150
         TabIndex        =   40
         Top             =   3720
         Width           =   6450
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   17
            Left            =   4230
            MaxLength       =   5
            TabIndex        =   19
            Tag             =   "% REC 1|N|S|0|99.90|facturas|porcrec1|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   18
            Left            =   4815
            MaxLength       =   15
            TabIndex        =   20
            Tag             =   "Importe REC 1|N|S|0||facturas|imporec1|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1560
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
            Index           =   5
            Left            =   3960
            MaxLength       =   15
            TabIndex        =   21
            Tag             =   "Total Factura|N|S|0||facturas|totalfac|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   2055
            Width           =   2325
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   16
            Left            =   2730
            MaxLength       =   15
            TabIndex        =   18
            Tag             =   "Importe IVA 1|N|S|0||facturas|impoiva1|#,###,###,##0.00|N|"
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
            TabIndex        =   17
            Tag             =   "% IVA 1|N|S|0|99.90|facturas|porciva1|#0.00|N|"
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
            TabIndex        =   16
            Tag             =   "Base Imponible 1|N|S|0||facturas|baseimp1|#,###,###,##0.00|N|"
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
            Tag             =   "IVA 1|N|S|0|9|facturas|codiiva1|0|N|"
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   500
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   9
            Left            =   630
            MaxLength       =   15
            TabIndex        =   12
            Tag             =   "Bruto Factura|N|S|0||facturas|brutofac|#,###,###,##0.00|N|"
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
            Tag             =   "Importe Descuento|N|S|0||facturas|impordto|#,###,###,##0.00|N|"
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
            TabIndex        =   53
            Top             =   1350
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "% Rec"
            Height          =   255
            Index           =   15
            Left            =   4230
            TabIndex        =   52
            Top             =   1350
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "% IVA"
            Height          =   255
            Index           =   41
            Left            =   2145
            TabIndex        =   51
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
            TabIndex        =   50
            Top             =   2115
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
            TabIndex        =   49
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
            TabIndex        =   48
            Top             =   1035
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Importe IVA"
            Height          =   255
            Index           =   33
            Left            =   2730
            TabIndex        =   47
            Top             =   1350
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   13
            Left            =   645
            TabIndex        =   46
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Cod."
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   45
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Bruto Factura"
            Height          =   255
            Index           =   5
            Left            =   630
            TabIndex        =   44
            Top             =   270
            Width           =   1110
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   7
            Left            =   4815
            TabIndex        =   43
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
            TabIndex        =   42
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
            TabIndex        =   41
            Top             =   540
            Width           =   135
         End
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1020
         ToolTipText     =   "Buscar tipo mercado"
         Top             =   1875
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "T.Mercado"
         Height          =   255
         Index           =   6
         Left            =   210
         TabIndex        =   60
         Top             =   1875
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Variedad"
         Height          =   255
         Index           =   3
         Left            =   210
         TabIndex        =   58
         Top             =   1485
         Width           =   690
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1020
         ToolTipText     =   "Buscar Variedad"
         Top             =   1485
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Iva"
         Height          =   255
         Index           =   14
         Left            =   225
         TabIndex        =   39
         Top             =   2190
         Width           =   810
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1020
         ToolTipText     =   "Buscar Destino"
         Top             =   1095
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F.Pago"
         Height          =   255
         Index           =   8
         Left            =   210
         TabIndex        =   38
         Top             =   1095
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "%Dto 2"
         Height          =   255
         Index           =   4
         Left            =   2700
         TabIndex        =   36
         Top             =   2235
         Width           =   1050
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   29
         Left            =   4290
         TabIndex        =   35
         Top             =   330
         Width           =   705
      End
      Begin VB.Label Label1 
         Caption         =   "%Dto 1"
         Height          =   255
         Index           =   2
         Left            =   1710
         TabIndex        =   34
         Top             =   2235
         Width           =   1140
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   5100
         Picture         =   "frmVtasFacturasCta.frx":0A0E
         ToolTipText     =   "Buscar fecha"
         Top             =   315
         Width           =   240
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   1350
         ToolTipText     =   "Zoom descripción"
         Top             =   2850
         Width           =   240
      End
      Begin VB.Label Label29 
         Caption         =   "Observaciones"
         Height          =   255
         Left            =   225
         TabIndex        =   32
         Top             =   2850
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   0
         Left            =   195
         TabIndex        =   31
         Top             =   660
         Width           =   540
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1020
         ToolTipText     =   "Buscar Cliente"
         Top             =   660
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Factura"
         Height          =   255
         Index           =   28
         Left            =   210
         TabIndex        =   30
         Top             =   285
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   150
      TabIndex        =   25
      Top             =   7020
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
         TabIndex        =   26
         Top             =   180
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5880
      TabIndex        =   24
      Top             =   7140
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4740
      TabIndex        =   22
      Top             =   7140
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   7155
      _ExtentX        =   12621
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
         Left            =   5760
         TabIndex        =   28
         Top             =   90
         Width           =   1125
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   5910
      TabIndex        =   23
      Top             =   7140
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
      Left            =   600
      Top             =   7050
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
      Left            =   480
      Top             =   7020
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
      Left            =   5340
      MaxLength       =   15
      TabIndex        =   54
      Text            =   "Text1 7"
      Top             =   1740
      Width           =   1485
   End
   Begin VB.Label Label1 
      Caption         =   "Imp.Descuento 2"
      Height          =   255
      Index           =   10
      Left            =   5610
      TabIndex        =   55
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
Attribute VB_Name = "frmVtasFacturasCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)


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
Private WithEvents frmVar As frmManVariedad 'Form Mto de variedades
Attribute frmVar.VB_VarHelpID = -1

Private WithEvents frmCli As frmClientes 'Form Mto de Clientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmFPag As frmManFpago 'Form Mto de Formas de Pago
Attribute frmFPag.VB_VarHelpID = -1
Private WithEvents frmMer As frmManTipMerc 'Form Mto de Tipos de Mercado
Attribute frmMer.VB_VarHelpID = -1



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


Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos
Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal
Dim indice As Byte
Dim ClienAnt As String

Dim TipoFactura As Byte
Private BuscaChekc As String

Private Sub check1_Click(Index As Integer)
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
Dim I As Integer

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda

        Case 3  'AÑADIR
            If DatosOk Then InsertarCabecera

        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaCabecera Then
                    espera 0.2
                    TerminaBloquear
                    PosicionarData
                    PonerCampos
                End If
            End If
            
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
            PonerModo 0
            PonerFoco text1(3)
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco text1(3)
            
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
    
    text1(1).Text = Format(Now, "dd/mm/yyyy")
        
    
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
        PonerFoco text1(3)
        text1(3).BackColor = vbYellow
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
        MandaBusquedaPrevia "facturas.codtipom='" & CodTipoMov & "'"
    Else
        LimpiarCampos
        CadenaConsulta = "Select facturas.* "
        CadenaConsulta = CadenaConsulta & "from " & NombreTabla
        CadenaConsulta = CadenaConsulta & " WHERE facturas.codtipom='" & CodTipoMov & "' " & Ordenacion
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
    
    If FacturaDescontada(Data1.Recordset!codTipoM, Data1.Recordset!Numfactu, Data1.Recordset!fecfactu) Then
        MsgBox "Esta factura a cuenta ya ha sido descontada. No se puede modificar.", vbExclamation
        Exit Sub
    End If
    
    ClienAnt = CStr(Data1.Recordset!CodClien)
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    PonerFoco text1(1) '*** 1r camp visible que siga PK ***
        
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
    
    If FacturaDescontada(Data1.Recordset!codTipoM, Data1.Recordset!Numfactu, Data1.Recordset!fecfactu) Then
        MsgBox "Esta factura a cuenta ya ha sido descontada. No se puede eliminar.", vbExclamation
        Exit Sub
    End If
    
    
    
'    'solo se puede modificar la factura si no esta contabilizada
'    If FactContabilizada Then Exit Sub
    
    cad = "Cabecera de Facturas." & vbCrLf
    cad = cad & "-------------------------------------      " & vbCrLf & vbCrLf
    cad = cad & "Va a eliminar la Factura:            "
    cad = cad & vbCrLf & "Tipo:  " & text1(6).Text
    cad = cad & vbCrLf & "Nº Factura:  " & Format(text1(0).Text, "0000000")
    cad = cad & vbCrLf & "Fecha:  " & Format(text1(1).Text, "dd/mm/yyyy")

    cad = cad & vbCrLf & vbCrLf & " ¿Desea Eliminarla? "

    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
'        On Error GoTo EEliminar
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
Dim I As Integer

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
    
   'IMAGES para zoom
    For I = 0 To Me.imgZoom.Count - 1
        Me.imgZoom(I).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    Next I
    
    LimpiarCampos   'Limpia los campos TextBox
    CargaCombo

    CodTipoMov = "EAC"
    
    VieneDeBuscar = False
    
        
    '## A mano
    NombreTabla = "facturas"
    
    '[Monica]03/02/2012: cambiamos el orden de busqueda
    If vParamAplic.Cooperativa <> 2 Then
        Ordenacion = " ORDER BY facturas.codtipom, facturas.numfactu, facturas.fecfactu"
    Else
        Ordenacion = " ORDER BY facturas.fecfactu, facturas.numfactu, facturas.codtipom"
    End If
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    CadenaConsulta = "select * from facturas "
    If hcoCodMovim <> "" Then
        CadenaConsulta = CadenaConsulta & " where codtipom = '" & CodTipoMov & "' and numfactu = " & DBSet(hcoCodMovim, "N") & " and fecfactu = " & DBSet(hcoFechaMov, "F")
    Else
        CadenaConsulta = CadenaConsulta & " where codtipom = '" & CodTipoMov & "' and numfactu = -1 "
    End If
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
    
    
'    If DatosADevolverBusqueda <> "" Then
'        Text1(0).Text = DatosADevolverBusqueda
'        HacerBusqueda
'        SSTab1.Tab = 1
'    Else
'        PonerModo 0
'    End If
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
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

Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        CadB = ""
        Aux = ValorDevueltoFormGrid(text1(6), CadenaDevuelta, 1)
        CadB = Aux
        Aux = ValorDevueltoFormGrid(text1(0), CadenaDevuelta, 2)
        CadB = CadB & " and  " & Aux
        Aux = ValorDevueltoFormGrid(text1(1), CadenaDevuelta, 3)
        CadB = CadB & " and " & Aux
        
        
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
    text2(indice).Text = RecuperaValor(CadenaSeleccion, 3) 'Nombre del cliente
    PonerFoco text1(indice)
End Sub

Private Sub frmFPag_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
    text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00") 'Codigo
    text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nombre
End Sub

Private Sub frmMer_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
    text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Codigo
    text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nombre
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     text1(indice).Text = vCampo
End Sub

Private Sub imgBuscar_Click(Index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. de Cliente
            indice = 3
            PonerFoco text1(indice)
            Set frmCli = New frmClientes
            frmCli.DatosADevolverBusqueda = "0|1|2|"
            frmCli.Show vbModal
            Set frmCli = Nothing
            PonerFoco text1(indice)
        
        Case 1 'Forma de Pago
            indice = 4
            PonerFoco text1(indice)
            Set frmFPag = New frmManFpago
            frmFPag.DatosADevolverBusqueda = "0|1|"
            frmFPag.Show vbModal
            Set frmFPag = Nothing
            PonerFoco text1(indice)
            
        Case 2 'Variedad
            indice = 19
            PonerFoco text1(indice)
            Set frmVar = New frmManVariedad
            frmVar.DatosADevolverBusqueda = "0|1|"
            frmVar.Show vbModal
            Set frmVar = Nothing
            PonerFoco text1(indice)
            
        Case 3 ' tipo de mercado
            indice = 20
            Set frmMer = New frmManTipMerc
            frmMer.DatosADevolverBusqueda = "0|1|2|"
            frmMer.Show vbModal
            Set frmMer = Nothing
            PonerFoco text1(indice)
            
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
        indice = 2
        frmZ.pTitulo = "Observaciones de la Factura"
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

Private Sub mnImprimir_Click()
'Imprimir Factura
    
    If Data1.Recordset.EOF Then Exit Sub
    
    BotonImprimir
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub


Private Sub mnModificar_Click()
    'bloquea la tabla cabecera de factura: scafac
    If BLOQUEADesdeFormulario(Me) Then
        'bloquear la tabla cabecera de albaranes de la factura: scafac1
        BotonModificar
    End If
End Sub


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




'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 2 Or (Index = 2 And text1(2).Text = "") Then KEYpress KeyAscii
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
Dim nRegs As Long
Dim vTipoMov As CTiposMov
        
    If Not PerderFocoGnral(text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1 'Fecha factura
            If text1(Index).Text <> "" Then PonerFormatoFecha text1(Index)
            
        Case 3 'Cliente
            If PonerFormatoEntero(text1(Index)) Then
                If Modo = 1 Then
                    text2(Index).Text = PonerNombreDeCod(text1(Index), "clientes", "nomclien")
                Else
                    If Modo = 4 And text1(Index).Text = ClienAnt Then Exit Sub
                    PonerDatosCliente (text1(Index).Text)
                    If text2(Index).Text = "" Then
                        cadMen = "No existe el Cliente: " & text1(Index).Text & vbCrLf
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
                    End If
                    If text1(Index).Text <> "" Then CalcularDatosFacturaVenta
               End If
            End If
                
            
        Case 4 ' Forma de Pago
            If PonerFormatoEntero(text1(Index)) Then
                text2(Index).Text = PonerNombreDeCod(text1(Index), "forpago", "nomforpa")
                If text2(Index).Text = "" Then
                    cadMen = "No existe la Forma de Pago: " & text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmFPag = New frmManFpago
                        frmFPag.DatosADevolverBusqueda = "0|1|"
                        text1(Index).Text = ""
                        TerminaBloquear
                        frmFPag.Show vbModal
                        Set frmFPag = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        text1(Index).Text = ""
                    End If
                    PonerFoco text1(Index)
                End If
            End If
            
        Case 19 ' variedad
            If PonerFormatoEntero(text1(Index)) Then
                text2(Index).Text = PonerNombreDeCod(text1(Index), "variedades", "nomvarie")
                If text2(Index).Text = "" Then
                    cadMen = "No existe la Variedad: " & text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmVar = New frmManVariedad
                        frmVar.DatosADevolverBusqueda = "0|1|"
                        text1(Index).Text = ""
                        TerminaBloquear
                        frmVar.Show vbModal
                        Set frmVar = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        text1(Index).Text = ""
                    End If
                    PonerFoco text1(Index)
                End If
            End If
         
        Case 20 'Tipo de Mercado
            If PonerFormatoEntero(text1(Index)) Then
                text2(Index).Text = PonerNombreDeCod(text1(Index), "tipomer", "nomtimer")
                If text2(Index).Text = "" Then
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
         
         Case 7, 8 'descuentos
            If Modo = 1 Then Exit Sub
            If text1(Index).Text <> "" Then PonerFormatoDecimal text1(Index), 4
            CalcularDatosFacturaVenta
            
         Case 9 ' bruto de la factura
            If Modo = 1 Then Exit Sub
            If text1(Index).Text <> "" Then PonerFormatoDecimal text1(Index), 3
    
            CalcularDatosFacturaVenta

    End Select
End Sub


Private Sub HacerBusqueda()
Dim CadB As String
Dim cadAux As String
    
    CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    If CadB <> "" Then
        CadB = CadB & " and facturas.codtipom = '" & CodTipoMov & "'"
    Else
        CadB = "facturas.codtipom = '" & CodTipoMov & "'"
    End If

    If chkVistaPrevia = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select facturas.* from " & NombreTabla
        CadenaConsulta = CadenaConsulta & " WHERE " & CadB & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim Tabla As String
Dim Titulo As String
Dim Desc As String, devuelve As String
    'Llamamos a al form
    '##A mano
    cad = ""
    cad = cad & "Tipo|facturas.codtipom|N||10·"
    cad = cad & "Nº.Factura|facturas.numfactu|N||15·"
    cad = cad & "Cliente|facturas.codclien|N||10·" 'ParaGrid(Text1(3), 10, "Cliente")
    cad = cad & "Nombre Cliente|clientes.nomclien|N||45·"
    cad = cad & ParaGrid(text1(1), 15, "F.Factura")
    Tabla = NombreTabla & " INNER JOIN clientes ON facturas.codclien=clientes.codclien "
    
    Titulo = "Facturas"
    devuelve = "0|1|4|"
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vtabla = Tabla
        frmB.vSQL = CadB
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
            PonerFoco text1(kCampo)
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
            PonerFoco text1(kCampo)
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


Private Sub PonerCampos()
Dim BrutoFac As Single
Dim b As Boolean

    On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    
    b = PonerCamposForma2(Me, Data1, 2, "Frame2")
    b = PonerCamposForma2(Me, Data1, 2, "FrameFactura")

    text1(12).Text = Format(Round2(DBLet(Data1.Recordset!BrutoFac, "N") - DBLet(Data1.Recordset!impordto, "N"), 2), FormatoImporte)

    CodTipoMov = text1(6).Text
    
'    FormatoDatosTotales
    
    'poner descripcion campos
    Modo = 4
    
    
    text2(3).Text = PonerNombreDeCod(text1(3), "clientes", "nomclien", "codclien", "N") 'cliente
    text2(4).Text = DevuelveDesdeBDNew(cAgro, "forpago", "nomforpa", "codforpa", text1(4), "N") 'forma de pago
    text2(19).Text = PonerNombreDeCod(text1(19), "variedades", "nomvarie", "codvarie", "N") ' variedades
    text2(20).Text = PonerNombreDeCod(text1(20), "tipomer", "nomtimer", "codtimer", "N") ' tipo de mercado
    
    
    Modo = 2
    
    
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
Dim I As Byte, NumReg As Byte
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
    If DatosADevolverBusqueda <> "" Or hcoCodMovim <> "" Then
        cmdRegresar.visible = b
    Else
        cmdRegresar.visible = False
    End If
        
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
          
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    'si estamos en modificar bloquea las compos que son clave primaria
    BloquearText1 Me, Modo
    
    For I = 10 To 18
        BloquearTxt text1(I), Not (Modo = 1)
        text1(I).Enabled = (Modo = 1)
    Next I
    BloquearTxt text1(5), Not (Modo = 1)
    text1(5).Enabled = (Modo = 1)


'    Me.Check1.Enabled = (Modo = 1)
    
    b = (Modo <> 1) And (Modo <> 3)
    
    b = (Modo = 1 Or Modo = 3 Or (Modo = 4 And Check1(1).Value = 0))
    
    'Campos Nº Factura bloqueado y en azul
    BloquearTxt text1(0), (Modo <> 1), True 'And (TipoFactura = 0)  'numero factura
    BloquearTxt text1(1), Not b 'fechafactura
    BloquearTxt text1(3), Not b 'cliente
    
    imgFec(0).Enabled = b
    imgFec(0).visible = b
    
    BloquearCmb Me.Combo1(0), (Modo <> 1)
    BloquearChk Me.Check1(0), (Modo <> 1)
    BloquearChk Me.Check1(1), (Modo <> 1)
    
    Me.imgZoom(0).Enabled = Not (Modo = 0)
    
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2)
    CmdCancelar.visible = b
    cmdAceptar.visible = b
    
    BloquearImgBuscar Me, Modo, ModificaLineas
                    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
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
Dim Sql As String

    On Error GoTo EDatosOK

    DatosOk = False
    
'    ComprobarDatosTotales

    'comprobamos datos OK de la tabla scafac
    b = CompForm2(Me, 2, "Frame2") ' , 1) 'Comprobar formato datos ok de la cabecera: opcion=1
    If Not b Then Exit Function
    
    If Modo = 3 Then
        'si el tipo de factura es manual y no hemos introducido valor en numero de factura
'        If TipoFactura = 1 And Text1(0).Text = "" Then
'            MsgBox "El Número de Factura no puede estar vacio. Reintroduzca.", vbExclamation
'            PonerFoco Text1(0)
'            b = False
'        End If

        'comprobamos que no exista ya la factura en la tabla facturas de ariagro
        Sql = ""
        Sql = DevuelveDesdeBDNew(cAgro, "facturas", "numfactu", "codtipom", text1(6).Text, "T", , "numfactu", text1(0).Text, "N", "fecfactu", text1(1).Text, "F")
        If Sql <> "" Then
            MsgBox "Factura ya existente. Reintroduzca.", vbExclamation
            PonerFoco text1(0)
            b = False
        End If
        If Not b Then Exit Function
        'comprobamos que no exista ya en la tabla facturas de contabilidad
        Serie = ""
'--monica:10/02/2009 stipom
'        Serie = DevuelveDesdeBDNew(cAgro, "stipom", "letraser", "codtipom", Text1(6).Text, "T")
'++monica
        Serie = ObtenerLetraSerie(text1(6).Text)
'++
        If Serie <> "" Then
            Sql = ""
            Sql = DevuelveDesdeBDNew(cConta, "cabfact", "codfaccl", "numserie", Serie, "T", , "codfaccl", text1(0).Text, "N", "fecfaccl", Mid(text1(1).Text, 7, 4), "N")
            If Sql <> "" Then
                MsgBox "Factura existente en contabilidad. Reintroduzca.", vbExclamation
                PonerFoco text1(0)
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
    Me.CmdCancelar.visible = Not b
    Me.cmdRegresar.visible = b
    Me.cmdRegresar.Caption = "Cabecera"
    If b Then
        Me.lblIndicador.Caption = "Líneas " & TituloLinea
        PonerFocoBtn Me.cmdRegresar
    End If
    'Habilitar las opciones correctas del menu segun Modo
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Function eliminar() As Boolean
Dim Sql As String, LEtra As String, Sql2 As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim Mens As String

    On Error GoTo FinEliminar

    If Data1.Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    'Eliminar en tablas de cabecera de albaran
    '------------------------------------------
    Sql = " " & ObtenerWhereCP(True)
    
    'Cabecera de factura
    conn.Execute "Delete from " & NombreTabla & Sql
    
    'Decrementar contador si borramos el ult. palet
    Set vTipoMov = New CTiposMov
    vTipoMov.DevolverContador text1(6).Text, Val(text1(0).Text)
    Set vTipoMov = Nothing
    
    b = True
FinEliminar:
    If Err.Number <> 0 Or Not b Then
        MuestraError Err.Number, "Eliminar Factura", Err.Description & " " & Mens
        b = False
    End If
    If Not b Then
        conn.RollbackTrans
        eliminar = False
    Else
        conn.CommitTrans
        eliminar = True
    End If
End Function


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
    
    Sql = "codtipom = " & DBSet(text1(6).Text, "T") & " and numfactu= " & DBSet(text1(0).Text, "N") & " and fecfactu= " & DBSet(text1(1).Text, "F")
    If conWhere Then Sql = " WHERE " & Sql
    ObtenerWhereCP = Sql
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo cadena WHERE.", Err.Description
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean, bAux As Boolean
Dim I As Integer

        b = ((Modo = 2) Or (Modo = 0)) And (hcoCodMovim = "")  'Or (Modo = 5 And ModificaLineas = 0)
        'Buscar
        Toolbar1.Buttons(1).Enabled = b
        Me.mnBuscar.Enabled = b
        'Vore Tots
        Toolbar1.Buttons(2).Enabled = b
        Me.mnVerTodos.Enabled = b
        'Añadir
        Toolbar1.Buttons(4).Enabled = b
        Me.mnModificar.Enabled = b
        
        
        b = (Modo = 2 And Data1.Recordset.RecordCount > 0) And (hcoCodMovim = "")
        'Modificar
        Toolbar1.Buttons(5).Enabled = b And Not (Check1(0).Value = 1 Or Check1(1).Value = 1)
        Me.mnModificar.Enabled = b And Not (Check1(0).Value = 1 Or Check1(1).Value = 1)
        'eliminar
        Toolbar1.Buttons(6).Enabled = b And Not (Check1(0).Value = 1 Or Check1(1).Value = 1)
        Me.mnEliminar.Enabled = b And Not (Check1(0).Value = 1 Or Check1(1).Value = 1)
        'Impresión de factura
        Toolbar1.Buttons(8).Enabled = ((Modo = 2)) Or (hcoCodMovim <> "")
        Me.mnImprimir.Enabled = ((Modo = 2)) Or (hcoCodMovim <> "")
        

End Sub


Private Sub BotonImprimir()
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String

    If text1(0).Text = "" Then
        MsgBox "Debe seleccionar una Factura para Imprimir.", vbInformation
        Exit Sub
    End If
    
    cadFormula = ""
    cadParam = ""
    cadSelect = ""
    numParam = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    indRPT = 12 'Impresion de Factura a cliente
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
      
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de factura
    '---------------------------------------------------
    If text1(0).Text <> "" Then
        'Tipo de factura
        devuelve = "{" & NombreTabla & ".codtipom}='" & text1(6).Text & "'"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        devuelve = "codtipom = '" & text1(6).Text & "'"
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
        'Nº Factura
        devuelve = "{" & NombreTabla & ".numfactu}=" & Val(text1(0).Text)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        devuelve = "numfactu = " & Val(text1(0).Text)
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
        'Fecha Factura
        devuelve = "{" & NombreTabla & ".fecfactu}=Date(" & Year(text1(1).Text) & "," & Month(text1(1).Text) & "," & Day(text1(1).Text) & ")"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        devuelve = "fecfactu = " & DBSet(text1(1).Text, "F")
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
    End If
    
    If Not HayRegParaInforme(NombreTabla, cadSelect) Then Exit Sub
     
     With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 0
            .Titulo = "Impresión de Factura"
            .ConSubInforme = True
            .Show vbModal
    End With
End Sub


Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim I As Byte
    
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
    
    
    b = ModificaDesdeFormulario(Me)
    
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
Dim Sql As String

    On Error GoTo EInsertarCab
    
    CodTipoMov = text1(6).Text
    
    Set vTipoMov = New CTiposMov
    If vTipoMov.leer(CodTipoMov) Then
        text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
        Sql = CadenaInsertarDesdeForm(Me)
        If Sql <> "" Then
            If InsertarOferta(Sql, vTipoMov) Then
                CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                PonerCadenaBusqueda
                PonerModo 2
                'Ponerse en Modo Insertar Lineas
'                BotonMtoLineas 0, "Variedades"
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
    'para ello vemos si existe una factura con ese contador y si existe la incrementamos
    Do
        devuelve = DevuelveDesdeBDNew(cAgro, NombreTabla, "numfactu", "codtipom", text1(6).Text, "T", , "numfactu", text1(0), "N", "fecfactu", text1(1), "F")
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




Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " codtipom = " & DBSet(text1(6).Text, "T") & " and numfactu= " & Val(text1(0).Text) & " and fecfactu = " & DBSet(text1(1).Text, "F")
    ' *******************************************************
    
    ObtenerWhereCab = vWhere
End Function


' **********************************************
Private Sub PonerDatosCliente(CodClien As String, Optional nifClien As String)
Dim vCliente As CCliente
Dim observaciones As String
    
    On Error GoTo EPonerDatos
    

    If CodClien = "" Then
        LimpiarDatosCliente
        Exit Sub
    End If

    Set vCliente = New CCliente
    
    'si se ha modificado el cliente volver a cargar los datos
    If vCliente.Existe(CodClien) Then
        If vCliente.LeerDatos(CodClien) Then
            text1(3).Text = vCliente.codigo
            FormateaCampo text1(3)
            If (Modo = 3) Or (Modo = 4) Then
                text2(3).Text = vCliente.Nombre  'Nom clien
                text1(4).Text = vCliente.ForPago
                text2(4).Text = PonerNombreDeCod(text1(4), "forpago", "nomforpa")
                text1(7).Text = Format(vCliente.Dto1, FormatoDescuento)
                text1(8).Text = Format(vCliente.Dto2, FormatoDescuento)
                Me.Combo1(0).ListIndex = vCliente.TipoIva
                
                TipoFactura = vCliente.TipoFactu
                text1(6).Text = "EAC"
            End If

            observaciones = DBLet(vCliente.observaciones)
            If observaciones <> "" Then
                MsgBox observaciones, vbInformation, "Observaciones del cliente"
            End If
        End If
    Else
        LimpiarDatosCliente
    End If
    Set vCliente = Nothing

EPonerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Datos Cliente", Err.Description
End Sub


Private Sub LimpiarDatosCliente()
Dim I As Byte

    text1(2).Text = ""
    text1(4).Text = ""
    text1(7).Text = ""
    text1(8).Text = ""

    text2(3).Text = ""
    text2(4).Text = ""
    Me.Combo1(0).ListIndex = -1
End Sub
    

'
'##Monica
'
Private Function CalcularDatosFacturaVenta() As Boolean
'cadWhere: cad para la where de la SQL que selecciona las lineas del albaran o la factura
'nomTabla: nombre de la tabla de albaranes(scaalp) o de AlbaranesXFactura(scafpa)
'           segun llamemos desde recepcion de facturas o desde Hco de Facturas
Dim Rs As ADODB.Recordset
Dim I As Integer

Dim Sql As String
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
    
Dim BrutoFac As Currency
Dim ImpIVA1 As Currency
Dim PorceIVA1 As Currency
Dim ImpREC1 As Currency
Dim PorceREC1 As Currency
Dim TipoIVA1 As Currency
    
Dim ImpDto1 As Currency
Dim ImpDto2 As Currency
Dim TotalFac As Currency

Dim IvaAnt As Integer
Dim cadwhere1 As String
    
Dim Nulo2 As String
Dim Nulo3 As String

Dim vClien As CCliente


Dim vImp As Currency
Dim vDto1 As Currency, vDto2 As Currency
Dim Dto1 As Currency, Dto2 As Currency
Dim Bruto As Currency


    CalcularDatosFacturaVenta = False
    On Error GoTo ECalcular

    BaseImp = 0
    BaseIVA1 = 0
    
    BrutoFac = 0
    
    ImpIVA1 = 0
    PorceIVA1 = 0
    ImpREC1 = 0
    PorceREC1 = 0
    TipoIVA1 = 0
    
    ImpDto1 = 0
    ImpDto2 = 0
    TotalFac = 0

    TotBruto = 0
    TotNeto = 0
    TotImpIVA = 0
    vBruto = 0
    vNeto = 0
    I = 1
       
    TotBruto = TotBruto + vBruto
    TotNeto = TotNeto + vNeto
    ImpBImIVA = vNeto

    Set vClien = New CCliente
    If vClien.LeerDatos(text1(3).Text) Then
    
        'Como son de tipo string comprobar que si vale "" lo ponemos a 0
        Bruto = CCur(ComprobarCero(text1(9).Text))
        Dto1 = CCur(ComprobarCero(text1(7).Text))
        Dto2 = CCur(ComprobarCero(text1(8).Text))
        
        vImp = Bruto
            
        If vClien.TipoDto = 0 Then 'Dto Aditivo
            vDto1 = (CCur(Dto1) * vImp) / 100
            vDto2 = (CCur(Dto2) * vImp) / 100
            vImp = vImp - vDto1 - vDto2
        ElseIf vClien.TipoDto = 1 Then 'Sobre Resto
            vDto1 = (CCur(Dto1) * vImp) / 100
            vImp = vImp - vDto1
            vDto2 = (CCur(Dto2) * vImp) / 100
            vImp = vImp - vDto2
        End If
        vImp = Round2(vImp, 2)
    
        text1(12).Text = vImp
        
        ' importe de descuento
        text1(10).Text = ""
        If vDto1 + vDto2 <> 0 Then text1(10).Text = Round2(vDto1 + vDto2, 2)
        
        ' base importe iva
        ImpBImIVA = CCur(text1(12).Text)
        
        'Obtener el % de IVA
        Select Case Combo1(0).ListIndex
            Case 0
                cadAux = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(vParamAplic.CodIvaNormal), "N")
                TipoIVA1 = vParamAplic.CodIvaNormal  'RS!codigiva
            Case 1
                cadAux = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(vParamAplic.CodIvaExento), "N")
                TipoIVA1 = vParamAplic.CodIvaExento  'RS!codigiva
            Case 2
                cadAux = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(vParamAplic.CodIvaRecargo), "N")
                TipoIVA1 = vParamAplic.CodIvaRecargo  'RS!codigiva
        End Select
        
        'aplicar el IVA a la base imponible de ese tipo
        impiva = CalcularPorcentaje(ImpBImIVA, CCur(ComprobarCero(cadAux)), 2)
        
        'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
        'los vamos acumulando
        TotImpIVA = TotImpIVA + impiva
    
        If CInt(Combo1(0).ListIndex) = 2 Then
            'Obtener el % de RECARGO
            cadAux1 = DevuelveDesdeBDNew(cConta, "tiposiva", "porcerec", "codigiva", CStr(TipoIVA1), "N")
    
            'aplicar el RECARGO a la base imponible de ese tipo
            ImpREC = CalcularPorcentaje(ImpBImIVA, CCur(cadAux1), 2)
            
            'sumamos todos los RECARGOS para sumarselo a la base imponible total de la factura
            'los vamos acumulando
            TotImpIVA = TotImpIVA + ImpREC
        Else
            cadAux1 = "0"
            ImpREC = 0
        End If
    
    
    
        BaseIVA1 = ImpBImIVA 'BASE IMPONIBLE
    
        PorceIVA1 = CCur(ComprobarCero(cadAux)) '% de IVA
    
        'Importe total con IVA
        ImpIVA1 = impiva
        
        PorceREC1 = CCur(ComprobarCero(cadAux1)) '% de REC
    
        'Importe total con RECARGO
        ImpREC1 = ImpREC
    
        'TOTAL de la factura
        TotalFac = BaseIVA1 + ImpIVA1 + ImpREC1
     
        text1(12).Text = Format(BaseIVA1, "###,###,##0.00") ' base imponible
        text1(13).Text = TipoIVA1
        text1(15).Text = text1(12).Text
        text1(14).Text = PorceIVA1
        text1(16).Text = ImpIVA1
        text1(17).Text = ""
        text1(18).Text = ""
        
        If PorceREC1 <> 0 Then
            text1(17).Text = PorceREC1
            text1(18).Text = ImpREC1
        End If
        text1(5).Text = Format(TotalFac, "###,###,##0.00") ' total factura
    End If
    
    Set vClien = Nothing

    CalcularDatosFacturaVenta = True

ECalcular:
    If Err.Number <> 0 Then
        CalcularDatosFacturaVenta = False
    Else
        CalcularDatosFacturaVenta = True
    End If
End Function




Private Function CalcularImporteDescuento(Importe As String, Dto1 As String, Dto2 As String, TipoDto As Byte) As String
'Calcula el Importe de una linea de Oferta, Pedido, Albaran, ...
'Importe=cantidad * precio - (descuentos)
'Si DtoProv=sprove.tipodtos, calcular Importe para Proveedores y obtener el tipo de descuento
'del campo sprove.tipodtos, si es para Clientes obtener el tipo de descuento del
'parametro spara1.tipodtos
'Tipo Descuento: 0=aditivo, 1=sobre resto
Dim vImp As Currency
Dim vDto1 As Currency, vDto2 As Currency
Dim vPre As Currency
Dim Bruto As Currency

On Error Resume Next



    'Como son de tipo string comprobar que si vale "" lo ponemos a 0
    Bruto = ComprobarCero(Importe)
    Dto1 = ComprobarCero(Dto1)
    Dto2 = ComprobarCero(Dto2)
    
    If Bruto <> "" Then
        vImp = CCur(Bruto)
    End If
        
    If TipoDto = 0 Then 'Dto Aditivo
        vDto1 = (CCur(Dto1) * vImp) / 100
        vDto2 = (CCur(Dto2) * vImp) / 100
        vImp = vImp - vDto1 - vDto2
    ElseIf TipoDto = 1 Then 'Sobre Resto
        vDto1 = (CCur(Dto1) * vImp) / 100
        vImp = vImp - vDto1
        vDto2 = (CCur(Dto2) * vImp) / 100
        vImp = vImp - vDto2
    End If
    vImp = Round(vImp, 2)
    CalcularImporteDescuento = CStr(vImp)
End Function



Private Function FacturaDescontada(TipoM As String, NumFact As Long, FecFact As Date) As Boolean
Dim Sql As String
    
    FacturaDescontada = False
    
    Sql = "select count(*) from facturas_acuenta where codtipomcta = " & DBSet(TipoM, "T")
    Sql = Sql & " and numfactucta = " & DBSet(NumFact, "N")
    Sql = Sql & " and fecfactucta = " & DBSet(FecFact, "F")
    
    FacturaDescontada = (TotalRegistros(Sql) <> 0)

End Function
