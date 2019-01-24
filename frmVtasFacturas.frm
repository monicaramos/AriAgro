VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmVtasFacturas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturas Clientes"
   ClientHeight    =   11055
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   17850
   Icon            =   "frmVtasFacturas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11055
   ScaleWidth      =   17850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   90
      TabIndex        =   147
      Top             =   45
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   148
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
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   3720
      TabIndex        =   145
      Top             =   45
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   146
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
      Left            =   15075
      TabIndex        =   144
      Top             =   315
      Width           =   1605
   End
   Begin VB.CheckBox chkBio 
      Caption         =   "BIO"
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
      Left            =   6930
      TabIndex        =   143
      Top             =   315
      Width           =   1215
   End
   Begin VB.CheckBox chkRectifica 
      Caption         =   "Rectificativa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8235
      TabIndex        =   142
      Top             =   315
      Width           =   1665
   End
   Begin VB.Frame Frame2 
      Height          =   4185
      Left            =   90
      TabIndex        =   37
      Top             =   810
      Width           =   17605
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
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
         Left            =   2295
         MaxLength       =   7
         TabIndex        =   6
         Tag             =   "Cambio Divisa|N|N|||facturas|cambiodivisa|###0.0000||"
         Text            =   "Text1 7"
         Top             =   1890
         Width           =   980
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   225
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "Tipo Moneda|N|N|||facturas|codmoneda||N|"
         Top             =   1890
         Width           =   1980
      End
      Begin VB.CheckBox Check1 
         Caption         =   "FacturaE"
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
         Left            =   6300
         TabIndex        =   14
         Tag             =   "FacturaE|N|N|0|1|facturas|enfacturae|||"
         Top             =   2655
         Width           =   1725
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Pasa EDI"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   6300
         TabIndex        =   13
         Tag             =   "Pasa Edicom|N|N|||facturas|pasedicom|0||"
         Top             =   2325
         Width           =   1635
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
         Left            =   6705
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Fecha Factura|F|N|||facturas|fecfactu|dd/mm/yyyy|S|"
         Top             =   720
         Width           =   1350
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   225
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Tag             =   "Tipo Iva|N|N|||facturas|tipoivac||N|"
         Top             =   2520
         Width           =   1980
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
         Left            =   1350
         MaxLength       =   12
         TabIndex        =   1
         Tag             =   "Tipo Movimiento|T|N|||facturas|codtipom||S|"
         Top             =   720
         Width           =   855
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
         Height          =   240
         Index           =   1
         Left            =   6300
         TabIndex        =   11
         Tag             =   "Contabilizado|N|N|||facturas|intconta|0||"
         Top             =   1665
         Width           =   1725
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
         Height          =   240
         Index           =   0
         Left            =   6300
         TabIndex        =   12
         Tag             =   "Pasa Aridoc|N|N|||facturas|pasaridoc|0||"
         Top             =   1995
         Width           =   1635
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
         Left            =   1350
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "Forma Pago|N|N|0|999|facturas|codforpa|000||"
         Text            =   "Text1"
         Top             =   1170
         Width           =   860
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
         Left            =   2295
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   46
         Text            =   "Text2"
         Top             =   1170
         Width           =   5745
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
         Left            =   3375
         MaxLength       =   5
         TabIndex        =   9
         Tag             =   "%Dto 2|N|S|0|100|facturas|dtocom2|##0.00||"
         Top             =   2520
         Width           =   945
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
         Left            =   2295
         MaxLength       =   5
         TabIndex        =   8
         Tag             =   "%Dto 1|N|S|0|100|facturas|dtocom1|##0.00||"
         Top             =   2520
         Width           =   945
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
         Left            =   4455
         MaxLength       =   10
         TabIndex        =   10
         Tag             =   "Importe Dto|N|S|||facturas|impdtoc|###,##0.00||"
         Text            =   "Text3"
         Top             =   2520
         Width           =   1500
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
         Left            =   2295
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   41
         Text            =   "Text2"
         Top             =   270
         Width           =   5790
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
         Height          =   690
         Index           =   2
         Left            =   225
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Tag             =   "Observaciones|T|S|||facturas|observac|||"
         Top             =   3270
         Width           =   7815
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
         Left            =   1350
         MaxLength       =   6
         TabIndex        =   0
         Tag             =   "Cod. Cliente|N|N|0|999999|facturas|codclien|000000||"
         Text            =   "Text1"
         Top             =   270
         Width           =   860
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
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
         Left            =   3465
         MaxLength       =   7
         TabIndex        =   2
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Text            =   "Text1 7"
         Top             =   720
         Width           =   980
      End
      Begin VB.Frame FrameFactura 
         Height          =   3780
         Left            =   8460
         TabIndex        =   51
         Top             =   180
         Width           =   8450
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
            Left            =   5475
            MaxLength       =   5
            TabIndex        =   117
            Tag             =   "% REC 1|N|S|0|99.90|facturas|porcrec1|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   660
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
            Index           =   30
            Left            =   6300
            MaxLength       =   15
            TabIndex        =   116
            Tag             =   "Importe REC 1|N|S|0||facturas|imporec1|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   1650
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
            Index           =   29
            Left            =   5475
            MaxLength       =   5
            TabIndex        =   115
            Tag             =   "% REC 2|N|S|0|99.90|facturas|porcrec2|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   2010
            Width           =   660
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
            Left            =   6300
            MaxLength       =   15
            TabIndex        =   114
            Tag             =   "Importe REC 2|N|S|0||facturas|imporec2|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   2010
            Width           =   1650
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
            Index           =   27
            Left            =   5475
            MaxLength       =   5
            TabIndex        =   113
            Tag             =   "% REC 3|N|S|0|99.90|facturas|porcrec3|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   2475
            Width           =   660
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
            Index           =   26
            Left            =   6300
            MaxLength       =   15
            TabIndex        =   112
            Tag             =   "Importe REC 3|N|S|0||facturas|imporec3|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   2475
            Width           =   1650
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
            Height          =   350
            Index           =   25
            Left            =   6300
            MaxLength       =   15
            TabIndex        =   31
            Tag             =   "Total Factura|N|S|0||facturas|totalfac|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   3225
            Width           =   1650
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
            Index           =   24
            Left            =   3585
            MaxLength       =   15
            TabIndex        =   30
            Tag             =   "Importe IVA 3|N|S|0||facturas|impoiva3|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   2475
            Width           =   1710
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
            Left            =   2790
            MaxLength       =   5
            TabIndex        =   28
            Tag             =   "% IVA 3|N|S|0|99.90|facturas|porciva3|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   2475
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
            Left            =   855
            MaxLength       =   15
            TabIndex        =   29
            Tag             =   "Base Imponible 3|N|S|0||facturas|baseimp3|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   2475
            Width           =   1710
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
            Index           =   20
            Left            =   3585
            MaxLength       =   15
            TabIndex        =   26
            Tag             =   "Importe IVA 2|N|S|0||facturas|impoiva2|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   2010
            Width           =   1710
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
            Left            =   2775
            MaxLength       =   5
            TabIndex        =   24
            Tag             =   "& IVA 2|N|S|0|99.90|facturas|porciva2|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   2010
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
            Index           =   19
            Left            =   855
            MaxLength       =   15
            TabIndex        =   25
            Tag             =   "Base Imponible 2 |N|S|0||facturas|baseimp2|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   2025
            Width           =   1710
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
            Index           =   16
            Left            =   3585
            MaxLength       =   15
            TabIndex        =   22
            Tag             =   "Importe IVA 1|N|S|0||facturas|impoiva1|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   1710
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
            Left            =   2775
            MaxLength       =   5
            TabIndex        =   20
            Tag             =   "% IVA 1|N|S|0|99.90|facturas|porciva1|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   1560
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
            Index           =   15
            Left            =   855
            MaxLength       =   15
            TabIndex        =   21
            Tag             =   "Base Imponible 1|N|S|0||facturas|baseimp1|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   1710
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
            Left            =   300
            MaxLength       =   5
            TabIndex        =   19
            Tag             =   "IVA 1|N|S|0|9|facturas|codiiva1|0|N|"
            Text            =   "Text1 7"
            Top             =   1560
            Width           =   500
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
            Left            =   300
            MaxLength       =   5
            TabIndex        =   23
            Tag             =   "IVA 2|N|S|0|9|facturas|codiiva2|0|N|"
            Text            =   "Text1 7"
            Top             =   2010
            Width           =   500
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
            Left            =   300
            MaxLength       =   5
            TabIndex        =   27
            Tag             =   "IVA 3|N|S|0|9|facturas|codiiva3|0|N|"
            Text            =   "Text1 7"
            Top             =   2475
            Width           =   500
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
            Left            =   855
            MaxLength       =   15
            TabIndex        =   16
            Tag             =   "Bruto Factura|N|S|0||facturas|brutofac|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   540
            Width           =   1710
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
            Left            =   3600
            MaxLength       =   15
            TabIndex        =   17
            Tag             =   "Impòrte Descuento|N|S|0||facturas|impordto|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   540
            Width           =   1665
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
            Index           =   12
            Left            =   6300
            MaxLength       =   15
            TabIndex        =   18
            Text            =   "Text1 7"
            Top             =   540
            Width           =   1650
         End
         Begin VB.Line Line1 
            X1              =   6300
            X2              =   7920
            Y1              =   990
            Y2              =   990
         End
         Begin VB.Label Label1 
            Caption         =   "Importe Recargo"
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
            Left            =   6300
            TabIndex        =   119
            Top             =   1260
            Width           =   1650
         End
         Begin VB.Label Label1 
            Caption         =   "% Rec"
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
            Left            =   5475
            TabIndex        =   118
            Top             =   1260
            Width           =   660
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
            Left            =   2775
            TabIndex        =   62
            Top             =   1260
            Width           =   810
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
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   39
            Left            =   4005
            TabIndex        =   61
            Top             =   3285
            Width           =   2025
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
            TabIndex        =   60
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
            Height          =   210
            Index           =   37
            Left            =   7020
            TabIndex        =   59
            Top             =   990
            Width           =   210
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
            Left            =   3585
            TabIndex        =   58
            Top             =   1260
            Width           =   1425
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
            Index           =   13
            Left            =   870
            TabIndex        =   57
            Top             =   1230
            Width           =   1530
         End
         Begin VB.Label Label1 
            Caption         =   "Cod."
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
            Left            =   300
            TabIndex        =   56
            Top             =   1230
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Bruto Factura"
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
            Left            =   855
            TabIndex        =   55
            Top             =   270
            Width           =   1110
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
            Index           =   7
            Left            =   6300
            TabIndex        =   54
            Top             =   270
            Width           =   1530
         End
         Begin VB.Label Label1 
            Caption         =   "Imp.Descuento"
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
            Left            =   3600
            TabIndex        =   53
            Top             =   270
            Width           =   1710
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
            Left            =   3150
            TabIndex        =   52
            Top             =   540
            Width           =   135
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Cambio Divisa"
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
         Left            =   2295
         TabIndex        =   151
         Top             =   1620
         Width           =   1620
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Moneda"
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
         Left            =   225
         TabIndex        =   150
         Top             =   1620
         Width           =   1575
      End
      Begin VB.Label Label1 
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
         Index           =   3
         Left            =   225
         TabIndex        =   50
         Top             =   765
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Iva"
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
         Left            =   225
         TabIndex        =   48
         Top             =   2250
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
         Left            =   225
         TabIndex        =   47
         Top             =   1215
         Width           =   720
      End
      Begin VB.Label Label1 
         Caption         =   "%Dto 2"
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
         Left            =   3375
         TabIndex        =   45
         Top             =   2295
         Width           =   1050
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Factura"
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
         Left            =   4905
         TabIndex        =   44
         Top             =   780
         Width           =   1500
      End
      Begin VB.Label Label1 
         Caption         =   "%Dto 1"
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
         Left            =   2295
         TabIndex        =   43
         Top             =   2295
         Width           =   1140
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   6435
         Picture         =   "frmVtasFacturas.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   765
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Dtos"
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
         Left            =   4500
         TabIndex        =   42
         Top             =   2295
         Width           =   1290
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   1710
         ToolTipText     =   "Zoom descripción"
         Top             =   2970
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
         Left            =   225
         TabIndex        =   40
         Top             =   2970
         Width           =   1485
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
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
         Left            =   225
         TabIndex        =   39
         Top             =   315
         Width           =   765
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1035
         ToolTipText     =   "Buscar Cliente"
         Top             =   315
         Width           =   240
      End
      Begin VB.Label Label1 
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
         Index           =   28
         Left            =   2340
         TabIndex        =   38
         Top             =   780
         Width           =   1125
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5100
      Left            =   90
      TabIndex        =   49
      Top             =   5175
      Width           =   17610
      _ExtentX        =   31062
      _ExtentY        =   8996
      _Version        =   393216
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
      TabCaption(0)   =   "Variedades"
      TabPicture(0)   =   "frmVtasFacturas.frx":0097
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameAux0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Envases"
      TabPicture(1)   =   "frmVtasFacturas.frx":00B3
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameAux1"
      Tab(1).Control(1)=   "txtAux(13)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Facturas a Cuenta"
      TabPicture(2)   =   "frmVtasFacturas.frx":00CF
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameAux2"
      Tab(2).ControlCount=   1
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   13
         Left            =   -63450
         MaxLength       =   10
         TabIndex        =   141
         Tag             =   "Fecha Albaran|F|S|||facturas_envase|fecalbar|dd/mm/yyyy||"
         Text            =   "Fec.Alb"
         Top             =   2580
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Frame FrameAux2 
         BorderStyle     =   0  'None
         Height          =   4155
         Left            =   -74955
         TabIndex        =   122
         Top             =   345
         Width           =   16625
         Begin VB.TextBox txtAux1 
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
            Left            =   10260
            MaxLength       =   12
            TabIndex        =   137
            Tag             =   "Importe REC 1|N|S|0||facturas|imporec1|#,###,###,##0.00|N|"
            Text            =   "Impor.Rec"
            Top             =   2250
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.TextBox txtAux1 
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
            Left            =   9570
            MaxLength       =   12
            TabIndex        =   136
            Tag             =   "% REC 1|N|S|0|99.90|facturas|porcrec1|#0.00|N|"
            Text            =   "%rec"
            Top             =   2250
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.TextBox txtAux1 
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
            Left            =   8550
            MaxLength       =   12
            TabIndex        =   135
            Tag             =   "Importe IVA 1|N|S|0||facturas|impoiva1|#,###,###,##0.00|N|"
            Text            =   "importe iva"
            Top             =   2250
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.TextBox txtAux1 
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
            Left            =   7920
            MaxLength       =   12
            TabIndex        =   134
            Tag             =   "% IVA 1|N|S|0|99.90|facturas|porciva1|#0.00|N|"
            Text            =   "%iva"
            Top             =   2250
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.TextBox txtAux1 
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
            Left            =   6630
            MaxLength       =   12
            TabIndex        =   133
            Tag             =   "Base Imponible|N|N|||facturas_acuenta|totalfaccta|###,##0.00||"
            Text            =   "Baseimp"
            Top             =   2250
            Visible         =   0   'False
            Width           =   1215
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
            Index           =   2
            Left            =   4680
            MaskColor       =   &H00000000&
            TabIndex        =   130
            ToolTipText     =   "Buscar Fecha"
            Top             =   2250
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtAux1 
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
            Left            =   270
            MaxLength       =   12
            TabIndex        =   129
            Tag             =   "Tipo Movim.|T|N|||facturas_acuenta|codtipom||S|"
            Text            =   "TipoMov"
            Top             =   2250
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.TextBox txtAux1 
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
            Left            =   990
            MaxLength       =   7
            TabIndex        =   128
            Tag             =   "Num.Factura|N|N|||facturas_acuenta|numfactu|0000000|S|"
            Text            =   "NumFact"
            Top             =   2250
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.TextBox txtAux1 
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
            Left            =   3645
            MaxLength       =   12
            TabIndex        =   127
            Tag             =   "Fec.Factu|F|N|||facturas_acuenta|fecfactucta|dd/mm/yyyy|S|"
            Text            =   "Fecfac"
            Top             =   2250
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtAux1 
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
            Left            =   4980
            MaxLength       =   12
            TabIndex        =   126
            Tag             =   "Total Factura|N|N|||facturas_acuenta|totalfaccta|###,##0.00||"
            Text            =   "totalfac"
            Top             =   2250
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox txtAux1 
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
            Left            =   1710
            MaxLength       =   12
            TabIndex        =   125
            Tag             =   "Fec.Factu|F|N|||facturas_acuenta|fecfactu|dd/mm/yyyy|S|"
            Text            =   "FecFact"
            Top             =   2250
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.TextBox txtAux1 
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
            Left            =   2430
            MaxLength       =   12
            TabIndex        =   124
            Tag             =   "Cod.Tipo Movimiento|T|N|||facturas_acuenta|codtipomcta||S|"
            Text            =   "Tipom"
            Top             =   2250
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.TextBox txtAux1 
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
            Left            =   3105
            MaxLength       =   7
            TabIndex        =   123
            Tag             =   "Num.Factura|N|N|||facturas_acuenta|numfactucta|0000000|S|"
            Text            =   "Fac"
            Top             =   2250
            Visible         =   0   'False
            Width           =   420
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   2
            Left            =   90
            TabIndex        =   131
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
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "frmVtasFacturas.frx":00EB
            Height          =   3480
            Left            =   90
            TabIndex        =   132
            Top             =   630
            Width           =   16490
            _ExtentX        =   29078
            _ExtentY        =   6138
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
            Index           =   2
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
      End
      Begin VB.Frame FrameAux1 
         BorderStyle     =   0  'None
         Height          =   4290
         Left            =   -74955
         TabIndex        =   79
         Top             =   345
         Width           =   16825
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
            Left            =   10860
            MaxLength       =   7
            TabIndex        =   140
            Tag             =   "Lin.Albaran|N|S|||facturas_envase|numlinealbar|00||"
            Text            =   "linealb"
            Top             =   2250
            Visible         =   0   'False
            Width           =   555
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
            Left            =   10260
            MaxLength       =   7
            TabIndex        =   139
            Tag             =   "Num.Albaran|N|S|||facturas_envase|numalbar|0000000||"
            Text            =   "albaran"
            Top             =   2250
            Visible         =   0   'False
            Width           =   555
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
            Left            =   3105
            MaxLength       =   3
            TabIndex        =   91
            Tag             =   "Almacen|N|N|||facturas_envase|codalmac|000||"
            Text            =   "Alm"
            Top             =   2250
            Visible         =   0   'False
            Width           =   420
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
            Left            =   1575
            Locked          =   -1  'True
            MaxLength       =   500
            TabIndex        =   96
            Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
            Top             =   3975
            Width           =   9915
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
            Index           =   10
            Left            =   9630
            MaxLength       =   2
            TabIndex        =   89
            Tag             =   "CodIva|N|N|||facturas_envase|codigiva|00||"
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
            Left            =   8820
            MaxLength       =   12
            TabIndex        =   97
            Tag             =   "Importe|N|N|||facturas_envase|importel|##,###,##0.00||"
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
            Left            =   8010
            MaxLength       =   5
            TabIndex        =   95
            Tag             =   "Dto.Linea|N|N|||facturas_envase|dtolinea|#0.00||"
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
            Left            =   7155
            MaxLength       =   12
            TabIndex        =   94
            Tag             =   "Precio|N|N|||facturas_envase|precioar|###,##0.0000||"
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
            Left            =   2430
            MaxLength       =   12
            TabIndex        =   88
            Tag             =   "Num.Linea|N|N|||facturas_variedad|numlinea|000|S|"
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
            Left            =   1710
            MaxLength       =   12
            TabIndex        =   87
            Tag             =   "Fec.Factu|F|N|||facturas_variedad|fecfactu||S|"
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
            Left            =   6255
            MaxLength       =   12
            TabIndex        =   93
            Tag             =   "Cantidad|N|N|||facturas_envase|cantidad|###,##0.00||"
            Text            =   "cantidad"
            Top             =   2250
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
            Index           =   5
            Left            =   3645
            MaxLength       =   16
            TabIndex        =   92
            Tag             =   "Artículo|T|N|||facturas_envase|codartic||N|"
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
            Left            =   1020
            MaxLength       =   12
            TabIndex        =   83
            Tag             =   "Num.Factura|N|N|||facturas_variedad|numfactu|0000000|S|"
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
            Left            =   270
            MaxLength       =   12
            TabIndex        =   82
            Tag             =   "Tipo Movim.|T|N|||facturas_variedad|codtipom||S|"
            Text            =   "TipoMov"
            Top             =   2250
            Visible         =   0   'False
            Width           =   675
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
            Height          =   330
            Index           =   0
            Left            =   4905
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   81
            Text            =   "Nombre articulo"
            Top             =   2250
            Width           =   1200
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
            Left            =   4680
            MaskColor       =   &H00000000&
            TabIndex        =   80
            ToolTipText     =   "Buscar Envase"
            Top             =   2250
            Visible         =   0   'False
            Width           =   195
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   1
            Left            =   90
            TabIndex        =   84
            Top             =   135
            Width           =   1500
            _ExtentX        =   2646
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
                  Object.ToolTipText     =   "Insertar Envases"
                  Object.Tag             =   "2"
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid DataGrid3 
            Bindings        =   "frmVtasFacturas.frx":0100
            Height          =   3105
            Left            =   90
            TabIndex        =   85
            Top             =   630
            Width           =   16595
            _ExtentX        =   29263
            _ExtentY        =   5477
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
            Left            =   225
            TabIndex        =   90
            Top             =   4020
            Width           =   1335
         End
      End
      Begin VB.Frame FrameAux0 
         BorderStyle     =   0  'None
         Height          =   4065
         Left            =   45
         TabIndex        =   63
         Top             =   420
         Width           =   17430
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
            Left            =   5670
            MaxLength       =   11
            TabIndex        =   72
            Tag             =   "Unidades|N|S|||facturas_variedad|unidades|#,##0||"
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
            Left            =   7380
            MaxLength       =   11
            TabIndex        =   111
            Tag             =   "Iva|N|N|||facturas_variedad|codigiva|00|N|"
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
            Left            =   6660
            MaxLength       =   11
            TabIndex        =   110
            Tag             =   "Dto2|N|S|||facturas_variedad|dtocom2|##0.00|N|"
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
            Left            =   5940
            MaxLength       =   11
            TabIndex        =   109
            Tag             =   "Dto1|N|S|||facturas_variedad|dtocom1|##0.00|N|"
            Text            =   "Dto1"
            Top             =   2250
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Frame Frame3 
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
            Height          =   3615
            Left            =   11295
            TabIndex        =   98
            Top             =   360
            Width           =   5595
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
               Left            =   450
               MaxLength       =   30
               TabIndex        =   108
               Text            =   "Text1 7"
               Top             =   2790
               Width           =   4680
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
               Index           =   3
               Left            =   450
               MaxLength       =   15
               TabIndex        =   106
               Text            =   "Text1 7"
               Top             =   2070
               Width           =   4680
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
               Index           =   2
               Left            =   450
               MaxLength       =   30
               TabIndex        =   104
               Text            =   "Text1 7"
               Top             =   1350
               Width           =   4680
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
               Index           =   1
               Left            =   3060
               MaxLength       =   12
               TabIndex        =   102
               Text            =   "Text1 7"
               Top             =   585
               Width           =   2070
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
               Index           =   0
               Left            =   450
               MaxLength       =   10
               TabIndex        =   101
               Text            =   "Text1 7"
               Top             =   630
               Width           =   1400
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
               Height          =   195
               Left            =   450
               TabIndex        =   107
               Top             =   2520
               Width           =   1365
            End
            Begin VB.Label Label5 
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
               Height          =   195
               Left            =   450
               TabIndex        =   105
               Top             =   1800
               Width           =   915
            End
            Begin VB.Label Label4 
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
               Height          =   195
               Left            =   450
               TabIndex        =   103
               Top             =   1080
               Width           =   915
            End
            Begin VB.Label Label3 
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
               Height          =   195
               Left            =   3060
               TabIndex        =   100
               Top             =   315
               Width           =   1455
            End
            Begin VB.Label Label2 
               Caption         =   "Fecha Albarán"
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
               Left            =   450
               TabIndex        =   99
               Top             =   360
               Width           =   1455
            End
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
            Index           =   1
            Left            =   4005
            MaskColor       =   &H00000000&
            TabIndex        =   86
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
            Left            =   3420
            MaxLength       =   3
            TabIndex        =   69
            Tag             =   "Lin.Albaran|N|N|||facturas_variedad|numlinealbar|000||"
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
            Left            =   8415
            MaxLength       =   11
            TabIndex        =   76
            Tag             =   "Imp.Neto|N|N|||facturas_variedad|impornet|##,###,##0.00|N|"
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
            Left            =   7785
            MaxLength       =   11
            TabIndex        =   75
            Tag             =   "Imp.Bruto|N|N|||facturas_variedad|imporbru|##,###,##0.00|N|"
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
            Left            =   4185
            MaxLength       =   8
            TabIndex        =   70
            Tag             =   "Cant.Real|N|N|||facturas_variedad|cantreal|###,##0||"
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
            Left            =   7200
            MaxLength       =   11
            TabIndex        =   74
            Tag             =   "Precio Neto|N|N|||facturas_variedad|precinet|###,##0.0000||"
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
            Left            =   6435
            MaxLength       =   11
            TabIndex        =   73
            Tag             =   "Precio Bruto|N|N|||facturas_variedad|precibru|###,##0.0000||"
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
            Left            =   4950
            MaxLength       =   8
            TabIndex        =   71
            Tag             =   "Cant.Fact|N|N|||facturas_variedad|cantfact|###,##0||"
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
            Left            =   2745
            MaxLength       =   7
            TabIndex        =   68
            Tag             =   "Num.Albaran|N|N|||facturas_variedad|numalbar|000000||"
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
            Left            =   2205
            MaxLength       =   30
            TabIndex        =   67
            Tag             =   "Num.Linea|N|N|||facturas_variedad|numlinea|000|S|"
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
            Left            =   225
            MaxLength       =   7
            TabIndex        =   66
            Tag             =   "Tipo Movim.|T|N|||facturas_variedad|codtipom||S|"
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
            Left            =   855
            MaxLength       =   15
            TabIndex        =   65
            Tag             =   "Num.Factura|N|N|||facturas_variedad|numfactu|0000000|S|"
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
            Left            =   1440
            MaxLength       =   30
            TabIndex        =   64
            Tag             =   "Fec.Factu|F|N|||facturas_variedad|fecfactu||S|"
            Text            =   "Fec.Factu"
            Top             =   1800
            Visible         =   0   'False
            Width           =   765
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   0
            Left            =   135
            TabIndex        =   77
            Top             =   45
            Width           =   1140
            _ExtentX        =   2011
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
            Bindings        =   "frmVtasFacturas.frx":0115
            Height          =   3480
            Left            =   135
            TabIndex        =   78
            Top             =   495
            Width           =   10965
            _ExtentX        =   19341
            _ExtentY        =   6138
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
            Left            =   1770
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
      TabIndex        =   35
      Top             =   10395
      Width           =   2490
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
         Left            =   135
         TabIndex        =   36
         Top             =   180
         Width           =   2025
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
      Left            =   16515
      TabIndex        =   33
      Top             =   10485
      Width           =   1065
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
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
      Left            =   15345
      TabIndex        =   32
      Top             =   10485
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
      Left            =   16515
      TabIndex        =   34
      Top             =   10485
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
      TabIndex        =   120
      Text            =   "Text1 7"
      Top             =   1710
      Width           =   1485
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   17115
      TabIndex        =   149
      Top             =   255
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
   Begin VB.Label Label7 
      Caption         =   "Pulse F2 para introducir valores por calibre"
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
      Height          =   315
      Left            =   2790
      TabIndex        =   138
      Top             =   10515
      Visible         =   0   'False
      Width           =   4965
   End
   Begin VB.Label Label1 
      Caption         =   "Imp.Descuento 2"
      Height          =   255
      Index           =   10
      Left            =   6705
      TabIndex        =   121
      Top             =   1440
      Width           =   1215
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
Attribute VB_Name = "frmVtasFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public Facturas As String  ' venimos de albaranes para ver las facturas donde aparece el albaran

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
Private WithEvents frmAlb As frmAlbVtas 'Form Mto de Albaranes
Attribute frmAlb.VB_VarHelpID = -1

Private WithEvents frmCli As frmClientes 'Form Mto de Clientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmFPag As frmManFpago 'Form Mto de Formas de Pago
Attribute frmFPag.VB_VarHelpID = -1
Private WithEvents frmAlm As frmManAlmProp 'Form Mto de almacenes propios
Attribute frmAlm.VB_VarHelpID = -1

Private WithEvents frmFac As frmBasico2 'manda busqueda previa facturas
Attribute frmFac.VB_VarHelpID = -1


Private WithEvents frmMens As frmMensajes ' devolvemos las facturas a cuenta que vamos a descontar
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

Dim FechaAnt As String
'[Monica]28/08/2013
Dim ClienteAnt As String
Dim FPagoAnt As String
Dim Dto1Ant As String
Dim Dto2Ant As String
Dim IDtoAnt As String
Dim ObsAnt As String
Dim EdiAnt As String
Dim ContaAnt As String

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
Dim PulsadoF2 As Boolean
Private BuscaChekc As String

Dim ArticAnt As String


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
            Set frmAlb = New frmAlbVtas
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
        

        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaCabecera Then
                    espera 0.2
                    TerminaBloquear
                    PosicionarData
                    PonerCampos
                    PonerCamposLineas
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
                        If Not AdoAux(0).Recordset.EOF Then AdoAux(0).Recordset.MoveFirst
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
                        If Not AdoAux(1).Recordset.EOF Then AdoAux(1).Recordset.MoveFirst
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
                        DataGrid1.AllowAddNew = False
                        If Not AdoAux(2).Recordset.EOF Then AdoAux(2).Recordset.MoveFirst
                    End If
                    ModificaLineas = 0
                    LLamaLineas Modo, 0, "DataGrid3"
                    PonerModo 2
                    DataGrid1.Enabled = True
                    If Not Data1.Recordset.EOF Then _
                        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            
                    'Habilitar las opciones correctas del menu segun Modo
                    PonerModoOpcionesMenu (Modo)
                    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
                    DataGrid1.Enabled = True
                    PonerFocoGrid DataGrid1
                
                    
             End Select
            
'            PonerBotonCabecera True
    
            
            
            
'            Me.DataGrid1.Enabled = True
    End Select
End Sub

Private Sub BotonAnyadir()

    LimpiarCampos 'Huida els TextBox
    
'    TipoFactura = 1
    
    PonerModo 3
    
    If Me.chkRectifica Then
        Text1(6).Enabled = False
        If vParamAplic.Cooperativa = 15 Then
            Text1(6).Text = "FR3"
        Else
            Text1(6).Text = "FAR"
        End If
        Text1(0).Enabled = False
    Else
        If Me.chkBio Then
            Text1(6).Text = "FBI"
        End If
    End If
    
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la capçalera *******
'    Text1(0).Text = SugerirCodigoSiguienteStr("palets", "numpalet")
'    FormateaCampo Text1(0)
'    cmbAux(0).ListIndex = -1
    
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
        
    '[Monica]17/12/2018: tipo de moneda y cambio de divisa por defecdto euro y 1
    Combo1(1).ListIndex = 0
    Text1(32).Text = Format(1, "###0.0000")
        
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
        Text1(3).BackColor = vbLightBlue 'vbYellow
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
        MandaBusquedaPrevia "facturas.codtipom <> 'EAC'"
    Else
        LimpiarCampos
        LimpiarDataGrids
        CadenaConsulta = "Select facturas.* "
        CadenaConsulta = CadenaConsulta & "from " & NombreTabla & " where codtipom <> 'EAC' " & Ordenacion
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

'    '[Monica]23/03/2012: sólo se puede modificar
'    'solo se puede modificar la factura si no esta contabilizada
'    If FactContabilizada(True) Then
'        TerminaBloquear
'        Exit Sub
'    End If
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    '[Monica]20/09/2012: antes ponia el foco en el 4
    FechaAnt = Text1(1).Text
    
    '[Monica]28/08/2013: dejamos modificar la factura aunque este contabilizada
    ClienteAnt = Text1(3).Text
    FPagoAnt = Text1(4).Text
    Dto1Ant = Text1(7).Text
    Dto2Ant = Text1(8).Text
    IDtoAnt = Text1(5).Text
    ObsAnt = Text1(2).Text
    EdiAnt = Check1(2).Value
    ContaAnt = Check1(1).Value
    
    
    '[Monica]20/09/2012: antes ponia el foco en el 4
    PonerFoco Text1(3) '*** 1r camp visible que siga PK ***
        
End Sub


Private Sub BotonModificarLinea(Index As Integer)
'Modificar una linea
Dim vWhere As String
Dim anc As Single
Dim J As Byte

    On Error GoTo eModificarLinea

    PulsadoF2 = False
'     'solo se puede modificar la factura si no esta contabilizada
'    If FactContabilizada(True) Then
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
            vWhere = Replace(ObtenerWhereCP(False), "facturas", "facturas_variedad")
            vWhere = vWhere & " and numlinea=" & AdoAux(0).Recordset!NumLinea
            If Not BloqueaRegistro("facturas_variedad", vWhere) Then
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
            
            '[Monica]15/07/2011: dejamos modificar cantidades
            BloquearTxt txtAux3(6), False
            BloquearTxt txtAux3(7), False
            
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
            PonerFoco txtAux3(6)
        
            BloquearBtn Me.btnBuscar(1), True
        Case 1 ' envases
            vWhere = Replace(ObtenerWhereCP(False), "facturas", "facturas_envases")
            vWhere = vWhere & " and numlinea=" & AdoAux(1).Recordset!NumLinea
            If Not BloqueaRegistro("facturas_envases", vWhere) Then
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
            
            '[Monica]20/03/2012: numero de albaran y linea de albaran
            txtAux(11).Text = DataGrid3.Columns(13).Text
            txtAux(12).Text = DataGrid3.Columns(14).Text
            txtAux(13).Text = DataGrid3.Columns(15).Text

            BloquearTxt txtAux(4), True
'[Monica]13/05/2015: dejamos modificar el articulo
'            BloquearTxt txtAux(5), True
'[Monica]28/07/2014: dejamos modificar el precio
'            BloquearTxt txtAux(7), True
            BloquearTxt txtAux(9), True
            txtAux(4).Enabled = False
'[Monica]13/05/2015: dejamos modificar el articulo
'            txtAux(5).Enabled = False
'[Monica]28/07/2014: dejamos modificar el precio
'            txtAux(7).Enabled = False
            txtAux(9).Enabled = False
            
            BloquearTxt txtAux(6), False
            BloquearTxt txtAux(8), False
            
'[Monica]13/05/2015: dejamos modificar el articulo
           BloquearBtn Me.btnBuscar(0), False
            
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

            ArticAnt = txtAux(5).Text
       
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
        
            Label7.visible = b
        
        Case "DataGrid3"
            DeseleccionaGrid Me.DataGrid3
            b = (xModo = 1 Or xModo = 2)
            For jj = 4 To 9
                txtAux(jj).Height = DataGrid3.RowHeight - 10
                txtAux(jj).Top = alto + 5
                txtAux(jj).visible = b
                txtAux(jj).Enabled = b
            Next jj
            Text2(0).Height = DataGrid3.RowHeight - 10
            Text2(0).Top = alto + 5
            Text2(0).visible = b
           
            btnBuscar(0).Height = DataGrid3.RowHeight - 10
            btnBuscar(0).Top = alto + 5
            btnBuscar(0).visible = b
            
            txtAux(11).Height = DataGrid3.RowHeight - 10
            txtAux(11).Top = alto + 5
            txtAux(11).visible = b
            
        Case "DataGrid1"
            DeseleccionaGrid Me.DataGrid1
            b = (xModo = 1 Or xModo = 2)
            For jj = 3 To 6
                txtAux1(jj).Height = DataGrid1.RowHeight - 10
                txtAux1(jj).Top = alto + 5
                txtAux1(jj).visible = b
            Next jj
            btnBuscar(2).Height = DataGrid1.RowHeight - 10
            btnBuscar(2).Top = alto + 5
            btnBuscar(2).visible = b
            
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
    
'    'solo se puede modificar la factura si no esta contabilizada
'    If FactContabilizada Then Exit Sub
    
    Cad = "Cabecera de Facturas." & vbCrLf
    Cad = Cad & "-------------------------------------      " & vbCrLf & vbCrLf
    Cad = Cad & "Va a eliminar la Factura:            "
    Cad = Cad & vbCrLf & "Tipo:  " & Text1(6).Text
    Cad = Cad & vbCrLf & "Nº Factura:  " & Format(Text1(0).Text, "0000000")
    Cad = Cad & vbCrLf & "Fecha:  " & Format(Text1(1).Text, "dd/mm/yyyy")

    Cad = Cad & vbCrLf & vbCrLf & " ¿Desea Eliminarla? "

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

Private Sub Combo1_Click(Index As Integer)
   Select Case Index
        Case 1
            Text1(32).Enabled = (Combo1(Index).ListIndex <> 0)
            If Not Text1(32).Enabled Then Text1(32).Text = "1,0000"
    End Select
End Sub

Private Sub Combo1_GotFocus(Index As Integer)
    If Modo = 1 Then Combo1(Index).BackColor = vbLightBlue
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    If Combo1(Index).BackColor = vbLightBlue Then Combo1(Index).BackColor = vbWhite
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub DataGrid2_DblClick()
            frmVtasLinFacturas.TipoM = Me.AdoAux(0).Recordset!codTipoM
            frmVtasLinFacturas.NumFactu = Me.AdoAux(0).Recordset!NumFactu
            frmVtasLinFacturas.FecFactu = Me.AdoAux(0).Recordset!FecFactu
            frmVtasLinFacturas.NumLinea = Me.AdoAux(0).Recordset!NumLinea
            frmVtasLinFacturas.Albaran = Me.AdoAux(0).Recordset!NumAlbar
            frmVtasLinFacturas.Linea = Me.AdoAux(0).Recordset!numlinealbar
            frmVtasLinFacturas.ModoExt = 0
            frmVtasLinFacturas.Variedad = Text3(3).Text
            frmVtasLinFacturas.Confeccion = Text3(4).Text
            frmVtasLinFacturas.Show vbModal

End Sub

Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not Me.AdoAux(0).Recordset.EOF And ModificaLineas <> 1 Then
        CargarDatosAlbaran Me.AdoAux(0).Recordset!NumAlbar, AdoAux(0).Recordset!numlinealbar
    End If
End Sub

Private Sub DataGrid3_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not AdoAux(1).Recordset.EOF Then
        If Not IsNull(AdoAux(1).Recordset!ampliaci) Then
            Text2(16).Text = AdoAux(1).Recordset!ampliaci
        End If
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
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'l'1 i el 2 son separadors
        .Buttons(5).Image = 1   'Buscar
        .Buttons(6).Image = 2   'Totss
        'el 5 i el 6 son separadors
        .Buttons(1).Image = 3   'Insertar
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Borrar
        'el 10  son separadors
        .Buttons(8).Image = 10  'Imprimir
    End With
    
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
            
            If kCampo = 1 Then .Buttons(4).Image = 16 ' Insertar envases
                
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
    NombreTabla = "facturas"
    NomTablaLineas = "facturas_variedad" 'Tabla lineas de variedades
    
    '[Monica]03/02/2012: cambiamos el orden de las facturas, Picassent las quiere por fecha
    If vParamAplic.Cooperativa <> 2 Then
        Ordenacion = " ORDER BY facturas.codtipom, facturas.numfactu, facturas.fecfactu"
    Else
        Ordenacion = " ORDER BY facturas.fecfactu, facturas.numfactu, facturas.codtipom"
    End If
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    CadenaConsulta = "select * from facturas "
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
Dim i As Integer
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.Combo1(0).ListIndex = -1
    Me.Combo1(1).ListIndex = -1
    For i = 0 To 3
        Me.Check1(i).Value = 0
    Next i
    
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
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(6), CadenaDevuelta, 1)
        CadB = Aux
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 2)
        CadB = CadB & " and  " & Aux
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 3)
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

Private Sub frmFac_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String
Dim Aux As String
      
    If CadenaSeleccion <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        CadB = "codtipom = " & DBSet(RecuperaValor(CadenaSeleccion, 1), "T")
        CadB = CadB & " and numfactu = " & RecuperaValor(CadenaSeleccion, 2)
        CadB = CadB & " and fecfactu = " & DBSet(RecuperaValor(CadenaSeleccion, 3), "F")
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmFPag_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00") 'Codigo
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nombre
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String
Dim CADENA As String

    If CadenaSeleccion <> "" Then
        FacturasADescontar = " (codtipom, numfactu, fecfactu) in (" & CadenaSeleccion & ")"
    Else
        FacturasADescontar = ""
    End If
  
    If FacturasADescontar <> "" Then
        Sql = "insert into facturas_acuenta (codtipom,numfactu,fecfactu,codtipomcta,numfactucta,fecfactucta,totalfaccta)"
        Sql = Sql & " select " & DBSet(Text1(6).Text, "T") & "," & DBSet(Text1(0).Text, "N") & "," & DBSet(Text1(1).Text, "F") & ", codtipom, numfactu, fecfactu, totalfac from facturas "
        Sql = Sql & " where " & FacturasADescontar
        
        conn.Execute Sql
        
        '[Monica]28/08/2013: insertamos en el log que hemos insertado una linea en una factura contabilizada
        If Check1(1).Value = 1 Then
            '------------------------------------------------------------------------------
            '  LOG de acciones.
            Set LOG = New cLOG
            
            CADENA = "Inserta Linea de Facturas a Cuenta "
            CADENA = CADENA & Text1(6).Text & " " & Text1(0).Text & " " & Text1(1).Text & " de " & Text1(25).Text & " Fras." & CadenaSeleccion
            
            LOG.Insertar 9, vUsu, CADENA
            Set LOG = Nothing
              '-----------------------------------------------------------------------------
        End If
        
    End If
  
End Sub


Private Sub frmZ_Actualizar(vCampo As String)
     Text1(indice).Text = vCampo
End Sub

Private Sub imgBuscar_Click(Index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. de Cliente
            indice = 3
            PonerFoco Text1(indice)
            Set frmCli = New frmClientes
            frmCli.DatosADevolverBusqueda = "0|1|2|"
            frmCli.Show vbModal
            Set frmCli = Nothing
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
    '[Monica]28/08/2013: añadida la pregunta de continuar a pesar de estar en la contabilidad y arimoney
                '[Monica]05/10/2018: no decimos nada si es una factura en b
    If Check1(1).Value = 1 And Text1(6).Text <> TipoFactB Then
        If MsgBox("Esta factura está en Contabilidad y Arimoney. " & vbCrLf & vbCrLf & "Si la modifica realice los cambios en estas aplicaciones." & vbCrLf & vbCrLf & "¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Sub
        Else
            If CDate(Text1(1).Text) <= vEmpresa.FechaUltIVA Then
                If MsgBox("La factura es de un período liquidado. " & vbCrLf & vbCrLf & "¿ Seguro que desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                    Exit Sub
                End If
            End If
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
Dim Sql As String
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
            
        Case 3 'Cliente
            If PonerFormatoEntero(Text1(Index)) Then
                '[Monica]28/08/2013: añado or modo= 4
                If Modo = 1 Or Modo = 4 Then
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), "clientes", "nomclien")
                Else
                    PonerDatosCliente (Text1(Index).Text)
                    If Text2(Index).Text = "" Then
                        cadMen = "No existe el Cliente: " & Text1(Index).Text & vbCrLf
                        cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                        If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                            Set frmCli = New frmClientes
                            frmCli.DatosADevolverBusqueda = "0|1|"
                            Text1(Index).Text = ""
                            TerminaBloquear
                            frmCli.Show vbModal
                            Set frmCli = Nothing
                            If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                        Else
                            Text1(Index).Text = ""
                        End If
                        PonerFoco Text1(Index)
                    Else
                        If (TipoFactura = 1 And Modo = 3) Then '[Monica]31/07/2012: añadidos los enables
                            Text1(6).Enabled = True
                            Text1(0).Enabled = True
                            PonerFoco Text1(6)
                        Else
                            Text1(6).Enabled = False
                            Text1(0).Enabled = False
                            PonerFoco Text1(1)
                        End If
                    End If
               End If
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
            If Modo = 1 Then Exit Sub
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
                        If vTipoMov.Leer(Text1(6).Text) Then
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
    
         Case 32 ' cambio de divisa
            PonerFormatoDecimal Text1(Index), 7
    

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
    
'--monica
'    CadB = ObtenerBusqueda(Me)
'++monica
    If Facturas = "" Then
        CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    Else
        CadB = Facturas
    End If
    
    '[Monica]22/06/2010 seleccionamos unicamente las facturas que no sean de tipo a cuenta
    If CadB <> "" Then
        CadB = CadB & " and facturas.codtipom <> 'EAC' "
    Else
        CadB = "facturas.codtipom <> 'EAC'"
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
Dim Cad As String
Dim tabla As String
Dim Titulo As String
Dim Desc As String, devuelve As String
'    'Llamamos a al form
'    '##A mano
'    Cad = ""
'    Cad = Cad & "Tipo|facturas.codtipom|N||10·"
'    Cad = Cad & "Nº.Factura|facturas.numfactu|N||15·"
'    Cad = Cad & "Cliente|facturas.codclien|N||10·" 'ParaGrid(Text1(3), 10, "Cliente")
'    Cad = Cad & "Nombre Cliente|clientes.nomclien|N||45·"
'    Cad = Cad & ParaGrid(Text1(1), 15, "F.Factura")
'    tabla = NombreTabla & " INNER JOIN clientes ON facturas.codclien=clientes.codclien "
'
'    Titulo = "Facturas"
'    devuelve = "0|1|4|"
'
'    If Cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = Cad
'        frmB.vtabla = tabla
'        frmB.vSQL = CadB
'        HaDevueltoDatos = False
'        '###A mano
'        frmB.vDevuelve = "0|1|4|"
'        frmB.vDevuelve = devuelve
'        frmB.vTitulo = Titulo
'        frmB.vSelElem = 0
''        frmB.vConexionGrid = cAgro  'Conexión a BD: Ariagro
'        If Not EsCabecera Then frmB.Label1.FontSize = 11
''        frmB.vBuscaPrevia = chkVistaPrevia
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
''        If EsCabecera Then
''            PonerCadenaBusqueda
''            Text1(0).Text = Format(Text1(0).Text, "0000000")
''        End If
'        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
'        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
'            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
'        End If
'    End If
'    Screen.MousePointer = vbDefault
'
    Set frmFac = New frmBasico2

    AyudaFacturas frmFac, , CadB

    Set frmFac = Nothing




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
                    CargaGrid DataGrid2, AdoAux(0), True
                Else
                    CargaGrid DataGrid2, AdoAux(0), False
                End If
                If Not AdoAux(0).Recordset.EOF Then CargarDatosAlbaran AdoAux(0).Recordset!NumAlbar, AdoAux(0).Recordset!numlinealbar
            Case 1  ' envases
                If Data1.Recordset.RecordCount > 0 Then
                    CargaGrid DataGrid3, AdoAux(1), True
                Else
                    CargaGrid DataGrid3, AdoAux(1), False
                End If
                If Not AdoAux(1).Recordset.EOF Then
                    Text2(16).Text = DBLet(AdoAux(1).Recordset!ampliaci, "T")
                Else
                    Text2(16).Text = ""
                End If
            Case 2  ' facturas a cuenta
                If Data1.Recordset.RecordCount > 0 Then
                    CargaGrid DataGrid1, AdoAux(2), True
                Else
                    CargaGrid DataGrid1, AdoAux(2), False
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
    
    TipoFactura = DevuelveDesdeBDNew(cAgro, "clientes", "tipofact", "codclien", Text1(3).Text, "N")
    
    Text2(3).Text = PonerNombreDeCod(Text1(3), "clientes", "nomclien", "codclien", "N") 'cliente
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
Dim i As Byte, NumReg As Byte
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
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, Numreg
    DesplazamientoVisible b And Data1.Recordset.RecordCount > 1
        
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
    BloquearCmb Me.Combo1(1), Not (Modo = 1 Or Modo = 3 Or Modo = 4)
    BloquearChk Me.Check1(0), (Modo <> 1)
    BloquearChk Me.Check1(1), (Modo <> 1)
    
    '[Monica]10/09/2018: para el caso de Natural de momento he de dejarle que lo marquen como contabilizada
    '
    '                   ESTO LO TENDREMOS QUE QUITAR CUANDO HAYAN INTRODUCIDO TODAS LAS FACTURAS
    '
    If vParamAplic.Cooperativa = 9 And vUsu.Nivel = 0 And (Modo = 4 Or Modo = 1) Then
        BloquearChk Me.Check1(1), False
    End If
    ' HASTA AQUI
    
    BloquearChk Me.Check1(2), Not (Modo = 1 Or Modo = 4)
    BloquearChk Me.Check1(3), (Modo <> 1)
    
    imgFec(0).Enabled = b
    imgFec(0).visible = b
    
    
    '[Monica]20/09/2012:desbloqueo a mano el campo de fecha para poder modificarlo
    If Modo = 4 Then
        Text1(1).Locked = False
        Text1(1).BackColor = vbWhite
        
        '[Monica]28/08/2013: desbloqueo a mano si estamos modificando
        If Not HayFacturasACuenta(Text1(6).Text, Text1(0).Text, Text1(1).Text) Then
            Text1(3).Locked = False
            Text1(3).BackColor = vbWhite
        End If
        
        Text1(4).Locked = False
        Text1(4).BackColor = vbWhite
        Text1(5).Locked = False
        Text1(5).BackColor = vbWhite
        Text1(7).Locked = False
        Text1(7).BackColor = vbWhite
        Text1(8).Locked = False
        Text1(8).BackColor = vbWhite
        
    End If
    
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
    
    
    For i = 0 To txtAux1.Count - 1
        txtAux1(i).visible = False
        BloquearTxt txtAux1(i), True
    Next i
    
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    BloquearImgBuscar Me, Modo, ModificaLineas
    BloquearImgFec Me, 0, Modo
                    
    '[Monica]20/09/2012:desbloqueo a mano el campo de fecha para poder modificarlo
    If Modo = 4 And HayFacturasACuenta(Text1(6).Text, Text1(0).Text, Text1(1).Text) Then
        imgBuscar(0).Enabled = False
        imgBuscar(0).visible = False
    End If
                    
                    
                    
                    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
    Me.chkRectifica.Enabled = (Modo <= 2)
    Me.chkBio.Enabled = (Modo <= 2)
    
    If Modo <> 3 Then Me.chkRectifica.Value = 0
    If Modo <> 3 Then Me.chkBio.Value = 0
    
    ' *** si n'hi han llínies i algún tab que no te datagrid ***
    Select Case NumTabMto
        Case 0
            BloquearFrameAux Me, "FrameAux0", Modo, NumTabMto
        Case 1
            BloquearFrameAux Me, "FrameAux1", Modo, NumTabMto
        Case 2
            BloquearFrameAux Me, "FrameAux2", Modo, NumTabMto
    End Select
    'If Modo = 3 And Me.chkRectifica.Value = 1 Then Me.SSTab1.Tab = 3
    
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
Dim Serie As String
Dim Sql As String
Dim Rs As ADODB.Recordset

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

        'comprobamos que no exista ya la factura en la tabla facturas de ariagro
        Sql = ""
        '[Monica]04/03/2013: antes no me miraba la fecha de factura, estaba comentado
        Sql = DevuelveDesdeBDNew(cAgro, "facturas", "numfactu", "codtipom", Text1(6).Text, "T", , "numfactu", Text1(0).Text, "N", "fecfactu", Text1(1).Text, "F")
        If Sql <> "" Then
            MsgBox "Factura ya existente. Reintroduzca.", vbExclamation
            PonerFoco Text1(0)
            b = False
        End If
        If Not b Then Exit Function
        'comprobamos que no exista ya en la tabla facturas de contabilidad
        Serie = ""
'--monica:10/02/2009 stipom
'        Serie = DevuelveDesdeBDNew(cAgro, "stipom", "letraser", "codtipom", Text1(6).Text, "T")
'++monica
        Serie = ObtenerLetraSerie(Text1(6).Text)
'++
        If Serie <> "" Then
            Sql = ""
            '[Monica]04/03/2013: en la contabilidad hemos de mirar el año de la factura, no la fecha
            If vParamAplic.ContabilidadNueva Then
                Sql = DevuelveDesdeBDNew(cConta, "factcli", "numfactu", "numserie", Serie, "T", , "numfactu", Text1(0).Text, "N", "anofactu", Year(CDate(Text1(1).Text)), "N")
            Else
                Sql = DevuelveDesdeBDNew(cConta, "cabfact", "codfaccl", "numserie", Serie, "T", , "codfaccl", Text1(0).Text, "N", "anofaccl", Year(CDate(Text1(1).Text)), "N")
            End If
            If Sql <> "" Then
                MsgBox "Factura existente en contabilidad. Reintroduzca.", vbExclamation
                PonerFoco Text1(0)
                b = False
            End If
        Else
            MsgBox "El tipo de factura no tiene serie asociada. Revise.", vbExclamation
            b = False
        End If
        If Not b Then Exit Function
    
    
        '[Monica]20/06/2017: control de fechas que antes no estaba
        If vParamAplic.NumeroConta <> 0 Then
            ResultadoFechaContaOK = EsFechaOKConta(CDate(Text1(1).Text))
            If ResultadoFechaContaOK > 0 Then
                If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
                Exit Function
            End If
        End If
    
    
    
    End If
    
    '[Monica]27/01/2015: damos aviso de si hay albaranes que los cambie por los del nuevo cliente
    If Modo = 4 And ClienteAnt <> Text1(3).Text Then
    
        Sql = "select distinct codclien from albaran where numalbar in (select numalbar from facturas_variedad where codtipom = " & DBSet(Text1(6).Text, "T") & " and numfactu = " & DBSet(Text1(0).Text, "N") & " and fecfactu = " & DBSet(Text1(1).Text, "F") & ")"
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF And b
            If Rs.Fields(0).Value <> Text1(3).Text Then
                MsgBox "Recuerde que los albaranes incluidos han de ser del nuevo cliente. Revise.", vbExclamation
                b = False
            End If
            Rs.MoveNext
        Wend
        Set Rs = Nothing
        
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

    '[Monica]28/08/2013: añadida la pregunta de continuar a pesar de estar en la contabilidad y arimoney
        '[Monica]05/10/2018: en el caso de que sea b no damos mensajes
    If Check1(1).Value = 1 And Text1(6).Text <> TipoFactB Then
        If MsgBox("Esta factura está en Contabilidad y Arimoney. " & vbCrLf & vbCrLf & "Si la modifica realice los cambios en estas aplicaciones." & vbCrLf & vbCrLf & "¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Sub
        Else
            If CDate(Text1(1).Text) <= vEmpresa.FechaUltIVA Then
                If MsgBox("La factura es de un período liquidado. " & vbCrLf & vbCrLf & "¿ Seguro que desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                    Exit Sub
                End If
            End If
        End If
    End If
    
    If BloqueaRegistro(NombreTabla, "numfactu = " & Data1.Recordset!NumFactu) Then
'    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
        Select Case Button.Index
            Case 1
                BotonAnyadirLinea Index
            Case 2
                BotonModificarLinea Index
            Case 3
                BotonEliminarLinea Index
            Case 4
                BotonActualizarLinea
            Case Else
        End Select
    End If

End Sub

Private Sub BotonActualizarLinea()
Dim Sql As String
Dim NumLinea As Long
Dim CadValues As String
Dim Rs As ADODB.Recordset
Dim Precio As Currency

    Sql = "select albaran_envase.numalbar, albaran_envase.numlinea, albaran.fechaalb, albaran_envase.codartic, sartic.codigiva, sum(albaran_envase.cantidad) cantidad, sum(albaran_envase.impfianza) impfianza from (albaran_envase inner join albaran on albaran_envase.numalbar = albaran.numalbar) inner join sartic on albaran_envase.codartic = sartic.codartic where impfianza <> 0 and tipomovi = 0 and (albaran_envase.numalbar) "
    Sql = Sql & " in (select numalbar from facturas_variedad where numfactu = " & DBSet(Text1(0).Text, "N")
    Sql = Sql & " and codtipom = " & DBSet(Text1(6).Text, "T") & " and fecfactu = " & DBSet(Text1(1).Text, "F") & ")"
    Sql = Sql & " group by 1,2,3,4,5 "
    Sql = Sql & " order by 1,2,3,4,5 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    NumLinea = 0
    CadValues = ""
    While Not Rs.EOF
        NumLinea = NumLinea + 1
        Precio = 0
        If DBLet(Rs!Cantidad) <> 0 Then
            Precio = Round2(DBLet(Rs!ImpFianza) / DBLet(Rs!Cantidad), 4)
        End If
        
        txtAux(0).Text = Text1(6).Text
        txtAux(1).Text = Text1(0).Text
        txtAux(2).Text = Text1(1).Text
        txtAux(3).Text = NumLinea
        txtAux(4).Text = vParamAplic.Almacen
        txtAux(5).Text = DBLet(Rs!codArtic)
        txtAux(6).Text = DBLet(Rs!Cantidad)
        txtAux(7).Text = Precio
        txtAux(8).Text = "0"
        txtAux(9).Text = DBLet(Rs!ImpFianza)
        
        '[Monica]20/10/2016: si era exento no traia correcto el iva
        If Combo1(0).ListIndex = 1 Then
            txtAux(10).Text = vParamAplic.CodIvaExento
        Else
            txtAux(10).Text = DBLet(Rs!Codigiva)
        End If
        
        txtAux(11).Text = DBLet(Rs!NumAlbar)
        txtAux(12).Text = DBLet(Rs!NumLinea)
        txtAux(13).Text = DBLet(Rs!FechaAlb)
        
        ModificaLineas = 1
        If InsertarLineaEnv(txtAux(3).Text) Then
             CalcularDatosFactura
'                b = BloqueaRegistro("facturas", "numfactu = " & Data1.Recordset!NumFactu)
        End If
        ModificaLineas = 0
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing

'    Sql2 = "insert into facturas_envases (codtipom,numfactu,fecfactu,numlinea,codalmac,codartic,cantidad,precioar,dtolinea,importel,ampliaci,codigiva,numalbar,numlinealbar,fecalbar) values "
'    Sql2 = Sql2 & Sql
    
End Sub

Private Sub BotonEliminarLinea(Index As Integer)
Dim Cad As String
Dim Sql As String
Dim Mens As String
Dim b As Boolean
Dim CADENA As String

    On Error GoTo EEliminarLinea

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
    ' ***************************************************************************

'    '[Monica]23/03/2012: sólo se puede modificar
'    'solo se puede modificar la factura si no esta contabilizada
'    If FactContabilizada(True) Then
'        TerminaBloquear
'        Exit Sub
'    End If

    b = True

    Select Case Index
        Case 0 'variedades
            ' *************** canviar la pregunta ****************
            Cad = "¿Seguro que desea eliminar la Variedad?"
            Cad = Cad & vbCrLf & "Factura: " & AdoAux(0).Recordset.Fields(0)
            Cad = Cad & vbCrLf & "Albarán: " & AdoAux(0).Recordset.Fields(4) & "-" & AdoAux(0).Recordset.Fields(5)
            
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
                On Error GoTo EEliminarLinea
                Screen.MousePointer = vbHourglass
                NumRegElim = AdoAux(0).Recordset.AbsolutePosition
                TerminaBloquear
                
                
                If Not EliminarLineaVariedades Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                
                Else
                    CalcularDatosFactura
                
                    SituarDataTrasEliminar AdoAux(0), NumRegElim
                    
                    If Me.AdoAux(0).Recordset.EOF Then
                        CargarDatosAlbaran "", ""
                    End If
                    
                    CargaGrid DataGrid2, AdoAux(0), True
                    SSTab1.Tab = 0
                End If
            End If
            Screen.MousePointer = vbDefault
       
       Case 1 'envases
            ' *************** canviar la pregunta ****************
            Cad = "¿Seguro que desea eliminar el Envase?"
            Cad = Cad & vbCrLf & "Factura: " & AdoAux(1).Recordset.Fields(1)
            Cad = Cad & vbCrLf & "Artículo: " & AdoAux(1).Recordset.Fields(5) & " - " & AdoAux(1).Recordset.Fields(6)
            
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
                On Error GoTo EEliminarLinea
                Screen.MousePointer = vbHourglass
                NumRegElim = AdoAux(1).Recordset.AbsolutePosition
                
                If Not EliminarLinea Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                Else
                    CalcularDatosFactura
                    If SituarDataTrasEliminar(AdoAux(1), NumRegElim) Then
                        PonerCampos
                    Else
                        PonerCampos
'                        LimpiarCampos
'                        PonerModo 0
                    End If
                End If
            End If
            Screen.MousePointer = vbDefault
       
       Case 2 'facturas a cuenta
            ' *************** canviar la pregunta ****************
            Cad = "¿Seguro que desea eliminar la factura a cuenta?"
            Cad = Cad & vbCrLf & "Factura: " & AdoAux(2).Recordset.Fields(4)
            
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
                On Error GoTo EEliminarLinea
                Screen.MousePointer = vbHourglass
                NumRegElim = AdoAux(2).Recordset.AbsolutePosition
                
                If Not EliminarLineaFacCta Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                Else
                
                    CalcularDatosFactura
                    
                    If SituarDataTrasEliminar(AdoAux(2), NumRegElim) Then
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
        Case 1  'Añadir
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
        Case 5  'Buscar
            mnBuscar_Click
        Case 6  'Todos
            BotonVerTodos
        Case 8  ' Impresion de albaran
            mnImprimir_Click
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

    b = DataGrid3.Enabled
    Select Case vDataGrid.Name
        Case "DataGrid2" 'variedades
            Opcion = 0
        Case "DataGrid3" 'envases
            Opcion = 1
        Case "DataGrid1"  'facturas a cuenta
            Opcion = 2
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
                       
         Case "DataGrid2" 'facturas_variedad
'select codtipom,numfactu,fecfactu,numlinea,numalbar,numlinealbar,cantreal,cantfact,precibru,precinet,
'dtocom1,dtocom2,imporbru,impornet,codigiva
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux3(4)|T|Albarán|1000|;S|txtAux3(5)|T|Linea|700|;"
            tots = tots & "S|btnBuscar(1)|B|||;S|txtAux3(6)|T|Cant.Real|1200|;S|txtAux3(7)|T|Cant.Fact.|1200|;S|txtAux3(15)|T|Uds|700|;S|txtAux3(8)|T|Prec.Bruto|1300|;"
            tots = tots & "S|txtAux3(9)|T|Prec.Neto|1200|;S|txtAux3(10)|T|Importe Bruto|1550|;"
            tots = tots & "S|txtAux3(11)|T|Importe Neto|1500|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
            arregla tots, DataGrid2, Me, 350
            
            
                     
         Case "DataGrid3" 'facturas_envases
'select codtipom,numfactu,fecfactu,numlinea,codartic,sartic.nomartic,cantidad,precioar,dtolinea,importel,ampliaci,codigiva,numalbar,numlinea
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;S|txtAux(4)|T|Alm|700|;"
            tots = tots & "S|txtAux(5)|T|Articulo|2000|;S|btnBuscar(0)|B|||;"
            tots = tots & "S|Text2(0)|T|Nombre|5000|;S|txtAux(6)|T|Cantidad|1900|;"
            tots = tots & "S|txtAux(7)|T|Precio|1900|;S|txtAux(8)|T|Dto|1100|;S|txtAux(9)|T|Importe|1900|;N||||0|;N||||0|;S|txtAux(11)|T|Albaran|1450|;N||||0|;N||||0|;"
            arregla tots, DataGrid3, Me, 350
            
            
            If AdoAux(0).Recordset.EOF Then CargarDatosAlbaran "", ""
            
     
         Case "DataGrid1" 'facturas_acuenta
'select codtipom,numfactu,fecfactu,numlinea,codartic,sartic.nomartic,cantidad,precioar,dtolinea,importel,ampliaci,codigiva
            tots = "N||||0|;N||||0|;N||||0|;S|txtAux1(3)|T|Tipo Mov.|1100|;"
            tots = tots & "S|txtAux1(4)|T|Factura|1500|;S|txtAux1(5)|T|Fecha |1500|;S|btnBuscar(2)|B|||;S|txtAux1(6)|T|Total Factura|2500|;"
            tots = tots & "S|txtAux1(7)|T|Base Imponible|2500|;S|txtAux1(8)|T|%Iva|850|;S|txtAux1(9)|T|Importe Iva|2500|;S|txtAux1(10)|T|%Rec|850|;"
            tots = tots & "S|txtAux1(11)|T|Imp.Retencion|2500|;"
            arregla tots, DataGrid1, Me, 350
            
    End Select
    
    vDataGrid.HoldFields
    Exit Sub
    
ECargaGRid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

'
'Private Sub TxtAux_Change(Index As Integer)
'    If Index = 6 And ModificaLineas = 2 Then 'Precio y Modo Borrar Lineas
'        txtAux(5).Text = "M"
'    End If
'End Sub

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
Dim cadMen As String
Dim Sql As String
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
                If Not AdoAux(1).Recordset.EOF Then devuelve = AdoAux(1).Recordset!codArtic
            End If
        
            If Not PonerArticulo(txtAux(5), Text2(0), txtAux(4).Text, CodTipoMov, ModificaLineas, devuelve) Then
                PonerFoco txtAux(Index)
            Else
                
                '[Monica]15/05/2015: si modificamos no traemos nada
                If ModificaLineas = 2 Then Exit Sub
                
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
'[Monica]22/05/2012: quitamos lo de control de stock
'                Set vCStock = New CStock
'                If Not InicializarCStock(vCStock, "S") Then Exit Sub
'                If vCStock.MueveStock Then 'Comprobar si el articulo mueve stock: tiene control de stock y no es instalacion
'                  If Not vCStock.MoverStock Then
'                    PonerFoco txtAux(Index)
'                    Set vCStock = Nothing
'                    Exit Sub
'                  End If
'                End If
'
'                Set vCStock = Nothing
            End If
            
        Case 7 ' Precio
            PonerFormatoDecimal txtAux(Index), 10
            
        Case 8  'Descuentos
            PonerFormatoDecimal txtAux(Index), 4 'Tipo 4: Decimal(4,2)
            
        Case 9 'Importe Linea
            PonerFormatoDecimal txtAux(Index), 3 'Tipo 3: Decimal(10,2)
            
        Case 11 ' número de albarán de donde viene el envase (albaran de envase o albaran de variedades)
            PonerFormatoEntero txtAux(Index)
            
    End Select
    If (Index = 6 Or Index = 7 Or Index = 8 Or Index = 9) Then 'Cant., Precio, Dto1, Dto2
'        If Trim(TxtAux(3).Text) = "" Or Trim(TxtAux(4).Text) = "" Then Exit Sub
'        If Trim(TxtAux(6).Text) = "" Or Trim(TxtAux(7).Text) = "" Then Exit Sub
        If txtAux(8).Text = "" Then txtAux(8).Text = 0
        TipoDto = DevuelveDesdeBDNew(cAgro, "clientes", "tipodtos", "codclien", Text1(3).Text, "N")
        
        txtAux(9).Text = CalcularImporteFClien(txtAux(6).Text, txtAux(7).Text, txtAux(8).Text, 0, TipoDto, 0)
        PonerFormatoDecimal txtAux(9), 3
    End If
    
End Sub




Private Function Eliminar() As Boolean
Dim Sql As String, LEtra As String, Sql2 As String
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
        Sql = " " & ObtenerWhereCP(True)
        
        'Lineas de calibres (facturas_calibre)
        conn.Execute "Delete from facturas_calibre " & Replace(Sql, "facturas", "facturas_calibre")
        
        'Lineas de envases (facturas_variedad)
        conn.Execute "Delete from facturas_variedad " & Replace(Sql, "facturas", "facturas_variedad")
        
        'Lineas de coste (facturas_envases)
        conn.Execute "Delete from facturas_envases " & Replace(Sql, "facturas", "facturas_envases")
        
        'Lineas de facturas a cuenta (facturas_acuenta)
        conn.Execute "delete from facturas_acuenta " & Replace(Sql, "facturas", "facturas_acuenta")
    

        'Cabecera de factura
        conn.Execute "Delete from " & NombreTabla & Sql
        
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
Dim Sql As String, LEtra As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim Mens As String
Dim vCStock As CStock
Dim CADENA As String

    On Error GoTo FinEliminar

    b = False
    If AdoAux(1).Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    Mens = ""
    
    
    '[Monica]28/08/2013: insertamos en el log que hemos eliminado una linea en una factura contabilizada
    If Check1(1).Value = 1 Then
        '------------------------------------------------------------------------------
        '  LOG de acciones.
        Set LOG = New cLOG
        'campo = "Facturas de Clientes: " & """"
        
        CADENA = "Eliminar Linea de Envases "
        LOG.Insertar 9, vUsu, CADENA & AdoAux(1).Recordset.Fields(0) & " " & AdoAux(1).Recordset.Fields(1) & " " & AdoAux(1).Recordset.Fields(2) & " de " & Text1(25).Text & " Linea " & AdoAux(1).Recordset.Fields(3) & " " & AdoAux(1).Recordset.Fields(5)
        Set LOG = Nothing
          '-----------------------------------------------------------------------------
    End If
    
    
    
    
    
    'Eliminar en tablas de facturas_envases
    '------------------------------------------
    Sql = " where codtipom = " & DBSet(AdoAux(1).Recordset.Fields(0), "T")
    Sql = Sql & " and numfactu = " & AdoAux(1).Recordset.Fields(1)
    Sql = Sql & " and fecfactu = " & DBSet(AdoAux(1).Recordset.Fields(2), "F")
    Sql = Sql & " and numlinea = " & DBSet(AdoAux(1).Recordset.Fields(3), "N")


     ' borramos el movimiento y aumentamos el stock
    Set vCStock = New CStock
    
    txtAux(11).Text = DBLet(AdoAux(1).Recordset!NumAlbar, "N")
    txtAux(12).Text = DBLet(AdoAux(1).Recordset!numlinealbar, "N")
    txtAux(13).Text = DBLet(AdoAux(1).Recordset!FecAlbar, "F")
    
    
    If Not InicializarCStock(vCStock, "E") Then Exit Function

     'en actualizar stock comprobamos si el articulo tiene control de stock
     b = vCStock.DevolverStock
     Set vCStock = Nothing

    'Lineas de variedades
    conn.Execute "Delete from facturas_envases " & Sql
    
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
Dim Sql As String, LEtra As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim Mens As String
Dim vCStock As CStock
Dim CADENA As String

    On Error GoTo FinEliminar

    b = False
    If AdoAux(2).Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    Mens = ""
    
    b = True
    
    
    '[Monica]28/08/2013: insertamos en el log que hemos insertado una linea en una factura contabilizada
    If Check1(1).Value = 1 Then
        '------------------------------------------------------------------------------
        '  LOG de acciones.
        Set LOG = New cLOG
        'campo = "Facturas de Clientes: " & """"
        
        CADENA = "Eliminar Linea Facturas a Cuenta "
        LOG.Insertar 9, vUsu, CADENA & AdoAux(2).Recordset.Fields(0) & " " & AdoAux(2).Recordset.Fields(1) & " " & AdoAux(2).Recordset.Fields(2) & " de " & Text1(25).Text & " " & AdoAux(2).Recordset.Fields(3) & " " & AdoAux(2).Recordset.Fields(4) & " " & AdoAux(2).Recordset.Fields(5)
        Set LOG = Nothing
          '-----------------------------------------------------------------------------
    End If
    
    'Eliminar en tablas de facturas_acuenta
    '------------------------------------------
    Sql = " where codtipom = " & DBSet(AdoAux(2).Recordset.Fields(0), "T")
    Sql = Sql & " and numfactu = " & AdoAux(2).Recordset.Fields(1)
    Sql = Sql & " and fecfactu = " & DBSet(AdoAux(2).Recordset.Fields(2), "F")
    Sql = Sql & " and codtipomcta = " & DBSet(AdoAux(2).Recordset.Fields(3), "T")
    Sql = Sql & " and numfactucta = " & AdoAux(2).Recordset.Fields(4)
    Sql = Sql & " and fecfactucta = " & DBSet(AdoAux(2).Recordset.Fields(5), "F")


    'Lineas de variedades
    conn.Execute "Delete from facturas_acuenta " & Sql
    
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



Private Function EliminarLineaVariedades() As Boolean
Dim Sql As String, LEtra As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim Mens As String
Dim vCStock As CStock
Dim CADENA As String

    On Error GoTo FinEliminar

    b = False
    If AdoAux(0).Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    Mens = ""
   
   '[Monica]28/08/2013: insertamos en el log que hemos eliminado una linea en una factura contabilizada
    If Check1(1).Value = 1 Then
        '------------------------------------------------------------------------------
        '  LOG de acciones.
        Set LOG = New cLOG
        'campo = "Facturas de Clientes: " & """"
        
        CADENA = "Eliminar Linea de Variedades "
        LOG.Insertar 9, vUsu, CADENA & AdoAux(0).Recordset.Fields(0) & " " & AdoAux(0).Recordset.Fields(1) & " " & AdoAux(0).Recordset.Fields(2) & " de " & Text1(25).Text & " Linea " & AdoAux(0).Recordset.Fields(3) & " Alb " & AdoAux(0).Recordset.Fields(4) & " " & AdoAux(0).Recordset.Fields(5)
        Set LOG = Nothing
          '-----------------------------------------------------------------------------
    End If
        
    Sql = "delete from facturas_calibre where codtipom = " & DBSet(AdoAux(0).Recordset.Fields(0), "T")
    Sql = Sql & " and numfactu = " & AdoAux(0).Recordset.Fields(1)
    Sql = Sql & " and fecfactu = " & DBSet(AdoAux(0).Recordset.Fields(2), "F")
    Sql = Sql & " and numlinea = " & DBSet(AdoAux(0).Recordset.Fields(3), "N")
    conn.Execute Sql
        
        
    Sql = "delete from facturas_variedad where codtipom = " & DBSet(AdoAux(0).Recordset.Fields(0), "T")
    Sql = Sql & " and numfactu = " & AdoAux(0).Recordset.Fields(1)
    Sql = Sql & " and fecfactu = " & DBSet(AdoAux(0).Recordset.Fields(2), "F")
    Sql = Sql & " and numlinea = " & DBSet(AdoAux(0).Recordset.Fields(3), "N")
    conn.Execute Sql
                   
    Mens = "Recalcular Dtos calibres"
    b = RecalcularDtosLineas(AdoAux(0).Recordset.Fields(0), AdoAux(0).Recordset.Fields(1), AdoAux(0).Recordset.Fields(2), Mens)
                   
    Mens = "Recalcular Dtos lineas"
    b = RecalcularDtos(AdoAux(0).Recordset.Fields(0), AdoAux(0).Recordset.Fields(1), AdoAux(0).Recordset.Fields(2), Mens)
    
FinEliminar:
    If Err.Number <> 0 Or Not b Then
        MuestraError Err.Number, "Eliminar Linea Albarán de la Factura ", Err.Description & " " & Mens
        b = False
    End If
    If Not b Then
        conn.RollbackTrans
        EliminarLineaVariedades = False
    Else
        conn.CommitTrans
        EliminarLineaVariedades = True
    End If
End Function






Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next

    CargaGrid DataGrid2, Me.AdoAux(1), False 'envases
    CargaGrid DataGrid3, Me.AdoAux(0), False 'variedades
    CargaGrid DataGrid1, Me.AdoAux(2), False 'facturas a cuenta
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & Replace(ObtenerWhereCP(False), "facturas.", "") & ")"
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
    
    Sql = "facturas.codtipom = " & DBSet(Text1(6).Text, "T") & " and facturas.numfactu= " & DBSet(Text1(0).Text, "N") & " and facturas.fecfactu= " & DBSet(Text1(1).Text, "F")
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
        Case 0  'variedades
'select codtipom,numfactu,fecfactu,numlinea,numalbar,numlinealbar,cantreal,cantfact,precibru,precinet,
'dtocom1,dtocom2,imporbru,impornet,codigiva
            Sql = "SELECT facturas_variedad.codtipom,numfactu,fecfactu,facturas_variedad.numlinea,"
            Sql = Sql & " facturas_variedad.numalbar,numlinealbar,cantreal,"
            Sql = Sql & " cantfact,facturas_variedad.unidades, precibru, precinet, imporbru,impornet, fechaalb, matrirem, "
            Sql = Sql & " destinos.nomdesti,variedades.nomvarie, forfaits.nomconfe, dtocom1, dtocom2, facturas_variedad.codigiva  "
            Sql = Sql & " FROM facturas_variedad, albaran, albaran_variedad, variedades, forfaits, destinos " 'lineas de variedades de la factura
            Sql = Sql & " WHERE facturas_variedad.numalbar = albaran.numalbar "
            Sql = Sql & " and albaran.numalbar = albaran_variedad.numalbar "
            Sql = Sql & " and facturas_variedad.numlinealbar = albaran_variedad.numlinea "
            Sql = Sql & " and albaran_variedad.codvarie = variedades.codvarie "
            Sql = Sql & " and albaran_variedad.codforfait = forfaits.codforfait "
            Sql = Sql & " and albaran.codclien = destinos.codclien "
            Sql = Sql & " and albaran.coddesti = destinos.coddesti "
            
            If enlaza Then
                Sql = Sql & " and " & Replace(ObtenerWhereCP(False), "facturas", "facturas_variedad")
            Else
                '[Monica]19/04/2017: cambio de condicion por rapidez
                Sql = Sql & " and numfactu is null " '= -1"
            End If
            Sql = Sql & " ORDER BY codtipom,numfactu,fecfactu,numlinea"
                    
        Case 1  'envases
'select codtipom,numfactu,fecfactu,numlinea,codalmac,codartic,sartic.nomartic,cantidad,precioar,dtolinea,importel,ampliaci,codigiva
            Sql = "SELECT codtipom,numfactu,fecfactu,numlinea,codalmac,facturas_envases.codartic,sartic.nomartic,cantidad,"
            Sql = Sql & "precioar,dtolinea,importel,ampliaci,facturas_envases.codigiva,facturas_envases.numalbar,facturas_envases.numlinealbar, facturas_envases.fecalbar "
            Sql = Sql & " FROM facturas_envases, sartic "
            Sql = Sql & " WHERE facturas_envases.codartic = sartic.codartic "
    
            If enlaza Then
                Sql = Sql & " and " & Replace(ObtenerWhereCP(False), "facturas", "facturas_envases")
            Else
                '[Monica]19/04/2017: cambio de condicion por rapidez
                Sql = Sql & " and numfactu is null " ' = -1"
            End If
            Sql = Sql & " ORDER BY codtipom,numfactu,fecfactu,numlinea"
    
        Case 2 ' facturas a cuenta
            Sql = "SELECT facturas_acuenta.codtipom, facturas_acuenta.numfactu, facturas_acuenta.fecfactu, facturas_acuenta.codtipomcta, "
            Sql = Sql & " facturas_acuenta.numfactucta, facturas_acuenta.fecfactucta, facturas_acuenta.totalfaccta, facturas.baseimp1, facturas.porciva1, facturas.impoiva1,  "
            Sql = Sql & " facturas.porcrec1, facturas.imporec1 "
            Sql = Sql & " FROM facturas_acuenta, facturas "
            Sql = Sql & " WHERE facturas_acuenta.codtipomcta = facturas.codtipom  "
            Sql = Sql & " and facturas_acuenta.numfactucta = facturas.numfactu "
            Sql = Sql & " and facturas_acuenta.fecfactucta = facturas.fecfactu "
    
            If enlaza Then
                Sql = Sql & " and " & Replace(ObtenerWhereCP(False), "facturas", "facturas_acuenta")
            Else
                '[Monica]19/04/2017: cambio de condicion por rapidez
                Sql = Sql & " and facturas_acuenta.numfactu is null " '= -1"
            End If
            Sql = Sql & " ORDER BY facturas_acuenta.codtipom,facturas_acuenta.numfactu,facturas_acuenta.fecfactu,facturas_acuenta.codtipomcta,facturas_acuenta.numfactucta,facturas_acuenta.fecfactucta"
    
    End Select
    
    MontaSQLCarga = Sql
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean, bAux As Boolean
Dim i As Integer

        b = ((Modo = 2) Or (Modo = 0)) And (Facturas = "") And (hcoCodMovim = "") 'Or (Modo = 5 And ModificaLineas = 0)
        'Buscar
        Toolbar1.Buttons(5).Enabled = b
        Me.mnBuscar.Enabled = b
        'Vore Tots
        Toolbar1.Buttons(6).Enabled = b
        Me.mnVerTodos.Enabled = b
        'Añadir
        Toolbar1.Buttons(1).Enabled = b
        Me.mnModificar.Enabled = b
        
        
        b = (Modo = 2 And Data1.Recordset.RecordCount > 0) And (Facturas = "") And (hcoCodMovim = "")
        'Modificar
        '[Monica]28/08/2013: dejo modificar la factura aunque esté contabilizada y demas
        '                    quito la condicion de todos los checks
        Toolbar1.Buttons(2).Enabled = b 'And Not (Check1(0).Value = 1 Or Check1(1).Value = 1 Or Check1(2).Value = 1)
        Me.mnModificar.Enabled = b 'And Not (Check1(0).Value = 1 Or Check1(1).Value = 1 Or Check1(2).Value = 1)
        'eliminar
        Toolbar1.Buttons(3).Enabled = b And (Not (Check1(0).Value = 1 Or Check1(1).Value = 1 Or Check1(2).Value = 1) Or (Check1(1).Value = 1 And Text1(6).Text = TipoFactB))
        Me.mnEliminar.Enabled = b And (Not (Check1(0).Value = 1 Or Check1(1).Value = 1 Or Check1(2).Value = 1) Or (Check1(1).Value = 1 And Text1(6).Text = TipoFactB))
        'Impresión de factura
        Toolbar1.Buttons(8).Enabled = ((Modo = 2) And (Facturas = "")) Or (hcoCodMovim <> "")
        Me.mnImprimir.Enabled = ((Modo = 2) And (Facturas = "")) Or (hcoCodMovim <> "")

    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
    
    '[Monica]28/08/2013: dejo modificar la factura aunque esté contabilizada y demas
    '                    quito la condicion de todos los checks
    b = (Modo = 2) And (Facturas = "") And (hcoCodMovim = "")  ' And Not (Check1(0).Value = 1 Or (Check1(1).Value = 1 And vUsu.Nivel >= 1) Or Check1(2).Value = 1)
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = b
        
        If b Then
            Select Case i
              Case 0
                bAux = (b And Me.AdoAux(0).Recordset.RecordCount > 0)
              Case 1
                bAux = (b And Me.AdoAux(1).Recordset.RecordCount > 0)
              Case 2
                bAux = (b And Me.AdoAux(2).Recordset.RecordCount > 0)
            End Select
        End If
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i
    '[Monica]25/06/2012: actualizar precios de fianza
    ToolAux(1).Buttons(4).Enabled = b And (Me.AdoAux(1).Recordset.RecordCount = 0)


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
    indRPT = 12 'Impresion de Factura
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
      
    '[Monica]05/10/2018: en cso de que sea una factura en B no lleva ni iva ni logos
    If Text1(6).Text = TipoFactB Then nomDocu = Replace(nomDocu, ".rpt", "B.rpt")
      
      
    '[Monica]18/12/2018: en el caso de que sea una factura extranjera
    If Combo1(1).ListIndex > 0 Then nomDocu = Replace(nomDocu, ".rpt", "Extr.rpt")
    
      
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
     
    NroCopias = DevuelveValor("select nrocopias from clientes where codclien = " & DBSet(Text1(3).Text, "N"))
     
     
    With frmImprimir
          '[Monica]11/01/2012: añadido la siguientes 3 lineas para el envio por el outlook
            .outClaveNombreArchiv = Text1(6).Text & Format(Text1(0).Text, "0000000")
            .outCodigoCliProv = Text1(3).Text
            '[Monica]06/05/2015: destino para sacar email
            .outCodigoDestino = ""
            .outTipoDocumento = 2
     
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 0
            .Titulo = "Impresión de Factura"
            .ConSubInforme = True
            If vParamAplic.Cooperativa = 11 Then .NroCopias = NroCopias
            .Show vbModal
    End With
End Sub



Private Sub TxtAux3_GotFocus(Index As Integer)
    ConseguirFoco txtAux3(Index), Modo
End Sub

Private Sub TxtAux3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim V As Long

    If Index <> 0 And KeyCode <> 38 Then KEYdown KeyCode
    
    '[Monica]13/02/2012: si pulsamos F2 nos vamos a introducir las lineas por calibre
    If KeyCode = 113 Then
        If txtAux3(4).Text <> "" And txtAux3(5).Text <> "" Then
            PulsadoF2 = True
        
            frmVtasLinFacturas.TipoM = txtAux3(0).Text
            frmVtasLinFacturas.NumFactu = txtAux3(1).Text
            frmVtasLinFacturas.FecFactu = txtAux3(2).Text
            frmVtasLinFacturas.NumLinea = txtAux3(3).Text
            frmVtasLinFacturas.Albaran = txtAux3(4).Text
            frmVtasLinFacturas.Linea = txtAux3(5).Text
            frmVtasLinFacturas.ModoExt = 2
            frmVtasLinFacturas.ImpDtoc = Text1(5).Text
            frmVtasLinFacturas.Dto1 = Text1(7).Text
            frmVtasLinFacturas.Dto2 = Text1(8).Text
            frmVtasLinFacturas.Variedad = Text3(3).Text
            frmVtasLinFacturas.Confeccion = Text3(4).Text
            frmVtasLinFacturas.TipoIva = txtAux3(14).Text
            frmVtasLinFacturas.Show vbModal
            
            cmdCancelar_Click
            
            V = txtAux3(3).Text 'el 2 es el nº de llinia
                
            
            
            CargaGrid DataGrid2, AdoAux(0), True
            CalcularDatosFactura
            
            DataGrid2.SetFocus
            AdoAux(0).Recordset.Find (AdoAux(0).Recordset.Fields(3).Name & " =" & V)
            
            
            
        Else
            MsgBox "Debe de introducir el número de albaran y línea para introducir los importes por calibre.", vbExclamation
        End If
    End If
End Sub

Private Sub TxtAux3_KeyPress(Index As Integer, KeyAscii As Integer)
    
    KEYpress KeyAscii
End Sub

Private Sub TxtAux3_LostFocus(Index As Integer)
Dim TipoDto As Byte
Dim ImpDto As String
Dim Unidades As String
Dim Cantidad As String
Dim Cad As String
Dim campo2 As String

    'Quitar espacios en blanco
    If Not PerderFocoGnralLineas(txtAux3(Index), ModificaLineas) Then Exit Sub
    
    
    Select Case Index
        Case 4 'Albaran
            If txtAux3(Index) <> "" Then
                PonerFormatoEntero txtAux3(Index)
            Else
                Exit Sub
            End If
            If ModificaLineas = 2 Then Exit Sub
            
            CargarDatosAlbaran txtAux3(4).Text, txtAux3(5).Text
            TxtAux3_LostFocus (8)

            
        Case 5 'Linea de albaran
            If txtAux3(Index) <> "" Then PonerFormatoEntero txtAux3(Index)
            
            If txtAux3(4).Text <> "" And txtAux3(5).Text <> "" Then
                If AlbaranFacturado(txtAux3(4).Text, txtAux3(5).Text) Then
                    Cad = "Esta línea de Albarán está facturada. " & vbCrLf & vbCrLf & "    ¿ Desea continuar ? "
                    If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                        CargarDatosAlbaran txtAux3(4).Text, txtAux3(5).Text
                        TxtAux3_LostFocus (8)

                    Else
                        txtAux3(4).Text = ""
                        txtAux3(5).Text = ""
                    End If
                Else
                    CargarDatosAlbaran txtAux3(4).Text, txtAux3(5).Text
                    TxtAux3_LostFocus (8)

                End If
            End If
            
            If txtAux3(4).Text = "" Or txtAux3(5).Text = "" Then
                PonerFoco txtAux3(4)
            Else
                PonerFoco txtAux3(8)
            End If
            
        Case 6 ' Cantidad real
            PonerFormatoEntero txtAux3(Index)
        Case 7 ' Cantidad facturada
            PonerFormatoEntero txtAux3(Index)
        
        Case 8 'precio bruto
            If txtAux3(Index).Text <> "" Then
                If PonerFormatoDecimal(txtAux3(Index), 7) Then
                    
                    Select Case TipoFacturarForfaits(txtAux3(4).Text, txtAux3(5).Text)
                        Case 0  'por unidades
                            '[Monica]27/08/2013: he añadido comprobarcero
                            txtAux3(10).Text = Round2(CCur(ImporteSinFormato(txtAux3(Index).Text)) * CCur(ImporteSinFormato(ComprobarCero(txtAux3(15).Text))), 2)
                            PonerFormatoDecimal txtAux3(10), 3
                        Case 1  'por kilos
                            '[Monica]27/08/2013: he añadido comprobarcero
                            txtAux3(10).Text = Round2(CCur(ImporteSinFormato(txtAux3(Index).Text)) * CCur(ImporteSinFormato(ComprobarCero(txtAux3(6).Text))), 2)
                            PonerFormatoDecimal txtAux3(10), 3
                        Case Else
                            
                    End Select
                    
'                    cmdAceptar.SetFocus
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
                        
                    If Not PulsadoF2 Then cmdAceptar.SetFocus
               Else
                    Exit Sub
               End If
            End If
    End Select

    If ((Index = 8 And txtAux3(Index).Text <> "") Or (Index = 10 And txtAux3(Index).Text <> "")) Then
        campo2 = "nrodecprec"
        TipoDto = DevuelveDesdeBDNew(cAgro, "clientes", "tipodtos", "codclien", Text1(3).Text, "N", campo2)
        Select Case TipoFacturarForfaits(txtAux3(4).Text, txtAux3(5).Text)
            Case 0 ' unidades
'                ImpDto = CalcularImporteDto(txtAux3(15).Text, txtAux3(8).Text, txtAux3(0).Text, txtAux3(1).Text, txtAux3(2).Text, CStr(DBLet(Data1.Recordset!impdtoc, "N")), False)
'                txtAux3(11).Text = CalcularImporte(txtAux3(15).Text, txtAux3(8).Text, txtAux3(12).Text, txtAux3(13).Text, TipoDto, ImpDto)
                Unidades = ComprobarCero(txtAux3(15).Text)
                '[Monica]24/11/2011: añadida condicion para evitar division por cero
                If CCur(Unidades) <> 0 Then
                    ImpDto = CalcularImporteDto(txtAux3(15).Text, CStr(CCur(ImporteSinFormato(txtAux3(10).Text)) / CCur(Unidades)), txtAux3(0).Text, txtAux3(1).Text, txtAux3(2).Text, CStr(DBLet(Data1.Recordset!ImpDtoc, "N")), False)
                    txtAux3(11).Text = CalcularImporteFClien(txtAux3(15).Text, CStr(CCur(ImporteSinFormato(txtAux3(10).Text)) / CCur(Unidades)), txtAux3(12).Text, txtAux3(13).Text, TipoDto, ImpDto, txtAux3(10).Text)
                Else
                    ImpDto = CalcularImporteDto(txtAux3(15).Text, CStr(0), txtAux3(0).Text, txtAux3(1).Text, txtAux3(2).Text, CStr(DBLet(Data1.Recordset!ImpDtoc, "N")), False)
                    txtAux3(11).Text = CalcularImporteFClien(txtAux3(15).Text, CStr(0), txtAux3(12).Text, txtAux3(13).Text, TipoDto, ImpDto, txtAux3(10).Text)
                End If
                PonerFormatoDecimal txtAux3(11), 1
                
                'precio neto
                If ComprobarCero(txtAux3(15).Text) <> "0" Then
                    txtAux3(9).Text = Round2(CCur(ImporteSinFormato(txtAux3(11).Text)) / CCur(ImporteSinFormato(txtAux3(15).Text)), CCur(campo2))
                Else
                    txtAux3(9).Text = "0"
                End If
                PonerFormatoDecimal txtAux3(9), 7
            
            Case 1 ' kilos
'                ImpDto = CalcularImporteDto(txtAux3(6).Text, txtAux3(8).Text, txtAux3(0).Text, txtAux3(1).Text, txtAux3(2).Text, CStr(DBLet(Data1.Recordset!impdtoc, "N")), False)
'                txtAux3(11).Text = CalcularImporte(txtAux3(6).Text, txtAux3(8).Text, txtAux3(12).Text, txtAux3(13).Text, TipoDto, ImpDto)
                Cantidad = ComprobarCero(txtAux3(6).Text)
                If Cantidad <> "0" Then
                    ImpDto = CalcularImporteDto(txtAux3(6).Text, CStr(CCur(ImporteSinFormato(txtAux3(10).Text)) / CCur(Cantidad)), txtAux3(0).Text, txtAux3(1).Text, txtAux3(2).Text, CStr(DBLet(Data1.Recordset!ImpDtoc, "N")), False)
                    txtAux3(11).Text = CalcularImporteFClien(txtAux3(6).Text, CStr(CCur(ImporteSinFormato(txtAux3(10).Text)) / CCur(Cantidad)), txtAux3(12).Text, txtAux3(13).Text, TipoDto, ImpDto, txtAux3(10).Text)
                Else
                    ImpDto = CalcularImporteDto(txtAux3(6).Text, CStr(0), txtAux3(0).Text, txtAux3(1).Text, txtAux3(2).Text, CStr(DBLet(Data1.Recordset!ImpDtoc, "N")), False)
                    txtAux3(11).Text = CalcularImporteFClien(txtAux3(6).Text, CStr(0), txtAux3(12).Text, txtAux3(13).Text, TipoDto, ImpDto, txtAux3(10).Text)
                End If
                PonerFormatoDecimal txtAux3(11), 1
                
                'precio neto
                If ComprobarCero(txtAux3(6).Text) <> "0" Then
                    txtAux3(9).Text = Round2(CCur(ImporteSinFormato(txtAux3(11).Text)) / CCur(ImporteSinFormato(txtAux3(6).Text)), CCur(campo2))
                Else
                    txtAux3(9).Text = "0"
                End If
                PonerFormatoDecimal txtAux3(9), 7
            
            Case Else
            
        End Select
        
    End If
    
End Sub

Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim i As Byte
Dim Cad As String

    Combo1(0).Clear
    
    Combo1(0).AddItem "Normal"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    
    Combo1(0).AddItem "Exento"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    
    Combo1(0).AddItem "Recargo Equiv."
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    
    'Tipo de divisa
    Combo1(1).Clear
    
    Cad = "SELECT * FROM moneda ORDER BY codmoneda"
    Set Rs = New ADODB.Recordset
'    Rs.Open Cad, conn, OpenForwardOnly, adLockPessimistic, adCmdText
    Rs.Open Cad, conn, adOpenDynamic, adLockReadOnly, adCmdText
    
    While Not Rs.EOF
        Combo1(1).AddItem Rs!nommoneda
        Combo1(1).ItemData(Combo1(1).NewIndex) = Rs!codmoneda
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
End Sub

Private Function ModificarFactura(MenError As String) As Boolean
Dim Sql As String
Dim Sql2 As String
    
    
    On Error GoTo eModificarFactura
    
    ModificarFactura = False
    
    'Como hay actualizacion en cascada no hace falta modificar las fechas en las tablas de las lineas
    
    Sql = "update facturas set codforpa = " & DBSet(Text1(4).Text, "N")
    Sql = Sql & " ,dtocom1 = " & DBSet(Text1(7).Text, "N")
    Sql = Sql & " ,dtocom2 = " & DBSet(Text1(8).Text, "N")
    Sql = Sql & " ,impdtoc = " & DBSet(Text1(5).Text, "N")
    Sql = Sql & " ,observac = " & DBSet(Text1(2).Text, "T")
    Sql = Sql & " ,fecfactu = " & DBSet(Text1(1).Text, "F")
    Sql = Sql & ", pasedicom = " & Check1(2).Value
    '[Monica]27/01/2015: faltaba que actualizara el codclien
    Sql = Sql & ", codclien = " & DBSet(Text1(3).Text, "N")
    
    '[Monica]10/09/2018: actualizamos intconta
    Sql = Sql & ", intconta = " & Check1(1).Value
    
    
    
    Sql = Sql & " where codtipom = " & DBSet(Text1(6).Text, "T")
    Sql = Sql & " and numfactu = " & DBSet(Text1(0).Text, "N")
    Sql = Sql & " and fecfactu = " & DBSet(FechaAnt, "F")
    
    conn.Execute Sql

    ModificarFactura = True
    Exit Function

eModificarFactura:
    MenError = MenError & vbCrLf & Err.Description
End Function


Private Function ModificaCabecera() As Boolean
Dim b As Boolean
Dim MenError As String
Dim vFactura As CFactura
Dim CADENA As String

    On Error GoTo EModificarCab

    conn.BeginTrans
    
    '[Monica]28/08/2013: Dejamos modificar los datos de la factura aunque esté contabilizada
    Set LOG = Nothing
    
    CADENA = Text1(6).Text & " " & Text1(0).Text & " " & Text1(1).Text & " de " & Text1(25).Text & " "
    If Text1(1).Text <> FechaAnt Then CADENA = CADENA & FechaAnt & " por " & Text1(1).Text
    If ClienteAnt <> Text1(3).Text Then CADENA = CADENA & " Cli " & ClienteAnt & " por " & Text1(3).Text
    If FPagoAnt <> Text1(4).Text Then CADENA = CADENA & " FPago " & FPagoAnt & " por " & Text1(4).Text
    If Dto1Ant <> Text1(7).Text Then CADENA = CADENA & " Dto1 " & Dto1Ant & " por " & Text1(7).Text
    If Dto2Ant <> Text1(8).Text Then CADENA = CADENA & " Dto2 " & Dto2Ant & " por " & Text1(8).Text
    If IDtoAnt <> Text1(5).Text Then CADENA = CADENA & " IDto " & IDtoAnt & " por " & Text1(5).Text
    If ObsAnt <> Text1(2).Text Then CADENA = CADENA & " Obs " & Trim(ObsAnt) & " por " & Trim(Text1(2).Text)
    If EdiAnt <> Check1(2).Value Then CADENA = CADENA & " Edi " & EdiAnt & " por " & Check1(2).Value
    If ContaAnt <> Check1(1).Value Then CADENA = CADENA & " IntConta " & ContaAnt & " por " & Check1(1).Value
    

    If CADENA <> "" And ((Text1(1).Text <> FechaAnt) Or Check1(1).Value = 1) Then
        '-----------------------------------------------------------------------------
        '  LOG de acciones.
        Set LOG = New cLOG
        LOG.Insertar 9, vUsu, "Modificación Cabecera: " & CADENA & vbCrLf
        Set LOG = Nothing
        '-----------------------------------------------------------------------------
        
        MenError = "Modificando Cabecera Factura"
        b = ModificarFactura(MenError)
    Else
        b = ModificaDesdeFormulario2(Me, 2, "Frame2")
    End If
    
'    '[Monica]20/09/2012
'    Set LOG = Nothing
'    If Text1(1).Text <> FechaAnt Then
'        '------------------------------------------------------------------------------
'        '  LOG de acciones.
'        Set LOG = New cLOG
'        'campo = "Facturas de Clientes: " & """"
'        LOG.Insertar 9, vUsu, "FecAnt: " & FechaAnt & " Nueva: " & Text1(1).Text
'        Set LOG = Nothing
'          '-----------------------------------------------------------------------------
'
'        MenError = "Modificando Cabecera Factura"
'        b = ModificarFactura(MenError)
'    Else
'        b = ModificaDesdeFormulario2(Me, 2, "Frame2")
'    End If
    
    
    MenError = "Recalcular Dtos calibres"
    If b Then b = RecalcularDtosLineas(Text1(6).Text, Text1(0).Text, Text1(1).Text, MenError)
    
    MenError = "Actualizar Variedades"
    If b Then b = ActualizarVariedades(Text1(6).Text, Text1(0).Text, Text1(1).Text, MenError)
    
    CalcularDatosFactura

    
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
    
    CodTipoMov = Text1(6).Text
    
    '[Monica]05/10/2018: si el tipo de factura es de B la dejamos como contabilizada,
    '                    podremos modificarla
    If CodTipoMov = TipoFactB Then Check1(1).Value = 1
    
    
    If TipoFactura = 0 Then
        Set vTipoMov = New CTiposMov
        If vTipoMov.Leer(CodTipoMov) Then
            Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
            Sql = CadenaInsertarDesdeForm(Me)
            If Sql <> "" Then
                If InsertarOferta(Sql, vTipoMov) Then
                    CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                    PonerCadenaBusqueda
                    PonerModo 2
                    'Ponerse en Modo Insertar Lineas
    '                BotonMtoLineas 0, "Variedades"
    
                    If CodTipoMov = "FAR" Then
                        
                    Else
                        DescontarFacturasACuenta Text1(6).Text, Text1(0).Text, Text1(1).Text, Text1(3).Text
                    End If
                    
                    BotonAnyadirLinea 0
                End If
            End If
            Text1(0).Text = Format(Text1(0).Text, "0000000")
        End If
        Set vTipoMov = Nothing
    Else
        Sql = CadenaInsertarDesdeForm(Me)
        conn.Execute Sql

        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
        PonerCadenaBusqueda
        PonerModo 2
        'Ponerse en Modo Insertar Lineas
'                BotonMtoLineas 0, "Variedades"

        DescontarFacturasACuenta Text1(6), Text1(0).Text, Text1(1).Text, Text1(3).Text
        
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
Dim nomFrame As String
Dim b As Boolean
Dim Mens As String
Dim CADENA As String


    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case Index
        Case 0: nomFrame = "FrameAux0" 'variedades
        Case 1: nomFrame = "FrameAux1" 'envases
        Case 2: nomFrame = "FrameAux2" 'facturas_acuenta
    End Select
    ' ***************************************************************
    
    
    
    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        
        Select Case Index
            Case 0
                '[Monica]22/02/2012: insertamos la transaccion
                conn.BeginTrans
                If InsertarDesdeForm2(Me, 2, nomFrame) Then
                    
                    '[Monica]28/08/2013: insertamos en el log que hemos insertado una linea en una factura contabilizada
                    If Check1(1).Value = 1 Then
                        '------------------------------------------------------------------------------
                        '  LOG de acciones.
                        Set LOG = New cLOG
                        
                        CADENA = "Inserta Linea de Variedades "
                        CADENA = CADENA & Text1(6).Text & " " & Text1(0).Text & " " & Text1(1).Text & " de " & Text1(25).Text & " Linea " & txtAux3(3).Text
                        CADENA = CADENA & " Albaran " & txtAux3(4).Text & " " & txtAux3(5).Text & " de " & txtAux3(11).Text
                        
                        LOG.Insertar 9, vUsu, CADENA
                        Set LOG = Nothing
                          '-----------------------------------------------------------------------------
                    End If
                    
                    ' *** si n'hi ha que fer alguna cosa abas d'insertar
                    Mens = "Recalcular Dtos lineas"
                    b = RecalcularDtos(txtAux3(0).Text, txtAux3(1).Text, txtAux3(2).Text, Mens)
                    
                    '[Monica]22/02/2012: insertamos en facturas_calibre
                    Mens = "Insertar Calibres"
                    If b Then b = InsertarModificarCalibres(True, txtAux3(0).Text, txtAux3(1).Text, txtAux3(2).Text, txtAux3(3).Text, txtAux3(4).Text, txtAux3(5).Text, txtAux3(6).Text, txtAux3(15).Text, txtAux3(10).Text, txtAux3(11).Text, Mens, txtAux3(7).Text)
                    
                    If b Then
                        
                        conn.CommitTrans
                        
                        CalcularDatosFactura
                        ' *************************************************
                        b = BloqueaRegistro("facturas", "codtipom = " & DBSet(Data1.Recordset!codTipoM, "T") & " and numfactu = " & DBSet(Data1.Recordset!NumFactu, "N") & " and fecfactu = " & DBSet(Data1.Recordset!FecFactu, "F"))
                        CargaGrid DataGrid2, AdoAux(0), True
                        If b Then BotonAnyadirLinea NumTabMto
                        SSTab1.Tab = NumTabMto
                    Else
                        conn.RollbackTrans
                    End If
                Else
                    conn.RollbackTrans
                End If
            
            Case 1
                If InsertarLineaEnv(txtAux(3).Text) Then
                    CalcularDatosFactura
                    b = BloqueaRegistro("facturas", "numfactu = " & Data1.Recordset!NumFactu)
                    CargaGrid DataGrid3, AdoAux(1), True
                    If b Then BotonAnyadirLinea NumTabMto
                    SSTab1.Tab = NumTabMto
                End If
            
        End Select
    End If
End Sub

Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vtabla As String
Dim anc As Single
Dim i As Integer
    
    ModificaLineas = 1 'Posem Modo Afegir Llínia
    
    PulsadoF2 = False
    
'    '[Monica]23/03/2012: sólo se puede modificar
'    'solo se puede modificar la factura si no esta contabilizada
'    If FactContabilizada(True) Then
'        TerminaBloquear
'        Exit Sub
'    End If
    
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
        Case 0: vtabla = "facturas_variedad"
        Case 1: vtabla = "facturas_envases"
        Case 2: vtabla = "facturas_acuenta"
    End Select
    ' ********************************************************
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case NumTabMto
        Case 0 ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", vWhere)
            ' ***************************************************************

            AnyadirLinea DataGrid2, AdoAux(0)
    
            anc = DataGrid2.Top
            If DataGrid2.Row < 0 Then
                anc = anc + 240 '210
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
'[Monica]15/07/2011: dejamos cambiar la cantidad
            BloquearTxt txtAux3(6), False
            BloquearTxt txtAux3(7), False
            
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
            
            
            '[Monica]05/10/2018: si es B el tipo de iva es exento
            If Text1(6).Text = TipoFactB Then
                txtAux3(14).Text = vParamAplic.CodIvaExento
            Else
                Select Case Me.Combo1(0).ListIndex
                    Case 0  'normal
                        '++monica:27/05/08 el iva se cogerá de la variedad cuando carguemos datos de albaran
                        txtAux3(14).Text = "" 'vParamAplic.CodIvaNormal
                    Case 1  'exento
                        txtAux3(14).Text = vParamAplic.CodIvaExento
                    Case 2  'recargo de equivalencia
                        txtAux3(14).Text = vParamAplic.CodIvaRecargo
                End Select
            End If
            
            BloquearBtn Me.btnBuscar(1), False
            
'                    BloquearTxt txtaux(12), False
            PonerFoco txtAux3(4)
                    
        ' *** si n'hi han llínies sense datagrid ***
        Case 1
            NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", vWhere)
            ' ***************************************************************

            AnyadirLinea DataGrid3, AdoAux(1)
    
            anc = DataGrid3.Top
            If DataGrid3.Row < 0 Then
                anc = anc + 240 '210
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
        
        ' [Monica] 22/06/2010 : descontamos las facturas de a cuenta que tengamos
        Case 2
             DescontarFacturasACuenta Text1(6).Text, Text1(0).Text, Text1(1).Text, Text1(3).Text
             TerminaBloquear
             CalcularDatosFactura
             CargaGrid Me.DataGrid1, Me.AdoAux(2), True
             If Not BloqueaRegistro("facturas_acuenta", vWhere) Then Exit Sub
    
    End Select
End Sub

Private Sub LimpiarCamposLin(FrameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, FrameAux  'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Function ModificarLinea() As Boolean
'Modifica registre en les taules de Llínies
Dim nomFrame As String
Dim V As Integer
Dim Cad As String
Dim Sql As String
Dim vCStock As CStock
Dim b As Boolean
Dim Mens As String
Dim CADENA As String
    
    On Error GoTo eModificarLinea

    ModificarLinea = False
    Sql = ""

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomFrame = "FrameAux0" 'variedades
        Case 1: nomFrame = "FrameAux1" 'envases
        Case 2: nomFrame = "FrameAux2" 'facturas a cuenta
    End Select
    ' **************************************************************

    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        
        Select Case NumTabMto
        Case 0
            conn.BeginTrans
        
            '[Monica]28/08/2013: insertamos en el log que hemos eliminado una linea en una factura contabilizada
            If Check1(1).Value = 1 Then
                '------------------------------------------------------------------------------
                '  LOG de acciones.
                Set LOG = New cLOG
                
                CADENA = "Modificar Linea de Variedades "
                
                LOG.Insertar 9, vUsu, CADENA & AdoAux(0).Recordset.Fields(0) & " " & AdoAux(0).Recordset.Fields(1) & " " & AdoAux(0).Recordset.Fields(2) & " de " & Text1(25).Text & " Linea " & AdoAux(0).Recordset.Fields(3) & " Alb " & AdoAux(0).Recordset.Fields(4) & " " & AdoAux(0).Recordset.Fields(5) & " de " & AdoAux(0).Recordset.Fields(12) & "-" & txtAux3(11).Text
                Set LOG = Nothing
                  '-----------------------------------------------------------------------------
            End If
        
        
            If ModificaDesdeFormulario2(Me, 2, nomFrame) Then
                ' *** si cal que fer alguna cosa abas d'insertar ***
                ' ******************************************************
' antes
'                Mens = "Recalcular Dtos lineas"
'                b = RecalcularDtos(txtAux3(0).Text, txtAux3(1).Text, txtAux3(2).Text, Mens)

                '[Monica]22/02/2012: modificamos en facturas_calibre
                Mens = "Insertar Calibres"
                b = InsertarModificarCalibres(False, txtAux3(0).Text, txtAux3(1).Text, txtAux3(2).Text, txtAux3(3).Text, txtAux3(4).Text, txtAux3(5).Text, txtAux3(6).Text, txtAux3(15).Text, txtAux3(10).Text, txtAux3(11).Text, Mens, txtAux3(7).Text)
    
                If b Then
                    Mens = "Recalcular Dtos lineas calibres"
                    b = RecalcularDtosLineas(txtAux3(0).Text, txtAux3(1).Text, txtAux3(2).Text, Mens)
                End If
                
                If b Then
                    Mens = "Actualizar Variedades"
                    b = ActualizarVariedades(txtAux3(0).Text, txtAux3(1).Text, txtAux3(2).Text, Mens)
                End If
    
    
    '            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
                If b Then
                
                    conn.CommitTrans
                    
                    V = AdoAux(0).Recordset.Fields(3) 'el 2 es el nº de llinia
                
                    CalcularDatosFactura
                    ModificaLineas = 0
        
'                    V = AdoAux(0).Recordset.Fields(3) 'el 2 es el nº de llinia
                    CargaGrid DataGrid2, AdoAux(0), True
    
                    ' *** si n'hi han tabs ***
                    SSTab1.Tab = 0
    
                    DataGrid2.SetFocus
                    AdoAux(0).Recordset.Find (AdoAux(0).Recordset.Fields(3).Name & " =" & V)
    
                    LLamaLineas ModificaLineas, 0, "DataGrid2"
               End If
               
            End If
    
        Case 1
            Set vCStock = New CStock
            If Not InicializarCStock(vCStock, "S") Then Exit Function
            
            If DatosOkLineaEnv(vCStock) Then
                '#### LAURA 15/11/2006
                conn.BeginTrans
                
                '[Monica]28/08/2013: insertamos en el log que hemos eliminado una linea en una factura contabilizada
                If Check1(1).Value = 1 Then
                    '------------------------------------------------------------------------------
                    '  LOG de acciones.
                    Set LOG = New cLOG
                    
                    CADENA = "Modificar Linea de Envases "
                    LOG.Insertar 9, vUsu, CADENA & AdoAux(1).Recordset.Fields(0) & " " & AdoAux(1).Recordset.Fields(1) & " " & AdoAux(1).Recordset.Fields(2) & " de " & Text1(25).Text & " Linea " & AdoAux(1).Recordset.Fields(3) & " " & AdoAux(1).Recordset.Fields(5) & " de " & AdoAux(1).Recordset.Fields(10) & "-" & txtAux(9).Text
                    Set LOG = Nothing
                    '-----------------------------------------------------------------------------
                End If
    
        '        Set vCStock = New CStock
                'iniciamos la clase con los valores anteriores para deshacer lo q insertamos antes
                b = InicializarCStock(vCStock, "E")
                If b Then
                    b = vCStock.DevolverStock 'eliminamos de smoval y devolvemos stock valores anteriores
                    'ahora leemos los valores nuevos
                    If b Then b = InicializarCStock(vCStock, "S")
                    'insertamos en smoval y actualizamos stock a los valores nuevos
                    vCStock.Cantidad = CSng(ComprobarCero(txtAux(6).Text))
                    '[Monica]22/05/2012: indicamos que es ticket para que no salga que no hay stock, añadido (true)
                    If b Then b = vCStock.ActualizarStock(True)
            
                    'actualizar la linea de Albaran
                    If b Then
                        Sql = "UPDATE facturas_envases Set codalmac = " & txtAux(4).Text & ", codartic=" & DBSet(txtAux(5).Text, "T") & ", "
                        Sql = Sql & "ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
                        Sql = Sql & "cantidad= " & DBSet(txtAux(6).Text, "N") & ", "
                        Sql = Sql & "precioar= " & DBSet(txtAux(7).Text, "N") & ", " 'precio
                        Sql = Sql & "dtolinea= " & DBSet(txtAux(8).Text, "N") & ", "
                        Sql = Sql & "importel= " & DBSet(txtAux(9).Text, "N") & ", " 'Importe
                        '[Monica]19/12/2016: faltaba modificar el numero de albarán
                        Sql = Sql & "numalbar= " & DBSet(txtAux(11).Text, "N") & ", " ' numero de albaran
                        
                        Sql = Sql & "codigiva= " & DBSet(txtAux(10).Text, "N") & " " 'codigo de iva
                        Sql = Sql & Replace(ObtenerWhereCP(True), NombreTabla, "facturas_envases") & " AND numlinea=" & AdoAux(1).Recordset!NumLinea
                        conn.Execute Sql
                    End If
                End If
                
                If b Then conn.CommitTrans
            End If
            Set vCStock = Nothing
                
            ModificaLineas = 0
            
            V = AdoAux(1).Recordset.Fields(3) 'el 2 es el nº de llinia
            
            CalcularDatosFactura
            
            CargaGrid DataGrid3, AdoAux(1), True

            ' *** si n'hi han tabs ***
            SSTab1.Tab = 1

            DataGrid3.SetFocus
            AdoAux(1).Recordset.Find (AdoAux(1).Recordset.Fields(3).Name & " =" & V)

            LLamaLineas ModificaLineas, 0, "DataGrid3"
        
        End Select
    
    End If
        
eModificarLinea:
    If Err.Number <> 0 Or Not b Then
        MuestraError Err.Number, "Modificar Lineas Factura" & vbCrLf & Err.Description & vbCrLf & Mens
        b = False
    End If
    
    If b Then
'        conn.CommitTrans
        ModificarLinea = True
    Else
        conn.RollbackTrans
        ModificarLinea = False
    End If
End Function
        

Private Function DatosOkLlin(nomFrame As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim b As Boolean
Dim Cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte
Dim Cliente As String
Dim TipoDto As Byte

    DatosOkLlin = True
    
    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False


    b = CompForm2(Me, 2, nomFrame) 'Comprovar formato datos ok
    If Not b Then Exit Function
    
    'en variedades comprobamos que el albaran introducido corresponde al cliente
   Select Case nomFrame
        Case "FrameAux0"
            Cliente = ""
            Cliente = DevuelveDesdeBDNew(cAgro, "albaran", "codclien", "numalbar", txtAux3(4).Text, "N")
            
            If CLng(Cliente) <> CLng(Data1.Recordset!CodClien) Then
                MsgBox "El albarán introducido no es del cliente del la factura. Revise.", vbExclamation
                b = False
            End If
            
'        '++
'        '[Monica]15/02/2011: Problema con el Alt+A
'            TxtAux3_LostFocus (8)
'            TxtAux3_LostFocus (10)
'        '++
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
Private Sub PonerDatosCliente(CodClien As String, Optional nifClien As String)
Dim vCliente As CCliente
Dim Observaciones As String
    
    On Error GoTo EPonerDatos
    
    If CodClien = "" Then
        LimpiarDatosCliente
        Exit Sub
    End If

    Set vCliente = New CCliente
    
    'si se ha modificado el cliente volver a cargar los datos
    If vCliente.Existe(CodClien) Then
        If vCliente.LeerDatos(CodClien) Then
            Text1(3).Text = vCliente.Codigo
            FormateaCampo Text1(3)
            If (Modo = 3) Or (Modo = 4) Then
                Text2(3).Text = vCliente.Nombre  'Nom clien
                Text1(4).Text = vCliente.ForPago
                Text2(4).Text = PonerNombreDeCod(Text1(4), "forpago", "nomforpa")
                Text1(7).Text = Format(vCliente.Dto1, FormatoDescuento)
                Text1(8).Text = Format(vCliente.Dto2, FormatoDescuento)
                Me.Combo1(0).ListIndex = vCliente.TipoIva
                
                
                If Me.chkRectifica.Value = 0 Then
                    TipoFactura = vCliente.TipoFactu
                    If TipoFactura = 1 Then
                        '[Monica]30/07/2012: traemos el tipo de movimiento del cliente
                        If vCliente.tipoMov = "" Then
                            Text1(6).Text = "FAV"
                        Else
                            Text1(6).Text = vCliente.tipoMov '"FAV"
                        End If
                    Else
                        Text1(6).Text = vCliente.tipoMov
                    End If
                    
                    '[Monica]29/12/2017: si es bio
                    If Me.chkBio.Value = 1 Then
                        Text1(6).Text = "FBI"
                        TipoFactura = 0
                    End If
                    
                    
                Else
                    TipoFactura = 0
                End If
                
                
            End If

            Observaciones = DBLet(vCliente.Observaciones)
            If Observaciones <> "" Then
                MsgBox Observaciones, vbInformation, "Observaciones del cliente"
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
Dim Sql As String
Dim vPrecio As Currency

    If Albaran = "" Or Linea = "" Then
        txtAux3(6).Text = ""
        txtAux3(7).Text = ""
        txtAux3(15).Text = ""
        For i = 0 To Text3.Count - 1
            Text3(i).Text = ""
        Next i
        
        '[Monica]23/04/2013: limpiamos el precio bruto
        txtAux3(8).Text = ""
    Else
        Sql = "select albaran.fechaalb, albaran.matrirem, forfaits.kiloscaj, albaran_variedad.pesoneto, "
        Sql = Sql & " albaran_variedad.numcajas, variedades.nomvarie, destinos.nomdesti, forfaits.nomconfe, albaran_variedad.unidades, albaran_variedad.codvarie "
        '[Monica]23/04/2013: traemos el precio provisional o definitivo si lo tiene
        Sql = Sql & " ,albaran_variedad.preciopro, albaran_variedad.preciodef "
        Sql = Sql & " from albaran, albaran_variedad, variedades, destinos, forfaits "
        Sql = Sql & " where albaran_variedad.numalbar = " & DBSet(Albaran, "N")
        Sql = Sql & " and albaran_variedad.numlinea = " & DBSet(Linea, "N")
        Sql = Sql & " and albaran.numalbar = albaran_variedad.numalbar "
        Sql = Sql & " and albaran_variedad.codforfait = forfaitS.codforfait "
        Sql = Sql & " and albaran_variedad.codvarie = variedades.codvarie "
        Sql = Sql & " and albaran.codclien = destinos.codclien "
        Sql = Sql & " and albaran.coddesti = destinos.coddesti "
        
        Set Rs = New ADODB.Recordset
        
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs.EOF Then
            '[Monica]10/06/2013: en las facturas rectificativas la cantidad en negativo y el precio en posivivo, se lo damos por defecto
            If Text1(6).Text = "FAR" Then
                txtAux3(6).Text = DBLet(Rs.Fields(3).Value * (-1), "N")
                '[Monica]15/07/2011: añadido el round en la siguiente linea
                txtAux3(7).Text = Round2(DBLet(Rs.Fields(4).Value, "N") * DBLet(Rs.Fields(2).Value * (-1), "N"), 0)
                txtAux3(15).Text = DBLet(Rs.Fields(8).Value * (-1), "N")
            Else
                txtAux3(6).Text = DBLet(Rs.Fields(3).Value, "N")
                '[Monica]15/07/2011: añadido el round en la siguiente linea
                txtAux3(7).Text = Round2(DBLet(Rs.Fields(4).Value, "N") * DBLet(Rs.Fields(2).Value, "N"), 0)
                txtAux3(15).Text = DBLet(Rs.Fields(8).Value, "N")
            End If
            
            '++monica:27/05/08: metemos el codigo iva de la variedad si el cliente es normal
            If Me.Combo1(0).ListIndex = 0 Then
                txtAux3(14).Text = DevuelveDesdeBDNew(cAgro, "variedades", "codigiva", "codvarie", Rs.Fields(9).Value, "N")
                
                '[Monica]15/10/2018: tipo de iva exento
                If Text1(6).Text = TipoFactB Then
                    txtAux3(14).Text = vParamAplic.CodIvaExento
                End If
                
            End If
            '++
            
            '[Monica]23/04/2013: en precio bruto traigo el precio provisional o definitivo dependiendo
            If DBLet(Rs.Fields(11).Value, "N") <> 0 Then
                vPrecio = DBLet(Rs.Fields(11).Value, "N")
            Else
                vPrecio = DBLet(Rs.Fields(10).Value, "N")
            End If
            
            Text3(0).Text = DBLet(Rs.Fields(0).Value, "F")
            Text3(1).Text = DBLet(Rs.Fields(1).Value, "T")
            Text3(2).Text = DBLet(Rs.Fields(6).Value, "T")
            Text3(3).Text = DBLet(Rs.Fields(5).Value, "T")
            Text3(4).Text = DBLet(Rs.Fields(7).Value, "T")
        
            '[Monica]23/04/2013: recalcula
            txtAux3(8).Text = vPrecio
        
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



Private Function InsertarLineaEnv(NumLinea As String) As Boolean
'Inserta un registro en la tabla de lineas de Albaranes: slialb
Dim Sql As String
Dim vWhere As String
Dim b As Boolean
Dim vCStock As CStock
Dim DentroTRANS As Boolean
Dim CADENA As String

    InsertarLineaEnv = False
    Sql = ""
    DentroTRANS = False
    
    'Conseguir el siguiente numero de linea
    vWhere = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
'    numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
    
    Set vCStock = New CStock
    If Not InicializarCStock(vCStock, "S", NumLinea) Then Exit Function
    
    If DatosOkLineaEnv(vCStock) Then 'Lineas de factura
        'Inserta en tabla "facturas_envases"
        Sql = "INSERT INTO facturas_envases "
        Sql = Sql & "(codtipom,numfactu,fecfactu,numlinea,codalmac,codartic,cantidad,precioar,dtolinea,importel,ampliaci,codigiva,numalbar,numlinealbar,fecalbar) "
        Sql = Sql & "VALUES ('" & txtAux(0).Text & "', " & DBSet(txtAux(1).Text, "N") & ", " & DBSet(txtAux(2).Text, "F") & ", " & NumLinea & ", " & DBSet(txtAux(4).Text, "N") & ","
        Sql = Sql & DBSet(txtAux(5).Text, "T") & ", "
        Sql = Sql & DBSet(txtAux(6).Text, "N") & ", "
        Sql = Sql & DBSet(txtAux(7).Text, "N") & ", " & DBSet(txtAux(8).Text, "N") & ", "
        Sql = Sql & DBSet(txtAux(9).Text, "N") & ","
        Sql = Sql & DBSet(Text2(16).Text, "T") & ","
        Sql = Sql & DBSet(txtAux(10).Text, "N") & ","
        Sql = Sql & DBSet(txtAux(11).Text, "N") & ","
        Sql = Sql & DBSet(txtAux(12).Text, "N") & ","
        Sql = Sql & DBSet(txtAux(13).Text, "F") & ")"
     Else
        Exit Function
     End If
    
    If Sql <> "" Then
        On Error GoTo EInsertarLineaEnv
        conn.BeginTrans
        DentroTRANS = True
        
        'insertar la linea
        conn.Execute Sql
        
        
        '[Monica]28/08/2013: insertamos en el log que hemos insertado una linea en una factura contabilizada
        If Check1(1).Value = 1 Then
            '------------------------------------------------------------------------------
            '  LOG de acciones.
            Set LOG = New cLOG
            
            CADENA = "Inserta Linea de Envases "
            CADENA = CADENA & Text1(6).Text & " " & Text1(0).Text & " " & Text1(1).Text & " de " & Text1(25).Text & " Linea " & txtAux(3).Text
            CADENA = CADENA & " " & txtAux(5).Text

            LOG.Insertar 9, vUsu, CADENA
            Set LOG = Nothing
              '-----------------------------------------------------------------------------
        End If
        
        
        'si hay control de stock para el articulo actualizar en salmac e insertar en smoval
        'en actualizar stock comprobamos si el articulo tiene control de stock
        
        '[Monica]22/05/2012: indicamos que es ticket para que no salga que no hay stock, añadido (true)
        b = vCStock.ActualizarStock(True)
        
    
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
'        'Inserta en tabla "facturas_envases"
'        Sql = "INSERT INTO facturas_acuenta "
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



Private Function InicializarCStock(ByRef vCStock As CStock, TipoM As String, Optional NumLinea As String) As Boolean
   On Error Resume Next
   'On Error GoTo eInicializar

    vCStock.tipoMov = TipoM
    '[Monica]20/03/2012: guardamos el numero de albaran o de factura (dependiendo de de donde viene)
    If ComprobarCero(txtAux(11).Text) = 0 Or ComprobarCero(txtAux(12).Text) = 0 Or Not EsFacturaEnvases(Text1(6).Text, Text1(0).Text, Text1(1).Text) Then
        vCStock.DetaMov = Text1(6).Text
    Else
        vCStock.DetaMov = vParamAplic.CodTipomAlb '"ALV"
    End If
    
    vCStock.Trabajador = CInt(Text1(3).Text) 'guardamos el cliente de la factura
    
    '[Monica]20/03/2012: guardamos el numero de albaran o de factura (dependiendo de de donde viene)
    If ComprobarCero(txtAux(11).Text) = 0 Or ComprobarCero(txtAux(12).Text) = 0 Or Not EsFacturaEnvases(Text1(6).Text, Text1(0).Text, Text1(1).Text) Then
        vCStock.Documento = CLng(Text1(0).Text) 'Nº Factura
        vCStock.Fechamov = Text1(1).Text 'Fecha de la Factura
    Else
        vCStock.Documento = CLng(txtAux(11).Text) 'Nº albaran
        vCStock.Fechamov = txtAux(13).Text 'Fecha del albaran
    End If
    
    
    '1=Insertar, 2=Modificar
    If ModificaLineas = 1 Or (ModificaLineas = 2 And TipoM = "S") Then
        vCStock.codArtic = txtAux(5).Text
        vCStock.codAlmac = CInt(txtAux(4).Text)
        If ModificaLineas = 1 Then '1=Insertar
            vCStock.Cantidad = CSng(ComprobarCero(txtAux(6).Text))
        Else '2=Modificar(Debe haber en stock la diferencia)
            If AdoAux(1).Recordset!codArtic = txtAux(5).Text Then
                vCStock.Cantidad = CSng(ComprobarCero(txtAux(6).Text)) - AdoAux(1).Recordset!Cantidad
            Else
                vCStock.Cantidad = CSng(ComprobarCero(txtAux(6).Text))
            End If
        End If
        vCStock.Importe = CCur(ComprobarCero(txtAux(9).Text))
    Else
        vCStock.codArtic = AdoAux(1).Recordset!codArtic
        vCStock.codAlmac = CInt(AdoAux(1).Recordset!codAlmac)
        vCStock.Cantidad = CSng(AdoAux(1).Recordset!Cantidad)
        vCStock.Importe = CCur(AdoAux(1).Recordset!ImporteL)
    End If
    If ModificaLineas = 1 Then
        vCStock.LineaDocu = CInt(ComprobarCero(NumLinea))
    Else
        '[Monica]20/03/2012: dependiendo de si viene de factura o de un albaran de envase (añado la condicion)
        If ComprobarCero(txtAux(11).Text) = 0 Or ComprobarCero(txtAux(12).Text) = 0 Or Not EsFacturaEnvases(Text1(6).Text, Text1(0).Text, Text1(1).Text) Then
            vCStock.LineaDocu = CInt(AdoAux(1).Recordset!NumLinea)
        Else
            vCStock.LineaDocu = CInt(AdoAux(1).Recordset!numlinealbar)
        End If
    End If
    If Err.Number <> 0 Then
        MsgBox "No se han podido inicializar la clase para actualizar Stock", vbExclamation
        InicializarCStock = False
    Else
        InicializarCStock = True
    End If

'eInicializar:
'    MuestraError Err.Number, "inicializar", Err.Description
End Function

Private Function DatosOkLineaEnv(ByRef vCStock As CStock) As Boolean
Dim b As Boolean
Dim i As Byte
    
    On Error GoTo EDatosOkLineaEnv

    DatosOkLineaEnv = False
    b = True

    'Comprobar si se mueve mover stock (hay stock, o si no hay pero no control de stock)
'[Monica]22/05/2012: no saca cartel de que no hay stock
'    If vCStock.MueveStock Then
'        b = vCStock.MoverStock
'    End If
    DatosOkLineaEnv = b
    
EDatosOkLineaEnv:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function



Private Function EliminarStock() As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim vCStock As CStock
Dim b As Boolean

    On Error GoTo eEliminarStock
    
    Sql = "select * from facturas_envases where " & Replace(ObtenerWhereCP(False), "facturas", "facturas_envases")
    Set Rs = New ADODB.Recordset
    
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    b = True
    While Not Rs.EOF And b
        Set vCStock = New CStock
        
        vCStock.Cantidad = DBLet(Rs!Cantidad, "N")
        vCStock.codAlmac = DBLet(Rs!codAlmac, "N")
        vCStock.codArtic = DBLet(Rs!codArtic, "T")
        '[Monica]20/03/2012: lo que se guardó es el numero de albaran
        If DBLet(Rs!NumAlbar, "N") <> 0 And EsFacturaEnvases(Text1(6).Text, Text1(0).Text, Text1(1).Text) Then
            vCStock.Documento = CLng(DBLet(Rs!NumAlbar, "N"))
            vCStock.DetaMov = "ALV"
            vCStock.LineaDocu = DBLet(Rs!numlinealbar, "N")
            vCStock.Fechamov = DBLet(Rs!FecAlbar, "F")
        Else
            vCStock.Documento = CLng(DBLet(Rs!NumFactu, "N"))
            vCStock.DetaMov = DBLet(Rs!codTipoM, "T")
            vCStock.Fechamov = DBLet(Rs!FecFactu, "F")
            vCStock.LineaDocu = DBLet(Rs!NumLinea, "N")
        End If
        vCStock.Importe = DBLet(Rs!ImporteL, "N")
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
Dim cadwhere As String, Sql As String

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For i = 9 To 31
         Text1(i).Text = ""
    Next i
    
    'Comprobar que hay lineas de facturas_variedad para calcular totales
    cadwhere = ObtenerWhereCP(False)
    Sql = "Select count(*) from " & NomTablaLineas & " Where " & Replace(cadwhere, NombreTabla, NomTablaLineas)
    If RegistrosAListar(Sql) = 0 Then
        'Comprobar que hay lineas de facturas_envases para calcular totales
        Sql = "Select count(*) from facturas_envases Where " & Replace(cadwhere, NombreTabla, "facturas_envases")
        If RegistrosAListar(Sql) = 0 Then
            '[Monica]22/06/2010 añadido por facturas_acuenta -- antes sólo: If RegistrosAListar(Sql) = 0 Then exit sub
            'Comprobar que hay lineas de facturas_acuenta para calcular totales
            Sql = "Select count(*) from facturas_acuenta Where " & Replace(cadwhere, NombreTabla, "facturas_acuenta")
            '[Monica]23/08/2013 he quitado el exit sub
            If RegistrosAListar(Sql) = 0 Then 'Exit Sub
            End If
        Else
'            Exit Sub
        End If
    End If
    
    
    If CalcularDatosFacturaVenta(cadwhere, NombreTabla, NomTablaLineas) Then
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
Private Function CalcularDatosFacturaVenta(cadwhere As String, NomTabla As String, NomTablaLin As String) As Boolean
'cadWhere: cad para la where de la SQL que selecciona las lineas del albaran o la factura
'nomTabla: nombre de la tabla de albaranes(scaalp) o de AlbaranesXFactura(scafpa)
'           segun llamemos desde recepcion de facturas o desde Hco de Facturas
Dim Rs As ADODB.Recordset
Dim i As Integer

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
    cadwhere1 = Replace(cadwhere, "facturas", "facturas_variedad")
    Sql = "SELECT facturas_variedad.codigiva, sum(imporbru) as bruto, sum(impornet) as neto"
    Sql = Sql & " FROM facturas_variedad "
    Sql = Sql & " WHERE " & cadwhere1
    Sql = Sql & " GROUP BY 1 "
    Sql = Sql & " UNION "
    cadwhere1 = Replace(cadwhere, "facturas", "facturas_envases")
    Sql = Sql & "SELECT facturas_envases.codigiva, sum(importel) as bruto, sum(importel) as neto"
    Sql = Sql & " FROM facturas_envases "
    Sql = Sql & " WHERE " & cadwhere1
    Sql = Sql & " GROUP BY 1 "
    Sql = Sql & " UNION "
    cadwhere1 = Replace(cadwhere, "facturas", "facturas_acuenta")
    Sql = Sql & "SELECT facturas.codiiva1 as codigiva, sum(brutofac * (-1)) as bruto, sum(brutofac * (-1)) as neto"
    Sql = Sql & " FROM facturas_acuenta, facturas "
    Sql = Sql & " WHERE " & cadwhere1
    Sql = Sql & " and facturas.codtipom = facturas_acuenta.codtipomcta "
    Sql = Sql & " and facturas.numfactu = facturas_acuenta.numfactucta "
    Sql = Sql & " and facturas.fecfactu = facturas_acuenta.fecfactucta "
    Sql = Sql & " GROUP BY 1 "
    Sql = Sql & " ORDER BY 1 "

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    TotBruto = 0
    TotNeto = 0
    TotImpIVA = 0
    vBruto = 0
    vNeto = 0
    i = 1

    '[Monica]23/08/2013: he metido en el if la instruccion de ivaant
    If Not Rs.EOF Then
        Rs.MoveFirst
        IvaAnt = Rs.Fields(0).Value
        
    Else
        '[Monica]23/08/2013 añado el else
        
        Sql = "update facturas "
        Sql = Sql & "set baseimp1 = " & ValorNulo
        Sql = Sql & ",impoiva1 = " & ValorNulo
        Sql = Sql & ",imporec1 = " & ValorNulo
        Sql = Sql & ",porciva1 = " & ValorNulo
        Sql = Sql & ",porcrec1 = " & ValorNulo
        Sql = Sql & ",codiiva1 = " & ValorNulo
        Nulo2 = "N"
        Nulo3 = "N"
        If DBSet(TipoIVA2, "N", "S") = ValorNulo Then Nulo2 = "S"
        If DBSet(TipoIVA3, "N", "S") = ValorNulo Then Nulo3 = "S"
        Sql = Sql & ",baseimp2 = " & ValorNulo
        Sql = Sql & ",impoiva2 = " & ValorNulo
        Sql = Sql & ",imporec2 = " & ValorNulo
        Sql = Sql & ",porciva2 = " & ValorNulo
        Sql = Sql & ",porcrec2 = " & ValorNulo
        Sql = Sql & ",codiiva2 = " & ValorNulo
        Sql = Sql & ",baseimp3 = " & ValorNulo
        Sql = Sql & ",impoiva3 = " & ValorNulo
        Sql = Sql & ",imporec3 = " & ValorNulo
        Sql = Sql & ",porciva3 = " & ValorNulo
        Sql = Sql & ",porcrec3 = " & ValorNulo
        Sql = Sql & ",codiiva3 = " & ValorNulo
        Sql = Sql & ",brutofac = " & ValorNulo
        Sql = Sql & ",impordto = " & ValorNulo
        Sql = Sql & ",totalfac = " & ValorNulo
        Sql = Sql & " where " & cadwhere
        
        conn.Execute Sql
    
        CalcularDatosFacturaVenta = True
        Exit Function
        
    End If
    While Not Rs.EOF
                                        '[Monica]05/05/2015: añadimos la condicion de que la suma de brutos de la factura sea distinta de
                                        '                    para que no llegue como base a la contabilidad
        If IvaAnt <> Rs.Fields(0).Value And vBruto <> 0 Then
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
    '[Monica]23/08/2013: añado el comprobar cero
    impiva = CalcularPorcentaje(ImpBImIVA, CCur(ComprobarCero(cadAux)), 2)
    
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

    'ACTUALIZAMOS LA FACTURA (tabla facturas)
    Sql = "update facturas "
    Sql = Sql & "set baseimp1 = " & DBSet(BaseIVA1, "N")
    Sql = Sql & ",impoiva1 = " & DBSet(ImpIVA1, "N")
    Sql = Sql & ",imporec1 = " & DBSet(ImpREC1, "N")
    Sql = Sql & ",porciva1 = " & DBSet(PorceIVA1, "N")
    Sql = Sql & ",porcrec1 = " & DBSet(PorceREC1, "N")
    Sql = Sql & ",codiiva1 = " & DBSet(TipoIVA1, "N")
    Nulo2 = "N"
    Nulo3 = "N"
    If DBSet(TipoIVA2, "N", "S") = ValorNulo Then Nulo2 = "S"
    If DBSet(TipoIVA3, "N", "S") = ValorNulo Then Nulo3 = "S"
    Sql = Sql & ",baseimp2 = " & DBSet(BaseIVA2, "N", Nulo2)
    Sql = Sql & ",impoiva2 = " & DBSet(ImpIVA2, "N", Nulo2)
    Sql = Sql & ",imporec2 = " & DBSet(ImpREC2, "N", Nulo2)
    Sql = Sql & ",porciva2 = " & DBSet(PorceIVA2, "N", Nulo2)
    Sql = Sql & ",porcrec2 = " & DBSet(PorceREC2, "N", Nulo2)
    Sql = Sql & ",codiiva2 = " & DBSet(TipoIVA2, "N", Nulo2)
    Sql = Sql & ",baseimp3 = " & DBSet(BaseIVA3, "N", Nulo3)
    Sql = Sql & ",impoiva3 = " & DBSet(ImpIVA3, "N", Nulo3)
    Sql = Sql & ",imporec3 = " & DBSet(ImpREC3, "N", Nulo3)
    Sql = Sql & ",porciva3 = " & DBSet(PorceIVA3, "N", Nulo3)
    Sql = Sql & ",porcrec3 = " & DBSet(PorceREC3, "N", Nulo3)
    Sql = Sql & ",codiiva3 = " & DBSet(TipoIVA3, "N", Nulo3)
    Sql = Sql & ",brutofac = " & DBSet(TotBruto, "N")
    Sql = Sql & ",impordto = " & DBSet(Round2(TotBruto - TotNeto, 2), "N")
    Sql = Sql & ",totalfac = " & DBSet(TotalFac, "N")
    Sql = Sql & " where " & cadwhere
    
    conn.Execute Sql

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
Dim Sql As String
Dim SumaBruto As Currency

On Error Resume Next

    '[Monica]16/09/2011: antes estaba puesto que solo se hiciera para Castelduc, ahora lo he parametrizado
    '                    porque tambien lo van a hacer de esta manera en Alzira
    'If vParamAplic.Cooperativa = 5 Then
    If vParamAplic.TipoCalculoComision = 1 Then
        ' Para Castelduc y Alzira el importe de descuento se prorratea con respecto a los kilos
        
        Sql = "select sum(cantreal) from facturas_variedad where codtipom = " & DBSet(TipoM, "T")
        Sql = Sql & " and numfactu = " & DBSet(Factura, "N") & " and fecfactu = " & DBSet(FecFactu, "F")
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
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
    
        Sql = "select sum(imporbru) from facturas_variedad where codtipom = " & DBSet(TipoM, "T")
        Sql = Sql & " and numfactu = " & DBSet(Factura, "N") & " and fecfactu = " & DBSet(FecFactu, "F")
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
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
Dim Sql As String
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
Dim TipoFactFor As Byte

Dim vHayReg As Byte

    On Error GoTo eRecalcularDtos

    Sql = "select * from facturas_variedad where codtipom = " & DBSet(TipoM, "T")
    Sql = Sql & " and numfactu = " & DBSet(Factura, "N") & " and fecfactu = " & DBSet(FecFactu, "F")
    Sql = Sql & " order by numlinea "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    Sql = ""
    Sql = DevuelveDesdeBDNew(cAgro, "facturas", "impdtoc", "codtipom", TipoM, "T", , "numfactu", Factura, "N", "fecfactu", FecFactu, "F")
    vImpDto = ComprobarCero(Sql)
    
    Sql = ""
    Sql = DevuelveDesdeBDNew(cAgro, "facturas", "dtocom1", "codtipom", TipoM, "T", , "numfactu", Factura, "N", "fecfactu", FecFactu, "F")
    vDto1 = ComprobarCero(Sql)
    
    Sql = ""
    Sql = DevuelveDesdeBDNew(cAgro, "facturas", "dtocom2", "codtipom", TipoM, "T", , "numfactu", Factura, "N", "fecfactu", FecFactu, "F")
    vDto2 = ComprobarCero(Sql)
    
    '++monica:030608:traemos el redondeo del precio
    Sql = ""
    Sql = DevuelveDesdeBDNew(cAgro, "facturas", "codclien", "codtipom", TipoM, "T", , "numfactu", Factura, "N", "fecfactu", FecFactu, "F")
    Cliente = ComprobarCero(Sql)
    Sql = ""
    Sql = DevuelveDesdeBDNew(cAgro, "clientes", "nrodecprec", "codclien", Cliente, "N")
    Rdo = ComprobarCero(Sql)
    
    vHayReg = 0
    
    While Not Rs.EOF
        vHayReg = 1
        
        TipoDto = DevuelveDesdeBDNew(cAgro, "clientes", "tipodtos", "codclien", Text1(3).Text, "N")
        If TipoFacturarForfaits(CStr(Rs!NumAlbar), CStr(Rs!numlinealbar)) = 1 Then 'kilos
            TipoFactFor = 1
            ImpDto = CalcularImporteDto(DBLet(Rs!cantreal, "N"), DBLet(Rs!precibru, "N"), TipoM, Factura, FecFactu, CStr(vImpDto), True)
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
            ImpDto = CalcularImporteDto(DBLet(Rs!Unidades, "N"), DBLet(Rs!precibru, "N"), TipoM, Factura, FecFactu, CStr(vImpDto), True)
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
        
        Sql2 = "update facturas_variedad set impornet = " & DBSet(vImpNeto, "N")
        Sql2 = Sql2 & ",precinet = " & DBSet(vPrecNeto, "N")
        Sql2 = Sql2 & ",dtocom1 = " & DBSet(vDto1, "N")
        Sql2 = Sql2 & ",dtocom2 = " & DBSet(vDto2, "N")
        Sql2 = Sql2 & " where codtipom = " & DBSet(TipoM, "T")
        Sql2 = Sql2 & " and numfactu = " & DBSet(Factura, "N")
        Sql2 = Sql2 & " and fecfactu = " & DBSet(FecFactu, "F")
        Sql2 = Sql2 & " and numlinea = " & DBSet(Rs!NumLinea, "N")
    
        conn.Execute Sql2
    
        UltimaLinea = DBLet(Rs!NumLinea, "N")
    
        Rs.MoveNext
    Wend
    
    Rs.Close
    
    '[Monica]16/09/2011: si no coincide la suma de dtos con el total descuento redondeamos en la ultima linea
    If vHayReg = 1 Then
        Sql2 = "select sum(imporbru) bruto, sum(impornet) neto from facturas_variedad "
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
                
                Sql2 = "update facturas_variedad set impornet = impornet + " & DBSet(Diferencia, "N")
                Sql2 = Sql2 & " where codtipom = " & DBSet(TipoM, "T")
                Sql2 = Sql2 & " and numfactu = " & DBSet(Factura, "N")
                Sql2 = Sql2 & " and fecfactu = " & DBSet(FecFactu, "F")
                Sql2 = Sql2 & " and numlinea = " & DBSet(UltimaLinea, "N")
            
                conn.Execute Sql2
        
                If TipoFactFor = 1 Then 'kilos
                    Sql2 = "update facturas_variedad set precinet = round(impornet / cantreal, " & DBSet(Rdo, "N") & ") "
                    Sql2 = Sql2 & " where codtipom = " & DBSet(TipoM, "T")
                    Sql2 = Sql2 & " and numfactu = " & DBSet(Factura, "N")
                    Sql2 = Sql2 & " and fecfactu = " & DBSet(FecFactu, "F")
                    Sql2 = Sql2 & " and numlinea = " & DBSet(UltimaLinea, "N")
                
                    conn.Execute Sql2
                Else 'unidades
                    'precio neto
                    Sql2 = "update facturas_variedad set precinet = round(impornet / unidades, " & DBSet(Rdo, "N") & ") "
                    Sql2 = Sql2 & " where codtipom = " & DBSet(TipoM, "T")
                    Sql2 = Sql2 & " and numfactu = " & DBSet(Factura, "N")
                    Sql2 = Sql2 & " and fecfactu = " & DBSet(FecFactu, "F")
                    Sql2 = Sql2 & " and numlinea = " & DBSet(UltimaLinea, "N")
                
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


Private Function RecalcularDtosLineas(TipoM As String, Factura As String, FecFactu As String, MenError As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
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

    On Error GoTo eRecalcularDtosLineas

    Sql = "select * from facturas_calibre where codtipom = " & DBSet(TipoM, "T")
    Sql = Sql & " and numfactu = " & DBSet(Factura, "N") & " and fecfactu = " & DBSet(FecFactu, "F")
    Sql = Sql & " order by numlinea, numline1 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    Sql = ""
    Sql = DevuelveDesdeBDNew(cAgro, "facturas", "impdtoc", "codtipom", TipoM, "T", , "numfactu", Factura, "N", "fecfactu", FecFactu, "F")
    vImpDto = ComprobarCero(Sql)
    
    Sql = ""
    Sql = DevuelveDesdeBDNew(cAgro, "facturas", "dtocom1", "codtipom", TipoM, "T", , "numfactu", Factura, "N", "fecfactu", FecFactu, "F")
    vDto1 = ComprobarCero(Sql)
    
    Sql = ""
    Sql = DevuelveDesdeBDNew(cAgro, "facturas", "dtocom2", "codtipom", TipoM, "T", , "numfactu", Factura, "N", "fecfactu", FecFactu, "F")
    vDto2 = ComprobarCero(Sql)
    
    '++monica:030608:traemos el redondeo del precio
    Sql = ""
    Sql = DevuelveDesdeBDNew(cAgro, "facturas", "codclien", "codtipom", TipoM, "T", , "numfactu", Factura, "N", "fecfactu", FecFactu, "F")
    Cliente = ComprobarCero(Sql)
    Sql = ""
    Sql = DevuelveDesdeBDNew(cAgro, "clientes", "nrodecprec", "codclien", Cliente, "N")
    Rdo = ComprobarCero(Sql)
    
    vHayReg = 0
    
    While Not Rs.EOF
        vHayReg = 1
        
        TipoDto = DevuelveDesdeBDNew(cAgro, "clientes", "tipodtos", "codclien", Text1(3).Text, "N")
        If TipoFacturarForfaits(CStr(Rs!NumAlbar), CStr(Rs!numlinealbar)) = 1 Then 'kilos
            TipoFactFor = 1
            ImpDto = CalcularImporteDto(DBLet(Rs!cantreal, "N"), DBLet(Rs!precibru, "N"), TipoM, Factura, FecFactu, CStr(vImpDto), True)
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
            ImpDto = CalcularImporteDto(DBLet(Rs!Unidades, "N"), DBLet(Rs!precibru, "N"), TipoM, Factura, FecFactu, CStr(vImpDto), True)
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
    
    RecalcularDtosLineas = True
    Exit Function

eRecalcularDtosLineas:
    If Err.Number <> 0 Then
        MenError = MenError & vbCrLf & Err.Description
        RecalcularDtosLineas = False
    End If
End Function


Private Sub DescontarFacturasACuenta(TipoM As String, Factu As String, fecFac As String, Cliente As String)
Dim Sql As String
Dim cadwhere As String

    Sql = "select codtipom, numfactu, fecfactu, totalfac from facturas "
    cadwhere = "where codtipom = 'EAC' and codclien = " & DBSet(Cliente, "N")
    cadwhere = cadwhere & " and (codtipom, numfactu, fecfactu) not in (select codtipomcta, numfactucta, fecfactucta from facturas_acuenta) "

    
    Sql = Sql & cadwhere
    
    If TotalRegistrosConsulta(Sql) <> 0 Then
        
        Set frmMens = New frmMensajes
        
        frmMens.OpcionMensaje = 22
        frmMens.cadwhere = cadwhere
        
        frmMens.Show vbModal
        
        Set frmMens = Nothing
        
    End If

End Sub


Private Function ActualizarVariedades(codTipoM As String, NumFactu As String, FecFactu As String, MensError As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim SQL1 As String

    On Error GoTo eActualizarVariedades

    ActualizarVariedades = False

    Sql = "select numlinea, numalbar, numlinealbar from facturas_variedad where codtipom = " & DBSet(codTipoM, "T")
    Sql = Sql & " and numfactu = " & DBSet(NumFactu, "N") & " and fecfactu = " & DBSet(FecFactu, "F")

    Set Rs2 = New ADODB.Recordset
    Rs2.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    While Not Rs2.EOF
        SQL1 = "select sum(if(cantreal is null,0,cantreal)), sum(if(cantfact is null,0,cantfact)), sum(if(imporbru is null,0,imporbru)), sum(if(impornet is null,0,impornet))"
        SQL1 = SQL1 & " from facturas_calibre "
        SQL1 = SQL1 & " where codtipom = " & DBSet(codTipoM, "T")
        SQL1 = SQL1 & " and numfactu = " & DBSet(NumFactu, "N")
        SQL1 = SQL1 & " and fecfactu = " & DBSet(FecFactu, "F")
        SQL1 = SQL1 & " and numlinea = " & DBSet(Rs2!NumLinea, "N")
        conn.Execute SQL1
    
        Set Rs = New ADODB.Recordset
        Rs.Open SQL1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs.EOF Then
            Sql = "update facturas_variedad set cantreal = " & DBSet(Rs.Fields(0).Value, "N")
            Sql = Sql & " where codtipom = " & DBSet(codTipoM, "T")
            Sql = Sql & " and numfactu = " & DBSet(NumFactu, "N")
            Sql = Sql & " and fecfactu = " & DBSet(FecFactu, "F")
            Sql = Sql & " and numlinea = " & DBSet(Rs2!NumLinea, "N")
    
            conn.Execute Sql
                
            Sql = "update facturas_variedad set cantfact = " & DBSet(Rs.Fields(1).Value, "N")
            Sql = Sql & " where codtipom = " & DBSet(codTipoM, "T")
            Sql = Sql & " and numfactu = " & DBSet(NumFactu, "N")
            Sql = Sql & " and fecfactu = " & DBSet(FecFactu, "F")
            Sql = Sql & " and numlinea = " & DBSet(Rs2!NumLinea, "N")
    
            conn.Execute Sql
            
            Sql = "update facturas_variedad set imporbru = " & DBSet(Rs.Fields(2).Value, "N")
            Sql = Sql & " where codtipom = " & DBSet(codTipoM, "T")
            Sql = Sql & " and numfactu = " & DBSet(NumFactu, "N")
            Sql = Sql & " and fecfactu = " & DBSet(FecFactu, "F")
            Sql = Sql & " and numlinea = " & DBSet(Rs2!NumLinea, "N")
    
            conn.Execute Sql
            
            Sql = "update facturas_variedad set impornet = " & DBSet(Rs.Fields(3).Value, "N")
            Sql = Sql & " where codtipom = " & DBSet(codTipoM, "T")
            Sql = Sql & " and numfactu = " & DBSet(NumFactu, "N")
            Sql = Sql & " and fecfactu = " & DBSet(FecFactu, "F")
            Sql = Sql & " and numlinea = " & DBSet(Rs2!NumLinea, "N")
    
            conn.Execute Sql
        
            If TipoFacturarForfaits(DBLet(Rs2!NumAlbar, "N"), DBLet(Rs2!numlinealbar, "N")) = 1 Then
                Sql = "update facturas_variedad set precibru = round(imporbru / cantreal,4), "
                Sql = Sql & " precinet = round(impornet / cantreal,4) "
            Else
                Sql = "update facturas_variedad set precibru = round(imporbru / unidades,4), "
                Sql = Sql & " precinet = round(impornet / unidades,4) "
            End If
            Sql = Sql & " where codtipom = " & DBSet(codTipoM, "T")
            Sql = Sql & " and numfactu = " & DBSet(NumFactu, "N")
            Sql = Sql & " and fecfactu = " & DBSet(FecFactu, "F")
            Sql = Sql & " and numlinea = " & DBSet(Rs2!NumLinea, "N")
    
            conn.Execute Sql
        
        End If
        Rs.Close
        Set Rs = Nothing
        
        Rs2.MoveNext
    Wend
    
    Set Rs2 = Nothing
    
eActualizarVariedades:
    If Err.Number = 0 Then ActualizarVariedades = True
End Function


Private Function FactContabilizada(ByRef EstaEnTesoreria As String) As Boolean
Dim LEtra As String, numasien As String
    
    On Error GoTo EContab
    
    'NO deberia poder modificar fras anteriors a fecha inicio ejercicio
    'Cojo la letra de serie
    LEtra = ObtenerLetraSerie(Text1(6).Text)
    
    'Primero comprobaremos que esta el cobro en contabilidad
    EstaEnTesoreria = ""
    If Not ComprobarCobroArimoney(EstaEnTesoreria, LEtra, CLng(Text1(0).Text), CDate(Text1(1).Text)) Then
        FactContabilizada = True
        Exit Function
    End If

    'comprobar que se puede modificar/eliminar la factura
    If Me.Check1(1).Value = 1 Then 'si esta contabilizada
        'comprobar en la contabilidad si esta contabilizada
        If LEtra <> "" Then
            If vParamAplic.ContabilidadNueva Then
                numasien = DevuelveDesdeBDNew(cConta, "factcli", "numasien", "numserie", LEtra, "T", , "numfactu", Text1(0).Text, "N", "anofactu", Year(Text1(1).Text), "N")
            Else
                numasien = DevuelveDesdeBDNew(cConta, "cabfact", "numasien", "numserie", LEtra, "T", , "codfaccl", Text1(0).Text, "N", "anofaccl", Year(Text1(1).Text), "N")
            End If
            If Val(ComprobarCero(numasien)) <> 0 Then
'                FactContabilizada = True
'                MsgBox "La factura esta contabilizada y no se puede modificar.", vbInformation
'                Exit Function
            Else
                numasien = ""
            End If
        Else
'            MsgBox "Las factura de venta no tienen asignada una letra de serie", vbInformation
            numasien = ""
        End If
        
        LEtra = "La factura esta en la contabilidad"
        If numasien <> "" Then LEtra = LEtra & vbCrLf & "Nº asiento: " & numasien
        LEtra = LEtra & vbCrLf & vbCrLf & "¿Continuar?"
        
        numasien = String(50, "*") & vbCrLf
        numasien = numasien & numasien & vbCrLf & vbCrLf
        LEtra = numasien & LEtra & vbCrLf & vbCrLf & numasien
        If MsgBox(LEtra, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            FactContabilizada = False
        Else
            FactContabilizada = True
        End If
    Else
        FactContabilizada = False
    End If
    
EContab:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Factura contabilizada", Err.Description
End Function



'En vTesoreria pondremos como estaba el recibo
'Es decir. El  msgbox que pondra al final lo guardo en esta variable
Private Function ComprobarCobroArimoney(vTesoreria As String, LEtra As String, Codfaccl As Long, fecha As Date) As Boolean
Dim vR As ADODB.Recordset
Dim Cad As String

On Error GoTo EComprobarCobroArimoney
    ComprobarCobroArimoney = False
    Set vR = New ADODB.Recordset
    
    If vParamAplic.ContabilidadNueva Then
        Cad = "Select * from cobros where numserie='" & LEtra & "'"
        Cad = Cad & " AND numfactu =" & Codfaccl
        Cad = Cad & " AND fecfactu =" & DBSet(fecha, "F")
    
    Else
        Cad = "Select * from scobro where numserie='" & LEtra & "'"
        Cad = Cad & " AND codfaccl =" & Codfaccl
        Cad = Cad & " AND fecfaccl =" & DBSet(fecha, "F")
    End If
    '
    vTesoreria = ""
    vR.Open Cad, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    If vR.EOF Then
        vTesoreria = "NO se ha encotrado ningun vencimiento en la tesoreria"
    Else
        While Not vR.EOF
            Cad = ""
            If DBLet(vR!codrem, "T") <> "" Then
                Cad = "El cobro asociado a la factura esta remesado(" & vR!codrem & ")"
            Else
                If DBLet(vR!recedocu, "N") = 1 Then
                    Cad = "Documento recibido"
                Else
                    If vParamAplic.ContabilidadNueva Then
                        If DBLet(vR!transfer, "N") = 1 Then
                            Cad = "Esta en una transferencia"
                        Else
                           If DBLet(vR!impcobro, "N") > 0 Then Cad = "Esta parcialmente cobrado: " & vR!impcobro
                        
                            
                                    'Si hubeira que poner mas coas iria aqui
                        End If 'transfer
                    
                    Else
                        If DBLet(vR!Estacaja, "N") = 1 Then
                            Cad = "Cobrado por caja"
                        Else
                            If DBLet(vR!transfer, "N") = 1 Then
                                Cad = "Esta en una transferencia"
                            Else
                               If DBLet(vR!impcobro, "N") > 0 Then Cad = "Esta parcialmente cobrado: " & vR!impcobro
                            
                                
                                        'Si hubeira que poner mas coas iria aqui
                            End If 'transfer
                        End If 'estacaja
                    End If
                End If 'recdedocu
            End If 'remesado
            If Cad <> "" Then vTesoreria = vTesoreria & "Vto: " & vR!numorden & "      " & Cad & vbCrLf
            vR.MoveNext
        Wend
    End If
    vR.Close
    
    If vTesoreria <> "" Then
        Cad = vTesoreria & vbCrLf & vbCrLf
        If vUsu.Nivel > 1 Then
            MsgBox Cad, vbExclamation
        Else
            Cad = Cad & "¿Continuar?"
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then ComprobarCobroArimoney = True
        End If
    Else
        ComprobarCobroArimoney = True
    End If
            
EComprobarCobroArimoney:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set vR = Nothing
End Function


Private Function EsFacturaEnvases(TipoFac As String, NumFact As String, fecFac As String) As Boolean
Dim Sql As String

    Sql = "select count(*) from facturas_variedad where codtipom = " & DBSet(TipoFac, "T")
    Sql = Sql & " and numfactu = " & DBSet(NumFact, "N")
    Sql = Sql & " and fecfactu = " & DBSet(fecFac, "F")
    
    EsFacturaEnvases = (TotalRegistros(Sql) = 0)

End Function


Private Function HayFacturasACuenta(vTipo As String, vNumfac As String, vFecfac As String) As Boolean
Dim Sql As String

    Sql = "select count(*) from facturas_acuenta where codtipom = " & DBSet(Trim(vTipo), "T")
    Sql = Sql & " and numfactu = " & DBSet(vNumfac, "N") & " and fecfactu = " & DBSet(vFecfac, "F")
    
    HayFacturasACuenta = (TotalRegistros(Sql) <> 0)
    
End Function
