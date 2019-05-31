VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFacturasCompra 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturas Compra"
   ClientHeight    =   11055
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   17850
   Icon            =   "frmFacturasCompra.frx":0000
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
      TabIndex        =   67
      Top             =   45
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   68
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
      TabIndex        =   65
      Top             =   45
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   66
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
      TabIndex        =   64
      Top             =   315
      Width           =   1605
   End
   Begin VB.Frame Frame2 
      Height          =   4230
      Left            =   90
      TabIndex        =   30
      Top             =   810
      Width           =   17610
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
         Left            =   4050
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Factura|F|N|||facturascom|fecfactu|dd/mm/yyyy|S|"
         Top             =   810
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
         Index           =   6
         Left            =   6930
         MaxLength       =   12
         TabIndex        =   3
         Tag             =   "Fecha Recepción|F|N|||facturascom|fecrecep|dd/mm/yyyy||"
         Top             =   810
         Width           =   1350
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
         Index           =   0
         Left            =   6615
         TabIndex        =   7
         Tag             =   "Contabilizado|N|N|||facturascom|intconta|0||"
         Top             =   2025
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
         Index           =   4
         Left            =   1350
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "Forma Pago|N|N|0|999|facturascom|codforpa|000||"
         Text            =   "Text1"
         Top             =   1395
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
         TabIndex        =   38
         Text            =   "Text2"
         Top             =   1395
         Width           =   5970
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
         Left            =   3915
         MaxLength       =   5
         TabIndex        =   6
         Tag             =   "%Dto GNRAL|N|S|0|100|facturascom|dtognral|##0.00||"
         Top             =   1935
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
         Left            =   1350
         MaxLength       =   5
         TabIndex        =   5
         Tag             =   "%Dto PPago|N|S|0|100|facturascom|dtoppago|##0.00||"
         Top             =   1935
         Width           =   945
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
         Left            =   2790
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   34
         Text            =   "Text2"
         Top             =   270
         Width           =   5475
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
         Height          =   1320
         Index           =   2
         Left            =   225
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Tag             =   "Observaciones|T|S|||facturascom|observac|||"
         Top             =   2730
         Width           =   8040
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
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   0
         Tag             =   "Proveedor|N|N|0|999999|facturascom|codprove|000000|S|"
         Text            =   "Text1"
         Top             =   270
         Width           =   1080
      End
      Begin VB.TextBox Text1 
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
         Left            =   1350
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Nº Factura|T|S|||facturascom|numfactu||S|"
         Text            =   "Text1 7"
         Top             =   855
         Width           =   1350
      End
      Begin VB.Frame FrameFactura 
         Height          =   3825
         Left            =   8685
         TabIndex        =   42
         Top             =   270
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
            Index           =   32
            Left            =   4365
            MaxLength       =   15
            TabIndex        =   92
            Tag             =   "Importe DtoGnral|N|S|0||facturascom|impgnral|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   540
            Width           =   1260
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
            Left            =   5475
            MaxLength       =   5
            TabIndex        =   59
            Tag             =   "% REC 1|N|S|0|99.90|facturascom|porcrec1|#0.00|N|"
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
            TabIndex        =   58
            Tag             =   "Importe REC 1|N|S|0||facturascom|imporec1|#,###,###,##0.00|N|"
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
            TabIndex        =   57
            Tag             =   "% REC 2|N|S|0|99.90|facturascom|porcrec2|#0.00|N|"
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
            TabIndex        =   56
            Tag             =   "Importe REC 2|N|S|0||facturascom|imporec2|#,###,###,##0.00|N|"
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
            TabIndex        =   55
            Tag             =   "% REC 3|N|S|0|99.90|facturascom|porcrec3|#0.00|N|"
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
            TabIndex        =   54
            Tag             =   "Importe REC 3|N|S|0||facturascom|imporec3|#,###,###,##0.00|N|"
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
            TabIndex        =   24
            Tag             =   "Total Factura|N|S|0||facturascom|totalfac|#,###,###,##0.00|N|"
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
            TabIndex        =   23
            Tag             =   "Importe IVA 3|N|S|0||facturascom|impoiva3|#,###,###,##0.00|N|"
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
            TabIndex        =   21
            Tag             =   "% IVA 3|N|S|0|99.90|facturascom|porciva3|#0.00|N|"
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
            TabIndex        =   22
            Tag             =   "Base Imponible 3|N|S|0||facturascom|baseimp3|#,###,###,##0.00|N|"
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
            TabIndex        =   19
            Tag             =   "Importe IVA 2|N|S|0||facturascom|impoiva2|#,###,###,##0.00|N|"
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
            TabIndex        =   17
            Tag             =   "& IVA 2|N|S|0|99.90|facturascom|porciva2|#0.00|N|"
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
            TabIndex        =   18
            Tag             =   "Base Imponible 2 |N|S|0||facturascom|baseimp2|#,###,###,##0.00|N|"
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
            TabIndex        =   15
            Tag             =   "Importe IVA 1|N|S|0||facturascom|impoiva1|#,###,###,##0.00|N|"
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
            TabIndex        =   13
            Tag             =   "% IVA 1|N|S|0|99.90|facturascom|porciva1|#0.00|N|"
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
            TabIndex        =   14
            Tag             =   "Base Imponible 1|N|S|0||facturascom|baseimp1|#,###,###,##0.00|N|"
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
            TabIndex        =   12
            Tag             =   "IVA 1|N|S|0|9|facturascom|codiiva1|0|N|"
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
            TabIndex        =   16
            Tag             =   "IVA 2|N|S|0|9|facturascom|codiiva2|0|N|"
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
            TabIndex        =   20
            Tag             =   "IVA 3|N|S|0|9|facturascom|codiiva3|0|N|"
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
            TabIndex        =   9
            Tag             =   "Bruto Factura|N|S|0||facturascom|brutofac|#,###,###,##0.00|N|"
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
            Left            =   2790
            MaxLength       =   15
            TabIndex        =   10
            Tag             =   "Importe Dto.PPago|N|S|0||facturascom|impppago|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   540
            Width           =   1260
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
            TabIndex        =   11
            Text            =   "Text1 7"
            Top             =   540
            Width           =   1650
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
            Index           =   14
            Left            =   5895
            TabIndex        =   95
            Top             =   540
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
            Index           =   11
            Left            =   4185
            TabIndex        =   94
            Top             =   540
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Imp.DtoGnral"
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
            Left            =   4365
            TabIndex        =   93
            Top             =   270
            Width           =   1710
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
            TabIndex        =   61
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
            TabIndex        =   60
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
            TabIndex        =   53
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
            TabIndex        =   52
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
            TabIndex        =   51
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
            TabIndex        =   50
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
            TabIndex        =   49
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
            TabIndex        =   48
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
            TabIndex        =   47
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
            TabIndex        =   46
            Top             =   270
            Width           =   1695
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
            TabIndex        =   45
            Top             =   270
            Width           =   1530
         End
         Begin VB.Label Label1 
            Caption         =   "Imp.PPago"
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
            Left            =   2790
            TabIndex        =   44
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
            Left            =   2655
            TabIndex        =   43
            Top             =   540
            Width           =   135
         End
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
         Left            =   5040
         MaxLength       =   5
         TabIndex        =   91
         Top             =   3195
         Width           =   945
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   6660
         Picture         =   "frmFacturasCompra.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   855
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F.Recepción"
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
         Left            =   5445
         TabIndex        =   41
         Top             =   855
         Width           =   1275
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1035
         ToolTipText     =   "Buscar Destino"
         Top             =   1440
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
         TabIndex        =   39
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label Label1 
         Caption         =   "%Dto GNRAL"
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
         Left            =   2475
         TabIndex        =   37
         Top             =   1980
         Width           =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "F.Factura"
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
         Left            =   2790
         TabIndex        =   36
         Top             =   855
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "%Dto PP"
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
         Left            =   225
         TabIndex        =   35
         Top             =   1980
         Width           =   1140
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   3780
         Picture         =   "frmFacturasCompra.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   855
         Width           =   240
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   1710
         ToolTipText     =   "Zoom descripción"
         Top             =   2430
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
         TabIndex        =   33
         Top             =   2430
         Width           =   1485
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
         Left            =   225
         TabIndex        =   32
         Top             =   315
         Width           =   990
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1350
         ToolTipText     =   "Buscar Proveedor"
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
         Left            =   225
         TabIndex        =   31
         Top             =   855
         Width           =   1125
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5100
      Left            =   90
      TabIndex        =   40
      Top             =   5175
      Width           =   17610
      _ExtentX        =   31062
      _ExtentY        =   8996
      _Version        =   393216
      Tabs            =   1
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
      TabPicture(0)   =   "frmFacturasCompra.frx":0122
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameAux1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame FrameAux1 
         BorderStyle     =   0  'None
         Height          =   4290
         Left            =   135
         TabIndex        =   70
         Top             =   480
         Width           =   17190
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
            Left            =   11070
            MaxLength       =   5
            TabIndex        =   78
            Tag             =   "Dto.Linea2|N|N|||facturascom_variedad|dtoline2|#0.00||"
            Text            =   "Dto.Linea"
            Top             =   2250
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
            Height          =   330
            Index           =   1
            Left            =   6210
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   90
            Text            =   "Nombre forfait"
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
            Index           =   1
            Left            =   5985
            MaskColor       =   &H00000000&
            TabIndex        =   89
            ToolTipText     =   "Buscar Forfait"
            Top             =   2250
            Visible         =   0   'False
            Width           =   195
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
            Left            =   3510
            MaskColor       =   &H00000000&
            TabIndex        =   86
            ToolTipText     =   "Buscar Variedad"
            Top             =   2250
            Visible         =   0   'False
            Width           =   195
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
            Left            =   3735
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   85
            Text            =   "Nombre variedad"
            Top             =   2250
            Width           =   1200
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
            TabIndex        =   84
            Tag             =   "Proveedor|N|N|||facturascom_variedad|codprove||S|"
            Text            =   "Proveedor"
            Top             =   2250
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
            Index           =   1
            Left            =   1020
            MaxLength       =   12
            TabIndex        =   83
            Tag             =   "Num.Factura|T|N|||facturascom_variedad|numfactu|0000000|S|"
            Text            =   "NumFact"
            Top             =   2250
            Visible         =   0   'False
            Width           =   690
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
            Left            =   5040
            MaxLength       =   16
            TabIndex        =   72
            Tag             =   "Forfait|T|N|||facturascom_variedad|codforfait|||"
            Text            =   "Forfait"
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
            Index           =   6
            Left            =   7470
            MaxLength       =   12
            TabIndex        =   73
            Tag             =   "Cajas|N|N|||facturascom_variedad|numcajas|###,##0||"
            Text            =   "cajas"
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
            Index           =   2
            Left            =   1710
            MaxLength       =   12
            TabIndex        =   81
            Tag             =   "Fec.Factu|F|N|||facturascom_variedad|fecfactu||S|"
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
            Index           =   3
            Left            =   2430
            MaxLength       =   12
            TabIndex        =   80
            Tag             =   "Num.Linea|N|N|||facturascom_variedad|numlinea|000|S|"
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
            Index           =   7
            Left            =   9495
            MaxLength       =   12
            TabIndex        =   76
            Tag             =   "Precio|N|N|||facturascom_variedad|precio|###,##0.0000||"
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
            Index           =   8
            Left            =   10305
            MaxLength       =   5
            TabIndex        =   77
            Tag             =   "Dto.Linea|N|N|||facturascom_variedad|dtoline1|#0.00||"
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
            Index           =   9
            Left            =   11835
            MaxLength       =   12
            TabIndex        =   79
            Tag             =   "Importe|N|N|||facturascom_variedad|importe|##,###,##0.00||"
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
            Index           =   10
            Left            =   12600
            MaxLength       =   2
            TabIndex        =   82
            Tag             =   "CodIva|N|N|||facturascom_variedad|codigiva|00||"
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
            Index           =   4
            Left            =   3105
            MaxLength       =   6
            TabIndex        =   71
            Tag             =   "Variedad|N|N|||facturascom_variedad|codvarie|000000||"
            Text            =   "Varied"
            Top             =   2250
            Visible         =   0   'False
            Width           =   420
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
            Left            =   8280
            MaxLength       =   7
            TabIndex        =   74
            Tag             =   "Peso Bruto|N|S|||facturascom_variedad|pesobrut|###,##0||"
            Text            =   "P.Bruto"
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
            Index           =   12
            Left            =   8910
            MaxLength       =   7
            TabIndex        =   75
            Tag             =   "Peso Neto|N|S|||facturascom_variedad|pesoneto|###,##0||"
            Text            =   "Peso Ne"
            Top             =   2250
            Visible         =   0   'False
            Width           =   555
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   1
            Left            =   90
            TabIndex        =   87
            Top             =   135
            Width           =   1230
            _ExtentX        =   2170
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
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "frmFacturasCompra.frx":013E
            Height          =   3105
            Left            =   90
            TabIndex        =   88
            Top             =   630
            Width           =   16860
            _ExtentX        =   29739
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
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   90
      TabIndex        =   28
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
         TabIndex        =   29
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
      TabIndex        =   26
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
      TabIndex        =   25
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
      TabIndex        =   27
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
      TabIndex        =   62
      Text            =   "Text1 7"
      Top             =   1710
      Width           =   1485
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   17115
      TabIndex        =   69
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
   Begin VB.Label Label1 
      Caption         =   "Imp.Descuento 2"
      Height          =   255
      Index           =   10
      Left            =   6705
      TabIndex        =   63
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
Attribute VB_Name = "frmFacturasCompra"
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
Public hcoCodProve As String 'cod. movim
Public hcoCodTipoM As String 'Codigo detalle de Movimiento(ALC)
Public hcoFechaMov As String 'fecha del movimiento

'========== VBLES PRIVADAS ====================
Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmVar As frmBasico2 'Form Mto de variedades
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmFor As frmManForfaits 'Form Mto de Forfaits
Attribute frmFor.VB_VarHelpID = -1

Private WithEvents frmPro As frmManProve 'Form Mto de Proveedores
Attribute frmPro.VB_VarHelpID = -1
Private WithEvents frmFPag As frmManFpago 'Form Mto de Formas de Pago
Attribute frmFPag.VB_VarHelpID = -1

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
Dim ProveedorAnt As String
Dim FPagoAnt As String
Dim DtoPPagoAnt As String
Dim DtoGnralAnt As String
Dim ObsAnt As String
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

Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos
Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal
Dim indice As Byte

Dim TipoFactura As Byte
Dim PulsadoF2 As Boolean
Private BuscaChekc As String

Dim VarieAnt As String
Dim CajasLinAnt As Currency

Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 0 'Variedades
            Set frmVar = New frmBasico2
            
            AyudaVariedad frmVar
            
            Set frmVar = Nothing
            
            PonerFoco txtAux(4)
        Case 1 'forfaits
            Set frmFor = New frmManForfaits
            frmFor.DatosADevolverBusqueda = "0|"
            frmFor.CodigoActual = txtAux(5).Text
            frmFor.Show vbModal
            Set frmFor = Nothing
            PonerFoco txtAux(5)
    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1

End Sub

Private Sub Check1_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "check1(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "check1(" & Index & ")|"
    End If
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
            LLamaLineas Modo, 0, "DataGrid1"
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
                        DataGrid1.AllowAddNew = False
                        If Not Adoaux(1).Recordset.EOF Then Adoaux(1).Recordset.MoveFirst
                    End If
                    ModificaLineas = 0
                    LLamaLineas Modo, 0, "DataGrid1"
                    PonerModo 2
                    DataGrid1.Enabled = True
                    If Not Data1.Recordset.EOF Then _
                        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            
                    'Habilitar las opciones correctas del menu segun Modo
                    PonerModoOpcionesMenu (Modo)
                    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
                    DataGrid1.Enabled = True
             End Select
            
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
        MandaBusquedaPrevia ""
    Else
        LimpiarCampos
        LimpiarDataGrids
        CadenaConsulta = "Select facturascom.* "
        CadenaConsulta = CadenaConsulta & "from " & NombreTabla & " where (1=1) " & Ordenacion
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
    ProveedorAnt = Text1(3).Text
    FPagoAnt = Text1(4).Text
    DtoPPagoAnt = Text1(7).Text
    DtoGnralAnt = Text1(8).Text
    ObsAnt = Text1(2).Text
    ContaAnt = Check1(0).Value
    
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
    'If Index = 2 Then NumTabMto = 3
    PonerModo 5, Index
 

    
    Select Case NumTabMto
        Case 1 ' variedades
            vWhere = Replace(ObtenerWhereCP(False), "facturascom", "facturascom_variedad")
            vWhere = vWhere & " and numlinea=" & Adoaux(1).Recordset!NumLinea
            If Not BloqueaRegistro("facturascom_variedad", vWhere) Then
                TerminaBloquear
                Exit Sub
            End If
            If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
                J = DataGrid1.Bookmark - DataGrid1.FirstRow
                DataGrid1.Scroll 0, J
                DataGrid1.Refresh
            End If
            
        '    anc = ObtenerAlto(Me.DataGrid1)
            anc = DataGrid1.Top
            If DataGrid1.Row < 0 Then
                anc = anc + 240
            Else
                anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 10
            End If
        
            For J = 0 To 4
                txtAux(J).Text = DataGrid1.Columns(J).Text
            Next J
            
            
            Text2(0).Text = DataGrid1.Columns(5).Text  ' nombre de variedad
            txtAux(5).Text = DataGrid1.Columns(6).Text
            Text2(1).Text = DataGrid1.Columns(7).Text  ' nombre de forfait
            txtAux(6).Text = DataGrid1.Columns(8).Text
            txtAux(11).Text = DataGrid1.Columns(9).Text
            txtAux(12).Text = DataGrid1.Columns(10).Text
            txtAux(7).Text = DataGrid1.Columns(11).Text
            txtAux(8).Text = DataGrid1.Columns(12).Text
            txtAux(13).Text = DataGrid1.Columns(13).Text
            txtAux(9).Text = DataGrid1.Columns(14).Text
            txtAux(10).Text = DataGrid1.Columns(15).Text
            
            
'[Monica]13/05/2015: dejamos modificar la variedad
           BloquearBtn Me.btnBuscar(0), False
           BloquearBtn Me.btnBuscar(1), False
            
            ModificaLineas = 2 'Modificar
            LLamaLineas ModificaLineas, anc, "DataGrid1"
            
            'Añadiremos el boton de aceptar y demas objetos para insertar
            Me.lblIndicador.Caption = "MODIFICAR"
            PonerModoOpcionesMenu (Modo)
            PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
            DataGrid1.Enabled = True
            
            PonerFoco txtAux(10)
            Me.DataGrid1.Enabled = False

            VarieAnt = txtAux(5).Text
            CajasLinAnt = ComprobarCero(txtAux(6).Text)
      
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
            b = (xModo = 1 Or xModo = 2)
            For jj = 4 To 13
                txtAux(jj).Height = DataGrid1.RowHeight - 10
                txtAux(jj).Top = alto + 5
                txtAux(jj).visible = b
                txtAux(jj).Enabled = b
            Next jj
            Text2(0).Height = DataGrid1.RowHeight - 10
            Text2(0).Top = alto + 5
            Text2(0).visible = b
           
            btnBuscar(0).Height = DataGrid1.RowHeight - 10
            btnBuscar(0).Top = alto + 5
            btnBuscar(0).visible = b
            
            Text2(1).Height = DataGrid1.RowHeight - 10
            Text2(1).Top = alto + 5
            Text2(1).visible = b
           
            btnBuscar(1).Height = DataGrid1.RowHeight - 10
            btnBuscar(1).Top = alto + 5
            btnBuscar(1).visible = b
            
            
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
    Cad = Cad & vbCrLf & "Proveedor:  " & Text1(6).Text
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
        DataGrid1.Enabled = True
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

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
    'Viene de DblClick en documentos en la ficha de cliente
    If hcoCodProve <> "" And Not Data1.Recordset.EOF Then PonerCadenaBusqueda
    
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
    For kCampo = 1 To ToolAux.Count
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

'--monica
'    CodTipoMov = "ALV" 'hcoCodTipoM
    VieneDeBuscar = False
    
        
    '## A mano
    NombreTabla = "facturascom"
    NomTablaLineas = "facturascom_variedad" 'Tabla lineas de variedades
    
    Ordenacion = " ORDER BY facturascom.codprove, facturascom.numfactu, facturascom.fecfactu"
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    CadenaConsulta = "select * from facturascom "
    If hcoCodProve <> "" Then
        CadenaConsulta = CadenaConsulta & " where codprove = " & hcoCodProve & " and numfactu = " & DBSet(hcoCodProve, "N") & " and fecfactu = " & DBSet(hcoFechaMov, "F") & " and codtipom <> 'EAC'"
    Else
        CadenaConsulta = CadenaConsulta & " where numfactu is null "
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
    NumTabMto = 1
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
    For i = 0 To Check1.Count - 1
        Me.Check1(i).Value = 0
    Next i
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    If Modo = 4 Then TerminaBloquear
End Sub


Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod almacen
    Text2(indice + 2).Text = RecuperaValor(CadenaSeleccion, 2) 'Nombre del almacen
End Sub


Private Sub frmFor_DatoSeleccionado(CadenaSeleccion As String)
'Forfaits
    txtAux(5).Text = RecuperaValor(CadenaSeleccion, 1) 'codforfait
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 2) 'nomconfe

End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Articulos
    txtAux(4).Text = RecuperaValor(CadenaSeleccion, 1) 'codvarie
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
    If txtAux(4) <> "" Then
        txtAux(10) = DevuelveDesdeBDNew(cAgro, "variedades", "codigiva", "codvarie", txtAux(4), "N")
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
        Text1(CByte(imgFec(0).Tag) + 5).Text = Format(vFecha, "dd/mm/yyyy") '<===
    End If
    ' ********************************************
End Sub

Private Sub frmPro_DatoSeleccionado(CadenaSeleccion As String)
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
        CadB = "codprove = " & DBSet(RecuperaValor(CadenaSeleccion, 3), "N")
        CadB = CadB & " and numfactu = " & DBSet(RecuperaValor(CadenaSeleccion, 1), "T")
        CadB = CadB & " and fecfactu = " & DBSet(RecuperaValor(CadenaSeleccion, 2), "F")
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


Private Sub frmZ_Actualizar(vCampo As String)
     Text1(indice).Text = vCampo
End Sub

Private Sub imgBuscar_Click(Index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. de Proveedor
            indice = 3
            PonerFoco Text1(indice)
            Set frmPro = New frmManProve
            frmPro.DatosADevolverBusqueda = "0|1|2|"
            frmPro.Show vbModal
            Set frmPro = Nothing
            PonerFoco Text1(indice)
        
        Case 1 'Forma de Pago
            indice = 4
            PonerFoco Text1(indice)
            Set frmFPag = New frmManFpago
            frmFPag.DatosADevolverBusqueda = "0|1|"
            frmFPag.Show vbModal
            Set frmFPag = Nothing
            PonerFoco Text1(indice)
            
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

    If Index = 0 Then
        indice = 1
    Else
        indice = 6
    End If
    imgFec(0).Tag = Index
    If Text1(indice).Text <> "" Then frmC.NovaData = Text1(indice).Text

    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco Text1(indice) '<===
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
    AbrirListado (17)
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub


Private Sub mnModificar_Click()
    '[Monica]28/08/2013: añadida la pregunta de continuar a pesar de estar en la contabilidad y arimoney
                '[Monica]05/10/2018: no decimos nada si es una factura en b
    If Check1(0).Value = 1 And Text1(6).Text <> TipoFactB Then
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
        Case 1, 6 'Fecha factura
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
            
        Case 3 'Proveedor
            If PonerFormatoEntero(Text1(Index)) Then
                '[Monica]28/08/2013: añado or modo= 4
                If Modo = 1 Or Modo = 4 Then
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), "proveedor", "nomprove")
                Else
                    PonerDatosProveedor (Text1(Index).Text)
                    If Text2(Index).Text = "" Then
                        cadMen = "No existe el Proveedor: " & Text1(Index).Text & vbCrLf
                        cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                        If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                            Set frmPro = New frmManProve
                            frmPro.DatosADevolverBusqueda = "0|1|"
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
            
         Case 7, 8 'descuentos
            If Modo = 1 Then Exit Sub
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 4

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
    
    If chkVistaPrevia = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select facturascom.* from " & NombreTabla
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

    AyudaFacturasComprasPrev frmFac, , CadB

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
    
    For i = 0 To 0
        Select Case i
            Case 0 'variedades
                If Data1.Recordset.RecordCount > 0 Then
                    CargaGrid DataGrid1, Adoaux(1), True
                Else
                    CargaGrid DataGrid1, Adoaux(1), False
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
    
    
    Text2(3).Text = PonerNombreDeCod(Text1(3), "proveedor", "nomprove", "codprove", "N") 'proveedores
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
    BuscaChekc = ""
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Or Facturas <> "" Or hcoCodProve <> "" Then
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
    
    For i = 9 To 32
        BloquearTxt Text1(i), Not (Modo = 1)
        Text1(i).Enabled = (Modo = 1)
    Next i
'    Me.Check1.Enabled = (Modo = 1)
    
    b = (Modo <> 1) And (Modo <> 3)
    
    'Campos Nº Factura bloqueado y en azul
    BloquearTxt Text1(0), b, True 'And (TipoFactura = 0)  'numero factura
    BloquearTxt Text1(6), b 'And (TipoFactura = 0) 'fecha recepcion
    BloquearTxt Text1(1), b 'fechafactura
    BloquearTxt Text1(3), b 'proveedor
    BloquearChk Me.Check1(0), (Modo <> 1)
    
    imgFec(0).Enabled = b
    imgFec(0).visible = b
    imgFec(1).Enabled = b
    imgFec(1).visible = b
    
    
    '[Monica]20/09/2012:desbloqueo a mano el campo de fecha para poder modificarlo
    If Modo = 4 Then
        Text1(1).Locked = False
        Text1(1).BackColor = vbWhite
        
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
    
    
    'Si no es modo lineas Boquear los TxtAux
    For i = 0 To txtAux.Count - 1
        txtAux(i).visible = False
        BloquearTxt txtAux(i), True
    Next i
    
    For i = 0 To 1
        Text2(i).visible = ((Modo = 5) And (indFrame = 1))
        Text2(i).Enabled = False
    Next i
    
    BloquearTxt Text2(16), (Modo <> 5)
    
    BloquearBtn Me.btnBuscar(0), True
    BloquearBtn Me.btnBuscar(1), True
    
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    BloquearImgBuscar Me, Modo, ModificaLineas
    BloquearImgFec Me, 0, Modo
                    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
    
    
    ' *** si n'hi han llínies i algún tab que no te datagrid ***
'    Select Case NumTabMto
'        Case 1
            BloquearFrameAux Me, "FrameAux1", Modo, NumTabMto
'    End Select
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
Dim Codmacta As String

    On Error GoTo EDatosOK

    DatosOk = False
    
'    ComprobarDatosTotales

    'comprobamos datos OK de la tabla scafac
    b = CompForm2(Me, 2, "Frame2") ' , 1) 'Comprobar formato datos ok de la cabecera: opcion=1
    If Not b Then Exit Function
    
    If Modo = 3 Then
        'si el tipo de factura es manual y no hemos introducido valor en numero de factura
        If Text1(0).Text = "" Then
            MsgBox "El Número de Factura no puede estar vacio. Reintroduzca.", vbExclamation
            PonerFoco Text1(0)
            b = False
        End If

        'comprobamos que no exista ya la factura en la tabla facturas de ariagro
        Sql = ""
        '[Monica]04/03/2013: antes no me miraba la fecha de factura, estaba comentado
        Sql = DevuelveDesdeBDNew(cAgro, "facturascom", "numfactu", "codprove", Text1(3).Text, "N", , "numfactu", Text1(0).Text, "T", "fecfactu", Text1(1).Text, "F")
        If Sql <> "" Then
            MsgBox "Factura ya existente. Reintroduzca.", vbExclamation
            PonerFoco Text1(0)
            b = False
        End If
        If Not b Then Exit Function
        'comprobamos que no exista ya en la tabla facturas de contabilidad
        
        Codmacta = DevuelveDesdeBDNew(cAgro, "proveedor", "codmacta", "codprove", Text1(3), "N")
'++
        If Codmacta <> "" Then
            Sql = ""
            '[Monica]04/03/2013: en la contabilidad hemos de mirar el año de la factura, no la fecha
            If vParamAplic.ContabilidadNueva Then
                Sql = DevuelveDesdeBDNew(cConta, "factpro", "numfactu", "codmacta", Codmacta, "T", , "numfactu", Text1(0).Text, "T", "anofactu", Year(CDate(Text1(1).Text)), "N")
            Else
                Sql = DevuelveDesdeBDNew(cConta, "cabfactprov", "numfacpr", "codmacta", Codmacta, "T", , "codfacpr", Text1(0).Text, "T", "anofacpr", Year(CDate(Text1(1).Text)), "N")
            End If
            If Sql <> "" Then
                MsgBox "Factura existente en contabilidad. Reintroduzca.", vbExclamation
                PonerFoco Text1(0)
                b = False
            End If
        Else
            MsgBox "El Proveedor no tiene cuenta contable asociada. Revise.", vbExclamation
            b = False
        End If
        If Not b Then Exit Function
    
    
        '[Monica]20/06/2017: control de fechas que antes no estaba
        If vParamAplic.NumeroConta <> 0 Then
            ResultadoFechaContaOK = EsFechaOKConta(CDate(Text1(6).Text))
            If ResultadoFechaContaOK > 0 Then
                If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
                Exit Function
            End If
        End If
    
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
    If Check1(0).Value = 1 Then
        If MsgBox("Esta factura está en Contabilidad y Arimoney. " & vbCrLf & vbCrLf & "Si la modifica realice los cambios en estas aplicaciones." & vbCrLf & vbCrLf & "¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Sub
        Else
            If CDate(Text1(6).Text) <= vEmpresa.FechaUltIVA Then
                If MsgBox("La factura es de un período liquidado. " & vbCrLf & vbCrLf & "¿ Seguro que desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                    Exit Sub
                End If
            End If
        End If
    End If
    
    If BloqueaRegistro(NombreTabla, "numfactu = " & DBSet(Data1.Recordset!NumFactu, "T")) Then
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
       Case 1 'variedades
            ' *************** canviar la pregunta ****************
            Cad = "¿Seguro que desea eliminar la Variedad?"
            Cad = Cad & vbCrLf & "Factura: " & Adoaux(1).Recordset.Fields(1)
            Cad = Cad & vbCrLf & "Variedad: " & Adoaux(1).Recordset.Fields(5) & " - " & Adoaux(1).Recordset.Fields(6)
            
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
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
        Case 8  ' Listado de facturas
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
    DataGrid1.Enabled = Not b
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim b As Boolean
Dim Opcion As Byte
Dim Sql As String

    On Error GoTo ECargaGRid

    b = DataGrid1.Enabled
    Select Case vDataGrid.Name
        Case "DataGrid1" 'variedades
            Opcion = 0
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
                       
         Case "DataGrid1" 'facturascom_variedad
'select codtipom,numfactu,fecfactu,numlinea,codartic,sartic.nomartic,cantidad,precioar,dtolinea,importel,ampliaci,codigiva,numalbar,numlinea
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;S|txtAux(4)|T|Codigo|1000|;S|btnBuscar(0)|B|||;S|Text2(0)|T|Variedad|2400|;"
            tots = tots & "S|txtAux(5)|T|Forfait|1500|;S|btnBuscar(1)|B|||;"
            tots = tots & "S|Text2(1)|T|Nombre|3000|;S|txtAux(6)|T|Cajas|1200|;S|txtAux(11)|T|Peso Bruto|1400|;S|txtAux(12)|T|Peso Neto|1400|;"
            tots = tots & "S|txtAux(7)|T|Precio|1000|;S|txtAux(8)|T|Dto1|700|;S|txtAux(13)|T|Dto2|700|;S|txtAux(9)|T|Importe|1900|;N||||0|;"
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
Dim Variedad As String



    'Quitar espacios en blanco
    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    Select Case Index
        Case 4 'variedad
            If PonerFormatoEntero(txtAux(Index)) Then
                Text2(0).Text = DevuelveDesdeBDNew(cAgro, "variedades", "nomvarie", "codvarie", txtAux(Index).Text, "N")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Variedad: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        indice = Index + 2
                        Set frmVar = New frmManVariedad
                        frmVar.DatosADevolverBusqueda = "0|1|"
                        frmVar.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        
                        frmVar.Show vbModal
                        Set frmVar = Nothing
                        '++monica
                        BloqueaRegistro "facturascom", "numfactu = " & Text1(0).Text
'                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(0).Text = ""
            End If
            
            '[Monica]15/05/2015: si modificamos no traemos nada
            If ModificaLineas = 2 Then Exit Sub
            
            txtAux(10).Text = DevuelveDesdeBDNew(cAgro, "variedades", "codigiva", "codvarie", txtAux(4), "N")
            
            
        Case 5 'forfait
            If txtAux(Index).Text <> "" Then
                Text2(1) = PonerNombreDeCod(txtAux(Index), "forfaits", "nomconfe", , "T")
                If Text2(1).Text = "" Then
                    cadMen = "No existe el Forfait: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        indice = Index + 2
                        Set frmFor = New frmManForfaits
                        frmFor.DatosADevolverBusqueda = "0|1|"
                        frmFor.NuevoCodigo = txtAux(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        
                        frmFor.Show vbModal
                        Set frmFor = Nothing
                        '++monica
                        BloqueaRegistro "facturascom", "numfactu = " & Text1(0).Text
'                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                '++monica:02/12/2008 control d que el forfait sea de la variedad introducida
                Else
                    Variedad = ""
                    Variedad = DevuelveDesdeBDNew(cAgro, "forfaits", "codvarie", "codforfait", txtAux(Index).Text, "T")
                    If Variedad <> "" Then
                        If CInt(Variedad) <> CInt(txtAux(4).Text) Then
                            MsgBox "El Forfait no es de la Variedad introducida.", vbExclamation
                        End If
                    End If
                '++
                End If
            Else
                Text2(1).Text = ""
            End If
        
        
        Case 6 ' cajas
            If txtAux(Index).Text <> "" Then PonerFormatoEntero txtAux(Index)
            
            '[Monica]11/12/2018: solo en el caso de que me modifiquen las cajas o inserten linea
            If (ModificaLineas = 2 And CajasLinAnt <> ComprobarCero(txtAux(Index))) Or ModificaLineas = 1 Then
                CalculoPesoNeto
            End If
        Case 11, 12 'peso bruto, peso neto
            PonerFormatoEntero txtAux(Index)
        
        Case 7 ' Precio
            PonerFormatoDecimal txtAux(Index), 10
            
        Case 8, 13  'Descuentos
            PonerFormatoDecimal txtAux(Index), 4 'Tipo 4: Decimal(4,2)
            
        Case 9 'Importe Linea
            If PonerFormatoDecimal(txtAux(Index), 3) Then   'Tipo 3: Decimal(10,2)
                PonerFocoBtn cmdAceptar
            End If
            
            
    End Select
    If (Index = 12 Or Index = 13 Or Index = 7 Or Index = 8 Or Index = 9) Then 'Cant., Precio, Dto1, Dto2
'        If Trim(TxtAux(3).Text) = "" Or Trim(TxtAux(4).Text) = "" Then Exit Sub
'        If Trim(TxtAux(6).Text) = "" Or Trim(TxtAux(7).Text) = "" Then Exit Sub
        If txtAux(8).Text = "" Then txtAux(8).Text = 0
        If txtAux(13).Text = "" Then txtAux(13).Text = 0
        TipoDto = DevuelveDesdeBDNew(cAgro, "proveedor", "tipodtos", "codprove", Text1(3).Text, "N")
        
        txtAux(9).Text = CalcularImporte(txtAux(12).Text, txtAux(7).Text, txtAux(8).Text, txtAux(13).Text, TipoDto, 0)
        PonerFormatoDecimal txtAux(9), 3
    End If
    
End Sub

Private Sub CalculoPesoNeto()
Dim Sql As String
Dim Kilos1 As Currency
Dim Kilos2 As Currency
Dim Rs As ADODB.Recordset

    Sql = "select kiloscaj, kilosuni  from forfaits where codforfait = " & DBSet(txtAux(5).Text, "T")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Kilos1 = 0
    Kilos2 = 0
    If Not Rs.EOF Then
        Kilos1 = DBLet(Rs!kiloscaj, "N")
        Kilos2 = DBLet(Rs!KilosUni, "N")
    End If
    
    'si hay cajas
    If ComprobarCero(txtAux(6).Text) <> 0 Then
        If Kilos1 <> 0 Then
            txtAux(11).Text = Round2(Kilos1 * ImporteSinFormato(txtAux(6).Text), 0)
            PonerFormatoEntero txtAux(11)
            txtAux(12).Text = txtAux(11).Text
        End If
    End If
    ' si hay unidades
    ' No hacemos nada pq en el formulario no se pide

End Sub



Private Function Eliminar() As Boolean
Dim Sql As String, LEtra As String, Sql2 As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim Mens As String

    On Error GoTo FinEliminar

    If Data1.Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    b = True

    If b Then
        'Eliminar en tablas de cabecera de albaran
        '------------------------------------------
        Sql = " " & ObtenerWhereCP(True)
        
        'Lineas de envases (facturas_variedad)
        conn.Execute "Delete from facturascom_variedad " & Replace(Sql, "facturascom", "facturascom_variedad")
        

        'Cabecera de factura
        conn.Execute "Delete from " & NombreTabla & Sql
        
        
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

    b = True
    If Adoaux(1).Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    Mens = ""
    
    
    '[Monica]28/08/2013: insertamos en el log que hemos eliminado una linea en una factura contabilizada
    If Check1(0).Value = 1 Then
        '------------------------------------------------------------------------------
        '  LOG de acciones.
        Set LOG = New cLOG
        'campo = "Facturas de Clientes: " & """"
        
        CADENA = "Eliminar Linea de Variedad "
        LOG.Insertar 9, vUsu, CADENA & Adoaux(1).Recordset.Fields(0) & " " & Adoaux(1).Recordset.Fields(1) & " " & Adoaux(1).Recordset.Fields(2) & " de " & Text1(25).Text & " Linea " & Adoaux(1).Recordset.Fields(3) & " " & Adoaux(1).Recordset.Fields(5)
        Set LOG = Nothing
          '-----------------------------------------------------------------------------
    End If
    
    
    'Eliminar en tablas de facturascom_variedad
    '------------------------------------------
    Sql = " where codprove = " & Adoaux(1).Recordset.Fields(0)
    Sql = Sql & " and numfactu = " & DBSet(Adoaux(1).Recordset.Fields(1), "T")
    Sql = Sql & " and fecfactu = " & DBSet(Adoaux(1).Recordset.Fields(2), "F")
    Sql = Sql & " and numlinea = " & Adoaux(1).Recordset.Fields(3)


    'Lineas de variedades
    conn.Execute "Delete from facturascom_variedad " & Sql
    
FinEliminar:
    If Err.Number <> 0 Or Not b Then
        MuestraError Err.Number, "Eliminar Variedad de la Factura ", Err.Description & " " & Mens
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
    If Adoaux(2).Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    Mens = ""
    
    b = True
    
    
    '[Monica]28/08/2013: insertamos en el log que hemos insertado una linea en una factura contabilizada
    If Check1(0).Value = 1 Then
        '------------------------------------------------------------------------------
        '  LOG de acciones.
        Set LOG = New cLOG
        'campo = "Facturas de Clientes: " & """"
        
        CADENA = "Eliminar Linea Facturas a Cuenta "
        LOG.Insertar 9, vUsu, CADENA & Adoaux(2).Recordset.Fields(0) & " " & Adoaux(2).Recordset.Fields(1) & " " & Adoaux(2).Recordset.Fields(2) & " de " & Text1(25).Text & " " & Adoaux(2).Recordset.Fields(3) & " " & Adoaux(2).Recordset.Fields(4) & " " & Adoaux(2).Recordset.Fields(5)
        Set LOG = Nothing
          '-----------------------------------------------------------------------------
    End If
    
    'Eliminar en tablas de facturas_acuenta
    '------------------------------------------
    Sql = " where codtipom = " & DBSet(Adoaux(2).Recordset.Fields(0), "T")
    Sql = Sql & " and numfactu = " & Adoaux(2).Recordset.Fields(1)
    Sql = Sql & " and fecfactu = " & DBSet(Adoaux(2).Recordset.Fields(2), "F")
    Sql = Sql & " and codtipomcta = " & DBSet(Adoaux(2).Recordset.Fields(3), "T")
    Sql = Sql & " and numfactucta = " & Adoaux(2).Recordset.Fields(4)
    Sql = Sql & " and fecfactucta = " & DBSet(Adoaux(2).Recordset.Fields(5), "F")


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
    If Adoaux(0).Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    Mens = ""
   
   '[Monica]28/08/2013: insertamos en el log que hemos eliminado una linea en una factura contabilizada
    If Check1(0).Value = 1 Then
        '------------------------------------------------------------------------------
        '  LOG de acciones.
        Set LOG = New cLOG
        'campo = "Facturas de Clientes: " & """"
        
        CADENA = "Eliminar Linea de Variedades "
        LOG.Insertar 9, vUsu, CADENA & Adoaux(0).Recordset.Fields(0) & " " & Adoaux(0).Recordset.Fields(1) & " " & Adoaux(0).Recordset.Fields(2) & " de " & Text1(25).Text & " Linea " & Adoaux(0).Recordset.Fields(3) & " Alb " & Adoaux(0).Recordset.Fields(4) & " " & Adoaux(0).Recordset.Fields(5)
        Set LOG = Nothing
          '-----------------------------------------------------------------------------
    End If
        
    Sql = "delete from facturascom_variedad where codprove = " & DBSet(Adoaux(0).Recordset.Fields(0), "N")
    Sql = Sql & " and numfactu = " & Adoaux(0).Recordset.Fields(1)
    Sql = Sql & " and fecfactu = " & DBSet(Adoaux(0).Recordset.Fields(2), "F")
    Sql = Sql & " and numlinea = " & DBSet(Adoaux(0).Recordset.Fields(3), "N")
    conn.Execute Sql
                   
    
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

    CargaGrid DataGrid1, Me.Adoaux(1), False 'lineas
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & Replace(ObtenerWhereCP(False), "facturascom.", "") & ")"
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
    
    Sql = "facturascom.codprove = " & DBSet(Text1(3).Text, "N") & " and facturascom.numfactu= " & DBSet(Text1(0).Text, "T") & " and facturascom.fecfactu= " & DBSet(Text1(1).Text, "F")
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
            Sql = "SELECT facturascom_variedad.codprove, numfactu,fecfactu, facturascom_variedad.numlinea,"
            Sql = Sql & " facturascom_variedad.codvarie, variedades.nomvarie, facturascom_variedad.codforfait,"
            Sql = Sql & " forfaits.nomconfe, facturascom_variedad.numcajas, facturascom_variedad.pesobrut,"
            Sql = Sql & " facturascom_variedad.pesoneto, facturascom_variedad.precio, facturascom_variedad.dtoline1,"
            Sql = Sql & " facturascom_variedad.dtoline2, facturascom_variedad.importe,"
            Sql = Sql & " facturascom_variedad.codigiva  "
            Sql = Sql & " FROM facturascom_variedad, variedades, forfaits " 'lineas de variedades de la factura
            Sql = Sql & " WHERE facturascom_variedad.codvarie = variedades.codvarie "
            Sql = Sql & " and facturascom_variedad.codforfait = forfaits.codforfait "
            
            If enlaza Then
                Sql = Sql & " and " & Replace(ObtenerWhereCP(False), "facturascom", "facturascom_variedad")
            Else
                '[Monica]19/04/2017: cambio de condicion por rapidez
                Sql = Sql & " and numfactu is null " '= -1"
            End If
            Sql = Sql & " ORDER BY codprove,numfactu,fecfactu,numlinea"
                    
    End Select
    
    MontaSQLCarga = Sql
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean, bAux As Boolean
Dim i As Integer

    b = ((Modo = 2) Or (Modo = 0)) And (Facturas = "") And (hcoCodProve = "") 'Or (Modo = 5 And ModificaLineas = 0)
    'Buscar
    Toolbar1.Buttons(5).Enabled = b
    Me.mnBuscar.Enabled = b
    'Vore Tots
    Toolbar1.Buttons(6).Enabled = b
    Me.mnVerTodos.Enabled = b
    'Añadir
    Toolbar1.Buttons(1).Enabled = b
    Me.mnModificar.Enabled = b
    
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0) And (Facturas = "") And (hcoCodProve = "")
    'Modificar
    '[Monica]28/08/2013: dejo modificar la factura aunque esté contabilizada y demas
    '                    quito la condicion de todos los checks
    Toolbar1.Buttons(2).Enabled = b 'And Not (Check1(0).Value = 1 )
    Me.mnModificar.Enabled = b 'And Not (Check1(0).Value = 1 Or Check1(1).Value = 1 Or Check1(2).Value = 1)
    'eliminar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnEliminar.Enabled = b
    'Impresión de factura
    Toolbar1.Buttons(8).Enabled = True '((Modo = 2) And (Facturas = "")) Or (hcoCodProve <> "")
    Me.mnImprimir.Enabled = True '((Modo = 2) And (Facturas = "")) Or (hcoCodProve <> "")

    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
    
    '[Monica]28/08/2013: dejo modificar la factura aunque esté contabilizada y demas
    '                    quito la condicion de todos los checks
    b = (Modo = 2) And (Facturas = "") And (hcoCodProve = "")  ' And Not (Check1(0).Value = 1 Or (Check1(1).Value = 1 And vUsu.Nivel >= 1) Or Check1(2).Value = 1)
    For i = 1 To ToolAux.Count
        ToolAux(i).Buttons(1).Enabled = b
        
        If b Then
            Select Case i
              Case 1
                bAux = (b And Me.Adoaux(1).Recordset.RecordCount > 0)
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




Private Function ModificarFactura(MenError As String) As Boolean
Dim Sql As String
Dim Sql2 As String
    
    
    On Error GoTo eModificarFactura
    
    ModificarFactura = False
    
    'Como hay actualizacion en cascada no hace falta modificar las fechas en las tablas de las lineas
    
    Sql = "update facturascom set codforpa = " & DBSet(Text1(4).Text, "N")
    Sql = Sql & " ,dtoppago = " & DBSet(Text1(7).Text, "N")
    Sql = Sql & " ,dtognral = " & DBSet(Text1(8).Text, "N")
    Sql = Sql & " ,observac = " & DBSet(Text1(2).Text, "T")
    Sql = Sql & " ,fecfactu = " & DBSet(Text1(1).Text, "F")
    Sql = Sql & " ,fecrecep = " & DBSet(Text1(6).Text, "F")
    Sql = Sql & ", codprove = " & DBSet(Text1(3).Text, "N")
    
    '[Monica]10/09/2018: actualizamos intconta
    Sql = Sql & ", intconta = " & Check1(0).Value
    
    Sql = Sql & " where codprove = " & DBSet(Text1(6).Text, "T")
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
    If ProveedorAnt <> Text1(3).Text Then CADENA = CADENA & " Pro " & ProveedorAnt & " por " & Text1(3).Text
    If FPagoAnt <> Text1(4).Text Then CADENA = CADENA & " FPago " & FPagoAnt & " por " & Text1(4).Text
    If DtoPPagoAnt <> Text1(7).Text Then CADENA = CADENA & " DtoPPago " & DtoPPagoAnt & " por " & Text1(7).Text
    If DtoGnralAnt <> Text1(8).Text Then CADENA = CADENA & " DtoGnralAnt " & DtoGnralAnt & " por " & Text1(8).Text
    If ObsAnt <> Text1(2).Text Then CADENA = CADENA & " Obs " & Trim(ObsAnt) & " por " & Trim(Text1(2).Text)
    If ContaAnt <> Check1(0).Value Then CADENA = CADENA & " IntConta " & ContaAnt & " por " & Check1(0).Value
    

    If CADENA <> "" And ((Text1(1).Text <> FechaAnt) Or Check1(0).Value = 1) Then
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
    
    
'    MenError = "Recalcular Dtos calibres"
'    If b Then b = RecalcularDtosLineas(Text1(6).Text, Text1(0).Text, Text1(1).Text, MenError)
    
    
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
    
    Sql = CadenaInsertarDesdeForm(Me)
    conn.Execute Sql


    CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
    PonerCadenaBusqueda
    PonerModo 2
    'Ponerse en Modo Insertar Lineas
'                BotonMtoLineas 0, "Variedades"
    
    BotonAnyadirLinea 1
    
EInsertarCab:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
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
        Case 1: nomFrame = "FrameAux1" 'variedades
    End Select
    ' ***************************************************************
    
    
    
    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        
        Select Case Index
            Case 1
                '[Monica]22/02/2012: insertamos la transaccion
                conn.BeginTrans
                If InsertarDesdeForm2(Me, 2, nomFrame) Then
                    
                    '[Monica]28/08/2013: insertamos en el log que hemos insertado una linea en una factura contabilizada
                    If Check1(0).Value = 1 Then
                        '------------------------------------------------------------------------------
                        '  LOG de acciones.
                        Set LOG = New cLOG
                        
                        CADENA = "Inserta Linea de Variedades "
                        CADENA = CADENA & Text1(6).Text & " " & Text1(0).Text & " " & Text1(1).Text & " de " & Text1(25).Text
                        CADENA = CADENA & " Variedad " & txtAux(4).Text & " " & Text2(0).Text
                        
                        LOG.Insertar 9, vUsu, CADENA
                        Set LOG = Nothing
                          '-----------------------------------------------------------------------------
                    End If
                    
                    ' *** si n'hi ha que fer alguna cosa abas d'insertar
                    Mens = "Recalcular Dtos lineas"
'                    b = RecalcularDtos(txtAux3(0).Text, txtAux3(1).Text, txtAux3(2).Text, Mens)
                    b = True
                    
                    If b Then
                        
                        conn.CommitTrans
                        
                        CalcularDatosFactura
                        ' *************************************************
                        b = BloqueaRegistro("facturascom", "codprove = " & DBSet(Data1.Recordset!codProve, "N") & " and numfactu = " & DBSet(Data1.Recordset!NumFactu, "T") & " and fecfactu = " & DBSet(Data1.Recordset!FecFactu, "F"))
                        CargaGrid DataGrid1, Adoaux(1), True
                        If b Then BotonAnyadirLinea NumTabMto
                        'SSTab1.Tab = NumTabMto
                    Else
                        conn.RollbackTrans
                    End If
                Else
                    conn.RollbackTrans
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
    BloquearTxt Text1(3), True
    BloquearTxt Text1(1), True
    
    
    ' **************************************************

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case NumTabMto
        Case 1: vtabla = "facturascom_variedad"
    End Select
    ' ********************************************************
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case NumTabMto
        Case 1
            NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", vWhere)
            ' ***************************************************************

            AnyadirLinea DataGrid1, Adoaux(1)
    
            anc = DataGrid1.Top
            If DataGrid1.Row < 0 Then
                anc = anc + 240 '210
            Else
                anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
            End If
          
            LLamaLineas ModificaLineas, anc, "DataGrid1"
        
            LimpiarCamposLin "FrameAux1"
            txtAux(0).Text = Text1(3).Text 'codprove
            txtAux(1).Text = Text1(0).Text 'numfactu
            txtAux(2).Text = Text1(1).Text 'fecfactu
            txtAux(3).Text = NumF
            PonerFoco txtAux(4)
            For i = 0 To 1
                Text2(i).Text = ""
            Next i
            txtAux(10).Enabled = False
            txtAux(10).visible = False
'            BloquearTxt txtAux(9), True
'            BloquearTxt Text2(16), False
            BloquearBtn Me.btnBuscar(0), False
            BloquearBtn Me.btnBuscar(1), False
        ' ******************************************
        
        ' *** si n'hi han llínies sense datagrid ***
    
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
        Case 1: nomFrame = "FrameAux1" 'variedades
    End Select
    ' **************************************************************

    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        
        Select Case NumTabMto
        Case 1
            conn.BeginTrans
        
            '[Monica]28/08/2013: insertamos en el log que hemos eliminado una linea en una factura contabilizada
            If Check1(0).Value = 1 Then
                '------------------------------------------------------------------------------
                '  LOG de acciones.
                Set LOG = New cLOG
                
                CADENA = "Modificar Linea de Variedades "
                
                LOG.Insertar 9, vUsu, CADENA & Adoaux(0).Recordset.Fields(0) & " " & Adoaux(0).Recordset.Fields(1) & " " & Adoaux(0).Recordset.Fields(2) & " de " & Text1(25).Text & " Linea " & Adoaux(0).Recordset.Fields(3) & " Alb " & Adoaux(0).Recordset.Fields(4) & " " & Adoaux(0).Recordset.Fields(5)
                Set LOG = Nothing
                  '-----------------------------------------------------------------------------
            End If
        
        
            If ModificaDesdeFormulario2(Me, 2, nomFrame) Then
                ' *** si cal que fer alguna cosa abas d'insertar ***
                ' ******************************************************
' antes
'                Mens = "Recalcular Dtos lineas"
'                b = RecalcularDtos(txtAux3(0).Text, txtAux3(1).Text, txtAux3(2).Text, Mens)

                
                b = True
    
    '            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
                If b Then
                
                    conn.CommitTrans
                    
                    V = Adoaux(1).Recordset.Fields(3) 'el 2 es el nº de llinia
                
                    CalcularDatosFactura
                    ModificaLineas = 0
        
'                    V = AdoAux(0).Recordset.Fields(3) 'el 2 es el nº de llinia
                    CargaGrid DataGrid1, Adoaux(1), True
    
                    ' *** si n'hi han tabs ***
                    SSTab1.Tab = 0
    
                    DataGrid1.SetFocus
                    Adoaux(1).Recordset.Find (Adoaux(1).Recordset.Fields(3).Name & " =" & V)
    
                    LLamaLineas ModificaLineas, 0, "DataGrid1"
               End If
               
            End If
    
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
    DatosOkLlin = b

EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " codprove = " & DBSet(Text1(3).Text, "N") & " and numfactu= " & DBSet(Text1(0).Text, "T") & " and fecfactu = " & DBSet(Text1(1).Text, "F")
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
Private Sub PonerDatosProveedor(codProve As String, Optional nifProve As String)
Dim vProveedor As CProveedor
Dim Observaciones As String
    
    On Error GoTo EPonerDatos
    
    If codProve = "" Then
        LimpiarDatosProveedor
        Exit Sub
    End If

    Set vProveedor = New CProveedor
    
    'si se ha modificado el cliente volver a cargar los datos
    If vProveedor.Existe(codProve) Then
        If vProveedor.LeerDatos(codProve) Then
            Text1(3).Text = vProveedor.Codigo
            FormateaCampo Text1(3)
            If (Modo = 3) Or (Modo = 4) Then
                Text2(3).Text = vProveedor.Nombre  'Nom clien
                Text1(4).Text = vProveedor.Forpago
                Text2(4).Text = PonerNombreDeCod(Text1(4), "forpago", "nomforpa")
                Text1(7).Text = Format(vProveedor.DtoPPago, FormatoDescuento)
                Text1(8).Text = Format(vProveedor.DtoGnral, FormatoDescuento)
                
                
            End If

            Observaciones = DBLet(vProveedor.Observaciones)
            If Observaciones <> "" Then
                MsgBox Observaciones, vbInformation, "Observaciones del proveedor"
            End If
        End If
    Else
        LimpiarDatosProveedor
    End If
    Set vProveedor = Nothing

EPonerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Datos Proveedor", Err.Description
End Sub


Private Sub LimpiarDatosProveedor()
Dim i As Byte

    Text1(2).Text = ""
    Text1(4).Text = ""
    Text1(7).Text = ""
    Text1(8).Text = ""

    Text2(3).Text = ""
    Text2(4).Text = ""
End Sub
    



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
    
    
    If CalcularDatosFacturaCompra(cadwhere, NombreTabla, NomTablaLineas) Then
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
'Private Function CalcularDatosFacturaCompra_old(cadwhere As String, NomTabla As String, NomTablaLin As String) As Boolean
''cadWhere: cad para la where de la SQL que selecciona las lineas del albaran o la factura
''nomTabla: nombre de la tabla de albaranes(scaalp) o de AlbaranesXFactura(scafpa)
''           segun llamemos desde recepcion de facturas o desde Hco de Facturas
'Dim Rs As ADODB.Recordset
'Dim i As Integer
'
'Dim Sql As String
'Dim cadAux As String
'Dim cadAux1 As String
'
''Aqui vamos acumulando los totales
'Dim TotBruto As Currency
'Dim TotNeto As Currency
'Dim TotImpIVA As Currency
'
'Dim ImpAux As Currency
'Dim impiva As Currency
'Dim ImpREC As Currency
'Dim ImpBImIVA As Currency 'Importe Base imponible a la que hay q aplicar el IVA
'
'Dim vBruto As Currency
'Dim vNeto As Currency
'
'Dim exentoIVA As Boolean
'Dim conDesplaz As Boolean
'
'Dim BaseImp As Currency
'Dim BaseIVA1 As Currency
'Dim BaseIVA2 As Currency
'Dim BaseIVA3 As Currency
'
'Dim BrutoFac As Currency
'
'Dim ImpIVA1 As Currency
'Dim ImpIVA2 As Currency
'Dim ImpIVA3 As Currency
'
'Dim PorceIVA1 As Currency
'Dim PorceIVA2 As Currency
'Dim PorceIVA3 As Currency
'
'Dim ImpREC1 As Currency
'Dim ImpREC2 As Currency
'Dim ImpREC3 As Currency
'
'Dim PorceREC1 As Currency
'Dim PorceREC2 As Currency
'Dim PorceREC3 As Currency
'
'Dim TipoIVA1 As Currency
'Dim TipoIVA2 As Currency
'Dim TipoIVA3 As Currency
'
'Dim ImpDto1 As Currency
'Dim ImpDto2 As Currency
'Dim TotalFac As Currency
'
'Dim IvaAnt As Integer
'Dim cadwhere1 As String
'
'Dim Nulo2 As String
'Dim Nulo3 As String
'
'    CalcularDatosFacturaCompra_old = False
'    On Error GoTo ECalcular
'
'    BaseImp = 0
'    BaseIVA1 = 0
'    BaseIVA2 = 0
'    BaseIVA3 = 0
'
'    BrutoFac = 0
'
'    ImpIVA1 = 0
'    ImpIVA2 = 0
'    ImpIVA3 = 0
'
'    PorceIVA1 = 0
'    PorceIVA2 = 0
'    PorceIVA3 = 0
'
'    ImpREC1 = 0
'    ImpREC2 = 0
'    ImpREC3 = 0
'
'    PorceREC1 = 0
'    PorceREC2 = 0
'    PorceREC3 = 0
'
'    TipoIVA1 = 0
'    TipoIVA2 = 0
'    TipoIVA3 = 0
'
'    ImpDto1 = 0
'    ImpDto2 = 0
'    TotalFac = 0
'
'    'Agrupar el importe bruto por tipos de iva
'    cadwhere1 = Replace(cadwhere, "facturascom", "facturascom_variedad")
'    Sql = "SELECT facturascom_variedad.codigiva, sum(importe) as bruto"
'    Sql = Sql & " FROM facturascom_variedad "
'    Sql = Sql & " WHERE " & cadwhere1
'    Sql = Sql & " GROUP BY 1 "
'    Sql = Sql & " ORDER BY 1 "
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'    TotBruto = 0
'    TotNeto = 0
'    TotImpIVA = 0
'    vBruto = 0
'    vNeto = 0
'    i = 1
'
'    '[Monica]23/08/2013: he metido en el if la instruccion de ivaant
'    If Not Rs.EOF Then
'        Rs.MoveFirst
'        IvaAnt = Rs.Fields(0).Value
'
'    Else
'        '[Monica]23/08/2013 añado el else
'
'        Sql = "update facturascom "
'        Sql = Sql & "set baseimp1 = " & ValorNulo
'        Sql = Sql & ",impoiva1 = " & ValorNulo
'        Sql = Sql & ",imporec1 = " & ValorNulo
'        Sql = Sql & ",porciva1 = " & ValorNulo
'        Sql = Sql & ",porcrec1 = " & ValorNulo
'        Sql = Sql & ",codiiva1 = " & ValorNulo
'        Nulo2 = "N"
'        Nulo3 = "N"
'        If DBSet(TipoIVA2, "N", "S") = ValorNulo Then Nulo2 = "S"
'        If DBSet(TipoIVA3, "N", "S") = ValorNulo Then Nulo3 = "S"
'        Sql = Sql & ",baseimp2 = " & ValorNulo
'        Sql = Sql & ",impoiva2 = " & ValorNulo
'        Sql = Sql & ",imporec2 = " & ValorNulo
'        Sql = Sql & ",porciva2 = " & ValorNulo
'        Sql = Sql & ",porcrec2 = " & ValorNulo
'        Sql = Sql & ",codiiva2 = " & ValorNulo
'        Sql = Sql & ",baseimp3 = " & ValorNulo
'        Sql = Sql & ",impoiva3 = " & ValorNulo
'        Sql = Sql & ",imporec3 = " & ValorNulo
'        Sql = Sql & ",porciva3 = " & ValorNulo
'        Sql = Sql & ",porcrec3 = " & ValorNulo
'        Sql = Sql & ",codiiva3 = " & ValorNulo
'        Sql = Sql & ",brutofac = " & ValorNulo
'        Sql = Sql & ",impordto = " & ValorNulo
'        Sql = Sql & ",totalfac = " & ValorNulo
'        Sql = Sql & " where " & cadwhere
'
'        conn.Execute Sql
'
'        CalcularDatosFacturaCompra = True
'        Exit Function
'
'    End If
'    While Not Rs.EOF
'                                        '[Monica]05/05/2015: añadimos la condicion de que la suma de brutos de la factura sea distinta de
'                                        '                    para que no llegue como base a la contabilidad
'        If IvaAnt <> Rs.Fields(0).Value And vBruto <> 0 Then
'            TotBruto = TotBruto + vBruto
'            TotNeto = TotNeto + vNeto
'            ImpBImIVA = vNeto
'
'
'            'Obtener el % de IVA
'            cadAux = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(IvaAnt), "N")
'
'            'aplicar el IVA a la base imponible de ese tipo
'            impiva = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
'
'            'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
'            'los vamos acumulando
'            TotImpIVA = TotImpIVA + impiva
'
'            If CInt(Data1.Recordset!TipoIvac) = 2 Then
'                'Obtener el % de RECARGO
'                cadAux1 = DevuelveDesdeBDNew(cConta, "tiposiva", "porcerec", "codigiva", CStr(IvaAnt), "N")
'
'                'aplicar el RECARGO a la base imponible de ese tipo
'                ImpREC = CalcularPorcentaje(ImpBImIVA, CCur(cadAux1), 2)
'
'                'sumamos todos los RECARGOS para sumarselo a la base imponible total de la factura
'                'los vamos acumulando
'                TotImpIVA = TotImpIVA + ImpREC
'            Else
'                cadAux1 = "0"
'                ImpREC = 0
'            End If
'
'
'            Select Case i
'                Case 1  'IVA 1
'                    TipoIVA1 = IvaAnt 'RS!codigiva
'
'                    BaseIVA1 = ImpBImIVA 'BASE IMPONIBLE
'
'                    PorceIVA1 = cadAux '% de IVA
'
'                    'Importe total con IVA
'                    ImpIVA1 = impiva
'
'                    PorceREC1 = cadAux1 '% de REC
'
'                    'Importe total con RECARGO
'                    ImpREC1 = ImpREC
'
'                Case 2  'IVA 2
'                    TipoIVA2 = IvaAnt 'RS!codigiva
'
'                    BaseIVA2 = ImpBImIVA 'BASE IMPONIBLE
'
'                    PorceIVA2 = cadAux '% de IVA
'
'                    'Importe total con IVA
'                    ImpIVA2 = impiva
'
'                    PorceREC2 = cadAux1 '% de REC
'
'                    'Importe total con RECARGO
'                    ImpREC2 = ImpREC
'                Case 3  'IVA 3
'                    TipoIVA3 = IvaAnt 'RS!codigiva
'
'                    BaseIVA3 = ImpBImIVA 'BASE IMPONIBLE
'
'                    PorceIVA3 = cadAux '% de IVA
'
'                    'Importe total con IVA
'                    ImpIVA3 = impiva
'
'                    PorceREC3 = cadAux1 '% de REC
'
'                    'Importe total con RECARGO
'                    ImpREC3 = ImpREC
'            End Select
'
'
'            i = i + 1
'            IvaAnt = Rs.Fields(0).Value
'            vBruto = DBLet(Rs.Fields(1).Value, "N")
''            vNeto = DBLet(Rs.Fields(2).Value, "N")
'        Else
'            vBruto = vBruto + DBLet(Rs.Fields(1).Value, "N")
''            vNeto = vNeto + DBLet(Rs.Fields(2).Value, "N")
'        End If
'
'
'        Rs.MoveNext
'    Wend
'    Rs.Close
'    Set Rs = Nothing
'
'    ' ULTIMO REGISTRO
'    TotBruto = TotBruto + vBruto
'    TotNeto = TotNeto + vNeto
'    ImpBImIVA = vNeto
'
'
'    'Obtener el % de IVA
'    cadAux = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(IvaAnt), "N")
'
'    'aplicar el IVA a la base imponible de ese tipo
'    '[Monica]23/08/2013: añado el comprobar cero
'    impiva = CalcularPorcentaje(ImpBImIVA, CCur(ComprobarCero(cadAux)), 2)
'
'    'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
'    'los vamos acumulando
'    TotImpIVA = TotImpIVA + impiva
'
'    If CInt(Data1.Recordset!TipoIvac) = 2 Then
'        'Obtener el % de RECARGO
'        cadAux1 = DevuelveDesdeBDNew(cConta, "tiposiva", "porcerec", "codigiva", CStr(IvaAnt), "N")
'
'        'aplicar el RECARGO a la base imponible de ese tipo
'        ImpREC = CalcularPorcentaje(ImpBImIVA, CCur(cadAux1), 2)
'    Else
'        cadAux1 = "0"
'        ImpREC = 0
'    End If
'    'sumamos todos los RECARGOS para sumarselo a la base imponible total de la factura
'    'los vamos acumulando
'    TotImpIVA = TotImpIVA + ImpREC
'
'
'
'    Select Case i
'        Case 1  'IVA 1
'            TipoIVA1 = IvaAnt
'
'            BaseIVA1 = ImpBImIVA 'BASE IMPONIBLE
'
'            PorceIVA1 = cadAux '% de IVA
'
'            'Importe total con IVA
'            ImpIVA1 = impiva
'
'            PorceREC1 = cadAux1 '% de REC
'
'            'Importe total con RECARGO
'            ImpREC1 = ImpREC
'
'        Case 2  'IVA 2
'            TipoIVA2 = IvaAnt
'
'            BaseIVA2 = ImpBImIVA 'BASE IMPONIBLE
'
'            PorceIVA2 = cadAux '% de IVA
'
'            'Importe total con IVA
'            ImpIVA2 = impiva
'
'            PorceREC2 = cadAux1 '% de REC
'
'            'Importe total con RECARGO
'            ImpREC2 = ImpREC
'        Case 3  'IVA 3
'            TipoIVA3 = IvaAnt
'
'            BaseIVA3 = ImpBImIVA 'BASE IMPONIBLE
'
'            PorceIVA3 = cadAux '% de IVA
'
'            'Importe total con IVA
'            ImpIVA3 = impiva
'
'            PorceREC3 = cadAux1 '% de REC
'
'            'Importe total con RECARGO
'            ImpREC3 = ImpREC
'    End Select
'
'    'Base Imponible
'    BaseImp = TotNeto
'
'    'TOTAL de la factura
'    TotalFac = BaseImp + TotImpIVA
'
'    'ACTUALIZAMOS LA FACTURA (tabla facturas)
'    Sql = "update facturascom "
'    Sql = Sql & "set baseimp1 = " & DBSet(BaseIVA1, "N")
'    Sql = Sql & ",impoiva1 = " & DBSet(ImpIVA1, "N")
'    Sql = Sql & ",imporec1 = " & DBSet(ImpREC1, "N")
'    Sql = Sql & ",porciva1 = " & DBSet(PorceIVA1, "N")
'    Sql = Sql & ",porcrec1 = " & DBSet(PorceREC1, "N")
'    Sql = Sql & ",codiiva1 = " & DBSet(TipoIVA1, "N")
'    Nulo2 = "N"
'    Nulo3 = "N"
'    If DBSet(TipoIVA2, "N", "S") = ValorNulo Then Nulo2 = "S"
'    If DBSet(TipoIVA3, "N", "S") = ValorNulo Then Nulo3 = "S"
'    Sql = Sql & ",baseimp2 = " & DBSet(BaseIVA2, "N", Nulo2)
'    Sql = Sql & ",impoiva2 = " & DBSet(ImpIVA2, "N", Nulo2)
'    Sql = Sql & ",imporec2 = " & DBSet(ImpREC2, "N", Nulo2)
'    Sql = Sql & ",porciva2 = " & DBSet(PorceIVA2, "N", Nulo2)
'    Sql = Sql & ",porcrec2 = " & DBSet(PorceREC2, "N", Nulo2)
'    Sql = Sql & ",codiiva2 = " & DBSet(TipoIVA2, "N", Nulo2)
'    Sql = Sql & ",baseimp3 = " & DBSet(BaseIVA3, "N", Nulo3)
'    Sql = Sql & ",impoiva3 = " & DBSet(ImpIVA3, "N", Nulo3)
'    Sql = Sql & ",imporec3 = " & DBSet(ImpREC3, "N", Nulo3)
'    Sql = Sql & ",porciva3 = " & DBSet(PorceIVA3, "N", Nulo3)
'    Sql = Sql & ",porcrec3 = " & DBSet(PorceREC3, "N", Nulo3)
'    Sql = Sql & ",codiiva3 = " & DBSet(TipoIVA3, "N", Nulo3)
'    Sql = Sql & ",brutofac = " & DBSet(TotBruto, "N")
'    Sql = Sql & ",impordto = " & DBSet(Round2(TotBruto - TotNeto, 2), "N")
'    Sql = Sql & ",totalfac = " & DBSet(TotalFac, "N")
'    Sql = Sql & " where " & cadwhere
'
'    conn.Execute Sql
'
'    CalcularDatosFacturaCompra_old = True
'
'ECalcular:
'    If Err.Number <> 0 Then
'        CalcularDatosFacturaCompra_old = False
'    Else
'        CalcularDatosFacturaCompra_old = True
'    End If
'End Function

Public Function CalcularDatosFacturaCompra(cadwhere As String, NomTabla As String, NomTablaLin As String) As Boolean
'cadWhere: cad para la where de la SQL que selecciona las lineas del albaran o la factura
'nomTabla: nombre de la tabla de albaranes(scaalp) o de AlbaranesXFactura(scafpa)
'           segun llamemos desde recepcion de facturas o desde Hco de Facturas
Dim Rs As ADODB.Recordset
Dim i As Integer

Dim Sql As String
Dim cadAux As String
Dim cadwhere1 As String

'Aqui vamos acumulando los totales
Dim TotBruto As Currency
Dim TotImpIVA As Currency

Dim ImpAux As Currency
Dim impiva As Currency
Dim ImpBImIVA As Currency 'Importe Base imponible a la que hay q aplicar el IVA

Dim DtoPPago As Currency
Dim DtoGnral As Currency
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

Dim ImpPPago As Currency
Dim ImpGnral As Currency

Dim TotalFac As Currency
Dim Nulo2 As String
Dim Nulo3 As String


    CalcularDatosFacturaCompra = False
    On Error GoTo ECalcular

   
    DtoPPago = ComprobarCero(Text1(7))
    DtoGnral = ComprobarCero(Text1(8))
   
    cadwhere1 = Replace(cadwhere, "facturascom", "facturascom_variedad")
    Sql = "SELECT facturascom_variedad.codigiva, sum(importe) as bruto"
    Sql = Sql & " FROM facturascom_variedad "
    Sql = Sql & " WHERE " & cadwhere1
    Sql = Sql & " GROUP BY 1 "
    Sql = Sql & " ORDER BY 1 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    TotBruto = 0
    i = 1
    While Not Rs.EOF
        'Aqui vamos acumulando la suma del importe bruto de las lineas
        TotBruto = TotBruto + Rs!Bruto
        ImpBImIVA = Rs!Bruto
        
        'Aplicarle el dto ppago
'        ImpAux = CCur(CalcularDto(CStr(RS!bruto), CStr(DtoPPago)))
'        '---- Laura: 27/09/2006
'        ImpAux = Round(ImpAux, 2)
'        '----
        '---- Laura: 24/10/2006
        ImpAux = CalcularPorcentaje(Rs!Bruto, DtoPPago, 2)
        ImpBImIVA = ImpBImIVA - ImpAux '(bruto - DtoPP)
        
        'Aplicarle el dto grnal
'        ImpAux = CCur(CalcularDto(CStr(RS!bruto), CStr(DtoGnral)))
'        '---- Laura: 27/09/2006
'        ImpAux = Round(ImpAux, 2)
'        '----
        '---- Laura: 24/10/2006
        ImpAux = CalcularPorcentaje(Rs!Bruto, DtoGnral, 2)
        ImpBImIVA = ImpBImIVA - ImpAux '(bruto - Dtogn)
        
        'Obtener el % de IVA
        If vParamAplic.NumeroConta <> 0 Then
            cadAux = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(Rs!Codigiva), "N")
        Else
            cadAux = DevuelveDesdeBDNew(cAgro, "tiposiva", "porceiva", "codigiva", CStr(Rs!Codigiva), "N")
        End If
        
        'aplicar el IVA a la base imponible de ese tipo
'        ImpAux = CalcularDto(CStr(ImpBImIVA), cadAux)
'        '---- Laura: modificado 27/09/2006
''        ImpIVA = ImpAux
'        ImpIVA = Round(ImpAux, 2)
        '----
        '---- Laura: 24/10/2006
        If cadAux = "" Then cadAux = "0"
        impiva = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
        
        'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
        'los vamos acumulando
        TotImpIVA = TotImpIVA + impiva
        
        
        Select Case i
            Case 1  'IVA 1
                TipoIVA1 = Rs!Codigiva
                
                BaseIVA1 = ImpBImIVA 'BASE IMPONIBLE
                
                PorceIVA1 = cadAux '% de IVA
                
                'Importe total con IVA
                ImpIVA1 = impiva
                
            Case 2  'IVA 2
                TipoIVA2 = Rs!Codigiva
                
                BaseIVA2 = ImpBImIVA 'BASE IMPONIBLE
                
                PorceIVA2 = cadAux '% de IVA
                
                'Importe total con IVA
                ImpIVA2 = impiva

            Case 3  'IVA 3
                TipoIVA3 = Rs!Codigiva
                
                BaseIVA3 = ImpBImIVA 'BASE IMPONIBLE
                
                PorceIVA3 = cadAux '% de IVA
                
                'Importe total con IVA
                ImpIVA3 = impiva
        End Select
        i = i + 1
        Rs.MoveNext
    Wend
    
    Rs.Close
    Set Rs = Nothing

    'TOTALES
    BrutoFac = TotBruto
    
    'Aplicarle el dto ppago
'    ImpPPago = CCur(CalcularDto(CStr(TotBruto), CStr(DtoPPago)))
'    '---- Laura: 27/09/2006
'    ImpPPago = Round(ImpPPago, 2)
'    '----
    '---- Laura: 24/10/2006
    ImpPPago = CalcularPorcentaje(TotBruto, DtoPPago, 2)
    '----
    
    'Aplicarle el dto general
'    ImpGnral = CCur(CalcularDto(CStr(TotBruto), CStr(DtoGnral)))
'    '---- Laura: 27/09/2006
'    ImpGnral = Round(ImpGnral, 2)
'    '----
    '---- Laura: 24/10/2006
    ImpGnral = CalcularPorcentaje(TotBruto, DtoGnral, 2)
    '----
    
    'Base Imponible
    BaseImp = TotBruto - ImpPPago - ImpGnral
    
    'TOTAL de la factura
    TotalFac = BaseImp + TotImpIVA
    
    'ACTUALIZAMOS LA FACTURA (tabla facturas)
    Sql = "update facturascom "
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
    Sql = Sql & ",impppago = " & DBSet(Round2(ImpPPago), "N")
    Sql = Sql & ",impgnral = " & DBSet(Round2(ImpGnral), "N")
    Sql = Sql & ",totalfac = " & DBSet(TotalFac, "N")
    Sql = Sql & " where " & cadwhere
    
    conn.Execute Sql

    
    
    
    CalcularDatosFacturaCompra = True
    
ECalcular:
    If Err.Number <> 0 Then
        CalcularDatosFacturaCompra = False
    Else
        CalcularDatosFacturaCompra = True
    End If
End Function


Private Function FactContabilizada(ByRef EstaEnTesoreria As String) As Boolean
Dim Codmacta As String, numasien As String
    
    On Error GoTo EContab
    
    'NO deberia poder modificar fras anteriors a fecha inicio ejercicio
    'Cojo la letra de serie
    Codmacta = DevuelveDesdeBDNew(cAgro, "proveedor", "codmacta", "codprove", Text1(6).Text)
    
    'Primero comprobaremos que esta el cobro en contabilidad
    EstaEnTesoreria = ""
    If Not ComprobarPagoArimoney(EstaEnTesoreria, Codmacta, CLng(Text1(0).Text), CDate(Text1(1).Text)) Then
        FactContabilizada = True
        Exit Function
    End If

    'comprobar que se puede modificar/eliminar la factura
    If Me.Check1(0).Value = 1 Then 'si esta contabilizada
        'comprobar en la contabilidad si esta contabilizada
        If Codmacta <> "" Then
            If vParamAplic.ContabilidadNueva Then
                numasien = DevuelveDesdeBDNew(cConta, "factpro", "numasien", "codmacta", Codmacta, "T", , "numfactu", Text1(0).Text, "N", "anofactu", Year(Text1(1).Text), "N")
            Else
                numasien = DevuelveDesdeBDNew(cConta, "cabfactpro", "numasien", "codmacta", Codmacta, "T", , "codfacpr", Text1(0).Text, "N", "anofacpr", Year(Text1(1).Text), "N")
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
        
        Codmacta = "La factura esta en la contabilidad"
        If numasien <> "" Then Codmacta = Codmacta & vbCrLf & "Nº asiento: " & numasien
        Codmacta = Codmacta & vbCrLf & vbCrLf & "¿Continuar?"
        
        numasien = String(50, "*") & vbCrLf
        numasien = numasien & numasien & vbCrLf & vbCrLf
        Codmacta = numasien & Codmacta & vbCrLf & vbCrLf & numasien
        If MsgBox(Codmacta, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
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
Private Function ComprobarPagoArimoney(vTesoreria As String, LEtra As String, Codfacpr As Long, fecha As Date) As Boolean
Dim vR As ADODB.Recordset
Dim Cad As String

On Error GoTo EComprobarPagoArimoney
    ComprobarPagoArimoney = False
    Set vR = New ADODB.Recordset
    
    If vParamAplic.ContabilidadNueva Then
        Cad = "Select * from pagos where codmacta='" & LEtra & "'"
        Cad = Cad & " AND numfactu =" & Codfacpr
        Cad = Cad & " AND fecfactu =" & DBSet(fecha, "F")
    
    Else
        Cad = "Select * from spagop where codmacta='" & LEtra & "'"
        Cad = Cad & " AND codfacpr =" & Codfacpr
        Cad = Cad & " AND fecfacpr =" & DBSet(fecha, "F")
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
            If DBLet(vR!recedocu, "N") = 1 Then
                Cad = "Documento recibido"
            Else
                If vParamAplic.ContabilidadNueva Then
                    If DBLet(vR!transfer, "N") = 1 Then
                        Cad = "Esta en una transferencia"
                    Else
                       If DBLet(vR!ImpPago, "N") > 0 Then Cad = "Esta parcialmente pagado: " & vR!ImpPago
                    
                        
                                'Si hubeira que poner mas coas iria aqui
                    End If 'transfer
                
                Else
                    If DBLet(vR!Estacaja, "N") = 1 Then
                        Cad = "Pagado por caja"
                    Else
                        If DBLet(vR!transfer, "N") = 1 Then
                            Cad = "Esta en una transferencia"
                        Else
                           If DBLet(vR!ImpPago, "N") > 0 Then Cad = "Esta parcialmente pagado: " & vR!ImpPago
                        
                            
                                    'Si hubeira que poner mas coas iria aqui
                        End If 'transfer
                    End If 'estacaja
                End If
            End If 'recdedocu
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
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then ComprobarPagoArimoney = True
        End If
    Else
        ComprobarPagoArimoney = True
    End If
            
EComprobarPagoArimoney:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set vR = Nothing
End Function


