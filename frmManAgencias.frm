VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmManAgencias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Agencias de Transporte / Comisionistas"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   14865
   Icon            =   "frmManAgencias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   14865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   90
      TabIndex        =   71
      Top             =   90
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   72
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
      TabIndex        =   69
      Top             =   90
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   70
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
      Left            =   12285
      TabIndex        =   68
      Top             =   285
      Width           =   1605
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos Contabilidad"
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
      Height          =   2655
      Index           =   4
      Left            =   6750
      TabIndex        =   58
      Top             =   1890
      Width           =   7995
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
         Index           =   27
         Left            =   1755
         MaxLength       =   4
         TabIndex        =   15
         Tag             =   "IBAN|T|S|||agencias|iban|||"
         Text            =   "Text1"
         Top             =   360
         Width           =   615
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
         Left            =   1755
         MaxLength       =   5
         TabIndex        =   23
         Tag             =   "Porc.Retencion|N|S|0|100|agencias|porcereten|##0.00||"
         Top             =   2160
         Width           =   660
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
         Index           =   20
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   62
         Text            =   "0"
         Top             =   1260
         Width           =   5295
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
         Index           =   25
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   61
         Text            =   "Text2"
         Top             =   810
         Width           =   5295
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
         Left            =   1740
         MaxLength       =   3
         TabIndex        =   20
         Tag             =   "Forma Pago|N|N|0|999|agencias|codforpa|000|N|"
         Text            =   "Text1"
         Top             =   810
         Width           =   660
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
         Index           =   24
         Left            =   4650
         MaxLength       =   10
         TabIndex        =   19
         Tag             =   "Cuenta Bancaria|T|S|||agencias|cuentaba|0000000000||"
         Text            =   "1234567890"
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
         Index           =   23
         Left            =   4005
         MaxLength       =   2
         TabIndex        =   18
         Tag             =   "Digito Control|T|S|||agencias|digcontr|00||"
         Text            =   "Text1"
         Top             =   360
         Width           =   495
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
         Index           =   22
         Left            =   3270
         MaxLength       =   4
         TabIndex        =   17
         Tag             =   "Sucursal|N|S|0|9999|agencias|codsucur|0000||"
         Text            =   "Text1"
         Top             =   360
         Width           =   615
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
         Index           =   21
         Left            =   2505
         MaxLength       =   4
         TabIndex        =   16
         Tag             =   "Banco|N|S|0|9999|agencias|codbanco|0000||"
         Text            =   "Text1"
         Top             =   360
         Width           =   615
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
         Left            =   1755
         MaxLength       =   4
         TabIndex        =   21
         Tag             =   "Banco Propio|N|N|0|9999|agencias|codbanpr|0000||"
         Text            =   "Text1"
         Top             =   1260
         Width           =   660
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
         Index           =   19
         Left            =   3150
         TabIndex        =   59
         Top             =   1710
         Width           =   4650
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
         Index           =   19
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   22
         Tag             =   "Cta.Contable|T|S|||agencias|codmacta|||"
         Text            =   "0000000000"
         Top             =   1710
         Width           =   1290
      End
      Begin VB.Label Label2 
         Caption         =   "%Retencion"
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
         Left            =   135
         TabIndex        =   66
         Top             =   2205
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "IBAN Agencia"
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
         Index           =   21
         Left            =   135
         TabIndex        =   65
         Top             =   405
         Width           =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "Banco Propio"
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
         Index           =   19
         Left            =   135
         TabIndex        =   64
         Top             =   1305
         Width           =   1305
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1485
         ToolTipText     =   "Buscar banco propio"
         Top             =   1305
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1485
         ToolTipText     =   "Buscar forma de pago"
         Top             =   855
         Width           =   240
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
         Index           =   14
         Left            =   135
         TabIndex        =   63
         Top             =   855
         Width           =   1185
      End
      Begin VB.Label Label20 
         Caption         =   "Cta.Contable"
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
         Left            =   135
         TabIndex        =   60
         Top             =   1755
         Width           =   1320
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1485
         ToolTipText     =   "Buscar Cta.Contable"
         Top             =   1755
         Width           =   240
      End
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
      Height          =   720
      Index           =   18
      Left            =   6795
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   24
      Tag             =   "Observaciones|T|S|||agencias|obstrans|||"
      Top             =   4905
      Width           =   7905
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos Administraci�n"
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
      Height          =   2430
      Index           =   3
      Left            =   6750
      TabIndex        =   50
      Top             =   5715
      Width           =   7950
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
         Index           =   17
         Left            =   225
         MaxLength       =   10
         TabIndex        =   25
         Tag             =   "Tel�fono|T|S|||agencias|teltrans1|||"
         Top             =   630
         Width           =   1695
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
         Index           =   16
         Left            =   4095
         MaxLength       =   10
         TabIndex        =   27
         Tag             =   "Fax|T|S|||agencias|faxtrans1|||"
         Top             =   630
         Width           =   1785
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
         Index           =   15
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   29
         Tag             =   "E-mail|T|S|||agencias|maitrans1|||"
         Top             =   1710
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
         Height          =   360
         Index           =   14
         Left            =   2070
         MaxLength       =   35
         TabIndex        =   28
         Tag             =   "Persona de Contacto|T|S|||agencias|pertrans1|||"
         Top             =   1170
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
         Height          =   360
         Index           =   13
         Left            =   2070
         MaxLength       =   10
         TabIndex        =   26
         Tag             =   "M�vil|T|S|||agencias|movtrans1|||"
         Top             =   630
         Width           =   1650
      End
      Begin VB.Label Label1 
         Caption         =   "Fax"
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
         Left            =   4095
         TabIndex        =   55
         Top             =   360
         Width           =   1185
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
         Index           =   17
         Left            =   225
         TabIndex        =   54
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "E-mail"
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
         Left            =   225
         TabIndex        =   53
         Top             =   1710
         Width           =   915
      End
      Begin VB.Image imgMail 
         Height          =   240
         Index           =   1
         Left            =   1710
         Top             =   1755
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Persona Contacto"
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
         TabIndex        =   52
         Top             =   1170
         Width           =   1770
      End
      Begin VB.Label Label1 
         Caption         =   "M�vil"
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
         Left            =   2070
         TabIndex        =   51
         Top             =   360
         Width           =   1185
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos b�sicos"
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
      Height          =   3750
      Index           =   1
      Left            =   90
      TabIndex        =   44
      Top             =   1890
      Width           =   6630
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
         Index           =   10
         Left            =   1305
         MaxLength       =   40
         TabIndex        =   9
         Tag             =   "Web|T|S|||agencias|wwwtrans|||"
         Top             =   3015
         Width           =   4935
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
         Index           =   2
         Left            =   1320
         MaxLength       =   9
         TabIndex        =   4
         Tag             =   "NIF|T|N|||agencias|ciftrans|||"
         Top             =   360
         Width           =   1335
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
         Left            =   1320
         MaxLength       =   35
         TabIndex        =   5
         Tag             =   "Domicilio|T|S|||agencias|domtrans|||"
         Top             =   885
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
         Height          =   360
         Index           =   5
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   7
         Tag             =   "Poblaci�n|T|S|||agencias|pobtrans|||"
         Top             =   1950
         Width           =   4935
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
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   6
         Tag             =   "C�digo Postal|T|S|||agencias|codpobla|||"
         Top             =   1425
         Width           =   735
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
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   8
         Tag             =   "Provincia|T|S|||agencias|protrans|||"
         Top             =   2490
         Width           =   4935
      End
      Begin VB.Image imgWeb 
         Height          =   255
         Left            =   765
         Picture         =   "frmManAgencias.frx":000C
         Stretch         =   -1  'True
         Tag             =   "-1"
         ToolTipText     =   "Abrir web"
         Top             =   3060
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Web"
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
         Left            =   180
         TabIndex        =   57
         Top             =   3060
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "N.I.F."
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
         Left            =   195
         TabIndex        =   49
         Top             =   360
         Width           =   495
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
         Index           =   2
         Left            =   195
         TabIndex        =   48
         Top             =   900
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
         Index           =   4
         Left            =   180
         TabIndex        =   47
         Top             =   1980
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "C.P."
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
         Left            =   195
         TabIndex        =   46
         Top             =   1440
         Width           =   495
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
         Index           =   8
         Left            =   180
         TabIndex        =   45
         Top             =   2520
         Width           =   1050
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos Comercial"
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
      Height          =   2430
      Index           =   2
      Left            =   90
      TabIndex        =   38
      Top             =   5715
      Width           =   6600
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
         Index           =   12
         Left            =   2070
         MaxLength       =   10
         TabIndex        =   11
         Tag             =   "M�vil|T|S|||agencias|movtrans|||"
         Top             =   630
         Width           =   1650
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
         Index           =   11
         Left            =   2070
         MaxLength       =   35
         TabIndex        =   13
         Tag             =   "Persona de Contactol|T|S|||agencias|pertrans|||"
         Top             =   1170
         Width           =   4170
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
         Left            =   2070
         MaxLength       =   40
         TabIndex        =   14
         Tag             =   "E-mail|T|S|||agencias|maitrans|||"
         Top             =   1710
         Width           =   4170
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
         Left            =   3915
         MaxLength       =   10
         TabIndex        =   12
         Tag             =   "Fax|T|S|||agencias|faxtrans|||"
         Top             =   630
         Width           =   1785
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
         Left            =   225
         MaxLength       =   10
         TabIndex        =   10
         Tag             =   "Tel�fono|T|S|||agencias|teltrans|||"
         Top             =   630
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "M�vil"
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
         Left            =   2070
         TabIndex        =   43
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Persona Contacto"
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
         Left            =   225
         TabIndex        =   42
         Top             =   1170
         Width           =   1770
      End
      Begin VB.Image imgMail 
         Height          =   240
         Index           =   0
         Left            =   1665
         Top             =   1755
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "E-mail"
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
         Left            =   225
         TabIndex        =   41
         Top             =   1710
         Width           =   915
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
         Index           =   13
         Left            =   225
         TabIndex        =   40
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label Label1 
         Caption         =   "Fax"
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
         Left            =   3915
         TabIndex        =   39
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Index           =   0
      Left            =   75
      TabIndex        =   35
      Top             =   855
      Width           =   14670
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
         ItemData        =   "frmManAgencias.frx":0596
         Left            =   9180
         List            =   "frmManAgencias.frx":0598
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "Tipo|N|S|0||agencias|codfortrans|||"
         Top             =   450
         Width           =   3480
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
         ItemData        =   "frmManAgencias.frx":059A
         Left            =   6750
         List            =   "frmManAgencias.frx":059C
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Tag             =   "Tipo|N|N|0|1|agencias|tipo|||"
         Top             =   450
         Width           =   1995
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
         Index           =   1
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   1
         Tag             =   "Nombre|T|N|||agencias|nomtrans|||"
         Top             =   450
         Width           =   4965
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
         Index           =   0
         Left            =   240
         MaxLength       =   3
         TabIndex        =   0
         Tag             =   "C�digo Agencia|N|N|0|999|agencias|codtrans|000|S|"
         Top             =   450
         Width           =   865
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Transporte"
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
         Left            =   9180
         TabIndex        =   74
         Top             =   180
         Width           =   2445
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo"
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
         Left            =   6750
         TabIndex        =   67
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "C�digo"
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
         Left            =   240
         TabIndex        =   37
         Top             =   195
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre "
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
         Left            =   1320
         TabIndex        =   36
         Top             =   195
         Width           =   735
      End
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
      Left            =   12465
      TabIndex        =   30
      Top             =   8370
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
      Left            =   13635
      TabIndex        =   31
      Top             =   8370
      Visible         =   0   'False
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
      Left            =   13635
      TabIndex        =   34
      Top             =   8370
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   90
      TabIndex        =   32
      Top             =   8190
      Width           =   2385
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
         TabIndex        =   33
         Top             =   195
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   330
      Left            =   4440
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Left            =   14325
      TabIndex        =   73
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
      Left            =   6795
      TabIndex        =   56
      Top             =   4590
      Width           =   1530
   End
   Begin VB.Image imgZoom 
      Height          =   240
      Index           =   0
      Left            =   8370
      ToolTipText     =   "Zoom descripci�n"
      Top             =   4590
      Width           =   240
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
Attribute VB_Name = "frmManAgencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MANOLO   +-+-
' +-+- Fecha: 23/05/06 +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-+

Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'codi per al registe que s'afegix al cridar des d'altre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String
Public DeConsulta As Boolean
Public CodigoActual As String

Private WithEvents frmCtas As frmCtasConta 'cuenta contable
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmFPa As frmManFpago ' formas de pago
Attribute frmFPa.VB_VarHelpID = -1
Private WithEvents frmBan As frmManBanco 'Banco Propio
Attribute frmBan.VB_VarHelpID = -1

Private HaDevueltoDatos As Boolean
Private CadenaSelect As String
Private CadenaConsulta As String
Private CadB As String

Private WithEvents frmAge As frmBasico2
Attribute frmAge.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1

Dim Modo As Byte
'-------------- MODOS ---------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la b�squeda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edici� del camp
'   3.-  Inserci� de nou registre
'   4.-  Modificar
'------------------------------------------------
Dim FormatoCod As String 'formato del campo c�digo
Dim NomTabla As String
Dim Ordenacion As String

Dim btnPrimero As Byte 'Variable que indica el n� del Bot� PrimerRegistro en la Toolbar1
Dim indice As Byte 'Indice del text1 donde se ponen los datos devueltos desde otros Formularios de Mtos

Private BuscaChekc As String

'Cambio en cuentas de la contabilidad
Dim IbanAnt As String
Dim NombreAnt As String
Dim NomComerAnt As String
Dim BancoAnt  As String
Dim SucurAnt As String
Dim DigitoAnt As String
Dim CuentaAnt As String

Dim DirecAnt As String
Dim cPostalAnt As String
Dim PoblaAnt As String
Dim ProviAnt As String
Dim NifAnt As String
Dim forpaant As String


Dim EMaiAnt As String
Dim WebAnt As String

Dim CtaBancoAnt As String




Private Sub PonerModo(vModo)
Dim b As Boolean
Dim NumReg As Byte

    On Error GoTo EPonerModo
    
    Modo = vModo
    If Modo = 2 Then
        lblIndicador.Caption = PonerContRegistros(Me.adodc1)
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    
    b = (Modo = 2)
    
    '=======================================
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Me.adodc1.Recordset.EOF Then
        If adodc1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, Numreg
    DesplazamientoVisible b And adodc1.Recordset.RecordCount > 1

     '---------------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = (Modo = 2)
    Else
        cmdRegresar.visible = False
    End If
    
    
    BloquearText1 Me, Modo
    
    ' **** si n'hi han imagens de buscar en la cap�alera *****
    BloquearImgBuscar Me, Modo
    BloquearImgZoom Me, Modo
    ' ********************************************************
    BloquearCombo Me, Modo
    
    
    'Si es regresar
'    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = b
    
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    PonerLongCampos 'Pone el Maxlength de los campos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu 'Activar/Desact botones de menu segun permisos del usuario
    
EPonerModo:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner modo.", Err.Description
End Sub

Private Sub DesplazamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub

Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botones del la toobar y del menu, segun el modo en que estemos
Dim b As Boolean

    b = (Modo = 2) Or Modo = 0
    'Busqueda
    Toolbar1.Buttons(5).Enabled = b
    Me.mnBuscar.Enabled = b
    'Ver Todos
    Toolbar1.Buttons(6).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    'Insertar
    Toolbar1.Buttons(1).Enabled = b And Not DeConsulta
    Me.mnNuevo.Enabled = b And Not DeConsulta
    
    b = (Modo = 2 And adodc1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnModificar.Enabled = b
    'Eliminar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnEliminar.Enabled = b
    'Imprimir
    Toolbar1.Buttons(8).Enabled = b
    Me.mnImprimir.Enabled = b
    
End Sub

Private Sub BotonAnyadir()
Dim NumF As String
    
    LimpiarCampos 'Vac�a los TextBox
    CadB = ""
    
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    
     '******************** canviar taula i camp **************************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        NumF = NuevoCodigo
    Else
        NumF = SugerirCodigoSiguienteStr("agencias", "codtrans")
    End If
    
    ' ******* Canviar el nom de la taula, el nom de la clau primaria, i el
    ' nom del camp que te la clau primaria si no es Text1(0) *************
    Text1(0).Text = NumF
    FormateaCampo Text1(0)
    
    'PosarDescripcions
    PonerFoco Text1(1)
    ' ********************************************************************
End Sub

Private Sub BotonVerTodos()
    CadB = ""
    LimpiarCampos 'Limpia los Text1
    
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NomTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
Dim Cad As String

    Set frmAge = New frmBasico2
    
    AyudaAgenciasPrev frmAge, , CadB
    
    Set frmAge = Nothing

End Sub

Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    Me.adodc1.RecordSource = CadenaConsulta
    adodc1.Refresh
    If adodc1.Recordset.RecordCount <= 0 Then
        If CadB = "" Then
            MsgBox "No hay ning�n registro en la tabla " & NomTabla, vbInformation
'            Screen.MousePointer = vbDefault
'            Exit Sub
        Else
            If Modo = 1 Then MsgBox "Ning�n registro encontrado para el criterio de b�squeda.", vbInformation
            PonerFoco Text1(indice)
        End If
    Else
        PonerModo 2
        'Data1.Recordset.MoveLast
        adodc1.Recordset.MoveFirst
        PonerCampos
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub

EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub BotonBuscar()
   If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        PonerFoco Text1(0)
'        PosicionarCombo Combo1(0), 754
        Text1(0).BackColor = vbLightBlue
    End If
End Sub

Private Sub BotonModificar()
    
    NombreAnt = Text1(1).Text
    IbanAnt = Text1(27).Text
    BancoAnt = Text1(21).Text
    SucurAnt = Text1(22).Text
    DigitoAnt = Text1(23).Text
    CuentaAnt = Text1(24).Text
    
    DirecAnt = Text1(3).Text
    cPostalAnt = Text1(4).Text
    PoblaAnt = Text1(5).Text
    ProviAnt = Text1(6).Text
    NifAnt = Text1(2).Text
    
    EMaiAnt = Text1(15).Text
    WebAnt = Text1(10).Text
    
    CtaBancoAnt = DevuelveValor("select codmacta from banpropi where codbanpr = " & DBSet(Text1(20).Text, "N"))
    
    forpaant = Text1(25).Text
    
    
    PonerModo 4
   
    'Como es modificar
    ' *** primer control que no siga clau primaria ***
    PonerFoco Text1(1)
    ' ************************************************
    Screen.MousePointer = vbDefault
End Sub

Private Sub BotonEliminar()
Dim Sql As String

    On Error GoTo EEliminar
    
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registro de codigo 0 no se puede Modificar ni Eliminar
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCod) Then Exit Sub
    
    If Not SepuedeBorrar Then Exit Sub
    
    '*************** canviar els noms i el DELETE **********************************
    Sql = "�Seguro que desea eliminar la Agencia?"
    Sql = Sql & vbCrLf & "C�digo: " & Text1(0).Text
    Sql = Sql & vbCrLf & "Nombre: " & adodc1.Recordset.Fields(1)
    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = adodc1.Recordset.AbsolutePosition
        
        Sql = "Delete from " & NomTabla & " where codtrans=" & adodc1.Recordset!codTrans
        conn.Execute Sql
        
        If SituarDataTrasEliminar(adodc1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
    
EEliminar:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    PonerLongCamposGnral Me, Modo, 1
End Sub

Private Sub cmdAceptar_Click()

    Select Case Modo
         Case 1  'BUSQUEDA
            HacerBusqueda
    
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    Text2(19).Text = PonerNombreCuenta(Text1(19), Modo, Text1(0).Text)
                
                    CadenaConsulta = "select * from " & NomTabla
                    CadenaConsulta = CadenaConsulta & " WHERE codtrans=" & Text1(0).Text
                    CadenaConsulta = CadenaConsulta & Ordenacion
                    Me.adodc1.RecordSource = CadenaConsulta '"Select * from " & NomTabla & Ordenacion
                    Me.adodc1.Refresh
                    PosicionarData
                End If
            End If
        
        Case 4 'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me) Then
                    TerminaBloquear
                    
                    '[Monica]28/10/2016: Si han cambiado nombre o CCC pregunto si quieren cambiar los datos de la cuenta en la seccion de horto
                    ModificarDatosCuentaContable
                    
                    
                    PosicionarData
                End If
            End If
    End Select
End Sub

Private Sub cmdCancelar_Click()
    On Error Resume Next

    Select Case Modo
        Case 1, 3 'Busqueda, Insertar
            LimpiarCampos
            If Me.adodc1.Recordset.EOF Then
                PonerModo 0
            Else
                PonerModo 2
                PonerCampos
            End If
            PonerFoco Text1(0)

        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
    End Select

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String
Dim i As Integer
Dim J As Integer
Dim Aux As String

    If adodc1.Recordset.EOF Then
        MsgBox "Ning�n registro devuelto.", vbExclamation
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

Private Sub Combo1_Validate(Index As Integer, Cancel As Boolean)
    '[Monica]25/04/2019: si es un transportista hemos de poner que forma de transporte tiene
    If Index = 0 Then
        Combo1(1).Enabled = (Combo1(0).ListIndex = 0)
        If Not Combo1(1).Enabled Then Combo1(1).ListIndex = -1
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim i As Integer

    ' ICONITOS DE LA BARRA
'    btnPrimero = 15 'index del bot� "primero"
'    With Me.Toolbar1
'        .HotImageList = frmPpal.imgListComun_OM
'        .DisabledImageList = frmPpal.imgListComun_BN
'        .ImageList = frmPpal.imgListComun
'        'el 1 es separadors
'        .Buttons(2).Image = 1   'Buscar
'        .Buttons(3).Image = 2   'Todos
'        'el 4 i el 5 son separadors
'        .Buttons(6).Image = 3   'Insertar
'        .Buttons(7).Image = 4   'Modificar
'        .Buttons(8).Image = 5   'Borrar
'        'el 9 i el 10 son separadors
'        .Buttons(11).Image = 10  'Imprimir
'        .Buttons(13).Image = 11  'Salir
'        '14 y 15 separadors
'        .Buttons(btnPrimero).Image = 6  'Primero
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
'        .Buttons(btnPrimero + 3).Image = 9 '�ltimo
'    End With
    
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
    
    
    'cargar IMAGES de busqueda
    For i = 0 To imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
 
    'cargar IMAGE de mail
    Me.ImgMail(0).Picture = frmPpal.imgListImages16.ListImages(2).Picture
    Me.ImgMail(1).Picture = frmPpal.imgListImages16.ListImages(2).Picture
    
    'IMAGES para zoom
    Me.imgZoom(0).Picture = frmPpal.imgListImages16.ListImages(3).Picture

    LimpiarCampos   'Limpia los campos TextBox
    
    CargaCombo
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)

    '****************** canviar la consulta *********************************+
    NomTabla = "agencias"
    Ordenacion = " ORDER BY codtrans"
    CadenaConsulta = "select * from " & NomTabla
    
    Me.adodc1.ConnectionString = conn
    Me.adodc1.RecordSource = CadenaConsulta & " where codtrans=-1"
    Me.adodc1.Refresh
    
    CadB = ""

    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1
        Text1(0).BackColor = vbYellow 'codclien
    End If
    
        

End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    If Modo = 4 Then TerminaBloquear
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim Aux As String

    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadB = Aux
        '   Como la clave principal es unica, con poner el sql apuntando
        '   al valor devuelto sobre la clave ppal es suficiente
        ' *** canviar o llevar el WHERE ***
        CadenaConsulta = "select * from " & NomTabla & " WHERE " & CadB & " " & Ordenacion
        ' **********************************
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmAge_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String
Dim Aux As String
      
    If CadenaSeleccion <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        CadB = "codtrans = '" & RecuperaValor(CadenaSeleccion, 1) & "'"
        CadenaConsulta = "select * from agencias WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
    Screen.MousePointer = vbDefault


End Sub

Private Sub frmBan_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento Bancos Propios
    Text1(20).Text = RecuperaValor(CadenaSeleccion, 1) 'codbanpr
    FormateaCampo Text1(20)
    Text2(20).Text = RecuperaValor(CadenaSeleccion, 2) 'nombanpr
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    FormateaCampo Text1(indice)
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nommacta
End Sub

Private Sub frmFpa_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento F.Pago
    Text1(25).Text = RecuperaValor(CadenaSeleccion, 1) 'codforpa
    FormateaCampo Text1(25)
    Text2(25).Text = RecuperaValor(CadenaSeleccion, 2) 'nomforpa
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(indice).Text = vCampo
End Sub

Private Sub imgBuscar_Click(Index As Integer)
    TerminaBloquear
    
    Select Case Index
        Case 0 'Cuentas Contables (de contabilidad)
            If vParamAplic.NumeroConta = 0 Then Exit Sub
            
            indice = Index + 19
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(indice).Text
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(indice)
        
        Case 1 'formas de pago
            Set frmFPa = New frmManFpago
            frmFPa.DatosADevolverBusqueda = "0|1|"
            frmFPa.CodigoActual = Text1(25).Text
            frmFPa.Show vbModal
            Set frmFPa = Nothing
            PonerFoco Text1(25)
        
        Case 2 'banco propio
            Set frmBan = New frmManBanco
            frmBan.DatosADevolverBusqueda = "0|1|"
            frmBan.CodigoActual = Text1(20).Text
            frmBan.Show vbModal
            Set frmBan = Nothing
            PonerFoco Text1(20)
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, adodc1, 1
End Sub

Private Sub imgMail_Click(Index As Integer)
    Select Case Index
        Case 0
            If Text1(9).Text <> "" Then
                LanzaMailGnral Text1(9).Text
            End If
        Case 1
            If Text1(15).Text <> "" Then
                LanzaMailGnral Text1(15).Text
            End If
    End Select
End Sub

Private Sub imgWeb_Click()
    'Abrimos el explorador de windows con la pagina Web del cliente
    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    If LanzaHomeGnral(Text1(10).Text) Then espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        indice = 18
        frmZ.pTitulo = "Observaciones de la Agencia"
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
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    printNou
End Sub

Private Sub mnModificar_Click()
    'Comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
    
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registro de codigo 0 no se puede Modificar ni Eliminar
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCod) Then Exit Sub

    'Preparar para modificar
    '-----------------------
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

Private Sub Text1_GotFocus(Index As Integer)
    indice = Index
    ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 18 Or (Index = 18 And Text1(18).Text = "") Then KEYpress KeyAscii
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim Nuevo As Boolean
Dim cadMen As String

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0 'codigo trabajador
            PonerFormatoEntero Text1(0)
        
        Case 1 'NOMBRE
            Text1(Index).Text = UCase(Text1(Index).Text)
        
        Case 2 'NIF
            Text1(Index).Text = UCase(Text1(Index).Text)
            ValidarNIF Text1(Index).Text
        
        Case 20 'BANCO PROPIO
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "banpropi", "nombanpr")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Banco Propio: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "�Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmBan = New frmManBanco
                        frmBan.DatosADevolverBusqueda = "0|1|"
                        frmBan.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmBan.Show vbModal
                        Set frmBan = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, adodc1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
        
        Case 25 'FORMA DE PAGO
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "forpago", "nomforpa")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Forma de Pago: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "�Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmFPa = New frmManFpago
                        frmFPa.DatosADevolverBusqueda = "0|1|"
                        frmFPa.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmFPa.Show vbModal
                        Set frmFPa = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, adodc1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 21, 22 'ENTIDAD Y SUCURSAL BANCARIA
            PonerFormatoEntero Text1(Index)
            
        Case 19 'cuenta contable
            If Text1(Index).Text = "" Then Exit Sub
            If Modo = 3 Then
                Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo, "") 'text1(0).Text)
            Else
                Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo, Text1(0).Text)
            End If
            
        Case 26 ' porcentaje de retencion
            If Modo = 1 Then Exit Sub
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 4
            
        Case 27 ' codigo de iban
            Text1(Index).Text = UCase(Text1(Index).Text)
            
    End Select
    
    '[Monica]: calculo del iban si no lo ponen
    If Index = 21 Or Index = 22 Or Index = 23 Or Index = 24 Then
        Dim cta As String
        Dim CC As String
        If Text1(21).Text <> "" And Text1(22).Text <> "" And Text1(23).Text <> "" And Text1(24).Text <> "" Then
            
            cta = Format(Text1(21).Text, "0000") & Format(Text1(22).Text, "0000") & Format(Text1(23).Text, "00") & Format(Text1(24).Text, "0000000000")
            If Len(cta) = 20 Then
    '        Text1(42).Text = Calculo_CC_IBAN(cta, Text1(42).Text)
    
                If Text1(27).Text = "" Then
                    'NO ha puesto IBAN
                    If DevuelveIBAN2("ES", cta, cta) Then Text1(27).Text = "ES" & cta
                Else
                    CC = CStr(Mid(Text1(27).Text, 1, 2))
                    If DevuelveIBAN2(CStr(CC), cta, cta) Then
                        If Mid(Text1(27).Text, 3) <> cta Then
                            
                            MsgBox "Codigo IBAN distinto del calculado [" & CC & cta & "]", vbExclamation
                        End If
                    End If
                End If
            End If
        End If
    End If
            
            
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 5
                BotonBuscar
        Case 6
                BotonVerTodos
        Case 1
                BotonAnyadir
        Case 2
                mnModificar_Click
        Case 3
                BotonEliminar
        Case 8 'Imprimir
                mnImprimir_Click
    End Select
End Sub

Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Me.adodc1.Recordset.EOF Then Exit Sub
    DesplazamientoData adodc1, Index, True
    PonerCampos
End Sub

Private Sub PonerCampos()

    If adodc1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Me.adodc1
    
    If vParamAplic.NumeroConta <> 0 Then
        Text2(19).Text = PonerNombreCuenta(Text1(19), Modo)
    End If

    Text2(25).Text = PonerNombreDeCod(Text1(25), "forpago", "nomforpa")
    Text2(20).Text = PonerNombreDeCod(Text1(20), "banpropi", "nombanpr")

    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = PonerContRegistros(Me.adodc1)
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim cta As String
Dim cadMen As String

    b = CompForm(Me)
    If Not b Then Exit Function
    
    If (Modo = 3) Then 'Estem insertant
         If ExisteCP(Text1(0)) Then b = False
    End If
    
    If b And (Modo = 3 Or Modo = 4) Then
        
        
        '[Monica]22/08/2013: a�adida la comprobacion de que la cuenta contable sea correcta
        If Text1(21).Text = "" Or Text1(22).Text = "" Or Text1(23).Text = "" Or Text1(24).Text = "" Then
            '[Monica]20/11/2013: a�adido el codigo de iban
            Text1(27).Text = ""
            Text1(21).Text = ""
            Text1(22).Text = ""
            Text1(23).Text = ""
            Text1(24).Text = ""
        Else
            cta = Format(Text1(21).Text, "0000") & Format(Text1(22).Text, "0000") & Format(Text1(23).Text, "00") & Format(Text1(24).Text, "0000000000")
            If Val(ComprobarCero(cta)) = 0 Then
                cadMen = "La agencia no tiene asignada cuenta bancaria."
                MsgBox cadMen, vbExclamation
            End If
            If Not Comprueba_CC(cta) Then
                cadMen = "La cuenta bancaria de la agencia no es correcta. � Desea continuar ?."
                If MsgBox(cadMen, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    b = True
                Else
                    PonerFoco Text1(21)
                    b = False
                End If
            Else
'                '[Monica]20/11/2013: a�adimos el tema de la comprobacion del IBAN
'                If Not Comprueba_CC_IBAN(cta, Text1(42).Text) Then
'                    cadMen = "La cuenta IBAN del cliente no es correcta. � Desea continuar ?."
'                    If MsgBox(cadMen, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
'                        b = True
'                    Else
'                        PonerFoco Text1(42)
'                        b = False
'                    End If
'                End If

'       sustituido por lo de David
                BuscaChekc = ""
                If Me.Text1(27).Text <> "" Then BuscaChekc = Mid(Text1(27).Text, 1, 2)
                    
                If DevuelveIBAN2(BuscaChekc, cta, cta) Then
                    If Me.Text1(27).Text = "" Then
                        If MsgBox("Poner IBAN ?", vbQuestion + vbYesNo) = vbYes Then Me.Text1(27).Text = BuscaChekc & cta
                    Else
                        If Mid(Text1(27).Text, 3) <> cta Then
                            cta = "Calculado : " & BuscaChekc & cta
                            cta = "Introducido: " & Me.Text1(27).Text & vbCrLf & cta & vbCrLf
                            cta = "Error en codigo IBAN" & vbCrLf & cta & "Continuar?"
                            If MsgBox(cta, vbQuestion + vbYesNo) = vbNo Then
                                PonerFoco Text1(27)
                                b = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    
    DatosOk = b
End Function

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me

End Sub

Private Function SepuedeBorrar() As Boolean
    SepuedeBorrar = True
End Function

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub HacerBusqueda()

    CadB = ObtenerBusqueda(Me, , False)
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NomTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        ' ******** Si la clau primaria no es Text1(0), canviar-ho ***********
        PonerFoco Text1(1)
        ' *******************************************************************
    End If
End Sub

Private Sub LimpiarCampos()
Dim i As Integer

    On Error Resume Next

    limpiar Me
    Combo1(0).ListIndex = -1
    Combo1(1).ListIndex = -1
    
    ' ****************************************************
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    Cad = "(codtrans=" & Text1(0).Text & ")"
    If SituarData(Me.adodc1, Cad, Indicador) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
End Sub

Private Sub printNou()
    
    With frmImprimir2
        .cadTabla2 = "agencias"
        .Informe2 = "rManAgencias.rpt"
        If CadB <> "" Then
            .cadRegSelec = SQL2SF(CadB)
        Else
            .cadRegSelec = ""
        End If
        .cadRegActua = POS2SF(adodc1, Me)
        .cadTodosReg = ""
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomempre & "'|" '& "pOrden={agencias.codtrans}|"
        .NumeroParametros2 = 1
        .MostrarTree2 = False
        .InfConta2 = False
        .ConSubInforme2 = False
        .SubInformeConta = ""
        .Show vbModal
    End With
End Sub


Private Sub CargaCombo()
Dim Cad As String
Dim i As Byte
Dim Rs As ADODB.Recordset

    On Error GoTo ErrCarga
    
    Combo1(0).Clear
    
    Combo1(0).AddItem "Transportista"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Comisionista"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    
    
    
    'Tipo de transporte
    Combo1(1).Clear
    
    Cad = "SELECT * FROM fortranspor ORDER BY codfortrans"
    Set Rs = New ADODB.Recordset
'    Rs.Open Cad, conn, OpenForwardOnly, adLockPessimistic, adCmdText
    Rs.Open Cad, conn, adOpenDynamic, adLockReadOnly, adCmdText
    
    While Not Rs.EOF
        Combo1(1).AddItem Rs!Nombre
        Combo1(1).ItemData(Combo1(1).NewIndex) = Rs!codfortrans
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
        
    
ErrCarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar datos combo.", Err.Description
End Sub



'[Monica]26/03/2015: no se modificaban los datos de la cuenta de proveedor


Private Sub ModificarDatosCuentaContable()
Dim Sql As String
Dim Cad As String
Dim CtaBancoPropio As String

    On Error GoTo eModificarDatosCuentaContable

    '[Monica]06/02/2019: si no me han puesto cuenta contable no actualizo nada
    If Text1(19).Text = "" Then Exit Sub

    CtaBancoPropio = DevuelveValor("select codmacta from banpropi where codbanpr = " & DBSet(Text1(20).Text, "N"))


    If Text1(1).Text <> NombreAnt Or Text1(21).Text <> BancoAnt Or Text1(22).Text <> SucurAnt Or Text1(23).Text <> DigitoAnt Or Text1(24).Text <> CuentaAnt Or _
       DirecAnt <> Text1(3).Text Or cPostalAnt <> Text1(4).Text Or PoblaAnt <> Text1(5).Text Or ProviAnt <> Text1(6).Text Or NifAnt <> Text1(2).Text Or _
       EMaiAnt <> Text1(15).Text Or WebAnt <> Text1(10).Text Or _
       forpaant <> Text1(25).Text Or _
       IbanAnt <> Text1(27).Text Or _
       CtaBancoPropio <> CtaBancoAnt Then
        
        Cad = "Se han producido cambios en datos de la Agencia de Transporte. " '& vbCrLf
        
        Cad = Cad & vbCrLf & vbCrLf & "� Desea actualizarlos en la Contabilidad ?" & vbCrLf & vbCrLf
        
        If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                        
            Sql = "update cuentas set nommacta = " & DBSet(Trim(Text1(1).Text), "T")
            
            Sql = Sql & ", dirdatos = " & DBSet(Trim(Text1(3).Text), "T")
            Sql = Sql & ", codposta = " & DBSet(Trim(Text1(4).Text), "T")
            Sql = Sql & ", despobla = " & DBSet(Trim(Text1(5).Text), "T")
            Sql = Sql & ", desprovi = " & DBSet(Trim(Text1(6).Text), "T")
            Sql = Sql & ", maidatos = " & DBSet(Trim(Text1(15).Text), "T")
            Sql = Sql & ", webdatos = " & DBSet(Trim(Text1(10).Text), "T")
            Sql = Sql & ", nifdatos = " & DBSet(Trim(Text1(2).Text), "T")
            '[Monica]26/03/2015: antes no grababamos la forma de pago de la cuenta
            Sql = Sql & ", forpa = " & DBSet(Trim(Text1(25).Text), "N", "S")
            
            If vParamAplic.ContabilidadNueva Then
                Dim vIban As String
                
                vIban = MiFormat(Text1(27).Text, "") & MiFormat(Text1(21).Text, "0000") & MiFormat(Text1(22).Text, "0000") & MiFormat(Text1(23).Text, "00") & MiFormat(Text1(24).Text, "0000000000")
            
                Sql = Sql & ", iban = " & DBSet(vIban, "T")
                Sql = Sql & ", codpais = 'ES' "
            Else
                Sql = Sql & ", entidad = " & DBSet(Trim(Text1(21).Text), "T", "S")
                Sql = Sql & ", oficina = " & DBSet(Trim(Text1(22).Text), "T", "S")
                Sql = Sql & ", cc = " & DBSet(Trim(Text1(23).Text), "T", "S")
                Sql = Sql & ", cuentaba = " & DBSet(Trim(Text1(24).Text), "T", "S")
                
                '[Monica]22/11/2013: tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    Sql = Sql & ", iban = " & DBSet(Trim(Text1(27).Text), "T", "S")
                End If
            End If
            
            '[Monica]27/10/2016: si han cambiado la cta de pago hay que cambiarla
            Sql = Sql & ", ctabanco = " & DBSet(CtaBancoPropio, "T")
            
            Sql = Sql & " where codmacta = " & DBSet(Trim(Text1(19).Text), "T")
                        
            ConnConta.Execute Sql
                        
'            MsgBox "Datos de Cuenta modificados correctamente.", vbExclamation
                        
        End If
    End If
    
    
    '[Monica]30/08/2013: modificamos los datos de tesoreria sobre los cobros y pagos pendientes
    If Text1(21).Text <> BancoAnt Or Text1(22).Text <> SucurAnt Or Text1(23).Text <> DigitoAnt Or Text1(24).Text <> CuentaAnt _
        Or Text1(27).Text <> IbanAnt Or Text1(25).Text <> forpaant Then
        Cad = "Se han producido cambios en la Cta.Bancaria la agencia de transporte."
        Cad = Cad & vbCrLf & vbCrLf & "� Desea actualizar los Cobros y Pagos pendientes en Tesoreria ?" & vbCrLf & vbCrLf
        
        If HayCobrosPagosPendientes(Text1(19).Text) Then
            If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                If ActualizarCobrosPagosPdtes(Text1(19), Text1(21).Text, Text1(22).Text, Text1(23).Text, Text1(24).Text, Text1(27).Text, Text1(25).Text) Then
'                    MsgBox "Datos en Tesoreria modificados correctamente.", vbExclamation
                End If
            End If
        End If
    End If
    
    Exit Sub
    
eModificarDatosCuentaContable:
    MuestraError Err.Number, "Modificar Datos Cuenta Contable", Err.Description
End Sub


Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub
