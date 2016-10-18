VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMensajes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mensajes"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14160
   Icon            =   "frmMensajes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   14160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameAnecoop 
      Height          =   5790
      Left            =   0
      TabIndex        =   87
      Top             =   60
      Width           =   7050
      Begin VB.CommandButton CmdAcepFrasAnecoop 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   89
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton CmdCancelFrasAnecoop 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   5520
         TabIndex        =   88
         Top             =   5160
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView11 
         Height          =   4155
         Left            =   225
         TabIndex        =   90
         Top             =   810
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   7329
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Factura"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Base Imp"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "%Iva"
            Object.Width           =   3706
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "IVA"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Total"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label10 
         Caption         =   "Facturas Anecoop a integrar "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   270
         TabIndex        =   91
         Top             =   270
         Width           =   5145
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   13
         Left            =   600
         Picture         =   "frmMensajes.frx":000C
         Top             =   5160
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   12
         Left            =   240
         Picture         =   "frmMensajes.frx":0156
         Top             =   5160
         Width           =   240
      End
   End
   Begin VB.Frame FrameCobrosPtes 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   8655
      Begin VB.CommandButton cmdCancelarCobros 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   7080
         TabIndex        =   25
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox txtParam 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   23
         Text            =   "frmMensajes.frx":02A0
         Top             =   240
         Width           =   6615
      End
      Begin VB.CommandButton cmdAceptarCobros 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5400
         TabIndex        =   1
         Top             =   4440
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3135
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   5530
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "¿Desea continuar?"
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   26
         Top             =   4440
         Width           =   7215
      End
      Begin VB.Label Label1 
         Caption         =   "Departamento:"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   7215
      End
   End
   Begin VB.Frame frameClaveAcceso 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1365
      Left            =   0
      TabIndex        =   79
      Top             =   0
      Width           =   3645
      Begin VB.TextBox Text7 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1290
         PasswordChar    =   "*"
         TabIndex        =   80
         Top             =   570
         Width           =   1665
      End
      Begin VB.Label Label16 
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   270
         TabIndex        =   81
         Top             =   600
         Width           =   945
      End
   End
   Begin VB.Frame FramePedidosSinAlbaran 
      Height          =   4620
      Left            =   0
      TabIndex        =   56
      Top             =   0
      Width           =   8655
      Begin VB.CommandButton CmdPedSinAlb 
         Caption         =   "&Regresar"
         Height          =   375
         Left            =   7290
         TabIndex        =   57
         Top             =   4005
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView6 
         Height          =   3555
         Left            =   150
         TabIndex        =   58
         Top             =   330
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   6271
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Pedidos sin Albarán Asignado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   59
         Top             =   90
         Visible         =   0   'False
         Width           =   7215
      End
   End
   Begin VB.Frame FramePaletsAsociados 
      Height          =   4620
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Width           =   8655
      Begin VB.CommandButton CmdAceptarPal 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5400
         TabIndex        =   52
         Top             =   4005
         Width           =   975
      End
      Begin VB.CommandButton CmdCanPal 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   7080
         TabIndex        =   51
         Top             =   4005
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView5 
         Height          =   3135
         Left            =   120
         TabIndex        =   53
         Top             =   720
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   5530
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "¿ Desea Continuar ?"
         Height          =   285
         Index           =   2
         Left            =   360
         TabIndex        =   55
         Top             =   4050
         Width           =   2715
      End
      Begin VB.Label Label1 
         Caption         =   "Palets Asociados al Pedido:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   54
         Top             =   240
         Width           =   7215
      End
   End
   Begin VB.Frame FrameCorreccionPrecios 
      Height          =   6375
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   12975
      Begin VB.ComboBox cmbActualizarTar 
         Height          =   315
         ItemData        =   "frmMensajes.frx":02A6
         Left            =   7800
         List            =   "frmMensajes.frx":02B3
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   5960
         Width           =   2175
      End
      Begin VB.CommandButton cmdCorrecotrPrecios 
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   11760
         TabIndex        =   38
         Top             =   5880
         Width           =   975
      End
      Begin VB.CommandButton cmdCorrecotrPrecios 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   10560
         TabIndex        =   37
         Top             =   5880
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   5175
         Left            =   120
         TabIndex        =   36
         Top             =   600
         Width           =   12660
         _ExtentX        =   22331
         _ExtentY        =   9128
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   3246
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Denominación"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "U.P.Compra"
            Object.Width           =   2364
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "% M"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "PVP"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "%T"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "P.Tarifa"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "PVP Correcto"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Tarifa correc."
            Object.Width           =   2011
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Actualizar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   6480
         TabIndex        =   42
         Top             =   6000
         Width           =   1215
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   11760
         Picture         =   "frmMensajes.frx":02E9
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   12360
         Picture         =   "frmMensajes.frx":0433
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lblIndicadorCorregir 
         Caption         =   "Label3"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   5880
         Width           =   5055
      End
      Begin VB.Label Label2 
         Caption         =   "Corrección de errores y actualización de tarifas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   8175
      End
   End
   Begin VB.Frame FrameErrores 
      Height          =   5535
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   8415
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   375
         Left            =   7080
         TabIndex        =   29
         Top             =   4920
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   4335
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   28
         Text            =   "frmMensajes.frx":057D
         Top             =   360
         Width           =   7695
      End
   End
   Begin VB.Frame FrameAcercaDe 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6375
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C/ Franco Tormo, 3 Bajo Izda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   3000
         TabIndex        =   10
         Top             =   2640
         Width           =   2535
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "46007 - VALENCIA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   3000
         TabIndex        =   9
         Top             =   2925
         Width           =   2535
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tfno: 96 358 05 47"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   3000
         TabIndex        =   8
         Top             =   3195
         Width           =   2535
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax: 96 378 82 01"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   3000
         TabIndex        =   7
         Top             =   3480
         Width           =   2535
      End
      Begin VB.Image Image2 
         Height          =   540
         Left            =   720
         Top             =   2640
         Width           =   2160
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -120
         TabIndex        =   6
         Top             =   1260
         Width           =   4155
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   81.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1725
         Left            =   4260
         TabIndex        =   5
         Top             =   0
         Width           =   1350
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "ARIGES"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   915
         Left            =   1080
         TabIndex        =   4
         Top             =   300
         Width           =   3495
      End
   End
   Begin VB.Frame FrameNSeries 
      Height          =   5000
      Left            =   480
      TabIndex        =   11
      Top             =   120
      Width           =   6975
      Begin VB.CommandButton cmdSelTodos 
         Caption         =   "&Todos"
         Height          =   315
         Left            =   720
         TabIndex        =   22
         Top             =   4320
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.CommandButton cmdDeselTodos 
         Caption         =   "&Ninguno"
         Height          =   315
         Left            =   1680
         TabIndex        =   21
         Top             =   4320
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5040
         TabIndex        =   20
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarNSeries 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3960
         TabIndex        =   12
         Top             =   4320
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3015
         Left            =   720
         TabIndex        =   13
         Top             =   840
         Width           =   4020
         _ExtentX        =   7091
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "Empresas en el sistema"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   1
         Left            =   720
         TabIndex        =   30
         Top             =   240
         Visible         =   0   'False
         Width           =   5295
      End
   End
   Begin VB.Frame FrameComponentes 
      Height          =   3975
      Left            =   0
      TabIndex        =   14
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton cmdAceptarComp 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3240
         TabIndex        =   19
         Top             =   3240
         Width           =   975
      End
      Begin VB.Frame FrameComponentes2 
         Caption         =   "Mostrar Equipos del :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2175
         Left            =   960
         TabIndex        =   15
         Top             =   720
         Width           =   3255
         Begin VB.OptionButton OptCompXClien 
            Caption         =   "Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   18
            Top             =   1440
            Width           =   2535
         End
         Begin VB.OptionButton OptCompXDpto 
            Caption         =   "Departamento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   17
            Top             =   960
            Width           =   2535
         End
         Begin VB.OptionButton OptCompXMant 
            Caption         =   "Mantenimiento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   16
            Top             =   480
            Width           =   2055
         End
      End
   End
   Begin VB.Frame FrameTraspasoMante 
      Height          =   3135
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   5415
      Begin VB.TextBox txtMante 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   48
         Top             =   1200
         Width           =   855
      End
      Begin VB.CheckBox chkMante 
         Caption         =   "Copiar importes en siguiente"
         Height          =   255
         Left            =   360
         TabIndex        =   46
         Top             =   1800
         Value           =   1  'Checked
         Width           =   4695
      End
      Begin VB.CommandButton cmdMante 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   45
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton cmdMante 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   2520
         TabIndex        =   44
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Año a traspasar"
         Height          =   195
         Left            =   600
         TabIndex        =   49
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Pasar importes mantenimiento a historico."
         Height          =   735
         Left            =   240
         TabIndex        =   47
         Top             =   360
         Width           =   4695
      End
   End
   Begin VB.Frame FrameProductos 
      Height          =   5790
      Left            =   30
      TabIndex        =   82
      Top             =   30
      Width           =   7050
      Begin VB.CommandButton CmdCanProductos 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   5520
         TabIndex        =   84
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton CmdAcepProductos 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   83
         Top             =   5160
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView10 
         Height          =   4155
         Left            =   225
         TabIndex        =   85
         Top             =   810
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   7329
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Variedad"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Clase "
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descripcion"
            Object.Width           =   3706
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   11
         Left            =   240
         Picture         =   "frmMensajes.frx":0583
         Top             =   5160
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   10
         Left            =   600
         Picture         =   "frmMensajes.frx":06CD
         Top             =   5160
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Productos del grupo a modificar "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   270
         TabIndex        =   86
         Top             =   270
         Width           =   5145
      End
   End
   Begin VB.Frame FrameVariedades 
      Height          =   5790
      Left            =   -45
      TabIndex        =   60
      Top             =   0
      Width           =   7050
      Begin VB.CommandButton cmdAcepVariedades 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   63
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton cmdCanVariedades 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   5520
         TabIndex        =   62
         Top             =   5160
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView7 
         Height          =   4155
         Left            =   225
         TabIndex        =   61
         Top             =   810
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   7329
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Variedad"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Clase "
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descripcion"
            Object.Width           =   3706
         EndProperty
      End
      Begin VB.Label Label5 
         Caption         =   "Variedades"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   270
         TabIndex        =   64
         Top             =   270
         Width           =   5145
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   5
         Left            =   600
         Picture         =   "frmMensajes.frx":0817
         Top             =   5160
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   4
         Left            =   240
         Picture         =   "frmMensajes.frx":0961
         Top             =   5160
         Width           =   240
      End
   End
   Begin VB.Frame FrameFacturasACuenta 
      Height          =   5790
      Left            =   0
      TabIndex        =   65
      Top             =   0
      Width           =   7050
      Begin VB.CommandButton CmdCancelFactACta 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   5550
         TabIndex        =   67
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton CmdAceptarFactACta 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   66
         Top             =   5160
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView8 
         Height          =   4155
         Left            =   225
         TabIndex        =   68
         Top             =   810
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   7329
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Variedad"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Clase "
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descripcion"
            Object.Width           =   3706
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   7
         Left            =   240
         Picture         =   "frmMensajes.frx":0AAB
         Top             =   5160
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   6
         Left            =   600
         Picture         =   "frmMensajes.frx":0BF5
         Top             =   5160
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Facturas a Cuenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   270
         TabIndex        =   69
         Top             =   270
         Width           =   5145
      End
   End
   Begin VB.Frame FramePMP 
      Height          =   6615
      Left            =   0
      TabIndex        =   75
      Top             =   0
      Width           =   10215
      Begin VB.CommandButton cmdCancelPMP 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   8880
         TabIndex        =   77
         Top             =   6000
         Width           =   1095
      End
      Begin VB.CommandButton cmdActualizaPMP 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   7560
         TabIndex        =   76
         Top             =   6000
         Width           =   1095
      End
      Begin MSComctlLib.ListView lw 
         Height          =   5475
         Index           =   0
         Left            =   210
         TabIndex        =   78
         Top             =   300
         Width           =   9750
         _ExtentX        =   17198
         _ExtentY        =   9657
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   3246
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Denominación"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "U.P.Compra"
            Object.Width           =   2364
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "% M"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "PVP"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "%T"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "P.Tarifa"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "PVP Correcto"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Tarifa correc."
            Object.Width           =   2011
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   9
         Left            =   240
         Picture         =   "frmMensajes.frx":0D3F
         Top             =   6000
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   8
         Left            =   600
         Picture         =   "frmMensajes.frx":0E89
         Top             =   6000
         Width           =   240
      End
   End
   Begin VB.Frame FrameLineasCalibre 
      Height          =   5790
      Left            =   0
      TabIndex        =   70
      Top             =   0
      Width           =   7050
      Begin VB.CommandButton CmdAcepCalib 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   72
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   5550
         TabIndex        =   71
         Top             =   5160
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView9 
         Height          =   4155
         Left            =   225
         TabIndex        =   73
         Top             =   810
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   7329
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Variedad"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Clase "
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descripcion"
            Object.Width           =   3706
         EndProperty
      End
      Begin VB.Label Label8 
         Caption         =   "Líneas de Calibre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   270
         TabIndex        =   74
         Top             =   270
         Width           =   5145
      End
   End
   Begin VB.Frame FrameEtiqEstant 
      Height          =   7455
      Left            =   0
      TabIndex        =   31
      Top             =   -120
      Width           =   8535
      Begin VB.CommandButton cmdEtiqEstan 
         Caption         =   "Imprimir"
         Height          =   375
         Index           =   1
         Left            =   5520
         TabIndex        =   34
         Top             =   6960
         Width           =   1215
      End
      Begin VB.CommandButton cmdEtiqEstan 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   0
         Left            =   6960
         TabIndex        =   33
         Top             =   6960
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   6495
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   11456
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Descripcion"
            Object.Width           =   6703
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Precio"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Familia"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   600
         Picture         =   "frmMensajes.frx":0FD3
         Top             =   6960
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "frmMensajes.frx":111D
         Top             =   6960
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmMensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'====================== VBLES PUBLICAS ================================

Public Event DatoSeleccionado(CadenaSeleccion As String)

Public OpcionMensaje As Byte
'======================================
'==== FACTURACION =====================
' 1 .- Mensaje de Cobros Pendientes
' 2 .- Mensaje de No hay suficiente Stock para pasar de Pedido a Albaran
' 3 .- Mensaje Acerca de...
' 4 .- Listado de los Nº de Serie de un Articulo
' 5 .- Seleccionar tipo de Componente a Mostrar en Mant. de Nº de Series
' 6 .- Mostrar Prefacturacion de Albaranes
' 7 .- Mostrar Prefacturacion Mantenimientos
' 8 .- Mostrar lista clientes para seleccionar los que queremos imprimir (Etiquetas)
' 9 .- Mostrar lista Proveedores para seleccionar los que queremos imprimir (Etiquetas)
'10 .- Mostrar lista de Errores de las facturas NO contabilizadas
'11 .- Mostrar lista lineas de factura a Rectificar para seleccionar las q queremos traer al Albaran de FAct. Rectificativa
'12 .- Mostrar Albaranes del Rango que no se van a Facturar. (Facturar Albaranes Venta)

'13 .- Mostrar Errores
'14 .- Mostrar Empresas existentes en el sistema



'15 .- Mostrar lista de articulos para imprimir etiquetas estanteria
'16 .- Lista de articulos para corregir importes
'17 .- Etiquetas clientes. LO MISMO QUE EL 8 pero hecho por david
'18 .- Mantenimientos. paso ejercicio siguiente a actual
'19 .- Lista de palets de venta asociados al pedido del que se va a generar el albaran
'20 .- Lista de pedidos sin numero de albaran asignado.
'21 .


'22 .- Facturas a cuenta que se han hecho al cliente
'23 .- Lineas de calibre de un albaran

'24 .- Cambio del precio medio ponderado
'25 .- Cambio del ultimo precio

'27 .- Mostrar clave de acceso

'28 .- Mostramos los productos que hay del mismo grupo

'29 .- Facturas de anecoop a generar

Public cadWHERE As String 'Cadena para pasarle la WHERE de la SELECT de los cobros pendientes o de Pedido(para comp. stock)
                          'o CodArtic para seleccionar los Nº Series
                          'para cargar el ListView
                          
Public cadWHERE2 As String

Public vCampos As String 'Articulo y cantidad Empipados para Nº de Series
                         'Tambien para pasar el nombre de la tabla de lineas (sliped, slirep,...)
                         'Dependiendo desde donde llamemos, de Pedidos o Reparaciones


'====================== VBLES LOCALES ================================

Dim PulsadoSalir As Boolean 'Solo salir con el boton de Salir no con aspa del form
Dim PrimeraVez As Boolean

'Para los Nº de Serie
Dim TotalArray As Integer
Dim codArtic() As String
Dim Cantidad() As Integer

Dim IT As ListItem
Dim vCadena As String
Dim Sql As String

Private Sub CmdAcepCalib_Click()
Dim CADENA As String
    'Cargo las facturas a cuenta que hay que descontar
    CADENA = ""
    
    CADENA = CADENA & ListView9.SelectedItem.Text & "|" & ListView9.SelectedItem.SubItems(1) & "|" & ListView9.SelectedItem.SubItems(2) & "|" & ListView9.SelectedItem.SubItems(3) & "|" & ListView9.SelectedItem.SubItems(4) & "|" & ListView9.SelectedItem.SubItems(5) & "|" & ListView9.SelectedItem.SubItems(6) & "|"
    
    RaiseEvent DatoSeleccionado(CADENA)
    
    Unload Me
End Sub

Private Sub CmdAcepFrasAnecoop_Click()
Dim CADENA As String
    'Cargo las facturas marcadas
    CADENA = ""
    
    For NumRegElim = 1 To ListView11.ListItems.Count
        If ListView11.ListItems(NumRegElim).Checked And ListView11.ListItems(NumRegElim).Tag = 1 Then
            CADENA = CADENA & ListView11.ListItems(NumRegElim).Text & ","
        End If
    Next NumRegElim
    
    ' quitamos la ultima coma
    If CADENA <> "" Then
        CADENA = Mid(CADENA, 1, Len(CADENA) - 1)
    End If
    
    RaiseEvent DatoSeleccionado(CADENA)
    Unload Me

End Sub

Private Sub CmdAcepProductos_Click()
Dim CADENA As String
    'Cargo las variedades marcadas
    CADENA = ""
    
    For NumRegElim = 1 To ListView10.ListItems.Count
        If ListView10.ListItems(NumRegElim).Checked Then
            CADENA = CADENA & ListView10.ListItems(NumRegElim).Text & ","
        Else
            SeleccionadosTodos = False
        End If
    Next NumRegElim
    ' quitamos la ultima coma
    If CADENA <> "" Then
        CADENA = Mid(CADENA, 1, Len(CADENA) - 1)
    End If
    
    RaiseEvent DatoSeleccionado(CADENA)
    Unload Me

End Sub

Private Sub CmdAceptarCobros_Click()
    If OpcionMensaje = 12 Then vCampos = "1"
    Unload Me
End Sub


Private Sub CmdAceptarFactACta_Click()
Dim CADENA As String
    'Cargo las facturas a cuenta que hay que descontar
    CADENA = ""
    For NumRegElim = 1 To ListView8.ListItems.Count
        If ListView8.ListItems(NumRegElim).Checked Then
            CADENA = CADENA & "('" & ListView8.ListItems(NumRegElim).Text & "'," & ListView8.ListItems(NumRegElim).SubItems(1) & "," & DBSet(ListView8.ListItems(NumRegElim).SubItems(2), "F") & "),"
        End If
    Next NumRegElim
    ' quitamos la ultima coma
    If CADENA <> "" Then
        CADENA = Mid(CADENA, 1, Len(CADENA) - 1)
    End If
    
    RaiseEvent DatoSeleccionado(CADENA)
    Unload Me
End Sub

'Private Sub cmdAceptarComp_Click()
''Boton Aceptar de Componentes del Mant. de Nº de Series en Reparaciones
'Dim h As Integer, w As Integer
'
'    ponerFrameComponentesVisible False, h, w
'    PonerFrameCobrosPtesVisible True, h, w
'    Me.Height = h + 350
'    Me.Width = w + 70
'
'    If Me.OptCompXMant.Value Then
'        'Mostrar Resumen de los Nº de Serie del Mantenimiento
'        Me.Caption = "Equipos del Mantenimiento"
'        CargarListaComponentes (1)
'    ElseIf Me.OptCompXDpto.Value Then
'        'Mostrar Resumen de los Nº de Serie del Departamento
'        Me.Caption = "Equipos del Departamento"
'        CargarListaComponentes (2)
'    ElseIf Me.OptCompXClien.Value Then
'        'Mostrar Resumen de los Nº de Serie del Cliente
'        Me.Caption = "Equipos del Cliente"
'        CargarListaComponentes (3)
'    End If
'    PonerFocoBtn Me.cmdAceptarCobros
'End Sub


Private Sub cmdAceptarPal_Click()
    RaiseEvent DatoSeleccionado("1")
    Unload Me
End Sub

Private Sub cmdAceptarNSeries_Click()
Dim I As Byte, j As Byte
Dim Seleccionados As Integer
Dim cad As String, Sql As String
Dim articulo As String
Dim Rs As ADODB.Recordset
Dim C1 As String * 10, c2 As String * 10, c3 As String * 10


    If OpcionMensaje = 4 Then
        'Comprobar que se han seleccionado el nº correcto de  Nº de Serie para cada Articulo
        Seleccionados = 0
        articulo = ""
      
        'Si se ha seleccionado la cantidad correcta de Nº de series, empiparlos y
        'devolverlos al form de Albaranes(facturacion)
        cad = ""
        For j = 0 To TotalArray
            articulo = codArtic(j)
            cad = cad & articulo & "|"
            For I = 1 To ListView2.ListItems.Count
                If ListView2.ListItems(I).Checked Then
                    If articulo = ListView2.ListItems(I).ListSubItems(1).Text Then
                        If Seleccionados < Abs(Cantidad(j)) Then
                            Seleccionados = Seleccionados + 1
                            cad = cad & ListView2.ListItems(I).Text & "|"
                        End If
                   'cad = cad & Data1.Recordset.Fields(1) & "|"
                    End If
                End If
            Next I
            If Seleccionados < Abs(Cantidad(j)) Then
                'Comprobar que si tiene Nºs de serie de ese articulos cargados seleccione los
                'que corresponden
                Sql = "SELECT count(sserie.numserie)"
                Sql = Sql & " FROM sserie " 'INNER JOIN sartic ON sserie.codartic=sartic.codartic "
                Sql = Sql & " WHERE sserie.codartic=" & DBSet(articulo, "T")
                Sql = Sql & " AND (isnull(sserie.numfactu) or sserie.numfactu='') and (isnull(sserie.numalbar) or sserie.numalbar='') "
                Sql = Sql & " ORDER BY sserie.codartic, numserie "
                Set Rs = New ADODB.Recordset
                Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                If Rs.Fields(0).Value >= Abs(Cantidad(j)) - Seleccionados Then
                    MsgBox "Debe seleccionar " & Cantidad(j) & " Nº Series para el articulo " & codArtic(j), vbExclamation
                    Exit Sub
                Else
                    'No hay Nº Serie y Pedirlos
                End If
                Rs.Close
                Set Rs = Nothing
            
            End If
            cad = cad & "·"
            Seleccionados = 0
        Next j
      
    ElseIf OpcionMensaje = 8 Or OpcionMensaje = 9 Or OpcionMensaje = 17 Then
        'concatenar todos los clientes seleccionados para imprimir etiquetas
        If OpcionMensaje = 17 Then
            
            '----------------------------------------------------------------
            cad = "insert into tmpnlotes (codusu,numalbar,fechaalb,numlinea,codprove) values ("
            cad = cad & vUsu.codigo & ",1,'2005-04-12',1,"
            
            
            For I = 1 To ListView2.ListItems.Count
                If ListView2.ListItems(I).Checked Then
                    conn.Execute cad & (ListView2.ListItems(I).Text) & ")"
                    NumRegElim = NumRegElim + 1
                End If
            Next I
            
            
            '----------------------------------------------------------------
            
        Else
            cad = ""
            For I = 1 To ListView2.ListItems.Count
                If ListView2.ListItems(I).Checked Then
                    cad = cad & Val(ListView2.ListItems(I).Text) & ","
                     'cad = cad & Data1.Recordset.Fields(1) & "|"
                End If
            Next I
            If cad <> "" Then cad = Mid(cad, 1, Len(cad) - 1)
        End If
    ElseIf OpcionMensaje = 11 Then
    'Lineas Factura a rectificar
        'cad = "(" & cadWHERE & ")"
        cad = ""
        C1 = ""
        c2 = ""
        c3 = ""
        Sql = ""
        For I = 1 To ListView2.ListItems.Count
            If ListView2.ListItems(I).Checked Then
                If Sql = "" Then
                    C1 = DBSet(ListView2.ListItems(I), "T", "N")
                    c2 = ListView2.ListItems(I).ListSubItems(1)
'                    c3 = ListView2.ListItems(i).ListSubItems(2)
                    cad = "(codtipoa=" & Trim(C1) & " and numalbar=" & Val(c2) & " and numlinea IN (" & ListView2.ListItems(I).ListSubItems(2)

                Else
                    If Trim(DBSet(ListView2.ListItems(I), "T", "N")) = Trim(C1) And Trim(ListView2.ListItems(I).ListSubItems(1)) = Trim(c2) Then
                    'es el mismo albaran y concatenamos lineas
                        cad = "," & ListView2.ListItems(I).ListSubItems(2)

                    Else
                        If cad <> "" Then Sql = Sql & ")) "
                        C1 = DBSet(ListView2.ListItems(I), "T", "N")
                        c2 = ListView2.ListItems(I).ListSubItems(1)
'                    c3 = ListView2.ListItems(i).ListSubItems(2)
                        cad = " or (codtipoa=" & Trim(C1) & " and numalbar=" & Val(c2) & " and numlinea IN (" & ListView2.ListItems(I).ListSubItems(2)
                        
'                       cad=cad &
                    End If
                End If
                Sql = Sql & cad
'                If cad <> "" Then cad = cad & " OR "
'                cad = cad & "(codtipoa=" & DBSet(ListView2.ListItems(i), "T", "N") & " and numalbar=" & Val(ListView2.ListItems(i).ListSubItems(1)) & " and numlinea=" & ListView2.ListItems(i).ListSubItems(2) & ")"
            Else
'                cad = ""
            End If
        Next I
        If cad <> "" Then
            Sql = Sql & "))"
            cad = "(" & cadWHERE & ") AND (" & Sql & ")"
        End If
'        If cad <> "" Then cad = "(" & cadWHERE & ") AND (" & cad & ")"
    ElseIf OpcionMensaje = 14 Then
        cad = RegresarCargaEmpresas
    End If
    
    
    
     'Actualizar la tabla sseries asignando los valores correspondientes a los
      'campos: codclien, coddirec, tieneman, codtipom, numalbar, fechavta, numline1
      'y Salir (Volver a Mto Albaranes Clientes (Facturacion)
      PulsadoSalir = True
      'RaiseEvent CargarNumSeries
      RaiseEvent DatoSeleccionado(cad)
      Unload Me
End Sub


Private Sub cmdacepVariedades_Click()
Dim CADENA As String
    'Cargo las variedades marcadas
    CADENA = ""
    CategoriaValorNulo = False
    SeleccionadosTodos = True
    For NumRegElim = 1 To ListView7.ListItems.Count
        If ListView7.ListItems(NumRegElim).Checked Then
            If Label5.Caption = "Forfaits" Or Label5.Caption = "Categorias" Then
                If Trim(ListView7.ListItems(NumRegElim).Text) = "" Then
                    '[Monica]17/06/2013: solo para categorias que pueden ser null
                    CategoriaValorNulo = True
                Else
                    CADENA = CADENA & "'" & Trim(ListView7.ListItems(NumRegElim).Text) & "',"
                End If
            Else
                CADENA = CADENA & ListView7.ListItems(NumRegElim).Text & ","
            End If
        Else
            SeleccionadosTodos = False
        End If
    Next NumRegElim
    ' quitamos la ultima coma
    If CADENA <> "" Then
        CADENA = Mid(CADENA, 1, Len(CADENA) - 1)
    End If
    
    RaiseEvent DatoSeleccionado(CADENA)
    Unload Me
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    RaiseEvent DatoSeleccionado("")
    Unload Me
End Sub


Private Sub cmdCancelarCobros_Click()
    vCampos = "0"
    Unload Me
End Sub

Private Sub CmdCancelFactACta_Click()
    RaiseEvent DatoSeleccionado("")
    Unload Me
End Sub

Private Sub CmdCancelFrasAnecoop_Click()
    RaiseEvent DatoSeleccionado("")
    Unload Me
End Sub

Private Sub cmdCancelPMP_Click(Index As Integer)
    PulsadoSalir = True
    RaiseEvent DatoSeleccionado("")
    Unload Me
End Sub

Private Sub CmdCanPal_Click()
    RaiseEvent DatoSeleccionado("0")
    Unload Me
End Sub

Private Sub CmdCanProductos_Click()
    RaiseEvent DatoSeleccionado("")
    Unload Me
End Sub

Private Sub cmdCanVariedades_Click()
    RaiseEvent DatoSeleccionado("")
    Unload Me
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

'---monica
'Private Sub cmdCorrecotrPrecios_Click(Index As Integer)
'Dim SQL As String
'
'
'    If Index = 0 Then
'
'
'        'Compruebo si ha seleccionado algun articulo de los de precio ultima compra=0
'        cadWHERE2 = ""
'        SQL = ""
'        For TotalArray = 1 To Me.ListView4.ListItems.Count
'            If ListView4.ListItems(TotalArray).Checked Then
'                If ListView4.ListItems(TotalArray).Tag = "" Then
'                    SQL = SQL & "M"
'                Else
'                    cadWHERE2 = cadWHERE2 & "M"
'                End If
'            End If
'        Next
'
'        If SQL <> "" Then
'            MsgBox "No puede actualizar los articulos cuyo precio ultima compra sea 0", vbExclamation
'            Exit Sub
'        End If
'
'        If cadWHERE2 = "" Then
'            MsgBox "Seleccione algun articulo para actualizar", vbExclamation
'            Exit Sub
'        End If
'
'        'Llegado aqui todo correcto. Hacemos la pregunta de actualizar y a correr
'        SQL = "artículo"
'        If Len(cadWHERE2) > 1 Then SQL = SQL & "s"
'        SQL = "Va a actualizar los precios de " & Len(cadWHERE2) & " " & SQL & vbCrLf & vbCrLf & "¿Desea continuar?"
'        If MsgBox(SQL, vbQuestion + vbYesNo) <> vbYes Then Exit Sub
'
'
'        'Aqui esta el proceso de actualizacion de articulos
'        Me.lblIndicadorCorregir.Caption = "Actualización precios"
'        Me.Refresh
'        espera 0.5
'
'       'Para el LOG
'       SQL = cadWHERE & vbCrLf
'       For TotalArray = 1 To Me.ListView4.ListItems.Count
'            If ListView4.ListItems(TotalArray).Checked Then
'                If ListView4.ListItems(TotalArray).Tag <> "" Then SQL = SQL & ListView4.ListItems(TotalArray).Text & "|"
'            End If
'        Next
'        SQL = Mid(SQL, 1, 237)
'
'        '------------------------------------------------------------------------------
'        '  LOG de acciones
'        Set Log = New cLOG
'        Log.Insertar 4, vUsu, "Correccion precios: " & vbCrLf & SQL
'        Set Log = Nothing
'        '-----------------------------------------------------------------------------
'
'
'
'
'
'
'
'
'
'
'        For TotalArray = 1 To Me.ListView4.ListItems.Count
'            If ListView4.ListItems(TotalArray).Checked Then
'                If ListView4.ListItems(TotalArray).Tag <> "" Then
'
'                    'lo metemos en transaccion. Si queremos vamos
'                    Me.lblIndicadorCorregir.Caption = ListView4.ListItems(TotalArray).Text
'                    Me.lblIndicadorCorregir.Refresh
'
'                    ActualizaPrecios TotalArray
'
'
'                End If
'            End If
'        Next
'
'
'    End If
'    Unload Me
'End Sub
'

Private Function ActualizaPrecios(NumeroItem As Integer) As Boolean

On Error GoTo EActualizaPrecios
    ActualizaPrecios = False
    With ListView4.ListItems(NumeroItem)
        If Me.cmbActualizarTar.ListIndex <> 2 Then
            cadWHERE2 = TransformaComasPuntos(CStr(ImporteFormateado(.SubItems(7))))
            cadWHERE2 = "UPDATE sartic set preciove=" & cadWHERE2 & " WHERE codartic = '" & ListView4.ListItems(NumeroItem).Tag & "'"
            conn.Execute cadWHERE2
        End If
        If Me.cmbActualizarTar.ListIndex <> 1 Then
            cadWHERE2 = TransformaComasPuntos(CStr(ImporteFormateado(.SubItems(8))))
            cadWHERE2 = "UPDATE slista set precioac=" & cadWHERE2 & " WHERE codartic = '" & ListView4.ListItems(NumeroItem).Tag & "' AND codlista =" & vCampos
            conn.Execute cadWHERE2
        End If
    End With
        
    ActualizaPrecios = True
    Exit Function
EActualizaPrecios:
    MuestraError Err.Number, ListView4.ListItems(NumeroItem).Text
End Function


Private Sub cmdDeselTodos_Click()
Dim I As Byte

    For I = 1 To ListView2.ListItems.Count
        ListView2.ListItems(I).Checked = False
    Next I
End Sub




Private Sub cmdEtiqEstan_Click(Index As Integer)
    If Index = 1 Then
        'Cargo la tabla temporal con los datos que qeuremos imprimir
        cadWHERE2 = "insert into `tmpnseries` (`codusu`,`codartic`,`numlinealb`,`numlinea`) VALUES "
        cadWHERE = ""
        For NumRegElim = 1 To ListView3.ListItems.Count
            '                                                En el tag YA esta grabado
            If ListView3.ListItems(NumRegElim).Checked Then
                cadWHERE = cadWHERE & ",(" & vUsu.codigo & ",'" & ListView3.ListItems(NumRegElim).Tag & "',0,0)"
                If (NumRegElim Mod 25) = 0 Then
                    conn.Execute cadWHERE2 & Mid(cadWHERE, 2) & ";"
                    cadWHERE = ""
                    DoEvents
                End If
            End If
        Next NumRegElim
        If cadWHERE <> "" Then conn.Execute cadWHERE2 & Mid(cadWHERE, 2) & ";"
    Else
        NumRegElim = 0
    End If
    Unload Me
End Sub

Private Sub cmdMante_Click(Index As Integer)
Dim b As Boolean
    If Index = 0 Then
        
        
        If Val(txtMante(0).Text) = 0 Then
            MsgBox "El campo Año a traspasar debe ser numérico", vbExclamation
            Exit Sub
        End If
        
        
        If MsgBox("El proceso es irreversible. Continuar de igual modo?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
        '-------------------------------------------
        Screen.MousePointer = vbHourglass
        Set miRsAux = New ADODB.Recordset
        conn.BeginTrans
        b = TraspasarMantenimientos
        Set miRsAux = Nothing
        Screen.MousePointer = vbDefault
        If b Then
            conn.CommitTrans
        Else
            conn.RollbackTrans
        End If
        
        
    End If
    Unload Me
End Sub

Private Sub CmdPedSinAlb_Click()
    Dim I As Byte

    If Not ListView6.SelectedItem Is Nothing Then
        RaiseEvent DatoSeleccionado(ListView6.SelectedItem)
    Else
        RaiseEvent DatoSeleccionado("")
    End If
    Unload Me
End Sub

Private Sub cmdSelTodos_Click()
    Dim I As Byte

    For I = 1 To ListView2.ListItems.Count
        ListView2.ListItems(I).Checked = True
    Next I
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
Dim OK As Boolean

    
    Select Case OpcionMensaje
        Case 4 'Mostrar Nº Series
            If PrimeraVez Then
                PrimeraVez = False
                Me.Refresh
                Screen.MousePointer = vbHourglass
                OK = ObtenerTamanyosArray
                If OK Then OK = SeparaCampos
                If Not OK Then
                    'Error en SQL
                    'Salimos
                    Unload Me
                    Exit Sub
                End If
                CargarListaNSeries
            End If
            
        Case 8, 9, 17 'Etiquetas de clientes/Proveedores
            CargarListaClientes
'        Case 10 'Errores al contabilizar facturas
'            CargarListaErrContab
        Case 11 'Lineas Factura a rectificar
            CargarListaLinFactu
            
        Case 14 'Mostrar Empresas del sistema
            CargarListaEmpresas
            
        Case 15
            'Etiquetas estanteria
            CargarArticulosEstanteria
            
        Case 16
            'Articulos para corregir
            CargarArticulosCorreccionPrecio
            
            If Me.ListView4.ListItems.Count = 0 Then
                MsgBox "Ningún dato para mostrar", vbExclamation
                Unload Me
            End If
        Case 18
            PonerFoco txtMante(0)
        
        Case 21  'Variedades viene de un rango de clases
            CargarListaFields False
        
        Case 22 ' facturas a cuenta del cliente
            CargarListaFacturas
        
        Case 23 ' Lineas de calibre del albaran de venta
            CargarListaCalibres
        
        Case 24, 25
            CargaLwPrecioMP
            
        Case 28 ' productos de la misma clase
            CargarListaProductos
            
        Case 29 ' facturas de anecoop
            CargarFacturasAnecoop
    End Select
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim cad As String
On Error Resume Next

    FrameCobrosPtes.visible = False
    FrameAcercaDe.visible = False
    FrameNSeries.visible = False
    FrameComponentes.visible = False
    FrameComponentes2.visible = False
    FrameErrores.visible = False
    FrameEtiqEstant.visible = False
    FrameCorreccionPrecios.visible = False
    FrameTraspasoMante.visible = False
    FramePaletsAsociados.visible = False
    FramePedidosSinAlbaran.visible = False
    FrameVariedades.visible = False
    FrameFacturasACuenta.visible = False
    FrameLineasCalibre.visible = False
    FramePMP.visible = False
    frameClaveAcceso.visible = False
    FrameProductos.visible = False
    FrameAnecoop.visible = False

    PulsadoSalir = True
    PrimeraVez = True
    
    
    Select Case OpcionMensaje
        Case 1 'Mensaje de Cobros Pendientes
            PonerFrameCobrosPtesVisible True, H, W
            CargarListaCobrosPtes
            Me.Caption = "Cobros Pendientes"
            PonerFocoBtn Me.cmdAceptarCobros
            
        Case 2 'Mensaje de no hay suficiente Stock
            PonerFrameCobrosPtesVisible True, H, W
            CargarListaArtSinStock (vCampos)
            Me.Caption = "Artículos sin stock suficiente"
            PonerFocoBtn Me.cmdAceptarCobros
            
        Case 3 'Mensaje ACERCA DE
            CargaImagen
            Me.Caption = "Acerca de ....."
            PonerFrameAcercaDeVisible True, H, W
            Me.lblVersion.Caption = "Versión:  " & App.Major & "." & App.Minor & "." & App.Revision & " "
        
        Case 4 'Listado Nº Series Articulo
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "Nº Serie"
            Me.Label7(1).Caption = "Seleccione los Nº de serie para el Albaran."
            Me.Label7(1).FontSize = 12
            PulsadoSalir = False
            
'        Case 5 'Seleccionar tipo de Componente que queremos mostrar en Resumen
'                'En mant. de Nº Series de Reparacion
'            ponerFrameComponentesVisible True, h, w
'            Me.Caption = "Componentes"
'            Me.OptCompXMant.Value = True
'            PonerFocoBtn Me.cmdAceptarComp
        
        Case 6 'Mostrar Prefacturacion de Albaranes
            PonerFrameCobrosPtesVisible True, H, W
            CargarListaPreFacturar
            Me.Caption = "Prefacturación Albaranes"
            cad = RecuperaValor(vCampos, 1)
            If cad <> "" Then cad = Mid(cad, 1, Len(cad) - 1)
            Me.txtParam.Text = cad
            cad = RecuperaValor(vCampos, 2)
            If cad <> "" Then
                cad = Mid(cad, 1, Len(cad) - 1)
                If Trim(Me.txtParam.Text) <> "" Then
                    txtParam.Text = Me.txtParam.Text & vbCrLf & cad
                Else
                    txtParam.Text = cad
                End If
            End If
            cad = RecuperaValor(vCampos, 3)
            If cad <> "" Then
                cad = Mid(cad, 1, Len(cad) - 1)
                If Trim(Me.txtParam.Text) <> "" Then
                    txtParam.Text = Me.txtParam.Text & vbCrLf & cad
                Else
                    txtParam.Text = cad
                End If
            End If
            
            PonerFocoBtn Me.cmdAceptarComp
            
        Case 8, 17 'Etiquetas de Clientes
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "Clientes"
            PulsadoSalir = False
            Me.cmdSelTodos.visible = True
            Me.cmdDeselTodos.visible = True
            
        Case 9 'Etiquetas de Proveedores
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "Proveedores"
            PulsadoSalir = False
            Me.cmdSelTodos.visible = True
            Me.cmdDeselTodos.visible = True
        
        Case 10 'Errores al contabilizar facturas
            PonerFrameCobrosPtesVisible True, H, W
            CargarListaErrContab
            Me.Caption = "Facturas NO contabilizadas: "
            PonerFocoBtn Me.cmdAceptarCobros
        
        Case 11 'Lineas Factura a Rectificar
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "Lineas Factura a Rectificar"
            PulsadoSalir = False
            Me.cmdSelTodos.visible = True
            Me.cmdDeselTodos.visible = True
            Me.cmdAceptarNSeries.Left = Me.cmdAceptarNSeries.Left + 1000
            Me.cmdCancelar.Left = Me.cmdCancelar.Left + 1000
        
        Case 12 'Mensaje Albaranes que no se van a Facturar
            PonerFrameCobrosPtesVisible True, H, W
            CargarListaAlbaranes
            Me.Caption = "Facturación Albaranes"
            Me.Label1(0).Caption = "Existen Albaranes que NO se van a Facturar:"
            Me.Label1(0).Top = 260
            Me.Label1(0).Left = 480
            PonerFocoBtn Me.cmdAceptarCobros
            
        Case 13 'Muestra Errores
            H = 6000
            W = 8800
            PonerFrameVisible Me.FrameErrores, True, H, W
            Me.text1.Text = vCampos
            Me.Caption = "Errores"
        
        Case 14 'Muestra Empresas del sistema
            PonerFrameNSeriesVisible True, H, W
            Me.Caption = "Selección"
            CargarListaEmpresas
        Case 15
            H = FrameEtiqEstant.Height
            W = FrameEtiqEstant.Width
            PonerFrameVisible FrameEtiqEstant, True, H, W
            
        Case 16
            Caption = "Corrección precios"
            H = FrameCorreccionPrecios.Height
            W = FrameCorreccionPrecios.Width
            PonerFrameVisible FrameCorreccionPrecios, True, H, W
            Me.cmdCorrecotrPrecios(1).Cancel = True
            lblIndicadorCorregir.Caption = ""
            
        Case 18
            
            Caption = "Mantenimientos"
            H = FrameTraspasoMante.Height
            W = FrameTraspasoMante.Width
            PonerFrameVisible FrameTraspasoMante, True, H, W

        Case 19 'Palets asociados al pedido del que se va a generar el albaran
            
            PonerFramePaletsAsociadosVisible True, H, W
            CargarListaPalets
            Me.Caption = "Palets Asociados al Pedido"
            PonerFocoBtn Me.CmdAceptarPal
    
        Case 20 'Pedidos sin nro de albaran asociado
            
            PonerFramePedidosSinAlbaranVisible True, H, W
            CargarListaPedidosSinAlbaran
            Me.Caption = "Pedidos sin Albarán Asignado"
            PonerFocoBtn Me.CmdPedSinAlb
    
        Case 21, 26 'variedades
            H = FrameVariedades.Height
            W = FrameVariedades.Width
            PonerFrameVisible FrameVariedades, True, H, W
                
        Case 22 'facturas a cuenta
            H = FrameFacturasACuenta.Height
            W = FrameFacturasACuenta.Width
            PonerFrameVisible FrameFacturasACuenta, True, H, W
    
        Case 23 ' Lineas de calibre de un albaran
            H = FrameLineasCalibre.Height
            W = FrameLineasCalibre.Width
            PonerFrameVisible FrameLineasCalibre, True, H, W
        
        Case 24, 25
            'Ajuste precio
            '   24: PrecioMP
            '   25: PrecioMP
            H = Me.FramePMP.Height
            W = Me.FramePMP.Width
            
            PonerFrameVisible FramePMP, True, H, W
            If OpcionMensaje = 24 Then
                Caption = "Actualizar precio medio ponderado"
            Else
                Caption = "Actualizar último precio compra"
            End If
            CadenaDesdeOtroForm = ""
        
        Case 27 ' clave de acceso
            H = frameClaveAcceso.Height
            W = frameClaveAcceso.Width
            PonerFrameVisible frameClaveAcceso, True, H, W

        Case 28 ' lista de productos del grupo
            H = FrameProductos.Height
            W = FrameProductos.Width
            PonerFrameVisible FrameProductos, True, H, W
        
        Case 29 ' lista de facturas para integrar en recoleccion
            H = FrameAnecoop.Height
            W = FrameAnecoop.Width
            PonerFrameVisible FrameAnecoop, True, H, W
        
    End Select
    'Me.cmdCancel(indFrame).Cancel = True
    Me.Height = H + 350
    Me.Width = W + 70
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PonerFrameCobrosPtesVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Cobros Pendientes Visible y Ajustado al Formulario, y visualiza los controles
'necesario para el Informe

    H = 4600
        
    Select Case OpcionMensaje
        Case 1
            H = 5000
            W = 8600
            Me.Label1(0).Caption = "CLIENTE: " & vCampos
        Case 2
            W = 8800
            Me.cmdAceptarCobros.Top = 4000
            Me.cmdAceptarCobros.Left = 4200
        Case 5 'Componentes
            W = 6000
            H = 5000
            Me.cmdAceptarCobros.Left = 4000

        Case 6, 7 'Prefacturar Albaranes
            W = 7000
            H = 6000
            Me.cmdAceptarCobros.Top = 5400
            Me.cmdAceptarCobros.Left = 4600

        Case 10, 12 'Errores al contabilizar facturas
            H = 6000
            W = 8400
            Me.cmdAceptarCobros.Top = 5300
            Me.cmdAceptarCobros.Left = 4900
            If OpcionMensaje = 12 Then
                Me.cmdCancelarCobros.Top = 5300
                Me.cmdCancelarCobros.Left = 4600
                Me.cmdAceptarCobros.Left = 3300
                Me.Label1(1).Top = 4800
                Me.Label1(1).Left = 3400
                Me.cmdAceptarCobros.Caption = "&SI"
                Me.cmdCancelarCobros.Caption = "&NO"
            End If
    End Select
            
    PonerFrameVisible Me.FrameCobrosPtes, visible, H, W

    If visible = True Then
        Me.txtParam.visible = (OpcionMensaje = 6 Or OpcionMensaje = 7)
        Me.Label1(0).visible = (OpcionMensaje = 1) Or (OpcionMensaje = 5) Or (OpcionMensaje = 12)
        Me.cmdCancelarCobros.visible = (OpcionMensaje = 12)
        Me.Label1(1).visible = (OpcionMensaje = 12)
    End If
End Sub

Private Sub PonerFramePaletsAsociadosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Cobros Pendientes Visible y Ajustado al Formulario, y visualiza los controles
'necesario para el Informe

        
    H = 6000
    W = 8400
    Me.CmdAceptarPal.Top = 5300
    Me.CmdAceptarPal.Left = 4900
    Me.CmdCanPal.Top = 5300
    Me.CmdCanPal.Left = 4600
    Me.CmdAceptarPal.Left = 3300
    Me.Label1(2).Top = 4800
    Me.Label1(2).Left = 3400
    Me.Label1(3).Caption = "Nº Pedido : " & vCampos
    Me.CmdAceptarPal.Caption = "&SI"
    Me.CmdCanPal.Caption = "&NO"
        
    PonerFrameVisible Me.FramePaletsAsociados, visible, H, W

End Sub


Private Sub PonerFrameAcercaDeVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame ACERCA DE visible y Ajustado al Formulario

    Me.FrameAcercaDe.visible = visible
    If visible = True Then
        'Ajustar Tamaño del Frame para ajustar tamaño de Formulario al del Frame
        Me.FrameAcercaDe.Top = -90
        Me.FrameAcercaDe.Left = 0
        Me.FrameAcercaDe.Height = 4555
        Me.FrameAcercaDe.Width = 6600
        
        W = Me.FrameAcercaDe.Width
        H = Me.FrameAcercaDe.Height
    End If
End Sub


Private Sub PonerFrameNSeriesVisible(visible As Boolean, H As Integer, W As Integer)
'Pone el Frame de Nº Serie Visible y Ajustado al Formulario, y visualiza los controles
'necesario para el Informe

    H = 5000
   
    If OpcionMensaje = 11 Then 'Lineas Factura a Rectificar
        W = 10900
    ElseIf OpcionMensaje = 14 Then
        W = 6500
        Me.Label7(1).visible = True
    Else
        W = 8500
        Me.Label7(1).visible = False
    End If
    PonerFrameVisible Me.FrameNSeries, visible, H, W
End Sub

Private Sub PonerFramePedidosSinAlbaranVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Cobros Pendientes Visible y Ajustado al Formulario, y visualiza los controles
'necesario para el Informe

        
    H = 4620
    W = 8655
        
    PonerFrameVisible Me.FramePedidosSinAlbaran, visible, H, W

End Sub

'Private Sub ponerFrameComponentesVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
''Pone el Frame de Componentes Visible y Ajustado al Formulario, y visualiza los controles
''necesario para el Informe
'
''    Me.FrameComponentes.visible = visible
'    Me.FrameComponentes2.visible = visible
'
'    h = 4000
'    w = 5300
'    PonerFrameVisible Me.FrameComponentes, visible, h, w
'
'    If visible = True Then
'        'Ajustar Tamaño del Frame para ajustar tamaño de Formulario al del Frame
'        If vParamAplic.Departamento Then
'            Me.OptCompXDpto.Caption = "Departemento"
'        Else
'            Me.OptCompXDpto.Caption = "Dirección"
'        End If
'    End If
'End Sub


Private Sub CargarListaCobrosPtes()
'Muestra la lista Detallada de cobros en un ListView
'Carga los valores de la tabla scobro de la Contabilidad
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String

    Sql = "SELECT numserie, codfaccl, fecfaccl, fecvenci, impvenci, impcobro "
    Sql = Sql & " FROM scobro INNER JOIN sforpa ON scobro.codforpa=sforpa.codforpa "
    Sql = Sql & cadWHERE

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
     
    ListView1.Top = 900
    ListView1.Height = 3250
    ListView1.Width = 8100
    ListView1.Left = 160
    
    'Los encabezados
    ListView1.ColumnHeaders.Clear
    
    ListView1.ColumnHeaders.Add , , "Nº Serie", 760
    ListView1.ColumnHeaders.Add , , "Nº Factura", 1100, 1
    ListView1.ColumnHeaders.Add , , "Fecha Factura", 1250, 2
    ListView1.ColumnHeaders.Add , , "Fecha Venci.", 1200, 2
    ListView1.ColumnHeaders.Add , , "Imp. Venci.(€)", 1250, 1
    ListView1.ColumnHeaders.Add , , "Imp. Cobro(€)", 1250, 1
    ListView1.ColumnHeaders.Add , , "Pte. Cobro(€)", 1250, 1
    
    While Not Rs.EOF
        Set ItmX = ListView1.ListItems.Add
        ItmX.Text = Rs.Fields(0).Value 'Nº Serie
        ItmX.SubItems(1) = Rs.Fields(1).Value 'Nº Factura
        ItmX.SubItems(2) = Rs.Fields(2).Value 'Fecha Factura
        ItmX.SubItems(3) = Rs.Fields(3).Value 'Fecha Vencimiento
        ItmX.SubItems(4) = Rs.Fields(4).Value 'Importe Vencido
        ItmX.SubItems(5) = DBLet(Rs.Fields(5).Value, "N") 'Importe Cobrado
        ItmX.SubItems(6) = Rs.Fields(4).Value - DBLet(Rs.Fields(5).Value, "N") 'Pendiente de cobro
        If ItmX.SubItems(6) > 0 Then
            ItmX.ListSubItems(6).ForeColor = vbRed
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
End Sub


Private Sub CargarListaArtSinStock(NomTabla As String)
'Muestra la lista Detallada de Articulos que no tienen stock suficiente en un ListView
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String

    Sql = "SELECT " & NomTabla & ".codalmac," & NomTabla & ".codartic, " & NomTabla & ".nomartic, salmac.canstock as canstock, SUM(cantidad) as cantidad, canstock-SUM(cantidad) as disp "
    Sql = Sql & "FROM ((" & NomTabla & " INNER JOIN sartic ON " & NomTabla & ".codartic=sartic.codartic) INNER JOIN sfamia ON sartic.codfamia=sfamia.codfamia) "
    Sql = Sql & "INNER JOIN salmac ON " & NomTabla & ".codalmac=salmac.codalmac and " & NomTabla & ".codartic=salmac.codartic "
    Sql = Sql & cadWHERE 'Where numpedcl = 2 And sfamia.instalac = 0
    Sql = Sql & "GROUP by " & NomTabla & ".codalmac, " & NomTabla & ".codartic "
    

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
     
    Me.ListView1.Top = 500
     
    'Los encabezados
    ListView1.Width = 8400
    ListView1.Height = 3150
    ListView1.ColumnHeaders.Clear
    
    ListView1.ColumnHeaders.Add , , "Alm.", 500
    ListView1.ColumnHeaders.Add , , "Articulo", 1800, 2
    ListView1.ColumnHeaders.Add , , "Dec. Artic", 3300
    ListView1.ColumnHeaders.Add , , "Stock", 950, 2
    ListView1.ColumnHeaders.Add , , "Cantidad", 900, 2
    ListView1.ColumnHeaders.Add , , "No Disp.", 900, 2
    
    While Not Rs.EOF
        If Rs!disp < 0 Then
            Set ItmX = ListView1.ListItems.Add
            ItmX.Text = Format(Rs.Fields(0).Value, "000") 'Cod Almacen
            ItmX.SubItems(1) = Rs.Fields(1).Value 'Cod Artic
            ItmX.SubItems(2) = Rs.Fields(2).Value 'Nom Artic
            ItmX.SubItems(3) = Rs.Fields(3).Value 'Stock
            ItmX.SubItems(4) = Rs.Fields(4).Value 'Cantidad
            ItmX.SubItems(5) = Rs.Fields(5).Value 'No Disp
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
End Sub


Private Sub CargarListaNSeries()
'Carga las lista con todos los Nº de serie encontrados en la tabla:sserie
'para el articulo pasado como parametro en la cadwhere: "codartic='00012'"
'y que esten disponibles: numfactu y numalbar no tengan valor
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String
Dim cadLista As String
Dim Dif As Single

    On Error GoTo ECargarLista

    If cadWHERE2 = "" Then
        'Mostramos los nº serie libres para seleccionar la cantidad
        Sql = "SELECT sserie.numserie, sserie.codartic, sartic.nomartic "
        Sql = Sql & "FROM sserie INNER JOIN sartic ON sserie.codartic=sartic.codartic "
        Sql = Sql & cadWHERE 'Where codartic='000012'
        'seleccionamos los que no esten asignados a ninguna factura ni albaran
        Sql = Sql & " AND ((isnull(sserie.numfactu) or sserie.numfactu='') and (isnull(sserie.numalbar) or sserie.numalbar='')) "
        Sql = Sql & " ORDER BY sserie.codartic, numserie "
        
    Else 'venimos de modificar la cantidad y seleccionamos los ya asignados
        If InStr(1, cadWHERE2, "|") > 0 Then
            Dif = CSng(RecuperaValor(cadWHERE2, 1))
            cadWHERE2 = RecuperaValor(cadWHERE2, 2)
        
            'seleccionamos nº serie del albaran que modificamos
            Sql = "SELECT sserie.numserie, sserie.codartic, sartic.nomartic "
            Sql = Sql & "FROM sserie INNER JOIN sartic ON sserie.codartic=sartic.codartic "
            Sql = Sql & cadWHERE2
                
            
            If Dif < 0 Then
                'Si la diferencia de cantidad es < 0, mostrar en la lista los nº serie que
                'tiene la linea de albaran asignado con todos marcados y desmarcar el que no queremos
                
            Else
                'si la diferencia de cantidad es > 0, mostrar en la lista los nº de serie que
                'ya tenia asignados la linea del albaran más los libres para seleccionar los que añadimos de mas
                cadLista = ""
                Set Rs = New ADODB.Recordset
                Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not Rs.EOF
                    cadLista = cadLista & ", " & Rs!numserie
                    Rs.MoveNext
                Wend
                Rs.Close
                Set Rs = Nothing
                
                'mostrar tambien los nº serie sin asignar
                Sql = Sql & " OR (" & Replace(cadWHERE, "WHERE", "") & " and (numalbar=''or isnull(numalbar)))"
            End If
        Else
            'viene de una factura rectificativa, seleccionamos los nº de serie de
            'esa factura y marcamos los que queremos quitar
            Sql = cadWHERE2
        End If
    End If
    

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Los encabezados
    ListView2.Width = 7400
    Me.ListView2.Height = 3100
    Me.ListView2.Left = 650
    ListView2.ColumnHeaders.Clear
    
    ListView2.ColumnHeaders.Add , , "Nº Serie", 1800
    ListView2.ColumnHeaders.Add , , "Articulo", 1800
    ListView2.ColumnHeaders.Add , , "Desc. Artic", 3650
        
    If Rs.EOF Then Unload Me
    
    While Not Rs.EOF
         Set ItmX = ListView2.ListItems.Add
         ItmX.Text = Rs.Fields(0).Value 'num serie
         If Dif < 0 Then
            ItmX.Checked = True
         ElseIf Dif > 0 Then
            If InStr(1, cadLista, CStr(Rs!numserie)) > 0 Then
                ItmX.Checked = True
            Else
                ItmX.Checked = False
            End If
         Else
            ItmX.Checked = False
         End If
         ItmX.SubItems(1) = Rs.Fields(1).Value 'Desc Artic
         ItmX.SubItems(2) = Rs.Fields(2).Value 'Nom Artic
         Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
ECargarLista:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar Nº Series", Err.Description
End Sub


'Private Sub CargarListaComponentes(opt As Byte)
''Muestra la lista Detallada de cobros en un ListView
''Carga los valores de la tabla scobro de la Contabilidad
'Dim RS As ADODB.Recordset
'Dim ItmX As ListItem
'Dim SQL As String
'Dim Codigo As String, cadCodigo As String
'
'    Select Case opt
'        Case 1 'Mantenimiento
'            Codigo = RecuperaValor(vCampos, 1)
'            If Codigo = "" Then
'                cadCodigo = " isnull(nummante) "
'            Else
'                cadCodigo = " nummante=" & DBSet(Codigo, "T")
'            End If
'            SQL = ObtenerSQLcomponentes(cadWHERE & " and " & cadCodigo)
'            Me.Label1(0).Caption = "Mantenimiento: " & Codigo
'
'        Case 2 'Departamento
'            Codigo = RecuperaValor(vCampos, 2)
'            If Codigo = "" Then
'                cadCodigo = "isnull(coddirec)"
'            Else
'                cadCodigo = " coddirec=" & Codigo
'            End If
'            SQL = ObtenerSQLcomponentes(cadWHERE & " and " & cadCodigo)
'            If vParamAplic.Departamento Then
'                Me.Caption = "Equipos del Departamento"
'                Me.Label1(0).Caption = " Departamento: " & RecuperaValor(vCampos, 3)
'            Else
'                Me.Caption = "Equipos de la Dirección"
'                Me.Label1(0).Caption = " Dirección: " & Codigo & " " & RecuperaValor(vCampos, 3)
'            End If
'
'        Case 3 'Cliente
'            SQL = ObtenerSQLcomponentes(cadWHERE)
'            Me.Caption = "Equipos del Cliente"
'            Me.Label1(0).Caption = "Cliente: " & RecuperaValor(vCampos, 4)
'    End Select
'
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    'Los encabezados
'    ListView1.Top = 800
'    ListView1.Left = 280
'    ListView1.Width = 4900
'    ListView1.Height = 3250
'    ListView1.ColumnHeaders.Clear
'
'    ListView1.ColumnHeaders.Add , , "TA", 760
'    ListView1.ColumnHeaders.Add , , "Tipo Articulo", 2800
'    ListView1.ColumnHeaders.Add , , "Cantidad", 1280, 2
'
'    If Not RS.EOF Then
'        While Not RS.EOF
'            Set ItmX = ListView1.ListItems.Add
'            ItmX.Text = RS.Fields(0).Value 'TA
'            ItmX.SubItems(1) = RS.Fields(1).Value 'Tipo Articulo
'            ItmX.SubItems(2) = RS.Fields(2).Value 'Cantidad
'            RS.MoveNext
'        Wend
'    End If
'    RS.Close
'    Set RS = Nothing
'End Sub


Private Sub CargarListaPreFacturar()
'Muestra la lista Detallada de Albaranes a Factura en un ListView
'Carga los valores de la tabla scobro de la Contabilidad
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String

    On Error GoTo ECargarList
    
    Sql = "CREATE TEMPORARY TABLE tmp ( "
    Sql = Sql & "codforpa SMALLINT(3) UNSIGNED  DEFAULT '0' NOT NULL, "
    Sql = Sql & "numalbar MEDIUMINT(7) UNSIGNED  DEFAULT '0' NOT NULL, "
    Sql = Sql & "dtoppago DECIMAL(4,2) UNSIGNED  DEFAULT '0.0' NOT NULL, "
    Sql = Sql & "dtopgnral DECIMAL(4,2) UNSIGNED  DEFAULT '0.0' NOT NULL, "
    Sql = Sql & "importe DECIMAL(12,2) UNSIGNED  DEFAULT '0.0' NOT NULL, "
    Sql = Sql & "bruto DECIMAL(12,2) UNSIGNED  DEFAULT '0.0' NOT NULL) "
    conn.Execute Sql
     
'     SQL = "LOCK TABLES scaalb READ, slialb READ;"
'     Conn.Execute SQL
     
    Sql = "SELECT scaalb.codforpa, scaalb.numalbar, dtoppago, dtognral, round(sum(importel),2) as importe, round(sum(importel),2) - round(((round(sum(importel),2)*dtoppago)/100),2) - round(((round(sum(importel),2)*dtognral)/100),2) as bruto "
    Sql = Sql & " FROM (scaalb INNER JOIN slialb ON scaalb.codtipom=slialb.codtipom and scaalb.numalbar=slialb.numalbar) "
    Sql = Sql & " WHERE " & cadWHERE
    Sql = Sql & " GROUP BY scaalb.numalbar "
    Sql = Sql & " ORDER BY scaalb.codforpa, scaalb.numalbar "

    Sql = " INSERT INTO tmp " & Sql
    conn.Execute Sql
     
    Sql = " SELECT tmp.codforpa, sforpa.nomforpa, sum(tmp.bruto) as bruto"
    Sql = Sql & " FROM tmp, sforpa WHERE tmp.codforpa=sforpa.codforpa "
    Sql = Sql & " GROUP BY tmp.codforpa "
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        ListView1.Height = 3850
        ListView1.Width = 5400
        ListView1.Left = 500
        ListView1.Top = 1200
    '    ListView1.GridLines = False
    
        'Los encabezados
        ListView1.ColumnHeaders.Clear
        
        ListView1.ColumnHeaders.Add , , " Forma de Pago", 3300
        ListView1.ColumnHeaders.Add , , "Base Imp.(€)", 2020, 1
     
        While Not Rs.EOF
            Set ItmX = ListView1.ListItems.Add
            ItmX.Text = Format(Rs!Codforpa.Value, "000") & "  " & Rs!nomforpa.Value
            
            ItmX.SubItems(1) = Rs!Bruto
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing
    
    'Borrar la tabla temporal
    Sql = " DROP TABLE IF EXISTS tmp;"
    conn.Execute Sql

ECargarList:
    If Err.Number <> 0 Then
         'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmp;"
        conn.Execute Sql
'        SQL = "UNLOCK TABLES "
'        Conn.Execute SQL
    End If
End Sub


Private Sub CargarListaClientes()
'Carga las lista con todos los clientes seleccionados en la tabla:sclien
'para imprimir etiquetas, pasando como parametro la cadwhere
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String, Men As String

    On Error GoTo ECargarLista

    Select Case OpcionMensaje
    Case 8
        'CLIENTES
        Sql = "SELECT codclien,nomclien,nifclien "
        Sql = Sql & "FROM sclien "
        If cadWHERE <> "" Then Sql = Sql & " WHERE " & cadWHERE
        Sql = Sql & " ORDER BY codclien "
        Men = "Cliente"
    Case 9
        'PROVEEDORES
        Sql = "SELECT codprove,nomprove,nifprove "
        Sql = Sql & "FROM proveedor "
        If cadWHERE <> "" Then Sql = Sql & " WHERE " & cadWHERE
        Sql = Sql & " ORDER BY codprove "
        Men = "Proveedor"
    Case 17
        'CLIENTES MANTENIMIENTO
        Sql = cadWHERE
    End Select
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        'Los encabezados
        ListView2.Width = 7000
        ListView2.Top = 500
        ListView2.Height = 3620
        ListView2.ColumnHeaders.Clear
        
        ListView2.ColumnHeaders.Add , , Men, 1350
        ListView2.ColumnHeaders.Add , , "Nombre", 4000
        ListView2.ColumnHeaders.Add , , "NIF", 1330
        
        While Not Rs.EOF
             Set ItmX = ListView2.ListItems.Add
             ItmX.Text = Format(Rs.Fields(0).Value, "000000") 'cod clien/prove
             ItmX.Checked = False
             ItmX.SubItems(1) = Rs.Fields(1).Value 'Nom clien/prove
             ItmX.SubItems(2) = Rs.Fields(2).Value 'NIF clien/prove
             Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing
    
ECargarLista:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargar " & Men, Err.Description
        PulsadoSalir = True
        Unload Me
    End If
End Sub



Private Sub CargarListaErrContab()
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String

    On Error GoTo ECargarList

    Sql = " SELECT  * "
    Sql = Sql & " FROM tmpErrFac "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        ListView1.Height = 4500
        ListView1.Width = 7400
        ListView1.Left = 500
        ListView1.Top = 500

        'Los encabezados
        ListView1.ColumnHeaders.Clear

        If Rs.Fields(0).Name = "codprove" Then
            'Facturas de Compra
             ListView1.ColumnHeaders.Add , , "Prove.", 700
        Else 'Facturas de Venta
            ListView1.ColumnHeaders.Add , , "Tipo", 600
        End If
        ListView1.ColumnHeaders.Add , , "Factura", 1000, 1
        ListView1.ColumnHeaders.Add , , "Fecha", 1100, 1
        ListView1.ColumnHeaders.Add , , "Error", 4620
    
        While Not Rs.EOF
            Set ItmX = ListView1.ListItems.Add
            'El primer campo será codtipom si llamamos desde Ventas
            ' y será codprove si llamamos desde Compras
            ItmX.Text = Rs.Fields(0).Value
            ItmX.SubItems(1) = Format(Rs!NumFactu, "0000000")
            ItmX.SubItems(2) = Rs!FecFactu
            ItmX.SubItems(3) = Rs!Error
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing

ECargarList:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub


Private Sub CargarListaLinFactu()
'Carga las lista con todas las lineas de la factura que estamos rectificando
'seleccionamos las que nos queremos llevar al Albaran de rectificacion
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String

    On Error GoTo ECargarLista

    Sql = "SELECT codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre"
    Sql = Sql & " FROM slifac "
    If cadWHERE <> "" Then Sql = Sql & " WHERE " & cadWHERE
    Sql = Sql & " ORDER BY codtipom,numfactu,fecfactu,codtipoa,numalbar,numlinea "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        
        ListView2.Top = 500
        ListView2.Left = 380
        ListView2.Width = 10100
        ListView2.Height = 3620
        
        'Los encabezados
        ListView2.ColumnHeaders.Clear
    
        ListView2.ColumnHeaders.Add , , "T.Alb", 660
        ListView2.ColumnHeaders.Add , , "Nº Alb", 840
        ListView2.ColumnHeaders.Add , , "Lin.", 450
         ListView2.ColumnHeaders.item(3).Alignment = lvwColumnRight
        ListView2.ColumnHeaders.Add , , "Alm", 460
        ListView2.ColumnHeaders.Add , , "Artic", 1380
        ListView2.ColumnHeaders.Add , , "Desc. Artic.", 2500
        ListView2.ColumnHeaders.Add , , "Cant.", 600
        ListView2.ColumnHeaders.item(7).Alignment = lvwColumnRight
        ListView2.ColumnHeaders.Add , , "Precio", 960
        ListView2.ColumnHeaders.item(8).Alignment = lvwColumnRight
        ListView2.ColumnHeaders.Add , , "Dto 1", 600
        ListView2.ColumnHeaders.item(9).Alignment = lvwColumnRight
        ListView2.ColumnHeaders.Add , , "Dto 2", 600
        ListView2.ColumnHeaders.item(10).Alignment = lvwColumnRight
        ListView2.ColumnHeaders.Add , , "Importe", 950
        ListView2.ColumnHeaders.item(11).Alignment = lvwColumnRight
    
        While Not Rs.EOF
             Set ItmX = ListView2.ListItems.Add
             ItmX.Text = Rs!codtipoa 'cod tipo alb
             ItmX.Checked = False
             ItmX.SubItems(1) = Format(Rs!NumAlbar, "0000000") 'Nº Albaran
             ItmX.SubItems(2) = Rs!numlinea 'linea Albaran
             ItmX.SubItems(3) = Format(Rs!codAlmac, "000") 'cod almacen
             ItmX.SubItems(4) = Rs!codArtic 'Cod Articulo
             ItmX.SubItems(5) = Rs!NomArtic 'Nombre del Articulo
             ItmX.SubItems(6) = Rs!Cantidad
             ItmX.SubItems(7) = Format(Rs!precioar, FormatoPrecio)
             ItmX.SubItems(8) = Rs!dtoline1
             ItmX.SubItems(9) = Rs!dtoline2
             ItmX.SubItems(10) = Format(Rs!ImporteL, FormatoImporte)
             Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing
    
    'si aparece la barra de desplazamiento ajustar el ancho
    If Me.ListView2.ListItems.Count > 11 Then
        Me.ListView2.ColumnHeaders(5).Width = 1200 'codartic
        Me.ListView2.ColumnHeaders(8).Width = 920  'precio
    End If
   
    
    
ECargarLista:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargar Lineas Factura", Err.Description
        PulsadoSalir = True
        Unload Me
    End If
End Sub




Private Sub CargarListaAlbaranes()
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String

    On Error GoTo ECargarList

    Sql = cadWHERE 'cadwhere ya le pasamos toda la SQL
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        ListView1.Height = 3900
        ListView1.Width = 7200
        ListView1.Left = 500
        ListView1.Top = 700

        'Los encabezados
        ListView1.ColumnHeaders.Clear

        ListView1.ColumnHeaders.Add , , "Tipo", 700
        ListView1.ColumnHeaders.Add , , "Nº Albaran", 1000, 1
        ListView1.ColumnHeaders.Add , , "Fecha", 1100, 1
        ListView1.ColumnHeaders.item(3).Alignment = lvwColumnCenter
        ListView1.ColumnHeaders.Add , , "Cod. Cli.", 900
        ListView1.ColumnHeaders.Add , , "Cliente", 3400
    
        While Not Rs.EOF
            Set ItmX = ListView1.ListItems.Add
            ItmX.Text = Rs.Fields(0).Value
            ItmX.SubItems(1) = Format(Rs!NumAlbar, "0000000")
            ItmX.SubItems(2) = Rs!FechaAlb
            ItmX.SubItems(3) = Format(Rs!CodClien, "000000")
            ItmX.SubItems(4) = Rs!Nomclien
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing

ECargarList:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub

Private Sub CargarListaPalets()
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String

    On Error GoTo ECargarList

    Sql = cadWHERE 'cadwhere ya le pasamos toda la SQL
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        ListView5.Height = 3900
        ListView5.Width = 7200
        ListView5.Left = 500
        ListView5.Top = 700

        'Los encabezados
        ListView5.ColumnHeaders.Clear

        ListView5.ColumnHeaders.Add , , "Nº Palet", 1000
        ListView5.ColumnHeaders.Add , , "Lin.Conf.", 1000, 1
        ListView5.ColumnHeaders.Add , , "F.Inicio", 1100, 1
        ListView5.ColumnHeaders.item(3).Alignment = lvwColumnCenter
        ListView5.ColumnHeaders.Add , , "Hora ", 900
        ListView5.ColumnHeaders.Add , , "F.Fin", 1100
        ListView5.ColumnHeaders.Add , , "Hora ", 900
        
    
        While Not Rs.EOF
            Set ItmX = ListView5.ListItems.Add
            ItmX.Text = Format(Rs!numpalet, "000000")
            ItmX.SubItems(1) = Format(Rs!linconfe, "00")
            ItmX.SubItems(2) = Rs!FechaIni
            ItmX.SubItems(3) = Format(Rs!HoraIni, "hh:mm:ss")
            ItmX.SubItems(4) = Rs!FechaFin
            ItmX.SubItems(5) = Format(Rs!HoraFin, "hh:mm:ss")
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing

ECargarList:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub


Private Sub CargarListaEmpresas()
'Carga las lista con todas las empresas que hay en el sistema
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String
Dim I As Integer

Dim Prohibidas As String

    On Error GoTo ECargarLista

    VerEmresasProhibidas Prohibidas
    
    Sql = "Select * from usuarios.empresasariagro order by codempre"
    Set ListView2.SmallIcons = frmPpal.ImageListB
    ListView2.Width = 5000
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "Empresa", 4900
    ListView2.HideColumnHeaders = True
    ListView2.GridLines = False
    ListView2.ListItems.Clear
    
    Set Rs = New ADODB.Recordset
    I = -1
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
        Sql = "|" & Rs!codempre & "|"
        If InStr(1, Prohibidas, Sql) = 0 Then
            Set ItmX = ListView2.ListItems.Add(, , Rs!nomempre, , 5)
            ItmX.Tag = Rs!codempre
            If ItmX.Tag = vEmpresa.codempre Then
                ItmX.Checked = True
                I = ItmX.Index
            End If
            ItmX.ToolTipText = Rs!Ariagro
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    If I > 0 Then Set ListView2.SelectedItem = ListView2.ListItems(I)

    
ECargarLista:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargando datos empresas", Err.Description
        PulsadoSalir = True
        Unload Me
    End If
End Sub


Private Sub VerEmresasProhibidas(ByRef VarProhibidas As String)
Dim Sql As String
Dim Rs As ADODB.Recordset

On Error GoTo EVerEmresasProhibidas
    VarProhibidas = "|"
    Sql = "Select codempre from usuarios.usuarioempresasariagro WHERE codusu = " & (vUsu.codigo Mod 1000)
    Sql = Sql & " order by codempre"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
          VarProhibidas = VarProhibidas & Rs!codempre & "|"
          Rs.MoveNext
    Wend
    Rs.Close
    Exit Sub
EVerEmresasProhibidas:
    MuestraError Err.Number, Err.Description & vbCrLf & " Consulte soporte técnico"
    Set Rs = Nothing
End Sub



Private Sub CargaImagen()
On Error Resume Next
    Image2.Picture = LoadPicture(App.path & "\minilogo.bmp")
    'Image1.Picture = LoadPicture(App.path & "\fondon.gif")
    Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If PulsadoSalir = False Then Cancel = 1
End Sub



Private Function ObtenerTamanyosArray() As Boolean
'Para el frame de los Nº de Serie de los Articulos
'En cada indice pone en CodArtic(i) el codigo del articulo
'y en Cantidad(i) la cantidad solicitada de cada codartic
Dim I As Integer, j As Integer

    ObtenerTamanyosArray = False
    'Primero a los campos de la tabla
    TotalArray = -1
    j = 0
    Do
        I = j + 1
        j = InStr(I, vCampos, "·")
        If j > 0 Then TotalArray = TotalArray + 1
    Loop Until j = 0
    
    If TotalArray < 0 Then Exit Function
    
    'Las redimensionaremos
    ReDim codArtic(TotalArray)
    ReDim Cantidad(TotalArray)
    
    ObtenerTamanyosArray = True
End Function


Private Function SeparaCampos() As Boolean
'Para el frame de los Nº de Serie de los Articulos
Dim Grupo As String
Dim I As Integer
Dim j As Integer
Dim C As Integer 'Contador dentro del array

    SeparaCampos = False
    I = 0
    C = 0
    Do
        j = I + 1
        I = InStr(j, vCampos, "·")
        If I > 0 Then
            Grupo = Mid(vCampos, j, I - j)
            'Y en la martriz
            InsertaGrupo Grupo, C
            C = C + 1
        End If
    Loop Until I = 0
    SeparaCampos = True
End Function


Private Sub InsertaGrupo(Grupo As String, Contador As Integer)
Dim j As Integer
Dim cad As String

    j = 0
    cad = ""
    
    'Cod Artic
    j = InStr(1, Grupo, "|")
    If j > 0 Then
        cad = Mid(Grupo, 1, j - 1)
        Grupo = Mid(Grupo, j + 1)
        j = 1
    End If
    codArtic(Contador) = cad
    
    'Cantidad
    j = InStr(1, Grupo, "|")
    If j > 0 Then
        cad = Mid(Grupo, 1, j - 1)
        Grupo = Mid(Grupo, j + 1)
    Else
        cad = Grupo
        Grupo = ""
    End If
    Cantidad(Contador) = cad
End Sub





Private Sub imgCheck_Click(Index As Integer)
Dim b As Boolean
    Select Case Index
        Case Is < 2
            'En el listview3
            b = Index = 1
            For TotalArray = 1 To ListView3.ListItems.Count
                ListView3.ListItems(TotalArray).Checked = b
                If (TotalArray Mod 50) = 0 Then DoEvents
            Next TotalArray
        Case 3
            'En el listview4
            b = Index = 3
            For TotalArray = 1 To ListView4.ListItems.Count
                If ListView4.ListItems(TotalArray).Tag <> "" Then
                    ListView4.ListItems(TotalArray).Checked = b
                Else
                    ListView4.ListItems(TotalArray).Checked = False
                End If
                If (TotalArray Mod 50) = 0 Then DoEvents
            Next TotalArray
       Case 4, 5
            'En el listview7
            b = (Index = 5)
            For TotalArray = 1 To ListView7.ListItems.Count
                ListView7.ListItems(TotalArray).Checked = b
                If (TotalArray Mod 50) = 0 Then DoEvents
            Next TotalArray
       Case 6, 7
            'En el listview8
            b = (Index = 6)
            For TotalArray = 1 To ListView8.ListItems.Count
                ListView8.ListItems(TotalArray).Checked = b
                If (TotalArray Mod 50) = 0 Then DoEvents
            Next TotalArray
       Case 8, 9
            b = (Index = 8)
            For TotalArray = 1 To Me.lw(0).ListItems.Count
                lw(0).ListItems(TotalArray).Checked = b
                If (TotalArray Mod 50) = 0 Then DoEvents
            Next TotalArray
       Case 10, 11
            b = (Index = 10)
            For TotalArray = 1 To Me.ListView10.ListItems.Count
                ListView10.ListItems(TotalArray).Checked = b
                If (TotalArray Mod 50) = 0 Then DoEvents
            Next TotalArray
       Case 12, 13
            b = (Index = 13)
            For TotalArray = 1 To Me.ListView11.ListItems.Count
                ListView11.ListItems(TotalArray).Checked = b
                
                If Index = 13 And ListView11.ListItems(TotalArray).Tag = 0 Then ListView11.ListItems(TotalArray).Checked = False
                
                If (TotalArray Mod 50) = 0 Then DoEvents
            Next TotalArray
       
    End Select
End Sub




Private Sub ListView11_ItemCheck(ByVal item As MSComctlLib.ListItem)
    For NumRegElim = 1 To ListView11.ListItems.Count
        If ListView11.ListItems(NumRegElim).Tag = 0 Then
            ListView11.ListItems(NumRegElim).Checked = False
        End If
    Next NumRegElim


End Sub

Private Sub OptCompXClien_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub OptCompXDpto_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub OptCompXMant_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub



Private Function RegresarCargaEmpresas() As String
Dim Sql As String
Dim Parametros As String
Dim I As Integer

    CadenaDesdeOtroForm = ""
    
        Sql = ""
        Parametros = ""
        For I = 1 To ListView2.ListItems.Count
            If Me.ListView2.ListItems(I).Checked Then
                Sql = Sql & Me.ListView2.ListItems(I).Text & "|"
                Parametros = Parametros & "1" 'Contador
            End If
        Next I
        CadenaDesdeOtroForm = Len(Parametros) & "|" & Sql
        'Vemos las conta
        Sql = ""
        For I = 1 To ListView2.ListItems.Count
            If Me.ListView2.ListItems(I).Checked Then
                Sql = Sql & Me.ListView2.ListItems(I).Tag & "|"
            End If
        Next I
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & Sql
    
    
        RegresarCargaEmpresas = CadenaDesdeOtroForm

End Function



Private Sub CargarArticulosEstanteria()
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim IT As ListItem

    Sql = "select sartic.*,nomfamia from sartic,sfamia where sartic.codfamia=sfamia.codfamia"
    If cadWHERE <> "" Then Sql = Sql & " AND " & cadWHERE
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    TotalArray = 0
    While Not Rs.EOF
        Set IT = ListView3.ListItems.Add
        IT.Tag = DevNombreSQL(Rs!codArtic)
        IT.Text = Rs!NomArtic
        IT.SubItems(1) = Format(Rs!preciove, cadWHERE2)
        IT.SubItems(2) = Rs!nomfamia
        IT.Checked = True
        Rs.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    Rs.Close
    
    
End Sub




Private Sub CargarArticulosCorreccionPrecio()
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim IT As ListItem
Dim margen As Currency
Dim MargenT As Currency
Dim ImpPVP As Currency
Dim ImpTar As Currency
Dim Aux As Currency
Dim decimales As Integer
Dim PrecioUC As Currency
Dim SoloImporteMenor As Boolean
    
    lblIndicadorCorregir = "LEYENDO BD"
    lblIndicadorCorregir.Refresh
    
    
    
    'Si NUMREGELIM=1 entonces esta marcada la opcion(check) de solo importes menores
    If NumRegElim = 1 Then SoloImporteMenor = True
    
    TotalArray = InStr(1, cadWHERE2, ",")
    Sql = Mid(cadWHERE2, TotalArray + 1)
    decimales = Len(Sql)
    'Formato
    cadWHERE2 = "#,##0." & Mid(cadWHERE2, TotalArray + 1)
    
    'Sql
    Sql = " SELECT sartic.nomartic,slista.codartic,sartic.preciove,sartic.preciouc,"
    Sql = Sql & "slista.precioac, slista.codlista, starif.nomlista,"
    Sql = Sql & "sartic.margecom as margenArt,starif.margecom as margetar"
    Sql = Sql & " FROM   (slista INNER JOIN sartic ON slista.codartic=sartic.codartic)"
    Sql = Sql & " INNER JOIN starif  ON slista.codlista=starif.codlista WHERE "

    Sql = Sql & cadWHERE '& " AND "
    ''SQL = SQL & " sartic.preciove <> sartic.preciouc + round((sartic.preciouc * if(isnull(sartic.margecom),0,sartic.margecom))/100," & Decimales & ")"
    
    Sql = Sql & " ORDER BY slista.codartic"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    

    '
  
    TotalArray = 0
    
    While Not Rs.EOF
        'Calculo los importes
        lblIndicadorCorregir.Caption = Rs!codArtic
        lblIndicadorCorregir.Refresh
        
        margen = DBLet(Rs!margenart, "N") / 100
        MargenT = DBLet(Rs!margetar, "N") / 100
        PrecioUC = DBLet(Rs!PrecioUC, "N")
        
        Aux = margen * PrecioUC
        ImpPVP = Round(PrecioUC + Aux, decimales)
        'El de la tarifa
        Aux = MargenT * ImpPVP
        ImpTar = Round(ImpPVP + Aux, decimales)
        
        Aux = Round(Rs!preciove, decimales)
        
        Sql = ""
        

        If SoloImporteMenor Then
            If Aux >= ImpPVP Then
                'El primero esta bien
                'Veamos el segundo. En la tarifa
                Aux = Round(Rs!precioac, decimales)
                If Aux < ImpTar Then Sql = "M"
            Else
                Sql = "M"
            End If
        
        
        Else
            If Aux = ImpPVP Then
                'El primero esta bien
                'Veamos el segundo. En la tarifa
                Aux = Round(Rs!precioac, decimales)
                If Aux <> ImpTar Then Sql = "M"
            Else
                Sql = "M"
            End If
        End If
        
        If Sql <> "" Then
            Set IT = ListView4.ListItems.Add
            IT.Tag = DevNombreSQL(Rs!codArtic)
            IT.ToolTipText = IT.Tag
            IT.Text = IT.Tag
            IT.SubItems(1) = Rs!NomArtic
            Aux = Round(PrecioUC, decimales)
            IT.SubItems(2) = Format(Aux, cadWHERE2)
            
            IT.SubItems(3) = Format(margen * 100, FormatoPorcen)
            Aux = Round(Rs!preciove, decimales)
            IT.SubItems(4) = Format(Aux, cadWHERE2)
            
            IT.SubItems(5) = Format(MargenT * 100, FormatoPorcen)
            Aux = Round(Rs!precioac, decimales)
            IT.SubItems(6) = Format(Aux, cadWHERE2)
            
            'Precio venta correcto
            
            IT.SubItems(7) = Format(ImpPVP, cadWHERE2)
            IT.SubItems(8) = Format(ImpTar, cadWHERE2)
            
            
            
            If PrecioUC = 0 Then
                'Precio ultima compra =0
                'NOOOOO se puede actualizar la tarifa
                IT.Tag = "" 'para no actualizar
                IT.Checked = False
                IT.Bold = True
                IT.ForeColor = vbRed
            Else
                
            End If
            IT.Checked = False
        End If
        Rs.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            Me.Refresh
            DoEvents
        End If
    Wend
    Rs.Close
    cmbActualizarTar.ListIndex = 0
    lblIndicadorCorregir.Caption = ""
End Sub




Private Function TraspasarMantenimientos() As Boolean
    
    On Error GoTo ETraspasarMantenimientos
    TraspasarMantenimientos = False

    

    cadWHERE = "Select count(*) from sliman where anomante =" & txtMante(0).Text
    miRsAux.Open cadWHERE, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    
    If NumRegElim > 0 Then
        MsgBox "Ya existen datos para el año " & txtMante(0).Text, vbExclamation
        Exit Function
    End If
    
    
    
    'Se divide en 3 pasos
    '1.- Introducir una linea en la sliman con los datos para el año
        cadWHERE = "insert into sliman (anomante,codclien,nummante,mes01man,mes02man,mes03man,mes04man,mes05man,mes06man,mes07man,mes08man,mes09man,mes10man,mes11man,mes12man)"
        cadWHERE = cadWHERE & " SELECT " & txtMante(0).Text & ",codclien,nummante,mes01act,mes02act,mes03act,mes04act,mes05act,mes06act,mes07act,mes08act,mes09act,mes10act,mes11act,mes12act FROM scaman"
        conn.Execute cadWHERE
    '2.- Updatear los campos de actual con siguiente
        cadWHERE = ""
        For TotalArray = 1 To 12
            cadWHERE = cadWHERE & ", mes" & Format(TotalArray, "00") & "act = mes" & Format(TotalArray, "00") & "sig"
        Next TotalArray
        cadWHERE = Mid(cadWHERE, 2) 'Para quitar la primera coma
        cadWHERE = "UPDATE scaman SET " & cadWHERE
        conn.Execute cadWHERE
        
    '3.- Si no han marcado la opcion copiar datos tengo que resetear a 0
        If chkMante.Value = 0 Then
            'NO SE COPIA, luego hay que resetear
            cadWHERE = ""
            For TotalArray = 1 To 12
                cadWHERE = cadWHERE & ", mes" & Format(TotalArray, "00") & "sig = 0 "
            Next TotalArray
            cadWHERE = Mid(cadWHERE, 2) 'Para quitar la primera coma
            cadWHERE = "UPDATE scaman SET " & cadWHERE
            conn.Execute cadWHERE
        End If
    TraspasarMantenimientos = True
    
    Exit Function
ETraspasarMantenimientos:
    MuestraError Err.Number
End Function

Private Sub txtMante_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Public Function ObtenerSQLcomponentes(cadWHERE As String) As String
'Obtiene la consulta SQL que selecciona los articulos con nº de serie
'agrupados por tipo de articulo
Dim Sql As String

    Sql = "Select distinct sserie.codtipar, nomtipar, count(numserie) as cantidad "
    Sql = Sql & "FROM sserie INNER JOIN stipar ON sserie.codtipar=stipar.codtipar "
    Sql = Sql & cadWHERE
    Sql = Sql & " GROUP by codtipar "
    
    ObtenerSQLcomponentes = Sql
End Function



Private Sub CargarListaPedidosSinAlbaran()
'Muestra la lista Detallada de pedidos sin numero de albaran asignado
'en un ListView
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String

    On Error GoTo ECargarList

    Sql = cadWHERE 'cadwhere ya le pasamos toda la SQL
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        'Los encabezados
        ListView6.ColumnHeaders.Clear

        ListView6.ColumnHeaders.Add , , "Nº Pedido", 1000
        ListView6.ColumnHeaders.Add , , "Fecha", 1100, 1
        ListView6.ColumnHeaders.item(2).Alignment = lvwColumnCenter
        ListView6.ColumnHeaders.Add , , "Código", 900
        ListView6.ColumnHeaders.Add , , "Cliente", 2100
        ListView6.ColumnHeaders.Add , , "Código  ", 700
        ListView6.ColumnHeaders.Add , , "Destino", 2100
        
        
    
        While Not Rs.EOF
            Set ItmX = ListView6.ListItems.Add
            ItmX.Text = Format(Rs!numpedid, "000000")
            ItmX.SubItems(1) = Rs!FechaPed
            ItmX.SubItems(2) = Format(Rs!CodClien, "000000")
            ItmX.SubItems(3) = Rs!Nomclien
            ItmX.SubItems(4) = Format(Rs!coddesti, "000")
            ItmX.SubItems(5) = Rs!nomdesti
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing

ECargarList:
    If Err.Number <> 0 Then
        Err.Clear
    End If



End Sub

Private Sub CargarListaFields(DadoProducto As Boolean)
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim IT As ListItem

    Select Case Label5.Caption
        Case "Clases"
            Sql = "select clases.codclase as codigo, clases.nomclase as descripcion from clases "
        Case "Variedades"
            Sql = "select variedades.codvarie as codigo, variedades.nomvarie as descripcion from variedades "
        Case "Clientes"
            Sql = "select clientes.codclien as codigo, clientes.nomclien as descripcion from clientes "
        Case "Destinos"
            Sql = "select destinos.coddesti as codigo, destinos.nomdesti as descripcion from destinos "
        Case "Forfaits"
            Sql = "select forfaits.codforfait as codigo, forfaits.nomconfe as descripcion from forfaits "
        Case "Marcas"
            Sql = "select marcas.codmarca as codigo, marcas.nommarca as descripcion from marcas "
        Case "Mercados"
            Sql = "select tipomer.codtimer as codigo, tipomer.nomtimer as descripcion from tipomer "
        Case "Paises"
            Sql = "select paises.codpaise as codigo, paises.nompaise as descripcion from paises "
        Case "Comisionistas"
            Sql = "select agencias.codtrans as codigo, agencias.nomtrans as descripcion from agencias "
        '[Monica]17/06/2013: añadimos las categorias
        Case "Categorias"
            Sql = "select distinct categori as codigo, '' as descripcion from albaran_variedad "
        '[Monica]20/09/2013: productos
        Case "Productos"
            Sql = "select productos.codprodu as codigo, productos.nomprodu as descripcion from productos "
        '[Monica]17/10/2016: contratos
        Case "Contratos"
            Sql = "select distinct nrocontra as codigo, '' as descripcion from albaran "
    End Select

'    ' viene de un rango de clases
'    Sql = "select variedades.codvarie, variedades.nomvarie, variedades.codclase, clases.nomclase from variedades, clases "
'    Sql = Sql & " where variedades.codclase = clases.codclase "
'
    If cadWHERE <> "" Then Sql = Sql & " where (1=1) " & cadWHERE
    
    If Label5 = "Comisionistas" Then Sql = Sql & " and agencias.tipo = 1"
    
    '[Monica]17/10/2016: contrato
    If Label5 = "Contratos" Then Sql = Sql & " group by 1 order by 1"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView7.ColumnHeaders.Clear
    
'    ListView7.ColumnHeaders.Add , , "Código", 1000.0631
'    ListView7.ColumnHeaders.Add , , "Variedad", 2200.2522, 1
'    ListView7.ColumnHeaders.Add , , "Clase", 799.9371, 1
'    ListView7.ColumnHeaders.Add , , "Descripción", 2101.0396
    
    ListView7.ColumnHeaders.Add , , "Código", 2000.0631
    ListView7.ColumnHeaders.Add , , "Descripción", 4101.0396
    TotalArray = 0
    While Not Rs.EOF
        Set IT = ListView7.ListItems.Add
            
'        It.Text = Format(DBLet(RS!codvarie, "N"), "000000")
'        It.SubItems(1) = DBLet(RS!nomvarie, "T")
'        It.SubItems(2) = Format(DBLet(RS!codclase, "N"), "000")
'        It.SubItems(3) = DBLet(RS!nomclase, "T")
        Select Case Label5.Caption
            Case "Clases"
                IT.Text = Format(DBLet(Rs!codigo, "N"), "000")
            Case "Variedades"
                IT.Text = Format(DBLet(Rs!codigo, "N"), "000000")
            Case "Clientes"
                IT.Text = Format(DBLet(Rs!codigo, "N"), "000000")
            Case "Destinos"
                IT.Text = Format(DBLet(Rs!codigo, "N"), "000000")
            Case "Forfaits"
                IT.Text = DBLet(Rs!codigo, "T")
            Case "Marcas"
                IT.Text = Format(DBLet(Rs!codigo, "N"), "000")
            Case "Mercados"
                IT.Text = Format(DBLet(Rs!codigo, "N"), "000")
            Case "Paises"
                IT.Text = Format(DBLet(Rs!codigo, "N"), "000")
            Case "Comisionistas"
                IT.Text = Format(DBLet(Rs!codigo, "N"), "000")
            Case "Categorias"
                IT.Text = DBLet(Rs!codigo, "T")
            Case "Productos"
                IT.Text = DBLet(Rs!codigo, "T")
            Case "Contratos"
                IT.Text = DBLet(Rs!codigo, "T")
        End Select
        IT.SubItems(1) = DBLet(Rs!Descripcion, "T")
         
        If Label5.Caption = "Categorias" Or Label5.Caption = "Contratos" Then
            IT.Checked = True
        Else
            IT.Checked = False
        End If
        
        Rs.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    Rs.Close
    
End Sub



Private Sub CargarListaFacturas()
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim IT As ListItem

     Sql = "select codtipom, numfactu, fecfactu, totalfac from facturas"

    If cadWHERE <> "" Then Sql = Sql & " " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView8.ColumnHeaders.Clear
    
    ListView8.ColumnHeaders.Add , , "Tipo", 1000.0631
    ListView8.ColumnHeaders.Add , , "Nro.Factura", 1200.2522, 1
    ListView8.ColumnHeaders.Add , , "Fecha Factura", 1799.9371, 1
    ListView8.ColumnHeaders.Add , , "Total Factura", 2101.0396
    
    TotalArray = 0
    While Not Rs.EOF
        Set IT = ListView8.ListItems.Add
            
        IT.Text = DBLet(Rs!codTipoM, "T")
        IT.SubItems(1) = Format(DBLet(Rs!NumFactu, "N"), "0000000")
        IT.SubItems(2) = Format(DBLet(Rs!FecFactu, "F"), "dd/mm/yyyy")
        IT.SubItems(3) = Format(DBLet(Rs!TotalFac, "N"), "###,###,##0.00")
        
        IT.Checked = False
        
        Rs.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    Rs.Close
    
End Sub

Private Sub CargarListaCalibres()
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim IT As ListItem
Dim I As Integer

    Sql = "select numline1, albaran_calibre.codcalib, nomcalib, numcajas, unidades, pesobrut, pesoneto, albaran_calibre.codvarie from albaran_calibre Inner join calibres on albaran_calibre.codvarie = calibres.codvarie and albaran_calibre.codcalib = calibres.codcalib "

    If cadWHERE <> "" Then Sql = Sql & " where " & cadWHERE
    
    Sql = Sql & " order by 1 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView9.ColumnHeaders.Clear
    
    ListView9.ColumnHeaders.Add , , "Linea", 600
    ListView9.ColumnHeaders.Add , , "Código", 800.2522, 1
    ListView9.ColumnHeaders.Add , , "Calibre", 1000
    ListView9.ColumnHeaders.Add , , "Cajas", 1000.0396, 1
    ListView9.ColumnHeaders.Add , , "Unidades", 1000.0396, 1
    ListView9.ColumnHeaders.Add , , "Peso Bruto", 1000.0396, 1
    ListView9.ColumnHeaders.Add , , "Peso Neto", 1000.0396, 1
    
    TotalArray = 0
    I = -1
    While Not Rs.EOF
        Set IT = ListView9.ListItems.Add
            
        IT.Text = DBLet(Rs!numline1, "T")
        IT.SubItems(1) = Format(DBLet(Rs!codcalib, "N"), "00")
        IT.SubItems(2) = DBLet(Rs!nomcalib, "T")
        IT.SubItems(3) = Format(DBLet(Rs!NumCajas, "N"), "###,##0")
        IT.SubItems(4) = Format(DBLet(Rs!Unidades, "N"), "###,##0")
        IT.SubItems(5) = Format(DBLet(Rs!pesobrut, "N"), "###,##0")
        IT.SubItems(6) = Format(DBLet(Rs!Pesoneto, "N"), "###,##0")
        
        If I = -1 Then
            I = IT.Index
            IT.Checked = True
        End If
        
        Rs.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    Rs.Close
    Set Rs = Nothing
    
    If I > 0 Then Set ListView9.SelectedItem = ListView9.ListItems(I)
    
    
End Sub


Private Sub cmdActualizaPMP_Click()

    Sql = ""
    For NumRegElim = 1 To lw(0).ListItems.Count
        If Me.lw(0).ListItems(NumRegElim).Checked Then Sql = Sql & "X"
    Next NumRegElim
    
    
    If Sql = "" Then
        MsgBox "Seleccione algún articulo para actualizar", vbExclamation
        Exit Sub
    End If
    
    
    Sql = "Va a actualizar " & Len(Sql) & " referencia(s)"
    Sql = Sql & vbCrLf & vbCrLf & "¿Continuar?"
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    
    Screen.MousePointer = vbHourglass
    ActualizarReferencias
    Screen.MousePointer = vbDefault
    If Sql = "" Then
        CadenaDesdeOtroForm = "OK"
        Unload Me  'ha ido bien
    End If
    
End Sub

Private Sub CargaLwPrecioMP()
Dim Sql As String

    lw(0).ColumnHeaders.Clear

    lw(0).ColumnHeaders.Add , , "Código", 1100.2522, 0
    lw(0).ColumnHeaders.Add , , "Nombre", 4000
    lw(0).ColumnHeaders.Add , , "Proveedor", 1000.0396, 1
    lw(0).ColumnHeaders.Add , , "Familia", 1000.0396, 1
    lw(0).ColumnHeaders.Add , , "PMP Actual", 1200.0396, 1
    lw(0).ColumnHeaders.Add , , "PMP Calculado", 1300.0396, 1
    



    Set miRsAux = New ADODB.Recordset
    Me.lw(0).ListItems.Clear

    Sql = "Select * from tmpinformes where codusu = " & vUsu.codigo & " ORDER BY campo1,nombre1"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lw(0).ListItems.Add()
        IT.Text = miRsAux!nombre1  'codartic
        IT.SubItems(1) = miRsAux!nombre2 'nomartic
        IT.SubItems(2) = miRsAux!campo1
        IT.SubItems(3) = miRsAux!campo2
        IT.SubItems(4) = Format(miRsAux!precio1, "###,##0.0000")
        IT.SubItems(5) = Format(miRsAux!precio2, "###,##0.0000")


        IT.Checked = False

        If miRsAux!importeb1 <> 0 And miRsAux!importeb2 <> 0 Then IT.Checked = True



        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub


Private Sub ActualizarReferencias()
Dim HayError As Boolean


    vCadena = ""
    HayError = False
    For NumRegElim = lw(0).ListItems.Count To 1 Step -1
        If lw(0).ListItems(NumRegElim).Checked Then

            If OpcionMensaje = 24 Then
                Sql = "preciomp"
            Else
                Sql = "preciouc"
            End If

            Sql = "UPDATE sartic set " & Sql & " = " & DBSet(lw(0).ListItems(NumRegElim).SubItems(5), "N")
            Sql = Sql & " WHERE codartic = " & DBSet(lw(0).ListItems(NumRegElim).Text, "T")
            If Not ejecutar(Sql) Then
                HayError = True
                NumRegElim = Me.lw(0).ListItems.Count + 1
            Else
                vCadena = vCadena & "  ·  " & DBSet(lw(0).ListItems(NumRegElim).Text, "T")
                lw(0).ListItems.Remove lw(0).ListItems(NumRegElim).Index

                If Len(vCadena) > 230 Then InsertaLog  'y pone vcdena a ""

            End If
        End If
    Next NumRegElim

    If vCadena <> "" Then InsertaLog 'y pone vcdena a ""

    'Si llega aqui... tutto benne
    If Not HayError Then Sql = ""

End Sub

Private Function ejecutar(Sql As String) As Boolean
    On Error GoTo eEjecutar
    
    ejecutar = True
    conn.Execute Sql
    Exit Function
    
eEjecutar:
    ejecutar = False
End Function


Private Sub InsertaLog()

        '------------------------------------------------------------------------------
        '  LOG de acciones
        Set LOG = New cLOG
        vCadena = Mid(vCadena, 6) 'quitamos el primer separador
        If OpcionMensaje = 24 Then
            vCadena = "PMP: " & vCadena
        Else
            vCadena = "UPC: " & vCadena
        End If
        vCadena = Replace(vCadena, "'", "")
        LOG.Insertar 19, vUsu, vCadena
        vCadena = ""
        Set LOG = Nothing
        '-----------------------------------------------------------------------------
        espera 0.6
End Sub


Private Sub Text7_LostFocus()
    RaiseEvent DatoSeleccionado(Text7.Text)
    Unload Me
End Sub


Private Sub Text7_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        Text7_LostFocus
    ElseIf KeyAscii = 27 Then 'ESC
            Text7_LostFocus
    End If
End Sub
    


Private Sub CargarListaProductos()
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim IT As ListItem

     Sql = "select codprodu, nomprodu from productos "

    If cadWHERE <> "" Then Sql = Sql & " " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView10.ColumnHeaders.Clear
    
    ListView10.ColumnHeaders.Add , , "Código", 1000.0631
    ListView10.ColumnHeaders.Add , , "Nombre", 3200.2522
    
    TotalArray = 0
    While Not Rs.EOF
        Set IT = ListView10.ListItems.Add
            
        IT.Checked = True
        IT.Text = Format(DBLet(Rs!codprodu, "N"), "000")
        IT.SubItems(1) = DBLet(Rs!nomprodu, "T")
        
        
        Rs.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    Rs.Close
    
End Sub


Private Sub CargarFacturasAnecoop()
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim IT As ListItem

    Sql = "select fra_liq, fecha_liq, sum(importe_liq) importe_liq, sum(importe_iva_liq) importe_iva_liq,  sum(importe_iva_liq + importe_liq) total from anecoop "

    If cadWHERE <> "" Then Sql = Sql & " where " & cadWHERE
    
    ' miramos que tengan todas sus lineas con nro de linea de albaran
'    Sql = Sql & " and not (fra_liq) in (select distinct fra_liq from anecoop where " & cadWHERE & " and (numlinea is null and nombre_variedad <> '') )"
    
    Sql = Sql & " group by 1,2 "
    Sql = Sql & " order by 1,2 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView11.ColumnHeaders.Clear
    
    ListView11.ColumnHeaders.Add , , "Nro.Factura", 1100.0631
    ListView11.ColumnHeaders.Add , , "Fecha ", 1150.2522
    ListView11.ColumnHeaders.Add , , "Base Imp. ", 1400.2522, 1
    ListView11.ColumnHeaders.Add , , "Iva ", 1300.25221, 1
    ListView11.ColumnHeaders.Add , , "TOTAL ", 1400.2522, 1
    
    TotalArray = 0
    While Not Rs.EOF
        Set IT = ListView11.ListItems.Add
            
        IT.Text = DBLet(Rs!fra_liq, "T")
        IT.SubItems(1) = DBLet(Rs!fecha_liq, "F")
        IT.SubItems(2) = Format(DBLet(Rs!importe_liq, "N"), "###,###,##0.00")
        IT.SubItems(3) = Format(DBLet(Rs!importe_iva_liq, "N"), "###,###,##0.00")
        IT.SubItems(4) = Format(DBLet(Rs!Total, "N"), "###,###,##0.00")
                
        Sql2 = "select count(*) from anecoop where fra_liq = " & DBSet(Rs!fra_liq, "N") & " and fra_liq in  (select distinct fra_liq from anecoop where " & cadWHERE & " and (numlinea is null and nombre_variedad <> '') )"
        If TotalRegistros(Sql2) <> 0 Then
            IT.ForeColor = &HC0&
            IT.Bold = True
            IT.ListSubItems(1).ForeColor = &HC0&
'            IT.ListSubItems(1).Bold = True
            IT.ListSubItems(2).ForeColor = &HC0&
'            IT.ListSubItems(2).Bold = True
            IT.ListSubItems(3).ForeColor = &HC0&
'            IT.ListSubItems(3).Bold = True
            IT.ListSubItems(4).ForeColor = &HC0&
'            IT.ListSubItems(4).Bold = True
            
            IT.Tag = 0
            
            IT.Checked = False
        Else
            IT.Tag = 1
            
            IT.Checked = True
        End If
        
        Rs.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    Rs.Close
    
End Sub
