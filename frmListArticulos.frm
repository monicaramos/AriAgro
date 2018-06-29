VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListArticulos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   9435
   Icon            =   "frmListArticulos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   8730
      Top             =   5580
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameInfArticulos 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9435
      Begin VB.Frame FrameOrden 
         BorderStyle     =   0  'None
         Height          =   1935
         Left            =   6435
         TabIndex        =   34
         Top             =   585
         Width           =   2655
         Begin VB.CommandButton cmdBajar 
            Height          =   510
            Left            =   2055
            Picture         =   "frmListArticulos.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   1080
            Width           =   510
         End
         Begin VB.CommandButton cmdSubir 
            Caption         =   " "
            Height          =   510
            Left            =   2055
            Picture         =   "frmListArticulos.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   375
            Width           =   510
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   1335
            Left            =   255
            TabIndex        =   37
            Top             =   255
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   2355
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Orden del Informe"
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
            Height          =   240
            Index           =   31
            Left            =   255
            TabIndex        =   38
            Top             =   15
            Width           =   1770
         End
      End
      Begin VB.Frame FrameStockMaxMin 
         Caption         =   "Imprimir Stocks"
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
         Height          =   750
         Left            =   510
         TabIndex        =   31
         Top             =   5760
         Width           =   3135
         Begin VB.OptionButton optStockMin 
            Caption         =   "Mínimos"
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
            Left            =   1620
            TabIndex        =   33
            Top             =   280
            Width           =   1425
         End
         Begin VB.OptionButton optStockMax 
            Caption         =   "Máximos"
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
            Left            =   240
            TabIndex        =   32
            Top             =   280
            Width           =   1425
         End
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   69
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "Text5"
         Top             =   4245
         Width           =   5865
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   68
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "Text5"
         Top             =   3840
         Width           =   5865
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   69
         Left            =   1755
         MaxLength       =   2
         TabIndex        =   8
         Top             =   4245
         Width           =   795
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   68
         Left            =   1755
         MaxLength       =   2
         TabIndex        =   7
         Top             =   3840
         Width           =   795
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   63
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "Text5"
         Top             =   1845
         Width           =   3780
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   62
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "Text5"
         Top             =   1440
         Width           =   3780
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   71
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   5355
         Width           =   4980
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   70
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text5"
         Top             =   4950
         Width           =   4980
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   63
         Left            =   1755
         MaxLength       =   4
         TabIndex        =   4
         Top             =   1845
         Width           =   795
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   62
         Left            =   1755
         MaxLength       =   4
         TabIndex        =   3
         Top             =   1440
         Width           =   795
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   71
         Left            =   1755
         MaxLength       =   16
         TabIndex        =   10
         Top             =   5355
         Width           =   1635
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   70
         Left            =   1755
         MaxLength       =   16
         TabIndex        =   9
         Top             =   4950
         Width           =   1635
      End
      Begin VB.CommandButton cmdAceptarArtic 
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
         Left            =   6195
         TabIndex        =   11
         Top             =   5970
         Width           =   1065
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   11
         Left            =   7320
         TabIndex        =   12
         Top             =   5970
         Width           =   1065
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   66
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   5
         Top             =   2610
         Width           =   915
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   67
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   6
         Top             =   3015
         Width           =   915
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   66
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "Text5"
         Top             =   2610
         Width           =   5700
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   67
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "Text5"
         Top             =   3015
         Width           =   5700
      End
      Begin VB.Label Label1 
         Caption         =   "Informe de Artículos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   495
         TabIndex        =   39
         Top             =   495
         Width           =   6735
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   26
         Left            =   1470
         Top             =   4260
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   25
         Left            =   1470
         Top             =   3840
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Artículo"
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
         Height          =   240
         Index           =   30
         Left            =   555
         TabIndex        =   30
         Top             =   3495
         Width           =   1560
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
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
         Index           =   60
         Left            =   825
         TabIndex        =   29
         Top             =   4245
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
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
         Index           =   59
         Left            =   825
         TabIndex        =   28
         Top             =   3840
         Width           =   735
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   20
         Left            =   1470
         Top             =   1845
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   19
         Left            =   1470
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Familia"
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
         Height          =   240
         Index           =   39
         Left            =   555
         TabIndex        =   25
         Top             =   1110
         Width           =   660
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
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
         Index           =   56
         Left            =   825
         TabIndex        =   24
         Top             =   1845
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
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
         Index           =   55
         Left            =   825
         TabIndex        =   23
         Top             =   1440
         Width           =   735
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   28
         Left            =   1470
         Top             =   5355
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   27
         Left            =   1470
         Top             =   4995
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Artículo"
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
         Height          =   240
         Index           =   38
         Left            =   555
         TabIndex        =   22
         Top             =   4635
         Width           =   750
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
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
         Index           =   54
         Left            =   825
         TabIndex        =   21
         Top             =   5355
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
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
         Index           =   51
         Left            =   825
         TabIndex        =   20
         Top             =   4995
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
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
         Index           =   50
         Left            =   825
         TabIndex        =   19
         Top             =   2610
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
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
         Index           =   48
         Left            =   825
         TabIndex        =   18
         Top             =   3015
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00972E0B&
         Height          =   240
         Index           =   37
         Left            =   555
         TabIndex        =   17
         Top             =   2280
         Width           =   990
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   23
         Left            =   1470
         Top             =   2610
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   24
         Left            =   1470
         Top             =   3045
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmListArticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcionlistado As Integer

    '==== Listados de ALMACEN ====
    '=============================
    ' 1 .- Listados Marcas.
    ' 2 .- Listado de Almacenes Propios
    ' 3 .- Listado de Tipos de Unidad
    ' 4 .- Listado de Tipos de Artículos
    ' 5 .- Listado de Familias de artículos
    
    ' 6 .- Listado de Artículos
    ' 7 .- Informe de Traspaso de Almacenes
    ' 8 .- Informe de Movimientos de Almacen
    ' 9 .- Listado Busquedas de movimientos de Artículos
    '10 .-
    
    '11 .-
    '12 .- Listado Toma de Inventario Articulos
    '13 .- Listado de Diferencias de Inventario Articulos
    '14 .- Actualizar Diferencias de Inventario (No IMPRIME INFORME)
    '15 .- Listado de Articulos Inactivos.
    
    '16 .- Listado Valoracion de Stocks Inventariados
    '17 .- Listado Valoración Stocks
    '18 .- Informe Stocks Maximos y Minimos
    '19 .- Informe de Stocks a una fecha
    
    '110 .- Listado de Ubicaciones
    
    
    
    
    '==== Listados de FACTURACION ====
    '=================================
    '20 .- Listado de Actividades de Clientes
    '21 .- Listado de Zonas de Clientes
    '22 .- Listado de Rutas de Asistencia
    '23 .- Listado de Formas de Envío
    '24 .- Listado de Tarifas Ventas
    '25 .-
    
    '26 .-
    '27 .- Listado de Situaciones Especiales
    '28 .- Informe de Tarifas de Articulos
    '29 .- Informe de Promociones de Tarifas
    '30 .- Informe de Precios Especiales
    
    '31 .- Informe de Ofertas
    '32 .- Informe de Recordatorio de Ofertas
    '33 .- Informe de Valoración de Ofertas
    '34 .- Informe de Ofertas Efectuadas
    '35 .- Informe Historico de Ofertas
    
    '36 .- Traspaso de Ofertas al Historico (NO IMPRIME INFORME)
    '37 .- Solicitar datos para pasar de Oferta a Pedido (NO IMPRIME INFORME)
    '38 .- Informe de Pedidos
    '239 .- Hco de Pedidos de venta (Historico)
    '39 .- Orden de Instalacion
    '40 .- Cartas Confirmacion de Pedidos
    
    '41 .- Informe de Pedidos por Articulo
    '42 .- Informe de Disponibilidad de Stocks
    '43 .- Generar Albaran desde Pedido (NO IMPRIME LISTADO)
    '44 .- Informe de Pedidos por Cliente
    '45 .- Informe de Albaran
    
    '46 .- Informe de Clientes Inactivos
    '47 .- Informe de Clientes
    '48 .- Informe de Altas de Nuevos Cliente
    '49 .- Informe de Albaranes por Articulo
    '50 .- Prevision de Facturacion de ALbaranes
    
    '51 .- Informe Incumplimiento Plazos de Entrega
    '52 .- Facturacion de Albaranes (NO IMPRIME LISTADO?)
    '53 .- Informe de Factura
    '54 .- Listado de Descuentos Familia/Marca
    
    '59 .- Informe de Factura ProForma
    '222 .- Informe de Factura Mostrador
    '223 .- Pedir datos para contabilizar facturas CLIENTES
    '224 .- Pedir datos para contabilizar facturas PROVEEDOR
    '225 .- Pedir datos para generar Facturas Rectificativas
    '226 .- Pedir datos para reimprimir Facturas
    '227 .- Informe estadistica Ventas por cliente
    '228 .- Informe estadistica Ventas por Trabajador
    '229 .- Informe estadistica Ventas por meses
    
    '240 .- Informe Cierre de Caja del TPV
    
    
    
    '==== Listados de COMPRAS ====
    '=============================
    '55 .- Informe de Pedido Proveedor
    '56 .- Inf. Historico Pedido Proveedor
    '57 .- Pasa Pedido a Albaran compras (NO IMPRIME LISTADO)
    '58 .- Listado de Proveedores
    
    
    '305 .- Listado Etiquetas de Proveedores
    '306 .- Listado Cartas a Proveedores
    '307 .- Listado Material pendiente de recibir
    '308 .- Listado Albaranes pendientes de facturar
    '309 .- Listado  Precios de Compra
    
    
    '==== Listados de REPARACIONES ====
    '==================================
    '60 .- Informe de Numeros de Serie
    '61 .- Listado Motivos Pend. Rep.
    '62 .- Listado Resguardo Reparacion
    '63 .- Listado Reparaciones por Día
    '64 .- Listado Reparaciones por Cliente
    
    '406 .- Listado Frecuencia de reparaciones
    '407 .- Sustitución Nº de Serie
    '408 .- Informe Aviso de Averia
    '409 .- Listado Avisos de averia pendientes
    
    
    '==== Listados de MANTENIMIENTOS ====
    '==================================
    '70 .- Listado Mantenimiento
    '71 .- Listado Revisiones de Mantenimientos
    '72 .- Informe Fichas de Mantenimientos
    '73 .- Listado Altas de Mantenimientos
    '74 .- Prefacturación Mantenimientos
    '75 .- Facturación de Mantenimientos
    
    
    
    '==== Listados OTROS ====
    '==================================
    
    '80 .- Pasar Albaranes Ventas al historico (NO IMPRIME)
    '81 .- Pasar Pedidos Ventas al historico (NO IMPRIME)
       
    '90 .- Etiquetas de Clientes
    '91 .- Cartas a Clientes
    
    '92 .- Informe de Gastos Técnicos
    '93 .- Ticket del TPV
    
    
    
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmArt As frmManArtic
Attribute frmArt.VB_VarHelpID = -1
Private WithEvents frmFam As frmManFamilias
Attribute frmFam.VB_VarHelpID = -1
Private WithEvents frmPro As frmManProve
Attribute frmPro.VB_VarHelpID = -1
Private WithEvents frmMarcas As frmManMarcas
Attribute frmMarcas.VB_VarHelpID = -1
Private WithEvents frmTArt As frmManTipArtic
Attribute frmTArt.VB_VarHelpID = -1

Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1

'---- Variables para el INFORME ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadselect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para el frmImprimir
Private conSubRPT As Boolean 'Si el informe tiene subreports
Private cadNombreRPT As String 'Nombre del informe
'-----------------------------------

Dim TipCod As String
Dim indCodigo As Integer 'indice para txtCodigo

Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean
Dim indFrame As Single


Private Sub KEYpress(KeyAscii As Integer)
    Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cmdAceptarArtic_Click()
'Listado de Articulos
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim campo As String
Dim Opcion As Byte, numOp As Byte

    InicializarVbles
    
    cadNombreRPT = "rListArticulos.rpt"  'Nombre fichero .rpt a Imprimir
    cadTABLA = "sartic"
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion D/H FAMILIA
    '--------------------------------------------
    cDesde = Trim(txtCodigo(62).Text)
    cHasta = Trim(txtCodigo(63).Text)
    nDesde = txtNombre(62).Text
    nHasta = txtNombre(63).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & cadTABLA & ".codfamia}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFamilia= """) Then Exit Sub
    End If

'    'Cadena para seleccion D/H MARCA
'    '--------------------------------------------
'    cDesde = Trim(txtCodigo(64).Text)
'    cHasta = Trim(txtCodigo(65).Text)
'    nDesde = txtNombre(64).Text
'    nHasta = txtNombre(65).Text
'    If Not (cDesde = "" And cHasta = "") Then
'        'Cadena para seleccion Desde y Hasta
'        Codigo = "{" & cadTABLA & ".codmarca}"
'        TipCod = "N"
'        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHMarca= """) Then Exit Sub
'    End If
    
    'Cadena para seleccion D/H PROVEEDOR
    '--------------------------------------------
    cDesde = Trim(txtCodigo(66).Text)
    cHasta = Trim(txtCodigo(67).Text)
    nDesde = txtNombre(66).Text
    nHasta = txtNombre(67).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & cadTABLA & ".codprove}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHProveedor= """) Then Exit Sub
    End If
    
    'Cadena para seleccion D/H TIPO ARTICULO
    '--------------------------------------------
    cDesde = Trim(txtCodigo(68).Text)
    cHasta = Trim(txtCodigo(69).Text)
    nDesde = txtNombre(68).Text
    nHasta = txtNombre(69).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & cadTABLA & ".codtipar}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHTipoArt= """) Then Exit Sub
    End If
    
    'Cadena para seleccion D/H ARTICULO
    '--------------------------------------------
    cDesde = Trim(txtCodigo(70).Text)
    cHasta = Trim(txtCodigo(71).Text)
    nDesde = txtNombre(70).Text
    nHasta = txtNombre(71).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & cadTABLA & ".codartic}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHArticulo= """) Then Exit Sub
    End If

    
    'Obtener el parametro con el Orden del Informe
    '---------------------------------------------
        numOp = PonerGrupo(1, ListView2.ListItems(1).Text)
        If numOp <> 0 Then Opcion = numOp
        numOp = PonerGrupo(2, ListView2.ListItems(2).Text)
        If numOp <> 0 Then Opcion = numOp
        numOp = PonerGrupo(3, ListView2.ListItems(3).Text)
        If numOp <> 0 Then Opcion = numOp
'        numOp = PonerGrupo(4, ListView2.ListItems(4).Text)
'        If numOp <> 0 Then Opcion = numOp
        Opcion = Opcion - 1
    
        Select Case Opcion
            Case 1 'El group2 es el Proveedor
                campo = "pTitulo1=""" & ListView2.ListItems(3).Text & """" '3
                cadParam = cadParam & campo & "|"
                numParam = numParam + 1
                
                campo = "pTitulo2=""" & ListView2.ListItems(3).Text & """" '4
                cadParam = cadParam & campo & "|"
                numParam = numParam + 1
            Case 2 'El Group3 es el Proveedor
                campo = "pTitulo1=""" & ListView2.ListItems(2).Text & """" '2
                cadParam = cadParam & campo & "|"
                numParam = numParam + 1
                
                campo = "pTitulo2=""" & ListView2.ListItems(3).Text & """" '4
                cadParam = cadParam & campo & "|"
                numParam = numParam + 1
            Case 3, 0 'El Group4 es el Proveedor
                      '0 'El Group1 es el Proveedor
                campo = "pTitulo1=""" & ListView2.ListItems(2).Text & """" '2
                cadParam = cadParam & campo & "|"
                numParam = numParam + 1
                
                campo = "pTitulo2=""" & ListView2.ListItems(3).Text & """" '3
                cadParam = cadParam & campo & "|"
                numParam = numParam + 1
                
                If Opcion = 0 Then
                    campo = "pTitulo3=""" & ListView2.ListItems(3).Text & """" '4
                    cadParam = cadParam & campo & "|"
                    numParam = numParam + 1
                End If
        End Select
   
    'Parametro Orden del Informe
    campo = "pOrden=" & Opcion
    cadParam = cadParam & campo & "|"
    numParam = numParam + 1
    
    If HayRegParaInforme(cadTABLA, cadselect) Then
       LlamarImprimir
    End If
    
End Sub

Private Sub cmdBajar_Click()
'Bajar el item seleccionado del listview2
    BajarItemList Me.ListView2
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdSubir_Click()
'Subir el item seleccionado del listview2 una posicion
    SubirItemList Me.ListView2
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(62)
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim i As Integer
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    limpiar Me

    For i = 19 To 20
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 23 To 28
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i

    'Ocultar todos los Frames de Formulario
    Me.FrameInfArticulos.visible = False
    
    CommitConexion
    
    cadTitulo = ""
    cadNombreRPT = ""
    
    ListadosAlmacen H, W
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub



Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Articulos
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtCodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmFam_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Familias
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMarcas_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Marcas
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmPro_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Proveedores
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTArt_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Tipo de articulos
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
'Buscar general: cada index llama a una tabla
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 19, 20 'cod. FAMILIA
            indCodigo = Index + 43
            Set frmFam = New frmManFamilias
            frmFam.DatosADevolverBusqueda = "0|1|"
            frmFam.Show vbModal
            Set frmFam = Nothing
            
        Case 21, 22 'cod. MARCA
            indCodigo = Index + 43
            AbrirFrmMarcas
            
        Case 27, 28 'cod. ARTICULO
            indCodigo = Index + 43
            Set frmArt = New frmManArtic
            frmArt.DatosADevolverBusqueda = "0|1|" 'Abrimos en Modo Busqueda
            frmArt.Show vbModal
            Set frmArt = Nothing
            
        Case 25, 26 'cod TIPO ARTICULO
            indCodigo = Index + 43
            AbrirFrmTipoArt

        Case 23, 24 'cod. PROVEEDOR
            indCodigo = Index + 43
            Set frmPro = New frmManProve
            frmManProve.DatosADevolverBusqueda = "0|1"
            frmManProve.Show vbModal
            Set frmManProve = Nothing
            
    End Select
    PonerFoco txtCodigo(indCodigo)
    Screen.MousePointer = vbDefault
End Sub



Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim tabla As String
Dim codCampo As String, nomCampo As String
Dim TipCampo As String, Formato As String
Dim Titulo As String
Dim EsNomCod As Boolean 'Si es campo Cod-Descripcion llama a PonerNombreDeCod


    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    EsNomCod = False
        
    Select Case Index
        Case 70, 71  'Cod. ARTICULO
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "sartic", "nomartic", "codartic", "T")
        
        Case 62, 63 'Cod. FAMILIA
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "sfamia", "nomfamia")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
        
        Case 66, 67 'PROVEEDOR
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "proveedor", "nomprove")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
        
        Case 64, 65 'MARCAS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "marcas", "nommarca")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
        
        Case 68, 69 'Tipos de Articulos
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = UCase(txtCodigo(Index).Text)
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "stipar", "nomtipar", "codtipar", "T")
    End Select
    
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda

    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        Select Case Opcionlistado
            Case 7, 8 'Informe Traspasos Almacen
                txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaDevuelta, 1), "0000000")
                PonerFoco txtCodigo(indCodigo)
            Case 9, 12, 13, 14, 15, 16, 17 '9: Informe Movimiento Articulos
                                'Inventario Articulos
                                '14: Actualizar diferencias Stock Inventariado
                                '16: Listado Valoracion stock inventariado
                txtCodigo(indCodigo).Text = RecuperaValor(CadenaDevuelta, 1)
                txtNombre(indCodigo).Text = RecuperaValor(CadenaDevuelta, 2)
                PonerFoco txtCodigo(indCodigo)
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub ponerFrameArticulosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el informe de Articulos, de tabla: sartic
Dim b As Boolean

    b = True
    H = 6735
    W = 9435
    
    PonerFrameVisible Me.FrameInfArticulos, visible, H, W

    If visible = True Then
        'visible orden campos si opcionlistado=6
        Me.FrameOrden.visible = b
    End If
End Sub


Private Sub CargarListViewOrden()
'Carga el List View del frame: frameInfArticulos
'para establecer el orden en que se van a mostrar los datos en el Informe
'Orden: Familia, MArca, Proveedor, Tipo de Articulo, Articulo
Dim ItmX As ListItem

    'Los encabezados
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "Campo", 1600
    Set ItmX = ListView2.ListItems.Add
    ItmX.Text = "Familia"
    Set ItmX = ListView2.ListItems.Add
    ItmX.Text = "Proveedor"
    Set ItmX = ListView2.ListItems.Add
    ItmX.Text = "Tipo Articulo"
End Sub


Private Sub InicializarVbles()
    cadFormula = ""
    cadselect = ""
    cadParam = ""
    numParam = 0
    conSubRPT = False
End Sub

Private Function PonerDesdeHasta(codD As String, codH As String, nomD As String, nomH As String, param As String) As Boolean
'IN: codD,codH --> codigo Desde/Hasta
'    nomD,nomH --> Descripcion Desde/Hasta
'Añade a cadFormula y cadSelect la cadena de seleccion:
'       "(codigo>=codD AND codigo<=codH)"
' y añade a cadParam la cadena para mostrar en la cabecera informe:
'       "codigo: Desde codD-nomd Hasta: codH-nomH"
Dim devuelve As String
Dim devuelve2 As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(codD, codH, Codigo, TipCod)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    If TipCod <> "F" Then 'Fecha
        If Not AnyadirAFormula(cadselect, devuelve) Then Exit Function
    Else
        devuelve2 = CadenaDesdeHastaBD(codD, codH, Codigo, TipCod)
        If devuelve2 = "Error" Then Exit Function
        If Not AnyadirAFormula(cadselect, devuelve2) Then Exit Function
    End If
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            cadParam = cadParam & AnyadirParametroDH(param, codD, codH, nomD, nomH)
            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function


Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .Titulo = cadTitulo
        .EnvioEMail = False
        .NombreRPT = cadNombreRPT
        .Opcion = 0 'Opcion
        .Show vbModal
    End With
End Sub


Private Function PonerGrupo(numGrupo As Byte, cadgrupo As String) As Byte
Dim campo As String
Dim nomCampo As String

    campo = "pGroup" & numGrupo & "="
    nomCampo = "pGroup" & numGrupo & "Name="
    PonerGrupo = 0
    
    Select Case cadgrupo
        Case "Familia"
            cadParam = cadParam & campo & "{sartic.codfamia}" & "|"
            If numGrupo = 1 Then
                cadParam = cadParam & nomCampo & " ""FAMILIA: "" & " & " totext({sartic.codfamia},""0000"") & " & """  """ & " & {sfamia.nomfamia}" & "|"
            Else
                cadParam = cadParam & nomCampo & " totext({sartic.codfamia},""0000"") & " & """ """ & " & {sfamia.nomfamia}" & "|"
            End If
            numParam = numParam + 1
'        Case "Marca"
'            cadParam = cadParam & campo & "{sartic.codmarca}" & "|"
'            If numGrupo = 1 Then
'                cadParam = cadParam & nomCampo & " ""MARCA: "" & " & " totext({sartic.codmarca},""0000"") & " & """  """ & " & {smarca.nommarca}" & "|"
'            Else
'                cadParam = cadParam & nomCampo & " totext({sartic.codmarca},""0000"") & " & """ """ & " & {smarca.nommarca}" & "|"
'            End If
'            numParam = numParam + 1
        Case "Proveedor"
            cadParam = cadParam & campo & "{sartic.codprove}" & "|"
            If numGrupo = 1 Then
                cadParam = cadParam & nomCampo & " ""PROVEEDOR: "" & " & " totext({sartic.codprove},""000000"") & " & """  """ & " & {proveedor.nomprove}" & "|"
            Else
                cadParam = cadParam & nomCampo & " totext({sartic.codprove},""000000"") & " & """ """ & " & {proveedor.nomprove}" & "|"
            End If
            numParam = numParam + 1
            PonerGrupo = numGrupo
        Case "Tipo Articulo"
            cadParam = cadParam & campo & "{sartic.codtipar}" & "|"
            If numGrupo = 1 Then
                cadParam = cadParam & nomCampo & " ""TIPO ARTICULO: "" & " & " {sartic.codtipar} & " & """  """ & " & {stipar.nomtipar}" & "|"
            Else
                cadParam = cadParam & nomCampo & " {sartic.codtipar} & " & """ """ & " & {stipar.nomtipar}" & "|"
            End If
            numParam = numParam + 1
    End Select

'Case "Familia"
'            cadParam = cadParam & "pGroup1=" & "{sartic.codfamia}" & "|"
'            cadParam = cadParam & "pGroup1Name= ""FAMILIA: "" & " & " totext({sartic.codfamia},""0000"") & " & """  """ & " & {sfamia.nomfamia}" & "|"
'            numParam = numParam + 1
'            Select Case ListView2.ListItems(2).Text
'                Case "Marca"
'                    cadParam = cadParam & "pGroup2=" & "{sartic.codmarca}" & "|"
'                    cadParam = cadParam & "pGroup2Name= ""MARCA: "" & " & " totext({sartic.codmarca},""0000"") & " & """  """ & " & {smarca.nommarca}" & "|"
'                    numParam = numParam + 1
'                    If ListView2.ListItems(3).Text = "Proveedor" Then
'                        Opcion = 1
'                    Else
'                        Opcion = 2
'                    End If
'                Case "Proveedor"
'                Case "Tipo Articulo"
'            End Select
End Function


Private Sub AbrirFrmMarcas()
    Set frmMarcas = New frmManMarcas
    frmMarcas.DatosADevolverBusqueda = "0|1"
    frmMarcas.DeConsulta = True
    frmMarcas.Show vbModal
    Set frmMarcas = Nothing
End Sub



Private Sub AbrirFrmTipoArt()
'Tipos de Articulos
    Set frmTArt = New frmManTipArtic
    frmTArt.DatosADevolverBusqueda = "0|1"
    frmTArt.DeConsulta = True
    frmTArt.Show vbModal
    Set frmTArt = Nothing
End Sub

Private Function ComprobarFechasConta(ind As Integer) As Boolean
'comprobar que el periodo de fechas a contabilizar esta dentro del
'periodo de fechas del ejercicio de la contabilidad
Dim FechaIni As String, FechaFin As String
Dim Cad As String
Dim Rs As ADODB.Recordset
    
On Error GoTo EComprobar

    ComprobarFechasConta = False
    
    If txtCodigo(ind).Text <> "" Then
        FechaIni = "Select fechaini,fechafin From parametros"
        Set Rs = New ADODB.Recordset
        Rs.Open FechaIni, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        If Not Rs.EOF Then
            FechaIni = DBLet(Rs!FechaIni, "F")
            FechaFin = DBLet(Rs!FechaFin, "F") + 365
            'nos guardamos los valores
            Orden1 = FechaIni
            Orden2 = FechaFin
        
            If Not EntreFechas(FechaIni, txtCodigo(ind).Text, FechaFin) Then
                 Cad = "El período de contabilización debe estar dentro del ejercicio:" & vbCrLf & vbCrLf
                 Cad = Cad & "    Desde: " & FechaIni & vbCrLf
                 Cad = Cad & "    Hasta: " & FechaFin
                 MsgBox Cad, vbExclamation
                 txtCodigo(ind).Text = ""
            Else
                ComprobarFechasConta = True
            End If
        End If
        Rs.Close
        Set Rs = Nothing
    Else
        ComprobarFechasConta = True
    End If
    
EComprobar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Fechas", Err.Description
End Function

Private Sub ListadosAlmacen(H As Integer, W As Integer)
   'Listado de Artículo
    ponerFrameArticulosVisible True, H, W
    CargarListViewOrden
    Codigo = "{sartic"
    indFrame = 11
    cadTitulo = "Listado de Artículos"
End Sub


