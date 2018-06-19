VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListMovArtFam 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   11175
   Icon            =   "frmListMovArtFam.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameMovArtic 
      Height          =   5865
      Left            =   0
      TabIndex        =   7
      Top             =   -30
      Width           =   7995
      Begin VB.CheckBox chkSaltaPag 
         Caption         =   "Salta pág. en Familia"
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
         Left            =   3870
         TabIndex        =   24
         Top             =   3915
         Width           =   2505
      End
      Begin VB.Frame FrameValorar 
         Caption         =   "Valorar Con:"
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
         Height          =   1575
         Left            =   540
         TabIndex        =   21
         Top             =   3585
         Width           =   3165
         Begin VB.OptionButton optPrecioUC 
            Caption         =   "Precio Ultima Compra"
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
            TabIndex        =   23
            Top             =   880
            Width           =   2820
         End
         Begin VB.OptionButton optPrecioMP 
            Caption         =   "Precio Medio Ponderado"
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
            TabIndex        =   22
            Top             =   450
            Value           =   -1  'True
            Width           =   2820
         End
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
         Index           =   3
         Left            =   6210
         TabIndex        =   6
         Top             =   5130
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
         Index           =   3
         Left            =   5040
         TabIndex        =   5
         Top             =   5130
         Width           =   1065
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
         Index           =   14
         Left            =   1920
         MaxLength       =   16
         TabIndex        =   1
         Top             =   1725
         Width           =   1500
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
         Index           =   15
         Left            =   1920
         MaxLength       =   16
         TabIndex        =   2
         Top             =   2085
         Width           =   1500
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
         Index           =   16
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   3
         Top             =   2685
         Width           =   660
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
         Index           =   17
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   4
         Top             =   3045
         Width           =   660
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
         Index           =   13
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   0
         Top             =   1080
         Width           =   540
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
         Index           =   14
         Left            =   3440
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   1725
         Width           =   3900
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
         Index           =   15
         Left            =   3440
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   2085
         Width           =   3900
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
         Index           =   16
         Left            =   2600
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text5"
         Top             =   2685
         Width           =   4755
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
         Index           =   17
         Left            =   2600
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text5"
         Top             =   3045
         Width           =   4755
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
         Index           =   13
         Left            =   2480
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Text5"
         Top             =   1080
         Width           =   4830
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
         Index           =   23
         Left            =   990
         TabIndex        =   20
         Top             =   1725
         Width           =   600
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
         Index           =   22
         Left            =   990
         TabIndex        =   19
         Top             =   2085
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Articulo"
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
         Index           =   2
         Left            =   600
         TabIndex        =   17
         Top             =   1440
         Width           =   750
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   11
         Left            =   1635
         Top             =   1725
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   12
         Left            =   1635
         Top             =   2085
         Width           =   240
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
         Index           =   20
         Left            =   990
         TabIndex        =   16
         Top             =   2685
         Width           =   600
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
         Index           =   19
         Left            =   990
         TabIndex        =   15
         Top             =   3045
         Width           =   555
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
         Index           =   3
         Left            =   600
         TabIndex        =   14
         Top             =   2445
         Width           =   660
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   13
         Left            =   1635
         Top             =   2685
         Width           =   240
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   14
         Left            =   1635
         Top             =   3045
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Almacen"
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
         Index           =   1
         Left            =   600
         TabIndex        =   13
         Top             =   1035
         Width           =   825
      End
      Begin VB.Image imgBuscarG 
         Height          =   240
         Index           =   10
         Left            =   1635
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label lbltituloInven 
         Caption         =   "Informe de Movimiento de Articulos"
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
         Left            =   600
         TabIndex        =   18
         Top             =   360
         Width           =   7260
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   10680
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmListMovArtFam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcionlistado As Integer

    '==== Listados de ALMACEN ====
    '=============================
    ' 9 .- Listado Busquedas de movimientos de Artículos
    '10 .-
    
    
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmMtoAlPropios As frmManAlmProp
Attribute frmMtoAlPropios.VB_VarHelpID = -1
Private WithEvents frmMtoTUnidad As frmManTipUnid
Attribute frmMtoTUnidad.VB_VarHelpID = -1
Private WithEvents frmMtoTArticulo As frmManTipArtic
Attribute frmMtoTArticulo.VB_VarHelpID = -1
Private WithEvents frmCta As frmCtasConta
Attribute frmCta.VB_VarHelpID = -1
Private WithEvents frmMtoFamilia As frmManFamilias
Attribute frmMtoFamilia.VB_VarHelpID = -1
Private WithEvents frmMtoProveedor As frmManProve
Attribute frmMtoProveedor.VB_VarHelpID = -1
Private WithEvents frmMtoArticulos As frmManArtic
Attribute frmMtoArticulos.VB_VarHelpID = -1
Private WithEvents frmMtoClientes As frmClientes
Attribute frmMtoClientes.VB_VarHelpID = -1
Private WithEvents frmMtoIncid As frmManInciden
Attribute frmMtoIncid.VB_VarHelpID = -1

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
Private cadNomRPT As String 'Nombre del informe a Imprimir
Private conSubRPT As Boolean 'Si el informe tiene subreports
'-----------------------------------


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

Private Sub cmdAceptar_Click(Index As Integer)
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim cadAux As String
Dim bol As Boolean
   InicializarVbles
   
   '========= Frame Listado Movimiento de Artículos ========================
   'Frame Listado Movimiento de Artículos
   'Nombre fichero .rpt a Imprimir
   cadNomRPT = "rAlmMovimFam.rpt"
    
   If Not PonerFormulaYParametrosInf9() Then Exit Sub
   'comprobar que hay datos para mostrar en el Informe
   cadAux = "sartic LEFT JOIN smoval ON smoval.codartic=sartic.codartic "
   If Not HayRegParaInforme(cadAux, cadselect) Then Exit Sub
   conSubRPT = False
    
   If Me.optPrecioMP.Value Then cadParam = cadParam & "pPrecio=0|"
   If Me.optPrecioUC.Value Then cadParam = cadParam & "pPrecio=1|"
   numParam = numParam + 1
   
   
   If chkSaltaPag.Value = 0 Then cadParam = cadParam & "pSalto=0|"
   If chkSaltaPag.Value = 1 Then cadParam = cadParam & "pSalto=1|"
   numParam = numParam + 1
   
    
    
        
   LlamarImprimir

   Screen.MousePointer = vbDefault
End Sub




Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub




Private Sub Form_Activate()
Dim IndiceFoco As Integer

    If PrimeraVez Then
        PrimeraVez = False
        IndiceFoco = -1
        IndiceFoco = 5
        If IndiceFoco >= 0 Then PonerFoco txtCodigo(IndiceFoco)
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer


    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    limpiar Me

    'Ocultar todos los Frames de Formulario
    FrameMovArtic.visible = False
    
    CommitConexion
    
    CargarIconos
    
    cadTitulo = ""
    cadNomRPT = ""
    
    ListadosAlmacen H, W
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(3).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub

Private Sub Form_Unload(Cancel As Integer)
    NumCod = ""
End Sub

Private Sub frmcta_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtCodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmMtoAlPropios_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Almacenes Propios
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoArticulos_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Articulos
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoClientes_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Clientes
    If indCodigo > 0 Then
        txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
        txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmMtoFamilia_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Familia de Articulos
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoIncid_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoProveedor_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Proveedores
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoTArticulo_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Tipos de Artículo
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoTUnidad_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Tipos de Unidad
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub



Private Sub imgBuscarG_Click(Index As Integer)
'Buscar general: cada index llama a una tabla
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 13, 14 'cod. FAMILIA
            Select Case Index
                Case 13, 14: indCodigo = Index + 3
            End Select
            Set frmMtoFamilia = New frmManFamilias
            frmMtoFamilia.DatosADevolverBusqueda = "0|1|"
            frmMtoFamilia.Show vbModal
            Set frmMtoFamilia = Nothing
        
        Case 10 'cod. ALMACEN
            Select Case Index
                Case 10: indCodigo = Index + 3
            End Select
            AbrirFrmAlmPropios
            
        Case 11, 12 'cod. ARTICULO
            Select Case Index
                Case 11, 12: indCodigo = Index + 3
            End Select
            Set frmMtoArticulos = New frmManArtic
            frmMtoArticulos.DatosADevolverBusqueda = "0|1|" 'Abrimos en Modo Busqueda
            frmMtoArticulos.Show vbModal
            Set frmMtoArticulos = Nothing
            
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
Dim Tabla As String
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
    If Index = 1 Or Index = 2 Then
    'el mismo frame ( y por tanto los mismos campos) se utilizan para distintos
    'informes. Según de donde llamemos código de una tabla u otra
        Select Case Opcionlistado
            Case 1 'Listado MARCAS
                EsNomCod = True
                Tabla = "smarca"
                codCampo = "codmarca"
                nomCampo = "nommarca"
                TipCampo = "N"
                Formato = "0000"
                Titulo = "Marca"
                
            Case 2 'Listado ALMACENES Propios
                EsNomCod = True
                Tabla = "salmpr"
                codCampo = "codalmac"
                nomCampo = "nomalmac"
                TipCampo = "N"
                Formato = "000"
                Titulo = "Almacen Propio"
                
            Case 3 'Listado Tipos UNIDADES
                EsNomCod = True
                Tabla = "sunida"
                codCampo = "codunida"
                nomCampo = "nomunida"
                TipCampo = "N"
                Formato = "00"
                Titulo = "Tipo Unidad"
                
            Case 4 'Listado Tipos Artículos
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), 1, "stipar", "nomtipar", "codtipar", "Tipo de Artículo", "T")
    
            Case 110 'Listado Ubicaciones Almacen
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "subica", "nomubica", "codubica", "Ubicaciones Almacen", "T")
            
            
            Case 20 'Listado ACTIVIDADES de Clientes
                EsNomCod = True
                Tabla = "sactiv"
                codCampo = "codactiv"
                nomCampo = "nomactiv"
                TipCampo = "N"
                Formato = "000"
                Titulo = "Actividad de Cliente"
            
            Case 21 'Listado ZONAS de Clientes
                EsNomCod = True
                Tabla = "szonas"
                codCampo = "codzonas"
                nomCampo = "nomzonas"
                TipCampo = "N"
                Formato = "000"
                Titulo = "Zona de Cliente"
            
            Case 22 'Listado RUTAS de Asistencia
                EsNomCod = True
                Tabla = "srutas"
                codCampo = "codrutas"
                nomCampo = "nomrutas"
                TipCampo = "N"
                Formato = "000"
                Titulo = "Ruta de Asistencia"
            
            Case 23 'Listado Formas de Envío
                EsNomCod = True
                Tabla = "senvio"
                codCampo = "codenvio"
                nomCampo = "nomenvio"
                TipCampo = "N"
                Formato = "000"
                Titulo = "Forma de Envío"
            
            Case 24 'Listado Tarifas Venta
                EsNomCod = True
                Tabla = "starif"
                codCampo = "codlista"
                nomCampo = "nomlista"
                TipCampo = "N"
                Formato = "000"
                Titulo = "Tarifa de Venta"
            
            Case 27 'Listado SITUACIONES Especiales
                EsNomCod = True
                Tabla = "ssitua"
                codCampo = "codsitua"
                nomCampo = "nomsitua"
                TipCampo = "N"
                Formato = "00"
                Titulo = "Situación Especial"
            
            Case 58 'Listado PROVEEDORES
                EsNomCod = True
                Tabla = "proveedor"
                codCampo = "codprove"
                nomCampo = "nomprove"
                TipCampo = "N"
                Formato = "000000"
                Titulo = "Proveedor"
            
            Case 61 'Listado MOTIVOS Pend. Rep.
                EsNomCod = True
                Tabla = "smotre"
                codCampo = "codmotre"
                nomCampo = "nommotre"
                TipCampo = "N"
                Formato = "00"
                Titulo = "Motivos Pend. Rep."
        End Select
        
    ElseIf Index = 3 Or Index = 4 Then
         '7: Informe Traspaso Almacenes
         '8: Informe Movimientos Almacen
         txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
    Else
        Select Case Index
        Case 0, 86, 87
            If txtCodigo(Index).Text <> "" Then
                
                If Index = 0 Then
                    txtNombre(Index).Text = PonerNombreCuenta(txtCodigo(Index), 2)
                    If txtNombre(Index).Text = "" Then
                        MsgBox "Número de Cuenta contable no existe en la contabilidad. Reintroduzca.", vbExclamation
                    End If
                Else
                    PonerFormatoEntero txtCodigo(Index)
                    If (Index = 86 Or Index = 87) Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
                End If
            End If
            
        Case 5, 6, 14, 15, 29, 30, 70, 71, 90, 91, 92, 93 'Cod. ARTICULO
            EsNomCod = True
            Tabla = "sartic"
            codCampo = "codartic"
            nomCampo = "nomartic"
            TipCampo = "T"
            Titulo = "Artículo"
            txtCodigo(Index).Text = UCase(txtCodigo(Index).Text)
        
        Case 7, 8, 16, 17, 25, 26, 62, 63, 75, 76, 88, 89, 94, 95 'Cod. FAMILIA
            EsNomCod = True
            Tabla = "sfamia"
            codCampo = "codfamia"
            nomCampo = "nomfamia"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Familia"
        
        Case 9, 10, 20, 22, 31, 32, 43, 44, 53, 54, 82, 83, 109, 110, 115, 116, 119, 120  'FECHA Desde Hasta
            If txtCodigo(Index).Text <> "" Then
                If Index = 22 And Opcionlistado = 19 Then 'Este campo sera Hora y no Fecha
                    PonerFormatoHora txtCodigo(Index)
                Else
                    PonerFormatoFecha txtCodigo(Index)
                    If Opcionlistado = 223 And txtCodigo(Index).Text <> "" Then
                        'Contabilizar facturas
                        If Not ComprobarFechasConta(Index) Then
                            PonerFoco txtCodigo(Index)
'                        Else '++monica
'                            If OptClientes.Value Then
'                                PonerFoco txtCodigo(0)
'                            Else
'                                cmdCancel(7).SetFocus
'                            End If
                        End If '++
                    End If
                    
                End If
            End If
        
        Case 11, 12, 13, 72 'ALMACENES Propios
            EsNomCod = True
            Tabla = "salmpr"
            codCampo = "codalmac"
            nomCampo = "nomalmac"
            TipCampo = "N"
            Formato = "000"
            Titulo = "Almacen Propio"
            
        Case 18, 19, 66, 67, 79, 80 'PROVEEDOR
            EsNomCod = True
            Tabla = "sprove"
            codCampo = "codprove"
            nomCampo = "nomprove"
            TipCampo = "N"
            Formato = "000000"
            Titulo = "Proveedor"
        
        Case 21, 96, 97, 111 'Cod. Operario/Trabajador
            EsNomCod = True
            Tabla = "straba"
            codCampo = "codtraba"
            nomCampo = "nomtraba"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Trabajador"
        
        Case 23, 24, 107
            EsNomCod = True
            TipCampo = "N"
            If Opcionlistado = 30 Then 'Precios Especiales
                Tabla = "sclien"
                codCampo = "codclien"
                nomCampo = "nomclien"
                Formato = "000000"
                Titulo = "Cliente"
            Else   'Tarifas Precios
                Tabla = "starif"
                codCampo = "codlista"
                nomCampo = "nomlista"
                Formato = "000"
                Titulo = "Tarifa de Venta"
            End If
        
        Case 27, 28, 64, 65, 77, 78 'MARCAS
            EsNomCod = True
            Tabla = "smarca"
            codCampo = "codmarca"
            nomCampo = "nommarca"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Marca"
        
        Case 31 'Nº de Oferta
            If txtCodigo(Index).Text = "" Then Exit Sub
            codCampo = DevuelveDesdeBDNew(cAgro, "scapre", "numofert", "numofert", txtCodigo(Index).Text, "N")
            If codCampo = "" Then
                MsgBox "No existe el código de Oferta: " & NumCod, vbInformation
                PonerFoco txtCodigo(Index)
            Else
                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
            End If
            
        Case 32, 43 'Carta de la Oferta
            EsNomCod = True
            Tabla = "scartas"
            codCampo = "codcarta"
            nomCampo = "descarta"
            TipCampo = "N"
            Formato = "000"
            Titulo = "Cartas para Ofertas"
            
        Case 37, 38, 34, 47, 48, 55, 56, 73, 74, 98, 101, 102, 103, 117, 118 'Cod. CLIENTE
            EsNomCod = True
            Tabla = "sclien"
            codCampo = "codclien"
            nomCampo = "nomclien"
            TipCampo = "N"
            Formato = "000000"
            Titulo = "Cliente"
            
        Case 112, 113, 114
            EsNomCod = True
            Tabla = "inciden"
            codCampo = "codincid"
            nomCampo = "nomincid"
            TipCampo = "T"
            'Formato = "0000"
            Titulo = "Incidencias"
        
        Case 41, 42, 59, 60 'Nº Contrato
'            If txtCodigo(Index).Text <> "" Then
'                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
'            End If

        Case 45, 46, 106, 108 'ZONAS del Cliente
            EsNomCod = True
            Tabla = "szonas"
            codCampo = "codzonas"
            nomCampo = "nomzonas"
            TipCampo = "N"
            Formato = "000"
            Titulo = "Zonas de Clientes"
        
        Case 49, 50 'Cod. AGENTE
            EsNomCod = True
            Tabla = "sagent"
            codCampo = "codagent"
            nomCampo = "nomagent"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Agente"
            
        Case 51, 52, 57, 58, 104, 105 'Tipos Contratos/MAntenimientos
            EsNomCod = True
            Tabla = "stipco"
            codCampo = "codtipco"
            nomCampo = "nomtipco"
            TipCampo = "T"
            Titulo = "Tipos de Contratos"
            
        Case 61 'Año Ejercicio
            If txtCodigo(Index).Text = "" Then Exit Sub
            If Not IsNumeric(txtCodigo(Index).Text) Then
                MsgBox "El Ejercicio debe ser un Año", vbInformation
                Exit Sub
            End If
        
        Case 68, 69 'Tipos de Articulos
            EsNomCod = True
            Tabla = "stipar"
            codCampo = "codtipar"
            nomCampo = "nomtipar"
            TipCampo = "T"
            Titulo = "Tipo de Articulo"
            
        Case 84, 85 'RUTAS del cliente
            EsNomCod = True
            Tabla = "srutas"
            codCampo = "codrutas"
            nomCampo = "nomrutas"
            TipCampo = "N"
            Formato = "000"
            Titulo = "Ruta de Asistencia"
        End Select
    End If
    
    If EsNomCod Then
        If TipCampo = "N" Then
            If PonerFormatoEntero(txtCodigo(Index)) Then
                
                
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), Tabla, nomCampo, codCampo)
'                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), Tabla, NomCampo, codCampo, Titulo, TipCampo)
            
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, Formato)
            Else
                txtNombre(Index).Text = ""
            End If
        Else
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), Tabla, nomCampo, codCampo)
'            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), Tabla, NomCampo, codCampo, Titulo, TipCampo)
        End If
    End If
    
   
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim Tabla As String
Dim Titulo As String

    'Llamamos a al form
    Cad = ""
    Conexion = cAgro    'Conexión a BD: Ariges
    Select Case Opcionlistado
        Case 7 'Traspaso de Almacenes
            Cad = Cad & "Nº Trasp|scatra.codtrasp|N|0000000|40·Almacen Origen|scatra.almaorig|N|000|20·Almacen Destino|scatra.almadest|N|000|20·Fecha|scatra.fechatra|F||20·"
            Tabla = "scatra"
            Titulo = "Traspaso Almacenes"
        Case 8 'Movimientos de Almacen
            Cad = Cad & "Nº Movim.|scamov.codmovim|N|0000000|40·Almacen|scamov.codalmac|N|000|30·Fecha|scamov|fecmovim|F||30·"
            Tabla = "scamov"
            Titulo = "Movimientos Almacen"
        Case 9, 12, 13, 14, 15, 16, 17 '9: Movimientos Articulos
                   '12: Inventario Articulos
                   '14:Actualizar Diferencias de Stock Inventariado
                   '16: Listado Valoracion stock inventariado
            Cad = Cad & "Código|sartic.codartic|T||30·Denominacion|sartic.nomartic|T||70·"
            Tabla = "sartic"
            Titulo = "Articulos"
    End Select
          
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vtabla = Tabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        '###A mano
        'frmB.vDevuelve = "0|1|"
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = Titulo
        frmB.vSelElem = 0
'        frmB.vConexionGrid = Conexion
'        frmB.vBuscaPrevia = 1
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
'        End If
    End If
    Screen.MousePointer = vbDefault
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




Private Function PonerFormulaYParametrosInf9() As Boolean
Dim Cad As String
Dim todosMarcados As Boolean
Dim devuelve As String
Dim i As Byte

    PonerFormulaYParametrosInf9 = False
    InicializarVbles
    
    'Parametro EMPRESA
    cadParam = "|pNomEmpre=""" & vParam.NombreEmpresa & """|"
    numParam = 1
        
        
    'Cadena para seleccion De ALMACEN
    '-----------------------------------
    Codigo = "{salmac.codalmac}"
    If Trim(txtCodigo(13).Text) <> "" Then _
    devuelve = Codigo & " = " & Val(txtCodigo(13).Text)
    If devuelve <> "" Then
        cadFormula = devuelve
        Cad = "pAlmacen= ""Almacen: " & Format(txtCodigo(13).Text, "000") & " " & txtNombre(13).Text
        cadParam = cadParam & Cad & """|"
        numParam = numParam + 1
    End If
        
    'Cadena para seleccion Desde y Hasta ARTICULO
    If txtCodigo(14).Text <> "" Or txtCodigo(15).Text <> "" Then
        Codigo = "{sartic.codartic}"
        devuelve = "pDHArticulo=""Artículo: "
        If Not PonerDesdeHasta(Codigo, "T", 14, 15, devuelve) Then Exit Function
    End If
                    
    'Cadena para seleccion Desde y Hasta FAMILIA
    If txtCodigo(16).Text <> "" Or txtCodigo(17).Text <> "" Then
        Codigo = "{sartic.codfamia}"
        devuelve = "pDHFamilia=""Familia: "
        If Not PonerDesdeHasta(Codigo, "N", 16, 17, devuelve) Then Exit Function
    End If
    
    PonerFormulaYParametrosInf9 = True
    
End Function



Private Function AnyadirParametroDH(Cad As String, indD As Byte, indH As Byte) As String
On Error Resume Next
    
     If txtCodigo(indD).Text <> "" Then
        Cad = Cad & "desde " & txtCodigo(indD).Text
        If txtNombre(indD).Text <> "" Then Cad = Cad & " - " & txtNombre(indD).Text
    End If
    If txtCodigo(indH).Text <> "" Then
        Cad = Cad & "  hasta " & txtCodigo(indH).Text
        If txtNombre(indH).Text <> "" Then Cad = Cad & " - " & txtNombre(indH).Text
    End If
    
    AnyadirParametroDH = Cad
    If Err.Number <> 0 Then Err.Clear
End Function


Private Sub InicializarVbles()
    cadFormula = ""
    cadselect = ""
    cadParam = ""
    numParam = 0
'    cadTitulo = ""
'    cadNomRPT = ""
    conSubRPT = False
End Sub


Private Function PonerDesdeHasta(campo As String, Tipo As String, indD As Byte, indH As Byte, param As String) As Boolean
Dim devuelve As String
Dim Cad As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    
    If Tipo <> "F" Then
        'Fecha para Crystal Report
        If Not AnyadirAFormula(cadselect, devuelve) Then Exit Function
    Else
        'Fecha para la Base de Datos
        Cad = CadenaDesdeHastaBD(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
        If Not AnyadirAFormula(cadselect, Cad) Then Exit Function
    End If
    
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            cadParam = cadParam & AnyadirParametroDH(param, indD, indH) & """|"
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
        .EnvioEMail = False
        .Opcion = Opcionlistado
        .Titulo = cadTitulo
        .NombreRPT = cadNomRPT
        .ConSubInforme = conSubRPT
        .Show vbModal
    End With
End Sub




Private Sub AbrirFrmAlmPropios()
    Set frmMtoAlPropios = New frmManAlmProp
    frmMtoAlPropios.DatosADevolverBusqueda = "0|1|"
    frmMtoAlPropios.DeConsulta = True
    frmMtoAlPropios.Show vbModal
    Set frmMtoAlPropios = Nothing
End Sub


Private Sub AbrirFrmTipoArt()
'Tipos de Articulos
    Set frmMtoTArticulo = New frmManTipArtic
    frmMtoTArticulo.DatosADevolverBusqueda = "0|1|"
    frmMtoTArticulo.DeConsulta = True
    frmMtoTArticulo.Show vbModal
    Set frmMtoTArticulo = Nothing
End Sub

Private Sub AbrirFrmClientes()
'Clientes
    Set frmMtoClientes = New frmClientes
    frmMtoClientes.DatosADevolverBusqueda = "0|1|"
    frmMtoClientes.Show vbModal
    Set frmMtoClientes = Nothing
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
            FechaFin = DateAdd("yyyy", 1, DBLet(Rs!FechaFin, "F"))
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
    'LISTADOS DE ALMACENES
    'Informe Movimiento Artículos
     W = 7995
     H = 5865
     PonerFrameVisible Me.FrameMovArtic, True, H, W
     indFrame = 3
     Codigo = "{smoval.codartic}"
     cadTitulo = "Informe Movimientos Articulos por Familia"
     conSubRPT = False
End Sub


Private Sub CargarIconos()
Dim i As Integer
    
    For i = 10 To 14
        Me.imgBuscarG(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i

End Sub
