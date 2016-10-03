VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListOficiales 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6630
   Icon            =   "frmListOficiales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameCobros 
      Height          =   6285
      Left            =   45
      TabIndex        =   8
      Top             =   0
      Width           =   6555
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "Text5"
         Top             =   1395
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Text5"
         Top             =   1020
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1725
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1395
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1725
         MaxLength       =   6
         TabIndex        =   0
         Top             =   1020
         Width           =   830
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo Hanegadas"
         ForeColor       =   &H00972E0B&
         Height          =   1455
         Left            =   360
         TabIndex        =   19
         Top             =   3750
         Width           =   1815
         Begin VB.OptionButton Option3 
            Caption         =   "Catastro"
            Height          =   225
            Index           =   2
            Left            =   300
            TabIndex        =   22
            Top             =   990
            Width           =   1035
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Sigpac"
            Height          =   225
            Index           =   1
            Left            =   300
            TabIndex        =   21
            Top             =   660
            Width           =   1305
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Cooperativa"
            Height          =   225
            Index           =   0
            Left            =   300
            TabIndex        =   20
            Top             =   330
            Width           =   1305
         End
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3390
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3030
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5040
         TabIndex        =   7
         Top             =   5610
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3870
         TabIndex        =   6
         Top             =   5610
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   2
         Top             =   2070
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   3
         Top             =   2445
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text5"
         Top             =   2070
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text5"
         Top             =   2445
         Width           =   3135
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1455
         MouseIcon       =   "frmListOficiales.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar producto"
         Top             =   1395
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1455
         MouseIcon       =   "frmListOficiales.frx":015E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar producto"
         Top             =   1020
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   27
         Top             =   780
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   840
         TabIndex        =   26
         Top             =   1395
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   825
         TabIndex        =   25
         Top             =   1020
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   195
         Left            =   585
         TabIndex        =   18
         Top             =   5340
         Width           =   5145
      End
      Begin VB.Label Label1 
         Caption         =   "O.P.A. 4"
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
         Left            =   450
         TabIndex        =   17
         Top             =   270
         Width           =   5160
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Albarán"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   420
         TabIndex        =   16
         Top             =   2730
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   780
         TabIndex        =   15
         Top             =   3030
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   780
         TabIndex        =   14
         Top             =   3390
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1440
         Picture         =   "frmListOficiales.frx":02B0
         ToolTipText     =   "Buscar fecha"
         Top             =   3030
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1440
         Picture         =   "frmListOficiales.frx":033B
         ToolTipText     =   "Buscar fecha"
         Top             =   3390
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   810
         TabIndex        =   13
         Top             =   2070
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   825
         TabIndex        =   12
         Top             =   2445
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   11
         Left            =   435
         TabIndex        =   11
         Top             =   1830
         Width           =   630
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1440
         MouseIcon       =   "frmListOficiales.frx":03C6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   2100
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1440
         MouseIcon       =   "frmListOficiales.frx":0518
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   2445
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmListOficiales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MANOLO +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmVar As frmManVariedad 'Variedad
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmPro As frmManProductos 'Productos
Attribute frmPro.VB_VarHelpID = -1
Private WithEvents frmCli As frmClientes 'Clientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmTra As frmManAgencias 'Agencias de transporte
Attribute frmTra.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmMensVariedad As frmMensajes 'mensajes
Attribute frmMensVariedad.VB_VarHelpID = -1
Private WithEvents frmMensProducto As frmMensajes 'mensajes
Attribute frmMensProducto.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadselect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean
Dim vsqlVariedad As String
Dim vsqlProducto As String
Dim SqlProds As String


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim cadTABLA1 As String
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim cadSelect1 As String
Dim Sql As String
Dim NSocs As Long
Dim Sql2 As String

InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
     '======== FORMULA  ====================================
    'Seleccionar registros de la empresa conectada
'    Codigo = "{" & tabla & ".codempre}=" & vEmpresa.codEmpre
'    If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
'    If Not AnyadirAFormula(cadSelect, Codigo) Then Exit Sub
    
    '[Monica]20/09/2013: Introducimos el producto
    'D/H Producto
    cDesde = Trim(txtCodigo(4).Text)
    cHasta = Trim(txtCodigo(5).Text)
    nDesde = txtNombre(4).Text
    nHasta = txtNombre(5).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{variedades.codprodu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHProducto= """) Then Exit Sub
    End If
    vsqlProducto = ""
    If txtCodigo(4).Text <> "" Then vsqlProducto = vsqlProducto & " and productos.codprodu >= " & DBSet(txtCodigo(4).Text, "N")
    If txtCodigo(5).Text <> "" Then vsqlProducto = vsqlProducto & " and productos.codprodu <= " & DBSet(txtCodigo(5).Text, "N")
    
    
    Set frmMensProducto = New frmMensajes

    frmMensProducto.OpcionMensaje = 21
    frmMensProducto.Label5 = "Productos"
    frmMensProducto.cadWHERE = vsqlProducto
    frmMensProducto.Show vbModal

    Set frmMensProducto = Nothing
    
    If SqlProds = " and variedades.codprodu = -1 " Then
        MsgBox "No hay datos para mostrar en el Informe.", vbInformation
        Exit Sub
    End If
    
    cadselect = ""
    cadFormula = ""
    
    'D/H Variedad
    cDesde = Trim(txtCodigo(0).Text)
    cHasta = Trim(txtCodigo(1).Text)
    nDesde = txtNombre(0).Text
    nHasta = txtNombre(1).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".codvarie}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad= """) Then Exit Sub
    End If
    
    vsqlVariedad = ""
    If txtCodigo(0).Text <> "" Then vsqlVariedad = vsqlVariedad & " and variedades.codvarie >= " & DBSet(txtCodigo(0).Text, "N")
    If txtCodigo(1).Text <> "" Then vsqlVariedad = vsqlVariedad & " and variedades.codvarie <= " & DBSet(txtCodigo(1).Text, "N")
    
    'D/H Fecha Albarán
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".fecalbar}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    ' en rhisfruta seleccionamos lo de horto
    cadTABLA = "((" & tabla & " INNER JOIN variedades ON rhisfruta.codvarie = variedades.codvarie) "
    cadTABLA = cadTABLA & " INNER JOIN productos ON variedades.codprodu = productos.codprodu) "
    cadTABLA = cadTABLA & " INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
    cadTABLA = cadTABLA & " and grupopro.codgrupo <> 5 and grupopro.codgrupo <> 6 " ' grupo no puede ser 5=almazara ni 6=bodega
    
    
    
    Set frmMensVariedad = New frmMensajes

    frmMensVariedad.OpcionMensaje = 21
    frmMensVariedad.Label5 = "Variedades"
    frmMensVariedad.cadWHERE = vsqlVariedad & SqlProds
    frmMensVariedad.Show vbModal

    Set frmMensVariedad = Nothing
    
    ' en albaranes está sólo lo de horto
    cadTABLA1 = "(albaran_variedad INNER JOIN variedades ON albaran_variedad.codvarie = variedades.codvarie) "
    cadTABLA1 = cadTABLA1 & " INNER JOIN albaran ON albaran_variedad.numalbar = albaran.numalbar "
    
    cadSelect1 = Replace(cadselect, "rhisfruta.fecalbar", "albaran.fechaalb")
    cadSelect1 = Replace(cadSelect1, "rhisfruta.codvarie", "albaran_variedad.codvarie")
    
    If HayRegistros(cadTABLA, cadselect, cadTABLA1, cadSelect1) Then
        If CargarTemporal(cadTABLA, cadselect, cadTABLA1, cadSelect1) Then
        
            Sql2 = "SELECT COUNT(*) from (SELECT DISTINCT codsocio  FROM  " & cadTABLA
            If cadselect <> "" Then Sql2 = Sql2 & " where " & cadselect
            Sql2 = Sql2 & ") aaaa "
            
            NSocs = DevuelveValor(Sql2)
            
            cadParam = cadParam & "pSocios=" & NSocs & "|"
            numParam = numParam + 1
            
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            cadTitulo = "Facturas Pendientes"
            cadNombreRPT = "rInfOPA4.rpt"
            
            LlamarImprimir
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
     For H = 0 To 3
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Next H

    '###Descomentar
'    CommitConexion
         
    Me.Option3(0).Value = True
         
    FrameCobrosVisible True, H, W
    indFrame = 5
    tabla = "rhisfruta"
    Me.Label2.Caption = ""
    Me.Refresh
        
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub


Private Sub frmMensVariedad_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        Sql = " {rhisfruta.codvarie} in (" & CadenaSeleccion & ")"
        Sql2 = " {rhisfruta.codvarie} in [" & CadenaSeleccion & "]"
    Else
        Sql = " {rhisfruta.codvarie} = -1 "
    End If
    If Not AnyadirAFormula(cadselect, Sql) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub
End Sub

Private Sub frmMensProducto_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        Sql = " {variedades.codprodu} in (" & CadenaSeleccion & ")"
        Sql2 = " {variedades.codprodu} in [" & CadenaSeleccion & "]"
        SqlProds = " and variedades.codprodu in (" & CadenaSeleccion & ")"
    Else
        Sql = " and variedades.codprodu = -1 "
        SqlProds = " and variedades.codprodu = -1 "
    End If
End Sub



Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de variedades
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgFec_Click(Index As Integer)
'FEchas
    Dim esq, dalt As Long
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
       
    ' es desplega dalt i cap a la esquerra
    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + 420 + 30

    ' ***canviar l'index de imgFec pel 1r index de les imagens de buscar data***
    imgFec(2).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtCodigo(Index).Text <> "" Then frmC.NovaData = txtCodigo(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtCodigo(CByte(imgFec(2).Tag))
    ' ***************************
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1 'variedades
            AbrirFrmVariedades (Index)
        
        
        Case 2, 3 ' Productos
            AbrirFrmProductos (Index)
        
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub Optcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub OptNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub Option3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Option3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'14/02/2007 antes
'    KEYpress KeyAscii
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'variedad desde
            Case 1: KEYBusqueda KeyAscii, 1 'variedad hasta
            Case 2: KEYFecha KeyAscii, 2 'fecha desde
            Case 3: KEYFecha KeyAscii, 3 'fecha hasta
            Case 4: KEYBusqueda KeyAscii, 2 'producto desde
            Case 5: KEYBusqueda KeyAscii, 3 'producto hasta
            
            Case 6: KEYFecha KeyAscii, 6 'fecha de calculo
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
            
        Case 0, 1 'VARIEDADES
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
        
        Case 4, 5 'PRODUCTOS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "productos", "nomprodu", "codprodu", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
        
        Case 2, 3, 6 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 6285
        Me.FrameCobros.Width = 6690
        W = Me.FrameCobros.Width
        H = Me.FrameCobros.Height
    End If
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadselect = ""
    cadParam = ""
    numParam = 0
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
        .FormulaSeleccion = "{tmpinformes.codusu} = " & vUsu.Codigo
        .OtrosParametros = cadParam
        .NumeroParametros = numParam + 1
        .SoloImprimir = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .EnvioEMail = False
        .Opcion = 0
        .ConSubInforme = True
        .Show vbModal
    End With
End Sub

Private Sub AbrirFrmClientes(indice As Integer)
    indCodigo = indice
    Set frmCli = New frmClientes
    frmCli.DatosADevolverBusqueda = "0|2|"
    frmCli.Show vbModal
    Set frmCli = Nothing
End Sub

Private Sub AbrirFrmVariedades(indice As Integer)
    indCodigo = indice
    Set frmVar = New frmManVariedad
    frmVar.DatosADevolverBusqueda = "0|1|"
    frmVar.DeConsulta = True
    frmVar.CodigoActual = txtCodigo(indCodigo)
    frmVar.Show vbModal
    Set frmVar = Nothing
End Sub

Private Sub AbrirFrmProductos(indice As Integer)
    indCodigo = indice
    Set frmPro = New frmManProductos
    frmPro.DatosADevolverBusqueda = "0|1|"
    frmPro.DeConsulta = True
    frmPro.CodigoActual = txtCodigo(indCodigo)
    frmPro.Show vbModal
    Set frmPro = Nothing
End Sub


Private Sub AbrirVisReport()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = cadFormula
        '.SoloImprimir = (Me.OptVisualizar(indFrame).Value = 1)
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        '##descomen
'        .MostrarTree = MostrarTree
'        .Informe = MIPATH & Nombre
'        .InfConta = InfConta
        '##
        
'        If NombreSubRptConta <> "" Then
'            .SubInformeConta = NombreSubRptConta
'        Else
'            .SubInformeConta = ""
'        End If
        '##descomen
'        .ConSubInforme = ConSubInforme
        '##
        .Opcion = ""
'        .ExportarPDF = (chkEMAIL.Value = 1)
        .Show vbModal
    End With
    
'    If Me.chkEMAIL.Value = 1 Then
'    '####Descomentar
'        If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
'    End If
    Unload Me
End Sub

Private Sub AbrirEMail()
    If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
End Sub


Private Function CargarTemporal(cTabla As String, cadWHERE As String, cTABLA1 As String, cadwhere1 As String) As Boolean
Dim Sql As String
Dim SQL1 As String
Dim Sql2 As String
Dim Sql4 As String
Dim Codiva As String
Dim PorcIva As String
Dim i As Integer
Dim HayReg As Integer
Dim b As Boolean
Dim Registro As String
Dim CADENA As String
Dim vCliente As CCliente
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Importe As Currency
Dim ImporteFacturado As Currency
Dim Dias As Long
Dim Variedad As Long
Dim VarieAnt As Long
Dim Superficie As Currency
Dim KilosNet As Long
Dim KilosNetVC As Long

Dim Interior As Long
Dim Industria As Long
Dim Exportacion As Long
Dim Retirada As Long

On Error GoTo eCargarTemporal

    HayReg = 0
    
    CargarTemporal = False
    
    conn.Execute "delete from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
        
    Sql = "Select rhisfruta.codvarie FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cadWHERE <> "" Then
        cadWHERE = QuitarCaracterACadena(cadWHERE, "{")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "}")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "_1")
        Sql = Sql & " WHERE " & cadWHERE
    End If
    
    Sql = Sql & " group by 1 "
    Sql = Sql & " union "
    Sql = Sql & " Select albaran_variedad.codvarie FROM " & QuitarCaracterACadena(cTABLA1, "_1")
    If cadwhere1 <> "" Then
        cadwhere1 = QuitarCaracterACadena(cadwhere1, "{")
        cadwhere1 = QuitarCaracterACadena(cadwhere1, "}")
        cadwhere1 = QuitarCaracterACadena(cadwhere1, "_1")
        Sql = Sql & " WHERE " & cadwhere1
    End If
    Sql = Sql & " group by 1 "
    Sql = Sql & " order by 1"
    
                                  '(codusu, variedad,  superficie,kilosnetEnt,kilosnetSal, kilosnetVC, kilosmer1, kilosmer2, kilosmer3, kilosmer4)
    Sql4 = "insert into tmpinformes (codusu, codigo1,  importe1,  importe2,   importe3,   importe4,  importeb1, importeb2, importeb3, importeb4) values "
    
    CADENA = ""
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        VarieAnt = DBLet(Rs!codvarie, "N")
                                  '(codusu, variedad)
        CADENA = "(" & vUsu.Codigo & ","
        CADENA = CADENA & DBSet(Rs!codvarie, "N") & ",0,0,0,0,0,0,0,0),"
    End If
    
    While Not Rs.EOF
        Variedad = DBLet(Rs!codvarie, "N")
        If VarieAnt <> Variedad Then
            CADENA = CADENA & "(" & vUsu.Codigo & ","
            CADENA = CADENA & DBSet(Variedad, "N") & ",0,0,0,0,0,0,0,0),"
            
            VarieAnt = DBLet(Variedad, "N")
        End If
    
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    If CADENA <> "" Then
                                      '(codusu, variedad, superficie,kilosnetEnt,kilosnetSal,kilosVC,  kilosmer1, kilosmer2, kilosmer3, kilosmer4)
        Sql = "insert into tmpinformes (codusu, codigo1,  importe1,  importe2,   importe3,   importe4, importeb1, importeb2, importeb3, importeb4) values "
        Sql = Sql & Mid(CADENA, 1, Len(CADENA) - 1) ' quitamos la ultima coma
        conn.Execute Sql
    
        Sql = "select codigo1 from tmpinformes where codusu = " & vUsu.Codigo & " order by codigo1 "
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs.EOF
            ' cargamos los kilos entrados distintos de venta campo y la superficie
            Sql2 = "select rhisfruta.codvarie, sum(kilosnet) kilosnet  "
            Sql2 = Sql2 & " from rhisfruta where rhisfruta.codvarie = " & DBSet(Rs!Codigo1, "N")
            Sql2 = Sql2 & " and rhisfruta.tipoentr <> 1 "
            Sql2 = Sql2 & " group by 1 "
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
            KilosNet = 0
            If Not Rs2.EOF Then
                KilosNet = DBLet(Rs2!KilosNet, "N")
            End If
            Set Rs2 = Nothing
            
            
            ' cargamos los kilos entrados de venta campo
            Sql2 = "select rhisfruta.codvarie, sum(kilosnet) kilosnet  "
            Sql2 = Sql2 & " from rhisfruta where rhisfruta.codvarie = " & DBSet(Rs!Codigo1, "N")
            Sql2 = Sql2 & " and rhisfruta.tipoentr = 1 "
            Sql2 = Sql2 & " group by 1 "
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
            KilosNetVC = 0
            If Not Rs2.EOF Then
                KilosNetVC = DBLet(Rs2!KilosNet, "N")
            End If
            Set Rs2 = Nothing
            
            
            Sql2 = "select rcampos.codvarie,   "
            
            If Option3(0).Value Then Sql2 = Sql2 & " sum(supcoope) superficie " 'cooperativa
            If Option3(1).Value Then Sql2 = Sql2 & " sum(supsigpa) superficie " 'sigpac
            If Option3(2).Value Then Sql2 = Sql2 & " sum(supcatas) superficie " 'catastro
            
            Sql2 = Sql2 & " from rcampos  where rcampos.codvarie = " & DBSet(Rs!Codigo1, "N")
            Sql2 = Sql2 & " and rcampos.fecbajas is null "
            Sql2 = Sql2 & " group by 1 "
        
            If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
                Sql2 = "select rcampos.codvarie,   "
                
                If Option3(0).Value Then
                    Sql2 = Sql2 & " sum(supcoope) superficie " 'cooperativa
                    Sql2 = Sql2 & " from rcampos  where rcampos.codvarie = " & DBSet(Rs!Codigo1, "N")
                    Sql2 = Sql2 & " and rcampos.fecbajas is null "
                    Sql2 = Sql2 & " group by 1 "
                End If
                If Option3(1).Value Then
                    Sql2 = Sql2 & " sum(supcultcatas) superficie " 'sigpac
                    Sql2 = Sql2 & " from rcampos inner join rcampos_parcelas on rcampos.codcampo = rcampos_parcelas where rcampos.codvarie = " & DBSet(Rs!Codigo1, "N")
                    Sql2 = Sql2 & " and rcampos.fecbajas is null "
                    Sql2 = Sql2 & " group by 1 "
                End If
                If Option3(2).Value Then
                    Sql2 = Sql2 & " sum(supcatas) superficie " 'catastro
                    Sql2 = Sql2 & " from rcampos  where rcampos.codvarie = " & DBSet(Rs!Codigo1, "N")
                    Sql2 = Sql2 & " and rcampos.fecbajas is null "
                    Sql2 = Sql2 & " group by 1 "
                End If
            End If
        
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
            Superficie = 0
            If Not Rs2.EOF Then
                Superficie = DBLet(Rs2!Superficie, "N")
            End If
            Set Rs2 = Nothing
            
            
            ' calculo de los kilos salidos de esa variedad
            Sql2 = "select tipomer.tiptimer, sum(pesoneto) pesoneto from albaran, albaran_variedad, tipomer " 'destinos, tipomer "
            Sql2 = Sql2 & " where albaran_variedad.codvarie = " & DBSet(Rs!Codigo1, "N")
            Sql2 = Sql2 & " and albaran.numalbar = albaran_variedad.numalbar "
            Sql2 = Sql2 & " and albaran.codtimer = tipomer.codtimer "
            
            If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
                Sql2 = Sql2 & " and albaran.codtimer <> 0 "
            End If
            
            
            If cadwhere1 <> "" Then
                Sql2 = Sql2 & " and " & cadwhere1
            End If
            Sql2 = Sql2 & " group by 1 "
            
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            Interior = 0
            Exportacion = 0
            Industria = 0
            Retirada = 0
            
            While Not Rs2.EOF
                If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
                    Select Case DBLet(Rs2!tiptimer, "N")
                        Case 0, 4
                            Interior = Interior + DBLet(Rs2!Pesoneto, "N")
                        Case 1
                            Exportacion = DBLet(Rs2!Pesoneto, "N")
                        Case 2
                            Industria = DBLet(Rs2!Pesoneto, "N")
                        Case 3
                            Retirada = DBLet(Rs2!Pesoneto, "N")
                    End Select
                Else
                    Select Case DBLet(Rs2!tiptimer, "N")
                        Case 0
                            Interior = DBLet(Rs2!Pesoneto, "N")
                        Case 1
                            Exportacion = DBLet(Rs2!Pesoneto, "N")
                        Case 2
                            Industria = DBLet(Rs2!Pesoneto, "N")
                        Case 3
                            Retirada = DBLet(Rs2!Pesoneto, "N")
                    End Select
                End If
                Rs2.MoveNext
            Wend
            Set Rs2 = Nothing
        
            Sql = "update tmpinformes set importe1 = " & DBSet(Superficie, "N")
            Sql = Sql & ", importe2 = " & DBSet(KilosNet, "N")
            Sql = Sql & ", importe3 = " & DBSet((Interior + Exportacion + Industria + Retirada), "N")
            Sql = Sql & ", importe4 = " & DBSet(KilosNetVC, "N")
            Sql = Sql & ", importeb1 = " & DBSet(Interior, "N")
            Sql = Sql & ", importeb2 = " & DBSet(Exportacion, "N")
            Sql = Sql & ", importeb3 = " & DBSet(Industria, "N")
            Sql = Sql & ", importeb4 = " & DBSet(Retirada, "N")
            Sql = Sql & " where codusu = " & vUsu.Codigo
            Sql = Sql & " and codigo1 = " & DBSet(Rs!Codigo1, "N")
        
            conn.Execute Sql
        
            Rs.MoveNext
        Wend
        Set Rs = Nothing
        
    End If
    
        
    CargarTemporal = True

    Exit Function

eCargarTemporal:
    MuestraError Err.Number, "Cargando Temporal", Err.Description
End Function

Private Function DatosOk() As Boolean
    DatosOk = True
End Function



Private Function HayRegistros(cTabla As String, cWhere As String, cTABLA1 As String, cWhere1 As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
Dim Rs As ADODB.Recordset

    Sql = "Select rhisfruta.codvarie FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    Sql = Sql & " group by 1 "
    Sql = Sql & " union "
    Sql = Sql & " Select albaran_variedad.codvarie FROM " & QuitarCaracterACadena(cTABLA1, "_1")
    If cWhere1 <> "" Then
        cWhere1 = QuitarCaracterACadena(cWhere1, "{")
        cWhere1 = QuitarCaracterACadena(cWhere1, "}")
        cWhere1 = QuitarCaracterACadena(cWhere1, "_1")
        Sql = Sql & " WHERE " & cWhere1
    End If
    Sql = Sql & " group by 1 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Rs.EOF Then
        MsgBox "No hay datos para mostrar en el Informe.", vbInformation
        HayRegistros = False
    Else
        HayRegistros = True
    End If

End Function

