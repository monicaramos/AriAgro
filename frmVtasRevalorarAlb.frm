VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmVtasRevalorarAlb 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6840
   Icon            =   "frmVtasRevalorarAlb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   6840
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
      Height          =   6060
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   6555
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   240
         Left            =   450
         TabIndex        =   22
         Top             =   5040
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "Text5"
         Top             =   3510
         Width           =   3495
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Text5"
         Top             =   3135
         Width           =   3495
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1650
         MaxLength       =   4
         TabIndex        =   3
         Top             =   3525
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   2
         Top             =   3135
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   4635
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   4275
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4995
         TabIndex        =   7
         Top             =   5385
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3810
         TabIndex        =   6
         Top             =   5385
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1650
         MaxLength       =   16
         TabIndex        =   0
         Top             =   2115
         Width           =   1635
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1650
         MaxLength       =   16
         TabIndex        =   1
         Top             =   2490
         Width           =   1635
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text5"
         Top             =   2115
         Width           =   2685
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text5"
         Top             =   2490
         Width           =   2685
      End
      Begin VB.Label Label2 
         Caption         =   "Este proceso es lento, se recomienda que no haya nadie en la aplicaci�n cuando se vaya a ejecutar para no ralentizarlo. "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1005
         Left            =   420
         TabIndex        =   24
         Top             =   840
         Width           =   5640
      End
      Begin VB.Label Label1 
         Caption         =   "Revalorar Albaranes"
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
         Left            =   420
         TabIndex        =   23
         Top             =   300
         Width           =   5160
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1335
         MouseIcon       =   "frmVtasRevalorarAlb.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   3510
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1350
         MouseIcon       =   "frmVtasRevalorarAlb.frx":015E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   3135
         Width           =   240
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
         Index           =   2
         Left            =   390
         TabIndex        =   21
         Top             =   2895
         Width           =   630
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   750
         TabIndex        =   20
         Top             =   3510
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   750
         TabIndex        =   19
         Top             =   3135
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Albar�n"
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
         Height          =   255
         Index           =   16
         Left            =   405
         TabIndex        =   16
         Top             =   3975
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   765
         TabIndex        =   15
         Top             =   4275
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   765
         TabIndex        =   14
         Top             =   4635
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1335
         Picture         =   "frmVtasRevalorarAlb.frx":02B0
         ToolTipText     =   "Buscar fecha"
         Top             =   4275
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1335
         Picture         =   "frmVtasRevalorarAlb.frx":033B
         ToolTipText     =   "Buscar fecha"
         Top             =   4635
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   765
         TabIndex        =   13
         Top             =   2115
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   765
         TabIndex        =   12
         Top             =   2490
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Forfait"
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
         Left            =   405
         TabIndex        =   11
         Top             =   1875
         Width           =   480
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1365
         MouseIcon       =   "frmVtasRevalorarAlb.frx":03C6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   2115
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1350
         MouseIcon       =   "frmVtasRevalorarAlb.frx":0518
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   2490
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmVtasRevalorarAlb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MONICA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar n� oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexi�n a BD Ariges  2.- Conexi�n a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmVar As frmManVariedad 'Variedad
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmFor As frmManForfaits 'Forfaits
Attribute frmFor.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'n� de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim Tabla As String
Dim codigo As String 'C�digo para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String

InicializarVbles
    
    '========= PARAMETROS  =============================
    'A�adir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
    numParam = numParam + 1
    
    cadParam = cadParam & "|pUsu=" & vUsu.codigo & "|"
    numParam = numParam + 1
    
    
     '======== FORMULA  ====================================
    'Seleccionar registros de la empresa conectada
'    Codigo = "{" & tabla & ".codempre}=" & vEmpresa.codEmpre
'    If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
'    If Not AnyadirAFormula(cadSelect, Codigo) Then Exit Sub
    
    
    'D/H Forfait
    cDesde = Trim(txtCodigo(0).Text)
    cHasta = Trim(txtCodigo(1).Text)
    nDesde = txtNombre(0).Text
    nHasta = txtNombre(1).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        codigo = "{" & Tabla & ".codforfait}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHForfait= """) Then Exit Sub
    End If
    
    'D/H Variedad
    cDesde = Trim(txtCodigo(4).Text)
    cHasta = Trim(txtCodigo(5).Text)
    nDesde = txtNombre(4).Text
    nHasta = txtNombre(5).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        codigo = "{albaran_variedad.codvarie}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad= """) Then Exit Sub
    End If
    
    'D/H Fecha albaran
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        codigo = "{albaran.fechaalb}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
   
    cadTABLA = Tabla & " INNER JOIN albaran ON albaran.numalbar = albaran_variedad.numalbar "
    
    If HayRegistros(cadTABLA, cadSelect) Then
        ProcesarCambios (cadSelect)
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(0)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
    Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Me.imgBuscar(1).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Me.imgBuscar(4).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Me.imgBuscar(5).Picture = frmPpal.imgListImages16.ListImages(1).Picture

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    Tabla = "albaran_variedad"
    
    Pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = w + 70
    Me.Height = h + 350
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Variedades
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Agencias de transporte
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
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
        Case 0, 1 'FORFAITS
            AbrirFrmForfait (Index)
        
        
        Case 4, 5 'VARIEDADES
            AbrirFrmVariedades (Index)
        
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
            Case 0: KEYBusqueda KeyAscii, 0 'forfait desde
            Case 1: KEYBusqueda KeyAscii, 1 'forfait hasta
            Case 4: KEYBusqueda KeyAscii, 4 'variedad desde
            Case 5: KEYBusqueda KeyAscii, 5 'variedad hasta
            Case 2: KEYFecha KeyAscii, 2 'fecha desde
            Case 3: KEYFecha KeyAscii, 3 'fecha hasta
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
            
        Case 0, 1 'FORFAIT
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "forfaits", "nomconfe", "codforfait", "T")
        
        Case 2, 3 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 4, 5 'VARIEDAD
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
            
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 6060
        Me.FrameCobros.Width = 6930
        w = Me.FrameCobros.Width
        h = Me.FrameCobros.Height
    End If
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
End Sub

Private Function PonerDesdeHasta(codD As String, codH As String, nomD As String, nomH As String, param As String) As Boolean
'IN: codD,codH --> codigo Desde/Hasta
'    nomD,nomH --> Descripcion Desde/Hasta
'A�ade a cadFormula y cadSelect la cadena de seleccion:
'       "(codigo>=codD AND codigo<=codH)"
' y a�ade a cadParam la cadena para mostrar en la cabecera informe:
'       "codigo: Desde codD-nomd Hasta: codH-nomH"
Dim devuelve As String
Dim devuelve2 As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(codD, codH, codigo, TipCod)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    If TipCod <> "F" Then 'Fecha
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        devuelve2 = CadenaDesdeHastaBD(codD, codH, codigo, TipCod)
        If devuelve2 = "Error" Then Exit Function
        If Not AnyadirAFormula(cadSelect, devuelve2) Then Exit Function
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
        .NombreRPT = cadNombreRPT
        .EnvioEMail = False
        .Opcion = 0
        .Show vbModal
    End With
End Sub

Private Sub AbrirFrmForfait(indice As Integer)
    indCodigo = indice
    Set frmFor = New frmManForfaits
    frmFor.DatosADevolverBusqueda = "0|1|"
    frmFor.Show vbModal
    Set frmFor = Nothing
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


Private Function HayRegistros(cTabla As String, cWhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
Dim RS As ADODB.Recordset

    SQL = "Select * FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If RS.EOF Then
        MsgBox "No hay datos para mostrar en el Informe.", vbInformation
        HayRegistros = False
    Else
        HayRegistros = True
    End If

End Function

Private Sub ProcesarCambios(cadWhere As String)
Dim SQL As String
Dim i As Integer
Dim HayReg As Integer
Dim b As Boolean
Dim RS As ADODB.Recordset
Dim Mens As String

On Error GoTo eProcesarCambios

    conn.BeginTrans
    
    
    If cadWhere <> "" Then
        cadWhere = QuitarCaracterACadena(cadWhere, "{")
        cadWhere = QuitarCaracterACadena(cadWhere, "}")
        cadWhere = QuitarCaracterACadena(cadWhere, "_1")
    End If
        
    SQL = "select count(*) from albaran_variedad inner join albaran on albaran_variedad.numalbar = albaran.numalbar " ' & cadWhere
    If cadWhere <> "" Then SQL = SQL & " where " & cadWhere
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Pb1.visible = True
    CargarProgres Pb1, RS.Fields(0).Value
    DoEvents
    RS.Close
    
    SQL = "select albaran_variedad.* from albaran_variedad  inner join albaran on albaran_variedad.numalbar = albaran.numalbar " '& cadWhere
    If cadWhere <> "" Then SQL = SQL & " where " & cadWhere
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    b = True
    While Not RS.EOF And b
        IncrementarProgres Pb1, 1
        DoEvents
        Mens = "Actualizando Costes" & vbCrLf & vbCrLf & "Albaran: " & RS!NumAlbar & vbCrLf & "Linea: " & RS!NumLinea
        b = ActualizarCostes(DBLet(RS!NumAlbar, "N"), DBLet(RS!NumLinea, "N"), True, DBLet(RS!codforfait, "T"), DBLet(RS!CodPalet, "N"))
        RS.MoveNext
    Wend
    Set RS = Nothing
eProcesarCambios:
    If Err.Number <> 0 Then
        b = False
    End If
    If b Then
        conn.CommitTrans
        MsgBox "Proceso realizado correctamente.", vbExclamation
        cmdCancel_Click
    Else
        conn.RollbackTrans
        MsgBox "Error " & Mens, vbExclamation
    End If
End Sub


Private Sub InsertaLineaEnTemporal(ByRef ItmX As ListItem)
Dim SQL As String
Dim Codmacta As String
Dim RS As ADODB.Recordset
Dim SQL1 As String

        SQL1 = "insert into tmpinformes(codusu, codigo1) values ("
        SQL1 = SQL1 & DBSet(vUsu.codigo, "N") & "," & DBSet(ItmX.Text, "N") & ")"

        conn.Execute SQL1
    
End Sub

