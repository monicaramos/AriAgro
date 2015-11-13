VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmVtasDiarioFact 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6645
   Icon            =   "frmVtasDiarioFact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   6645
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
      Height          =   5850
      Left            =   45
      TabIndex        =   11
      Top             =   0
      Width           =   6555
      Begin VB.Frame Frame1 
         Caption         =   "Orden"
         ForeColor       =   &H00972E0B&
         Height          =   1320
         Left            =   3825
         TabIndex        =   27
         Top             =   990
         Width           =   1995
         Begin VB.OptionButton Option1 
            Caption         =   "Serie"
            Height          =   195
            Index           =   1
            Left            =   630
            TabIndex        =   29
            Top             =   405
            Width           =   690
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Cliente"
            Height          =   195
            Index           =   0
            Left            =   630
            TabIndex        =   28
            Top             =   855
            Width           =   1095
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Detallar Albaranes"
         Height          =   195
         Left            =   495
         TabIndex        =   8
         Top             =   5130
         Width           =   1815
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1755
         MaxLength       =   1
         TabIndex        =   1
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   1485
         Width           =   510
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1755
         MaxLength       =   1
         TabIndex        =   0
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   1125
         Width           =   510
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1755
         MaxLength       =   7
         TabIndex        =   3
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   2340
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1755
         MaxLength       =   7
         TabIndex        =   2
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   1995
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3405
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3060
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5040
         TabIndex        =   10
         Top             =   5175
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3870
         TabIndex        =   9
         Top             =   5175
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   6
         Top             =   4035
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   7
         Top             =   4410
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text5"
         Top             =   4035
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   4410
         Width           =   3135
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   5
         Left            =   900
         TabIndex        =   26
         Top             =   1170
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   4
         Left            =   900
         TabIndex        =   25
         Top             =   1485
         Width           =   420
      End
      Begin VB.Label Label1 
         Caption         =   "Diario de Facturación"
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
         TabIndex        =   24
         Top             =   315
         Width           =   5160
      End
      Begin VB.Label Label4 
         Caption         =   "Serie"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   3
         Left            =   495
         TabIndex        =   23
         Top             =   990
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   900
         TabIndex        =   22
         Top             =   2340
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   900
         TabIndex        =   21
         Top             =   2025
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nro.Factura"
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
         Left            =   495
         TabIndex        =   20
         Top             =   1755
         Width           =   870
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   465
         TabIndex        =   19
         Top             =   2745
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   825
         TabIndex        =   18
         Top             =   3045
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   825
         TabIndex        =   17
         Top             =   3405
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1485
         Picture         =   "frmVtasDiarioFact.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   3045
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1485
         Picture         =   "frmVtasDiarioFact.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   3405
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   855
         TabIndex        =   16
         Top             =   4035
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   870
         TabIndex        =   15
         Top             =   4410
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Left            =   510
         TabIndex        =   14
         Top             =   3795
         Width           =   495
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1470
         MouseIcon       =   "frmVtasDiarioFact.frx":0122
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   4035
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1485
         MouseIcon       =   "frmVtasDiarioFact.frx":0274
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   4410
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmVtasDiarioFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MONICA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmCli As frmClientes 'Clientes
Attribute frmCli.VB_VarHelpID = -1
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
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
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
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
    numParam = numParam + 1
    
     '======== FORMULA  ====================================
    'Seleccionar registros de la empresa conectada
'    Codigo = "{" & tabla & ".codempre}=" & vEmpresa.codEmpre
'    If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
'    If Not AnyadirAFormula(cadSelect, Codigo) Then Exit Sub
    
    'D/H Cliente
    cDesde = Trim(txtCodigo(0).Text)
    cHasta = Trim(txtCodigo(1).Text)
    nDesde = txtNombre(0).Text
    nHasta = txtNombre(1).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".codclien}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCliente= """) Then Exit Sub
    End If
    
    'D/H Serie
    cDesde = Trim(txtCodigo(6).Text)
    cHasta = Trim(txtCodigo(7).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{stipom.letraser}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSerie= """) Then Exit Sub
    End If
    
    'D/H Nro de Factura
    cDesde = Trim(txtCodigo(4).Text)
    cHasta = Trim(txtCodigo(5).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".numfactu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFactura= """) Then Exit Sub
    End If
    
    'D/H Fecha factura
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    'Orden del informe
    If Option1(0).Value Then
        cadParam = cadParam & "pOrden=1|"
        cadParam = cadParam & "pGrupo={facturas.codclien}|"
    Else
        cadParam = cadParam & "pOrden=0|"
        cadParam = cadParam & "pGrupo={stipom.letraser}|"
    End If
    numParam = numParam + 2
    
    'Detalle
    cadParam = cadParam & "pDetalle= " & Check1.Value & "|"
    numParam = numParam + 1
    
    cadTABLA = tabla & " INNER JOIN usuarios.stipom stipom ON facturas.codtipom = stipom.codtipom "
    
    If HayRegistros(cadTABLA, cadSelect) Then
        cadNombreRPT = "rDiarioFacturacion.rpt"
        cadTitulo = "Diario de Facturación"
        LlamarImprimir
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
Dim h As Integer, w As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
     For h = 0 To 1
        Me.imgBuscar(h).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Next h

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    tabla = "facturas"
    
    Option1(1).Value = True
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = w + 70
    Me.Height = h + 350
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
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
        Case 0, 1 'CLIENTE
            AbrirFrmClientes (Index)
        
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
            Case 0: KEYBusqueda KeyAscii, 0 'cliente desde
            Case 1: KEYBusqueda KeyAscii, 1 'cliente hasta
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
            
        Case 0, 1 'CLIENTE
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "clientes", "nomclien", "codclien", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
        
        Case 2, 3 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 4, 5 'FACTURAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoEntero txtCodigo(Index)
        
        Case 6, 7 'SERIE
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = UCase(txtCodigo(Index).Text)
    
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 6390
        Me.FrameCobros.Width = 6735
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
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        devuelve2 = CadenaDesdeHastaBD(codD, codH, Codigo, TipCod)
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
        .EnvioEMail = False
        .NombreRPT = cadNombreRPT
        .ConSubInforme = True
        .Opcion = 0
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
Dim sql As String
Dim Rs As ADODB.Recordset

    sql = "Select * FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        sql = sql & " WHERE " & cWhere
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Rs.EOF Then
        MsgBox "No hay datos para mostrar en el Informe.", vbInformation
        HayRegistros = False
    Else
        HayRegistros = True
    End If

End Function

Private Function ProcesarCambios(cadWhere As String) As Boolean
Dim sql As String
Dim Sql1 As String
Dim i As Integer
Dim HayReg As Integer
Dim b As Boolean

On Error GoTo eProcesarCambios

    HayReg = 0
    
    conn.Execute "delete from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
        
    If cadWhere <> "" Then
        cadWhere = QuitarCaracterACadena(cadWhere, "{")
        cadWhere = QuitarCaracterACadena(cadWhere, "}")
        cadWhere = QuitarCaracterACadena(cadWhere, "_1")
    End If
        
    sql = "insert into tmpinformes (codusu, codigo1) select " & DBSet(vUsu.Codigo, "N")
    sql = sql & ", albaran.numalbar from albaran, albaran_variedad where albaran.numalbar not in (select numalbar from tcafpa) "
    sql = sql & " and albaran.numalbar = albaran_variedad.numalbar "
    
    If cadWhere <> "" Then sql = sql & " and " & cadWhere
    
    
    conn.Execute sql
        
    ProcesarCambios = HayRegistros("tmpinformes", "codusu = " & vUsu.Codigo)

eProcesarCambios:
    If Err.Number <> 0 Then
        ProcesarCambios = False
    End If
End Function


Private Sub InsertaLineaEnTemporal(ByRef ItmX As ListItem)
Dim sql As String
Dim Codmacta As String
Dim Rs As ADODB.Recordset
Dim Sql1 As String

        Sql1 = "insert into tmpinformes(codusu, codigo1) values ("
        Sql1 = Sql1 & DBSet(vUsu.Codigo, "N") & "," & DBSet(ItmX.Text, "N") & ")"

        conn.Execute Sql1
    
End Sub

Private Function DatosOk() As Boolean

    DatosOk = True


End Function
