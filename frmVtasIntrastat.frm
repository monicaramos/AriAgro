VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmVtasIntrastat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6840
   Icon            =   "frmVtasIntrastat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
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
      Height          =   5130
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   6555
      Begin VB.CheckBox Check1 
         Caption         =   "Fecha de Factura"
         Height          =   285
         Left            =   3510
         TabIndex        =   23
         Top             =   2370
         Width           =   2475
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "Text5"
         Top             =   1740
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Text5"
         Top             =   1365
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1755
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1830
         MaxLength       =   6
         TabIndex        =   0
         Top             =   1365
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2775
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2415
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5025
         TabIndex        =   7
         Top             =   4500
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   6
         Top             =   4500
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1845
         MaxLength       =   3
         TabIndex        =   4
         Top             =   3360
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1845
         MaxLength       =   3
         TabIndex        =   5
         Top             =   3735
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text5"
         Top             =   3360
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text5"
         Top             =   3735
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Relación Intrastat"
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
         TabIndex        =   22
         Top             =   450
         Width           =   5160
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1530
         MouseIcon       =   "frmVtasIntrastat.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   1740
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1545
         MouseIcon       =   "frmVtasIntrastat.frx":015E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   1365
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
         Left            =   585
         TabIndex        =   21
         Top             =   1125
         Width           =   630
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   945
         TabIndex        =   20
         Top             =   1740
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   945
         TabIndex        =   19
         Top             =   1365
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha albarán"
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
         Left            =   600
         TabIndex        =   16
         Top             =   2115
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   960
         TabIndex        =   15
         Top             =   2415
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   960
         TabIndex        =   14
         Top             =   2775
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1530
         Picture         =   "frmVtasIntrastat.frx":02B0
         ToolTipText     =   "Buscar fecha"
         Top             =   2415
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1530
         Picture         =   "frmVtasIntrastat.frx":033B
         ToolTipText     =   "Buscar fecha"
         Top             =   2775
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   960
         TabIndex        =   13
         Top             =   3360
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   960
         TabIndex        =   12
         Top             =   3735
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "País"
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
         Left            =   600
         TabIndex        =   11
         Top             =   3120
         Width           =   285
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1560
         MouseIcon       =   "frmVtasIntrastat.frx":03C6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar país"
         Top             =   3360
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1575
         MouseIcon       =   "frmVtasIntrastat.frx":0518
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar país"
         Top             =   3735
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmVtasIntrastat"
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
Private WithEvents frmPais As frmManPaises 'Paises
Attribute frmPais.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1

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

Dim cadSelect1 As String


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Label4(16).Caption = "Fecha Factura"
    Else
        Label4(16).Caption = "Fecha Albarán"
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim cadSelect1 As String

InicializarVbles
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
     '======== FORMULA  ====================================
    'Seleccionar registros de la empresa conectada
'    Codigo = "{" & tabla & ".codempre}=" & vEmpresa.codEmpre
'    If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
'    If Not AnyadirAFormula(cadSelect, Codigo) Then Exit Sub
    
    
    'D/H Pais
    cDesde = Trim(txtCodigo(0).Text)
    cHasta = Trim(txtCodigo(1).Text)
    nDesde = txtNombre(0).Text
    nHasta = txtNombre(1).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{clientes.codpaise}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHPais= """) Then Exit Sub
    End If
    
    'D/H Variedad
    cDesde = Trim(txtCodigo(4).Text)
    cHasta = Trim(txtCodigo(5).Text)
    nDesde = txtNombre(4).Text
    nHasta = txtNombre(5).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{albaran_variedad.codvarie}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad= """) Then Exit Sub
    End If
    
    
    If Check1.Value = 0 Then
        'D/H Fecha albaran
        cDesde = Trim(txtCodigo(2).Text)
        cHasta = Trim(txtCodigo(3).Text)
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{albaran.fechaalb}"
            TipCod = "F"
            If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
        End If
        
        '[Monica]03/11/2017: para el caso de ser por fecha de albaran
        cadParam = cadParam & "pNumfac=0|"
        numParam = numParam + 1
        
        
    Else
        cDesde = Trim(txtCodigo(2).Text)
        cHasta = Trim(txtCodigo(3).Text)
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{facturas.fecfactu}"
            TipCod = "F"
            cadParam = cadParam & AnyadirParametroDH("pDHFecha", cDesde, cHasta, "", "")
            numParam = numParam + 1
            
            cadSelect1 = " (1=1) "
            If cDesde <> "" Then
                cadSelect1 = cadSelect1 & " and facturas_variedad.fecfactu >= '" & Format(cDesde, "yyyy-mm-dd") & "'"
            End If
            If cHasta <> "" Then
                cadSelect1 = cadSelect1 & " and facturas_variedad.fecfactu <= '" & Format(cHasta, "yyyy-mm-dd") & "'"
            End If
        End If
        
        '[Monica]03/11/2017: para el caso de ser por fecha de factura
        cadParam = cadParam & "pNumfac=1|"
        numParam = numParam + 1
        
        
    End If
    
    
    cadTABLA = tabla & " INNER JOIN albaran_variedad ON albaran.numalbar = albaran_variedad.numalbar "
    cadTABLA = "(" & cadTABLA & ") INNER JOIN clientes ON albaran.codclien = clientes.codclien "
    cadTABLA = "(" & cadTABLA & ") INNER JOIN paises ON clientes.codpaise = paises.codpaise "
    
    If Not AnyadirAFormula(cadFormula, "({paises.intracom} = 1 or {paises.intrastad} = 1)") Then Exit Sub
    If Not AnyadirAFormula(cadselect, "({paises.intracom} = 1 or {paises.intrastad} = 1)") Then Exit Sub
    
    If CargarTablaTemporal(cadTABLA, cadselect, cadSelect1) Then
        If HayRegParaInforme("tmpinformes", "{tmpinformes.codusu}=" & vUsu.Codigo) Then
              'Nombre fichero .rpt a Imprimir
              cadTitulo = "Informe Intrastat"
              cadNombreRPT = "rIntrastat.rpt"
              LlamarImprimir
              'AbrirVisReport
        End If
   End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(4)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
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
         
    FrameCobrosVisible True, H, W
    indFrame = 5
    tabla = "albaran"
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmPais_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Paises
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Variedades
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
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
        Case 0, 1 'PAISES
            AbrirFrmPaises (Index)
        
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
            Case 0: KEYBusqueda KeyAscii, 0 'pais desde
            Case 1: KEYBusqueda KeyAscii, 1 'pais hasta
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
            
        Case 0, 1 'PAIS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "paises", "nompaise", "codpaise", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
        
        Case 2, 3 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 4, 5 'VARIEDAD
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "00")
            
                        
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 5850
        Me.FrameCobros.Width = 6930
        W = Me.FrameCobros.Width
        H = Me.FrameCobros.Height
    End If
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadselect = ""
    cadParam = ""
    numParam = 0
    cadSelect1 = ""
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
        .FormulaSeleccion = "{tmpinformes.codusu} = " & vUsu.Codigo 'cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .Titulo = cadTitulo
        .EnvioEMail = False
        .NombreRPT = cadNombreRPT
        .Opcion = 0
        .Show vbModal
    End With
End Sub

Private Sub AbrirFrmPaises(indice As Integer)
    indCodigo = indice
    Set frmPais = New frmManPaises
    frmPais.DatosADevolverBusqueda = "0|2|"
    frmPais.Show vbModal
    Set frmPais = Nothing
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


Private Function CargarTablaTemporal(cadTABLA As String, cWhere As String, cWhere1 As String) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Importe As Currency
Dim Sql2 As String
Dim Facturas As String

    On Error GoTo eCargarTablaTemporal


    CargarTablaTemporal = False

    SQL = "delete from tmpinformes where codusu= " & vUsu.Codigo
    conn.Execute SQL
    
    
    cadTABLA = QuitarCaracterACadena(cadTABLA, "{")
    cadTABLA = QuitarCaracterACadena(cadTABLA, "}")
    SQL = "Select count(*) FROM " & QuitarCaracterACadena(cadTABLA, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    SQL = "select albaran_variedad.numalbar, albaran_variedad.numlinea from " & cadTABLA & " where " & cWhere
    
    If cWhere1 <> "" Then SQL = SQL & " and (albaran_variedad.numalbar, albaran_variedad.numlinea) in (select numalbar, numlinealbar from facturas_variedad where " & cWhere1 & ")"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        
        Sql2 = "select sum(impornet) from facturas_variedad where numalbar=" & DBLet(Rs.Fields(0).Value, "N")
        Sql2 = Sql2 & " and numlinealbar =" & DBLet(Rs.Fields(1).Value, "N")
        
        '[Monica]13/10/2014: falta la condicion de que esté entre fechas de factura y/o variedades
        If cWhere1 <> "" Then
            Sql2 = Sql2 & " and " & cWhere1
        End If
        
        Importe = 0
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs2.EOF Then Importe = DBLet(Rs2.Fields(0).Value, "N")
        Set Rs2 = Nothing
        
        Sql2 = "insert into tmpinformes(codusu,codigo1,campo1,importe1) values ("
        Sql2 = Sql2 & vUsu.Codigo & "," & DBSet(Rs.Fields(0), "N") & "," & DBSet(Rs.Fields(1), "N") & ","
        Sql2 = Sql2 & DBSet(Importe, "N") & ")"
        
        conn.Execute Sql2
        
        
        '[Monica]03/11/2017: para el caso de que sea por fecha de factura
        Sql2 = "select distinct concat(letraser,right(concat('000000',facturas_variedad.numfactu),7)) from facturas_variedad, usuarios.stipom  "
        Sql2 = Sql2 & " where numalbar = " & DBSet(Rs.Fields(0).Value, "N")
        Sql2 = Sql2 & " and numlinealbar =" & DBLet(Rs.Fields(1).Value, "N")
        Sql2 = Sql2 & " and facturas_variedad.codtipom = stipom.codtipom "
        If cWhere1 <> "" Then
            Sql2 = Sql2 & " and " & cWhere1
        End If
        Facturas = ""
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs2.EOF
            Facturas = " " & Facturas & DBLet(Rs2.Fields(0).Value, "N")
            
            Rs2.MoveNext
        Wend
        Set Rs2 = Nothing
        
        If Facturas <> "" Then Facturas = Trim(Facturas)
        
        Sql2 = "update tmpinformes set nombre1 = " & DBSet(Facturas, "T") & " where codusu = "
        Sql2 = Sql2 & vUsu.Codigo & " and codigo1 = " & DBSet(Rs.Fields(0), "N") & " and campo1 =" & DBSet(Rs.Fields(1), "N")
        
        conn.Execute Sql2
        
        
        Rs.MoveNext
    Wend
    
    CargarTablaTemporal = True
    Exit Function
    
eCargarTablaTemporal:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Error Cargando la tabla intermedia"
    End If
End Function

