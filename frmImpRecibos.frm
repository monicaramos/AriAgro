VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImpRecibos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6570
   Icon            =   "frmImpRecibos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameHorasTrabajadas 
      Height          =   5820
      Left            =   45
      TabIndex        =   8
      Top             =   0
      Width           =   6435
      Begin VB.CheckBox Check1 
         Caption         =   "Sobre Horas Productivas"
         Height          =   195
         Index           =   1
         Left            =   780
         TabIndex        =   21
         Top             =   4860
         Width           =   2130
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Imprimir Resumen Recibo"
         Height          =   195
         Index           =   0
         Left            =   765
         TabIndex        =   20
         Top             =   4500
         Width           =   2130
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   1530
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "Tipo|N|N|||straba|codsecci||N|"
         Top             =   3870
         Width           =   1350
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   20
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   4
         Top             =   3420
         Width           =   1005
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4605
         TabIndex        =   7
         Top             =   5115
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   3525
         TabIndex        =   6
         Top             =   5100
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   19
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1665
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   0
         Top             =   1305
         Width           =   750
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   18
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text5"
         Top             =   1305
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   19
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text5"
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   3
         Top             =   2745
         Width           =   1005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   2
         Top             =   2340
         Width           =   1005
      End
      Begin VB.Label Label8 
         Caption         =   "Secci�n "
         ForeColor       =   &H00972E0B&
         Height          =   255
         Left            =   585
         TabIndex        =   19
         Top             =   3870
         Width           =   615
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   6
         Left            =   1560
         Picture         =   "frmImpRecibos.frx":000C
         Top             =   3420
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Recibo"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   30
         Left            =   585
         TabIndex        =   18
         Top             =   3240
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   29
         Left            =   960
         TabIndex        =   17
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   28
         Left            =   960
         TabIndex        =   16
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   27
         Left            =   600
         TabIndex        =   15
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label Label7 
         Caption         =   "Impresi�n de Recibos"
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
         Left            =   405
         TabIndex        =   14
         Top             =   405
         Width           =   5925
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   26
         Left            =   960
         TabIndex        =   13
         Top             =   2400
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   25
         Left            =   960
         TabIndex        =   12
         Top             =   2715
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   24
         Left            =   600
         TabIndex        =   11
         Top             =   2160
         Width           =   450
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   15
         Left            =   1620
         MouseIcon       =   "frmImpRecibos.frx":0097
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   1665
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   14
         Left            =   1620
         MouseIcon       =   "frmImpRecibos.frx":01E9
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   1305
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   1575
         Picture         =   "frmImpRecibos.frx":033B
         Top             =   2745
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   1575
         Picture         =   "frmImpRecibos.frx":03C6
         Top             =   2340
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmImpRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: LAURA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public OpcionListado As Byte
    '==== Listados BASICOS ====
    '=============================
    ' 10 .- Listado de Clientes
    ' 11 .- Listado de Proveedores
    ' 12 .- Listado de Variedades
    ' 13 .- Listado de Calibres
    ' 15 .- Listado de Horas trababajadas
    
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar n� oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public Event RectificarFactura(cliente As String, observaciones As String)

Private Conexion As Byte
'1.- Conexi�n a BD Ariges  2.- Conexi�n a BD Conta

Private HaDevueltoDatos As Boolean


Private WithEvents frmPro As frmManProve 'Proveedores
Attribute frmPro.VB_VarHelpID = -1
Private WithEvents frmCli As frmClientes 'Clientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmProd As frmManProductos 'Productos
Attribute frmProd.VB_VarHelpID = -1
Private WithEvents frmVar As frmManVariedad 'Variedades
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmCal As frmManCalibres 'Calibres
Attribute frmCal.VB_VarHelpID = -1
Private WithEvents frmTEn As frmManTipEnv 'tipos de envases
Attribute frmTEn.VB_VarHelpID = -1
Private WithEvents frmTra As frmManTraba 'mantenimiento de trabajadores
Attribute frmTra.VB_VarHelpID = -1

Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
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
Dim indFrame As Single 'n� de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'C�digo para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim tipo As String


Dim PrimeraVez As Boolean
Dim Contabilizada As Byte

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub
Private Sub cmdAceptar_Click(Index As Integer)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
    
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
    
    
    If Not DatosOk Then Exit Sub
    
    InicializarVbles
    
    'A�adir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
    numParam = numParam + 1
    
    Select Case Index
        Case 0 ' impresion de recibos
            '======== FORMULA  ====================================
            'D/H TRABAJADOR
            cDesde = Trim(txtCodigo(18).Text)
            cHasta = Trim(txtCodigo(19).Text)
            nDesde = txtNombre(18).Text
            nHasta = txtNombre(19).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{horas.codtraba}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHTrabajador=""") Then Exit Sub
            End If
            
            'D/H fecha
            cDesde = Trim(txtCodigo(16).Text)
            cHasta = Trim(txtCodigo(17).Text)
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{horas.fechahora}"
                TipCod = "F"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
            End If
                       
            'Tipo de seccion
            AnyadirAFormula cadFormula, "{straba.codsecci} = " & Me.Combo1(1).ListIndex
            AnyadirAFormula cadselect, "{straba.codsecci} = " & Me.Combo1(1).ListIndex
            
            
            tabla = "horas INNER JOIN straba ON horas.codtraba = straba.codtraba "
                       
            cadParam = cadParam & "pFecha=""" & txtCodigo(20).Text & """|"
            numParam = numParam + 1
            cadParam = cadParam & "pTitulo=""" & "Recibo Horas " & Combo1(1).Text & """|"
            numParam = numParam + 1
                       
            ' imprimir el resumen
            If Check1(0).Value Then ' se imprime el resumen
                cadParam = cadParam & "pImpIRPF=1|"
            Else
                cadParam = cadParam & "pImpIRPF=0|"
            End If
            numParam = numParam + 1
                       
            ' sobre horas productivas
            If Check1(1).Value Then ' se imprime el resumen
                cadParam = cadParam & "pHProductivas=1|"
            Else
                cadParam = cadParam & "pHProductivas=0|"
            End If
            numParam = numParam + 1
                       
                       
                       
            indRPT = 13 'Impresion de Recibos
            If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
            
            'Nombre fichero .rpt a Imprimir
            frmImprimir.NombreRPT = nomDocu
            
            AnyadirAFormula cadFormula, "isnull({horas.fecharec})"
            AnyadirAFormula cadselect, "isnull({horas.fecharec})"
 
            cadTitulo = "Impresi�n de Recibos"
    
    End Select
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(tabla, cadselect) Then
        LlamarImprimir
        If MsgBox("�Impresi�n correcta para actualizar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            If ActualizarRegistros(tabla, cadselect) Then
               MsgBox "Proceso realizado correctamente", vbExclamation
            End If
        End If
    End If

    cmdCancel_Click
    
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 10 ' Listado de Clientes
                PonerFoco txtCodigo(4)
                
            Case 11 ' Listado de Proveedores
                PonerFoco txtCodigo(2)
            
            Case 12 ' Listado de Variedades
                PonerFoco txtCodigo(6)
        
            Case 13 ' Listado de Calibres
                PonerFoco txtCodigo(8)
                
            Case 14 ' Imforme de Movimientos de calibres
                PonerFoco txtCodigo(12)
            
            Case 15 ' Informe de Horas Trabajadas
                PonerFoco txtCodigo(18)
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
    Set List = New Collection
    For H = 24 To 27
        List.Add H
    Next H
    For H = 1 To 10
        List.Add H
    Next H
    List.Add 12
    List.Add 13
    List.Add 14
    List.Add 15
    List.Add 18
    List.Add 19
    
    
'    For h = 1 To List.Count
'        Me.imgBuscar(List.item(h)).Picture = frmPpal.imgListImages16.ListImages(1).Picture
'    Next h
' ### [Monica] 09/11/2006    he sustituido el anterior
    For H = 14 To 15 'imgBuscar.Count - 1
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
     
    
    Set List = Nothing

    'Ocultar todos los Frames de Formulario
    Me.FrameHorasTrabajadas.visible = False
    
    '###Descomentar
'    CommitConexion
    H = 6360
    W = 6660
    FrameHorasTrabajadasVisible True, H, W
    indFrame = 0
    tabla = "horas"
        
    CargaCombo
        
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
    
    Me.Combo1(1).ListIndex = 1
End Sub



Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtCodigo(CByte(imgFecha(2).Tag) + 14).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmCal_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 14, 15 'Horas trabajadas
            AbrirFrmManTraba (Index)
    
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub


Private Sub imgFecha_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim obj As Object

    Set frmC = New frmCal
    
    esq = imgFecha(Index).Left
    dalt = imgFecha(Index).Top
        
    Set obj = imgFecha(Index).Container
      
      While imgFecha(Index).Parent.Name <> obj.Name
            esq = esq + obj.Left
            dalt = dalt + obj.Top
            Set obj = obj.Container
      Wend
    
    menu = Me.Height - Me.ScaleHeight 'ac� tinc el heigth del men� i de la toolbar

    frmC.Left = esq + imgFecha(Index).Parent.Left + 30
    frmC.Top = dalt + imgFecha(Index).Parent.Top + imgFecha(Index).Height + menu - 40

    imgFecha(2).Tag = Index '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtCodigo(Index + 14).Text <> "" Then frmC.NovaData = txtCodigo(Index + 14).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtCodigo(CByte(imgFecha(2).Tag) + 14) '<===
    ' ********************************************
End Sub



Private Sub ListView1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'15/02/2007
'    KEYpress KeyAscii
'ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 18: KEYBusqueda KeyAscii, 14 'trabajador desde
            Case 19: KEYBusqueda KeyAscii, 15 'trabajador hasta
            Case 2: KEYFecha KeyAscii, 16 'fecha desde
            Case 3: KEYFecha KeyAscii, 17 'fecha hasta
            Case 6: KEYFecha KeyAscii, 20 'fecha recibo

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
    imgFecha_Click (indice)
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
            
        Case 14, 15, 16, 17, 20 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 18, 19 'TRABAJADORES
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "straba", "nomtraba", "codtraba", "N")
            
    End Select
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim tabla As String
Dim Titulo As String

    'Llamamos a al form
    cad = ""
    Conexion = cAgro    'Conexi�n a BD: Ariges
'    Select Case OpcionListado
'        Case 7 'Traspaso de Almacenes
'            cad = cad & "N� Trasp|scatra|codtrasp|N|0000000|40�Almacen Origen|scatra|almaorig|N|000|20�Almacen Destino|scatra|almadest|N|000|20�Fecha|scatra|fechatra|F||20�"
'            Tabla = "scatra"
'            titulo = "Traspaso Almacenes"
'        Case 8 'Movimientos de Almacen
'            cad = cad & "N� Movim.|scamov|codmovim|N|0000000|40�Almacen|scamov|codalmac|N|000|30�Fecha|scamov|fecmovim|F||30�"
'            Tabla = "scamov"
'            titulo = "Movimientos Almacen"
'        Case 9, 12, 13, 14, 15, 16, 17 '9: Movimientos Articulos
'                   '12: Inventario Articulos
'                   '14:Actualizar Diferencias de Stock Inventariado
'                   '16: Listado Valoracion stock inventariado
'            cad = cad & "C�digo|sartic|codartic|T||30�Denominacion|sartic|nomartic|T||70�"
'            Tabla = "sartic"
'            titulo = "Articulos"
'    End Select
          
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vtabla = tabla
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


Private Sub FrameHorasTrabajadasVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de horas trabajadas
    Me.FrameHorasTrabajadas.visible = visible
    If visible = True Then
        Me.FrameHorasTrabajadas.Top = -90
        Me.FrameHorasTrabajadas.Left = 0
        Me.FrameHorasTrabajadas.Height = 5985
        Me.FrameHorasTrabajadas.Width = 6660
        W = Me.FrameHorasTrabajadas.Width
        H = Me.FrameHorasTrabajadas.Height
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
'A�ade a cadFormula y cadSelect la cadena de seleccion:
'       "(codigo>=codD AND codigo<=codH)"
' y a�ade a cadParam la cadena para mostrar en la cabecera informe:
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
'        .NombreRPT = cadNombreRPT
        .EnvioEMail = False
        .Opcion = OpcionListado
        .Show vbModal
    End With
End Sub

Private Function PonerGrupo(numGrupo As Byte, cadgrupo As String) As Byte
Dim campo As String
Dim nomcampo As String

    campo = "pGroup" & numGrupo & "="
    nomcampo = "pGroup" & numGrupo & "Name="
    PonerGrupo = 0

    Select Case cadgrupo
'        Case "Codigo"
'            cadParam = cadParam & campo & "{" & Tabla & ".codclien}" & "|"
'            cadParam = cadParam & nomcampo & " {" & "scoope" & ".nomcoope}" & "|"
'            cadParam = cadParam & "pTitulo1" & "=""C�digo""" & "|"
'            numParam = numParam + 3
'
'        Case "Alfabetico"
'            cadParam = cadParam & campo & "{" & Tabla & ".tipsocio}" & "|"
'            cadParam = cadParam & nomcampo & " {" & "tiposoci" & ".nomtipso}" & "|"
'            cadParam = cadParam & "pTitulo1" & "=""Colectivo""" & "|"
'            numParam = numParam + 3
'
        
        'Informe de variedades
        Case "Clase"
            cadParam = cadParam & campo & "{" & tabla & ".codclase}" & "|"
            cadParam = cadParam & nomcampo & " {" & "clases" & ".nomclase}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Producto""" & "|"
            numParam = numParam + 3
            
        Case "Producto"
            cadParam = cadParam & campo & "{" & tabla & ".codprodu}" & "|"
            cadParam = cadParam & nomcampo & " {" & "productos" & ".nomprodu}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Clase""" & "|"
            numParam = numParam + 3

        'Informe de calibres
        Case "Variedad"
            cadParam = cadParam & campo & "{" & tabla & ".codvarie}" & "|"
            cadParam = cadParam & nomcampo & " {" & "variedades" & ".nomvarie}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Variedad""" & "|"
            numParam = numParam + 3
            
        Case "Calibre"
            cadParam = cadParam & campo & "{" & tabla & ".codcalib}" & "|"
            cadParam = cadParam & nomcampo & " {" & "calibres" & ".nomcalib}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Calibre""" & "|"
            numParam = numParam + 3
            
'        'Informe de Horas Trabajadas
'        Case "Trabajador"
'            cadParam = cadParam & campo & "{" & Tabla & ".codtraba}" & "|"
'            cadParam = cadParam & nomcampo & " {" & "straba" & ".nomtraba}" & "|"
'            cadParam = cadParam & "pTitulo1" & "=""Fecha""" & "|"
'            numParam = numParam + 3
'
'        Case "Fecha"
'            cadParam = cadParam & "pGroup1=" & "{" & Tabla & ".fechahora}" & "|"
'            cadParam = cadParam & "pGroup1Name=" & " {" & "horas" & ".fechahora}" & "|"
'            cadParam = cadParam & "pTitulo1" & "=""Trabajadores""" & "|"
'            numParam = numParam + 3
        

End Select

End Function

Private Function PonerOrden(cadgrupo As String) As Byte
Dim campo As String
Dim nomcampo As String

    PonerOrden = 0

    Select Case cadgrupo
        Case "Codigo"
            cadParam = cadParam & "Orden" & "= {" & tabla
            Select Case OpcionListado
                Case 10
                    cadParam = cadParam & ".codclien}|"
                Case 11
                    cadParam = cadParam & ".codprove}|"
            End Select
            tipo = "C�digo"
        Case "Alfab�tico"
            cadParam = cadParam & "Orden" & "= {" & tabla
            Select Case OpcionListado
                Case 10
                    cadParam = cadParam & ".nomclien}|"
                Case 11
                    cadParam = cadParam & ".nomprove}|"
            End Select
            tipo = "Alfab�tico"
    End Select
    
    numParam = numParam + 1

End Function

Private Sub AbrirFrmManTraba(indice As Integer)
    indCodigo = indice + 4
    Set frmTra = New frmManTraba
    frmTra.DatosADevolverBusqueda = "0|2|"
'    frmCli.DeConsulta = True
'    frmCli.CodigoActual = txtCodigo(indCodigo)
    frmTra.Show vbModal
    Set frmTra = Nothing
End Sub

Private Sub AbrirVisReport()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = cadFormula
'        .SoloImprimir = (Me.OptVisualizar(indFrame).Value = 1)
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
        .Opcion = OpcionListado
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

Private Sub PonerValoresFactura()
Dim intconta As String
Dim cad As String
'    txtCodigo(9).Text = RecuperaValor(CadTag, 1)
'    txtCodigo(10).Text = RecuperaValor(CadTag, 2)
'    txtCodigo(11).Text = RecuperaValor(CadTag, 3)
'    txtCodigo(12).Text = RecuperaValor(CadTag, 4)
'    txtNombre(9).Text = RecuperaValor(CadTag, 5)
'    Contabilizada = RecuperaValor(CadTag, 6)
     intconta = "intconta"
     txtCodigo(12).Text = ""
     txtCodigo(12).Text = DevuelveDesdeBDNew(cAgro, "schfac", "codsocio", "letraser", txtCodigo(9).Text, "T", intconta, "numfactu", txtCodigo(10).Text, "N", "fecfactu", txtCodigo(11).Text, "F")
     If txtCodigo(12).Text <> "" Then
        txtNombre(9).Text = PonerNombreDeCod(txtCodigo(12), "ssocio", "nomsocio", "codsocio", "N")
        Contabilizada = CInt(intconta)
     Else
        cad = "No existe la factura. Reintroduzca. " & vbCrLf & vbCrLf
        cad = cad & "   Serie: " & txtCodigo(9).Text & vbCrLf
        cad = cad & "   Factura: " & txtCodigo(10).Text & vbCrLf
        cad = cad & "   Fecha: " & txtCodigo(11).Text & vbCrLf
        cad = cad & vbCrLf
        MsgBox cad, vbExclamation
        txtCodigo(9).Text = ""
        txtCodigo(10).Text = ""
        txtCodigo(11).Text = ""
        PonerFoco txtCodigo(9)
     End If
End Sub


Private Function ConTarjetaProfesional(letraser As String, numfactu As String, fecfactu As String) As Boolean
Dim sql As String
Dim Rs As ADODB.Recordset

    sql = "select count(*) from slhfac, starje where letraser = " & DBSet(letraser, "T") & " and numfactu = " & DBSet(numfactu, "N")
    sql = sql & " and fecfactu = " & DBSet(fecfactu, "F") & " and starje.tiptarje = 2 and slhfac.numtarje = starje.numtarje "
    
    ConTarjetaProfesional = (TotalRegistros(sql) <> 0)

End Function


' ********* si n'hi han combos a la cap�alera ************
Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim I As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
'    For I = 0 To Combo1.Count - 1
'        Combo1(I).Clear
'    Next I

    Combo1(1).Clear
    
    Combo1(1).AddItem "Campo"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Almac�n"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    
    
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim sql As String
'Dim Datos As String

    On Error GoTo EDatosOK

    b = True

    If txtCodigo(20).Text = "" Then
        MsgBox "Debe introducir una Fecha de Recibo.", vbExclamation
        txtCodigo(20).Text = ""
        PonerFoco txtCodigo(20)
        b = False
    End If
    
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function ActualizarRegistros(tabla As String, cWhere As String) As Boolean
Dim sql As String
    On Error GoTo eActualizarRegistros
    
    ActualizarRegistros = False
    
    cWhere = QuitarCaracterACadena(cWhere, "{")
    cWhere = QuitarCaracterACadena(cWhere, "}")
    cWhere = QuitarCaracterACadena(cWhere, "_1")

    sql = "update horas, straba set fecharec = " & DBSet(txtCodigo(20).Text, "F")
    sql = sql & " where " & cWhere
    sql = sql & " and horas.codtraba = straba.codtraba"
'    (codtraba, fechahora) in (select horas.codtraba, horas.fechahora from " & tabla & " where " & cWhere & ")"
    
    conn.Execute sql
        
    ActualizarRegistros = True
    
    Exit Function

eActualizarRegistros:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Error en la actualizacion de Registros" & vbCrLf & Err.Description
    End If
End Function

