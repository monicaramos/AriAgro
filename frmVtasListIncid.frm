VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmVtasListIncid 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   10365
   Icon            =   "frmVtasListIncid.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   10365
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
      Height          =   6300
      Left            =   45
      TabIndex        =   12
      Top             =   0
      Width           =   9885
      Begin VB.Frame FrameCategoria 
         BorderStyle     =   0  'None
         Height          =   1305
         Left            =   5130
         TabIndex        =   32
         Top             =   810
         Width           =   5565
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   6
            Left            =   1380
            MaxLength       =   3
            TabIndex        =   2
            Top             =   585
            Width           =   830
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   7
            Left            =   1380
            MaxLength       =   3
            TabIndex        =   3
            Top             =   960
            Width           =   830
         End
         Begin VB.Label Label4 
            Caption         =   "Desde"
            Height          =   195
            Index           =   5
            Left            =   450
            TabIndex        =   35
            Top             =   585
            Width           =   465
         End
         Begin VB.Label Label4 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   4
            Left            =   450
            TabIndex        =   34
            Top             =   960
            Width           =   420
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Categoria"
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
            Index           =   3
            Left            =   120
            TabIndex        =   33
            Top             =   270
            Width           =   705
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "Text5"
         Top             =   2775
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "Text5"
         Top             =   2430
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   5
         Top             =   2775
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   4
         Top             =   2430
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Text5"
         Top             =   1755
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Text5"
         Top             =   1380
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   1
         Top             =   1755
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   0
         Top             =   1380
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Text5"
         Top             =   3840
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "Text5"
         Top             =   3495
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   7
         Top             =   3855
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   6
         Top             =   3495
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   9
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   4905
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   4545
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5055
         TabIndex        =   11
         Top             =   5535
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3870
         TabIndex        =   10
         Top             =   5535
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Cargando Tabla Temporal"
         Height          =   195
         Index           =   6
         Left            =   540
         TabIndex        =   36
         Top             =   5370
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1485
         MouseIcon       =   "frmVtasListIncid.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   2775
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1485
         MouseIcon       =   "frmVtasListIncid.frx":015E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   2430
         Width           =   240
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
         Left            =   540
         TabIndex        =   31
         Top             =   2190
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   870
         TabIndex        =   30
         Top             =   2775
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   870
         TabIndex        =   29
         Top             =   2430
         Width           =   465
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   13
         Left            =   1485
         MouseIcon       =   "frmVtasListIncid.frx":02B0
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar incidencia"
         Top             =   1770
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   12
         Left            =   1485
         MouseIcon       =   "frmVtasListIncid.frx":0402
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar incidencia"
         Top             =   1380
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Incidencia"
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
         Index           =   21
         Left            =   540
         TabIndex        =   28
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   22
         Left            =   870
         TabIndex        =   27
         Top             =   1755
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   23
         Left            =   870
         TabIndex        =   26
         Top             =   1380
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Informe de Incidencias o Categorias"
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
         TabIndex        =   21
         Top             =   450
         Width           =   5655
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1485
         MouseIcon       =   "frmVtasListIncid.frx":0554
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   3840
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1485
         MouseIcon       =   "frmVtasListIncid.frx":06A6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   3495
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
         Left            =   510
         TabIndex        =   20
         Top             =   3255
         Width           =   630
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   870
         TabIndex        =   19
         Top             =   3840
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   870
         TabIndex        =   18
         Top             =   3495
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   510
         TabIndex        =   15
         Top             =   4245
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   870
         TabIndex        =   14
         Top             =   4545
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   870
         TabIndex        =   13
         Top             =   4905
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1485
         Picture         =   "frmVtasListIncid.frx":07F8
         ToolTipText     =   "Buscar fecha"
         Top             =   4545
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1485
         Picture         =   "frmVtasListIncid.frx":0883
         ToolTipText     =   "Buscar fecha"
         Top             =   4905
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmVtasListIncid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MANOLO +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public OpcionListado As Integer
' 0 = listado de incidencias
' 1 = listado de categorias

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmVar As frmManVariedad 'Variedad
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmCli As frmClientes 'Clientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmInc As frmManInciden 'Incidencias
Attribute frmInc.VB_VarHelpID = -1
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

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim I As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String

InicializarVbles
    
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
    numParam = numParam + 1
    
     '======== FORMULA  ====================================
    'Seleccionar registros de la empresa conectada
'    Codigo = "{" & tabla & ".codempre}=" & vEmpresa.codEmpre
'    If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
'    If Not AnyadirAFormula(cadSelect, Codigo) Then Exit Sub
    
    
    If OpcionListado = 0 Then
        'D/H Incidencia
        cDesde = Trim(txtCodigo(12).Text)
        cHasta = Trim(txtCodigo(13).Text)
        nDesde = txtNombre(12).Text
        nHasta = txtNombre(13).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{albaran_variedad.codincid}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHIncidencia= """) Then Exit Sub
        End If
    Else
        'D/H Categoria
        cDesde = Trim(txtCodigo(6).Text)
        cHasta = Trim(txtCodigo(7).Text)
        nDesde = ""
        nHasta = ""
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{albaran_variedad.categori}"
            TipCod = "T"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCategoria= """) Then Exit Sub
        End If
    End If
    
    'D/H Cliente
    cDesde = Trim(txtCodigo(0).Text)
    cHasta = Trim(txtCodigo(1).Text)
    nDesde = txtNombre(0).Text
    nHasta = txtNombre(1).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{albaran.codclien}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCliente= """) Then Exit Sub
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
    
    'D/H Fecha
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{albaran.fechaalb}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    cadTABLA = tabla & " INNER JOIN albaran_variedad ON albaran.numalbar = albaran_variedad.numalbar "
    
    If HayRegParaInforme(cadTABLA, cadselect) Then
        Select Case OpcionListado
            Case 0
                If CargarTemporal(cadTABLA, cadselect) Then
                    cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
                
                    cadTitulo = "Informe de Incidencias"
                    cadNombreRPT = "rVtasListIncid.rpt"
                Else

                End If
            Case 1
                cadTitulo = "Informe de Categorias"
                cadNombreRPT = "rVtasListCateg.rpt"
        End Select
        LlamarImprimir
    End If
    
End Sub

Private Function CargarTemporal(cadTABLA As String, cadselect As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim CadValues As String
Dim ImpFactu As Currency
Dim PrecioFact As Currency
Dim Diferencia As Currency
Dim Importe As Currency

Dim Rs As ADODB.Recordset

    On Error GoTo eCargarTemporal
    
    CargarTemporal = False
    Screen.MousePointer = vbHourglass
    Label4(6).visible = True
    DoEvents
    
    Sql = "delete from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
    conn.Execute Sql
    
    Sql3 = "insert into tmpinformes (codusu, importe1, importe2, precio1, precio2, importe3) values "
    
    Sql = "select albaran_variedad.numalbar, albaran_variedad.numlinea, albaran_variedad.preciopro, albaran_variedad.pesoneto from " & cadTABLA
    If cadselect <> "" Then Sql = Sql & " WHERE " & cadselect
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    CadValues = ""
    While Not Rs.EOF
        Sql2 = "select sum(if(impornet is null,0,impornet)) from facturas_variedad where numalbar = " & DBSet(Rs.Fields!NumAlbar, "N")
        Sql2 = Sql2 & " and numlinealbar = " & DBSet(Rs!NumLinea, "N")
        
        ImpFactu = DevuelveValor(Sql2)
        PrecioFact = 0
        Diferencia = 0
        Importe = 0
        If CCur(ImpFactu) <> 0 Then
            If DBLet(Rs!PesoNeto, "N") <> 0 Then
                PrecioFact = Round2(ImpFactu / DBLet(Rs!PesoNeto, "N"), 4)
            End If
            Diferencia = PrecioFact - DBLet(Rs!preciopro, "N")
            Importe = Round2(DBLet(Rs!PesoNeto, "N") * Diferencia, 2)
        End If
        CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(Rs!NumAlbar, "N") & "," & DBSet(Rs!NumLinea, "N") & ","
        CadValues = CadValues & DBSet(PrecioFact, "N") & "," & DBSet(Diferencia, "N") & "," & DBSet(Importe, "N") & "),"
        
        Rs.MoveNext
    Wend
    
    If CadValues <> "" Then
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
        conn.Execute Sql3 & CadValues
    End If
    
    Screen.MousePointer = vbDefault
    Label4(6).visible = False
    DoEvents
    Set Rs = Nothing
    CargarTemporal = True
    Exit Function
    
eCargarTemporal:
    Screen.MousePointer = vbDefault
    Label4(6).visible = False
    MuestraError Err.Number, "Cargar Temporal", Err.Description
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 0:
                PonerFoco txtCodigo(12)
            Case 1:
                PonerFoco txtCodigo(6)
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
    Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Me.imgBuscar(1).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Me.imgBuscar(4).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Me.imgBuscar(5).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Me.imgBuscar(12).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Me.imgBuscar(13).Picture = frmPpal.imgListImages16.ListImages(1).Picture


    Select Case OpcionListado
        Case 0  ' informe de incidencias
            Label1.Caption = "Informe de Incidencias"
            FrameCategoria.visible = False
            FrameCategoria.Enabled = False
        Case 1
            Label1.Caption = "Informe de Categorias"
        
    End Select


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

Private Sub Form_Unload(Cancel As Integer)
    CerrarConexionMultibase
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

Private Sub frmInc_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Incidencias
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Variedades
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
        Case 0, 1 'CLIENTES
            AbrirFrmClientes (Index)
            
        Case 4, 5 'VARIEDADES
            AbrirFrmVariedades (Index)
        
        Case 12, 13 'INCIDENCIAS
            AbrirFrmIncidencias (Index)
        
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
            Case 4: KEYBusqueda KeyAscii, 4 'variedad desde
            Case 5: KEYBusqueda KeyAscii, 5 'variedad hasta
            Case 2: KEYFecha KeyAscii, 2 'fecha desde
            Case 3: KEYFecha KeyAscii, 3 'fecha hasta
            Case 12: KEYBusqueda KeyAscii, 12 'incidencia desde
            Case 13: KEYBusqueda KeyAscii, 12 'incidencia hasta
            Case 0: KEYBusqueda KeyAscii, 0 'cliente desde
            Case 1: KEYBusqueda KeyAscii, 1 'cliente hasta
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
        Case 2, 3 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 4, 5 'VARIEDAD
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
            
        Case 0, 1 'CLIENTES
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "clientes", "nomclien", "codclien", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
                        
        Case 12, 13 'INCIDENCIAS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "inciden", "nomincid", "codincid", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
                        
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 6300
        Me.FrameCobros.Width = 6660
        W = Me.FrameCobros.Width
        H = Me.FrameCobros.Height
        
        If OpcionListado = 1 Then
            FrameCategoria.Top = 810
            FrameCategoria.Left = 420
        End If
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
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .ConSubInforme = True
        .Opcion = 0
        .Show vbModal
    End With
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
 
Private Sub AbrirFrmClientes(indice As Integer)
    indCodigo = indice
    Set frmCli = New frmClientes
    frmCli.DatosADevolverBusqueda = "0|2|"
'    frmCli.CodigoActual = txtCodigo(indCodigo)
    frmCli.Show vbModal
    Set frmCli = Nothing
End Sub
 
 
Private Sub AbrirFrmIncidencias(indice As Integer)
    indCodigo = indice
    Set frmInc = New frmManInciden
    frmInc.DatosADevolverBusqueda = "0|1|"
    frmInc.DeConsulta = True
    frmInc.CodigoActual = txtCodigo(indCodigo)
    frmInc.Show vbModal
    Set frmInc = Nothing
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
Dim Sql As String
Dim Rs As ADODB.Recordset

    Sql = "Select * FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    Sql = Sql & " group by 1 "
    Sql = Sql & " having sum(totalfac) > " & DBSet(txtCodigo(6).Text, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Rs.EOF Then
        MsgBox "No hay datos para mostrar en el Informe.", vbInformation
        HayRegistros = False
    Else
        HayRegistros = True
    End If

End Function

