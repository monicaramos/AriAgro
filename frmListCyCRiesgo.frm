VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListCyCRiesgo 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6645
   Icon            =   "frmListCyCRiesgo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
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
      Height          =   4275
      Left            =   45
      TabIndex        =   6
      Top             =   0
      Width           =   6555
      Begin VB.CheckBox Check1 
         Caption         =   "Sólo con riesgo concedido"
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
         Height          =   240
         Left            =   2565
         TabIndex        =   19
         Top             =   3240
         Width           =   3570
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Datos Seguro"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   495
         TabIndex        =   16
         Top             =   3240
         Width           =   2670
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
         Index           =   3
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1650
         Width           =   1350
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
         Index           =   2
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1245
         Width           =   1350
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
         Left            =   5175
         TabIndex        =   5
         Top             =   3645
         Width           =   1020
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
         Left            =   4005
         TabIndex        =   4
         Top             =   3645
         Width           =   1020
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
         Index           =   0
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   2
         Top             =   2325
         Width           =   870
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
         Index           =   1
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   3
         Top             =   2745
         Width           =   870
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   0
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Text5"
         Top             =   2325
         Width           =   3585
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   1
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "Text5"
         Top             =   2745
         Width           =   3585
      End
      Begin VB.Label Label4 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   18
         Left            =   495
         TabIndex        =   18
         Top             =   3870
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.Label Label4 
         Caption         =   "Cargando tabla temporal"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   495
         TabIndex        =   17
         Top             =   3645
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.Label Label1 
         Caption         =   "Listado de Riesgo"
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
         Left            =   465
         TabIndex        =   15
         Top             =   270
         Width           =   5160
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Albarán"
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
         Height          =   255
         Index           =   16
         Left            =   465
         TabIndex        =   14
         Top             =   900
         Width           =   1815
      End
      Begin VB.Label Label4 
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
         Index           =   15
         Left            =   735
         TabIndex        =   13
         Top             =   1245
         Width           =   690
      End
      Begin VB.Label Label4 
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
         Index           =   14
         Left            =   735
         TabIndex        =   12
         Top             =   1650
         Width           =   645
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1440
         Picture         =   "frmListCyCRiesgo.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   1245
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1440
         Picture         =   "frmListCyCRiesgo.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   1650
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Index           =   13
         Left            =   735
         TabIndex        =   11
         Top             =   2325
         Width           =   645
      End
      Begin VB.Label Label4 
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
         Index           =   12
         Left            =   735
         TabIndex        =   10
         Top             =   2745
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Height          =   195
         Index           =   11
         Left            =   465
         TabIndex        =   9
         Top             =   1995
         Width           =   675
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1440
         MouseIcon       =   "frmListCyCRiesgo.frx":0122
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   2325
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1440
         MouseIcon       =   "frmListCyCRiesgo.frx":0274
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   2745
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmListCyCRiesgo"
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
Private WithEvents frmCli As frmClientes 'Clientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmTra As frmManAgencias 'Agencias de transporte
Attribute frmTra.VB_VarHelpID = -1
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




Private Sub Check3_Click()
    Check1.Enabled = (Check3.Value = 1)
    If (Check3.Value = 0) Then Check1.Value = 0
End Sub


Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String

Dim Sql As String

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
    
    '[Monica]26/04/2019: para el caso del nuevo listado de riesgo por cliente
    If Check3.Value = 1 Then
        cadselect = Replace(cadselect, "albaran", "clientes")
        
        '[Monica]26/04/2019: para el caso de los datos de seguro
        If Check1.Value = 1 Then
            If Not AnyadirAFormula(cadselect, "(not isnull({clientes.nroseguro}) and {clientes.nroseguro}<> """")") Then Exit Sub
        End If
        
        If HayRegistros("clientes", cadselect) Then
            If CargarTemporal2("clientes", cadselect) Then
                cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
                
                cadTitulo = "Listado de Riesgo"
                indRPT = 124 'Listado de Riesgos
                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                
                cadNombreRPT = nomDocu
                cadselect = cadFormula
                If HayRegistros("tmpinformes", cadselect) Then
                    LlamarImprimir
                Else
                    MsgBox "No hay registros para imprimir", vbExclamation
                End If
            End If
        End If
        Exit Sub
    End If
    
    
    'D/H Fecha Albarán
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".fechaalb}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    cadTABLA = tabla & " INNER JOIN clientes ON albaran.codclien = clientes.codclien "
    
    
    cadParam = cadParam & "pNroPoliza=""" & Trim(vParamAplic.NroPolizaExp) & """|"
    numParam = numParam + 1
    
    If HayRegistros(cadTABLA, cadselect) Then
        If CargarTemporal(cadTABLA, cadselect) Then
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            
            If vParamAplic.Cooperativa = 4 Then
                cadFormula = cadFormula & " and {tmpinformes.importe5} = 0 "
            End If
            
            cadTitulo = "Listado de Riesgo"
            indRPT = 94 'Listado de Riesgos
            If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
            
            cadNombreRPT = nomDocu
            cadselect = cadFormula
            If HayRegistros("tmpinformes", cadselect) Then
                LlamarImprimir
            Else
                MsgBox "No hay registros para imprimir", vbExclamation
            End If
        End If
    End If
End Sub

Private Function CargarTemporal2(tabla As String, vWhere As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RSaux As ADODB.Recordset
Dim importe1 As Currency
Dim importe2 As Currency
Dim Importe3 As Currency
Dim Importe4 As Currency
Dim SqlValues As String
Dim SQLinsert As String
Dim ImporteRiesgo As Currency

    On Error GoTo eCargaTemporal2

    CargarTemporal2 = False


    Label4(17).Caption = "Cargando tabla temporal"
    Label4(17).visible = True
    Me.Refresh

    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    SQLinsert = "insert into tmpinformes (codusu, codigo1, nombre1, importe1, importe2, importe3, importe4, importe5) values "
    SqlValues = ""

    Sql = "select codclien, nomclien, limiteriesgos, codmacta from clientes where (1=1)"
    If vWhere <> "" Then Sql = Sql & " and " & vWhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    While Not Rs.EOF
        Label4(17).Caption = "Cargando cliente " & DBLet(Rs!CodClien)
        
        Label4(18).Caption = "Cobros Pendientes"
        Me.Refresh
        ' cobros pendientes
        Sql = "select sum(impvenci+coalesce(gastos,0)-coalesce(impcobro,0)) from cobros where codmacta = " & DBSet(Rs!Codmacta, "T")
        Set RSaux = New ADODB.Recordset
        importe1 = 0
        RSaux.Open Sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RSaux.EOF Then
            importe1 = DBLet(RSaux.Fields(0).Value, "N")
        End If
        
        Label4(18).Caption = "Facturas no contabilizadas"
        Me.Refresh
        ' facturas no contabilizadas
        Sql = "select sum(totalfac) from facturas where codclien = " & DBSet(Rs!CodClien, "N") & " and intconta = 0"
        importe2 = DevuelveValor(Sql)
        
        Label4(18).Caption = "Albaranes sin Factura"
        Me.Refresh
        ' albaranes sin factura
        Sql = "select sum(round(coalesce(pesoneto,0) * if(coalesce(preciodef,0)= 0 ,coalesce(preciopro,0),coalesce(preciodef,0)),2)) from albaran_variedad inner join albaran on albaran_variedad.numalbar = albaran.numalbar "
        Sql = Sql & " where albaran.codclien = " & DBSet(Rs!CodClien, "N")
        Sql = Sql & " and not (albaran_variedad.numalbar, albaran_variedad.numlinea) in (select numalbar, numlinealbar from facturas_variedad)"
        Importe3 = DevuelveValor(Sql)
        
        Label4(18).Caption = "Pedidos sin Albaran"
        Me.Refresh
        ' pedidos sin albaran
        Sql = "select sum(round(pesoneto * coalesce(preciopro,0),2)) from pedidos_variedad inner join pedidos on pedidos_variedad.numpedid = pedidos.numpedid "
        Sql = Sql & " where pedidos.codclien = " & DBSet(Rs!CodClien, "N")
        Sql = Sql & " and (pedidos.numalbar is null or pedidos.numalbar = 0)"
        Importe4 = DevuelveValor(Sql)
        
        ImporteRiesgo = DBLet(Rs!limiteRiesgos, "N")
        SqlValues = SqlValues & ",(" & vUsu.Codigo & "," & DBSet(Rs!CodClien, "N") & "," & DBSet(Rs!Nomclien, "T")
        SqlValues = SqlValues & "," & DBSet(ImporteRiesgo, "N") & "," & DBSet(importe1, "N") & "," & DBSet(importe2, "N")
        SqlValues = SqlValues & "," & DBSet(Importe3, "N") & "," & DBSet(Importe4, "N") & ")"
        
        Rs.MoveNext
    Wend
    
    If SqlValues <> "" Then
        SqlValues = Mid(SqlValues, 2)
        
        conn.Execute SQLinsert & SqlValues
    End If
    
    Set Rs = Nothing

    CargarTemporal2 = True
    Label4(17).visible = False
    Label4(18).visible = False
    
    Exit Function

eCargaTemporal2:
    MuestraError Err.Number, "Cargar Tabla Temporal", Err.Description
End Function


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
     For H = 0 To 1
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Next H

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, H, W
    indFrame = 5
    tabla = "albaran"
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
Dim Cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
            
        Case 0, 1 'CLIENTE
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "clientes", "nomclien", "codclien", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
        
        Case 2, 3, 6 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 4, 5 'ALBARANES
            If txtCodigo(Index).Text <> "" Then PonerFormatoEntero txtCodigo(Index)
        
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 4710
        Me.FrameCobros.Width = 6735
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
        .FormulaSeleccion = cadFormula ' "{tmpinformes.codusu} = " & vUsu.codigo
        .OtrosParametros = cadParam
        .NumeroParametros = numParam + 1
        .SoloImprimir = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .EnvioEMail = False
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
Dim Sql As String
Dim Rs As ADODB.Recordset

    Sql = "Select * FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Rs.EOF Then
        MsgBox "No hay datos para mostrar en el Informe.", vbInformation
        HayRegistros = False
    Else
        HayRegistros = True
    End If

End Function


' necesario cargar la tabla temporal pq los importes son con el iva del cliente de la contabilidad

Private Function CargarTemporal(cTabla As String, cadwhere As String) As Boolean
Dim Sql As String
Dim SQL1 As String
Dim Sql2 As String
Dim Codiva As String
Dim PorcIva As String
Dim i As Integer
Dim HayReg As Integer
Dim b As Boolean
Dim Registro As String
Dim vCliente As CCliente
Dim Rs As ADODB.Recordset
Dim Importe As Currency
Dim Dias As Long
Dim RsFact As ADODB.Recordset
Dim SqlFact As String
Dim BaseImpo As Currency
Dim ImpoIva As Currency


On Error GoTo eCargarTemporal

    HayReg = 0
    
    CargarTemporal = False
    
    conn.Execute "delete from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
        
    If cadwhere <> "" Then
        cadwhere = QuitarCaracterACadena(cadwhere, "{")
        cadwhere = QuitarCaracterACadena(cadwhere, "}")
        cadwhere = QuitarCaracterACadena(cadwhere, "_1")
    End If
    
    '[Monica]20/02/2012: Cambiamos los clientes asegurados y no asegurados.
    '           Los No Asegurados son los que tienen el importe a 0
    '           Los Asegurados son los que tienen el importe a 0 (antes era con el nro de seguro)
    '                    tipo: 0 los asegurados (clientes con importe a 0)
    '                          1 los no asegurados (clientes con importe <> 0)
    'antes
    '                             ' tipo: 0 los asegurados (clientes con nro de seguro)
    '                             '       1 los no asegurados (clientes sin el nro de seguro)
                                  '(codusu, tipo,     pais,     albaran,  linea,  fecalbar,cliente, imp.factu,imp.provi,dias , baseimp    impoiva,    nrofactu, fecfactu,codtipom
    Sql = "insert into tmpinformes (codusu, importe5, importe4, importe1, campo1, fecha1,  codigo1, importe2, importe3, campo2, importeb1, importeb2, importeb3, fecha2, nombre1) values "
    
    Sql2 = "select 0 tipo, clientes.codpaise, albaran.fechaalb, albaran.codclien, clientes.tipoiva, clientes.diasasegurados, albaran_variedad.* "
    Sql2 = Sql2 & " from albaran_variedad, albaran, clientes where albaran.numalbar = albaran_variedad.numalbar "
    '[Monica]20/02/2012: Cambiamos los clientes asegurados y no asegurados.
'    Sql2 = Sql2 & " and albaran.codclien = clientes.codclien and clientes.nroseguro <> '' and not clientes.nroseguro is null "
    Sql2 = Sql2 & " and albaran.codclien = clientes.codclien and clientes.limiteriesgos <> 0 and not clientes.limiteriesgos is null "

    If cadwhere <> "" Then Sql2 = Sql2 & " and " & cadwhere
    Sql2 = Sql2 & " union "
    Sql2 = Sql2 & "select 1 tipo, clientes.codpaise, albaran.fechaalb, albaran.codclien, clientes.tipoiva, clientes.diasasegurados, albaran_variedad.* "
    Sql2 = Sql2 & " from albaran_variedad, albaran, clientes where albaran.numalbar = albaran_variedad.numalbar "
    '[Monica]20/02/2012: Cambiamos los clientes asegurados y no asegurados.
'    Sql2 = Sql2 & " and albaran.codclien = clientes.codclien and (clientes.nroseguro = '' or clientes.nroseguro is null) "
    Sql2 = Sql2 & " and albaran.codclien = clientes.codclien and (clientes.limiteriesgos = 0 or clientes.limiteriesgos is null) "
    If cadwhere <> "" Then Sql2 = Sql2 & " and " & cadwhere
    
    Sql2 = Sql2 & " order by 1,2 "
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Registro = ""
    
    While Not Rs.EOF
        Registro = Registro & "(" & vUsu.Codigo & "," & DBSet(Rs!Tipo, "N") & "," & DBSet(Rs!CodPaise, "N") & "," & DBSet(Rs!NumAlbar, "N") & "," & DBSet(Rs!NumLinea, "N") & ","
        Registro = Registro & DBSet(Rs!FechaAlb, "F") & "," & DBSet(Rs!CodClien, "N") & ","
                
        If AlbaranFacturado(CCur(Rs!NumAlbar), CCur(Rs!NumLinea)) Then
            Importe = CCur(ImporteAlbaranFacturado(CCur(Rs!NumAlbar), CCur(Rs!NumLinea)))
            
            Sql2 = "select codigiva from facturas_variedad where numalbar = " & DBSet(Rs!NumAlbar, "N")
            Sql2 = Sql2 & " and numlinealbar = " & DBSet(Rs!NumLinea, "N")
            Codiva = DevuelveValor(Sql2)
            
            PorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", Codiva, "N")
            
            BaseImpo = Importe
            ImpoIva = Round2(Importe * CCur(PorcIva) / 100, 2)
            ' obtenemos el importe con iva
            Importe = Importe + Round2(Importe * CCur(PorcIva) / 100, 2)
            
            Registro = Registro & DBSet(Importe, "N") & ",0,"
            Registro = Registro & DBSet(Rs!Diasasegurados, "N") & ","
            
            '[Monica]23/05/2013: añadimos los datos para el listado de Alzira
            SqlFact = "select codtipom, numfactu, fecfactu from facturas_variedad where numalbar = " & DBSet(Rs!NumAlbar, "N") & " and numlinealbar = " & DBSet(Rs!NumLinea, "N")
            Set RsFact = New ADODB.Recordset
            RsFact.Open SqlFact, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RsFact.EOF Then
               Registro = Registro & DBSet(BaseImpo, "N") & "," & DBSet(ImpoIva, "N") & "," & DBSet(RsFact!NumFactu, "N") & "," & DBSet(RsFact!FecFactu, "F") & "," & DBSet(RsFact!codTipoM, "T") & "),"
            Else
               Registro = Registro & "0,0,0,null,null),"
            End If
            Set RsFact = Nothing
        Else
            ' calculamos el importe provisional
            Importe = Round2(DBLet(Rs!Pesoneto, "N") * DBLet(Rs!preciopro, "N"), 2)
            Select Case Rs!TipoIva
                Case 0
                    Codiva = DevuelveDesdeBDNew(cAgro, "variedades", "codigiva", "codvarie", Rs!codvarie, "N")
                Case 1
                    Codiva = vParamAplic.CodIvaExento
                Case 2
                    Codiva = vParamAplic.CodIvaRecargo
            End Select
            PorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", Codiva, "N")
            
            ' obtenemos el importe con iva
            Importe = Importe + Round2(Importe * CCur(PorcIva) / 100, 2)
        
            Registro = Registro & "0," & DBSet(Importe, "N") & ","
            Registro = Registro & DBSet(Rs!Diasasegurados, "N") & ","
            Registro = Registro & "0,0,0,null,null),"
        End If
        
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    ' quitamos la ultima coma
    If Len(Registro) > 0 Then
        Registro = Mid(Registro, 1, Len(Registro) - 1)
    
        conn.Execute Sql & Registro
    End If
        
        
    CargarTemporal = True

    Exit Function

eCargarTemporal:
    MuestraError Err.Number, "Cargando Temporal", Err.Description
End Function

Private Function DatosOk() As Boolean

    DatosOk = False
        
    DatosOk = True

End Function

