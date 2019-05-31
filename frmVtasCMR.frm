VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmVtasCMR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   8505
   Icon            =   "frmVtasCMR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   8505
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
      Height          =   8520
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   8310
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
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
         Index           =   8
         Left            =   7110
         MaxLength       =   8
         TabIndex        =   1
         Top             =   1020
         Visible         =   0   'False
         Width           =   840
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
         Index           =   7
         Left            =   2340
         MaxLength       =   60
         TabIndex        =   3
         Top             =   1890
         Width           =   5640
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Consignatario el Destino"
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
         Height          =   330
         Index           =   1
         Left            =   375
         TabIndex        =   20
         Top             =   7350
         Width           =   3075
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
         Index           =   6
         Left            =   2340
         MaxLength       =   40
         TabIndex        =   7
         Top             =   4500
         Width           =   5640
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
         Index           =   5
         Left            =   2340
         MaxLength       =   40
         TabIndex        =   6
         Top             =   3780
         Width           =   5640
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
         Index           =   4
         Left            =   2340
         MaxLength       =   40
         TabIndex        =   5
         Top             =   3375
         Width           =   5640
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Remitente la Cooperativa"
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
         Height          =   330
         Index           =   0
         Left            =   360
         TabIndex        =   9
         Top             =   7020
         Width           =   3075
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
         Height          =   1590
         Index           =   1
         Left            =   2340
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   5265
         Width           =   5640
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
         Index           =   0
         Left            =   2340
         MaxLength       =   80
         TabIndex        =   4
         Top             =   2595
         Width           =   5640
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
         Index           =   3
         Left            =   2340
         MaxLength       =   60
         TabIndex        =   2
         Top             =   1485
         Width           =   5640
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
         Left            =   2340
         MaxLength       =   10
         TabIndex        =   0
         Top             =   1035
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
         Left            =   6915
         TabIndex        =   11
         Top             =   7785
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
         Left            =   5730
         TabIndex        =   10
         Top             =   7785
         Width           =   1065
      End
      Begin VB.Label Label4 
         Caption         =   "Hora Carga"
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
         Index           =   5
         Left            =   5895
         TabIndex        =   21
         Top             =   1005
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label Label4 
         Caption         =   "16.- Nombre Transportista"
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
         Height          =   300
         Index           =   4
         Left            =   405
         TabIndex        =   19
         Top             =   4185
         Width           =   3300
      End
      Begin VB.Label Label4 
         Caption         =   "13.- Instrucciones del Remitente"
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
         Index           =   2
         Left            =   405
         TabIndex        =   18
         Top             =   3015
         Width           =   3570
      End
      Begin VB.Label Label1 
         Caption         =   "Listado CMR"
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
         TabIndex        =   17
         Top             =   315
         Width           =   5160
      End
      Begin VB.Label Label4 
         Caption         =   "19.- Estipulaciones Particulares"
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
         Height          =   300
         Index           =   1
         Left            =   405
         TabIndex        =   16
         Top             =   4950
         Width           =   3210
      End
      Begin VB.Label Label4 
         Caption         =   "5.- Documentos Anexos "
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
         Index           =   0
         Left            =   405
         TabIndex        =   15
         Top             =   2280
         Width           =   2670
      End
      Begin VB.Label Label4 
         Caption         =   "2.- Consignatario"
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
         Index           =   3
         Left            =   405
         TabIndex        =   14
         Top             =   1485
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Carga"
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
         Left            =   405
         TabIndex        =   13
         Top             =   1035
         Width           =   1365
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   2025
         Picture         =   "frmVtasCMR.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   1020
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmVtasCMR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor Monica +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public NumCod As String 'Para indicar el numero de albaran
Public NomTrans As String 'indicamos el nombre de transportista
Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

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

Dim CodTipoMov As String
'Codigo tipo de movimiento en función del valor en la tabla de parámetros: stipom


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
Dim Contador As Long
Dim vTipoMov As CTiposMov
Dim Mens As String
Dim ContCMR As String
Dim Sql As String


    InicializarVbles
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = "|pFechaCarga=""" & txtCodigo(2).Text & """"
    numParam = numParam + 1
    
    cadParam = cadParam & "|pHoraCarga=""" & txtCodigo(8).Text & """"
    numParam = numParam + 1
    
    ' consignatario
    cadParam = cadParam & "|pConsignatario=""" & txtCodigo(3).Text & """"
    numParam = numParam + 1
    cadParam = cadParam & "|pConsignatari1=""" & txtCodigo(7).Text & """"
    numParam = numParam + 1
    
    ' documentos anexos
    cadParam = cadParam & "|pDocumAnexos=""" & txtCodigo(0).Text & """"
    numParam = numParam + 1
    
    ' instrucciones del remitente
    cadParam = cadParam & "|pInstruc=""" & txtCodigo(4).Text & """"
    numParam = numParam + 1
    cadParam = cadParam & "|pInstruc2=""" & txtCodigo(5).Text & """"
    numParam = numParam + 1
    
    ' estipulaciones particulares
    cadParam = cadParam & "|pEstipulaciones=""" & txtCodigo(1).Text & """"
    numParam = numParam + 1
    
    ' remitente cooperativa
    cadParam = cadParam & "|pRemitente=" & Check1(0).Value
    numParam = numParam + 1
    
    ' Consignatario el destino
    cadParam = cadParam & "|pDestino=" & Check1(1).Value
    numParam = numParam + 1
    
    
    ' transportista
    cadParam = cadParam & "|pNomtrans=""" & txtCodigo(6).Text & """|"
    numParam = numParam + 1
    
    CodTipoMov = "CMR"
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        'contador del albaran
        ContCMR = ""
        ContCMR = DevuelveDesdeBDNew(cAgro, "albaran", "numerocmr", "numalbar", NumCod, "N")
        If ContCMR = "" Then
            Contador = vTipoMov.ConseguirContador(CodTipoMov)
        
            vTipoMov.IncrementarContador (CodTipoMov)
        
        Else
            Contador = CLng(ContCMR)
        End If
        cadParam = cadParam & "|pContador=" & Contador & "|"
        numParam = numParam + 1
'[Monica]03/01/2012: subo esta instruccion arriba dentro de : if ContCMR = "" then
'        vTipoMov.IncrementarContador (CodTipoMov)
    
        Set vTipoMov = Nothing
        
        If Not InsertarTemporal Then
            MsgBox "Error insertando en temporal. Llame a soporte.", vbExclamation
            Exit Sub
        End If
        
        AnyadirAFormula cadFormula, "{tmpcmr.numalbar} = " & NumCod & " and {tmpcmr.codusu} = " & vUsu.Codigo
        
        If CargarVariedades(Mens) Then
            'Nombre fichero .rpt a Imprimir
            cadTitulo = "Listado CMR"
            'antes cadNombreRPT = "rListCMR.rpt"
            indRPT = 11 'Impresion de listado CMR
            If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
            'Nombre fichero .rpt a Imprimir
            frmImprimir.NombreRPT = nomDocu
            cadNombreRPT = nomDocu
            LlamarImprimir
        Else
            MsgBox "Se ha producido un error " & Mens, vbExclamation
            Exit Sub
        End If
        
        If Not frmVisReport.EstaImpreso Then
            ' si el albaran no tenia contador asignado ya
            If ContCMR = "" Then
                'Devolvemos contador, si no estamos actualizando
                Set vTipoMov = New CTiposMov
                vTipoMov.DevolverContador CodTipoMov, Contador
                Set vTipoMov = Nothing
            End If
        Else
            ' actualizamos el contador del albaran
            Sql = "update albaran set numerocmr = " & DBSet(Contador, "N")
            Sql = Sql & " where numalbar = " & DBSet(NumCod, "N")
        
            conn.Execute Sql
        End If
        
        
        '[Monica]18/11/2015: en el caso de que tenga el cliente la marca de enviar por email
        If EnvioEMail(NumCod) Then
            '[Monica]26/11/2015: solo hay que imprimir la primera página
            cadParam = cadParam & "pEsCorreo=1|"
            numParam = numParam + 1
            With frmImprimir
               '[Monica]24/01/2012: añadido la siguientes 3 lineas para el envio por el outlook
                .outClaveNombreArchiv = Format(Contador, "0000000")
                .outCodigoCliProv = DevuelveValor("select codclien from albaran where numalbar = " & DBSet(NumCod, "N"))
                .outTipoDocumento = 7
                
                .FormulaSeleccion = cadFormula
                .OtrosParametros = cadParam
                .NumeroParametros = numParam
                .SoloImprimir = False
                .EnvioEMail = True
                .Titulo = cadTitulo
                .NombreRPT = cadNombreRPT
                .Opcion = 0
                .ConSubInforme = True
                .Show vbModal
            End With
        
        End If
        
        cmdCancel_Click
    
    End If
End Sub

Private Function EnvioEMail(Albaran As String) As Boolean
Dim Sql As String

    Sql = "select envcmremail from clientes, albaran where albaran.codclien = clientes.codclien and albaran.numalbar = " & DBSet(Albaran, "N")
    EnvioEMail = (DevuelveValor(Sql) = 1)

End Function


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(2)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection
Dim Sql As String

    PrimeraVez = True
    limpiar Me

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, H, W
    indFrame = 5
    tabla = "pedidos"
    
    txtCodigo(2).Text = Format(Now, "dd/mm/yyyy")
    txtCodigo(6).Text = NomTrans
    
    Sql = "Las partes intervinientes en este contrato se someten expresamente a  la Junta Arbitral del Transporte de "
    Sql = Sql & ProvAgenciaTransporte & "(España), incluso en controversias que excedan de 3005'06€"
    txtCodigo(1).Text = Sql
    
    txtCodigo(8).visible = (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16)
    txtCodigo(8).Enabled = (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16)
    txtCodigo(8).Text = Format(Time, "hh:mm:ss")
    Label4(5).visible = (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16)
    Label4(5).Enabled = (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16)
    
    
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(0).Tag) + 2).Text = Format(vFecha, "dd/MM/yyyy")
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
    imgFec(0).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtCodigo(Index + 2).Text <> "" Then frmC.NovaData = txtCodigo(Index + 2).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtCodigo(CByte(imgFec(0).Tag) + 2)
    ' ***************************
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
            Case 2: KEYFecha KeyAscii, 2 'fecha de carga
        End Select
    Else
        KEYpress KeyAscii
    End If

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
        Case 2 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
        Case 8 ' Hora
            If txtCodigo(Index).Text <> "" Then PonerFormatoHora txtCodigo(Index)
        
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 8520 '6585
        Me.FrameCobros.Width = 8310 '6930
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
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .Opcion = 0
        .EnvioEMail = False
        .ConSubInforme = True
        .Show vbModal
    End With
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


Private Function CargarVariedades(Mens As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim b As Boolean

    Sql = "delete from tmpinformes where codusu =" & vUsu.Codigo
    conn.Execute Sql
    
    Sql = "select sum(pesobrut) from albaran_variedad where numalbar = " & NumCod
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        If DBLet(Rs.Fields(0).Value, "N") > vParamAplic.LimPesoCMR Then
            Mens = "Repartir Pesos Brutos"
            b = RepartirBrutos(Mens, DBLet(Rs.Fields(0).Value, "N"))
        Else
            Mens = "Cargar Pesos Brutos"
            b = CargarBrutos(Mens)
        End If
    End If
    Set Rs = Nothing
    
    CargarVariedades = b
    
End Function

Private Function RepartirBrutos(Mens As String, SumaBrutos As Currency) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Importe As Currency
Dim CarImporte As String
Dim CadValues As String

    On Error GoTo eRepartirBrutos

    RepartirBrutos = False
    If SumaBrutos = 0 Then Exit Function

    Sql = "select numlinea, pesobrut from albaran_variedad where numalbar = " & NumCod
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Sql2 = "insert into tmpinformes (codusu, campo1, importe1) values "
    
    While Not Rs.EOF
        Importe = Round2(vParamAplic.LimPesoCMR * DBLet(Rs.Fields(1).Value, "N") / (SumaBrutos * 10), 0)
        CarImporte = CStr(Importe) & "0"
        
        CadValues = "(" & vUsu.Codigo & "," & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(CarImporte, "N") & ")"
        
        conn.Execute Sql2 & CadValues
    
        Rs.MoveNext
    Wend
    
    Rs.Close
    
    Set Rs = Nothing

eRepartirBrutos:
    If Err.Number = 0 Then
        RepartirBrutos = True
    End If
End Function

Private Function CargarBrutos(Mens As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Importe As Currency
Dim CarImporte As String
Dim CadValues As String
    
    On Error GoTo eCargarBrutos
    
    Sql = "select numlinea, pesobrut from albaran_variedad where numalbar = " & NumCod
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Sql2 = "insert into tmpinformes (codusu, campo1, importe1) values "
    
    While Not Rs.EOF
        Importe = Round2(DBLet(Rs.Fields(1).Value, "N") / 10, 0)
        CarImporte = CStr(Importe) & "0"
        
        CadValues = "(" & vUsu.Codigo & "," & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(CarImporte, "N") & ")"
        
        conn.Execute Sql2 & CadValues
        
        Rs.MoveNext
    Wend

eCargarBrutos:
    If Err.Number <> 0 Then
        Mens = Mens & vbCrLf & Err.Description
        CargarBrutos = False
    Else
        CargarBrutos = True
    End If
End Function

Private Function InsertarTemporal() As Boolean
Dim Sql As String
    
    On Error GoTo eInsertarTemporal
    
    InsertarTemporal = True

    conn.Execute "delete from tmpcmr where codusu = " & vUsu.Codigo
    
    Sql = "insert into tmpcmr(numlinea, codusu, numalbar) select numlinea, "
    Sql = Sql & vUsu.Codigo & "," & NumCod & " from tmpcopiascmr "
    conn.Execute Sql

eInsertarTemporal:
    If Err.Number <> 0 Then InsertarTemporal = False
End Function

Private Function ProvAgenciaTransporte() As String
Dim Sql As String
Dim Rs As ADODB.Recordset
    
    ProvAgenciaTransporte = ""
    
    Sql = "select protrans from agencias, albaran_transporte where albaran_transporte.numalbar = " & DBSet(NumCod, "N")
    Sql = Sql & " and albaran_transporte.codtrans = agencias.codtrans and albaran_transporte.esmedioppal = 1"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        ProvAgenciaTransporte = UCase(DBLet(Rs.Fields(0).Value, "T"))
    End If
    
End Function

