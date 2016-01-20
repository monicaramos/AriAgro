VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmANECOOPTrasPagos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6630
   Icon            =   "frmANECOOPTrasPagos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameIntegracion 
      Height          =   5415
      Left            =   30
      TabIndex        =   6
      Top             =   60
      Width           =   6555
      Begin VB.TextBox txtNombre 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1620
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos para la contabilización"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1785
         Left            =   210
         TabIndex        =   13
         Top             =   2190
         Width           =   6165
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   2655
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   1320
            Width           =   3225
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   3
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1320
            Width           =   1005
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   2655
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   900
            Width           =   3225
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   2
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   900
            Width           =   1005
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   14
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   1
            Top             =   510
            Width           =   1005
         End
         Begin VB.Label Label1 
            Caption         =   "Cta Banco"
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
            Index           =   1
            Left            =   300
            TabIndex        =   18
            Top             =   1350
            Width           =   735
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   1200
            ToolTipText     =   "Buscar cuenta"
            Top             =   1350
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Diario"
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
            Index           =   0
            Left            =   300
            TabIndex        =   16
            Top             =   930
            Width           =   735
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   0
            Left            =   1200
            ToolTipText     =   "Buscar diario"
            Top             =   930
            Width           =   240
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   0
            Left            =   1200
            Picture         =   "frmANECOOPTrasPagos.frx":000C
            Top             =   510
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Pago"
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
            Index           =   23
            Left            =   270
            TabIndex        =   14
            Top             =   480
            Width           =   840
         End
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   0
         Top             =   1620
         Width           =   1005
      End
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   255
         Left            =   270
         TabIndex        =   10
         Top             =   3990
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4050
         TabIndex        =   4
         Top             =   4770
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5235
         TabIndex        =   5
         Top             =   4770
         Width           =   975
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   240
         Top             =   3510
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "doc"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Importe Liquidación"
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
         Index           =   1
         Left            =   3060
         TabIndex        =   19
         Top             =   1620
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nro.Liquidación"
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
         Index           =   0
         Left            =   450
         TabIndex        =   12
         Top             =   1590
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Traspaso de Pagos Anecoop"
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
         TabIndex        =   11
         Top             =   420
         Width           =   5430
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Proceso que realiza la integración de Pagos Anecoop."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   525
         Index           =   37
         Left            =   240
         TabIndex        =   9
         Top             =   1020
         Width           =   5820
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   270
         TabIndex        =   8
         Top             =   4470
         Width           =   6195
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   0
         Left            =   270
         TabIndex        =   7
         Top             =   4200
         Width           =   6195
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   5280
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmANECOOPTrasPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: LAURA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public Opcionlistado As Byte
    
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public Event RectificarFactura(Cliente As String, Observaciones As String)

Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes ' seleccionamos que facturas vamos a generar
Attribute frmMens.VB_VarHelpID = -1
Private WithEvents frmCtas As frmCtasConta 'cuentas contables
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmBas As frmBasico ' frmbasico
Attribute frmBas.VB_VarHelpID = -1


'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe
Private cadSelect1 As String 'Cadena para comprobar si hay datos antes de abrir Informe


Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Tabla1 As String
Dim codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim Tipo As String

Dim indice As Integer

Dim PrimeraVez As Boolean
Dim Contabilizada As Byte
Dim ConSubInforme As Boolean

Dim Facturas As String

Dim vClien As CCliente

Dim CodtipomAnecoop As String
Dim Codforpa As Integer
Dim letraser As String
Dim TipoIvac As Byte
Dim Dto1 As Currency
Dim Dto2 As Currency

Dim TotalAnecoop As Currency
Dim Nregs As Currency
Dim FecLiq As String


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub cmdAceptar_Click()
Dim Sql As String
Dim I As Byte
Dim cadWHERE As String
Dim b As Boolean
Dim NomFic As String
Dim CADENA As String
Dim cadena1 As String
Dim Directorio As String
Dim fec As String
Dim nomDir As String

Dim cadTABLA As String
Dim NomFic1 As String

Dim Mens As String
Dim cadErr As String

'On Error GoTo eError

    If Not DatosOk Then Exit Sub

    ' comprobamos que esten metidos todos los datos
    If Not ComprobarDesdobles Then
        cadFormula = "{tmpinformes.codusu} = " & vUsu.codigo
        cadTitulo = "Desdobles de Expedientes sin Pagos"
        cadNombreRPT = "rErroresExpAnecoop1.rpt"
        LlamarImprimir
        
        Exit Sub
    End If

    ' comprobamos que se puede hacer el pago de la liquidacion
    If ComprobarLiquidacion() Then
        If Nregs <> 0 Then
            cmdAceptar.Enabled = False
            cmdCancelar.Enabled = False
        
            If InsertarAsientoDiario(txtcodigo(14).Text, txtcodigo(3).Text, txtcodigo(1).Text, txtcodigo(0).Text, FecLiq, cadErr) Then
                
                MsgBox "Proceso realizado correctamente.", vbExclamation
                
            Else
            
                MsgBox "No se ha realizado el proceso. " & vbCrLf & cadErr, vbExclamation
            
            End If
        Else
        
            MsgBox "No hay registros para el número de liquidación.", vbExclamation
            
        End If
    End If
        
    cmdAceptar.Enabled = True
    cmdCancelar.Enabled = True
    
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""
                
    Unload Me
    
End Sub


Private Function ComprobarDesdobles()
Dim Sql As String
Dim SqlValues As String
Dim RS As ADODB.Recordset

    On Error GoTo eComprobarDesdobles
        
    ComprobarDesdobles = False
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.codigo
    conn.Execute Sql
    
    Sql = "select distinct expediente_id"
    Sql = Sql & " from anecoop_pago ff"
    Sql = Sql & " where ff.num_liquidacion = " & DBSet(txtcodigo(0).Text, "N")
    Sql = Sql & " and (mid(ff.expediente_id,1,1) = '0' or length(ff.expediente_id) <> 18) "
    Sql = Sql & " order by 1 "
    
    SqlValues = ""
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        If Len(RS!expediente_id) <> 18 Then
            Sql = "select count(*) from anecoop_pago where mid(expediente_id,2,17) = right(concat('000000000000000'," & DBSet(RS!expediente_id, "T") & "),17) "
            Sql = Sql & " and mid(expediente_id,1,1)  <> '0' "
            Sql = Sql & " and num_liquidacion = " & DBSet(txtcodigo(0).Text, "N")
            
            If TotalRegistros(Sql) = 0 Then
                Sql = "select expediente_id from anecoop where mid(expediente_id,1,1) <> '0' and mid(expediente_id,2,17) = right(concat('000000000000000'," & DBSet(RS!expediente_id, "T") & "),17)  "
                If TotalRegistrosConsulta(Sql) <> 0 Then
                    SqlValues = SqlValues & "(" & vUsu.codigo & "," & DBSet(DevuelveValor(Sql), "T") & "),"
                End If
            End If
        
        
        Else
            If Mid(RS!expediente_id, 1, 1) = "0" Then
            
                Sql = "select count(*) from anecoop_pago where mid(expediente_id,2,17) = mid(" & DBSet(RS!expediente_id, "T") & ",2,17) "
                Sql = Sql & " and mid(expediente_id,1,1)  <> '0' "
                Sql = Sql & " and num_liquidacion = " & DBSet(txtcodigo(0).Text, "N")
                
                If TotalRegistros(Sql) = 0 Then
                    Sql = "select expediente_id from anecoop where mid(expediente_id,1,1) <> '0' and mid(expediente_id,2,17) = mid(" & DBSet(RS!expediente_id, "T") & ",2,17) "
                    If TotalRegistrosConsulta(Sql) <> 0 Then
                        SqlValues = SqlValues & "(" & vUsu.codigo & "," & DBSet(DevuelveValor(Sql), "T") & "),"
                    End If
                End If
            
            End If
        End If
        RS.MoveNext
    Wend
    Set RS = Nothing
    
    If SqlValues <> "" Then
        SqlValues = Mid(SqlValues, 1, Len(SqlValues) - 1)
        
        conn.Execute "insert into tmpinformes (codusu, nombre1) values " & SqlValues
    End If
    
'    Sql = " insert into tmpinformes (codusu, nombre1) "
'    Sql = Sql & " select distinct " & vUsu.Codigo & ", dd.expediente_id from anecoop dd, anecoop_pago ff "
'    Sql = Sql & " where ff.num_liquidacion = " & DBSet(txtCodigo(0).Text, "N")
'    Sql = Sql & " and mid(dd.expediente_id,1,1) = '0' "
'    Sql = Sql & " and mid(dd.expediente_id,2,17) = right(concat('000000000000000000',ff.expediente_id),17)"
'    Sql = Sql & " and (not dd.expediente_id in (select ss.expediente_id from anecoop_pago  ss) or "
'    Sql = Sql & " dd.expediente_id in (select ss.expediente_id from anecoop_pago  ss where ss.num_liquidacion <> " & DBSet(txtCodigo(0).Text, "N") & "))"
'
'    conn.Execute Sql
    
    Sql = "select count(*) from tmpinformes where codusu = " & vUsu.codigo
    
    ComprobarDesdobles = (TotalRegistros(Sql) = 0)
    Exit Function


eComprobarDesdobles:
    MuestraError Err.Number, "Comprobar Desdobles", Err.Description
End Function



Private Sub cmdCancelar_Click()
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
    
    ConSubInforme = False

    For H = 0 To imgBuscar.Count - 1
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    
    'Ocultar todos los Frames de Formulario
    FrameIntegracion.visible = False
    '###Descomentar
'    CommitConexion
        
    FrameIntegracionVisible True, H, W
    pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
'    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub


Private Sub FrameIntegracionVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de socios por seccion
    Me.FrameIntegracion.visible = visible
    If visible = True Then
        Me.FrameIntegracion.Top = -90
        Me.FrameIntegracion.Left = 0
        Me.FrameIntegracion.Height = 5295
        Me.FrameIntegracion.Width = 6555
        W = Me.FrameIntegracion.Width
        H = Me.FrameIntegracion.Height
    End If
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadSelect1 = ""
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
        .EnvioEMail = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .ConSubInforme = ConSubInforme
        .Opcion = Opcionlistado
        .Show vbModal
    End With
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
        .Opcion = Opcionlistado
'        .ExportarPDF = (chkEMAIL.Value = 1)
        .Show vbModal
    End With
    
'    If Me.chkEMAIL.Value = 1 Then
'    '####Descomentar
'        If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
'    End If
    Unload Me
End Sub



Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Sql As String
Dim Sql2 As String
' añadido
Dim Mens As String
Dim NumFactu As String
Dim numser As String
Dim fecha As Date
Dim vCont As CTiposMov
Dim tipoMov As String

    b = True
    ' Datos contables introducidos
    If txtcodigo(14).Text = "" Then
        MsgBox "Debe de introducir la fecha de pago. Reintroduzca.", vbExclamation
        b = False
        PonerFoco txtcodigo(14)
    End If
    If b And txtcodigo(3).Text = "" Then
        MsgBox "Debe de introducir el número de diario. Reintroduzca.", vbExclamation
        b = False
        PonerFoco txtcodigo(3)
    End If
    If b And txtcodigo(1).Text = "" Then
        MsgBox "Debe de introducir la cuenta de banco. Reintroduzca.", vbExclamation
        b = False
        PonerFoco txtcodigo(1)
    End If
    
    ' Introducido nro de liquidacion
    If b And txtcodigo(0).Text = "" Then
        MsgBox "Debe introducir el número de liquidación. Reintroduzca.", vbExclamation
        b = False
        PonerFoco txtcodigo(0)
    End If
    
    DatosOk = b

End Function


Private Function ComprobarLiquidacion() As Boolean
Dim Sql As String
Dim CadResult As String
Dim CadResult2 As String
Dim RS As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Total As Currency
Dim SqlValues As String
Dim SQLinsert As String
Dim Sql2 As String
Dim Sql3 As String
Dim SqlNue As String
Dim Mens As String
Dim I As Long

    On Error GoTo eComprobarLiquidacion

    ComprobarLiquidacion = False
    
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.codigo
    conn.Execute Sql
    
    Sql = "delete from tmpinformes2 where codusu = " & vUsu.codigo
    conn.Execute Sql
    
    '[Monica]30/06/2015: introduzco q cobros han sido revisados ya para no volverlos a coger.
    '                    Solucionamos el problema de envases repetidos en el mismo albaran de venta nuestro
    '                    Utilizo esto pq hago una insercion global al final del proceso sobre tmpinformes para hacer el asiento. Necesito que la
    '                       insercion se haga en linea para que no vuelva a coger el mismo ordinal de 2 envases q esten en el mismo albaran y solucionar
    '                       el problema.
    Sql = "delete from tmprfactsoc where codusu = " & vUsu.codigo
    conn.Execute Sql
    
    CadResult = ""
    CadResult2 = ""
    
    Set vClien = New CCliente
    If vClien.LeerDatos(vParamAplic.CodAnecoop) Then
        CodtipomAnecoop = vClien.tipoMov
        Codforpa = vClien.ForPago
        TipoIvac = vClien.TipoIva
        Dto1 = vClien.Dto1
        Dto2 = vClien.Dto2
    Else
        Exit Function
    End If
    
    Total = 0
    
    Sql = "select ll.*, cc.nombre_variedad, cc.numero_salida_cooperativa, cc.numlinea from anecoop_pago ll, anecoop cc where ll.idcontab = 0 and ll.num_liquidacion  = " & DBSet(txtcodigo(0).Text, "N")
    Sql = Sql & " and ll.expediente_id = cc.expediente_id"
    
    TotalAnecoop = DevuelveValor("select sum(importe) from (" & Sql & ") aaaa")
    
    Nregs = TotalRegistrosConsulta(Sql)
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Nregs <> 0 Then
        'fecha que pondremos en el documento del apunte del asiento al debe del banco
        FecLiq = DBLet(RS!fecha_pago, "F")
        
        lblProgres(0).Caption = "Comprobando datos: "
        pb1.visible = True
        pb1.Max = Nregs
        pb1.Value = 0
        DoEvents
        I = 0
        
        While Not RS.EOF
            I = I + 1
            lblProgres(1).Caption = "Linea " & I
            IncrementarProgresNew pb1, 1
            DoEvents
            
            
            letraser = ObtenerLetraSerie(CodtipomAnecoop)
        
            '[Monica]02/06/2015: si el nro de albaran es alfa damos un error y salimos
            If Not IsNumeric(DBLet(RS!numero_salida_cooperativa, "T")) Then
                MsgBox "El nro de albarán del expediente " & DBLet(RS!expediente_id) & " no es numérico." & vbCrLf & vbCrLf & "Revise.", vbExclamation
                Exit Function
            End If
        
            Sql = "select * from conta" & vParamAplic.NumeroConta & ".scobro where numserie = " & DBSet(letraser, "T") & " and codfaccl = " & DBSet(RS!num_factura, "N")
            Sql = Sql & " and fecfaccl = " & DBSet(RS!fecha_factura, "F")
            
            If Mid(DBLet(RS!tipo_pago), 1, 1) = "I" Then
                If DBLet(RS!nombre_variedad) = "" Then
                    Sql = Sql & " and coalesce(referencia,'') = 'IVA ENVASE'"
                    Sql = Sql & " and cast(referencia1 as unsigned) = " & DBSet(RS!numero_salida_cooperativa, "N")
                Else
                    Sql = Sql & " and coalesce(referencia,'') = 'IVA VARIEDAD'"
                    Sql = Sql & " and cast(referencia1 as unsigned) = " & DBSet(RS!numero_salida_cooperativa, "N")
                    Sql = Sql & " and cast(referencia2 as unsigned) = " & DBSet(RS!numlinea, "N")
                End If
            Else
                If DBLet(RS!nombre_variedad) = "" Then
                    Sql = Sql & " and coalesce(referencia,'') = 'ENVASES'"
                    Sql = Sql & " and cast(referencia1 as unsigned) = " & DBSet(RS!numero_salida_cooperativa, "N")
                Else
                    Sql = Sql & " and cast(referencia1 as unsigned) = " & DBSet(RS!numero_salida_cooperativa, "N")
                    Sql = Sql & " and cast(referencia2 as unsigned) = " & DBSet(RS!numlinea, "N")
                    Sql = Sql & " and coalesce(referencia,'') <> 'IVA VARIEDAD'"
                End If
            End If
            
            If TotalRegistrosConsulta(Sql) = 0 Then
                If InStr(1, CadResult, DBLet(RS!expediente_id)) = 0 Then CadResult = CadResult & DBLet(RS!expediente_id) & ", "
            Else
                SqlNue = Sql & " and not (numserie, codfaccl, fecfaccl, numorden) in (select codtipom, numfactu, fecfactu, baseimpo from tmprfactsoc where codusu = " & DBSet(vUsu.codigo, "N") & ")"
                
                Set Rs2 = New ADODB.Recordset
                Rs2.Open SqlNue, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
                If Not Rs2.EOF Then
                    If DBLet(Rs2!ImpVenci) <> DBLet(RS!Importe) Then
                        ' caso de que una linea de albaran corresponde a 2 o mas pagos de anecoop
                    
                        CadResult2 = CadResult2 & DBLet(RS!expediente_id) & ", "
                        
                        Total = Total + DBLet(RS!Importe)
                    
                        SqlValues = SqlValues & "(" & vUsu.codigo & "," & DBSet(Rs2!numserie, "T") & "," & DBSet(Rs2!Codfaccl, "N") & "," & DBSet(Rs2!fecfaccl, "F") & ","
                        SqlValues = SqlValues & DBSet(Rs2!numorden, "N") & "),"
                        
                        Sql2 = "insert into tmpinformes2 (codusu, nombre1, nombre2) values (" & vUsu.codigo & ","
                        Sql2 = Sql2 & DBSet(RS!expediente_id, "T") & "," & DBSet(RS!expediente_pagoid, "T") & ")"
                        conn.Execute Sql2
                    
                        '[Monica]30/06/2015: insertamos en rfactsoc.baseimpo el numorden de la scobro
                        '                    nos guardamos los cobros que ya han sido procesados
                        Sql3 = "insert into tmprfactsoc (codusu, codtipom, numfactu, fecfactu, baseimpo) values ( "
                        Sql3 = Sql3 & vUsu.codigo & "," & DBSet(Rs2!numserie, "T") & "," & DBSet(Rs2!Codfaccl, "N") & "," & DBSet(Rs2!fecfaccl, "F") & ","
                        Sql3 = Sql3 & DBSet(Rs2!numorden, "N") & ")"
                        conn.Execute Sql3
                    
                    Else
                
                        Total = Total + DBLet(Rs2!ImpVenci)
                        
                        SqlValues = SqlValues & "(" & vUsu.codigo & "," & DBSet(Rs2!numserie, "T") & "," & DBSet(Rs2!Codfaccl, "N") & "," & DBSet(Rs2!fecfaccl, "F") & ","
                        SqlValues = SqlValues & DBSet(Rs2!numorden, "N") & "),"
                        
                        Sql2 = "insert into tmpinformes2 (codusu, nombre1, nombre2) values (" & vUsu.codigo & ","
                        Sql2 = Sql2 & DBSet(RS!expediente_id, "T") & "," & DBSet(RS!expediente_pagoid, "T") & ")"
                        conn.Execute Sql2
                    
                        '[Monica]30/06/2015: insertamos en rfactsoc.baseimpo el numorden de la scobro
                        '                    nos guardamos los cobros que ya han sido procesados
                        Sql3 = "insert into tmprfactsoc (codusu, codtipom, numfactu, fecfactu, baseimpo) values ( "
                        Sql3 = Sql3 & vUsu.codigo & "," & DBSet(Rs2!numserie, "T") & "," & DBSet(Rs2!Codfaccl, "N") & "," & DBSet(Rs2!fecfaccl, "F") & ","
                        Sql3 = Sql3 & DBSet(Rs2!numorden, "N") & ")"
                        conn.Execute Sql3
                    
                    End If
                Else
'                    CadResult = CadResult & DBLet(Rs!expediente_id) & ", "
                    
                    Total = Total + DBLet(RS!Importe)
                    
                    Sql2 = "insert into tmpinformes2 (codusu, nombre1, nombre2) values (" & vUsu.codigo & ","
                    Sql2 = Sql2 & DBSet(RS!expediente_id, "T") & "," & DBSet(RS!expediente_pagoid, "T") & ")"
                    conn.Execute Sql2

                End If
                
                Set Rs2 = Nothing
            End If
        
            RS.MoveNext
        Wend
        
        Mens = ""
        If CadResult <> "" Or Total <> TotalAnecoop Then
            If CadResult <> "" Then
                Mens = "Las siguientes referencias no se encuentran en Cartera de Cobros: " & vbCrLf & vbCrLf & Mid(CadResult, 1, Len(CadResult) - 2)
            
                '[Monica]20/05/2015: sacamos en un listado lo que antes sacabamos en el mensaje
                '                    DEJAMOS CARGADO MENS PQ ES LO QUE INDICA Q HA HABIDO UN ERROR Y NO REALIZA EL PROCESO
                SQLinsert = "insert into tmpinformes (codusu, nombre1) select " & vUsu.codigo & ", expediente_id from anecoop where expediente_id in (" & Mid(CadResult, 1, Len(CadResult) - 2) & ")"
                conn.Execute SQLinsert
                
                '========= PARAMETROS  =============================
                'Añadir el parametro de Empresa
                cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
                numParam = numParam + 1
                
                cadFormula = "{tmpinformes.codusu} = " & vUsu.codigo
    
                Sql = "select count(*) from tmpinformes where codusu = " & vUsu.codigo
    
                If TotalRegistros(Sql) <> 0 Then
                    cadTitulo = "Expedientes no encontrados en Cobros"
                    cadNombreRPT = "rErroresExpAnecoop.rpt"
                    LlamarImprimir
                End If
                
            Else
                Mens = "No coinciden los totales con el total de Cartera de Cobros. Revise."
            
'                If Mens <> "" Then
'                    MsgBox Mens, vbExclamation
'                End If
                '[Monica]20/05/2015: sacamos en un listado lo que antes sacabamos en el mensaje
                '                    DEJAMOS CARGADO MENS PQ ES LO QUE INDICA Q HA HABIDO UN ERROR Y NO REALIZA EL PROCESO
                SQLinsert = "insert into tmpinformes (codusu, nombre1) select " & vUsu.codigo & ", expediente_id from anecoop where expediente_id in (" & Mid(CadResult2, 1, Len(CadResult2) - 2) & ")"
                conn.Execute SQLinsert
                
                '========= PARAMETROS  =============================
                'Añadir el parametro de Empresa
                cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
                numParam = numParam + 1
                
                cadFormula = "{tmpinformes.codusu} = " & vUsu.codigo
    
                Sql = "select count(*) from tmpinformes where codusu = " & vUsu.codigo
    
                If TotalRegistros(Sql) <> 0 Then
                    cadTitulo = "Expedientes de importes diferentes en Cobros"
                    cadNombreRPT = "rErroresExpAnecoop.rpt"
                    LlamarImprimir
                End If


            
            End If
        Else
            ' insertamos en la tabla temporal para crear el asiento
            SQLinsert = "insert into tmpinformes (codusu, nombre1, importe1, fecha1, importe2) values "
            conn.Execute SQLinsert & Mid(SqlValues, 1, Len(SqlValues) - 1)
        End If
        Set RS = Nothing
    
        
        ComprobarLiquidacion = (Mens = "")
    
    Else
        MsgBox "No existen pagos pendientes en esta liquidación.", vbExclamation
    End If
    
    lblProgres(0).visible = False
    lblProgres(1).visible = False
    pb1.visible = False
  
    Exit Function

eComprobarLiquidacion:
    MuestraError Err.Number, "Comprobar Liquidación", Err.Description
End Function

Private Sub frmBas_DatoSeleccionado(CadenaSeleccion As String)
'tipos de diario de la Contabilidad
    txtcodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'des macta
    
End Sub

Private Sub frmC_Selec(vFecha As Date)
    txtcodigo(CByte(imgFecha(0).Tag) + 14).Text = Format(vFecha, "dd/mm/yyyy") '<===
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    txtcodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'des macta

End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)

    If CadenaSeleccion <> "" Then
        'Facturas = " anecoop.fra_liq in (" & Mid(CadenaSeleccion, 1, Len(CadenaSeleccion) - 1) & ") "
        Facturas = CadenaSeleccion
    Else
        Facturas = ""
    End If

End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0 ' diario
            indCodigo = 3
            Set frmBas = New frmBasico
            
            frmBas.CadenaTots = "S|txtAux(0)|T|Código|800|;S|txtAux(1)|T|Descripción|3930|;"
            frmBas.CadenaConsulta = "SELECT tiposdiario.numdiari, tiposdiario.desdiari "
            frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM conta" & vParamAplic.NumeroConta & ".tiposdiario "
            frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
            frmBas.Tag1 = "Código|N|N|0|99|tiposdiario|numdiari|00|S|"
            frmBas.Tag2 = "Descripción|T|N|||tiposdiario|desdiari|||"
            frmBas.Maxlen1 = 4
            frmBas.Maxlen2 = 30
            frmBas.tabla = "tiposdiario"
            frmBas.CampoCP = "numdiari"
            'frmBas.Report = "rManCCZonas.rpt"
            frmBas.Caption = "Tipos de Diario"
            frmBas.DeConsulta = True
            frmBas.DatosADevolverBusqueda = "0|1|"
            frmBas.CodigoActual = txtcodigo(3).Text
            frmBas.Show vbModal
            
            Set frmBas = Nothing
        
        
        Case 1 'cuenta contable banco
            AbrirFrmCuentas (Index)
    End Select
    PonerFoco txtcodigo(indCodigo)

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
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    frmC.Left = esq + imgFecha(Index).Parent.Left + 30
    frmC.Top = dalt + imgFecha(Index).Parent.Top + imgFecha(Index).Height + menu - 40

    imgFecha(0).Tag = Index '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtcodigo(Index + 14).Text <> "" Then frmC.NovaData = txtcodigo(Index + 14).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtcodigo(CByte(imgFecha(0).Tag) + 14) '<===
    ' ********************************************
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtcodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 14: KEYFecha KeyAscii, 0 'fecha desde
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFecha_Click (indice)
End Sub
            
Private Sub txtCodigo_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0 ' nro de liquidacion
            CalcularImporte
    
    
        Case 1 ' cta banco
            If txtcodigo(Index).Text <> "" Then txtNombre(Index).Text = PonerNombreCuenta(txtcodigo(Index), 2)
        
        Case 3 ' numero de diario
            txtNombre(Index).Text = DevuelveDesdeBDNew(cConta, "tiposdiario", "desdiari", "numdiari", txtcodigo(Index).Text, "N")
            If txtNombre(Index).Text = "" Then
                MsgBox "Código de diario no existe. Reintroduzca.", vbExclamation
'                PonerFoco txtcodigo(Index)
            End If
        
        Case 14 'FECHAS
            If txtcodigo(Index).Text <> "" Then PonerFormatoFecha txtcodigo(Index)
    End Select
End Sub


Private Sub AbrirFrmCuentas(indice As Integer)
    indCodigo = indice
    Set frmCtas = New frmCtasConta
    frmCtas.DatosADevolverBusqueda = "0|1|"
    frmCtas.CodigoActual = txtcodigo(indCodigo)
    frmCtas.Show vbModal
    Set frmCtas = Nothing
End Sub

Private Sub CalcularImporte()
Dim Sql As String

    txtNombre(0).Text = ""
    If txtcodigo(0).Text = "" Then Exit Sub

    Sql = "select sum(coalesce(importe,0)) from anecoop_pago where num_liquidacion = " & DBSet(txtcodigo(0).Text, "N") & " and idcontab = 0"
    txtNombre(0).Text = Format(DevuelveValor(Sql), "###,###,##0.00")

End Sub
