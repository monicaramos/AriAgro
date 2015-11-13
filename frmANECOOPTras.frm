VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmANECOOPTras 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6630
   Icon            =   "frmANECOOPTras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameIntegracion 
      Height          =   4545
      Left            =   30
      TabIndex        =   4
      Top             =   60
      Width           =   6555
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   14
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   0
         Top             =   1890
         Width           =   1005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   15
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   1
         Top             =   2295
         Width           =   1005
      End
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   2820
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   2
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5235
         TabIndex        =   3
         Top             =   3960
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
      Begin VB.Label Label6 
         Caption         =   "Traspaso de Anecoop"
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
         TabIndex        =   13
         Top             =   420
         Width           =   5430
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   21
         Left            =   810
         TabIndex        =   12
         Top             =   1950
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   22
         Left            =   810
         TabIndex        =   11
         Top             =   2265
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Left            =   450
         TabIndex        =   10
         Top             =   1710
         Width           =   435
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1380
         Picture         =   "frmANECOOPTras.frx":000C
         Top             =   1890
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   1380
         Picture         =   "frmANECOOPTras.frx":0097
         Top             =   2295
         Width           =   240
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   2
         Left            =   270
         TabIndex        =   9
         Top             =   3630
         Width           =   6195
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Proceso que realiza la integración de Expedientes Anecoop para asociarlos con los albaranes"
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
         TabIndex        =   7
         Top             =   1020
         Width           =   5820
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   270
         TabIndex        =   6
         Top             =   3390
         Width           =   6195
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   0
         Left            =   270
         TabIndex        =   5
         Top             =   3090
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
Attribute VB_Name = "frmANECOOPTras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: LAURA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public OpcionListado As Byte
    
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


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub cmdAceptar_Click()
Dim Sql As String
Dim I As Byte
Dim cadWhere As String
Dim b As Boolean
Dim NomFic As String
Dim cadena As String
Dim cadena1 As String
Dim Directorio As String
Dim fec As String
Dim nomDir As String

Dim nRegs As Long
Dim cadTABLA As String
Dim NomFic1 As String

Dim File1 As FileSystemObject

'On Error GoTo eError


    If Not DatosOk Then Exit Sub
    
    cmdAceptar.Enabled = False
    cmdCancelar.Enabled = False

    If CargarExpedientes Then
    
        If Not CambiarFechasANull Then Exit Sub
    
        If AsociacionExpedientes Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
            '========= PARAMETROS  =============================
            'Añadir el parametro de Empresa
            cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
            numParam = numParam + 1
            
            cadTABLA = "tmpinformes"
            cadFormula = "{tmpinformes.codusu} = " & vUsu.codigo

            Sql = "select count(*) from tmpinformes where codusu = " & vUsu.codigo

            If TotalRegistros(Sql) <> 0 Then
                MsgBox "Hay errores en la Asignación a Albaranes." & vbCrLf, vbExclamation
                cadTitulo = "Errores de asignación a Albaranes"
                cadNombreRPT = "rErroresAnecoop.rpt"
                LlamarImprimir
            Else
                MsgBox "Proceso realizado correctamente.", vbExclamation
            End If
        
        End If
        
    End If
        
    cmdAceptar.Enabled = True
    cmdCancelar.Enabled = True
    
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""
    lblProgres(2).Caption = ""
                
    Unload Me
    
End Sub


Private Function CambiarFechasANull() As Boolean
Dim Sql As String

    On Error GoTo eCambiarFechasANull

    CambiarFechasANull = False

    Sql = "update anecoop set "
    Sql = Sql & "  fecha_liq = if(mid(fecha_liq,1,4) < 1900,null, fecha_liq), "
    Sql = Sql & "  fecha_cambio_liq = if(mid(fecha_cambio_liq,1,4) < 1900,null, fecha_cambio_liq), "
    Sql = Sql & "  fecha_sc_liq = if(mid(fecha_sc_liq,1,4) < 1900,null, fecha_sc_liq) "
    
    conn.Execute Sql

    Sql = "update anecoop_pago set "
    Sql = Sql & "  fecha_factura = if(mid(fecha_factura,1,4) < 1900,null, fecha_factura), "
    Sql = Sql & "  fecha_pago = if(mid(fecha_pago,1,4) < 1900,null, fecha_pago), "
    Sql = Sql & "  fecha_pago_sc = if(mid(fecha_pago_sc,1,4) < 1900,null, fecha_pago_sc) "
    
    conn.Execute Sql


    CambiarFechasANull = True
    Exit Function

eCambiarFechasANull:
    MuestraError Err.Number, "Cambiar Fechas Erroneas", Err.Description
End Function

Private Function AsociacionExpedientes() As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim numlinea As String
Dim Albaran As String
Dim Rs As ADODB.Recordset
Dim MenError As String
Dim PorPedido As Boolean

    On Error GoTo eAsociacionExpedientes

    AsociacionExpedientes = False
    
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.codigo
    conn.Execute Sql

    Sql = "Select * from anecoop where (numlinea is null or numlinea = 0) and fecha_salida between " & DBSet(txtCodigo(14).Text, "F") & " and " & DBSet(txtCodigo(15).Text, "F")
    Sql = Sql & " and nombre_variedad <> '' "
    
    If TotalRegistrosConsulta(Sql) = 0 Then
        MsgBox "No hay registros pendientes de asociar a albaranes"
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    Me.lblProgres(0).Caption = "Asociando Expedientes con Albaranes"
    DoEvents

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Me.lblProgres(1).Caption = "Expediente: " & Rs!expediente_id
        DoEvents
        
        MenError = ""
        
        If Not IsNumeric(Rs!numero_salida_cooperativa) Then
            MenError = "Albarán " & Rs!numero_salida_cooperativa & " albarán no numérico."
        Else
            Sql3 = "select numalbar from albaran where numalbar = " & DBSet(Rs!numero_salida_cooperativa, "N")
            Albaran = TotalRegistrosConsulta(Sql3)
            If Albaran = 0 Then ' si hay mas de un albaran o no hay albaran, error
                 MenError = "No hay albarán asociado"
            End If
        End If
        If MenError <> "" Then
            Sql3 = "insert into tmpinformes (codusu, nombre1, fecha1, nombre2, importe1, importe2) values (" & vUsu.codigo & ","
            Sql3 = Sql3 & DBSet(Rs!numero_salida_cooperativa, "T") & ","
            Sql3 = Sql3 & DBSet(Rs!fecha_salida, "F") & ","
            Sql3 = Sql3 & DBSet(MenError, "T") & ","
            Sql3 = Sql3 & DBSet(Rs!ncajas, "N") & ","
            Sql3 = Sql3 & DBSet(Rs!peso_neto, "N") & ")"
            
            conn.Execute Sql3
        Else
            Albaran = DevuelveValor(Sql3)
            
            Sql2 = "select numlinea from albaran_variedad where  numcajas = " & DBSet(Rs!ncajas, "N")
            '[Monica]13/05/2015: añado la condicion de que no tenga expediente insertado
            Sql2 = Sql2 & " and numalbar = " & DBSet(Albaran, "N") & " and (expediente is null or expediente = '') "
            numlinea = DevuelveValor(Sql2)
            
            If numlinea = 0 Then ' error
                MenError = "No hay línea de albarán asociado"
                
                Sql3 = "insert into tmpinformes (codusu, nombre1, fecha1, nombre2, importe1, importe2) values (" & vUsu.codigo & ","
                Sql3 = Sql3 & DBSet(Rs!numero_salida_cooperativa, "T") & ","
                Sql3 = Sql3 & DBSet(Rs!fecha_salida, "F") & ","
                Sql3 = Sql3 & DBSet(MenError, "T") & ","
                Sql3 = Sql3 & DBSet(Rs!ncajas, "N") & ","
                Sql3 = Sql3 & DBSet(Rs!peso_neto, "N") & ")"
                
                conn.Execute Sql3
            Else
                Sql3 = "update anecoop set numlinea = " & DBSet(numlinea, "N")
                Sql3 = Sql3 & " where expediente_id = " & DBSet(Rs!expediente_id, "T")
                Sql3 = Sql3 & " and linea_expediente = " & DBSet(Rs!linea_expediente, "T")
                Sql3 = Sql3 & " and codigo_campanya = " & DBSet(Rs!codigo_campanya, "T")
                
                conn.Execute Sql3
                
                '[Monica]13/05/2015: añado la actualizacion de la linea de albaran_variedad
                Sql3 = "update albaran_variedad set expediente = " & DBSet(Rs!expediente_id, "T")
                Sql3 = Sql3 & " where numalbar = " & DBSet(Albaran, "N")
                Sql3 = Sql3 & " and numlinea = " & DBSet(numlinea, "N")
                
                conn.Execute Sql3
                
            End If
        End If
    
        Rs.MoveNext
    Wend
    
    AsociacionExpedientes = True
    
    Screen.MousePointer = vbDefault
    Exit Function
           
eAsociacionExpedientes:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Asociación Expedientes", Err.Description
End Function



Private Function CargarExpedientes() As Boolean
Dim Sql As String
Dim Anyo As String
Dim temp As Boolean
Dim I As Integer

    On Error GoTo Error2
    
    CargarExpedientes = False

    If Dir(App.path & "\ConAnecoop.exe") = "" Then
        MsgBox "El programa de Carga de Expedientes no existe. Revise.", vbExclamation
    Else
        If Dir(App.path & "\aneccop.z") = "" Then
            MsgBox "El proceso de carga debe de estar realizándose.", vbExclamation
        Else
            Sql = "Se va a proceder a realizar la carga de Expedientes Anecoop. " & vbCrLf & vbCrLf & "¿ Desea continuar ?"
            If MsgBox(Sql, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            
                '------------------------------------------------------------------------------
                '  LOG de acciones
                Set LOG = New cLOG
                LOG.Insertar 11, vUsu, "Inserción de Expedientes de Anecoop: " & vbCrLf & vUsu.codigo & vbCrLf & Now
                Set LOG = Nothing
                '-----------------------------------------------------------------------------
                     
                ' eliminamos el registro chivato
                Kill App.path & "\aneccop.z"
                    
                Anyo = Mid(CStr(Year(vParam.FecIniCam)), 3, 2)
                    
                Shell App.path & "\ConAnecoop " & txtCodigo(14).Text & " " & txtCodigo(15).Text & " " & Anyo & " v ", vbNormalFocus
                
                Screen.MousePointer = vbHourglass
                
                I = 0
                While Dir(App.path & "\aneccop.z") = "" And I < 300
                    Me.lblProgres(0).Caption = "Procesando Insercion "
                    DoEvents
                    
                    espera 1
                    
                    I = I + 1
                Wend
                
                
                If Dir(App.path & "\aneccop.z") <> "" Then CargarExpedientes = True
                
                
            End If
        End If
    End If
    
    Screen.MousePointer = vbDefault
    Exit Function
    
    
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Cargar Expedientes", Err.Description
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
Dim h As Integer, w As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me
    
    ConSubInforme = False

    
    'Ocultar todos los Frames de Formulario
    FrameIntegracion.visible = False
    '###Descomentar
'    CommitConexion
        
    FrameIntegracionVisible True, h, w
    pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
'    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = w + 70
    Me.Height = h + 350
End Sub


Private Sub FrameIntegracionVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para el listado de socios por seccion
    Me.FrameIntegracion.visible = visible
    If visible = True Then
        Me.FrameIntegracion.Top = -90
        Me.FrameIntegracion.Left = 0
        Me.FrameIntegracion.Height = 4665
        Me.FrameIntegracion.Width = 6555
        w = Me.FrameIntegracion.Width
        h = Me.FrameIntegracion.Height
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
        .Opcion = OpcionListado
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


Private Function ComprobarErrores(ByRef pb1 As ProgressBar) As Boolean
Dim NF As Long
Dim cad As String
Dim I As Integer
Dim Longitud As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim Numreg As Long
Dim Sql As String
Dim SQL1 As String
Dim total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean
Dim Mens As String
Dim Tipo As Integer
Dim FechaEnt As String
Dim Variedad As String


    On Error GoTo eComprobarErrores

    ComprobarErrores = False
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.codigo
    conn.Execute Sql

    I = 0
    lblProgres(1).Caption = "Comprobando errores Tabla temporal entradas "
    
    Sql = "select * from tmpentradaS"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    b = True
    I = 0
    While Not Rs.EOF And b
        I = I + 1

        Me.pb1.Value = Me.pb1.Value + 1
        lblProgres(2).Caption = "Linea " & I
        Me.Refresh

        Variedad = Format(Rs!codprodu, "00") & Format(Rs!codvarie, "00")

        ' comprobamos la fecha
        FechaEnt = DBLet(Rs!FechaEnt, "T")
        If Not EsFechaOK(FechaEnt) Then
            Mens = "Fecha incorrecta"
            Sql = "insert into tmpinformes (codusu, campo1, codigo1, importe1, importe2, fecha1, nombre1) values (" & _
                  vUsu.codigo & "," & DBSet(Variedad, "N") & "," & DBSet(Rs!CodSocio, "N") & "," & _
                  DBSet(Rs!codCampo, "N") & "," & DBSet(Rs!Numnotac, "N") & "," & _
                  DBSet(FechaEnt, "F") & "," & DBSet(Mens, "T") & ")"
            conn.Execute Sql
        End If


        ' comprobamos que exista el socio
        Sql = "select count(*) from rsocios where codsocio = " & DBSet(Rs!CodSocio, "N")
        If TotalRegistros(Sql) = 0 Then
            Mens = "Socio no existe"
            Sql = "insert into tmpinformes (codusu, campo1, codigo1, importe1, importe2, fecha1, nombre1) values (" & _
                  vUsu.codigo & "," & DBSet(Variedad, "N") & "," & DBSet(Rs!CodSocio, "N") & "," & _
                  DBSet(Rs!codCampo, "N") & "," & DBSet(Rs!Numnotac, "N") & "," & _
                  DBSet(FechaEnt, "F") & "," & DBSet(Mens, "T") & ")"
            conn.Execute Sql
        End If

        ' comprobamos que exista el campo
        Sql = "select count(*) from rcampos where codsocio = " & DBSet(Rs!CodSocio, "N")
        Sql = Sql & " and nrocampo = " & DBSet(Rs!codCampo, "N")
        Sql = Sql & " and codvarie = " & DBSet(Variedad, "N")
        Sql = Sql & " and fecbajas is null "
        If TotalRegistros(Sql) = 0 Then
            Mens = "Campo no existe o con fecha de baja"
            Sql = "insert into tmpinformes (codusu, campo1, codigo1, importe1, importe2, fecha1, nombre1) values (" & _
                  vUsu.codigo & "," & DBSet(Variedad, "N") & "," & DBSet(Rs!CodSocio, "N") & "," & _
                  DBSet(Rs!codCampo, "N") & "," & DBSet(Rs!Numnotac, "N") & "," & _
                  DBSet(FechaEnt, "F") & "," & DBSet(Mens, "T") & ")"
            conn.Execute Sql
        End If

        ' comprobamos que no exista mas de un campo con ese numero de orden campo (scampo.codcampo MB)
        Sql = "select count(*) from rcampos where codsocio = " & DBSet(Rs!CodSocio, "N")
        Sql = Sql & " and nrocampo = " & DBSet(Rs!codCampo, "N")
        Sql = Sql & " and codvarie = " & DBSet(Variedad, "N")
        If TotalRegistros(Sql) > 1 Then
            Mens = "Campo con más de un registro"
            Sql = "insert into tmpinformes (codusu, campo1, codigo1, importe1, importe2, fecha1, nombre1) values (" & _
                  vUsu.codigo & "," & DBSet(Variedad, "N") & "," & DBSet(Rs!CodSocio, "N") & "," & _
                  DBSet(Rs!codCampo, "N") & "," & DBSet(Rs!Numnotac, "N") & "," & _
                  DBSet(FechaEnt, "F") & "," & DBSet(Mens, "T") & ")"
            conn.Execute Sql
        End If

        
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""
    lblProgres(2).Caption = ""

    ComprobarErrores = b
    Exit Function

eComprobarErrores:
    ComprobarErrores = False
End Function


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Sql As String
Dim Sql2 As String
Dim vClien As CSocio
' añadido
Dim Mens As String
Dim NumFactu As String
Dim numser As String
Dim fecha As Date
Dim vCont As CTiposMov
Dim tipoMov As String

    b = True
    
    If txtCodigo(14).Text = "" Or txtCodigo(15) = "" Then
        MsgBox "Debe de introducir las fechas de trapaso. Reintroduzca.", vbExclamation
        b = False
        PonerFoco txtCodigo(14)
    End If
    
    DatosOk = b

End Function


Private Sub frmC_Selec(vFecha As Date)
    txtCodigo(CByte(imgFecha(0).Tag) + 14).Text = Format(vFecha, "dd/mm/yyyy") '<===
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
    If txtCodigo(Index + 14).Text <> "" Then frmC.NovaData = txtCodigo(Index + 14).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtCodigo(CByte(imgFecha(0).Tag) + 14) '<===
    ' ********************************************
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 14: KEYFecha KeyAscii, 0 'fecha desde
            Case 15: KEYFecha KeyAscii, 1 'fecha hasta
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
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 14, 15 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
    End Select
End Sub

