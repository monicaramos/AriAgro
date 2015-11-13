VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmVtasRevalorarCostes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6840
   Icon            =   "frmVtasRevalorarCostes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
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
      Height          =   4230
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6555
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   4095
         MaxLength       =   10
         TabIndex        =   9
         Tag             =   "Año|N|S|||clientes|codposta|0000||"
         Text            =   "fecha"
         Top             =   1800
         Visible         =   0   'False
         Width           =   1050
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   240
         Left            =   450
         TabIndex        =   7
         Top             =   2790
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1530
         MaxLength       =   4
         TabIndex        =   1
         Tag             =   "Año|N|S|||clientes|codposta|0000||"
         Top             =   1830
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1530
         MaxLength       =   2
         TabIndex        =   0
         Tag             =   "Mes|N|S|1|12|clientes|codposta|00||"
         Top             =   1440
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5025
         TabIndex        =   3
         Top             =   3465
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   2
         Top             =   3465
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Asignar Costes Reales Albaranes"
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
         TabIndex        =   8
         Top             =   495
         Width           =   5160
      End
      Begin VB.Label Label4 
         Caption         =   "Mes"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   15
         Left            =   645
         TabIndex        =   6
         Top             =   1470
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Año"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   14
         Left            =   645
         TabIndex        =   5
         Top             =   1830
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmVtasRevalorarCostes"
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
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim Tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean

Dim FecIni As Date
Dim FecFin As Date

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


    cadSelect = ""
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
    numParam = numParam + 1
    
    cadParam = cadParam & "|pUsu=" & vUsu.Codigo & "|"
    numParam = numParam + 1
    
    FecIni = CDate(txtCodigo(0).Text)
    FecFin = DateAdd("m", 1, FecIni)
    FecFin = FecFin - 1
    
    'D/H Fecha albaran
    cDesde = CStr(FecIni)
    cHasta = CStr(FecFin)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{albaran.fechaalb}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    cadTABLA = "albaran INNER JOIN albaran_costes ON albaran.numalbar = albaran_costes.numalbar "
    
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
        PonerFoco txtCodigo(2)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    Tabla = "albaran"
    
    Pb1.visible = False
    
    txtCodigo(2).Text = Format(Month(Now), "00")
    txtCodigo(3).Text = Format(Year(Now), "0000")
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = w + 70
    Me.Height = h + 350
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
        Case 2 ' mes
            If txtCodigo(Index).Text <> "" Then PonerFormatoEntero txtCodigo(Index)
            
        Case 3 ' año
            If txtCodigo(Index).Text <> "" Then PonerFormatoEntero txtCodigo(Index)
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 4920
        Me.FrameCobros.Width = 6930
        w = Me.FrameCobros.Width
        h = Me.FrameCobros.Height
    End If
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

Private Sub ProcesarCambios(cadWhere As String)
Dim sql As String
Dim Sql2 As String
Dim i As Integer
Dim HayReg As Integer
Dim b As Boolean
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim Mens As String
Dim vSql As String

Dim ImpLinea As Currency
Dim TotContabilidad As Currency
Dim TotAlbaranes As Currency

On Error GoTo eProcesarCambios

    conn.BeginTrans
    
    ' borramos las temporales
    ' en tmpinformes meteremos la diferencia entre impmesde y impmesha de la contabilidad por coste
    ' en tmrinfcostes meteremos la suma de costes de albaran_costes por coste
    
    conn.Execute "delete from tmpinformes where codusu = " & vUsu.Codigo
    
    conn.Execute "delete from tmpinfcostes where codusu = " & vUsu.Codigo
    
    
    sql = "delete from albaran_costreal where numalbar in (select numalbar from albaran where fechaalb >= " & DBSet(FecIni, "F") & " and fechaalb <= " & DBSet(FecFin, "F") & ")"
    conn.Execute sql
    
    sql = "insert into tmpinformes (codusu, campo1, importe1) select " & vUsu.Codigo & ","
    sql = sql & " ctacoste.codcoste, sum(impmesde - impmesha) "
    sql = sql & " from ctacoste, conta" & vParamAplic.NumeroConta & ".hsaldos "
    sql = sql & " where ctacoste.codmacta = conta" & vParamAplic.NumeroConta & ".hsaldos.codmacta and "
    sql = sql & " conta" & vParamAplic.NumeroConta & ".hsaldos.anopsald = " & DBSet(txtCodigo(3).Text, "N") & " and"
    sql = sql & " Conta" & vParamAplic.NumeroConta & ".hsaldos.mespsald = " & DBSet(txtCodigo(2).Text, "N")
    sql = sql & " group by 1,2 "
    
    conn.Execute sql
    
    sql = "insert into tmpinfcostes(codusu, codcoste, importe) select " & vUsu.Codigo & ","
    sql = sql & " albaran_costes.codcoste, sum(albaran_costes.impcoste) from albaran_costes, albaran "
    sql = sql & " where albaran.fechaalb >= " & DBSet(FecIni, "F") & " and albaran.fechaalb <= " & DBSet(FecFin, "F")
    sql = sql & " and albaran_costes.tipogasto = 0 and albaran.numalbar = albaran_costes.numalbar "
    sql = sql & " group by 1,2 "
    
    conn.Execute sql
    
    Set Rs = New ADODB.Recordset
    
    sql = "select count(*) from albaran, albaran_costes where fechaalb >= " & DBSet(FecIni, "F") & " and fechaalb <= " & DBSet(FecFin, "F")
    sql = sql & " and albaran.numalbar = albaran_costes.numalbar and albaran_costes.tipogasto = 0 "
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Pb1.visible = True
    CargarProgres Pb1, Rs.Fields(0).Value
    
    Set Rs = Nothing
    
    Set Rs = New ADODB.Recordset
    
    sql = "select albaran_costes.* from albaran, albaran_costes where fechaalb >= " & DBSet(FecIni, "F") & " and fechaalb <= " & DBSet(FecFin, "F")
    sql = sql & " and albaran.numalbar = albaran_costes.numalbar and albaran_costes.tipogasto = 0 "
    
    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Sql2 = "insert into albaran_costreal(numalbar, numlinea, codcoste, impcoste) values "
    sql = ""
    While Not Rs.EOF
        IncrementarProgres Pb1, 1
        Mens = "Actualizando Costes" & vbCrLf & vbCrLf & "Albaran: " & Rs!numalbar & vbCrLf & "Linea: " & Rs!numlinea
    
        vSql = ""
        vSql = DevuelveDesdeBDNew(cAgro, "tmpinformes", "importe1", "codusu", vUsu.Codigo, "N", , "campo1", Rs!codCoste, "N")
        TotContabilidad = CCur(ComprobarCero(vSql))
        vSql = ""
        vSql = DevuelveDesdeBDNew(cAgro, "tmpinfcostes", "importe", "codusu", vUsu.Codigo, "N", , "codcoste", Rs!codCoste, "N")
        TotAlbaranes = CCur(ComprobarCero(vSql))
        
        ImpLinea = 0
        
        If TotAlbaranes <> 0 Then
            ImpLinea = Round2(TotContabilidad * DBLet(Rs!ImpCoste, "N") / TotAlbaranes, 2)
        End If
    
        sql = sql & "(" & DBSet(Rs!numalbar, "N") & ","
        sql = sql & DBSet(Rs!numlinea, "N") & "," & DBSet(Rs!codCoste, "N") & "," & DBSet(ImpLinea, "N") & "),"
        
        Rs.MoveNext
    Wend
    
    If sql <> "" Then
        'quitamos la ultima coma y ejecutamos
        sql = Mid(sql, 1, Len(sql) - 1)
        
        Sql2 = Sql2 & sql
        
        conn.Execute Sql2
    End If
    
    Rs.Close

eProcesarCambios:
    If Err.Number = 0 Then
        conn.CommitTrans
        MsgBox "Proceso realizado correctamente.", vbExclamation
        cmdCancel_Click
    Else
        conn.RollbackTrans
        MsgBox "Error " & Mens, vbExclamation
    End If
End Sub


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
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim b As Boolean
Dim Fecha As String

    On Error GoTo EDatosOK

    DatosOk = False

    b = False

    If txtCodigo(2).Text = "" Then
        MsgBox "El campo mes debe tener un valor. Reintroduzca.", vbExclamation
        Exit Function
    End If
    If txtCodigo(3).Text = "" Then
        MsgBox "El campo año debe tener un valor. Reintroduzca.", vbExclamation
        Exit Function
    End If

    Fecha = "01/" & Format(CInt(txtCodigo(2).Text), "00") & "/" & Format(CInt(txtCodigo(3).Text), "0000")
    
    txtCodigo(0).Text = Fecha
    
    b = PonerFormatoFecha(txtCodigo(0))


    If Not b Then Exit Function

    DatosOk = b

EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function



