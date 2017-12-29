VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmANECOOPTrasFras 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6630
   Icon            =   "frmANECOOPTrasFras.frx":0000
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
         Caption         =   "Traspaso de Facturas Anecoop"
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
         Picture         =   "frmANECOOPTrasFras.frx":000C
         Top             =   1890
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   1380
         Picture         =   "frmANECOOPTrasFras.frx":0097
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
         Caption         =   "Proceso que realiza la creación de Facturas Anecoop."
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
Attribute VB_Name = "frmANECOOPTrasFras"
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


'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadselect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe
Private cadSelect1 As String 'Cadena para comprobar si hay datos antes de abrir Informe


Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Tabla1 As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
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
Dim TipoIvac As Byte
Dim Dto1 As Currency
Dim Dto2 As Currency


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub cmdAceptar_Click()
Dim SQL As String
Dim i As Byte
Dim cadWHERE As String
Dim b As Boolean
Dim NomFic As String
Dim CADENA As String
Dim cadena1 As String
Dim Directorio As String
Dim fec As String
Dim nomDir As String

Dim Nregs As Long
Dim cadTABLA As String
Dim NomFic1 As String

Dim File1 As FileSystemObject

'On Error GoTo eError

    If Not DatosOk Then Exit Sub
    
    cmdAceptar.Enabled = False
    cmdCancelar.Enabled = False

    SQL = " not anecoop.fra_liq is null and anecoop.fra_liq <> '' " 'and not numlinea is null and numlinea <> ''"
    
    If txtCodigo(14).Text <> "" Then SQL = SQL & " and anecoop.fecha_liq >= " & DBSet(txtCodigo(14).Text, "F")
    If txtCodigo(15).Text <> "" Then SQL = SQL & " and anecoop.fecha_liq <= " & DBSet(txtCodigo(15).Text, "F")

    '[Monica]18/12/2017: nro de fra A000000X
    SQL = SQL & " and not mid(fra_liq,2,7) in (select numfactu "
    SQL = SQL & " from facturas where codtipom = " & DBSet(CodtipomAnecoop, "T")
    
    If txtCodigo(14).Text <> "" Then SQL = SQL & " and facturas.fecfactu >= " & DBSet(txtCodigo(14).Text, "F")
    If txtCodigo(15).Text <> "" Then SQL = SQL & " and facturas.fecfactu <= " & DBSet(txtCodigo(15).Text, "F")
        
    SQL = SQL & ") "

    Facturas = ""


    Set frmMens = New frmMensajes
    
    frmMens.OpcionMensaje = 29
    frmMens.cadWHERE = SQL
    frmMens.Show vbModal
    
    Set frmMens = Nothing


    If Facturas <> "" Then
    
        If Not ComprobarDesdobles Then
        
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            cadTitulo = "Expedientes sin Importe Factura"
            cadNombreRPT = "rErroresExpAnecoop1.rpt"
            LlamarImprimir
            
            cmdAceptar.Enabled = True
            cmdCancelar.Enabled = True
            
            lblProgres(0).Caption = ""
            lblProgres(1).Caption = ""
            lblProgres(2).Caption = ""
        
            Exit Sub
        End If
        
        If CreacionFacturasAnecoop(SQL) Then
            
            MsgBox "Proceso realizado correctamente.", vbExclamation
            
            '========= PARAMETROS  =============================
            'Añadir el parametro de Empresa
            cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
            numParam = numParam + 1
            
            cadTABLA = "tmpinformes"
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo

            SQL = "select count(*) from tmpinformes where codusu = " & vUsu.Codigo

            If TotalRegistros(SQL) <> 0 Then
                cadTitulo = "Facturas de Anecoop creadas"
                cadNombreRPT = "rFacturasAnecoop.rpt"
                LlamarImprimir
            End If
        
        End If
    Else
    
        MsgBox "No se ha creado ninguna factura.", vbExclamation
        
    End If
        
    cmdAceptar.Enabled = True
    cmdCancelar.Enabled = True
    
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""
    lblProgres(2).Caption = ""
                
    Unload Me
    
End Sub


Private Function ComprobarDesdobles()
Dim SQL As String

    On Error GoTo eComprobarDesdobles
        
    ComprobarDesdobles = False
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    SQL = " insert into tmpinformes (codusu, nombre1) "
    SQL = SQL & " select distinct " & vUsu.Codigo & ", dd.expediente_id from anecoop dd, anecoop ff "
    SQL = SQL & " where mid(dd.expediente_id,1,1) <> '0' and (dd.importe_liq = 0 or dd.importe_liq is null)"
    SQL = SQL & " and mid(dd.expediente_id,2,17) = right(concat('000000000000000000',ff.expediente_id),17)"
    SQL = SQL & " and ff.fra_liq in (" & Facturas & ") "
    
    conn.Execute SQL
    
    SQL = "select count(*) from tmpinformes where codusu = " & vUsu.Codigo
    
    ComprobarDesdobles = (TotalRegistros(SQL) = 0)
    Exit Function


eComprobarDesdobles:
    MuestraError Err.Number, "Comprobar Desdobles", Err.Description
End Function

Private Function CreacionFacturasAnecoop(vWhere As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim NumLinea As String
Dim Albaran As String
Dim Rs As ADODB.Recordset
Dim RSLineas As ADODB.Recordset
Dim RSEnvases As ADODB.Recordset
Dim MenError As String
Dim FacturaAnt As String
Dim FechaAnt As Date
Dim TipoIva As String
Dim Bruto As Currency
Dim Total As Currency
Dim HayReg As Byte
Dim i As Integer
Dim SqlValues As String
Dim SqlValues2 As String
Dim sqlLineas As String
Dim SqlEnvases As String
Dim vCStock As CStock
Dim b As Boolean

    On Error GoTo eCreacionFacturasAnecoop

    CreacionFacturasAnecoop = False

    conn.BeginTrans

    conn.Execute "set foreign_key_checks = 0 "

    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    '[Monica]18/12/2017:nro de fra A000000X
    sqlLineas = "select mid(fra_liq,2,7) numfactu, fecha_liq fecfactu, numero_salida_cooperativa, numlinea, porcent_iva_liq, sum(importe_liq) importe_liq, sum(ncajas) ncajas,  sum(peso_neto) peso_neto, if ( sum(peso_neto) is null or sum(peso_neto) = 0,0, round(sum(importe_liq) / sum(peso_neto),4))  precio_comercial from anecoop  "
    sqlLineas = sqlLineas & " where " & vWhere
    sqlLineas = sqlLineas & " and mid(fra_liq,2,7) in (" & Facturas & ") "
    sqlLineas = sqlLineas & " and not numlinea is null and nombre_variedad <> '' "
    sqlLineas = sqlLineas & " group by 1, 2, 3, 4, 5 "
    sqlLineas = sqlLineas & " order by 1, 2, 3, 4, 5 "
    
    Set RSLineas = New ADODB.Recordset
    RSLineas.Open sqlLineas, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    '[Monica]13/05/2015: no agrupamos por nada cada linea tiene que tener su linea de envase
    'SqlEnvases = "select fra_liq numfactu, fecha_liq fecfactu, numero_salida_cooperativa, numlinea, porcent_iva_liq, sum(importe_liq) importe_liq, sum(ncajas) ncajas,  sum(peso_neto) peso_neto, if (sum(peso_neto) is null or sum(peso_neto) = 0,0,round(sum(importe_liq) / sum(peso_neto) ,4)) precio_comercial from anecoop  "
                    '[Monica]18/12/2017:nro de fra A000000X
    SqlEnvases = "select mid(fra_liq,2,7) numfactu, fecha_liq fecfactu, numero_salida_cooperativa, numlinea, porcent_iva_liq, importe_liq importe_liq, ncajas ncajas,  peso_neto peso_neto, if (peso_neto is null or peso_neto = 0,0,round(importe_liq / peso_neto ,4)) precio_comercial from anecoop  "
    SqlEnvases = SqlEnvases & " where " & vWhere
    SqlEnvases = SqlEnvases & " and mid(fra_liq,2,7) in (" & Facturas & ") "
    SqlEnvases = SqlEnvases & " and numlinea is null and (nombre_variedad = '' or nombre_variedad is null) "
    'SqlEnvases = SqlEnvases & " group by 1, 2, 3, 4, 5 "
    SqlEnvases = SqlEnvases & " order by 1, 2, 3, 4, 5 "
    
    Set RSEnvases = New ADODB.Recordset
    RSEnvases.Open SqlEnvases, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    
    'CABECERAS DE FACTURAS
    '======================

    ' Cargamos el recordset para actualizar segun el tipo de iva las cabeceras
    SQL = "select mid(fra_liq,2,7) as numfactu, fecha_liq as fecfactu, porcent_iva_liq as porciva, sum(importe_liq) as importe_liq, sum(importe_iva_liq) as importe_iva_liq,  sum(importe_iva_liq + importe_liq) as total from anecoop  "
    SQL = SQL & " where " & vWhere
    SQL = SQL & " and mid(fra_liq,2,7) in (" & Facturas & ") "
    SQL = SQL & " group by 1, 2, 3 "
    SQL = SQL & " order by 1, 2, 3 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Insertamos todas las cabeceras de las facturas
    Sql3 = "select mid(fra_liq,2,7) numfactu, fecha_liq fecfactu, sum(importe_liq) importe_liq, sum(importe_iva_liq) importe_iva_liq,  sum(importe_iva_liq + importe_liq) total from anecoop  "
    Sql3 = Sql3 & " where " & vWhere
    Sql3 = Sql3 & " and mid(fra_liq,2,7) in (" & Facturas & ") "
    Sql3 = Sql3 & " group by 1, 2 "
    Sql3 = Sql3 & " order by 1, 2 "
    
    Sql2 = "insert into facturas (codtipom, numfactu, fecfactu, codclien, codforpa, tipoivac) "
    Sql2 = Sql2 & " select " & DBSet(CodtipomAnecoop, "T") & ", numfactu, fecfactu, " & DBSet(vParamAplic.CodAnecoop, "N") & "," & DBSet(Codforpa, "N") & "," & DBSet(TipoIvac, "N") & " from (" & Sql3 & ") aaaa"
    conn.Execute Sql2
    
    If Not Rs.EOF Then
        FacturaAnt = DBLet(Rs!NumFactu, "N")
        FechaAnt = CDate(DBLet(Rs!FecFactu))
    End If
    
    i = 0
    Bruto = 0
    Total = 0
    HayReg = 0
    
    While Not Rs.EOF
        HayReg = 1
        If FacturaAnt <> Rs!NumFactu Or FechaAnt <> Rs!FecFactu Then
            Sql2 = "update facturas set "
            Sql2 = Sql2 & " brutofac = " & DBSet(Bruto, "N")
            Sql2 = Sql2 & ", totalfac = " & DBSet(Total, "N")
            Sql2 = Sql2 & " where codtipom = " & DBSet(CodtipomAnecoop, "T")
            Sql2 = Sql2 & " and numfactu = " & DBSet(FacturaAnt, "N")
            Sql2 = Sql2 & " and fecfactu =" & DBSet(FechaAnt, "F")
            
            conn.Execute Sql2
            
            
            ' insertamos en la tabla temporal la factura que hemos insertado
            SQL = "insert into tmpinformes (codusu, nombre1, importe1, fecha1, importe2) values "
            SQL = SQL & "(" & vUsu.Codigo & "," & DBSet(CodtipomAnecoop, "T") & "," & DBSet(FacturaAnt, "N") & ","
            SQL = SQL & DBSet(FechaAnt, "F") & "," & DBSet(Total, "N") & ")"
            conn.Execute SQL
            
            i = 0
            Bruto = 0
            Total = 0
            FacturaAnt = DBLet(Rs!NumFactu)
            FechaAnt = DBLet(Rs!FecFactu)
        End If
        
        TipoIva = DevuelveDesdeBDNew(cConta, "tiposiva", "codigiva", "porceiva", DBLet(Rs!PorcIva), "N")
        
        If TipoIva = "" Then
            MsgBox "No existe el tipo de iva " & DBLet(Rs!PorcIva) & " en contabilidad. Revise.", vbExclamation
            conn.RollbackTrans
            Exit Function
        End If
        
        Bruto = Bruto + DBLet(Rs!importe_liq)
        Total = Total + DBLet(Rs!Total, "N")
        
        ' actualizamos bases
        SQL = "update facturas set "
        Select Case i
            Case 0
                SQL = SQL & " baseimp1 = " & DBSet(Rs!importe_liq, "N")
                SQL = SQL & ",impoiva1 = " & DBSet(Rs!importe_iva_liq, "N")
                SQL = SQL & ",porciva1 = " & DBSet(Rs!PorcIva, "N")
                SQL = SQL & ",codiiva1 = " & DBSet(TipoIva, "N")
            Case 1
                SQL = SQL & " baseimp2 = " & DBSet(Rs!importe_liq, "N")
                SQL = SQL & ",impoiva2 = " & DBSet(Rs!importe_iva_liq, "N")
                SQL = SQL & ",porciva2 = " & DBSet(Rs!PorcIva, "N")
                SQL = SQL & ",codiiva2 = " & DBSet(TipoIva, "N")
            Case 2
                SQL = SQL & " baseimp3 = " & DBSet(Rs!importe_liq, "N")
                SQL = SQL & ",impoiva3 = " & DBSet(Rs!importe_iva_liq, "N")
                SQL = SQL & ",porciva3 = " & DBSet(Rs!PorcIva, "N")
                SQL = SQL & ",codiiva3 = " & DBSet(TipoIva, "N")
        End Select
    
        SQL = SQL & " where codtipom = " & DBSet(CodtipomAnecoop, "T")
        SQL = SQL & " and numfactu = " & DBSet(Rs!NumFactu, "N")
        SQL = SQL & " and fecfactu =" & DBSet(Rs!FecFactu, "F")
        
        conn.Execute SQL
    
        i = i + 1
    
        Rs.MoveNext
    Wend
     
    ' en el ultimo registro actualizamos totales
    If HayReg Then
        Sql2 = "update facturas set "
        Sql2 = Sql2 & " brutofac = " & DBSet(Bruto, "N")
        Sql2 = Sql2 & ", totalfac = " & DBSet(Total, "N")
        Sql2 = Sql2 & " where codtipom = " & DBSet(CodtipomAnecoop, "T")
        Sql2 = Sql2 & " and numfactu = " & DBSet(FacturaAnt, "N")
        Sql2 = Sql2 & " and fecfactu =" & DBSet(FechaAnt, "F")
        
        conn.Execute Sql2
        
        ' insertamos en la tabla temporal la factura que hemos insertado
        SQL = "insert into tmpinformes (codusu, nombre1, importe1, fecha1, importe2) values "
        SQL = SQL & "(" & vUsu.Codigo & "," & DBSet(CodtipomAnecoop, "T") & "," & DBSet(FacturaAnt, "N") & ","
        SQL = SQL & DBSet(FechaAnt, "F") & "," & DBSet(Total, "N") & ")"
        
        conn.Execute SQL
        
    End If

    'LINEAS DE FACTURAS
    '==================
    'Insertamos las lineas de facturas por albaran
    If Not RSLineas.EOF Then
        SQL = "insert into facturas_variedad (codtipom,numfactu,fecfactu,numlinea,numalbar,numlinealbar,cantreal,cantfact,precibru,precinet,"
        SQL = SQL & "imporbru,impornet,codigiva,unidades) values "
        
        Sql2 = "insert into facturas_calibre (codtipom,numfactu,fecfactu,numlinea,numline1,numalbar,numlinealbar,numline1albar,cantreal,cantfact,"
        Sql2 = Sql2 & "precibru,precinet,imporbru,impornet,codigiva,unidades) values "
        
        SqlValues = ""
        SqlValues2 = ""
        
        i = 0
        FacturaAnt = RSLineas!NumFactu
        FechaAnt = RSLineas!FecFactu
        While Not RSLineas.EOF
            If FacturaAnt <> RSLineas!NumFactu Or FechaAnt <> RSLineas!FecFactu Then
                i = 0
            End If
            i = i + 1
            TipoIva = DevuelveDesdeBDNew(cConta, "tiposiva", "codigiva", "porceiva", CStr(DBLet(RSLineas!porcent_iva_liq)), "N")
            
            If TipoIva = "" Then
                MsgBox "No existe el tipo de iva " & DBLet(Rs!PorcIva) & " en contabilidad. Revise.", vbExclamation
                conn.RollbackTrans
                Exit Function
            End If
        
            SqlValues = SqlValues & "(" & DBSet(CodtipomAnecoop, "T") & "," & DBSet(RSLineas!NumFactu, "N") & "," & DBSet(RSLineas!FecFactu, "F") & "," & DBSet(i, "N") & ","
            SqlValues = SqlValues & DBSet(RSLineas!numero_salida_cooperativa, "N") & "," & DBSet(RSLineas!NumLinea, "N") & "," & DBSet(RSLineas!peso_neto, "N") & "," & DBSet(RSLineas!ncajas, "N") & ","
            SqlValues = SqlValues & DBSet(RSLineas!precio_comercial, "N") & "," & DBSet(RSLineas!precio_comercial, "N") & "," & DBSet(RSLineas!importe_liq, "N") & "," & DBSet(RSLineas!importe_liq, "N") & ","
            SqlValues = SqlValues & DBSet(TipoIva, "N") & ",0),"
        
        
            SqlValues2 = SqlValues2 & "(" & DBSet(CodtipomAnecoop, "T") & "," & DBSet(RSLineas!NumFactu, "N") & "," & DBSet(RSLineas!FecFactu, "F") & "," & DBSet(i, "N") & ",1,"
            SqlValues2 = SqlValues2 & DBSet(RSLineas!numero_salida_cooperativa, "N") & "," & DBSet(RSLineas!NumLinea, "N") & ",1," & DBSet(RSLineas!peso_neto, "N") & "," & DBSet(RSLineas!ncajas, "N") & ","
            SqlValues2 = SqlValues2 & DBSet(RSLineas!precio_comercial, "N") & "," & DBSet(RSLineas!precio_comercial, "N") & "," & DBSet(RSLineas!importe_liq, "N") & "," & DBSet(RSLineas!importe_liq, "N") & ","
            SqlValues2 = SqlValues2 & DBSet(TipoIva, "N") & ",0),"
        
        
            RSLineas.MoveNext
        Wend
        If SqlValues <> "" Then
            SqlValues = Mid(SqlValues, 1, Len(SqlValues) - 1)
            conn.Execute SQL & SqlValues
            
            SqlValues2 = Mid(SqlValues2, 1, Len(SqlValues2) - 1)
            conn.Execute Sql2 & SqlValues2
            
        End If
    End If
    Set RSLineas = Nothing
    
    
    b = True
    
    'LINEAS DE ENVASES
    '==================
    'Insertamos las lineas de facturas por albaran
    If Not RSEnvases.EOF Then
    
        SQL = "insert into facturas_envases (codtipom,numfactu,fecfactu,numlinea,codalmac,codartic,cantidad,precioar,"
        SQL = SQL & "dtolinea,importel,ampliaci,codigiva, numalbar) values "
        
        SqlValues = ""
        
        i = 0
        
        FacturaAnt = RSEnvases!NumFactu
        FechaAnt = RSEnvases!FecFactu
        While Not RSEnvases.EOF And b
            If FacturaAnt <> RSEnvases!NumFactu Or FechaAnt <> RSEnvases!FecFactu Then
                i = 0
            End If
            i = i + 1
            TipoIva = DevuelveDesdeBDNew(cConta, "tiposiva", "codigiva", "porceiva", CStr(DBLet(RSEnvases!porcent_iva_liq)), "N")
            
            If TipoIva = "" Then
                MsgBox "No existe el tipo de iva " & DBLet(Rs!PorcIva) & " en contabilidad. Revise.", vbExclamation
                conn.RollbackTrans
                Exit Function
            End If
        
            SqlValues = SqlValues & "(" & DBSet(CodtipomAnecoop, "T") & "," & DBSet(RSEnvases!NumFactu, "N") & "," & DBSet(RSEnvases!FecFactu, "F") & "," & DBSet(i, "N") & ","
            SqlValues = SqlValues & DBSet(vParamAplic.Almacen, "N") & "," & DBSet(vParamAplic.EnvAnecoop, "T") & "," & DBSet(RSEnvases!peso_neto, "N") & ","
            SqlValues = SqlValues & DBSet(RSEnvases!precio_comercial, "N") & ",0," & DBSet(RSEnvases!importe_liq, "N") & "," & ValorNulo & ","
            SqlValues = SqlValues & DBSet(TipoIva, "N") & "," & DBSet(RSEnvases!numero_salida_cooperativa, "N") & "),"
        
        
            'Tenemos que modificar el stock de los albaranes
            Set vCStock = New CStock
            
            vCStock.tipoMov = "S"
            vCStock.DetaMov = CodtipomAnecoop
            
            vCStock.Trabajador = CInt(vParamAplic.CodAnecoop)  'guardamos el cliente de la factura
            
            '[Monica]20/03/2012: guardamos el numero de albaran o de factura (dependiendo de de donde viene)
            vCStock.Documento = CLng(RSEnvases!NumFactu) 'Nº Factura
            vCStock.Fechamov = DBLet(RSEnvases!FecFactu) 'Fecha de la Factura
            
            vCStock.CodArtic = vParamAplic.EnvAnecoop
            vCStock.codAlmac = CInt(vParamAplic.Almacen)
            vCStock.Cantidad = DBLet(RSEnvases!peso_neto)
            vCStock.Importe = DBLet(RSEnvases!importe_liq)
            vCStock.LineaDocu = i
        
            b = vCStock.ActualizarStock(True)
    
            Set vCStock = Nothing
        
            RSEnvases.MoveNext
        Wend
        If SqlValues <> "" Then
            SqlValues = Mid(SqlValues, 1, Len(SqlValues) - 1)
            conn.Execute SQL & SqlValues
        End If
    End If
    
    Set RSEnvases = Nothing
    
    If b Then
        conn.CommitTrans
        CreacionFacturasAnecoop = True
        Exit Function
    End If

eCreacionFacturasAnecoop:
    MuestraError Err.Number, "Creación Facturas Anecoop", Err.Description
    conn.RollbackTrans
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
        Me.FrameIntegracion.Height = 4665
        Me.FrameIntegracion.Width = 6555
        W = Me.FrameIntegracion.Width
        H = Me.FrameIntegracion.Height
    End If
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadselect = ""
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


Private Function ComprobarErrores(ByRef pb1 As ProgressBar) As Boolean
Dim NF As Long
Dim cad As String
Dim i As Integer
Dim Longitud As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim Numreg As Long
Dim SQL As String
Dim SQL1 As String
Dim Total As Long
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
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    i = 0
    lblProgres(1).Caption = "Comprobando errores Tabla temporal entradas "
    
    SQL = "select * from tmpentradaS"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    b = True
    i = 0
    While Not Rs.EOF And b
        i = i + 1

        Me.pb1.Value = Me.pb1.Value + 1
        lblProgres(2).Caption = "Linea " & i
        Me.Refresh

        Variedad = Format(Rs!codprodu, "00") & Format(Rs!codvarie, "00")

        ' comprobamos la fecha
        FechaEnt = DBLet(Rs!FechaEnt, "T")
        If Not EsFechaOK(FechaEnt) Then
            Mens = "Fecha incorrecta"
            SQL = "insert into tmpinformes (codusu, campo1, codigo1, importe1, importe2, fecha1, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(Variedad, "N") & "," & DBSet(Rs!CodSocio, "N") & "," & _
                  DBSet(Rs!codCampo, "N") & "," & DBSet(Rs!Numnotac, "N") & "," & _
                  DBSet(FechaEnt, "F") & "," & DBSet(Mens, "T") & ")"
            conn.Execute SQL
        End If


        ' comprobamos que exista el socio
        SQL = "select count(*) from rsocios where codsocio = " & DBSet(Rs!CodSocio, "N")
        If TotalRegistros(SQL) = 0 Then
            Mens = "Socio no existe"
            SQL = "insert into tmpinformes (codusu, campo1, codigo1, importe1, importe2, fecha1, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(Variedad, "N") & "," & DBSet(Rs!CodSocio, "N") & "," & _
                  DBSet(Rs!codCampo, "N") & "," & DBSet(Rs!Numnotac, "N") & "," & _
                  DBSet(FechaEnt, "F") & "," & DBSet(Mens, "T") & ")"
            conn.Execute SQL
        End If

        ' comprobamos que exista el campo
        SQL = "select count(*) from rcampos where codsocio = " & DBSet(Rs!CodSocio, "N")
        SQL = SQL & " and nrocampo = " & DBSet(Rs!codCampo, "N")
        SQL = SQL & " and codvarie = " & DBSet(Variedad, "N")
        SQL = SQL & " and fecbajas is null "
        If TotalRegistros(SQL) = 0 Then
            Mens = "Campo no existe o con fecha de baja"
            SQL = "insert into tmpinformes (codusu, campo1, codigo1, importe1, importe2, fecha1, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(Variedad, "N") & "," & DBSet(Rs!CodSocio, "N") & "," & _
                  DBSet(Rs!codCampo, "N") & "," & DBSet(Rs!Numnotac, "N") & "," & _
                  DBSet(FechaEnt, "F") & "," & DBSet(Mens, "T") & ")"
            conn.Execute SQL
        End If

        ' comprobamos que no exista mas de un campo con ese numero de orden campo (scampo.codcampo MB)
        SQL = "select count(*) from rcampos where codsocio = " & DBSet(Rs!CodSocio, "N")
        SQL = SQL & " and nrocampo = " & DBSet(Rs!codCampo, "N")
        SQL = SQL & " and codvarie = " & DBSet(Variedad, "N")
        If TotalRegistros(SQL) > 1 Then
            Mens = "Campo con más de un registro"
            SQL = "insert into tmpinformes (codusu, campo1, codigo1, importe1, importe2, fecha1, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(Variedad, "N") & "," & DBSet(Rs!CodSocio, "N") & "," & _
                  DBSet(Rs!codCampo, "N") & "," & DBSet(Rs!Numnotac, "N") & "," & _
                  DBSet(FechaEnt, "F") & "," & DBSet(Mens, "T") & ")"
            conn.Execute SQL
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
Dim SQL As String
Dim Sql2 As String
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
    
    If vParamAplic.CodAnecoop = "" Then
        MsgBox "No existe en parámetros el cliente Anecoop.", vbExclamation
        b = False
    Else
        CodtipomAnecoop = ""
        Codforpa = 0
        TipoIvac = 0
        Dto1 = 0
        Dto2 = 0

        ' cargamos los datos del cliente anecoop que vayamos a necesitar
        Set vClien = New CCliente
        If vClien.LeerDatos(vParamAplic.CodAnecoop) Then
            CodtipomAnecoop = vClien.tipoMov
            Codforpa = vClien.ForPago
            TipoIvac = vClien.TipoIva
            Dto1 = vClien.Dto1
            Dto2 = vClien.Dto2
        Else
            MsgBox "No existe el cliente Anecoop. Revise.", vbExclamation
            b = False
        End If
        Set vClien = Nothing
    
    End If
    
    
    
    DatosOk = b

End Function


Private Sub frmC_Selec(vFecha As Date)
    txtCodigo(CByte(imgFecha(0).Tag) + 14).Text = Format(vFecha, "dd/mm/yyyy") '<===
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)

    If CadenaSeleccion <> "" Then
        'Facturas = " anecoop.fra_liq in (" & Mid(CadenaSeleccion, 1, Len(CadenaSeleccion) - 1) & ") "
        Facturas = CadenaSeleccion
    Else
        Facturas = ""
    End If

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


'Private Function InicializarCStock(ByRef vCStock As CStock, TipoM As String, Optional numlinea As String) As Boolean
'   On Error Resume Next
'   'On Error GoTo eInicializar
'
'    vCStock.tipoMov = TipoM
'    vCStock.DetaMov = CodtipomAnecoop
'
'    vCStock.Trabajador = CInt(vParamAplic.CodAnecoop)  'guardamos el cliente de la factura
'
'    '[Monica]20/03/2012: guardamos el numero de albaran o de factura (dependiendo de de donde viene)
'    vCStock.Documento = CLng(Text1(0).Text) 'Nº Factura
'    vCStock.Fechamov = Text1(1).Text 'Fecha de la Factura
'
'    vCStock.codArtic = txtAux(5).Text
'    vCStock.codAlmac = CInt(txtAux(4).Text)
'    vCStock.Cantidad = CSng(ComprobarCero(txtAux(6).Text))
'    vCStock.Importe = CCur(ComprobarCero(txtAux(9).Text))
'    vCStock.LineaDocu = CInt(ComprobarCero(numlinea))
'
'    If Err.Number <> 0 Then
'        MsgBox "No se han podido inicializar la clase para actualizar Stock", vbExclamation
'        InicializarCStock = False
'    Else
'        InicializarCStock = True
'    End If
'
''eInicializar:
''    MuestraError Err.Number, "inicializar", Err.Description
'End Function


