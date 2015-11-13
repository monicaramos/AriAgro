VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTrazabilidad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cargas Salidas MESURASOFT"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6600
   Icon            =   "frmTrazabilidad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameTrasTraza 
      Height          =   4545
      Left            =   30
      TabIndex        =   9
      Top             =   30
      Width           =   6555
      Begin MSComctlLib.ProgressBar pb2 
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1920
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmdCancelTras 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5145
         TabIndex        =   12
         Top             =   3690
         Width           =   975
      End
      Begin VB.CommandButton cmdAcepTras 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3930
         TabIndex        =   11
         Top             =   3720
         Width           =   975
      End
      Begin MSComDlg.CommonDialog cmd2 
         Left            =   570
         Top             =   3390
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "doc"
      End
      Begin VB.Label lblProgres 
         Height          =   405
         Index           =   4
         Left            =   270
         TabIndex        =   16
         Top             =   2340
         Width           =   6195
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   3
         Left            =   270
         TabIndex        =   15
         Top             =   2865
         Width           =   6195
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Proceso que realiza la lectura de expedientes de traza para incorporarlas a la aplicación."
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
         Left            =   300
         TabIndex        =   14
         Top             =   630
         Width           =   5820
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   2
         Left            =   270
         TabIndex        =   13
         Top             =   3270
         Width           =   6195
      End
   End
   Begin VB.Frame FrameTrazabilidad 
      Height          =   4665
      Left            =   0
      TabIndex        =   2
      Top             =   60
      Width           =   6555
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   1305
         MaxLength       =   6
         TabIndex        =   7
         Tag             =   "Cod.Almacen|N|N|0|999|albaran|codalmac|000||"
         Text            =   "Text1"
         Top             =   2340
         Width           =   780
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   4
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   2340
         Width           =   3900
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   570
         Top             =   3735
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "doc"
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4905
         TabIndex        =   1
         Top             =   3780
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   0
         Top             =   3780
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   210
         TabIndex        =   3
         Top             =   3195
         Width           =   6030
         _ExtentX        =   10636
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1035
         ToolTipText     =   "Buscar Almacén"
         Top             =   2385
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Almacén"
         Height          =   255
         Index           =   1
         Left            =   225
         TabIndex        =   8
         Top             =   2385
         Width           =   810
      End
      Begin VB.Label lblProgres 
         Caption         =   "Procesando"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   555
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   675
         Width           =   6195
      End
      Begin VB.Label lblProgres 
         Caption         =   "Fichero"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   645
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   6195
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
Attribute VB_Name = "frmTrazabilidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' PROGRAMA DE TRASPASO DE TRAZABILIDAD
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MONICA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit


Public Opcionlistado As Byte
'    0 = Traspaso de Trazabilidad de Valsur
'    1 = Traspaso de Trazabilidad de Castelduc

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim indCodigo As Integer 'indice para txtCodigo
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim cad As String
Dim cadTABLA As String

Dim vContad As Long

Dim PrimeraVez As Boolean
Dim NomFicheros(7) As String
Dim nompath As String
Dim LongitudCadena As Long
Dim indice As Long

Private WithEvents frmAlm As frmManAlmProp 'Form Mto de almacenes propios
Attribute frmAlm.VB_VarHelpID = -1

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cmdAceptar_Click()
Dim Sql As String
Dim i As Byte
Dim cadWHERE As String
Dim b As Boolean
Dim CorrectoA As Boolean
Dim CorrectoB As Boolean
Dim CorrectoC As Boolean
Dim NomFic As String
Dim cadena As String
Dim cadena1 As String
Dim Sql2 As String
Dim Mens As String

On Error GoTo eError

    If Not DatosOk Then Exit Sub
    
    Me.CommonDialog1.Flags = cdlOFNExplorer + cdlOFNAllowMultiselect + cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist

    Me.CommonDialog1.InitDir = App.path & "\traza\salidas"
    Me.CommonDialog1.DefaultExt = "txt"
    CommonDialog1.Filter = "Archivos TXT|*.txt|"
    CommonDialog1.FilterIndex = 1
    Me.CommonDialog1.FileName = "pedido.txt"
    Me.CommonDialog1.CancelError = True
    Me.CommonDialog1.ShowOpen
    lblProgres(0).Caption = "Directorio Seleccionado : "
    nompath = CargarPath(Me.CommonDialog1.FileName)
    lblProgres(1).Caption = "    " & nompath
    
'    nompath = GetFolder("Selecciona directorio")
     If Me.CommonDialog1.FileName <> "" Then
'    If nompath <> "" Then
        InicializarVbles
        InicializarTabla
            '========= PARAMETROS  =============================
        'Añadir el parametro de Empresa
        cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
        numParam = numParam + 1

'        nompath = CargarPath(Me.CommonDialog1.FileName)

        LongitudCadena = 100

        
        CorrectoA = ComprobarErrores(nompath & "\pedido.txt", 0)
        CorrectoB = ComprobarErrores(nompath & "\linea.txt", 1)
        CorrectoC = ComprobarErrores(nompath & "\slinea.txt", 2)
        
        b = (CorrectoA And CorrectoB And CorrectoC)
        
        If Not b Then Exit Sub
        AnyadirAFormula cadFormula, "{tmpinformes.codusu} = " & vUsu.codigo
        Sql = "select count(*) from tmpinformes where codusu = " & vUsu.codigo
        
        If TotalRegistros(Sql) <> 0 Then
            MsgBox "Hay errores en el Traspaso de Trazabilidad. Debe corregirlos previamente.", vbExclamation
            cadTitulo = "Errores de Traspaso"
            cadNombreRPT = "rErroresTraza.rpt"
            LlamarImprimir
            Exit Sub
        Else
            conn.BeginTrans
           
            b = True
            Mens = ""
            If b Then b = ProcesarFichero(nompath & "\PEDIDOS.TXT", 0, Mens)
            If b Then b = ProcesarFichero(nompath & "\LINEAS.TXT", 1, Mens)
            If b Then b = ProcesarFichero(nompath & "\SUBLINEAS.TXT", 2, Mens)
        
        End If
    End If

eError:
    If Err.Number = 32755 Then
        ' le hemos dado a cancelar, no hacemos nada.
        cmdCancel_Click
    Else
        If Err.Number <> 0 Or Not b Then
            
            conn.RollbackTrans
            MsgBox "No se ha podido realizar el proceso. " & vbCrLf & vbCrLf & Mens, vbExclamation
        Else
            conn.CommitTrans
            MsgBox "Proceso realizado correctamente.", vbExclamation
            pb1.visible = False
            lblProgres(0).Caption = ""
            lblProgres(1).Caption = ""
            ' borramos los ficheros del directorio
            BorrarArchivo Me.CommonDialog1.FileName
            BorrarArchivo Replace(LCase(Me.CommonDialog1.FileName), "pedido", "linea")
            BorrarArchivo Replace(LCase(Me.CommonDialog1.FileName), "pedido", "slinea")
            cmdCancel_Click
        End If
    End If
End Sub

Private Sub cmdAcepTras_Click()
Dim Sql As String
Dim i As Byte
Dim cadWHERE As String
Dim b As Boolean
Dim NomFic As String
Dim cadena As String
Dim cadena1 As String
Dim Directorio As String
Dim fec As String
Dim nomDir As String

Dim Nregs As Long
Dim cadTABLA As String


Dim File1 As FileSystemObject

On Error GoTo eError

    If Not DatosOk Then Exit Sub

    
    Me.CommonDialog1.Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist


    Me.CommonDialog1.DefaultExt = "txt"
    CommonDialog1.Filter = "Archivos TXT|*.txt|"
    CommonDialog1.FilterIndex = 1
    Me.CommonDialog1.FileName = "*.txt"
    
    
    Me.CommonDialog1.CancelError = True
    Me.CommonDialog1.ShowOpen
    Set File1 = New FileSystemObject
    
    Directorio = File1.GetParentFolderName(Me.CommonDialog1.FileName)

    If Directorio <> "" Then

        Sql = "DROP TABLE IF EXISTS tmpalbaran; "
        conn.Execute Sql
        Sql = "DROP TABLE IF EXISTS tmpalbaran_variedad; "
        conn.Execute Sql
        Sql = "DROP TABLE IF EXISTS tmpalbaran_calibre; "
        conn.Execute Sql
        
        
        Sql = "CREATE TEMPORARY TABLE `tmpalbaran` ("
        Sql = Sql & "`codclien` int(7), "
        Sql = Sql & "`codexped` int(7), "
        Sql = Sql & "`fecexped` varchar(8), "
        Sql = Sql & "`matriveh` varchar(15), "
        Sql = Sql & "`matrirem` varchar(15), "
        Sql = Sql & "`numerocmr` int(7), "
        Sql = Sql & "`totpalet` int(3), "
        Sql = Sql & "`codtimer` smallint(3), "
        Sql = Sql & "`coddesti` smallint(3), "
        Sql = Sql & "`codtrans` smallint(3) "
        Sql = Sql & " ) ENGINE=InnoDB DEFAULT CHARSET=latin1"
    
        conn.Execute Sql
    
        Sql = "CREATE TEMPORARY TABLE `tmpalbaran_variedad` ("
        Sql = Sql & "`codexped` int(7), "
        Sql = Sql & "`fecexped` varchar(8), "
        Sql = Sql & "`codforfait` smallint, "
        Sql = Sql & "`categori` varchar(3), "
        Sql = Sql & "`marcafru` varchar(15), "
        Sql = Sql & "`codprodu` smallint(3), "
        Sql = Sql & "`codvarie` smallint(3), "
        Sql = Sql & "`codcalib` varchar(3), "
        Sql = Sql & "`numcajas` int(7), "
        Sql = Sql & "`pesobrut` int(7), "
        Sql = Sql & "`pesoneto` int(7), "
        Sql = Sql & "`numpalet` int(7) "
        '[Monica]12/12/2013: introducimos el comisionista
        Sql = Sql & ", `codcomis` smallint(3) "
        Sql = Sql & " ) ENGINE=InnoDB DEFAULT CHARSET=latin1"
    
        conn.Execute Sql
        
        Sql = "CREATE TEMPORARY TABLE `tmpalbaran_calibre` ("
        Sql = Sql & "`codexped` int(7), "
        Sql = Sql & "`fecexped` varchar(8), "
        Sql = Sql & "`codforfait` smallint, "
        Sql = Sql & "`categori` varchar(3), "
        Sql = Sql & "`marcafru` varchar(15), "
        Sql = Sql & "`codprodu` smallint(3), "
        Sql = Sql & "`codvarie` smallint(3), "
        Sql = Sql & "`codcalib` varchar(3), "
        Sql = Sql & "`numcajas` int(7), "
        Sql = Sql & "`pesobrut` int(7), "
        Sql = Sql & "`pesoneto` int(7) "
        '[Monica]12/12/2013: introducimos el comisionista
        Sql = Sql & ", `codcomis` smallint(3) "
        Sql = Sql & " ) ENGINE=InnoDB DEFAULT CHARSET=latin1"
        
        conn.Execute Sql
        
        
        conn.BeginTrans

        nomDir = Directorio & "\"

        NomFic = Dir(nomDir & "*.txt")  ' Recupera la primera entrada.
    
        
        ' Cargamos en la tabla temporal todas las entradas de los ficheros del directorio seleccionado
        
        Do While NomFic <> ""   ' Inicia el bucle.
           ' Ignora el directorio actual y el que lo abarca.
           If NomFic <> "." And NomFic <> ".." Then
              Select Case UCase(Mid(NomFic, 1, 1))
                Case "C"
                    ' Realiza una comparación a nivel de bit para asegurarse de que MiNombre es un directorio.
'                    If GetAttr(nomDir & NomFic) And vbArchive = vbArchive Then
                          lblProgres(0).Caption = "Procesando Fichero: " & NomFic
                        
                          Sql = "load data local infile '" & Replace(nomDir & NomFic, "\", "/") & "' into table `tmpalbaran` fields terminated by '|' lines terminated by '\n' "
                          Sql = Sql & "(`codclien`,`codexped`,`fecexped`,`matriveh`,`matrirem`,`numerocmr`,`totpalet`,`codtimer`,`coddesti`,`codtrans`)  "
                          conn.Execute Sql
                      
'                    End If
                Case "L"
                    ' Realiza una comparación a nivel de bit para asegurarse de que MiNombre es un directorio.
'                    If GetAttr(nomDir & NomFic) And vbArchive = vbArchive Then
                          lblProgres(0).Caption = "Procesando Fichero: " & NomFic
                        
                          Sql = "load data local infile '" & Replace(nomDir & NomFic, "\", "/") & "' into table `tmpalbaran_variedad` fields terminated by '|' lines terminated by '\n' "
                          Sql = Sql & "(`codexped`,`fecexped`,`codforfait`,`categori`,`marcafru`,`codprodu`,`codvarie`,`codcalib`,`numcajas`,`pesobrut`,`pesoneto`,`numpalet`, `codcomis`)  "
                          conn.Execute Sql
                      
'                    End If
              End Select
           End If
           
           NomFic = Dir   ' Obtiene siguiente entrada.
        Loop

        Sql = "select count(distinct tmpalbaran.codexped) from tmpalbaran, tmpalbaran_variedad where tmpalbaran.codexped = tmpalbaran_variedad.codexped "
        Nregs = TotalRegistros(Sql)
        If Nregs <> 0 Then
'            Pb1.visible = True
'            Pb1.Max = Nregs
'            Pb1.Value = 0
'            Me.Refresh
'            DoEvents
                
            InicializarVbles
                
                '========= PARAMETROS  =============================
            'Añadir el parametro de Empresa
            cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
            numParam = numParam + 1
    
            If ComprobarErrores2(pb1) Then
                cadTABLA = "tmpinformes"
                cadFormula = "{tmpinformes.codusu} = " & vUsu.codigo
                
                Sql = "select count(*) from tmpinformes where codusu = " & vUsu.codigo
                
                If TotalRegistros(Sql) <> 0 Then
                    MsgBox "Hay errores en el Traspaso de Trazabilidad." & vbCrLf & "   Debe corregirlos previamente.", vbExclamation
                    cadTitulo = "Errores de Traspaso de TRAZABILIDAD"
                    cadNombreRPT = "rErroresTrasTraza.rpt"
                    LlamarImprimir
                    conn.RollbackTrans
                    lblProgres(0).Caption = ""
                    lblProgres(1).Caption = ""
                    lblProgres(2).Caption = ""
                    Exit Sub
                Else
                    b = CargarExpedientes()
                End If
            Else
                b = False
            End If
                
        End If

    End If
    
eError:
    If Err.Number = 32755 Then Exit Sub ' le han dado a cancelar

    If Err.Number <> 0 Or Not b Then
        conn.RollbackTrans
        MsgBox "No se ha podido realizar el proceso.", vbExclamation
    Else
        conn.CommitTrans
        MsgBox "Proceso realizado correctamente.", vbExclamation
        pb1.visible = False
        lblProgres(0).Caption = ""
        lblProgres(1).Caption = ""
        lblProgres(2).Caption = ""

        BorrarArchivo Directorio & "\C*.txt"
        BorrarArchivo Directorio & "\L*.txt"
        cmdCancelTras_Click
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancelTras_Click()
    Unload Me
End Sub

Private Sub Form_Activate()

On Error GoTo eError
    
    If PrimeraVez Then
        PrimeraVez = False
    End If
    Screen.MousePointer = vbDefault

    Select Case Opcionlistado
        Case 0 ' Programa de traspaso de trazabilidad
            lblProgres(0).Caption = ""
            lblProgres(1).Caption = ""
        
        Case 1 ' Programa de traspaso de trazabilidad de castelduc
        
    End Select

    
eError:
    If Err.Number = cdlCancel Then Unload Me
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection



    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
    '###Descomentar
'    CommitConexion
         
    Me.FrameTrasTraza.visible = False
    Me.FrameTrazabilidad.visible = False
    
    Select Case Opcionlistado
    
        Case 0 ' Traspaso de traza de valsur
            FrameTrazabilidadVisible True, H, W
            pb1.visible = False
        Case 1
            FrameTrasTrazaVisible True, H, W
            pb2.visible = False
                     
    End Select
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = W + 70
    Me.Height = H + 35
    Me.imgBuscar(4).Picture = frmPpal.imgListImages16.ListImages(1).Picture

End Sub

Private Sub FrameTrazabilidadVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de socios por seccion
    Me.FrameTrazabilidad.visible = visible
    If visible = True Then
        Me.FrameTrazabilidad.Top = -90
        Me.FrameTrazabilidad.Left = 0
        Me.FrameTrazabilidad.Height = 4665
        Me.FrameTrazabilidad.Width = 6555
        W = Me.FrameTrazabilidad.Width
        H = Me.FrameTrazabilidad.Height
    End If
End Sub

Private Sub FrameTrasTrazaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de socios por seccion
    Me.FrameTrasTraza.visible = visible
    If visible = True Then
        Me.FrameTrasTraza.Top = -90
        Me.FrameTrasTraza.Left = 0
        Me.FrameTrasTraza.Height = 4545
        Me.FrameTrasTraza.Width = 6555
        W = Me.FrameTrasTraza.Width
        H = Me.FrameTrasTraza.Height
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DesBloqueoManual ("TRASPOST")
End Sub



Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
End Sub


Private Function RecuperaFichero() As Boolean
Dim NF As Integer

    RecuperaFichero = False
    NF = FreeFile
    Open App.path For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    Line Input #NF, cad
    Close #NF
    If cad <> "" Then RecuperaFichero = True
    
End Function


Private Function ProcesarFichero(nomFich As String, Opcion As Integer, ByRef Mens As String) As Boolean
Dim NF As Long
Dim cad As String
Dim i As Integer
Dim Longitud As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim NumReg As Long
Dim Sql As String
Dim Codclave As String
Dim Codforpa As String


Dim Fin As Integer
Dim campo As Integer

Dim P_Pedido As String
Dim P_envio As String
Dim P_cliente As String
Dim P_subcliente As String
Dim P_fechaexp As String
Dim P_codtrans As String
Dim P_matriveh As String
Dim P_matrirem As String
Dim P_numpalet As String
Dim P_codtimer As String
Dim P_coddesti As String
Dim P_observa1 As String
Dim P_observa2 As String
Dim P_observa3 As String

Dim L_Pedido As String
Dim L_nrolinea As String
Dim L_codprodu As String
Dim L_codvarie As String
Dim L_codconfe As String
Dim L_codmarca As String
Dim L_categori As String
Dim L_NumCajas As String
Dim L_pesobrut As String
Dim L_PesoNeto As String

Dim S_Pedido  As String
Dim S_nrolinea As String
Dim S_sublinea As String
Dim S_codprodu As String
Dim S_codvarie As String
Dim S_codcalib As String
Dim S_pesobrut As String
Dim S_PesoNeto As String
Dim S_NumCajas As String
Dim S_numunida As String

Dim b As Boolean

    On Error GoTo eProcesarFichero

    ProcesarFichero = False
    NF = FreeFile

    Open nomFich For Input As #NF

    i = 0

    lblProgres(0).Caption = "Procesando Fichero: " & nomFich
    Longitud = FileLen(nomFich)

    Longitud = FileLen(nomFich)
    
    
    pb1.visible = True
    Me.pb1.Max = Longitud
    Me.Refresh
    Me.pb1.Value = 0

    b = True
    
    Do While Not EOF(NF) And b
        Line Input #NF, cad
         
        i = i + 1
        
        Me.pb1.Value = Me.pb1.Value + Len(cad)
        lblProgres(1).Caption = "Linea " & i
        Me.Refresh
        
        'cargamos los valores a comprobar
        Select Case Opcion
            Case 0 ' fichero pedidos
                Fin = InStr(1, cad, "|")
                campo = 0
                While Fin <> 0 Or cad <> ""
                    campo = campo + 1
                    Select Case campo
                        Case 1 'pedido
                            P_Pedido = Mid(cad, 1, Fin - 1)
                        Case 2 'nro envio
                            P_envio = Mid(cad, 1, Fin - 1)
                        Case 3 'codclien
                            P_cliente = Mid(cad, 1, Fin - 1)
                        Case 4
                            P_subcliente = Mid(cad, 1, Fin - 1)
                        Case 5 'fechaexped
                            P_fechaexp = Mid(cad, 1, Fin - 1)
                        Case 6 'codtrans
                            P_codtrans = Mid(cad, 1, Fin - 1)
                        Case 7 'matricula veh
                            P_matriveh = Mid(cad, 1, Fin - 1)
                        Case 8 'matricula rem
                            P_matrirem = Mid(cad, 1, Fin - 1)
                        Case 9 'numpalet
                            P_numpalet = Mid(cad, 1, Fin - 1)
                        Case 10 'codtimer
                            P_codtimer = Mid(cad, 1, Fin - 1)
                        Case 11 'coddesti
                            P_coddesti = Mid(cad, 1, Fin - 1)
                        Case 12 'observa1
                            P_observa1 = Mid(cad, 1, Fin - 1)
                        Case 13 'observa2
                            P_observa2 = Mid(cad, 1, Fin - 1)
                        Case 14 'observa3
                            P_observa3 = Mid(cad, 1, Fin - 1)
                    End Select
                    
                    If (Fin + 1) < Len(cad) Then
                        cad = Mid(cad, Fin + 1, Len(cad))
                    Else
                        cad = ""
                    End If
                    
                    Fin = InStr(1, cad, "|")
                Wend
            Case 1 ' fichero lineas
                Fin = InStr(1, cad, "|")
                campo = 0
                While Fin <> 0 Or cad <> ""
                    campo = campo + 1
                    Select Case campo
                        Case 1 'pedido
                            L_Pedido = Mid(cad, 1, Fin - 1)
                        Case 2 'nro linea
                            L_nrolinea = Mid(cad, 1, Fin - 1)
                        Case 3 'codprodu
                            L_codprodu = Mid(cad, 1, Fin - 1)
                        Case 4 'codvarie
                            L_codvarie = Mid(cad, 1, Fin - 1)
                        Case 5 'codconfe
                            L_codconfe = Mid(cad, 1, Fin - 1)
                        Case 6 'codmarca
                            L_codmarca = Mid(cad, 1, Fin - 1)
                        Case 7 'categori
                            L_categori = Mid(cad, 1, Fin - 1)
                        Case 8 'numcajas
                            L_NumCajas = Mid(cad, 1, Fin - 1)
                        Case 9 'pesobrut
                            L_pesobrut = Mid(cad, 1, Fin - 1)
                        Case 10 'pesoneto
                            L_PesoNeto = Mid(cad, 1, Fin - 1)
                    End Select
                    
                    If (Fin + 1) < Len(cad) Then
                        cad = Mid(cad, Fin + 1, Len(cad))
                    Else
                        cad = ""
                    End If
                    
                    Fin = InStr(1, cad, "|")
                Wend
                
            
            Case 2 ' fichero sublineas
                Fin = InStr(1, cad, "|")
                campo = 0
                While Fin <> 0 Or cad <> ""
                    campo = campo + 1
                    Select Case campo
                        Case 1 'pedido
                            S_Pedido = Mid(cad, 1, Fin - 1)
                        Case 2 'nro linea
                            S_nrolinea = Mid(cad, 1, Fin - 1)
                        Case 3 'sublinea
                            S_sublinea = Mid(cad, 1, Fin - 1)
                        Case 4 'codprodu
                            S_codprodu = Mid(cad, 1, Fin - 1)
                        Case 5 'codvarie
                            S_codvarie = Mid(cad, 1, Fin - 1)
                        Case 6 'codcalib
                            S_codcalib = Mid(cad, 1, Fin - 1)
                        Case 7 'pesobrut
                            S_pesobrut = Mid(cad, 1, Fin - 1)
                        Case 8 'pesoneto
                            S_PesoNeto = Mid(cad, 1, Fin - 1)
                        Case 9 'numcajas
                            S_NumCajas = Mid(cad, 1, Fin - 1)
                        Case 10 'numunida
                            S_numunida = Mid(cad, 1, Fin - 1)
                    End Select
                    
                    If (Fin + 1) < Len(cad) Then
                        cad = Mid(cad, Fin + 1, Len(cad))
                    Else
                        cad = ""
                    End If
                    
                    Fin = InStr(1, cad, "|")
                Wend
                
        End Select
        
        Select Case Opcion
            Case 0
                Sql = "insert into albaran (numalbar, fechaalb, codclien, coddesti, codtrans, matriveh, "
                Sql = Sql & "matrirem, refclien, codtimer, totpalet, portespre, nrocontra, nroactas,"
                Sql = Sql & "numpedid, fechaped, observac, pasaridoc, codalmac, portespag) values ("
                Sql = Sql & DBSet(P_Pedido, "N") & ","
                Sql = Sql & DBSet(P_fechaexp, "F") & ","
                Sql = Sql & DBSet(P_cliente, "N") & ","
                Sql = Sql & DBSet(P_coddesti, "N") & ","
                Sql = Sql & DBSet(P_codtrans, "N") & ","
                Sql = Sql & DBSet(P_matriveh, "T") & ","
                Sql = Sql & DBSet(P_matrirem, "T") & ","
                Sql = Sql & ValorNulo & ","
                Sql = Sql & DBSet(P_codtimer, "N") & ","
                Sql = Sql & ValorNulo & ","
                Sql = Sql & ValorNulo & ","
                Sql = Sql & ValorNulo & ","
                Sql = Sql & ValorNulo & ","
                Sql = Sql & ValorNulo & ","
                Sql = Sql & ValorNulo & ","
                Sql = Sql & DBSet(P_observa1 & P_observa2 & P_observa3, "T") & ",0,"
                Sql = Sql & DBSet(Text1(4).Text, "N") & ","
                Sql = Sql & ValorNulo & ")"
                                
                conn.Execute Sql
                
            Case 1
                Sql = "insert into albaran_variedad (numalbar, numlinea, codvarie, codvarco, codforfait,"
                Sql = Sql & "codmarca, categori, totpalet, numcajas, pesobrut, pesoneto, preciopro,"
                Sql = Sql & "preciodef, codincid, impcomis, observac, unidades) values ("
                Sql = Sql & DBSet(L_Pedido, "N") & ","
                Sql = Sql & DBSet(L_nrolinea, "N") & ","
                Sql = Sql & DBSet(L_codvarie, "N") & ","
                Sql = Sql & DBSet(L_codvarie, "N") & ","
                Sql = Sql & DBSet(L_codconfe, "T") & ","
                Sql = Sql & DBSet(L_codmarca, "N") & ","
                Sql = Sql & DBSet(L_categori, "T") & ","
                Sql = Sql & ValorNulo & ","
                Sql = Sql & DBSet(L_NumCajas, "N") & ","
                Sql = Sql & DBSet(L_pesobrut, "N") & ","
                Sql = Sql & DBSet(L_PesoNeto, "N") & ","
                Sql = Sql & ValorNulo & ","
                Sql = Sql & ValorNulo & ",0,"
                Sql = Sql & ValorNulo & ","
                Sql = Sql & ValorNulo & ","
                Sql = Sql & ValorNulo & ")"
                
                conn.Execute Sql
            
                b = ActualizarCostes(CInt(L_Pedido), CInt(L_nrolinea), True, L_codconfe, "")
            
            Case 2
                Sql = "insert into albaran_calibre (numalbar, numlinea, numline1, codvarie, codcalib,"
                Sql = Sql & "numcajas, pesobrut, pesoneto, unidades) values ("
                Sql = Sql & DBSet(S_Pedido, "N") & ","
                Sql = Sql & DBSet(S_nrolinea, "N") & ","
                Sql = Sql & DBSet(S_sublinea, "N") & ","
                Sql = Sql & DBSet(S_codvarie, "N") & ","
                Sql = Sql & DBSet(S_codcalib, "N") & ","
                Sql = Sql & DBSet(S_NumCajas, "N") & ","
                Sql = Sql & DBSet(S_pesobrut, "N") & ","
                Sql = Sql & DBSet(S_PesoNeto, "N") & ","
                Sql = Sql & DBSet(S_numunida, "N") & ")"
                
                conn.Execute Sql
            
            
        End Select
    Loop
    Close #NF
    
    ProcesarFichero = b
    
    pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

eProcesarFichero:
    If Not b Then
        Mens = "Error en el proceso de fichero, actualizando costes. Fichero linea.txt: Linea " & i
    Else
        If Err.Number <> 0 Then
            Mens = "Error en el proceso de fichero "
            Select Case Opcion
                Case 0
                    Mens = Mens & "pedido.txt  " & Err.Description
                Case 1
                    Mens = Mens & "linea.txt  " & Err.Description
                Case 2
                    Mens = Mens & "slinea.txt  " & Err.Description
            End Select
        End If
    End If
End Function
                
Private Function ComprobarErrores(nomFich As String, Opcion As Integer) As Boolean
Dim NF As Long
Dim cad As String
Dim i As Integer
Dim J As Integer
Dim Longitud As Long
Dim Rs As ADODB.Recordset
Dim NumReg As Long
Dim Sql As String
Dim SeProcesaLinea As Boolean

Dim Fin As Integer
Dim campo As Integer

Dim P_Pedido As String
Dim P_envio As String
Dim P_cliente As String
Dim P_subcliente As String
Dim P_fechaexp As String
Dim P_codtrans As String
Dim P_matriveh As String
Dim P_matrirem As String
Dim P_numpalet As String
Dim P_codtimer As String
Dim P_coddesti As String
Dim P_observa1 As String
Dim P_observa2 As String
Dim P_observa3 As String

Dim L_Pedido As String
Dim L_nrolinea As String
Dim L_codprodu As String
Dim L_codvarie As String
Dim L_codconfe As String
Dim L_codmarca As String
Dim L_categori As String
Dim L_NumCajas As String
Dim L_pesobrut As String
Dim L_PesoNeto As String

Dim S_Pedido  As String
Dim S_nrolinea As String
Dim S_sublinea As String
Dim S_codprodu As String
Dim S_codvarie As String
Dim S_codcalib As String
Dim S_pesobrut As String
Dim S_PesoNeto As String
Dim S_NumCajas As String
Dim S_numunida As String

Dim Mens As String
Dim Mens1 As String
Dim Caracter As String


    On Error GoTo eComprobarErrores
    
    ComprobarErrores = False

    If Dir(LCase(nomFich)) = "" Then
        MsgBox "No existe el fichero " & nomFich, vbExclamation
        Exit Function
    End If
    
    NF = FreeFile
    
    Open nomFich For Input As #NF
    
'    Line Input #NF, cad
    i = 0
    
    lblProgres(0).Caption = "Comprobando Fichero: " & nomFich
    Longitud = FileLen(nomFich)
    
    
    pb1.visible = True
    Me.pb1.Max = Longitud
    Me.Refresh
    Me.pb1.Value = 0

    Do While Not EOF(NF)
        Line Input #NF, cad
         
        i = i + 1
        
        Me.pb1.Value = Me.pb1.Value + Len(cad)
        lblProgres(1).Caption = "Linea " & i
        Me.Refresh
        
        'cargamos los valores a comprobar
        Select Case Opcion
            Case 0 ' fichero pedidos
                Fin = InStr(1, cad, "|")
                campo = 0
                While Fin <> 0 Or cad <> ""
                    campo = campo + 1
                    Select Case campo
                        Case 1 'pedido
                            P_Pedido = Mid(cad, 1, Fin - 1)
                        Case 2 'nro envio
                            P_envio = Mid(cad, 1, Fin - 1)
                        Case 3 'codclien
                            P_cliente = Mid(cad, 1, Fin - 1)
                        Case 4
                            P_subcliente = Mid(cad, 1, Fin - 1)
                        Case 5 'fechaexped
                            P_fechaexp = Mid(cad, 1, Fin - 1)
                        Case 6 'codtrans
                            P_codtrans = Mid(cad, 1, Fin - 1)
                        Case 7 'matricula veh
                            P_matriveh = Mid(cad, 1, Fin - 1)
                        Case 8 'matricula rem
                            P_matrirem = Mid(cad, 1, Fin - 1)
                        Case 9 'numpalet
                            P_numpalet = Mid(cad, 1, Fin - 1)
                        Case 10 'codtimer
                            P_codtimer = Mid(cad, 1, Fin - 1)
                        Case 11 'coddesti
                            P_coddesti = Mid(cad, 1, Fin - 1)
                        Case 12 'observa1
                            P_observa1 = Mid(cad, 1, Fin - 1)
                        Case 13 'observa2
                            P_observa2 = Mid(cad, 1, Fin - 1)
                        Case 14 'observa3
                            P_observa3 = Mid(cad, 1, Fin - 1)
                    End Select
                    
                    If (Fin + 1) < Len(cad) Then
                        cad = Mid(cad, Fin + 1, Len(cad))
                    Else
                        cad = ""
                    End If
                    
                    Fin = InStr(1, cad, "|")
                Wend
            Case 1 ' fichero lineas
                Fin = InStr(1, cad, "|")
                campo = 0
                While Fin <> 0 Or cad <> ""
                    campo = campo + 1
                    Select Case campo
                        Case 1 'pedido
                            L_Pedido = Mid(cad, 1, Fin - 1)
                        Case 2 'nro linea
                            L_nrolinea = Mid(cad, 1, Fin - 1)
                        Case 3 'codprodu
                            L_codprodu = Mid(cad, 1, Fin - 1)
                        Case 4 'codvarie
                            L_codvarie = Mid(cad, 1, Fin - 1)
                        Case 5 'codconfe
                            L_codconfe = Mid(cad, 1, Fin - 1)
                        Case 6 'codmarca
                            L_codmarca = Mid(cad, 1, Fin - 1)
                        Case 7 'categori
                            L_categori = Mid(cad, 1, Fin - 1)
                        Case 8 'numcajas
                            L_NumCajas = Mid(cad, 1, Fin - 1)
                        Case 9 'pesobrut
                            L_pesobrut = Mid(cad, 1, Fin - 1)
                        Case 10 'pesoneto
                            L_PesoNeto = Mid(cad, 1, Fin - 1)
                    End Select
                    
                    If (Fin + 1) < Len(cad) Then
                        cad = Mid(cad, Fin + 1, Len(cad))
                    Else
                        cad = ""
                    End If
                    
                    Fin = InStr(1, cad, "|")
                Wend
                
            
            Case 2 ' fichero sublineas
                Fin = InStr(1, cad, "|")
                campo = 0
                While Fin <> 0 Or cad <> ""
                    campo = campo + 1
                    Select Case campo
                        Case 1 'pedido
                            S_Pedido = Mid(cad, 1, Fin - 1)
                        Case 2 'nro linea
                            S_nrolinea = Mid(cad, 1, Fin - 1)
                        Case 3 'sublinea
                            S_sublinea = Mid(cad, 1, Fin - 1)
                        Case 4 'codprodu
                            S_codprodu = Mid(cad, 1, Fin - 1)
                        Case 5 'codvarie
                            S_codvarie = Mid(cad, 1, Fin - 1)
                        Case 6 'codcalib
                            S_codcalib = Mid(cad, 1, Fin - 1)
                        Case 7 'pesobrut
                            S_pesobrut = Mid(cad, 1, Fin - 1)
                        Case 8 'pesoneto
                            S_PesoNeto = Mid(cad, 1, Fin - 1)
                        Case 9 'numcajas
                            S_NumCajas = Mid(cad, 1, Fin - 1)
                        Case 10 'numunida
                            S_numunida = Mid(cad, 1, Fin - 1)
                    End Select
                    
                    If (Fin + 1) < Len(cad) Then
                        cad = Mid(cad, Fin + 1, Len(cad))
                    Else
                        cad = ""
                    End If
                    
                    Fin = InStr(1, cad, "|")
                Wend
                
        End Select
        
        'Comprobamos que el numero de albaran no exista
        If Opcion = 0 Then
            Sql = ""
            Sql = DevuelveDesdeBDNew(cAgro, "albaran", "numalbar", "numalbar", P_Pedido, "N")
            If Sql <> "" Then
                Mens = "Existe el número de albarán " & P_Pedido
                Mens1 = "Fichero Pedido.txt  Linea " & i
                Sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                      vUsu.codigo & "," & DBSet(Mens, "T") & "," & DBSet(Mens1, "T") & ")"
                      
                conn.Execute Sql
            End If
        End If
        If Opcion = 1 Then
            Sql = ""
            Sql = DevuelveDesdeBDNew(cAgro, "albaran", "numalbar", "numalbar", L_Pedido, "N")
            If Sql <> "" Then
                Mens = "Existe el número de albarán " & L_Pedido
                Mens1 = "Fichero Lineas.txt  Linea " & i
                Sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                      vUsu.codigo & "," & DBSet(Mens, "T") & "," & DBSet(Mens1, "T") & ")"
                      
                conn.Execute Sql
            End If
        End If
        If Opcion = 2 Then
            Sql = ""
            Sql = DevuelveDesdeBDNew(cAgro, "albaran", "numalbar", "numalbar", L_Pedido, "N")
            If Sql <> "" Then
                Mens = "Existe el número de albarán " & S_Pedido
                Mens1 = "Fichero Sublineas.txt  Linea " & i
                Sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                      vUsu.codigo & "," & DBSet(Mens, "T") & "," & DBSet(Mens1, "T") & ")"
                      
                conn.Execute Sql
            End If
        End If
        If Opcion = 0 Then
            'Comprobamos que exista el cliente
            Sql = ""
            Sql = DevuelveDesdeBDNew(cAgro, "clientes", "codclien", "codclien", P_cliente, "N")
            If Sql = "" Then
                Mens = "No Existe el cliente " & P_cliente
                Mens1 = "Fichero Pedido.txt  Linea " & i
                Sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                      vUsu.codigo & "," & DBSet(Mens, "T") & "," & DBSet(Mens1, "T") & ")"
                      
                conn.Execute Sql
            End If
                    
            'Comprobamos que exista la agencia de transporte
            Sql = ""
            Sql = DevuelveDesdeBDNew(cAgro, "agencias", "codtrans", "codtrans", P_codtrans, "N")
            If Sql = "" Then
                Mens = "No Existe la agencia de transporte " & P_codtrans
                Mens1 = "Fichero Pedido.txt  Linea " & i
                Sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                      vUsu.codigo & "," & DBSet(Mens, "T") & "," & DBSet(Mens1, "T") & ")"
                      
                conn.Execute Sql
            End If
        
            'Comprobamos que exista el tipo de mercado
            Sql = ""
            Sql = DevuelveDesdeBDNew(cAgro, "tipomer", "codtimer", "codtimer", P_codtimer, "N")
            If Sql = "" Then
                Mens = "No Existe el tipo de mercado " & P_codtimer
                Mens1 = "Fichero Pedido.txt  Linea " & i
                Sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                      vUsu.codigo & "," & DBSet(Mens, "T") & "," & DBSet(Mens1, "T") & ")"
                      
                conn.Execute Sql
            End If
        
            'Comprobamos que exista el destino
            Sql = ""
            Sql = DevuelveDesdeBDNew(cAgro, "destinos", "codclien", "codclien", P_cliente, "N", , "coddesti", P_coddesti, "N")
            If Sql = "" Then
                Mens = "No Existe el destino " & P_cliente & " " & P_coddesti
                Mens1 = "Fichero Pedido.txt  Linea " & i
                Sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                      vUsu.codigo & "," & DBSet(Mens, "T") & "," & DBSet(Mens1, "T") & ")"
                      
                conn.Execute Sql
            End If
            
            If Not EsFechaOK(P_fechaexp) Then
                Mens = "Fecha incorrecta"
                Mens1 = "Fichero Pedido.txt  Linea " & i
                Sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                      vUsu.codigo & "," & DBSet(Mens, "T") & "," & DBSet(Mens1, "T") & ")"
                
                conn.Execute Sql
            End If
            
        End If
        If Opcion = 1 Then
            'Comprobamos que exista el producto
            Sql = ""
            Sql = DevuelveDesdeBDNew(cAgro, "productos", "codprodu", "codprodu", L_codprodu, "N")
            If Sql = "" Then
                Mens = "No Existe el producto " & L_codprodu
                Mens1 = "Fichero Lineas.txt  Linea " & i
                Sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                      vUsu.codigo & "," & DBSet(Mens, "T") & "," & DBSet(Mens1, "T") & ")"
                      
                conn.Execute Sql
            End If
        End If
        If Opcion = 2 Then
            'Comprobamos que exista el producto
            Sql = ""
            Sql = DevuelveDesdeBDNew(cAgro, "productos", "codprodu", "codprodu", S_codprodu, "N")
            If Sql = "" Then
                Mens = "No Existe el producto " & S_codprodu
                Mens1 = "Fichero Sublineas.txt  Linea " & i
                Sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                      vUsu.codigo & "," & DBSet(Mens, "T") & "," & DBSet(Mens1, "T") & ")"
                      
                conn.Execute Sql
            End If
        End If
        If Opcion = 1 Then
            'Comprobamos que exista la variedad
            Sql = ""
            Sql = DevuelveDesdeBDNew(cAgro, "variedades", "codvarie", "codvarie", L_codvarie, "N")
            If Sql = "" Then
                Mens = "No Existe la variedad " & L_codvarie
                Mens1 = "Fichero Lineas.txt  Linea " & i
                Sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                      vUsu.codigo & "," & DBSet(Mens, "T") & "," & DBSet(Mens1, "T") & ")"
                      
                conn.Execute Sql
            End If
        End If
        If Opcion = 2 Then
            'Comprobamos que exista la variedad
            Sql = ""
            Sql = DevuelveDesdeBDNew(cAgro, "variedades", "codvarie", "codvarie", S_codvarie, "N")
            If Sql = "" Then
                Mens = "No Existe la variedad " & S_codprodu
                Mens1 = "Fichero Sublineas.txt  Linea " & i
                Sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                      vUsu.codigo & "," & DBSet(Mens, "T") & "," & DBSet(Mens1, "T") & ")"
                      
                conn.Execute Sql
            End If
        End If
        If Opcion = 1 Then
            'Comprobamos que exista la confeccion
            Sql = ""
            Sql = DevuelveDesdeBDNew(cAgro, "forfaits", "codforfait", "codforfait", L_codconfe, "T")
            If Sql = "" Then
                Mens = "No Existe la confeccion " & L_codconfe
                Mens1 = "Fichero Lineas.txt  Linea " & i
                Sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                      vUsu.codigo & "," & DBSet(Mens, "T") & "," & DBSet(Mens1, "T") & ")"
                      
                conn.Execute Sql
            End If
        End If
        If Opcion = 2 Then
            'Comprobamos que exista el calibre
            Sql = ""
            Sql = DevuelveDesdeBDNew(cAgro, "calibres", "codcalib", "codvarie", S_codvarie, "N", , "codcalib", S_codcalib, "N")
            If Sql = "" Then
                Mens = "No Existe el calibre " & S_codvarie & " " & S_codcalib
                Mens1 = "Fichero Sublineas.txt  Linea " & i
                Sql = "insert into tmpinformes (codusu, nombre1, nombre2) values (" & _
                      vUsu.codigo & "," & DBSet(Mens, "T") & "," & DBSet(Mens1, "T") & ")"
                      
                conn.Execute Sql
            End If
        End If
        
        
        
    Loop
    Close #NF
    
    ComprobarErrores = True
    
    pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

eComprobarErrores:
    If Err.Number <> 0 Then
        cad = "Se ha producido un error en el proceso de comprobación"
        MsgBox cad, vbExclamation
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
        .Opcion = 0
        .Show vbModal
    End With
End Sub

Private Sub InicializarTabla()
Dim Sql As String
    Sql = "delete from tmpinformes where codusu = " & vUsu.codigo
    
    conn.Execute Sql
End Sub


Private Function CargarPath(nomFich As String) As String
'las cadenas de entrada pueden ser las dos siguientes:
' C:\programas\Arigasol\envios credit1.cre credit0.cre
' C:\programas\Arigasol\envios\credit1.cre
Dim i As Integer
Dim J As Integer
Dim Ini As String
Dim Ini2 As String

Dim lon As Integer

    Ini = InStr(1, nomFich, ".txt")
    Ini2 = InStr(1, nomFich, ".TXT")
    
    If Ini2 > Ini Then Ini = Ini2
    
    ' recorremos el inicio de cadena hasta credit1.cre
    While Asc(Mid(nomFich, Ini, 1)) <> 0 And Mid(nomFich, Ini, 1) <> "\"
        Ini = Ini - 1
    Wend

    CargarPath = Mid(nomFich, 1, Ini - 1)
End Function


Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod almacen
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nombre del almacen
End Sub

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 0
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)
Dim devuelve As String
Dim cadMen As String
Dim Sql As String

    'Quitar espacios en blanco por los lados
    Text1(Index).Text = Trim(Text1(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

          
    'Si queremos hacer algo ..
    Select Case Index
        Case 4 'Almacen
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "salmpr", "nomalmac")
                If Text2(Index).Text = "" Then
                    MsgBox "No existe el Almacén. Reintroduzca."
                    PonerFoco Text1(Index)
                End If
            End If
    End Select
End Sub

Private Sub imgBuscar_Click(Index As Integer)

    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 4 ' Almacen
            indice = Index
            PonerFoco Text1(indice)
            Set frmAlm = New frmManAlmProp
            frmAlm.DatosADevolverBusqueda = "0|1|"
            frmAlm.Show vbModal
            Set frmAlm = Nothing
            PonerFoco Text1(indice)
    End Select
    
    Screen.MousePointer = vbDefault
End Sub

Private Function DatosOk() As Boolean
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim b As Boolean

    On Error GoTo EDatosOK

    b = True
    
    Select Case vParamAplic.Cooperativa
        Case 1 ' Valsur
            b = False
            If Text1(4).Text = "" Then
                MsgBox "Debe introducir obligatoriamente el código de almacén", vbExclamation
            Else
                Text2(4).Text = PonerNombreDeCod(Text1(4), "salmpr", "nomalmac")
                If Text2(4).Text = "" Then
                    MsgBox "No existe el código de almacén. Reintroduzca.", vbExclamation
                    PonerFoco Text1(4)
                Else
                    b = True
                End If
            End If
            
    End Select
    
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function



Private Function ComprobarErrores2(ByRef pb1 As ProgressBar) As Boolean
Dim NF As Long
Dim cad As String
Dim i As Integer
Dim Longitud As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim NumReg As Long
Dim Sql As String
Dim Sql1 As String
Dim Total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean
Dim bAux As Boolean
Dim Mens As String
Dim Tipo As Integer
Dim FechaEnt As String
Dim Variedad As String
Dim Nregs As Long

    On Error GoTo eComprobarErrores2

    ComprobarErrores2 = False
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.codigo
    conn.Execute Sql

    i = 0
    lblProgres(1).Caption = "Comprobando errores Tabla temporal expedientes "
    
    Sql = "select distinct tmpalbaran.codclien, tmpalbaran.codexped, tmpalbaran.fecexped, tmpalbaran.codtimer, tmpalbaran_variedad.codprodu, "
    Sql = Sql & " tmpalbaran.coddesti, tmpalbaran.codtrans, "
    Sql = Sql & " tmpalbaran_variedad.codvarie, tmpalbaran_variedad.codforfait, tmpalbaran_variedad.codcalib, tmpalbaran_variedad.marcafru  "
    '[Monica]12/12/2013: añadimos el comisionista
    Sql = Sql & ", tmpalbaran_variedad.codcomis "
    Sql = Sql & " from tmpalbaran, tmpalbaran_variedad "
    Sql = Sql & " where tmpalbaran.codexped = tmpalbaran_variedad.codexped "
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Nregs = TotalRegistrosConsulta(Sql)

    pb1.visible = True
    pb1.Max = Nregs
    pb1.Value = 0
    Me.Refresh
    DoEvents


    b = True
    i = 0
    While Not Rs.EOF And b
        i = i + 1

        Me.pb1.Value = Me.pb1.Value + 1
        lblProgres(2).Caption = "Linea " & i
        Me.Refresh

        Variedad = Format(DBLet(Rs!codprodu, "N"), "00") & Format(DBLet(Rs!codvarie, "N"), "00")

        ' comprobamos la fecha
        FechaEnt = DBLet(Rs!fecexped, "T")
        If Not EsFechaOK(FechaEnt) Then
            Mens = "Fecha incorrecta"
            Sql = "insert into tmpinformes (codusu, codigo1, importe1, fecha1, campo2, importe2, campo1, importe3, nombre2, importe4, importe5, nombre1 ) values (" & _
                  vUsu.codigo & "," & DBSet(Rs!CodClien, "N") & "," & DBSet(Rs!codexped, "N") & "," & DBSet(FechaEnt, "F") & "," & _
                  DBSet(Rs!Codtimer, "N") & "," & DBSet(Rs!codforfait, "N") & "," & DBSet(Variedad, "N") & "," & DBSet(Rs!codcalib, "T") & "," & _
                  DBSet(Rs!MarcaFru, "T") & "," & _
                  DBSet(Rs!coddesti, "N") & "," & DBSet(Rs!codTrans, "N") & "," & _
                  DBSet(Mens, "T") & ")"
            conn.Execute Sql
        End If

        ' comprobamos que no exista ya el expediente
        Sql = "select count(*) from albaran where numalbar = " & DBSet(Rs!codexped, "N")
        If TotalRegistros(Sql) <> 0 Then
            Mens = "Expediente ya existe"
            Sql = "insert into tmpinformes (codusu, codigo1, importe1, fecha1, campo2, importe2, campo1, importe3, nombre2, importe4, importe5, nombre1) values (" & _
                  vUsu.codigo & "," & DBSet(Rs!CodClien, "N") & "," & DBSet(Rs!codexped, "N") & "," & DBSet(FechaEnt, "F") & "," & _
                  DBSet(Rs!Codtimer, "N") & "," & DBSet(Rs!codforfait, "N") & "," & DBSet(Variedad, "N") & "," & DBSet(Rs!codcalib, "T") & "," & _
                  DBSet(Rs!MarcaFru, "T") & "," & _
                  DBSet(Rs!coddesti, "N") & "," & DBSet(Rs!codTrans, "N") & "," & _
                  DBSet(Mens, "T") & ")"
            conn.Execute Sql
        End If
        
        ' comprobamos que exista la variedad
        Sql = "select count(*) from variedades where codvarie = " & DBSet(Variedad, "N")
        If TotalRegistros(Sql) = 0 Then
            Mens = "Variedad no existe"
            Sql = "insert into tmpinformes (codusu, codigo1, importe1, fecha1, campo2, importe2, campo1, importe3, nombre2, importe4, importe5, nombre1) values (" & _
                  vUsu.codigo & "," & DBSet(Rs!CodClien, "N") & "," & DBSet(Rs!codexped, "N") & "," & DBSet(FechaEnt, "F") & "," & _
                  DBSet(Rs!Codtimer, "N") & "," & DBSet(Rs!codforfait, "N") & "," & DBSet(Variedad, "N") & "," & DBSet(Rs!codcalib, "T") & "," & _
                  DBSet(Rs!MarcaFru, "T") & "," & _
                  DBSet(Rs!coddesti, "N") & "," & DBSet(Rs!codTrans, "N") & "," & _
                  DBSet(Mens, "T") & ")"
            conn.Execute Sql
        End If

        ' comprobamos que exista el cliente del expediente
        Sql = "select count(*) from clientes where codclien = " & DBSet(Rs!CodClien, "N")
        If TotalRegistros(Sql) = 0 Then
            Mens = "Cliente no existe"
            Sql = "insert into tmpinformes (codusu, codigo1, importe1, fecha1, campo2, importe2, campo1, importe3, nombre2, importe4, importe5, nombre1) values (" & _
                  vUsu.codigo & "," & DBSet(Rs!CodClien, "N") & "," & DBSet(Rs!codexped, "N") & "," & DBSet(FechaEnt, "F") & "," & _
                  DBSet(Rs!Codtimer, "N") & "," & DBSet(Rs!codforfait, "N") & "," & DBSet(Variedad, "N") & "," & DBSet(Rs!codcalib, "T") & "," & _
                  DBSet(Rs!MarcaFru, "T") & "," & _
                  DBSet(Rs!coddesti, "N") & "," & DBSet(Rs!codTrans, "N") & "," & _
                  DBSet(Mens, "T") & ")"
            conn.Execute Sql
        End If

        '[Monica]12/12/2013: si me viene comisionista comprobamos que exista
        If DBLet(Rs!codcomis, "N") <> 0 Then
            Sql = "select count(*) from agencias where codtrans = " & DBSet(Rs!codcomis, "N")
            If TotalRegistros(Sql) = 0 Then
                Mens = "Comisionista no existe"
                Sql = "insert into tmpinformes (codusu, codigo1, importe1, fecha1, campo2, importe2, campo1, importe3, nombre2, importe4, importe5, nombre1) values (" & _
                      vUsu.codigo & "," & DBSet(Rs!CodClien, "N") & "," & DBSet(Rs!codexped, "N") & "," & DBSet(FechaEnt, "F") & "," & _
                      DBSet(Rs!Codtimer, "N") & "," & DBSet(Rs!codforfait, "N") & "," & DBSet(Variedad, "N") & "," & DBSet(Rs!codcalib, "T") & "," & _
                      DBSet(Rs!MarcaFru, "T") & "," & _
                      DBSet(Rs!coddesti, "N") & "," & DBSet(Rs!codTrans, "N") & "," & _
                      DBSet(Mens, "T") & ")"
                conn.Execute Sql
            End If
        End If



'        ' el cliente del expediente ha de tener por lo menos un destino pq (cliente, destino) es referencial
'        sql = "select count(*) from destinos where codclien = " & DBSet(RS!CodClien, "N")
'        If TotalRegistros(sql) = 0 Then
'            Mens = "Cliente sin destino"
'            sql = "insert into tmpinformes (codusu, codigo1, importe1, fecha1, campo2, importe2, campo1, importe3, nombre2, nombre1) values (" & _
'                  vUsu.codigo & "," & DBSet(RS!CodClien, "N") & "," & DBSet(RS!Codexped, "N") & "," & DBSet(FechaEnt, "F") & "," & _
'                  DBSet(RS!Codtimer, "N") & "," & DBSet(RS!codforfait, "N") & "," & DBSet(Variedad, "N") & "," & DBSet(RS!codcalib, "N") & "," & _
'                  DBSet(RS!MarcaFru, "T") & "," & _
'                  DBSet(Mens, "T") & ")"
'            conn.Execute sql
'        End If

        ' el destino del cliente ha de existir
        Sql = "select count(*) from destinos where codclien = " & DBSet(Rs!CodClien, "N") & " and coddesti = " & DBSet(Rs!coddesti, "N")
        If TotalRegistros(Sql) = 0 Then
            Mens = "Cliente sin destino"
            Sql = "insert into tmpinformes (codusu, codigo1, importe1, fecha1, campo2, importe2, campo1, importe3, nombre2, importe4, importe5, nombre1) values (" & _
                  vUsu.codigo & "," & DBSet(Rs!CodClien, "N") & "," & DBSet(Rs!codexped, "N") & "," & DBSet(FechaEnt, "F") & "," & _
                  DBSet(Rs!Codtimer, "N") & "," & DBSet(Rs!codforfait, "N") & "," & DBSet(Variedad, "N") & "," & DBSet(Rs!codcalib, "T") & "," & _
                  DBSet(Rs!MarcaFru, "T") & "," & _
                  DBSet(Rs!coddesti, "N") & "," & DBSet(Rs!codTrans, "N") & "," & _
                  DBSet(Mens, "T") & ")"
            conn.Execute Sql
        End If
        
        ' comprobamos que exista el transportista del expediente
        Sql = "select count(*) from agencias where codtrans = " & DBSet(Rs!codTrans, "N")
        If TotalRegistros(Sql) = 0 Then
            Mens = "Agencia transp.no existe"
            Sql = "insert into tmpinformes (codusu, codigo1, importe1, fecha1, campo2, importe2, campo1, importe3, nombre2, importe4, importe5, nombre1) values (" & _
                  vUsu.codigo & "," & DBSet(Rs!CodClien, "N") & "," & DBSet(Rs!codexped, "N") & "," & DBSet(FechaEnt, "F") & "," & _
                  DBSet(Rs!Codtimer, "N") & "," & DBSet(Rs!codforfait, "N") & "," & DBSet(Variedad, "N") & "," & DBSet(Rs!codcalib, "T") & "," & _
                  DBSet(Rs!MarcaFru, "T") & "," & _
                  DBSet(Rs!coddesti, "N") & "," & DBSet(Rs!codTrans, "N") & "," & _
                  DBSet(Mens, "T") & ")"
            conn.Execute Sql
        End If
        
        
        ' comprobamos que exista el mercado
        Sql = "select count(*) from tipomer where codtimer = " & DBSet(Rs!Codtimer, "N")
        If TotalRegistros(Sql) = 0 Then
            Mens = "Tipo de Mercado no existe"
            Sql = "insert into tmpinformes (codusu, codigo1, importe1, fecha1, campo2, importe2, campo1, importe3, nombre2 , importe4, importe5, nombre1) values (" & _
                  vUsu.codigo & "," & DBSet(Rs!CodClien, "N") & "," & DBSet(Rs!codexped, "N") & "," & DBSet(FechaEnt, "F") & "," & _
                  DBSet(Rs!Codtimer, "N") & "," & DBSet(Rs!codforfait, "N") & "," & DBSet(Variedad, "N") & "," & DBSet(Rs!codcalib, "T") & "," & _
                  DBSet(Rs!MarcaFru, "T") & "," & _
                  DBSet(Rs!coddesti, "N") & "," & DBSet(Rs!codTrans, "N") & "," & _
                  DBSet(Mens, "T") & ")"
            conn.Execute Sql
        End If

        ' comprobamos que exista la confeccion
        Sql = "select count(*) from forfaits where codforfait = " & DBSet(Rs!codforfait, "N")
        If TotalRegistros(Sql) = 0 Then
            Mens = "Confección no existe"
            Sql = "insert into tmpinformes (codusu, codigo1, importe1, fecha1, campo2, importe2, campo1, importe3, nombre2, importe4, importe5, nombre1) values (" & _
                  vUsu.codigo & "," & DBSet(Rs!CodClien, "N") & "," & DBSet(Rs!codexped, "N") & "," & DBSet(FechaEnt, "F") & "," & _
                  DBSet(Rs!Codtimer, "N") & "," & DBSet(Rs!codforfait, "N") & "," & DBSet(Variedad, "N") & "," & DBSet(Rs!codcalib, "T") & "," & _
                  DBSet(Rs!MarcaFru, "T") & "," & _
                  DBSet(Rs!coddesti, "N") & "," & DBSet(Rs!codTrans, "N") & "," & _
                  DBSet(Mens, "T") & ")"
            conn.Execute Sql
        End If

        ' comprobamos que exista el calibre
        Sql = "select count(*) from calibres where codvarie = " & DBSet(Variedad, "N")
        Sql = Sql & " and nomcalab = " & DBSet(Rs!codcalib, "T")
        If TotalRegistros(Sql) = 0 Then
            Mens = "Calibre no existe"
            Sql = "insert into tmpinformes (codusu, codigo1, importe1, fecha1, campo2, importe2, campo1, importe3, nombre2, importe4, importe5, nombre1) values (" & _
                  vUsu.codigo & "," & DBSet(Rs!CodClien, "N") & "," & DBSet(Rs!codexped, "N") & "," & DBSet(FechaEnt, "F") & "," & _
                  DBSet(Rs!Codtimer, "N") & "," & DBSet(Rs!codforfait, "N") & "," & DBSet(Variedad, "N") & "," & DBSet(Rs!codcalib, "T") & "," & _
                  DBSet(Rs!MarcaFru, "T") & "," & _
                  DBSet(Rs!coddesti, "N") & "," & DBSet(Rs!codTrans, "N") & "," & _
                  DBSet(Mens, "T") & ")"
            conn.Execute Sql
        End If

        ' comprobamos que exista la marca fruta
        bAux = True
        If DBLet(Rs!MarcaFru, "T") = "" Then
            bAux = False
        Else
            Sql = ""
            Sql = DevuelveDesdeBDNew(cAgro, "marcas", "codmarca", "nommarca", Rs!MarcaFru, "T")
            bAux = (Sql <> "")
        End If
        If Not bAux Then
            Mens = "Marca no existe"
            Sql = "insert into tmpinformes (codusu, codigo1, importe1, fecha1, campo2, importe2, campo1, importe3, nombre2, importe4, importe5, nombre1) values (" & _
                  vUsu.codigo & "," & DBSet(Rs!CodClien, "N") & "," & DBSet(Rs!codexped, "N") & "," & DBSet(FechaEnt, "F") & "," & _
                  DBSet(Rs!Codtimer, "N") & "," & DBSet(Rs!codforfait, "N") & "," & DBSet(Variedad, "N") & "," & DBSet(Rs!codcalib, "T") & "," & _
                  DBSet(Rs!MarcaFru, "T") & "," & _
                  DBSet(Rs!coddesti, "N") & "," & DBSet(Rs!codTrans, "N") & "," & _
                  DBSet(Mens, "T") & ")"
            conn.Execute Sql
        End If
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""
    lblProgres(2).Caption = ""

    ComprobarErrores2 = b
    Exit Function

eComprobarErrores2:
    ComprobarErrores2 = False
End Function


Private Function CargarExpedientes() As Boolean
Dim Sql As String
Dim Sql1 As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim RS3 As ADODB.Recordset
Dim Rs4 As ADODB.Recordset

Dim Precio As Currency
Dim Transporte As Currency
Dim Kilos As Long

Dim AlbarAnt As Long
Dim KilosAlbar As Long
Dim KilosNetAnt As Long
Dim VarieAnt As Long
Dim CalidAnt As Long
Dim Longitud As Long

Dim campo As Variant
Dim cadMen As String

Dim Variedad As String
Dim TipoEntr As Byte
Dim Recolect As Byte
Dim KilosNet As Long
Dim MarcaFru As String

Dim fecha As String

Dim NumF As Long
Dim NumF1 As Long
Dim Destino As String
Dim b As Boolean
Dim Lista As String
Dim Calibre As String

    On Error GoTo eCargarExpedientes
    
    CargarExpedientes = False
    
    lblProgres(1).Caption = "Cargando Expedientes"
    
    Sql = "select count(*) from tmpalbaran " ' , tmpalbaran_variedad "
'    sql = sql & " where tmpalbaran.codexped = tmpalbaran_variedad.codexped "
    Longitud = TotalRegistros(Sql)
    
    pb1.visible = True
    Me.pb1.Max = Longitud
    Me.Refresh
    Me.pb1.Value = 0
    
    b = True
    
    Sql = "select tmpalbaran.codclien, tmpalbaran.codexped, tmpalbaran.fecexped, tmpalbaran.matriveh, "
    Sql = Sql & " tmpalbaran.matrirem, tmpalbaran.numerocmr, tmpalbaran.codtimer, tmpalbaran.totpalet, "
    Sql = Sql & " tmpalbaran.coddesti, tmpalbaran.codtrans "
    Sql = Sql & " from tmpalbaran "
    Sql = Sql & " order by codexped"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    b = True
    
    While Not Rs.EOF And b
        Me.pb1.Value = Me.pb1.Value + 1
        lblProgres(2).Caption = "Expediente " & DBLet(Rs!codexped, "N")
        Me.Refresh
        
        ' fecha y hora en formato de mysql
        fecha = "20" & Mid(Rs!fecexped, 7, 2) & "-" & Mid(Rs!fecexped, 4, 2) & "-" & Mid(Rs!fecexped, 1, 2)
            
        Sql = "select count(*) from albaranes where numalbar =" & DBSet(Rs!codexped, "N")
        
        If TotalRegistros(Sql) <> 0 Then
            ' borramos las tablas de lineas para volver a insertarla
            Sql = "delete from albaran_calibre where codexped = " & DBSet(Rs!codexped, "N")
            conn.Execute Sql
            
            Sql = "select * from albaran_variedad where codexped = " & DBSet(Rs!codexped, "N")
            Set RS3 = New ADODB.Recordset
            RS3.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            While Not RS3.EOF And b
                b = ActualizarCostes(RS3!NumAlbar, RS3!NumLinea, False, RS3!codforfait, "")
            
                Sql = " delete from albaran_variedad where numalbar = " & DBSet(RS3!NumAlbar, "N")
                Sql = Sql & " and numlinea = " & DBSet(RS3!NumLinea, "N")
                
                conn.Execute Sql
            Wend
        Else
'            Destino = DevuelveValor("select min(coddesti) from destinos where codclien = " & DBSet(RS!CodClien, "N"))
        
            Sql = "insert into albaran (numalbar, fechaalb, codclien, coddesti, codtrans,"
            Sql = Sql & "matriveh, matrirem, codtimer, totpalet, numerocmr, pasaridoc, codalmac) values ("
            Sql = Sql & DBSet(Rs!codexped, "N") & ","
            Sql = Sql & DBSet(fecha, "F") & ","
            Sql = Sql & DBSet(Rs!CodClien, "N") & ","
            Sql = Sql & DBSet(Rs!coddesti, "N") & ","
            Sql = Sql & DBSet(Rs!codTrans, "N") & "," ' agencia de transporte
            Sql = Sql & DBSet(Rs!matriveh, "T") & ","
            Sql = Sql & DBSet(Rs!matrirem, "T") & ","
            Sql = Sql & DBSet(Rs!Codtimer, "N") & ","
            Sql = Sql & DBSet(Rs!TotPalet, "N") & ","
            Sql = Sql & DBSet(Rs!numerocmr, "N") & ","
            Sql = Sql & "0," & DBSet(vParamAplic.Almacen, "N") & ")"
        
            conn.Execute Sql
        End If
        
        
        If b Then
            Sql = "select tmpalbaran_variedad.categori, tmpalbaran_variedad.marcafru, tmpalbaran_variedad.codprodu, "
            '[Monica]12/12/2013: agrupamos tb por codigo de comisionista
            Sql = Sql & " tmpalbaran_variedad.codvarie, tmpalbaran_variedad.codforfait, tmpalbaran_variedad.marcafru, tmpalbaran_variedad.codcalib , tmpalbaran_variedad.codcomis, count(*) palet "
            Sql = Sql & " from tmpalbaran_variedad "
            Sql = Sql & " where codexped = " & DBSet(Rs!codexped, "N")
            Sql = Sql & " group by 1,2,3,4,5,6,7,8 "
            Sql = Sql & " order by 1,2,3,4,5,6,7,8 "
            
            
            Set Rs4 = New ADODB.Recordset
            Rs4.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
            While Not Rs4.EOF
                Variedad = Format(Rs4!codprodu, "00") & Format(Rs4!codvarie, "00")
                
                MarcaFru = ""
                MarcaFru = DevuelveDesdeBDNew(cAgro, "marcas", "codmarca", "nommarca", Rs4!MarcaFru, "T")
                ' insertamos en albaran_variedad
                NumF = SugerirCodigoSiguienteStr("albaran_variedad", "numlinea", "numalbar = " & DBSet(Rs!codexped, "N"))
                
                Sql = "insert into albaran_variedad (numalbar, numlinea, codvarie, codvarco, codforfait, codmarca, categori, "
                Sql = Sql & "totpalet, numcajas, pesobrut, pesoneto, codincid"
                '[Monica]12/12/2013: introducimos el codigo de comisionista si existe
                Sql = Sql & ", codcomis "
                Sql = Sql & ") values ("
                Sql = Sql & DBSet(Rs!codexped, "N") & ","  ' numero de albaran
                Sql = Sql & DBSet(NumF, "N") & ","         ' numero de linea
                Sql = Sql & DBSet(Variedad, "N") & ","  ' variedad
                Sql = Sql & DBSet(Variedad, "N") & ","  ' variedad comercial
                Sql = Sql & DBSet(Rs4!codforfait, "T") & "," ' forfait
                Sql = Sql & DBSet(MarcaFru, "N") & "," ' marca
                Sql = Sql & DBSet(Rs4!categori, "T") & "," ' categoria
                Sql = Sql & DBSet(Rs4!Palet, "N") & ","
                Sql = Sql & "0," 'DBSet(RS!Numcajas, "N") & ","
                Sql = Sql & "0," 'DBSet(RS!pesobrut, "N") & ","
                Sql = Sql & "0," 'DBSet(RS!Pesoneto, "N") & ","
                Sql = Sql & "1" ')" ' codigo de incidencia
                '[Monica]12/12/2013:codigo de comisionista
                Sql = Sql & "," & DBSet(Rs4!codcomis, "N") & ")"
                
                conn.Execute Sql
                
                
                Sql = "insert into tmpalbaran_calibre (codexped, fecexped, codforfait, categori, marcafru, "
                Sql = Sql & " codprodu, codvarie, codcalib, codcomis, numcajas, pesobrut, pesoneto) "
                Sql = Sql & " select codexped, fecexped, codforfait, categori, marcafru, codprodu, codvarie, "
                Sql = Sql & " tmpalbaran_variedad.codcalib, tmpalbaran_variedad.codcomis, sum(tmpalbaran_variedad.numcajas), sum(tmpalbaran_variedad.pesobrut), "
                Sql = Sql & " sum(tmpalbaran_variedad.pesoneto) from tmpalbaran_variedad "
                Sql = Sql & " where codexped = " & DBSet(Rs!codexped, "N")
                Sql = Sql & " and codprodu = " & DBSet(Rs4!codprodu, "N")
                Sql = Sql & " and codvarie = " & DBSet(Rs4!codvarie, "N")  ' variedad comercial
                Sql = Sql & " and codforfait = " & DBSet(Rs4!codforfait, "T")   ' forfait
                Sql = Sql & " and marcafru = " & DBSet(Rs4!MarcaFru, "T")  ' marca
                Sql = Sql & " and categori = " & DBSet(Rs4!categori, "T")  ' categoria
                '[Monica]23/06/2010: añadido para Castelduc, quieren una linea de variedad por cada calibre
                Sql = Sql & " and codcalib = " & DBSet(Rs4!codcalib, "N") ' calibre
                '[Monica]12/12/2013:codigo de comisionista
                Sql = Sql & " and codcomis = " & DBSet(Rs4!codcomis, "N") ' comisionista
                '
                Sql = Sql & " group by 1, 2, 3, 4, 5, 6, 7 ,8, 9 "
                Sql = Sql & " order by 9 "
                
                conn.Execute Sql
                
                Sql = "select codcalib, sum(tmpalbaran_calibre.numcajas) numcajas, sum(tmpalbaran_calibre.pesobrut) pesobrut,"
                Sql = Sql & " sum(tmpalbaran_calibre.pesoneto) pesoneto from tmpalbaran_calibre"
                Sql = Sql & " where codexped = " & DBSet(Rs!codexped, "N")
                Sql = Sql & " and codprodu = " & DBSet(Rs4!codprodu, "N")
                Sql = Sql & " and codvarie = " & DBSet(Rs4!codvarie, "N")  ' variedad comercial
                Sql = Sql & " and codforfait = " & DBSet(Rs4!codforfait, "T")   ' forfait
                Sql = Sql & " and marcafru = " & DBSet(Rs4!MarcaFru, "T")  ' marca
                Sql = Sql & " and categori = " & DBSet(Rs4!categori, "T")  ' categoria
                '[Monica]23/06/2010: añadido para Castelduc, quieren una linea de variedad por cada calibre
                Sql = Sql & " and codcalib = " & DBSet(Rs4!codcalib, "N") ' añadido el calibre (tendremos una linea de variedad por calibre)
                '[Monica]12/12/2013:codigo de comisionista
                Sql = Sql & " and codcomis = " & DBSet(Rs4!codcomis, "N") ' comisionista
                '
                Sql = Sql & " group by 1 "
                Sql = Sql & " order by codcalib "
                
                Set Rs2 = New ADODB.Recordset
                Rs2.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                While Not Rs2.EOF
                    ' insertamos albaran_calidad
                    NumF1 = SugerirCodigoSiguienteStr("albaran_calibre", "numline1", "numalbar = " & DBSet(Rs!codexped, "N") & " and numlinea = " & DBSet(NumF, "N"))
                    
                    Variedad = Format(Rs4!codprodu, "00") & Format(Rs4!codvarie, "00")
                    
                    Calibre = DevuelveValor("select codcalib from calibres where codvarie = " & DBSet(Variedad, "N") & " and nomcalab = " & DBSet(Rs2!codcalib, "T"))
                    
                    Sql = "insert into albaran_calibre (numalbar,numlinea,numline1,codvarie,codcalib,numcajas,pesobrut,pesoneto) values ("
                    Sql = Sql & DBSet(Rs!codexped, "N") & ","
                    Sql = Sql & DBSet(NumF, "N") & ","
                    Sql = Sql & DBSet(NumF1, "N") & ","
                    Sql = Sql & DBSet(Variedad, "N") & ","
                    Sql = Sql & DBSet(Calibre, "N") & ","
                    Sql = Sql & DBSet(Rs2!NumCajas, "N") & ","
                    Sql = Sql & DBSet(Rs2!pesobrut, "N") & ","
                    Sql = Sql & DBSet(Rs2!Pesoneto, "N") & ")"
                                
                    conn.Execute Sql
                
                    Rs2.MoveNext
                Wend
                
                Set Rs2 = Nothing
                
                ' actualizamos los pesos y cajas de albaran_variedad
                Sql = "select if(sum(numcajas) is null,0,sum(numcajas)) numcajas, if(sum(pesobrut) is null,0,sum(pesobrut)) pesobrut, if(sum(pesoneto) is null,0,sum(pesoneto)) pesoneto "
                Sql = Sql & " from albaran_calibre where numalbar= " & DBSet(Rs!codexped, "N")
                Sql = Sql & " and numlinea = " & DBSet(NumF, "N")
                
                Set Rs2 = New ADODB.Recordset
                Rs2.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                While Not Rs2.EOF And b
                    Sql2 = "update albaran_variedad set numcajas = " & DBSet(Rs2!NumCajas, "N")
                    Sql2 = Sql2 & ", pesobrut = " & DBSet(Rs2!pesobrut, "N")
                    Sql2 = Sql2 & ", pesoneto = " & DBSet(Rs2!Pesoneto, "N")
                    Sql2 = Sql2 & " where numalbar = " & DBSet(Rs!codexped, "N")
                    Sql2 = Sql2 & " and numlinea = " & DBSet(NumF, "N")
                    
                    conn.Execute Sql2
                
                    b = ActualizarCostes(DBLet(Rs!codexped, "N"), CInt(NumF), True, Rs4!codforfait, "")
                
                    Rs2.MoveNext
                Wend
                
                Set Rs2 = Nothing
            
                Rs4.MoveNext
            
            Wend
            Set Rs4 = Nothing
            
            ' actualizamos las observaciones del expediente con los numeros de palets que me hayan introducido
            Sql = "select distinct numpalet from tmpalbaran_variedad where codexped = " & DBSet(Rs!codexped, "N")
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Lista = ""
            
            While Not Rs2.EOF And b
                Lista = Lista & DBSet(Rs2!numpalet, "N") & ","
            
                Rs2.MoveNext
            Wend
            
            Set Rs2 = Nothing
        
            If Lista <> "" Then
                Lista = Mid(Lista, 1, Len(Lista) - 1)
                
                Sql = "update albaran set observac = " & DBSet(Lista, "T") & " where numalbar = " & DBSet(Rs!codexped, "N")
                conn.Execute Sql
            End If
        End If
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing

    pb1.visible = False
    lblProgres(1).Caption = ""
    lblProgres(2).Caption = ""

    CargarExpedientes = True And b
    Exit Function
    
eCargarExpedientes:
    MuestraError Err.Number, "Cargar expedientes", Err.Description
End Function


