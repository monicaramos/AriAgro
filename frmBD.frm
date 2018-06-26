VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acceso a MYSQL"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   Icon            =   "frmBD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   7785
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameNuevaEmpresa 
      Height          =   7110
      Left            =   45
      TabIndex        =   12
      Top             =   45
      Width           =   7650
      Begin VB.CheckBox Check1 
         Caption         =   "Inicializar Contadores"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   300
         TabIndex        =   26
         Top             =   4875
         Width           =   3060
      End
      Begin VB.TextBox Text2 
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
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   3150
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   0
         Tag             =   "admon"
         Top             =   2115
         Width           =   1545
      End
      Begin VB.TextBox Text2 
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
         Left            =   4725
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   4290
         Width           =   1350
      End
      Begin VB.TextBox Text2 
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
         Left            =   2025
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   4290
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
         Height          =   435
         Left            =   6270
         TabIndex        =   7
         Top             =   6375
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
         Height          =   435
         Left            =   5055
         TabIndex        =   6
         Top             =   6375
         Width           =   1065
      End
      Begin VB.TextBox Text2 
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
         Index           =   2
         Left            =   2025
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   3825
         Width           =   555
      End
      Begin VB.TextBox Text2 
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
         Left            =   2025
         MaxLength       =   15
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   3375
         Width           =   1815
      End
      Begin VB.TextBox Text2 
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
         Left            =   2025
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "Text2"
         Top             =   2910
         Width           =   5340
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   285
         TabIndex        =   24
         Top             =   5370
         Width           =   7140
         _ExtentX        =   12594
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Realice previamente una copia de seguridad de ariagro"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   330
         Left            =   420
         TabIndex        =   27
         Top             =   1095
         Width           =   6675
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   330
         TabIndex        =   25
         Top             =   5910
         Width           =   6975
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Este proceso crea una base de datos nueva."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   330
         Left            =   450
         TabIndex        =   23
         Top             =   780
         Width           =   6675
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Se recomienda que no haya nadie trabajando en la actual campaña. "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   510
         Index           =   5
         Left            =   495
         TabIndex        =   22
         Top             =   1410
         Width           =   6630
      End
      Begin VB.Label Label3 
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   345
         Left            =   2025
         TabIndex        =   21
         Top             =   2115
         Width           =   2235
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Creación nueva campaña"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   435
         Index           =   4
         Left            =   1395
         TabIndex        =   20
         Top             =   225
         Width           =   4740
      End
      Begin VB.Label Label6 
         Caption         =   "label6"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   330
         TabIndex        =   18
         Top             =   5625
         Width           =   7020
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha fin"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   3705
         TabIndex        =   17
         Top             =   4350
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha inicio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   300
         TabIndex        =   16
         Top             =   4350
         Width           =   1260
      End
      Begin VB.Label Label4 
         Caption         =   "Número BD"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   300
         TabIndex        =   15
         Top             =   3900
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre corto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   300
         TabIndex        =   14
         Top             =   3435
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre campaña"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   300
         TabIndex        =   13
         Top             =   2970
         Width           =   1830
      End
   End
   Begin VB.Frame Frame4 
      Height          =   3135
      Left            =   45
      TabIndex        =   19
      Top             =   135
      Width           =   6735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   3
      Left            =   5040
      TabIndex        =   11
      Top             =   0
      Width           =   1185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Inserta datos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   435
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   2400
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Borrar datos en BD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   435
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nueva BD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1755
   End
End
Attribute VB_Name = "frmBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const TamanyoMaximoParaSQL = 50000

Public ChequeaFacturas As Byte
    '0 No
    '1 Solo clientes
    '2 Solo proveedores
    '3 Ambas
    
Public Opcion As Byte
    '0 Nueva BD
    '1 Borrar datos
    '3 Insertar desde archivos tra
    
    '5 Nueva pero desde ariconta
    
    '6 Le diremos un fichero TRA y tendra que traspasarlo
    
    
    ' TESORERIA
    '7.- Nueva
    '8.- Eliminar datos
    '9.- Eliminar estructura y vinculacion
    '10.- TRASPASO de tesoreria
    
    
    '20.- Recargar ampconce
    
    '21.- Renumerar FRAPRO
    '22.- Cambio cta contable
    
    
Public Intercambio As String


Dim PrimeraVez As Boolean

Dim Tam2 As Long
Dim Tamanyo As Long
Dim NombreArchivo As String
Dim PrimeraLinea As Boolean
Dim SQL As String
Dim Insert As String
Dim Final As String
Dim Linea As String
Dim TablaAnt As String
Dim NF As Integer
Dim Errores  As String
Dim ErroresAux As String
Dim path As String

Dim ContadorInserciones As Long

Dim LineaInsert As String
Dim AuxCad As String


Dim RSCta As ADODB.Recordset

Dim Cnn As Connection


Dim SePuedeSalir As Boolean

'Para las inserciones masivas
Dim CadenaMasiva As String
Dim DatosMasivo As Long
Dim BdNueva As String

Dim NumTablas As Integer



'-------------------------------------
'Abrir conexion CNN


Private Function AbrirConexion(Usuario As String, Pass As String, Conta As String) As Boolean
Dim Cad As String
On Error GoTo EAbrirConexion

    AbrirConexion = False
    Set Cnn = New Connection
    'Conn.CursorLocation = adUseClient
    Cnn.CursorLocation = adUseServer
    Cad = "DSN=vUsuarios;DESC=MySQL ODBC 3.51 Driver DSN;DATABASE=" & Conta & ";SERVER=" & vConfig.SERVER
    Cad = Cad & ";UID=" & Usuario & ";PASSWORD=" & Pass & ";PORT=3306;OPTION=3;STMT=;"
    Cad = Cad & ";Persist Security Info=true"
    
    Cnn.ConnectionString = Cad
    Cnn.Open
    AbrirConexion = True
    Exit Function
    
EAbrirConexion:
    MuestraError Err.Number, "Abrir conexión " & Conta, Err.Description
End Function


Private Sub Command8_Click()
'DROP DATABASE
On Error Resume Next
SQL = "Va a eliminar una  BD" & vbCrLf & "¿Desea continuar?"
If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
    SQL = InputBox("Nombre BD")
    If SQL <> "" Then
        If LCase(SQL) = "conta" Then
            MsgBox "BD ppal", vbExclamation
            Exit Sub
        End If
        SQL = "DROP DATABASE " & SQL
        Cnn.Execute SQL
        If Err.Number <> 0 Then
            MuestraError Err.Number, "Eliminando BD"
        End If
    End If
End If

End Sub


Private Sub GeneracionNuevaBD(DesdeCopiarDatosDeOtraEmpresa As Boolean)
    
    
    NombreArchivo = "ariagro" & Text2(2).Text
    BdNueva = NombreArchivo
    If ComprobarEmpresa(NombreArchivo) Then
        MsgBox "Ya existe la empresa: ariagro" & Text2(2).Text, vbExclamation
        Exit Sub
    End If
    
    
    If Not GeneraNuevaBD Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    CrearEstructura
    
    If Not CrearTablas Then Exit Sub
    

    
'    Set Cnn = Nothing
    FormularioOK = "OK"
    Screen.MousePointer = vbDefault

End Sub


Private Function GeneraNuevaBD() As Boolean
On Error Resume Next
       GeneraNuevaBD = False
        SQL = "CREATE DATABASE " & NombreArchivo
        conn.Execute SQL
        If Err.Number <> 0 Then
            MuestraError Err.Number, "Creando BD"
        Else
            GeneraNuevaBD = True
        End If
End Function

Private Sub cmdAceptar_Click()
Dim OK As Boolean
Dim Mens As String

    
    For NF = 0 To Me.Text2.Count - 1
        Text2(NF).Text = Trim(Text2(NF).Text)
        If Text2(NF).Text = "" Then
            MsgBox "Campo " & Label4(NF).Caption & " esta vacio", vbExclamation
            Exit Sub
        End If
    Next NF

    If Not IsNumeric(Text2(2).Text) Then
        MsgBox "Número de empresa tiene que ser numérico, obviamente", vbExclamation
        Exit Sub
    End If
    
    If Text2(3).Text = "" Or Text2(4).Text = "" Then
        MsgBox "Debe introducir la Fecha Inicio y Fecha Fin de campaña.", vbExclamation
        Exit Sub
    Else
        If Not EsFechaOK(Text2(3).Text) Then
            MsgBox "Fecha inicio campaña incorrecta", vbExclamation
            Exit Sub
        End If
        
        If Not EsFechaOK(Text2(4).Text) Then
            MsgBox "Fecha fin campaña incorrecta", vbExclamation
            Exit Sub
        End If
        
        If CDate(Text2(3).Text) > CDate(Text2(4).Text) Then
            MsgBox "Fecha Inicio Campaña es superior a la Fecha Fin de Campaña", vbExclamation
            Exit Sub
        End If
    
'[Monica]24/12/2012: comprobamos que estamos en la campaña actual, (no en la ultima campaña como hacia el proceso anteriormente)
'        Mens = "Comprobar que estamos en la Ultima Campaña: " & vbCrLf & vbCrLf
'        If Not ComprobarUltimaCampanya(Mens) Then
'            MsgBox Mens, vbExclamation
'            Exit Sub
'        End If
       
        If InStr(1, conn.ConnectionString, "DATABASE=ariagro;") = 0 Then
            MsgBox "No se encuentra en la campaña actual para realizar el proceso. Revise.", vbExclamation
            Exit Sub
        End If
        
        Mens = "Comprobar Fechas de Campaña: " & vbCrLf & vbCrLf
        Select Case ComprobarFechasCampanya(Mens)
            Case 1
                Mens = Mens & vbCrLf & vbCrLf & "Llame a soporte."
                MsgBox Mens, vbExclamation
                Exit Sub
            Case 2
                If MsgBox(Mens & vbCrLf & vbCrLf & "¿ Desea continuar ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub
        End Select
    End If
        
    
    FormularioOK = "Va a generar una nueva campaña: " & Text2(0).Text
    FormularioOK = FormularioOK & vbCrLf & "¿ Desea continuar? "
    If MsgBox(FormularioOK, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    FormularioOK = ""
    OK = False
    Label6.visible = True
    Label6.Caption = "Generando estructura de BD"
    Me.Refresh
    
    conn.BeginTrans
    Cnn.BeginTrans
    
    GeneracionNuevaBD True
    'Si ha ido bien salimos
    OK = (FormularioOK = "OK")
    
    If OK Then
    
        Label6.Caption = ""
        Label6.Refresh
        '[Monica]24/12/2012: No insertamos datos basicos en la base de datos nueva, sino que borramos de la campaña actual.
'        OK = InsercionDatos
        OK = InsercionDatosOrigen
        If OK Then OK = BorradoDatos
    End If
    Screen.MousePointer = vbDefault
    If OK Then
        SePuedeSalir = True
        conn.CommitTrans
        Cnn.CommitTrans
        
        LeerParametros
        
        MsgBox "Proceso realizado correctamente.", vbExclamation
    Else
        conn.RollbackTrans
        Cnn.RollbackTrans
        
        MsgBox "No se ha realizado el proceso. Llame a Ariadna.", vbExclamation
    End If
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Label6.Caption = ""
    If PrimeraVez Then
        PrimeraVez = False
        
        PonerFoco Text2(5)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Long
Dim W As Long
    PrimeraVez = True
    SePuedeSalir = False
    Label6.visible = False
    Label6.Caption = ""
    Label1.Caption = ""
    Me.Refresh
    Pb1.visible = False
    Me.Check1.Value = 0
    
    ActivarCLAVE
    
    CargarDatosComandos
    
    frameNuevaEmpresa.visible = True
End Sub



'--------------------------------------------------------------------
'
'                    Crear estructura BD
'
'--------------------------------------------------------------------

Private Sub CrearEstructura()

    Errores = ""
    ErroresAux = ""
    
    If Not CargaFicheroEstructura(False) Then Exit Sub

End Sub


Private Function CargaFicheroEstructura(Tesoreria As Boolean) As Boolean
    On Error GoTo eCargaFicheroEstructura
    CargaFicheroEstructura = False
    Final = ""
    
    
    'Nuevo 17 MARZO 2005
    '--------------------------------------------------------------
    If CargaGeneracionEstructura(Tesoreria) Then
        CargaFicheroEstructura = True
        Exit Function
    End If
    
    
eCargaFicheroEstructura:
        MuestraError Err.Number, "Cargando fichero estructura"
End Function



Private Function CargaGeneracionEstructura(Tesoreria As Boolean) As Boolean
Dim L As Collection
    On Error GoTo ECargaGeneracionEstructura
    
    CargaGeneracionEstructura = False
    
    Set RSCta = New ADODB.Recordset
        
     NF = FreeFile
     Open vConfig.FichGene For Output As #NF
     
     Pb1.visible = True
     Label6.Caption = "Generando la estructura en fichero."
     Me.Refresh
     
     
     
     Set L = New Collection
     RSCta.Open "SHOW TABLES", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
     While Not RSCta.EOF

         L.Add CStr(RSCta.Fields(0))
         RSCta.MoveNext
     Wend
     RSCta.Close
     espera 0.2
     '++monica
     NumTablas = L.Count
     
     CargarProgres Pb1, NumTablas
     
     For ContadorInserciones = 1 To L.Count
         SQL = L.item(ContadorInserciones)
         IncrementarProgres Pb1, 1
         DoEvents
         RSCta.Open "SHOW CREATE TABLE " & SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
         If Not RSCta.EOF Then
             SQL = RSCta.Fields(1)
             'Voy a quitar las `
             Do
                 Tam2 = InStr(1, SQL, "`")
                 If Tam2 > 0 Then SQL = Mid(SQL, 1, Tam2 - 1) & Mid(SQL, Tam2 + 1)
             Loop Until Tam2 = 0
             Final = Final & SQL & ";     " & vbCrLf
             Print #NF, SQL & ";     " & vbCrLf
     
         End If
         RSCta.Close
         
     Next ContadorInserciones
     Close #NF
     
     Pb1.visible = False
     Me.Refresh
     
     CargaGeneracionEstructura = True
ECargaGeneracionEstructura:
    If Err.Number <> 0 Then MuestraError Err.Number, "Leyendo estructura SHOW TABLES"
    Set RSCta = Nothing
    Set L = Nothing
End Function

Private Sub Label4_DblClick(Index As Integer)
 If Index = 2 Then Text2(2).Enabled = Not Text2(2).Enabled
End Sub

Private Sub Label4_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 2 And Button = 2 Then
        Text2(2).Enabled = True
        Text2(2).BackColor = vbWhite
    End If
End Sub

Private Sub CargarDatosComandos()
Dim Rs As Recordset

    For NF = 0 To Me.Text2.Count - 1
        Text2(NF).Text = ""
    Next NF
    
    Label2(4).visible = True
    
    If AbrirConexion(vConfig.User, vConfig.password, "Usuarios") Then
        Set Rs = New ADODB.Recordset
        NF = 1
        SQL = "select max(codempre) from empresasariagro"
        Rs.Open SQL, Cnn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not Rs.EOF Then
            If Not IsNull(Rs.Fields(0)) Then
                NF = Rs.Fields(0)
            End If
        End If
        Rs.Close
        Set Rs = Nothing
        Text2(2).Text = NF + 1
    End If
    
    CargarFechasCampanya
    
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub Text2_LostFocus(Index As Integer)
    Text2(Index).Text = Trim(Text2(Index).Text)
    If Text2(Index).Text = "" Then Exit Sub
    
    
    Select Case Index
        Case 3, 4
            PonerFormatoFecha Text2(Index)
        Case 5
            If Text2(Index).Text = "" Then Exit Sub
            If Trim(Text2(Index).Text) <> Trim(Text2(Index).Tag) Then
                MsgBox "    ACCESO DENEGADO    ", vbExclamation
                Text2(Index).Text = ""
                PonerFoco Text2(Index)
            Else
                DesactivarCLAVE
                PonerFoco Text2(0)
            End If
        
    End Select
'    If Index > 2 Then
'        If Not EsFechaOK(Text2(Index)) Then
'            MsgBox "Fecha incorrecta: " & Text2(Index).Text, vbExclamation
'            Text2(Index).Text = ""
'            Text2(Index).SetFocus
'        End If
'    End If
End Sub



Private Function InsercionDatos() As Boolean
'incluiremos en este procedimiento todas las tablas maestras que vayan apareciendo

Dim Rs As Recordset

    On Error GoTo EInsercionDatos
    
    InsercionDatos = False
    
    Label6.Caption = "Cargando datos de tablas maestras."
    Pb1.visible = True
    CargarProgres Pb1, 116 ' 116 tablas maestras hasta el momento
    Me.Refresh
    
    ' No tenemos en cuenta las claves referenciales
    SQL = "set foreign_key_checks = 0"
    conn.Execute SQL
    
    '1 agencias
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".agencias SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".agencias "
    conn.Execute SQL
    
    '2 banpropi
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".banpropi SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".banpropi "
    conn.Execute SQL
    
    '3 cadenas
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".cadenas SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".cadenas "
    conn.Execute SQL
    
    '4 capacidad
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".capacida SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".capacida "
    conn.Execute SQL
    
    '5 clases
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".clases SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".clases "
    conn.Execute SQL
    
    '6 confenva
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".confenva SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".confenva "
    conn.Execute SQL
    
    '7 confmedi
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".confmedi SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".confmedi "
    conn.Execute SQL
    
    '8 confpale
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".confpale SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".confpale "
    conn.Execute SQL
    
    '9 confpres
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".confpres SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".confpres "
    conn.Execute SQL
    
    '10 conftipo
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".conftipo SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".conftipo "
    conn.Execute SQL
    
    '11 forpago
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".forpago SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".forpago "
    conn.Execute SQL
    
    '12 grupopro
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".grupopro SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".grupopro "
    conn.Execute SQL
    
    '13 inciden
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".inciden SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".inciden "
    conn.Execute SQL
    
    '14 marcas
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".marcas SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".marcas "
    conn.Execute SQL
    
    '15 nombcoste
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".nombcoste SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".nombcoste "
    conn.Execute SQL
    
    '16 paises
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".paises SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".paises "
    conn.Execute SQL
    
    '17 scartas
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".scartas SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".scartas "
    conn.Execute SQL
    
    '18 scryst
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".scryst SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".scryst "
    conn.Execute SQL
    
    '19 sdirpr
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".sdirpr SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".sdirpr "
    conn.Execute SQL
    
    '20 sfamia
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".sfamia SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".sfamia "
    conn.Execute SQL
    
'
'    Sql = "UPDATE " & Trim(BdNueva) & ".sparam SET codparam = " & DBSet(Text2(2).Text, "T")
'    Conn.Execute Sql
    
    '21 sprvar
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".sprvar SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".sprvar "
    conn.Execute SQL
    
    '22 stipar
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".stipar SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".stipar "
    conn.Execute SQL
    
'--monica:10/02/2009 la stipom pasa a estar en la bd usuarios
'    '24 stipom
'    IncrementarProgres Pb1, 1
'    Sql = "INSERT INTO " & Trim(BdNueva) & ".stipom SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".stipom "
'    conn.Execute Sql
    
    '23 sunida
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".sunida SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".sunida "
    conn.Execute SQL
    
    '24 tarifas
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".tarifas SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".tarifas "
    conn.Execute SQL
    
    '25 tipomer
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".tipomer SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".tipomer "
    conn.Execute SQL
    
    '26 tmpcopiascmr
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".tmpcopiascmr SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".tmpcopiascmr "
    conn.Execute SQL
    
    
    '27 empresas
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".empresas SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".empresas "
    conn.Execute SQL
    
    SQL = "UPDATE " & Trim(BdNueva) & ".empresas SET codempre = " & DBSet(Text2(2).Text, "T")
    SQL = SQL & " , fechaini = " & DBSet(Text2(3).Text, "F") & ", fechafin = " & DBSet(Text2(4).Text, "F")
    conn.Execute SQL
    
    '28 salmpr
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".salmpr SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".salmpr "
    conn.Execute SQL
    
    '29 sparam
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".sparam SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".sparam "
    conn.Execute SQL
    
    '[Monica]08/06/2012: tipos de variedad para las lineas de albaranes de comercial
    '30 tipos de variedad
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".tipovarie SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".tipovarie "
    conn.Execute SQL
    
    '31 clientes
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".clientes SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".clientes "
    conn.Execute SQL
    
    '32 destinos
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".destinos SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".destinos "
    conn.Execute SQL
    
    '33 productos
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".productos SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".productos "
    conn.Execute SQL

    '34 variedades
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".variedades SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".variedades "
    conn.Execute SQL
    
    '35 variedades anecoop
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".variane SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".variane "
    conn.Execute SQL
    
    '36 calibres
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".calibres SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".calibres "
    conn.Execute SQL
    
'    'calibres variedad
'    IncrementarProgres Pb1, 1
'    Sql = "INSERT INTO " & Trim(BdNueva) & ".calibvar SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".calibvar "
'    Conn.Execute Sql
    
    '37 proveedores
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".proveedor SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".proveedor "
    conn.Execute SQL
    
    '38 articulos
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".sartic SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".sartic "
    conn.Execute SQL
    
    '39 almacenes
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".salmac SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".salmac "
    conn.Execute SQL
    
    '40 forfaits
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".forfaits SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".forfaits "
    conn.Execute SQL
    
    '41 forfaits-lineas
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".forfaits_envases SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".forfaits_envases "
    conn.Execute SQL
    
    '42 costes de confeccion
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".forfaits_costes SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".forfaits_costes "
    conn.Execute SQL
    
    '43 tipos de iva
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".tiposiva SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".tiposiva "
    conn.Execute SQL
    
    '44 codigoean
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".codigoean SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".codigoean "
    conn.Execute SQL
    
    '45 salarios
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".salarios SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".salarios "
    conn.Execute SQL
    
    '46 straba
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".straba SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".straba "
    conn.Execute SQL
    
    '----------------------------------------
    'tablas de recoleccion
    '----------------------------------------
    '47 rcalidad
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rcalidad SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rcalidad "
    conn.Execute SQL
    
    '48 rcapataz
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rcapataz SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rcapataz "
    conn.Execute SQL
    
    '49 rcoope
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rcoope SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rcoope "
    conn.Execute SQL
    
    '50 rincidencia
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rincidencia SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rincidencia "
    conn.Execute SQL
    
    '51 rseccion
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rseccion SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rseccion "
    conn.Execute SQL
    
    '52 rparam
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rparam SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rparam "
    conn.Execute SQL
    
    '53 rzonas
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rzonas SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rzonas "
    conn.Execute SQL
    
    '54 rpueblos
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rpueblos SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rpueblos "
    conn.Execute SQL
    
    '55 rpartida
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rpartida SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rpartida "
    conn.Execute SQL
    
    '56 rportespobla
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rportespobla SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rportespobla "
    conn.Execute SQL
    
    
    '57 rsituacion
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rsituacion SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rsituacion "
    conn.Execute SQL
    
    '58 rsituacioncampo
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rsituacioncampo SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rsituacioncampo "
    conn.Execute SQL
    
    '59 rtransporte
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rtransporte SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rtransporte "
    conn.Execute SQL
    
    '60 rtarifatra
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rtarifatra SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rtarifatra "
    conn.Execute SQL
    
    '61 rsocios
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rsocios SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rsocios "
    conn.Execute SQL
    
    '62 rsocios_seccion
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rsocios_seccion SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rsocios_seccion "
    conn.Execute SQL
    
    '63 rsocios_telefonos
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rsocios_telefonos SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rsocios_telefonos "
    conn.Execute SQL
    
    '64 rsocios_pozos
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rsocios_pozos SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rsocios_pozos "
    conn.Execute SQL
    
    
    '65 rdesarrollo
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rdesarrollo SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rdesarrollo "
    conn.Execute SQL
    
    '66 rplantacion
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rplantacion SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rplantacion "
    conn.Execute SQL
    
    '67 rriego
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rriego SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rriego "
    conn.Execute SQL
    
    '68 rtierra
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rtierra SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rtierra "
    conn.Execute SQL
    
    '69 rseguroopcion
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rseguroopcion SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rseguroopcion "
    conn.Execute SQL
    
    '70 rproceriego
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rproceriego SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rproceriego "
    conn.Execute SQL
    
    '71 rpatronpie
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rpatronpie SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rpatronpie "
    conn.Execute SQL
    
    
    '72 rcampos
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rcampos SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rcampos "
    conn.Execute SQL
    
    '73 rcampos_clasif
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rcampos_clasif SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rcampos_clasif "
    conn.Execute SQL
    
    '74 rconcepgasto
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rconcepgasto SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rconcepgasto "
    conn.Execute SQL
    
    '75 rconcepgastonom
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rconcepgastonom SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rconcepgastonom "
    conn.Execute SQL
    
    '76 rcopropiedad
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rcopropiedad SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rcopropiedad "
    conn.Execute SQL
    
    '77 rdeposito
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rdeposito SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rdeposito "
    conn.Execute SQL
    
    '78 rpozos
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rpozos SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rpozos "
    conn.Execute SQL
    
    '79 rcalidad_calibrador
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rcalidad_calibrador SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rcalidad_calibrador "
    conn.Execute SQL
    
    '80 advfamia
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".advfamia SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".advfamia "
    conn.Execute SQL
    
    '81 advartic
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".advartic SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".advartic "
    conn.Execute SQL
    
    
    '82 advartic_salmac
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".advartic_salmac SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".advartic_salmac "
    conn.Execute SQL
    
    '83 trztipos
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".trztipos SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".trztipos "
    conn.Execute SQL
    
    '84 trzareas
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".trzareas SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".trzareas "
    conn.Execute SQL
    
    '85 trzdispositivos
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".trzdispositivos SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".trzdispositivos "
    conn.Execute SQL
    
    '86 trzlineas_rfid
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".trzlineas_rfid SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".trzlineas_rfid "
    conn.Execute SQL
    
    '87 trzlineas_confeccion
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".trzlineas_confeccion SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".trzlineas_confeccion "
    conn.Execute SQL
    
    '88 rplagasaux
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rplagasaux SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rplagasaux "
    conn.Execute SQL
    
    '[Monica] 14/02/2011 : dos nuevas tablas de adv alzira
    
    '89 advmatactiva
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".advmatactiva SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".advmatactiva "
    conn.Execute SQL
    
    '90 advartic_matactiva
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".advartic_matactiva SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".advartic_matactiva "
    conn.Execute SQL
    
    
    '91 raporreparto
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".raporreparto SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".raporreparto "
    conn.Execute SQL

    '92 raportacion
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".raportacion SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".raportacion "
    conn.Execute SQL
    
    '93 rbonifica
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rbonifica SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rbonifica "
    conn.Execute SQL
    
    '94 rbonifica_lineas
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rbonifica_lineas SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rbonifica_lineas "
    conn.Execute SQL
    
    '95 rcampos_cooprop
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rcampos_cooprop SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rcampos_cooprop "
    conn.Execute SQL
    
    '96 rcampos_parcelas
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rcampos_parcelas SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rcampos_parcelas "
    conn.Execute SQL
    
    '97 rpozos_cooprop
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rpozos_cooprop SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rpozos_cooprop "
    conn.Execute SQL

    '98 rtipoapor
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rtipoapor SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rtipoapor "
    conn.Execute SQL

    '99 rtipopozos
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rtipopozos SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rtipopozos "
    conn.Execute SQL

    '100 rtarifaett
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rtarifaett SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rtarifaett "
    conn.Execute SQL

    '101 advtrata
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".advtrata SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".advtrata "
    conn.Execute SQL

    '[Monica]20/10/2011: incrementamos en un año la fecha de inicio y fin de tratamientos
    SQL = "UPDATE " & Trim(BdNueva) & ".advtrata SET fechaini = date_add(fechaini, interval 1 year), "
    SQL = SQL & " fechafin = date_add(fechafin, interval 1 year)"
    conn.Execute SQL

    '102 advtrata_lineas
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".advtrata_lineas SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".advtrata_lineas "
    conn.Execute SQL

' actualizamos los campos de seguro de la campaña anterior de esta nueva bd e inicializamos los de esta campaña
    SQL = "UPDATE " & Trim(BdNueva) & ".rcampos SET codseguroant = codseguro, aseguradoant = asegurado, kilosaseant = kilosase, costeseguroant = costeseguro, "
    SQL = SQL & " codseguro = null, asegurado = 0, kilosase = null, costeseguro = null "
    conn.Execute SQL
    
    '[Monica]24/10/2011: nueva tabla de hco de campos (un campo por qué socios ha pasado)
    '103 rcampos_hco
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rcampos_hco SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rcampos_hco "
    conn.Execute SQL
    
    '[Monica]10/11/2011: nueva tabla entrega de ficha de cultivo de campos
    '104 entrega ficha de cultivo de campos
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rfichculti SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rfichculti "
    conn.Execute SQL
    
    '[Monica]17/11/2011: nueva tabla de rglobalgap
    '105 globalgap
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rglobalgap SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rglobalgap "
    conn.Execute SQL
    
    
    '[Monica]28/11/2011: nueva tabla de clientes_precio
    '106 clientes_precio
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".clientes_precio SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".clientes_precio "
    conn.Execute SQL
    
    
    '[Monica]27/12/2011: nueva tabla de rcampos_gastos
    '107 rcampos_gastos
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rcampos_gastos SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rcampos_gastos "
    conn.Execute SQL
    
    '[Monica]07/02/2012: tabla de lineas de envases de palets, faltaba
    '108 confpale_envases
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".confpale_envases SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".confpale_envases "
    conn.Execute SQL
    
    '[Monica]07/02/2012: tabla de lineas de envases de palets, faltaba
    '109 raporhco
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".raporhco SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".raporhco "
    conn.Execute SQL
    
    '[Monica]29/02/2012: tabla de rpozos_campos, faltaba
    '110 rpozos_campos
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".rpozos_campos SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".rpozos_campos "
    conn.Execute SQL
    
    '[Monica]10/04/2012: tabla de cctipocoste tabla maestra basica
    '111 cctipocoste
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".cctipocoste  SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".cctipocoste "
    conn.Execute SQL
    
    '[Monica]11/04/2012: tabla de ccareas
    '112 ccareas
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".ccareas  SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".ccareas "
    conn.Execute SQL
    
    '[Monica]11/04/2012: tabla de ccconcostes
    '113 ccconcostes
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".ccconcostes  SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".ccconcostes "
    conn.Execute SQL
    
    '[Monica]11/04/2012: tabla de cclinconf
    '114 cclinconf
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".cclinconf  SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".cclinconf "
    conn.Execute SQL
    
    '[Monica]31/07/2012: tabla de cchorario
    '115 cchorario
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".cchorario  SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".cchorario "
    conn.Execute SQL
    
    '[Monica]31/07/2012: tabla de cchorario_tramo
    '116 cchorario_tramo
    IncrementarProgres Pb1, 1
    DoEvents
    SQL = "INSERT INTO " & Trim(BdNueva) & ".cchorario_tramo  SELECT * FROM " & Trim(vEmpresa.BDAriagro) & ".cchorario_tramo "
    conn.Execute SQL
    
    Me.Refresh
    DoEvents
        
    '----------
    'parametros
    '----------------------------------------
    'Insertamos en tabla empresas
    SQL = "INSERT INTO empresasariagro (codempre, nomempre, nomresum, Usuario, Pass, ariagro) VALUES ("
    SQL = SQL & Text2(2).Text & ",'" & Text2(0).Text & "','" & Text2(1).Text
    SQL = SQL & "','',''," & DBSet(BdNueva, "T") & ")"
    Cnn.Execute SQL
    
    ' inicializamos los contadores de la stipom
    If Me.Check1.Value Then
        SQL = "update stipom set contador = 0"
        
        Cnn.Execute SQL
    End If
    
    InsercionDatos = True
    
    ' No tenemos en cuenta las claves referenciales
    SQL = "set foreign_key_checks = 1"
    conn.Execute SQL
    
    Exit Function
    
EInsercionDatos:
    MuestraError Err.Number, Label6.Caption & vbCrLf & vbCrLf & Err.Description
    
    ' No tenemos en cuenta las claves referenciales
    SQL = "set foreign_key_checks = 1"
    conn.Execute SQL

End Function


Private Sub FinalizaPorErrores(Donde As String)
    ErroresAux = ErroresAux & vbCrLf & vbCrLf & Donde & vbCrLf
    ErroresAux = ErroresAux & LineaInsert & vbCrLf & Err.Description
    frmErrores2.Opcion = 0
    frmErrores2.Text1 = ErroresAux
    frmErrores2.Show vbModal
    End
End Sub


Public Function ComprobarEmpresa(Empre As String) As Boolean
Dim Cad As String
Dim Conne As Connection
Dim Rs As ADODB.Recordset
Dim itemX As ListItem

    On Error Resume Next
    Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=" & Empre & ";SERVER=" & vConfig.SERVER & ";"
    Cad = Cad & ";UID=" & vConfig.User
    Cad = Cad & ";PWD=" & vConfig.password
    

    Set Conne = New Connection
    Conne.CursorLocation = adUseServer
    Conne.ConnectionString = Cad
    Conne.Open
    If Err.Number <> 0 Then
        'Ha sido error
        ComprobarEmpresa = False
    Else
        Set Rs = New ADODB.Recordset
        Rs.Open "Select * from Empresa", Conne, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Err.Number = 0 Then
            If Not Rs.EOF Then

            End If
            Rs.Close
        End If
        Set Rs = Nothing
        ComprobarEmpresa = True
    End If
    Set Conne = Nothing
End Function


Private Function CrearTablas() As Boolean
Dim SQL As String
Dim NF As Long
Dim J As Integer
Dim NVeces As Integer

    On Error Resume Next
    
    SQL = "USE ariagro" & Text2(2).Text
    conn.Execute SQL
    
    CrearTablas = True
    
    NF = FreeFile
    J = 0
    
    Pb1.visible = True
    Label6.Caption = "Creando Tablas en nueva Base de Datos "
    Label1.Caption = ""
    Me.Refresh
    CargarProgres Pb1, (NumTablas)
    
    
    NVeces = 0
    While J < NumTablas And NVeces < 6
        Open vConfig.FichGene For Input As #NF
        Do Until EOF(NF) Or J = NumTablas
            Line Input #NF, SQL
            If SQL <> "" Then
                conn.Execute SQL
                If Err.Number = 0 Then
                    J = J + 1
                    IncrementarProgres Pb1, 1
                    DoEvents
                Else
                    Err.Number = 0
                    
                End If
                
            End If
        Loop
        Close #NF
        NVeces = NVeces + 1
    Wend
    
    If J < NumTablas Then
        MsgBox "No se ha creado correctamente la nueva base de datos. Llame a soporte.", vbExclamation
        CrearTablas = False
    End If
    Pb1.visible = False
End Function

Private Sub ActivarCLAVE()
Dim i As Integer
    
    For i = 0 To 4
        Text2(i).Enabled = False
    Next i
    Me.Check1.Enabled = False
    Text2(5).Enabled = True
    ' fechas siempre inhibidas
    Text2(3).Enabled = False
    Text2(4).Enabled = False
    
    cmdAceptar.Enabled = False
    cmdCancel.Enabled = True

End Sub

Private Sub DesactivarCLAVE()
Dim i As Integer

    For i = 0 To 4
        If i <> 2 Then Text2(i).Enabled = True
    Next i
    Me.Check1.Enabled = True
    
    Text2(5).Enabled = False
    ' fechas siempre inhibidas
    Text2(3).Enabled = False
    Text2(4).Enabled = False
    
    cmdAceptar.Enabled = True
End Sub

Private Sub CargarFechasCampanya()

    Text2(3).Text = Format(CStr(DateAdd("yyyy", 1, CDate(vParam.FecIniCam))), "dd/mm/yyyy")
    Text2(4).Text = Format(CStr(DateAdd("yyyy", 1, CDate(vParam.FecFinCam))), "dd/mm/yyyy")

End Sub

Private Function ComprobarFechasCampanya(ByRef Mens As String) As Byte
Dim Rs As ADODB.Recordset, Rs1 As ADODB.Recordset
Dim vSQL As String, vSql2 As String
Dim Continuar As Boolean
Dim FechasFinCamp As Collection
Dim MaxFec As Date
Dim i As Integer
    
    On Error GoTo eComprobarFechasCampanya

    ComprobarFechasCampanya = 0

'    Conn.BeginTrans

    vSQL = "select codempre, ariagro from empresasariagro order by codempre"
    Set Rs = New ADODB.Recordset
    Rs.Open vSQL, Cnn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Set FechasFinCamp = New Collection
    
    
    Continuar = True
    While Not Rs.EOF And Continuar
        conn.Execute "use " & Rs.Fields(1).Value
        
        Set Rs1 = New ADODB.Recordset
        vSql2 = "select  fechaini, fechafin from empresas where codempre = " & DBSet(Rs.Fields(0).Value, "N")
        Rs1.Open vSql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs1.EOF Then
            FechasFinCamp.Add DBLet(Rs1.Fields(1).Value, "F")
        
            'si not (fechainicio de nueva campaña es superior a la fechafin de campañas anteriores) --> salimos
            If Not (CDate(Text2(3).Text) > DBLet(Rs1.Fields(1).Value, "F")) Then
                Continuar = False
                Mens = Mens & "La Fecha de inicio es inferior a la campaña anterior. Base de datos: " & DBLet(Rs.Fields(1).Value, "F")
                ComprobarFechasCampanya = 1
            End If
        End If
        Set Rs1 = Nothing
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    conn.Execute "use " & vEmpresa.BDAriagro
    
    'comprobamos que la fechainicio nueva campaña debe ser:
    'la mas alta fechafin de campañas anteriores más 1 dia
    If Continuar Then
        MaxFec = FechasFinCamp(1)
        For i = 2 To FechasFinCamp.Count
            If CDate(MaxFec) < CDate(FechasFinCamp.item(i)) Then MaxFec = FechasFinCamp(i)
        Next i
    End If
    
    Set FechasFinCamp = Nothing
    
    If Continuar And MaxFec <> DateAdd("d", -1, Text2(3).Text) Then
        Continuar = False
        Mens = Mens & "La nueva campaña no comienza al dia siguiente de la anterior."
        ComprobarFechasCampanya = 2
    End If
    
'    Conn.CommitTrans
    Exit Function
    
eComprobarFechasCampanya:
    If Err.Number <> 0 Then
        MuestraError Error.Number, "Error en Comprobar Fechas Campaña:" & Err.Description
'        Conn.RollbackTrans
    End If
End Function



Private Function ComprobarUltimaCampanya(ByRef Mens As String) As Boolean
Dim Rs As ADODB.Recordset, Rs1 As ADODB.Recordset
Dim vSQL As String, vSql2 As String
Dim Continuar As Boolean
Dim FechasFinCamp As Collection
Dim MaxFec As Date
Dim i As Integer
    
    On Error GoTo eComprobarUltimaCampanya

    ComprobarUltimaCampanya = True

    vSQL = "select max(codempre) from empresasariagro"
    Set Rs = New ADODB.Recordset
    Rs.Open vSQL, Cnn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        If Rs.Fields(0).Value <> vEmpresa.codempre Then
            Mens = Mens & vbCrLf & "No se encuentra en la ultima campaña. Revise."
            ComprobarUltimaCampanya = False
        End If
    End If
        
eComprobarUltimaCampanya:
    If Err.Number <> 0 Then
        MuestraError Error.Number, "Error en Comprobar Ultima Campaña:" & Err.Description
    End If
End Function


' Insertamos los datos de toda la base de datos origen
Private Function InsercionDatosOrigen() As Boolean
Dim L As Collection
Dim Origen As String
Dim Destino As String
Dim Sql2 As String

    On Error GoTo EInsercionDatosOrigen
    
    InsercionDatosOrigen = False
    
    Set RSCta = New ADODB.Recordset
        
    '----------
    'parametros
    '----------------------------------------
    'Insertamos en tabla empresas la campaña anterior
    SQL = "INSERT INTO empresasariagro (codempre, nomempre, nomresum, Usuario, Pass, ariagro) select  "
    SQL = SQL & Text2(2).Text & ", nomempre, nomresum,'','','ariagro" & Text2(2).Text & "' from usuarios.empresasariagro where codempre = 0"
    Cnn.Execute SQL
    
    ' modificamos el registro de la campaña actual (registro 0)
    SQL = "update empresasariagro set nomempre = " & DBSet(Text2(0).Text, "T") & ", nomresum = " & DBSet(Text2(1).Text, "T")
    SQL = SQL & " where codempre = 0 "
    Cnn.Execute SQL
    
    ' inicializamos los contadores de la stipom
    If Me.Check1.Value Then
        SQL = "update stipom set contador = 0"
        
        Cnn.Execute SQL
    End If
     
     
     Pb1.visible = True
     Label6.Caption = "Insertando Datos de Base de Datos Origen."
     Me.Refresh
     
     Origen = "ariagro"
     Destino = "ariagro" & Text2(2)
     
     ' Obviamos las referenciales del destino
     conn.Execute "USE " & Destino
     conn.Execute "SET FOREIGN_KEY_CHECKS = 0"
     
     conn.Execute "USE " & Origen
     
     Set L = New Collection
     RSCta.Open "SHOW TABLES", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
     While Not RSCta.EOF

         L.Add CStr(RSCta.Fields(0))
         RSCta.MoveNext
     Wend
     Set RSCta = Nothing
     
     NumTablas = L.Count
     CargarProgres Pb1, NumTablas
     
     For ContadorInserciones = 1 To L.Count
         SQL = L.item(ContadorInserciones)
         IncrementarProgres Pb1, 1
         DoEvents
         
         Sql2 = "insert into " & Destino & "." & SQL & " select * from " & Origen & "." & SQL
         conn.Execute Sql2
         
     Next ContadorInserciones
     Set L = Nothing
     
    SQL = "UPDATE " & Trim(Origen) & ".empresas SET "
    SQL = SQL & " fechaini = " & DBSet(Text2(3).Text, "F") & ", fechafin = " & DBSet(Text2(4).Text, "F")
    conn.Execute SQL
     
    SQL = "UPDATE " & Trim(Destino) & ".empresas SET codempre = " & DBSet(Text2(2).Text, "T")
    conn.Execute SQL
     
     
    '[Monica]20/10/2011: incrementamos en un año la fecha de inicio y fin de tratamientos
    SQL = "UPDATE " & Trim(Origen) & ".advtrata SET fechaini = date_add(fechaini, interval 1 year), "
    SQL = SQL & " fechafin = date_add(fechafin, interval 1 year)"
    conn.Execute SQL
     
    ' actualizamos los campos de seguro de la campaña anterior de esta nueva bd e inicializamos los de esta campaña
    SQL = "UPDATE " & Trim(Origen) & ".rcampos SET codseguroant = codseguro, aseguradoant = asegurado, kilosaseant = kilosase, costeseguroant = costeseguro, "
    SQL = SQL & " codseguro = null, asegurado = 0, kilosase = null, costeseguro = null "
    conn.Execute SQL
     
    '[Monica]26/08/2013: desmarcamos el campo como terminado de recolectar
    SQL = "UPDATE " & Trim(Origen) & ".rcampos SET acabadorecol = 0"
    conn.Execute SQL
    
    '[Monica]11/02/2015: ponemos ficha de cultivo no entregada
    SQL = "UPDATE " & Trim(Origen) & ".rcampos SET entregafichaculti = 0"
    conn.Execute SQL
    
    
    '[Monica]26/03/2014: modificamos el año de los costes fijos, incrementando el año
    SQL = "UPDATE " & Trim(Origen) & ".ccconcostes_mes SET año = año + 1"
    conn.Execute SQL
     
     
     'Volvemos a tener control de las claves referenciales
     conn.Execute "USE " & Destino
     conn.Execute "SET FOREIGN_KEY_CHECKS = 1"
     
     Pb1.visible = False
     Me.Refresh
     
     InsercionDatosOrigen = True
     Exit Function
     
EInsercionDatosOrigen:
    MuestraError Err.Number, "Inserción Datos de DB Origen"
End Function



' Insertamos los datos de toda la base de datos origen
Private Function BorradoDatos() As Boolean
Dim L As Collection
Dim Origen As String
Dim Destino As String
Dim Sql2 As String

    On Error GoTo EBorradoDatos
    
    BorradoDatos = False
    
    Set RSCta = New ADODB.Recordset
        
     
    Pb1.visible = True
    Label6.Caption = "Borrado Datos de Históricos."
    Me.Refresh
     
    Origen = "ariagro"
     
    ' Obviamos las referenciales del destino
    conn.Execute "USE " & Origen
    conn.Execute "SET FOREIGN_KEY_CHECKS = 0"
     
    ' Tendremos que ir añadiendo las tablas que no son maestras
    NumTablas = 192
    CargarProgres Pb1, NumTablas
     
    '1 advfacturas
    Sql2 = "delete from ariagro.advfacturas"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '2 advfacturas_lineas
    Sql2 = "delete from ariagro.advfacturas_lineas"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '3 advfacturas_partes
    Sql2 = "delete from ariagro.advfacturas_partes"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '4 advfacturas_trabajador
    Sql2 = "delete from ariagro.advfacturas_trabajador"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '5 advpartes
    Sql2 = "delete from ariagro.advpartes"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '6 advpartes_lineas
    Sql2 = "delete from ariagro.advpartes_lineas"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '7 advpartes_trabajador
    Sql2 = "delete from ariagro.advpartes_trabajador"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '8 advsmoval
    Sql2 = "delete from ariagro.advsmoval"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '9 albaran
    Sql2 = "delete from ariagro.albaran"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '10 albaran_calibre
    Sql2 = "delete from ariagro.albaran_calibre"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '11 albaran_costes
    Sql2 = "delete from ariagro.albaran_costes"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '12 albaran_costreal
    Sql2 = "delete from ariagro.albaran_costreal"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '13 albaran_envase
    Sql2 = "delete from ariagro.albaran_envase"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '14 albaran_palets
    Sql2 = "delete from ariagro.albaran_palets"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '15 albaran_variedad
    Sql2 = "delete from ariagro.albaran_variedad"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '16 cabcostreal
    Sql2 = "delete from ariagro.cabcostreal"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '17 cccabdia
    Sql2 = "delete from ariagro.cccabdia"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '18 cccabecera
    Sql2 = "delete from ariagro.cccabecera"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '19 cccaborden
    Sql2 = "delete from ariagro.cccaborden"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '20 ccinforme
'[Monica]26/03/2013: no hay que borrarla pq son los puntos del informe de costes de catadau
'    Sql2 = "delete from ariagro.ccinforme"
'    conn.Execute Sql2
' Utilizo esta posicion para eliminar la tabla de ingresos de las facturas de liquidacion de terceros de picassent
    Sql2 = "delete from ariagro.ringresos"
    conn.Execute Sql2

    IncrementarProgres Pb1, 1
    DoEvents
    '21 cclindia1
    Sql2 = "delete from ariagro.cclindia1"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '22 cclindia2
    Sql2 = "delete from ariagro.cclindia2"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '23 cclineas1
    Sql2 = "delete from ariagro.cclineas1"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '24 cclineas2
    Sql2 = "delete from ariagro.cclineas2"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '25 cclineas3
    Sql2 = "delete from ariagro.cclineas3"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '26 cclineas4
    Sql2 = "delete from ariagro.cclineas4"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '27 cclineas5
    Sql2 = "delete from ariagro.cclineas5"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '28 cclineas6
    Sql2 = "delete from ariagro.cclineas6"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '29 cclinorden1
    Sql2 = "delete from ariagro.cclinorden1"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '30 cclinorden2
    Sql2 = "delete from ariagro.cclinorden2"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '31 cclinorden3
    Sql2 = "delete from ariagro.cclinorden3"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '32 cclinorden4
    Sql2 = "delete from ariagro.cclinorden4"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '33 cclinorden5
    Sql2 = "delete from ariagro.cclinorden5"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '34 ccticajes
    Sql2 = "delete from ariagro.ccticajes"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '35 cctrabaconf
    Sql2 = "delete from ariagro.cctrabaconf"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '36 chivato
    Sql2 = "delete from ariagro.chivato"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '37 facturas
    Sql2 = "delete from ariagro.facturas"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '38 facturas_acuenta
    Sql2 = "delete from ariagro.facturas_acuenta"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '39 facturas_calibre
    Sql2 = "delete from ariagro.facturas_calibre"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '40 facturas_envases
    Sql2 = "delete from ariagro.facturas_envases"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '41 facturas_variedad
    Sql2 = "delete from ariagro.facturas_variedad"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '42 facturassocio
    Sql2 = "delete from ariagro.facturassocio"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '43 facturassocio_envases
    Sql2 = "delete from ariagro.facturassocio_envases"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '44 facturassocio_variedad
    Sql2 = "delete from ariagro.facturassocio_variedad"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '45 fvarcabfact
    Sql2 = "delete from ariagro.fvarcabfact"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '46 fvarcabfactpro
    Sql2 = "delete from ariagro.fvarcabfactpro"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '47 fvarlinfact
    Sql2 = "delete from ariagro.fvarlinfact"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '48 fvarlinfactpro
    Sql2 = "delete from ariagro.fvarlinfactpro"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '49 horas
    Sql2 = "delete from ariagro.horas"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '50 horasdestajo
    Sql2 = "delete from ariagro.horasdestajo"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '51 horasett
    Sql2 = "delete from ariagro.horasett"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '52 lincostreal
    Sql2 = "delete from ariagro.lincostreal"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '53 palets
    Sql2 = "delete from ariagro.palets"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '54 palets_calibre
    Sql2 = "delete from ariagro.palets_calibre"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '55 palets_variedad
    Sql2 = "delete from ariagro.palets_variedad"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '56 pedidos
    Sql2 = "delete from ariagro.pedidos"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '57 pedidos_calibre
    Sql2 = "delete from ariagro.pedidos_calibre"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '58 pedidos_variedad
    Sql2 = "delete from ariagro.pedidos_variedad"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '59 rbodalbaran
    Sql2 = "delete from ariagro.rbodalbaran"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '60 rbodalbaran_variedad
    Sql2 = "delete from ariagro.rbodalbaran_variedad"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '61 rbodfacturas
    Sql2 = "delete from ariagro.rbodfacturas"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '62 rbodfacturas_alb
    Sql2 = "delete from ariagro.rbodfacturas_alb"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '63 rbodfacturas_lineas
    Sql2 = "delete from ariagro.rbodfacturas_lineas"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '64 rbonifentradas
    Sql2 = "delete from ariagro.rbonifentradas"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '65 rcabfactalmz
    Sql2 = "delete from ariagro.rcabfactalmz"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '66 rcafter
    Sql2 = "delete from ariagro.rcafter"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '67 rclasifauto
    Sql2 = "delete from ariagro.rclasifauto"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '68 rclasifauto_clasif
    Sql2 = "delete from ariagro.rclasifauto_clasif"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '69 rclasifauto_plagas
    Sql2 = "delete from ariagro.rclasifauto_plagas"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '70 rclasifica
    Sql2 = "delete from ariagro.rclasifica"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '71 rclasifica_clasif
    Sql2 = "delete from ariagro.rclasifica_clasif"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '72 rclasifica_imp
    Sql2 = "delete from ariagro.rclasifica_imp"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '73 rclasifica_incidencia
    Sql2 = "delete from ariagro.rclasifica_incidencia"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '74 rcontrol
    Sql2 = "delete from ariagro.rcontrol"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '75 rcontrol_plagas
    Sql2 = "delete from ariagro.rcontrol_plagas"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '76 rentradas
    Sql2 = "delete from ariagro.rentradas"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '77 rfactsoc
    Sql2 = "delete from ariagro.rfactsoc"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '78 rfactsoc_albaran
    Sql2 = "delete from ariagro.rfactsoc_albaran"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '79 rfactsoc_anticipos
    Sql2 = "delete from ariagro.rfactsoc_anticipos"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '80 rfactsoc_calidad
    Sql2 = "delete from ariagro.rfactsoc_calidad"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '81 rfactsoc_fvarias
    Sql2 = "delete from ariagro.rfactsoc_fvarias"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '82 rfactsoc_gastos
    Sql2 = "delete from ariagro.rfactsoc_gastos"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '83 rfactsoc_retirada
    Sql2 = "delete from ariagro.rfactsoc_retirada"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '84 rfactsoc_variedad
    Sql2 = "delete from ariagro.rfactsoc_variedad"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '85 rfacttra
    Sql2 = "delete from ariagro.rfacttra"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '86 rfacttra_albaran
    Sql2 = "delete from ariagro.rfacttra_albaran"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '87 rhisfruta
    Sql2 = "delete from ariagro.rhisfruta"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '88 rhisfruta_clasif
    Sql2 = "delete from ariagro.rhisfruta_clasif"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '89 rhisfruta_entradas
    Sql2 = "delete from ariagro.rhisfruta_entradas"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '90 rhisfruta_gastos
    Sql2 = "delete from ariagro.rhisfruta_gastos"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '91 rhisfruta_incidencia
    Sql2 = "delete from ariagro.rhisfruta_incidencia"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '92 rhistrans
    Sql2 = "delete from ariagro.rhistrans"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    
    '93 rliantifter
    Sql2 = "delete from ariagro.rliantifter"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    
    '94 rlifter
    Sql2 = "delete from ariagro.rlifter"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '95 rlinfactalmz
    Sql2 = "delete from ariagro.rlinfactalmz"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '96 rpartes
    Sql2 = "delete from ariagro.rpartes"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '97 rpartes_gastos
    Sql2 = "delete from ariagro.rpartes_gastos"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '98 rpartes_trabajador
    Sql2 = "delete from ariagro.rpartes_trabajador"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '99 rpartes_variedad
    Sql2 = "delete from ariagro.rpartes_variedad"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '100 rpesadas
    Sql2 = "delete from ariagro.rpesadas"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '101 rprecios
    Sql2 = "delete from ariagro.rprecios"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '102 rprecios_calidad
    Sql2 = "delete from ariagro.rprecios_calidad"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '103 rrecasesoria
    Sql2 = "delete from ariagro.rrecasesoria"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '104 rrecibosnomina
    Sql2 = "delete from ariagro.rrecibosnomina"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '105 rrecibpozos
    Sql2 = "delete from ariagro.rrecibpozos"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '106 rrecibpozos_acc
    Sql2 = "delete from ariagro.rrecibpozos_acc"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '107 rrecibpozos_cam
    Sql2 = "delete from ariagro.rrecibpozos_cam"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '108 rrecibpozos_hid
    Sql2 = "delete from ariagro.rrecibpozos_hid"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '109 rrendim
    Sql2 = "delete from ariagro.rrendim"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '110 rsmsenviados
    Sql2 = "delete from ariagro.rsmsenviados"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '111 rtelmovil
    Sql2 = "delete from ariagro.rtelmovil"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '112 scaalb
    Sql2 = "delete from ariagro.scaalb"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '113 scaalp
    Sql2 = "delete from ariagro.scaalp"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '114 scafpa
    Sql2 = "delete from ariagro.scafpa"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '115 scafpc
    Sql2 = "delete from ariagro.scafpc"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '116 scamov
    Sql2 = "delete from ariagro.scamov"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '117 scappr
    Sql2 = "delete from ariagro.scappr"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '118 scaser
    Sql2 = "delete from ariagro.scaser"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '119 scatra
    Sql2 = "delete from ariagro.scatra"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '120 schalp
    Sql2 = "delete from ariagro.schalp"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '121 schmov
    Sql2 = "delete from ariagro.schmov"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '122 schppr
    Sql2 = "delete from ariagro.schppr"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '123 schser
    Sql2 = "delete from ariagro.schser"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '124 schtra
    Sql2 = "delete from ariagro.schtra"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
'    '125 scrmacciones
'    Sql2 = "delete from ariagro.scrmacciones"
'    conn.Execute Sql2
'    IncrementarProgres Pb1, 1
'    DoEvents

    '125 shinve
    Sql2 = "delete from ariagro.shinve"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '126 sinven
    Sql2 = "delete from ariagro.sinven"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '127 slhalp
    Sql2 = "delete from ariagro.slhalp"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '128 slhmov
    Sql2 = "delete from ariagro.slhmov"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '129 slhppr
    Sql2 = "delete from ariagro.slhppr"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '130 slhser
    Sql2 = "delete from ariagro.slhser"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '131 slhtra
    Sql2 = "delete from ariagro.slhtra"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '132 slialb
    Sql2 = "delete from ariagro.slialb"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '133 slialp
    Sql2 = "delete from ariagro.slialp"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '134 slifpc
    Sql2 = "delete from ariagro.slifpc"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '135 slimov
    Sql2 = "delete from ariagro.slimov"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '136 slippr
    Sql2 = "delete from ariagro.slippr"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '137 sliser
    Sql2 = "delete from ariagro.sliser"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '138 slisp1
    Sql2 = "delete from ariagro.slisp1"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '139 slispr
    Sql2 = "delete from ariagro.slispr"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '140 slitra
    Sql2 = "delete from ariagro.slitra"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '141 slog
    Sql2 = "delete from ariagro.slog"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '142 smoval
    Sql2 = "delete from ariagro.smoval"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '143 tcafpa
    Sql2 = "delete from ariagro.tcafpa"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '144 tcafpc
    Sql2 = "delete from ariagro.tcafpc"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '145 tcafpv
    Sql2 = "delete from ariagro.tcafpv"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '146 tlifpc
    Sql2 = "delete from ariagro.tlifpc"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '147 tmp346
    Sql2 = "delete from ariagro.tmp346"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '148 tmpalmazara
    Sql2 = "delete from ariagro.tmpalmazara"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '149 tmpclasifica
    Sql2 = "delete from ariagro.tmpclasifica"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '150 tmpclasifica2
    Sql2 = "delete from ariagro.tmpclasifica2"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '151 tmpcmr
    Sql2 = "delete from ariagro.tmpcmr"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '152 tmpcopiascmr
'[Monica]20/09/2013: no se borra son las hojas del CMR.
'    Sql2 = "delete from ariagro.tmpcopiascmr"
'    Conn.Execute Sql2
'[Monica]23/09/2013: utilizo esta posicion para borrar los anticipos de nomina
    Sql2 = "delete from ariagro.horasanticipos"
    conn.Execute Sql2
    
    IncrementarProgres Pb1, 1
    DoEvents
    '153 tmpenvasesret
    Sql2 = "delete from ariagro.tmpenvasesret"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '154 tmpexcel
    Sql2 = "delete from ariagro.tmpexcel"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '155 tmpfactura
    Sql2 = "delete from ariagro.tmpfactura"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '156 tmpinfcostes
    Sql2 = "delete from ariagro.tmpinfcostes"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '157 tmpinfkilos
    Sql2 = "delete from ariagro.tmpinfkilos"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '158 tmpinformes
    Sql2 = "delete from ariagro.tmpinformes"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '159 tmpinfventas
    Sql2 = "delete from ariagro.tmpinfventas"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '160 tmplineas
    Sql2 = "delete from ariagro.tmplineas"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '161 tmpliquidacion
    Sql2 = "delete from ariagro.tmpliquidacion"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '162 tmpliquidacion1
    Sql2 = "delete from ariagro.tmpliquidacion1"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '163 tmppesada
    Sql2 = "delete from ariagro.tmppesada"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '164 tmpportesv
    Sql2 = "delete from ariagro.tmpportesv"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '165 tmppreciosaux
    Sql2 = "delete from ariagro.tmppreciosaux"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '166 tmprfactsoc
    Sql2 = "delete from ariagro.tmprfactsoc"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '167 tmprfactsoc_variedad
    Sql2 = "delete from ariagro.tmprfactsoc_variedad"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '168 tmpscapla
    Sql2 = "delete from ariagro.tmpscapla"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '169 tmpstockfec
    Sql2 = "delete from ariagro.tmpstockfec"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '170 tmpsuperficies
    Sql2 = "delete from ariagro.tmpsuperficies"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '171 trzentradas
    Sql2 = "delete from ariagro.trzentradas"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '172 trzlineas_cargas
    Sql2 = "delete from ariagro.trzlineas_cargas"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '173 trzlineas_confeccion_entradas
    Sql2 = "delete from ariagro.trzlineas_confeccion_entradas"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '174 trzpalet_palets
    Sql2 = "delete from ariagro.trzpalet_palets"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '175 trzpalets
    Sql2 = "delete from ariagro.trzpalets"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '176 trztmp_palets_lineas_cargas
    Sql2 = "delete from ariagro.trztmp_palets_lineas_cargas"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '177 trztransitos
    Sql2 = "delete from ariagro.trztransitos"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '178 vtafrutacab
    Sql2 = "delete from ariagro.vtafrutacab"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '179 vtafrutalin
    Sql2 = "delete from ariagro.vtafrutalin"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '180 rhisfrutasin
    Sql2 = "delete from ariagro.rhisfrutasin"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '181 rhisfrutasin_clasif
    Sql2 = "delete from ariagro.rhisfrutasin_clasif"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '182 rhisfrutasin_entradas
    Sql2 = "delete from ariagro.rhisfrutasin_entradas"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '183 rhisfrutasin_gastos
    Sql2 = "delete from ariagro.rhisfrutasin_gastos"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '184 rhisfrutasin_incidencia
    Sql2 = "delete from ariagro.rhisfrutasin_incidencia"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    
'[Monica]02/10/2013: tabla ya no exite, sustituida por las dos siguientes
'    '[Monica]26/08/2013: nueva tabla de ordenes de recoleccion
'    '185 rcampos_ordrec
'    Sql2 = "delete from ariagro.rcampos_ordrec"
'    Conn.Execute Sql2
'    IncrementarProgres Pb1, 1
'    DoEvents
    '185 rordrecogida
    Sql2 = "delete from ariagro.rordrecogida"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    '186 rordrecogida_incid
    Sql2 = "delete from ariagro.rordrecogida_incid"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
     
     
    '187 anecoop
    Sql2 = "delete from ariagro.anecoop"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
     
    '188 anecoop_cobro
    Sql2 = "delete from ariagro.anecoop_cobro"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    
    '189 anecoop_pago
    Sql2 = "delete from ariagro.anecoop_pago"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
    
    
    '[Monica]29/01/2016: solo para el caso de catadau borramos los datos de asegurado
    '190 rcampos_seguros
    If vParamAplic.Cooperativa = 0 Then
        Sql2 = "delete from ariagro.rcampos_seguros"
        conn.Execute Sql2
        IncrementarProgres Pb1, 1
        DoEvents
    Else
        IncrementarProgres Pb1, 1
        DoEvents
    End If
     
    '[Monica]21/04/2017: faltan los ingresos en las facturas de socios
    '191 rfactsoc_ingresos
    Sql2 = "delete from ariagro.rfactsoc_ingresos"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
     
    '[Monica]21/04/2017: falta trzmovim
    '192 trzmovim
    Sql2 = "delete from ariagro.trzmovim"
    conn.Execute Sql2
    IncrementarProgres Pb1, 1
    DoEvents
     
     
     
     
    'Volvemos a tener control de las claves referenciales
    conn.Execute "SET FOREIGN_KEY_CHECKS = 1"
     
    Pb1.visible = False
    Me.Refresh
     
    BorradoDatos = True
    Exit Function
     
EBorradoDatos:
    MuestraError Err.Number, "Borrado Datos de tablas no maestras."
End Function








