VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPagoRecibos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6570
   Icon            =   "frmPagoRecibos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameHorasTrabajadas 
      Height          =   5130
      Left            =   45
      TabIndex        =   6
      Top             =   0
      Width           =   6435
      Begin VB.CheckBox Check1 
         Caption         =   "Sobre Horas Productivas"
         Height          =   195
         Index           =   1
         Left            =   540
         TabIndex        =   16
         Top             =   3870
         Width           =   2130
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Tag             =   "Tipo|N|N|||straba|codsecci||N|"
         Top             =   3420
         Width           =   1665
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   18
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   2790
         Width           =   3015
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   1830
         MaxLength       =   6
         TabIndex        =   3
         Top             =   2790
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1800
         Width           =   1005
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   20
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   2
         Top             =   2340
         Width           =   1005
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   1845
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Tag             =   "Tipo|N|N|||straba|codsecci||N|"
         Top             =   1035
         Width           =   1350
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4650
         TabIndex        =   5
         Top             =   4500
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   3570
         TabIndex        =   4
         Top             =   4485
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   240
         Left            =   480
         TabIndex        =   11
         Top             =   4170
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   90
         Top             =   3915
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Caption         =   "Concepto Transferencia "
         ForeColor       =   &H00972E0B&
         Height          =   255
         Left            =   540
         TabIndex        =   15
         Top             =   3150
         Width           =   1875
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   14
         Left            =   1515
         MouseIcon       =   "frmPagoRecibos.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar banco"
         Top             =   2790
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Banco "
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   27
         Left            =   540
         TabIndex        =   13
         Top             =   2700
         Width           =   510
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   1530
         Picture         =   "frmPagoRecibos.frx":015E
         ToolTipText     =   "Buscar fecha"
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Recibo"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   24
         Left            =   540
         TabIndex        =   10
         Top             =   1530
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Pago"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   30
         Left            =   540
         TabIndex        =   9
         Top             =   2160
         Width           =   870
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   6
         Left            =   1515
         Picture         =   "frmPagoRecibos.frx":01E9
         ToolTipText     =   "Buscar fecha"
         Top             =   2340
         Width           =   240
      End
      Begin VB.Label Label8 
         Caption         =   "Sección "
         ForeColor       =   &H00972E0B&
         Height          =   255
         Left            =   540
         TabIndex        =   8
         Top             =   1035
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Pago de Recibos"
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
         TabIndex        =   7
         Top             =   405
         Width           =   5925
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
Attribute VB_Name = "frmPagoRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: MONICA +-+-
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
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public Event RectificarFactura(cliente As String, observaciones As String)

Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean


Private WithEvents frmBan As frmManBanco 'Banco propio
Attribute frmBan.VB_VarHelpID = -1

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
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim tipo As String
Dim Repetir As Boolean

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
    
Dim cadSelect1 As String
Dim cadSelect2 As String
Dim cTabla As String
Dim sql As String

    
    If Not DatosOk Then Exit Sub
    
    cadselect = ""
    'Fecha de Recibo
    AnyadirAFormula cadselect, "horas.fecharec = " & DBSet(txtCodigo(16).Text, "F")
              
    'Tipo de seccion
    AnyadirAFormula cadselect, "straba.codsecci = " & Me.Combo1(1).ListIndex
    
    'La forma de pago tiene que ser de tipo Transferencia
    AnyadirAFormula cadselect, "forpago.tipoforp = 1"
    
    tabla = "(horas INNER JOIN straba ON horas.codtraba = straba.codtraba) INNER JOIN forpago ON straba.codforpa = forpago.codforpa "
               
    cTabla = tabla
    cadSelect1 = cadselect
    cadSelect2 = cadselect
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    sql = "Select count(*) FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cadSelect1 <> "" Then
        cadSelect1 = QuitarCaracterACadena(cadSelect1, "{")
        cadSelect1 = QuitarCaracterACadena(cadSelect1, "}")
        cadSelect1 = QuitarCaracterACadena(cadSelect1, "_1")
        sql = sql & " WHERE " & cadSelect1
    End If
    
    If RegistrosAListar(sql) = 0 Then
        MsgBox "No hay datos para mostrar en el Informe.", vbInformation
    Else
        AnyadirAFormula cadselect, "horas.intconta = 0"

        'Comprobar si hay registros a Mostrar antes de abrir el Informe
        If HayRegParaInforme(tabla, cadselect) Then
            ProcesarCambios (cadselect)
        Else
            Repetir = True
            If MsgBox("¿Desea repetir el proceso?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                ProcesarCambios (cadSelect2)
            End If
        End If
    End If
    
    cmdCancel_Click
    
End Sub

Private Sub ProcesarCambios(cadWhere As String)
Dim sql As String
Dim SQL2 As String
Dim Sql3 As String
Dim cad As String
Dim I As Integer
Dim HayReg As Integer
Dim b As Boolean
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim mens As String

Dim ImpHoras As Currency
Dim ImpHorasE As Currency
Dim ImpBruto As Currency
Dim IRPF As Currency
Dim SegSoc As Currency
Dim Neto As Currency
Dim Bruto As Currency
Dim CuentaPropia As String
Dim CodigoOrden34 As String

On Error GoTo eProcesarCambios
    
    BorrarTMP
    CrearTMP

    Conn.BeginTrans
    
    
    If cadWhere <> "" Then
        cadWhere = QuitarCaracterACadena(cadWhere, "{")
        cadWhere = QuitarCaracterACadena(cadWhere, "}")
        cadWhere = QuitarCaracterACadena(cadWhere, "_1")
    End If
        
    sql = "select count(distinct horas.codtraba) from (horas inner join straba on horas.codtraba = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where " & cadWhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Pb1.visible = True
    CargarProgres Pb1, Rs.Fields(0).Value
    
    Rs.Close
    
    If Check1(1).Value = 0 Then
        sql = "select horas.codtraba, sum(horasdia), sum(compleme), sum(horasext) from (horas inner join straba on horas.codtraba = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where " & cadWhere
    Else
        sql = "select horas.codtraba, sum(horasproduc), sum(compleme), sum(horasext) from (horas inner join straba on horas.codtraba = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where " & cadWhere
    End If
    sql = sql & " group by horas.codtraba "
    
'    BorrarTMP
'    CrearTMP
    
    Rs.Open sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        IncrementarProgres Pb1, 1
        mens = "Calculando Importes" & vbCrLf & vbCrLf & "Trabajador: " & Rs!codtraba & vbCrLf
        
        SQL2 = "select salarios.* from salarios, straba where straba.codtraba = " & DBSet(Rs!codtraba, "N")
        SQL2 = SQL2 & " and salarios.codcateg = straba.codcateg "
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open SQL2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        ImpHoras = Round2(DBLet(Rs.Fields(1).Value, "N") * DBLet(Rs2!impsalar, "N"), 2)
        ImpHorasE = Round2(DBLet(Rs.Fields(3).Value, "N") * DBLet(Rs2!imphorae, "N"), 2)
        ImpBruto = Round2(ImpHoras + ImpHorasE + DBLet(Rs.Fields(2).Value, "N"), 2)
        
        IRPF = Round2(ImpBruto * DBLet(Rs2!dtosirpf, "N") / 100, 2)
        SegSoc = Round2(ImpBruto * DBLet(Rs2!dtosegso, "N") / 100, 2)
        
        Neto = Round2(ImpBruto - IRPF - SegSoc, 2)
        
        Sql3 = "insert into tmpImpor (codtraba, importe) values ("
        Sql3 = Sql3 & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(ImporteSinFormato(CStr(Neto)), "N") & ")"
        
        Conn.Execute Sql3
        
        Set Rs2 = Nothing
            
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    sql = "select codbanco, codsucur, digcontr, cuentaba, codorden34 from banpropi where codbanpr = " & DBSet(txtCodigo(18).Text, "N")
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CodigoOrden34 = ""
    
    If Rs.EOF Then
        cad = ""
    Else
        If IsNull(Rs!codbanco) Then
            cad = ""
        Else
            cad = Format(Rs!codbanco, "0000") & "|" & Format(DBLet(Rs!codsucur, "T"), "0000") & "|" & DBLet(Rs!digcontr, "T") & "|" & Format(DBLet(Rs!CuentaBa, "T"), "0000000000") & "|"
        End If
        CodigoOrden34 = DBLet(Rs!codorden34, "T")
    End If
    
    Set Rs = Nothing
    
    CuentaPropia = cad
    b = GeneraFicheroNorma34New(vParam.CifEmpresa, CDate(txtCodigo(20).Text), CuentaPropia, 9, 0, "Pago Nómina", CodigoOrden34, Combo1(0).ListIndex)
    If b Then
        If CopiarFichero Then
            If Not Repetir Then
                If MsgBox("¿Proceso realizado correctamente para actualizar?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    sql = "update horas, straba, forpago set horas.intconta = 1 where horas.codtraba = straba.codtraba and straba.codforpa = forpago.codforpa and " & cadWhere
                    Conn.Execute sql
                End If
            End If
        End If
    End If

eProcesarCambios:
    If Err.Number <> 0 Then
        mens = Err.Description
        b = False
    End If
    If b Then
        Conn.CommitTrans
        MsgBox "Proceso realizado correctamente.", vbExclamation
        cmdCancel_Click
    Else
        Conn.RollbackTrans
        MsgBox "Error " & mens, vbExclamation
    End If
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
    For H = 14 To 14 'imgBuscar.Count - 1
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
     
    
    Set List = Nothing

    'Ocultar todos los Frames de Formulario
    Me.FrameHorasTrabajadas.visible = False
    
    '###Descomentar
'    CommitConexion
    H = 5055
    W = 6660
    FrameHorasTrabajadasVisible True, H, W
    indFrame = 0
    tabla = "horas"
        
    CargaCombo
    Combo1(0).ListIndex = 0
        
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
    
    Me.Combo1(1).ListIndex = 1
    
    Pb1.visible = False
End Sub

Private Sub frmBan_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de banco propio
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtCodigo(CByte(imgFecha(2).Tag) + 14).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 14  'Banco propio
            AbrirFrmManBanco (Index)
    
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
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

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
            Case 16: KEYFecha KeyAscii, 2  'fecha recibo
            Case 20: KEYFecha KeyAscii, 6 'fecha pago
            Case 18: KEYBusqueda KeyAscii, 14 'banco
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
            
        Case 16, 20   'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 18 ' banco propio
            If txtCodigo(Index).Text <> "" Then PonerFormatoEntero txtCodigo(Index)
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "banpropi", "nombanpr", "codbanpr", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            
    End Select
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim tabla As String
Dim Titulo As String

    'Llamamos a al form
    cad = ""
    Conexion = cAgro    'Conexión a BD: Ariges
'    Select Case OpcionListado
'        Case 7 'Traspaso de Almacenes
'            cad = cad & "Nº Trasp|scatra|codtrasp|N|0000000|40·Almacen Origen|scatra|almaorig|N|000|20·Almacen Destino|scatra|almadest|N|000|20·Fecha|scatra|fechatra|F||20·"
'            Tabla = "scatra"
'            titulo = "Traspaso Almacenes"
'        Case 8 'Movimientos de Almacen
'            cad = cad & "Nº Movim.|scamov|codmovim|N|0000000|40·Almacen|scamov|codalmac|N|000|30·Fecha|scamov|fecmovim|F||30·"
'            Tabla = "scamov"
'            titulo = "Movimientos Almacen"
'        Case 9, 12, 13, 14, 15, 16, 17 '9: Movimientos Articulos
'                   '12: Inventario Articulos
'                   '14:Actualizar Diferencias de Stock Inventariado
'                   '16: Listado Valoracion stock inventariado
'            cad = cad & "Código|sartic|codartic|T||30·Denominacion|sartic|nomartic|T||70·"
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
        Me.FrameHorasTrabajadas.Height = H
        Me.FrameHorasTrabajadas.Width = W
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
        .EnvioEMail = False
'        .NombreRPT = cadNombreRPT
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
'            cadParam = cadParam & "pTitulo1" & "=""Código""" & "|"
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
            tipo = "Código"
        Case "Alfabético"
            cadParam = cadParam & "Orden" & "= {" & tabla
            Select Case OpcionListado
                Case 10
                    cadParam = cadParam & ".nomclien}|"
                Case 11
                    cadParam = cadParam & ".nomprove}|"
            End Select
            tipo = "Alfabético"
    End Select
    
    numParam = numParam + 1

End Function

Private Sub AbrirFrmManBanco(indice As Integer)
    indCodigo = indice + 4
    Set frmBan = New frmManBanco
    frmBan.DatosADevolverBusqueda = "0|1|"
    frmBan.Show vbModal
    Set frmBan = Nothing
End Sub

Private Sub AbrirVisReport()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .Opcion = OpcionListado
        .Show vbModal
    End With
    
    Unload Me
End Sub

Private Sub AbrirEMail()
    If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
End Sub


' ********* si n'hi han combos a la capçalera ************
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
    Combo1(1).AddItem "Almacén"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    
    Combo1(0).Clear
    
    Combo1(0).AddItem "Nómina"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Pensión"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "Otros Conceptos"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    
    
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim sql As String
'Dim Datos As String

    On Error GoTo EDatosOK

    b = True

    If txtCodigo(16).Text = "" Then
        MsgBox "Debe introducir una Fecha de Recibo.", vbExclamation
        txtCodigo(16).Text = ""
        PonerFoco txtCodigo(16)
        b = False
    End If
    
    If b And txtCodigo(20).Text = "" Then
        MsgBox "Debe introducir una Fecha de Pago.", vbExclamation
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
    
    Conn.Execute sql
        
    ActualizarRegistros = True
    
    Exit Function

eActualizarRegistros:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Error en la actualizacion de Registros" & vbCrLf & Err.Description
    End If
End Function

Public Sub BorrarTMP()
On Error Resume Next

    Conn.Execute " DROP TABLE IF EXISTS tmpImpor;"
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Function CrearTMP() As Boolean
'Crea una temporal donde inserta la clave primaria de las
'facturas seleccionadas para facturar y trabaja siempre con ellas
Dim sql As String
    
    On Error GoTo ECrear
    
    CrearTMP = False
    
    sql = "CREATE TEMPORARY TABLE tmpImpor ( "
    sql = sql & "codtraba int(6) unsigned NOT NULL default '0',"
    sql = sql & "importe decimal(12,2)  NOT NULL default '0')"
    
    Conn.Execute sql
     
    CrearTMP = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMP = False
        'Borrar la tabla temporal
        sql = " DROP TABLE IF EXISTS tmpImpor;"
        Conn.Execute sql
    End If
End Function

Public Function CopiarFichero() As Boolean
Dim nomFich As String

On Error GoTo ecopiarfichero

    CopiarFichero = False
    ' abrimos el commondialog para indicar donde guardarlo
'    Me.CommonDialog1.InitDir = App.path

    Me.CommonDialog1.DefaultExt = "txt"
    
    CommonDialog1.Filter = "Archivos txt|txt|"
    CommonDialog1.FilterIndex = 1
    
    ' copiamos el primer fichero
    CommonDialog1.FileName = "norma34.txt"
    Me.CommonDialog1.ShowSave
    
    If CommonDialog1.FileName <> "" Then
        FileCopy App.path & "\norma34.txt", CommonDialog1.FileName
    End If
    
    CopiarFichero = True
    Exit Function

ecopiarfichero:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    End If
    Err.Clear
End Function
