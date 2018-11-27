VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmVtasOrdenConfec 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Orden de Confección"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   8160
   Icon            =   "frmVtasOrdenConfec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameSeleccion 
      Height          =   5505
      Left            =   30
      TabIndex        =   6
      Top             =   30
      Width           =   8055
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
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
         Left            =   5895
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   4980
         Width           =   1620
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
         Index           =   1
         Left            =   1500
         TabIndex        =   9
         Top             =   4980
         Width           =   1065
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "Aceptar"
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
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   4980
         Width           =   1065
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4425
         Left            =   150
         TabIndex        =   7
         Top             =   420
         Width           =   7725
         _ExtentX        =   13626
         _ExtentY        =   7805
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   390
         Picture         =   "frmVtasOrdenConfec.frx":000C
         ToolTipText     =   "Desmarcar todos"
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   150
         Picture         =   "frmVtasOrdenConfec.frx":0A0E
         ToolTipText     =   "Marcar todos"
         Top             =   120
         Width           =   240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         X1              =   5805
         X2              =   7560
         Y1              =   4890
         Y2              =   4890
      End
      Begin VB.Label Label1 
         Caption         =   "Pedidos Seleccionados"
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
         Index           =   1
         Left            =   3240
         TabIndex        =   11
         Top             =   5040
         Width           =   2685
      End
   End
   Begin VB.Frame FrameCobros 
      Height          =   5460
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   6375
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
         Index           =   6
         Left            =   3015
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1755
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
         Index           =   0
         Left            =   4890
         TabIndex        =   3
         Top             =   4485
         Width           =   1065
      End
      Begin VB.CommandButton CmdAceptar 
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
         Index           =   0
         Left            =   3735
         TabIndex        =   1
         Top             =   4485
         Width           =   1065
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   3825
         Width           =   5610
         _ExtentX        =   9895
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Carga"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   435
         Index           =   16
         Left            =   630
         TabIndex        =   4
         Top             =   1665
         Width           =   1905
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   6
         Left            =   2655
         Picture         =   "frmVtasOrdenConfec.frx":7260
         ToolTipText     =   "Buscar fecha"
         Top             =   1755
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmVtasOrdenConfec"
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
Dim NRegSelec As Integer
Dim CtaNuevoCliente As String



Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cmdAceptar_Click(Index As Integer)
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim cadMen As String
Dim i As Byte
Dim Sql As String
Dim Tipo As Byte
Dim Nregs As Integer
Dim NumError As Long
Dim db As BaseDatos

Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
Dim cadselect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal


    Select Case Index
        Case 0
            If Not DatosOk Then Exit Sub
            
            If CargarListView = 0 Then
                VisualizarListview False
                MsgBox "No existen datos entre esos límites.", vbExclamation
            Else
                VisualizarListview True
            End If
        Case 1
            If NRegSelec = 0 Then
                MsgBox "No ha seleccionado ningún pedido para realizar la Orden de Confección.", vbExclamation
                Exit Sub
            Else
                If ProcesarCambios Then
                       cadFormula = ""
                       cadParam = ""
                       cadselect = ""
                       numParam = 0
                         
                       indRPT = 6 'Impresion de Orden de Confección
                       If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                       
                       'Nombre fichero .rpt a Imprimir
                       frmImprimir.NombreRPT = nomDocu
                       
                       cadParam = cadParam & "pUsu=" & vUsu.Codigo & "|"
                       numParam = numParam + 1
                       
                       cadParam = cadParam & "pDHFecha= ""Fecha de Carga: " & txtCodigo(6).Text & """|"
                       numParam = numParam + 1
                       
                       AnyadirAFormula cadFormula, "{tmpinformes.codusu} =" & vUsu.Codigo
                       
                       With frmImprimir
                             .FormulaSeleccion = cadFormula
                             .OtrosParametros = cadParam
                             .NumeroParametros = numParam
                             .SoloImprimir = False
                             .EnvioEMail = False
                             .ConSubInforme = True
                             .Opcion = 0
                             .Titulo = "Impresión de Orden de Confección"
                             .Show vbModal
                       End With
                       If frmVisReport.EstaImpreso = True Then
                            ActualizarPedidos
                       End If

                       cmdCancel_Click (0)
                End If
            End If
    End Select

End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        ValoresPorDefecto
        PonerFoco txtCodigo(6)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection
Dim i As Integer

    PrimeraVez = True
    limpiar Me

    VisualizarListview False
    
    NRegSelec = 0
    
    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, H, W
    indFrame = 5
    
    Pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(0).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub


Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(6).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub Image1_Click(Index As Integer)
Dim i As Integer
Dim TotalCant As Currency
Dim TotalImporte As Currency

    Screen.MousePointer = vbHourglass
    
    TotalCant = 0
    TotalImporte = 0
    NRegSelec = 0
    
    Select Case Index
        Case 0
            For i = 1 To ListView1.ListItems.Count
                ListView1.ListItems(i).Checked = True
                NRegSelec = NRegSelec + 1
            Next i
        Case 1
            For i = 1 To ListView1.ListItems.Count
                ListView1.ListItems(i).Checked = False
            Next i

    End Select
    Screen.MousePointer = vbDefault

    Text1(0).Text = NRegSelec

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
    imgFec(6).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtCodigo(Index).Text <> "" Then frmC.NovaData = txtCodigo(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    If imgFec(6).Tag = 6 Then
        PonerFoco txtCodigo(3)
    Else
        PonerFoco txtCodigo(1)
    End If
    ' ***************************
End Sub

Private Sub ListView1_ItemCheck(ByVal item As MSComctlLib.ListItem)
Dim i As Integer
Dim TotalCant As Currency
Dim TotalImporte As Currency
    
    Screen.MousePointer = vbHourglass
    
    NRegSelec = 0
    
    ' vemos si lo podemos seleccionar
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then
            NRegSelec = NRegSelec + 1
        End If
    Next i
    
    Screen.MousePointer = vbDefault

    Text1(0).Text = NRegSelec
    
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'14/02/2007
'    KEYpress KeyAscii
'ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 6: KEYFecha KeyAscii, 6 'fecha desde
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
            
        Case 6 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
        
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 6015
        Me.FrameCobros.Width = 6555
        W = Me.FrameCobros.Width
        H = Me.FrameCobros.Height
    End If
End Sub

Private Sub ValoresPorDefecto()
    txtCodigo(6).Text = Format(Now, "dd/mm/yyyy")
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Sql As String

    b = True
    
    If txtCodigo(6).Text = "" Then
        MsgBox "Debe introducir obligatoriamente una Fecha de Carga.", vbExclamation
        b = False
    End If

    DatosOk = b
    
End Function

Private Function CargarListView() As Integer
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String
Dim HayReg As Integer

    On Error GoTo ECargarList
    
    HayReg = 0
    CargarListView = 0
    If ListView1.ListItems.Count <> 0 Then Exit Function

    Screen.MousePointer = vbHourglass

    Me.FrameSeleccion.Height = 5415
    Me.FrameSeleccion.Width = 8055
    Me.Height = 6120
    Me.Width = 8370
    

    Sql = " SELECT  pedidos.numpedid, pedidos.fechaped, pedidos.codclien, clientes.nomclien,  "
    Sql = Sql & "CASE pedidos.situacio WHEN 0 THEN ""Original"" WHEN 1 THEN ""Modificado"" WHEN 2 THEN ""Anulado"" END as situacio"
    Sql = Sql & " FROM pedidos, clientes where fechacar = " & DBSet(txtCodigo(6).Text, "F")
    Sql = Sql & " and pedidos.impresor = 0 "
    Sql = Sql & " and pedidos.codclien = clientes.codclien "
    Sql = Sql & " order by pedidos.fechaped, pedidos.numpedid "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        VisualizarListview True
    
        'Los encabezados
        
        ListView1.ColumnHeaders.Clear

        ListView1.ColumnHeaders.Add , , "Pedido", 1100
        ListView1.ColumnHeaders.Add , , "Fecha", 1100
        ListView1.ColumnHeaders.Add , , "Cliente", 1100
        ListView1.ColumnHeaders.Add , , "Nombre Cliente", 2500
        ListView1.ColumnHeaders.Add , , "Situacion", 1100
       
        ListView1.ListItems.Clear
        
        While Not Rs.EOF
            Set ItmX = ListView1.ListItems.Add
            'El primer campo será codtipom si llamamos desde Ventas
            ' y será codprove si llamamos desde Compras
            ItmX.Text = Format(Rs!numpedid, "00000000")
            ItmX.SubItems(1) = DBLet(Rs!FechaPed, "F")
            ItmX.SubItems(2) = DBLet(Rs!CodClien, "N")
            ItmX.SubItems(3) = DBLet(Rs!Nomclien, "T")
            ItmX.SubItems(4) = DBLet(Rs.Fields(4).Value, "T")
            HayReg = 1
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing
    CargarListView = HayReg
    
    Screen.MousePointer = vbDefault
    
ECargarList:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Function

Private Sub VisualizarListview(Modo As Boolean)
    If Modo = False Then
        Me.Width = 6570
    Else
        Me.Width = 8250
    End If
    FrameSeleccion.visible = Modo
    FrameCobros.visible = Not Modo
End Sub

Private Function ProcesarCambios() As Boolean
Dim Sql As String
Dim SQL1 As String
Dim i As Integer
Dim HayReg As Integer
Dim b As Boolean

On Error GoTo eProcesarCambios

    HayReg = 0
    
    VisualizarListview False
        
    conn.Execute "delete from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
         
        
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then
           
            InsertaLineaEnTemporal ListView1.ListItems(i)
            HayReg = 1
        End If
    Next i
    
    ProcesarCambios = (HayReg = 1)

eProcesarCambios:
    If Err.Number <> 0 Then
        ProcesarCambios = False
    End If
End Function


Private Sub InsertaLineaEnTemporal(ByRef ItmX As ListItem)
Dim Sql As String
Dim Codmacta As String
Dim Rs As ADODB.Recordset
Dim SQL1 As String

        SQL1 = "insert into tmpinformes(codusu, codigo1) values ("
        SQL1 = SQL1 & DBSet(vUsu.Codigo, "N") & "," & DBSet(ItmX.Text, "N") & ")"

        conn.Execute SQL1
    
End Sub

Private Sub ActualizarPedidos()
Dim Sql As String

    On Error GoTo EActPedidos
    
    If MsgBox("¿Impresión correcta para Actualizar Pedidos?", vbQuestion + vbYesNo) = vbYes Then
       Sql = "update pedidos set impresor = 1 where numpedid in (select codigo1 from tmpinformes where codusu = " & vUsu.Codigo & ")"
       conn.Execute Sql
    End If
    
    ' borramos la tabla temporal
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
    
EActPedidos:
    If Err.Number <> 0 Then Err.Clear
End Sub

