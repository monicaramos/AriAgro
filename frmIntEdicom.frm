VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmIntEdicom 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6630
   Icon            =   "frmIntEdicom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   6630
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
      Height          =   5535
      Left            =   45
      TabIndex        =   9
      Top             =   0
      Width           =   6555
      Begin VB.ComboBox Combo1 
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
         Left            =   1755
         TabIndex        =   0
         Text            =   "Combo1"
         Top             =   1215
         Width           =   1725
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
         Index           =   5
         Left            =   1755
         MaxLength       =   7
         TabIndex        =   2
         Tag             =   "N� Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   2115
         Width           =   1065
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
         Index           =   4
         Left            =   1755
         MaxLength       =   7
         TabIndex        =   1
         Tag             =   "N� Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   1725
         Width           =   1065
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
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   3135
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
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   2775
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
         Left            =   5040
         TabIndex        =   8
         Top             =   4950
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
         Left            =   3870
         TabIndex        =   7
         Top             =   4950
         Width           =   1065
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
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   5
         Top             =   3765
         Width           =   915
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
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   6
         Top             =   4140
         Width           =   915
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
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   3765
         Width           =   3540
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
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text5"
         Top             =   4140
         Width           =   3540
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
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
         Left            =   630
         TabIndex        =   23
         Top             =   4635
         Width           =   5145
      End
      Begin VB.Label Label1 
         Caption         =   "Integraci�n Facturas Edicom"
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
         Left            =   495
         TabIndex        =   22
         Top             =   315
         Width           =   5160
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo Movimiento"
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
         Left            =   495
         TabIndex        =   21
         Top             =   900
         Width           =   1815
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
         Index           =   1
         Left            =   780
         TabIndex        =   20
         Top             =   2115
         Width           =   645
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
         Index           =   0
         Left            =   780
         TabIndex        =   19
         Top             =   1755
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nro.Factura"
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
         Height          =   240
         Index           =   2
         Left            =   495
         TabIndex        =   18
         Top             =   1485
         Width           =   1170
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
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
         Left            =   495
         TabIndex        =   17
         Top             =   2475
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
         Left            =   780
         TabIndex        =   16
         Top             =   2775
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
         Left            =   780
         TabIndex        =   15
         Top             =   3135
         Width           =   645
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1485
         Picture         =   "frmIntEdicom.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   2775
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1485
         Picture         =   "frmIntEdicom.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   3135
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
         Left            =   780
         TabIndex        =   14
         Top             =   3765
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
         Index           =   12
         Left            =   780
         TabIndex        =   13
         Top             =   4140
         Width           =   645
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
         Height          =   240
         Index           =   11
         Left            =   510
         TabIndex        =   12
         Top             =   3525
         Width           =   675
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1485
         MouseIcon       =   "frmIntEdicom.frx":0122
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   3765
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1485
         MouseIcon       =   "frmIntEdicom.frx":0274
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   4140
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmIntEdicom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MANOLO +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar n� oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexi�n a BD Ariges  2.- Conexi�n a BD Conta

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
Dim indFrame As Single 'n� de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim Tabla As String
Dim Codigo As String 'C�digo para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String

Dim SQL As String

InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'A�adir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
     '======== FORMULA  ====================================
    'Seleccionar registros de la empresa conectada
'    Codigo = "{" & tabla & ".codempre}=" & vEmpresa.codEmpre
'    If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
'    If Not AnyadirAFormula(cadSelect, Codigo) Then Exit Sub
    
    'Tipo de movimiento
    Codigo = "{facturas.codtipom} = '" & Combo1(0).Text & "'"
    If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
    If Not AnyadirAFormula(cadselect, Codigo) Then Exit Sub
    
    'D/H Cliente
    cDesde = Trim(txtcodigo(0).Text)
    cHasta = Trim(txtcodigo(1).Text)
    nDesde = txtNombre(0).Text
    nHasta = txtNombre(1).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".codclien}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCliente= """) Then Exit Sub
    End If
    
    'D/H Nro de Factura
    cDesde = Trim(txtcodigo(4).Text)
    cHasta = Trim(txtcodigo(5).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{facturas.numfactu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFactura= """) Then Exit Sub
    End If
    
    'D/H Fecha factura
    cDesde = Trim(txtcodigo(2).Text)
    cHasta = Trim(txtcodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    cadselect = cadselect & " and facturas.pasedicom = 0"
    
    If HayRegistros(Tabla, cadselect) Then
        If Not HayRegistrosEnvases(cadselect) Then
            If vParamAplic.PathEdicom <> "" Then
                If Not ExistenFicheros Then
                    If ComprobarFicheros(cadselect) Then
                        SQL = "select count(*) from tmpinformes where codusu = " & vUsu.Codigo
                        
                        If TotalRegistros(SQL) <> 0 Then
                            MsgBox "Hay errores en la integraci�n EDICOM. Debe corregirlos previamente.", vbExclamation
                            cadTitulo = "Errores de integraci�n EDICOM"
                            cadNombreRPT = "rErroresEDICOM.rpt"
                            
                            cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
                            numParam = 1
                            
                            LlamarImprimir
                            Exit Sub
                        End If
                    
                        If GenerarFicheros(cadselect) Then
                            MsgBox "Proceso realizado correctamente", vbExclamation
                            cmdCancel_Click
                        Else
                            BorrarFicheros
                        End If
                    End If
                End If
            Else
                MsgBox "No existe directorio donde insertar los ficheros. Revise par�metros.", vbExclamation
            End If
        Else
            MsgBox "Hay Facturas con Envases. Revise.", vbExclamation
        End If
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
        Combo1(0).SetFocus
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
    Tabla = "facturas"
    Me.Label2.Caption = ""
    Me.Refresh
    CargaCombo
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdcancel.Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtcodigo(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Variedades
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Agencias de transporte
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
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
    If txtcodigo(Index).Text <> "" Then frmC.NovaData = txtcodigo(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtcodigo(CByte(imgFec(2).Tag))
    ' ***************************
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1 'CLIENTE
            AbrirFrmClientes (Index)
        
        Case 2, 3 'AGENCIAS DE TRANSPORTE
            AbrirFrmAgencias (Index)
        
        Case 4, 5 'VARIEDADES
            AbrirFrmVariedades (Index)
        
    End Select
    PonerFoco txtcodigo(indCodigo)
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
    ConseguirFoco txtcodigo(Index), 3
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
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
            
        Case 0, 1 'CLIENTE
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "clientes", "nomclien", "codclien", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000000")
        
        Case 2, 3 'FECHAS
            If txtcodigo(Index).Text <> "" Then PonerFormatoFecha txtcodigo(Index)
            
        Case 4, 5 'FACTURAS
            If txtcodigo(Index).Text <> "" Then PonerFormatoEntero txtcodigo(Index)
        
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 5760
        Me.FrameCobros.Width = 6690
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
'A�ade a cadFormula y cadSelect la cadena de seleccion:
'       "(codigo>=codD AND codigo<=codH)"
' y a�ade a cadParam la cadena para mostrar en la cabecera informe:
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
        .FormulaSeleccion = "{tmpinformes.codusu} = " & vUsu.Codigo
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
    frmVar.CodigoActual = txtcodigo(indCodigo)
    frmVar.Show vbModal
    Set frmVar = Nothing
End Sub

Private Sub AbrirFrmAgencias(indice As Integer)
    indCodigo = indice + 4
    Set frmTra = New frmManAgencias
    frmTra.DatosADevolverBusqueda = "0|1|"
    frmTra.Show vbModal
    Set frmTra = Nothing
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
Dim SQL As String
Dim Rs As ADODB.Recordset

    SQL = "Select * FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Rs.EOF Then
        MsgBox "No hay datos para mostrar en el Informe.", vbInformation
        HayRegistros = False
    Else
        HayRegistros = True
    End If

End Function

Private Function ProcesarCambios(cadwhere As String) As Boolean
Dim SQL As String
Dim SQL1 As String
Dim i As Integer
Dim HayReg As Integer
Dim b As Boolean

On Error GoTo eProcesarCambios

    HayReg = 0
    
    conn.Execute "delete from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
        
    If cadwhere <> "" Then
        cadwhere = QuitarCaracterACadena(cadwhere, "{")
        cadwhere = QuitarCaracterACadena(cadwhere, "}")
        cadwhere = QuitarCaracterACadena(cadwhere, "_1")
    End If
        
    SQL = "insert into tmpinformes (codusu, codigo1) select " & DBSet(vUsu.Codigo, "N")
    SQL = SQL & ", albaran.numalbar from albaran, albaran_variedad where albaran.numalbar not in (select numalbar from tcafpa) "
    SQL = SQL & " and albaran.numalbar = albaran_variedad.numalbar "
    
    If cadwhere <> "" Then SQL = SQL & " and " & cadwhere
    
    
    conn.Execute SQL
        
    ProcesarCambios = HayRegistros("tmpinformes", "codusu = " & vUsu.Codigo)

eProcesarCambios:
    If Err.Number <> 0 Then
        ProcesarCambios = False
    End If
End Function


Private Sub InsertaLineaEnTemporal(ByRef ItmX As ListItem)
Dim SQL As String
Dim Codmacta As String
Dim Rs As ADODB.Recordset
Dim SQL1 As String

        SQL1 = "insert into tmpinformes(codusu, codigo1) values ("
        SQL1 = SQL1 & DBSet(vUsu.Codigo, "N") & "," & DBSet(ItmX.Text, "N") & ")"

        conn.Execute SQL1
    
End Sub

Private Sub CargaCombo()
Dim Cad As String
Dim Rs As ADODB.Recordset
Dim i As Integer
    On Error GoTo ErrCarga
    Combo1(0).Clear
    'Conceptos
    
    Cad = "SELECT codtipom FROM usuarios.stipom WHERE codtipom like 'F%' ORDER BY codtipom"
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    'Combo1.AddItem "" 'pose uno en blanc sinse valor
    i = 0
    While Not Rs.EOF
        Combo1(0).AddItem Rs!codTipoM
        Combo1(0).ItemData(Combo1(0).NewIndex) = i
        Rs.MoveNext
        '.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    Exit Sub
    
ErrCarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar datos combo.", Err.Description
End Sub

Private Function DatosOk() As Boolean

    DatosOk = False
    If Combo1(0).ListIndex = -1 Then
        MsgBox "Debe selecionar un tipo de movimiento. Revise.", vbExclamation
        Combo1(0).SetFocus
    Else
        DatosOk = True
    End If


End Function


Private Function ExistenFicheros() As Boolean
Dim b As Boolean
Dim cadMen As String

    ExistenFicheros = False
    b = False
    
    
    If Dir(vParamAplic.PathEdicom, vbDirectory) = "" Then
        cadMen = "La carpeta de los ficheros de edicom " & vParamAplic.PathEdicom & " de par�metros no existe. Revise."
        MsgBox cadMen, vbExclamation
        ExistenFicheros = True
        Exit Function
    End If
    
    cadMen = "Los Ficheros : " & vbCrLf
    
    If Dir(vParamAplic.PathEdicom & "\cabfac.txt") <> "" Then
        cadMen = cadMen & vbCrLf & "        CABFAC.TXT"
        b = True
    End If
    If Dir(vParamAplic.PathEdicom & "\obsfac.txt") <> "" Then
        cadMen = cadMen & vbCrLf & "        OBSFAC.TXT"
        b = True
    End If
'--monica: 040608 ya no se genera el fichero de dtos.
    If Dir(vParamAplic.PathEdicom & "\dtofac.txt") <> "" Then
        cadMen = cadMen & vbCrLf & "        DTOFAC.TXT"
        b = True
    End If
    If Dir(vParamAplic.PathEdicom & "\impfac.txt") <> "" Then
        cadMen = cadMen & vbCrLf & "        IMPFAC.TXT"
        b = True
    End If
    If Dir(vParamAplic.PathEdicom & "\contenedfac.txt") <> "" Then
        cadMen = cadMen & vbCrLf & "        CONTENEDFAC.TXT"
        b = True
    End If
    If Dir(vParamAplic.PathEdicom & "\linfac.txt") <> "" Then
        cadMen = cadMen & vbCrLf & "        LINFAC.TXT"
        b = True
    End If
    If Dir(vParamAplic.PathEdicom & "\obslfac.txt") <> "" Then
        cadMen = cadMen & vbCrLf & "        OBSLFAC.TXT"
        b = True
    End If
    If Dir(vParamAplic.PathEdicom & "\dtolfac.txt") <> "" Then
        cadMen = cadMen & vbCrLf & "        DTOLFAC.TXT"
        b = True
    End If

    If b Then
        cadMen = cadMen & vbCrLf & vbCrLf & "existen en el directorio de edicom. Revise." & vbCrLf
        MsgBox cadMen, vbExclamation
    End If
    ExistenFicheros = b
End Function

Private Function GenerarFicheros(cadwhere As String) As Boolean
Dim b As Boolean
Dim SQL As String
Dim Mens As String
        
    On Error GoTo eGenerarFicheros
    
    b = True
    If b Then
        Mens = "Generando Cabecera de Factura"
        b = GeneraCABFAC(cadwhere, Mens)
    End If
    
    If b Then
        Mens = "Generando Observaciones de Factura"
        b = GeneraOBSFAC(cadwhere, Mens)
    End If

    If b Then
        Mens = "Generando Descuentos de Factura"
        b = GeneraDTOFAC(cadwhere, Mens)
    End If
    
'--monica:040608 nto se generan impuestos
'    If b Then
'        Mens = "Generando Impuestos de Factura"
'        b = GeneraIMPFAC(cadWhere)
'    End If
    
'    If b Then b = GeneraCONTENEDFAC(Sql)
    
    If b Then
        Mens = "Generando L�neas de Factura"
        b = GeneraLINFAC(cadwhere, Mens)
    End If
    
'    If b Then b = GeneraOBSLFAC(Sql)
'    If b Then b = GeneraDTOLFAC(Sql)
    
    If b Then
        SQL = "update facturas set pasedicom = 1 where " & cadwhere
        
        conn.Execute SQL
    End If
    
'    GenerarFicheros = b
'    Exit Function
    
eGenerarFicheros:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Error Actualizando Facturas. " & Err.Description
        GenerarFicheros = False
    Else
        If Not b Then
            MuestraError Err.Number, "Error en la Generaci�n de Ficheros: " & vbCrLf & Mens
            GenerarFicheros = False
        Else
            GenerarFicheros = True
        End If
    End If
End Function

Private Function RellenaABlancos(CADENA As String, PorLaDerecha As Boolean, Longitud As Integer) As String
Dim Cad As String
    
    Cad = Space(Longitud)
    If PorLaDerecha Then
        Cad = CADENA & Cad
        RellenaABlancos = Left(Cad, Longitud)
    Else
        Cad = Cad & CADENA
        RellenaABlancos = Right(Cad, Longitud)
    End If
    
End Function

Private Function GeneraCABFAC(cadwhere As String, Mens As String) As Boolean
Dim b As Boolean
Dim NF As Long
Dim Cad As String
Dim SQL As String
Dim SQL1 As String
Dim i As Integer
Dim Longitud As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim vCliente As CCliente
Dim Neto As Currency
Dim Impuestos As Currency
Dim DiasVto As String
Dim Dias As Integer
Dim FecVto As Date
Dim Descuen As Currency

    On Error GoTo eGeneraCABFAC
    
    b = True
    NF = FreeFile
    Open vParamAplic.PathEdicom & "\CABFAC.TXT" For Output As #NF
        
    '[Monica] 29/01/2010 enlazo con facturas_variedad para no coger facturas sin lineas
    SQL = "select distinct facturas.* from facturas INNER JOIN facturas_variedad ON "
    SQL = SQL & " facturas.codtipom = facturas_variedad.codtipom "
    SQL = SQL & " and facturas.numfactu = facturas_variedad.numfactu "
    SQL = SQL & " and facturas.fecfactu = facturas_variedad.fecfactu "
    If cadwhere <> "" Then SQL = SQL & " where " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Label2.Caption = "Generando Fichero CABFAC.TXT"
    Me.Refresh
    
    i = 0
    While Not Rs.EOF
        Cad = ""
        
        Set vCliente = New CCliente
    
        'si se ha modificado el cliente volver a cargar los datos
        If vCliente.Existe(Rs!CodClien) Then
            If vCliente.LeerDatos(Rs!CodClien) Then
                SQL1 = "select albaran.*, destinos.codigoedi, destinos.nomdesti, destinos.domdesti, destinos.pobdesti, destinos.codpobla "
                SQL1 = SQL1 & " from albaran, facturas_variedad, destinos "
                SQL1 = SQL1 & " where facturas_variedad.codtipom = " & DBSet(Rs!codTipoM, "T")
                SQL1 = SQL1 & " and facturas_variedad.numfactu = " & DBSet(Rs!NumFactu, "N")
                SQL1 = SQL1 & " and facturas_variedad.fecfactu = " & DBSet(Rs!FecFactu, "F")
                SQL1 = SQL1 & " and facturas_variedad.numalbar = albaran.numalbar "
                SQL1 = SQL1 & " and albaran.codclien = destinos.codclien "
                SQL1 = SQL1 & " and albaran.coddesti = destinos.coddesti "
                
                Set Rs1 = New ADODB.Recordset
                Rs1.Open SQL1, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
                Cad = Cad & RellenaABlancos(DBLet(Rs!NumFactu, "N"), True, 15) '1.-nro factura
                Cad = Cad & RellenaABlancos(vParamAplic.CodigoEdi, True, 17)  '2.-codigo edi vendedor
                Cad = Cad & RellenaABlancos(vParamAplic.CodigoEdi, True, 17) '3.-codigo edi emisor
                Cad = Cad & Space(17)                                       '4.-
                
                '[Monica]07/07/2016: en el caso de lidl el codigo edi del comprador es el del cliente
                If DBLet(Rs!CodClien, "N") = 104 Then
                    Cad = Cad & RellenaABlancos(vCliente.CodigoEdi, True, 17)   '5.-codigo edi comprador
                Else
                    Cad = Cad & RellenaABlancos(DBLet(Rs1!CodigoEdi, "T"), True, 17)   '5.-codigo edi comprador
                End If
                
                Cad = Cad & Space(13)                                       '6.-departamento
                Cad = Cad & RellenaABlancos(DBLet(Rs1!CodigoEdi, "T"), True, 17)  '7.-codigo edi receptor--> de la tabla de destinos
                
                If vCliente.DestEDI = 0 Then
                    Cad = Cad & RellenaABlancos(vCliente.CodigoEdi, True, 17)   '8.-codigo edi cliente, destinatario de la factura
                    Cad = Cad & RellenaABlancos(vCliente.CodigoEdi, True, 17)   '9.-pagador
                Else
                    Cad = Cad & RellenaABlancos(DBLet(Rs1!CodigoEdi, "T"), True, 17)   '8.-codigo edi cliente, destinatario de la factura
                    Cad = Cad & RellenaABlancos(DBLet(Rs1!CodigoEdi, "T"), True, 17)   '9.-pagador
                End If
                
                'cad = cad & Space(17)                                       '9.-
                Cad = Cad & RellenaABlancos(DBLet(Rs1!refclien, "N"), True, 17)  '10.-nro pedido-->albaran.refclien
                Cad = Cad & Format(DBLet(Rs!FecFactu, "F"), "YYYYMMDDhhmm") '11.-fecha emision de la factura
                Cad = Cad & "380"                                           '12.-tipo de factura
                Cad = Cad & Space(3)                                        '13.-
                
                If vCliente.DestEDI = 0 Then ' el destino de factura es el cliente
                    Cad = Cad & RellenaABlancos(vCliente.Nombre, True, 70)      '14.-nombre del cliente
                    Cad = Cad & RellenaABlancos(vCliente.Domicilio, True, 35)     '15.-domicilio del cliente
                    Cad = Cad & RellenaABlancos(vCliente.Poblacion, True, 35)     '16.-ciudad del cliente
                    Cad = Cad & Format(vCliente.CPostal, "00000")               '17.-codigo postal del cliente
                Else ' el destino de factura es el destinatario
'[Monica]23/02/2011: La direccion es la del cliente
'                    cad = cad & RellenaABlancos(RS1!nomdesti, True, 70)      '14.-nombre del cliente
'                    cad = cad & RellenaABlancos(RS1!domdesti, True, 35)     '15.-domicilio del cliente
'                    cad = cad & RellenaABlancos(RS1!pobdesti, True, 35)     '16.-ciudad del cliente
'                    cad = cad & Format(RS1!codPobla, "00000")               '17.-codigo postal del cliente
                    Cad = Cad & RellenaABlancos(vCliente.Nombre, True, 70)      '14.-nombre del cliente
                    Cad = Cad & RellenaABlancos(vCliente.Domicilio, True, 35)     '15.-domicilio del cliente
                    Cad = Cad & RellenaABlancos(vCliente.Poblacion, True, 35)     '16.-ciudad del cliente
                    Cad = Cad & Format(vCliente.CPostal, "00000")               '17.-codigo postal del cliente
                End If
                
                Cad = Cad & RellenaABlancos(vCliente.NIF, True, 17)         '18.-nif cliente
                Cad = Cad & Space(3)                                        '19.-razon
                Cad = Cad & RellenaABlancos(DBLet(Rs1!NumAlbar, "N"), True, 17) '20.-albaran
                Cad = Cad & Space(17)  '21.-
                Cad = Cad & Space(17)  '22.-
                Cad = Cad & Space(3)   '23.-
                '++monica:040608 antes no se cargaba
                Cad = Cad & "EUR" 'Space(3)   '24.-
                'cad = cad & String(11, "0") & "." & String(3, "0") '25.-
'--monica:030708
'                '25 y 26.-
'                Neto = DBLet(RS!baseimp1, "N") + DBLet(RS!baseimp2, "N") + DBLet(RS!baseimp3, "N")
'                If Neto >= 0 Then
'                    cad = cad & "+"
'                Else
'                    cad = cad & "-"
'                End If
'                cad = cad & Format(Neto, "0000000000.000")
'                '++ monica:040608 antes no se cargaba el bruto ahora s�, coincide con el neto.
'                If Neto >= 0 Then
'                    cad = cad & "+"
'                Else
'                    cad = cad & "-"
'                End If
'                cad = cad & Format(Neto, "0000000000.000")
                
'++monica:030708
                '25 y 26.-
                Neto = DBLet(Rs!baseimp1, "N") + DBLet(Rs!baseimp2, "N") + DBLet(Rs!baseimp3, "N")
                
                If DBLet(Rs!BrutoFac, "N") >= 0 Then
                    Cad = Cad & "+"
                Else
                    Cad = Cad & "-"
                End If
                Cad = Cad & Format(DBLet(Rs!BrutoFac, "N"), "0000000000.000")
                '++ monica:040608 antes no se cargaba el bruto ahora s�, coincide con el neto.
                If DBLet(Rs!BrutoFac, "N") >= 0 Then
                    Cad = Cad & "+"
                Else
                    Cad = Cad & "-"
                End If
                Cad = Cad & Format(DBLet(Rs!BrutoFac, "N"), "0000000000.000")

                
                Cad = Cad & String(11, "0") & "." & String(3, "0") '27.-
                
'--monica: 040708 ahora el campo descuen tiene que incluir el total de dtos.
'                cad = cad & String(11, "0") & "." & String(3, "0") '28.-
'++monica: 040708 ahora el campo descuen tiene que incluir el total de dtos.
                Descuen = DBLet(Rs!BrutoFac, "N") - Neto
                If Descuen >= 0 Then
                    Cad = Cad & "+"
                Else
                    Cad = Cad & "-"
                End If
                Cad = Cad & Format(Descuen, "0000000000.000")

                
                '29.- iva1
                If DBLet(Rs!baseimp1, "N") >= 0 Then
                    Cad = Cad & "+"
                Else
                    Cad = Cad & "-"
                End If
                Cad = Cad & Format(DBLet(Rs!baseimp1, "N"), "0000000000.000")
                '30.-
                Cad = Cad & "VAT"
                '31.-
                If DBLet(Rs!porciva1, "N") >= 0 Then
                    Cad = Cad & "+"
                Else
                    Cad = Cad & "-"
                End If
                Cad = Cad & Format(DBLet(Rs!porciva1, "N"), "000.000")
                '32.-
                If DBLet(Rs!impoiva1, "N") >= 0 Then
                    Cad = Cad & "+"
                Else
                    Cad = Cad & "-"
                End If
                Cad = Cad & Format(DBLet(Rs!impoiva1, "N"), "0000000000.000")
                
                
                '33.- iva 2
                If DBLet(Rs!baseimp2, "N") >= 0 Then
                    Cad = Cad & "+"
                Else
                    Cad = Cad & "-"
                End If
                Cad = Cad & Format(DBLet(Rs!baseimp2, "N"), "0000000000.000")
                '34.-
                If Not IsNull(Rs!codiiva2) Then
                    Cad = Cad & "VAT"
                Else
                    Cad = Cad & Space(3)
                End If
                '35.-
                If DBLet(Rs!porciva2, "N") >= 0 Then
                    Cad = Cad & "+"
                Else
                    Cad = Cad & "-"
                End If
                Cad = Cad & Format(DBLet(Rs!porciva2, "N"), "000.000")
                '36.-
                If DBLet(Rs!impoiva2, "N") >= 0 Then
                    Cad = Cad & "+"
                Else
                    Cad = Cad & "-"
                End If
                Cad = Cad & Format(DBLet(Rs!impoiva2, "N"), "0000000000.000")
                
                
                '37.-iva 3
                If DBLet(Rs!baseimp3, "N") >= 0 Then
                    Cad = Cad & "+"
                Else
                    Cad = Cad & "-"
                End If
                Cad = Cad & Format(DBLet(Rs!baseimp3, "N"), "0000000000.000")
                '38.-
                If Not IsNull(Rs!codiiva3) Then
                    Cad = Cad & "VAT"
                Else
                    Cad = Cad & Space(3)
                End If
                '39.-
                If DBLet(Rs!porciva3, "N") >= 0 Then
                    Cad = Cad & "+"
                Else
                    Cad = Cad & "-"
                End If
                Cad = Cad & Format(DBLet(Rs!porciva3, "N"), "000.000")
                '40.-
                If DBLet(Rs!impoiva3, "N") >= 0 Then
                    Cad = Cad & "+"
                Else
                    Cad = Cad & "-"
                End If
                Cad = Cad & Format(DBLet(Rs!impoiva3, "N"), "0000000000.000")
                
                
                Cad = Cad & Space(15) '41.-iva4
                Cad = Cad & Space(3) '42.-
                Cad = Cad & Space(8) '43.-
                Cad = Cad & Space(15) '44.-
                
                Cad = Cad & Space(15) '45.-iva4
                Cad = Cad & Space(3) '46.-
                Cad = Cad & Space(8) '47.-
                Cad = Cad & Space(15) '48.-
                
                Cad = Cad & Space(15) '49.-iva4
                Cad = Cad & Space(3) '50.-
                Cad = Cad & Space(8) '51.-
                Cad = Cad & Space(15) '52.-
                
                '53.-
                If Neto >= 0 Then
                    Cad = Cad & "+"
                Else
                    Cad = Cad & "-"
                End If
                Cad = Cad & Format(Neto, "0000000000.000")
                
                '54.-
                Impuestos = DBLet(Rs!impoiva1, "N") + DBLet(Rs!impoiva2, "N") + DBLet(Rs!impoiva3, "N")
                If Impuestos >= 0 Then
                    Cad = Cad & "+"
                Else
                    Cad = Cad & "-"
                End If
                Cad = Cad & Format(Impuestos, "0000000000.000")
                
                '55.-
                If DBLet(Rs!TotalFac, "N") >= 0 Then
                    Cad = Cad & "+"
                Else
                    Cad = Cad & "-"
                End If
                Cad = Cad & Format(DBLet(Rs!TotalFac, "N"), "0000000000.000")
                
                '56.- Fecha del primer vencimiento
                DiasVto = ""
                DiasVto = DevuelveDesdeBDNew(cAgro, "forpago", "primerve", "codforpa", Rs!Codforpa, "N")
                If DiasVto = "" Then
                    Dias = 0
                Else
                    Dias = CInt(DiasVto)
                End If
                FecVto = DateAdd("d", Dias, DBLet(Rs!FecFactu, "F"))
                Cad = Cad & Format(FecVto, "YYYYMMDD")
                
                '57.-
                If DBLet(Rs!TotalFac, "N") >= 0 Then
                    Cad = Cad & "+"
                Else
                    Cad = Cad & "-"
                End If
                Cad = Cad & Format(DBLet(Rs!TotalFac, "N"), "0000000000.000")
                
                
                Cad = Cad & Space(8) '58.-
                Cad = Cad & Space(15) '59.-
                Cad = Cad & Space(8) '60.-
                Cad = Cad & Space(15) '61.-
                
                Cad = Cad & Space(15) '62.-
                Cad = Cad & Space(3) '63.-
                
                Cad = Cad & Space(2) '64.-
                Cad = Cad & Space(3) '65.-
                Cad = Cad & Space(8) '66.-
                Cad = Cad & Space(15) '67.-
                Cad = Cad & Space(3) '68.-
                Cad = Cad & Space(2) '69.-
                Cad = Cad & Space(3) '70.-
                Cad = Cad & Space(8) '71.-
                Cad = Cad & Space(15) '72.-
                
                Cad = Cad & Space(3) '73.-
                Cad = Cad & Space(2) '74.-
                Cad = Cad & Space(3) '75.-
                Cad = Cad & Space(8) '76.-
                Cad = Cad & Space(15) '77.-
                Cad = Cad & Space(3) '78.-
                Cad = Cad & Space(2) '79.-
                Cad = Cad & Space(3) '80.-
                Cad = Cad & Space(8) '81.-
                Cad = Cad & Space(15) '82.-
                Cad = Cad & Space(3) '83.-
                Cad = Cad & Space(2) '84.-
                Cad = Cad & Space(3) '85.-
                Cad = Cad & Space(8) '86.-
                
                Cad = Cad & Space(15) '87.-
                
                Cad = Cad & RellenaABlancos(vParam.NombreEmpresa, True, 70) '88.-
                Cad = Cad & RellenaABlancos(vParam.DomicilioEmpresa, True, 35) '89.-
                Cad = Cad & RellenaABlancos(vParam.Poblacion, True, 35) '90.-
                Cad = Cad & Format(vParam.CPostal, "00000") '91.-
                Cad = Cad & RellenaABlancos(vParam.CifEmpresa, True, 17) '92.-
                Cad = Cad & RellenaABlancos(vParamAplic.RegMercantil, True, 70) '93.-
                Cad = Cad & Space(17) '94.-
                Cad = Cad & Space(17) '95.-
                Cad = Cad & Space(17) '96.-
                Cad = Cad & Space(17) '97.-
                
                Cad = Cad & Format(DBLet(Rs!FecFactu, "F"), "YYYYMMDDhhmm") '98.-
                Cad = Cad & Space(17) '99.-
                Cad = Cad & Space(17) '100.-
                Cad = Cad & Space(35) '101.-
                Cad = Cad & Space(35) '102.-
                Cad = Cad & Space(35) '103.-
                
                Cad = Cad & Space(25) '104.-
                '[Monica]05/07/2016: si codclien es lidl
                If DBLet(Rs!CodClien, "N") = 104 Then
                    Cad = Cad & Format(DBLet(Rs1!FechaAlb, "F"), "YYYYMMDDhhmm") '105.- fecha de albaran
                Else
                    Cad = Cad & Space(12) '105.-
                End If
                Cad = Cad & Space(35) '106.-
                Cad = Cad & Space(35) '107.-
                Cad = Cad & Space(35) '108.-
                Cad = Cad & Space(35) '109.-
                Cad = Cad & Space(35) '110.-
                Cad = Cad & Space(35) '111.-
                Cad = Cad & Space(8) '112.-
                Cad = Cad & Space(35) '113.-
                Cad = Cad & Space(17) '114.-
                Cad = Cad & Space(17) '115.-
                Cad = Cad & Space(17) '116.-
                Cad = Cad & Space(17) '117.-
                Cad = Cad & Space(17) '118.-
                Cad = Cad & Space(17) '119.-
                Cad = Cad & Space(17) '120.-
                Cad = Cad & Space(17) '121.-
                Cad = Cad & Space(15) '122.-
                Cad = Cad & Space(8) '123.-
                Cad = Cad & Space(8) '124.-
                Cad = Cad & Space(35) '125.-
                Cad = Cad & Space(16) '126.-
                Cad = Cad & Space(12) '127.-
                Cad = Cad & Space(17) '128.-
                Cad = Cad & Space(17) '129.-
                Cad = Cad & Space(12) '130.-
                Cad = Cad & Space(17) '131.-
                Cad = Cad & Space(12) '132.-
                Cad = Cad & Space(17) '133.-
                Cad = Cad & Space(35) '134.-
                Cad = Cad & Space(12) '135.-
                Cad = Cad & Space(35) '136.-
                Cad = Cad & Space(70) '137.-
                Cad = Cad & Space(35) '138.-
                Cad = Cad & Space(35) '139.-
                Cad = Cad & Space(17) '140.-
                Cad = Cad & Space(3) '141.-
                Cad = Cad & Space(35) '142.-
                Cad = Cad & Space(35) '143.-
                Cad = Cad & Space(12) '144.-
                Cad = Cad & Space(35) '145.-
                Cad = Cad & Space(17) '146.-
                Cad = Cad & Space(12) '147.-
                Cad = Cad & Space(35) '148.-
                
                Print #NF, Cad
                
                Set Rs1 = Nothing
                
            End If
        End If
        
        
        Rs.MoveNext
    Wend
    
    Rs.Close
    Set Rs = Nothing
    
    Set vCliente = Nothing
    
    Close #NF
    
    GeneraCABFAC = b
    Exit Function
    
eGeneraCABFAC:
    If Err.Number <> 0 Then
        Close #NF
        Set vCliente = Nothing
        GeneraCABFAC = False
        Mens = Mens & vbCrLf & Err.Description
    End If
    
End Function


Private Function GeneraOBSFAC(cadwhere As String, Mens As String) As Boolean
Dim b As Boolean
Dim Cad As String
Dim SQL As String
Dim Sql2 As String
Dim NF As Long
Dim i As Integer
Dim Longitud As Long
Dim Rs As ADODB.Recordset
Dim Neto As Currency
Dim Impuestos As Currency
    
    On Error GoTo eGeneraOBSFAC
    
    b = True
    
    '[Monica] 29/01/2010 enlazo con facturas_variedad para no coger facturas sin lineas
    Sql2 = "select distinct facturas.* from facturas INNER JOIN facturas_variedad ON "
    Sql2 = Sql2 & " facturas.codtipom = facturas_variedad.codtipom "
    Sql2 = Sql2 & " and facturas.numfactu = facturas_variedad.numfactu "
    Sql2 = Sql2 & " and facturas.fecfactu = facturas_variedad.fecfactu "
    Sql2 = Sql2 & " where not observac is null and observac <> '' "
    
    If cadwhere <> "" Then Sql2 = Sql2 & " and " & cadwhere
    
    If RegistrosAListar(Sql2) = 0 Then
        GeneraOBSFAC = True
        Exit Function
    End If
    
    '[Monica] 29/01/2010 enlazo con facturas_variedad para no coger facturas sin lineas
'    SQL = "select * from facturas where " & cadWhere & " and not observac is null and observac <> ''"
    SQL = "select distinct facturas.* from facturas INNER JOIN facturas_variedad ON "
    SQL = SQL & " facturas.codtipom = facturas_variedad.codtipom "
    SQL = SQL & " and facturas.numfactu = facturas_variedad.numfactu "
    SQL = SQL & " and facturas.fecfactu = facturas_variedad.fecfactu "
    SQL = SQL & " where not observac is null and observac <> '' "

    If cadwhere <> "" Then SQL = SQL & " and " & cadwhere

    NF = FreeFile
    Open vParamAplic.PathEdicom & "\OBSFAC.TXT" For Output As #NF
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Label2.Caption = "Generando Fichero OBSFAC.TXT"
    Me.Refresh
    
    i = 0
    While Not Rs.EOF
        Cad = ""
        Cad = Cad & RellenaABlancos(DBLet(Rs!NumFactu, "N"), True, 15) '1.-nro factura
        Cad = Cad & "00001" '2.- nro de observacion
        
        '[Monica]05/07/2016: para el caso de LIDL hay que poner en el tema SUR
        If DBLet(Rs!CodClien, "N") = 104 Then
            Cad = Cad & "SUR" '3.- tema
        Else
            Cad = Cad & Space(3) '3.- tema
        End If
        
        Cad = Cad & RellenaABlancos(Mid(DBLet(Rs!Observac, "T"), 1, 70), True, 70) '4.-observaciones
        
'        If Len(DBLet(RS!Observac, "T")) > 70 Then
        Cad = Cad & RellenaABlancos(Mid(DBLet(Rs!Observac, "T"), 71, 70), True, 70) '5.-observaciones
'        End If
'        If Len(DBLet(RS!Observac, "T")) > 140 Then
        Cad = Cad & RellenaABlancos(Mid(DBLet(Rs!Observac, "T"), 141, 70), True, 70) '5.-observaciones
'        End If
        Cad = Cad & Space(70)
        Cad = Cad & Space(70)
        
        Print #NF, Cad
        
        Rs.MoveNext
    Wend
    
    Rs.Close
    Set Rs = Nothing
    
    Close #NF
    
    GeneraOBSFAC = b
    Exit Function
    
eGeneraOBSFAC:
    If Err.Number <> 0 Then
        Close #NF
        GeneraOBSFAC = False
        Mens = Mens & vbCrLf & Err.Description
    End If
End Function



Private Function GeneraDTOFAC(cadwhere As String, Mens As String) As Boolean
Dim b As Boolean
Dim Cad As String
Dim NF As Long
Dim i As Integer
Dim Longitud As Long
Dim Rs As ADODB.Recordset
Dim Neto As Currency
Dim Impuestos As Currency
Dim SQL As String

Dim importe1 As Currency
Dim importe2 As Currency

    
    On Error GoTo eGeneraDTOFAC
    
    b = True
    NF = FreeFile
    Open vParamAplic.PathEdicom & "\DTOFAC.TXT" For Output As #NF
    
    '[Monica] 29/01/2010 enlazo con facturas_variedad para no coger facturas sin lineas
    SQL = "select distinct facturas.* from facturas INNER JOIN facturas_variedad ON "
    SQL = SQL & " facturas.codtipom = facturas_variedad.codtipom "
    SQL = SQL & " and facturas.numfactu = facturas_variedad.numfactu "
    SQL = SQL & " and facturas.fecfactu = facturas_variedad.fecfactu "
    SQL = SQL & " where (facturas.dtocom1 <> 0 or facturas.dtocom2 <> 0)"
    If cadwhere <> "" Then SQL = SQL & " and " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Label2.Caption = "Generando Fichero OBSFAC.TXT"
    Me.Refresh
    
    i = 0
    While Not Rs.EOF
        
        If DBLet(Rs!dtocom1, "N") <> 0 And DBLet(Rs!dtocom2, "N") <> 0 Then
            importe1 = Round2(DBLet(Rs!BrutoFac, "N") * DBLet(Rs!dtocom1, "N") / 100, 2)
            
            importe2 = DBLet(Rs!impordto, "N") - importe1
            
            Cad = ""
            Cad = Cad & RellenaABlancos(DBLet(Rs!NumFactu, "N"), True, 15) '1.-nro factura
            Cad = Cad & "00001" '2.- nro de descuento
            Cad = Cad & "A  " '3.- calificador de cargo o descuento
            Cad = Cad & Space(2) '4.- secuencia de aplicacion
            Cad = Cad & "TD " '5.- tipo de descuento: dtocom1 se corresponde con el descuento comercial
        
            '6.- porcentaje de descuento o cargo
            If DBLet(Rs!dtocom1, "N") >= 0 Then
                Cad = Cad & "+"
            Else
                Cad = Cad & "-"
            End If
            Cad = Cad & Format(DBLet(Rs!dtocom1, "N"), "000.000")
        
            '7.- importe de descuento o cargo
            If importe1 >= 0 Then
                Cad = Cad & "+"
            Else
                Cad = Cad & "-"
            End If
            Cad = Cad & Format(importe1, "0000000000.000")
        
        
            Print #NF, Cad
            
            Cad = ""
            Cad = Cad & RellenaABlancos(DBLet(Rs!NumFactu, "N"), True, 15) '1.-nro factura
            Cad = Cad & "00002" '2.- nro de descuento
            Cad = Cad & "A  "  '3.- calificador de cargo o descuento
            Cad = Cad & Space(2) '4.- secuencia de aplicacion
            Cad = Cad & "EAB" '5.- tipo de descuento: dtocom1 se corresponde con el descuento comercial
            '6.- porcentaje de descuento o cargo
            If DBLet(Rs!dtocom2, "N") >= 0 Then
                Cad = Cad & "+"
            Else
                Cad = Cad & "-"
            End If
            Cad = Cad & Format(DBLet(Rs!dtocom2, "N"), "000.000")
        
            '7.- importe de descuento o cargo
            If importe2 >= 0 Then
                Cad = Cad & "+"
            Else
                Cad = Cad & "-"
            End If
            Cad = Cad & Format(importe2, "0000000000.000")
        
            Print #NF, Cad
        
        Else
            If DBLet(Rs!dtocom1, "N") <> 0 Then
                importe1 = DBLet(Rs!impordto, "N")
                
                Cad = ""
                Cad = Cad & RellenaABlancos(DBLet(Rs!NumFactu, "N"), True, 15) '1.-nro factura
                Cad = Cad & "00001" '2.- nro de descuento
                Cad = Cad & "A  "  '3.- calificador de cargo o descuento
                Cad = Cad & Space(2) '4.- secuencia de aplicacion
                Cad = Cad & "TD " '5.- tipo de descuento: dtocom1 se corresponde con el descuento comercial
                '6.- porcentaje de descuento o cargo
                If DBLet(Rs!dtocom1, "N") >= 0 Then
                    Cad = Cad & "+"
                Else
                    Cad = Cad & "-"
                End If
                Cad = Cad & Format(DBLet(Rs!dtocom1, "N"), "000.000")
            
                '7.- importe de descuento o cargo
                If DBLet(Rs!impordto, "N") >= 0 Then
                    Cad = Cad & "+"
                Else
                    Cad = Cad & "-"
                End If
                Cad = Cad & Format(DBLet(Rs!impordto, "N"), "0000000000.000")
            
                Print #NF, Cad
            Else
                If DBLet(Rs!dtocom2, "N") <> 0 Then
                    importe2 = DBLet(Rs!impordto, "N")
                
                    Cad = ""
                    Cad = Cad & RellenaABlancos(DBLet(Rs!NumFactu, "N"), True, 15) '1.-nro factura
                    Cad = Cad & "00001" '2.- nro de descuento
                    Cad = Cad & "A  "  '3.- calificador de cargo o descuento
                    Cad = Cad & Space(2) '4.- secuencia de aplicacion
                    Cad = Cad & "EAB" '5.- tipo de descuento: dtocom1 se corresponde con el descuento comercial
                    '6.- porcentaje de descuento o cargo
                    If DBLet(Rs!dtocom2, "N") >= 0 Then
                        Cad = Cad & "+"
                    Else
                        Cad = Cad & "-"
                    End If
                    Cad = Cad & Format(DBLet(Rs!dtocom2, "N"), "000.000")
                
                    '7.- importe de descuento o cargo
                    If DBLet(Rs!impordto, "N") >= 0 Then
                        Cad = Cad & "+"
                    Else
                        Cad = Cad & "-"
                    End If
                    Cad = Cad & Format(DBLet(Rs!impordto, "N"), "0000000000.000")
                
                    Print #NF, Cad
                End If
            End If
        End If
        
        
        Rs.MoveNext
    Wend
    
    Rs.Close
    Set Rs = Nothing
    
    Close #NF
    
    GeneraDTOFAC = b
    Exit Function
    
eGeneraDTOFAC:
    If Err.Number <> 0 Then
        Close #NF
        GeneraDTOFAC = False
        Mens = Mens & vbCrLf & Err.Description
    End If
End Function


Private Function GeneraIMPFAC(cadwhere As String, Mens As String) As Boolean
Dim b As Boolean
Dim Cad As String
Dim SQL As String
Dim NF As Long
Dim i As Integer
Dim Longitud As Long
Dim Rs As ADODB.Recordset
Dim Neto As Currency
Dim Impuestos As Currency
    
    
    On Error GoTo eGeneraIMPFAC
    
    b = True
    NF = FreeFile
    Open vParamAplic.PathEdicom & "\IMPFAC.TXT" For Output As #NF
    
    SQL = "select * from facturas where " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Label2.Caption = "Generando Fichero IMPFAC.TXT"
    Me.Refresh
    
    i = 0
    While Not Rs.EOF
        If Not IsNull(Rs!codiiva1) Then
            Cad = ""
            Cad = Cad & RellenaABlancos(DBLet(Rs!NumFactu, "N"), True, 15) '1.-nro factura
            Cad = Cad & "00001" '2.- nro de impuesto
            
            '3.- base imponible
            If DBLet(Rs!baseimp1, "N") >= 0 Then
                Cad = Cad & "+"
            Else
                Cad = Cad & "-"
            End If
            Cad = Cad & Format(DBLet(Rs!baseimp1, "N"), "0000000000.000")
            
            Cad = Cad & "VAT"  '4.- tipo de iva
            
            '5.- porcentaje de impuesto
            If DBLet(Rs!porciva1, "N") >= 0 Then
                Cad = Cad & "+"
            Else
                Cad = Cad & "-"
            End If
            Cad = Cad & Format(DBLet(Rs!porciva1, "N"), "000.000")

            '6.- importe de impuesto
            If DBLet(Rs!impoiva1, "N") >= 0 Then
                Cad = Cad & "+"
            Else
                Cad = Cad & "-"
            End If
            Cad = Cad & Format(DBLet(Rs!impoiva1, "N"), "0000000000.000")
            
            Print #NF, Cad
        End If
        
        If Not IsNull(Rs!codiiva2) Then
            Cad = ""
            Cad = Cad & RellenaABlancos(DBLet(Rs!NumFactu, "N"), True, 15) '1.-nro factura
            Cad = Cad & "00002" '2.- nro de impuesto
            
            '3.- base imponible
            If DBLet(Rs!baseimp2, "N") >= 0 Then
                Cad = Cad & "+"
            Else
                Cad = Cad & "-"
            End If
            Cad = Cad & Format(DBLet(Rs!baseimp2, "N"), "0000000000.000")
            
            Cad = Cad & "VAT"  '4.- tipo de iva
            
            '5.- porcentaje de impuesto
            If DBLet(Rs!porciva2, "N") >= 0 Then
                Cad = Cad & "+"
            Else
                Cad = Cad & "-"
            End If
            Cad = Cad & Format(DBLet(Rs!porciva2, "N"), "000.000")

            '6.- importe de impuesto
            If DBLet(Rs!impoiva2, "N") >= 0 Then
                Cad = Cad & "+"
            Else
                Cad = Cad & "-"
            End If
            Cad = Cad & Format(DBLet(Rs!impoiva2, "N"), "0000000000.000")
            
            Print #NF, Cad
        End If
        
        If Not IsNull(Rs!codiiva3) Then
            Cad = ""
            Cad = Cad & RellenaABlancos(DBLet(Rs!NumFactu, "N"), True, 15) '1.-nro factura
            Cad = Cad & "00003" '2.- nro de impuesto
            
            '3.- base imponible
            If DBLet(Rs!baseimp3, "N") >= 0 Then
                Cad = Cad & "+"
            Else
                Cad = Cad & "-"
            End If
            Cad = Cad & Format(DBLet(Rs!baseimp3, "N"), "0000000000.000")
            
            Cad = Cad & "VAT"  '4.- tipo de iva
            
            '5.- porcentaje de impuesto
            If DBLet(Rs!porciva3, "N") >= 0 Then
                Cad = Cad & "+"
            Else
                Cad = Cad & "-"
            End If
            Cad = Cad & Format(DBLet(Rs!porciva3, "N"), "000.000")

            '6.- importe de impuesto
            If DBLet(Rs!impoiva3, "N") >= 0 Then
                Cad = Cad & "+"
            Else
                Cad = Cad & "-"
            End If
            Cad = Cad & Format(DBLet(Rs!impoiva3, "N"), "0000000000.000")
            
            Print #NF, Cad
        End If
        
        Rs.MoveNext
    Wend
    
    Rs.Close
    Set Rs = Nothing
    
    Close #NF
    
    GeneraIMPFAC = b
    Exit Function
    
eGeneraIMPFAC:
    If Err.Number <> 0 Then
        Close #NF
        GeneraIMPFAC = False
        Mens = Mens & vbCrLf & Err.Description
    End If
End Function


Private Function GeneraLINFAC(cadwhere As String, Mens As String) As Boolean
Dim b As Boolean
Dim Cad As String
Dim SQL As String
Dim NF As Long
Dim i As Integer
Dim Longitud As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim vCliente As CCliente
Dim Neto As Currency
Dim Impuestos As Currency
Dim REFEAN As String
Dim NomArtic As String
Dim PrecioNeto As Currency
Dim TipFac As Byte
                
Dim cadAux As String
Dim PorcIva As Currency

    On Error GoTo eGeneraLINFAC

    b = True
    NF = FreeFile
    Open vParamAplic.PathEdicom & "\LINFAC.TXT" For Output As #NF
    
    '[Monica] 29/01/2010 enlazo con facturas_variedad para no coger facturas sin lineas
    SQL = "select distinct facturas.* from facturas INNER JOIN facturas_variedad ON "
    SQL = SQL & "facturas.codtipom = facturas_variedad.codtipom "
    SQL = SQL & " and facturas.numfactu = facturas_variedad.numfactu "
    SQL = SQL & " and facturas.fecfactu = facturas_variedad.fecfactu "
    If cadwhere <> "" Then SQL = SQL & " where " & cadwhere
    
    Set Rs1 = New ADODB.Recordset
    Rs1.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Label2.Caption = "Generando Fichero LINFAC.TXT"
    Me.Refresh
    
    
    While Not Rs1.EOF
        i = 0
        
        '***** INSERTAMOS LAS LINEAS DE VARIEDADES *****
        SQL = "select facturas_variedad.*, albaran_variedad.codforfait, albaran_variedad.codvarie from facturas_variedad, albaran_variedad where "
        SQL = SQL & " facturas_variedad.codtipom = " & DBSet(Rs1!codTipoM, "T")
        SQL = SQL & " and facturas_variedad.numfactu = " & DBSet(Rs1!NumFactu, "N")
        SQL = SQL & " and facturas_variedad.fecfactu = " & DBSet(Rs1!FecFactu, "F")
        SQL = SQL & " and facturas_variedad.numalbar = albaran_variedad.numalbar "
        SQL = SQL & " and facturas_variedad.numlinealbar = albaran_variedad.numlinea "
        SQL = SQL & " order by facturas_variedad.codtipom, facturas_variedad.numfactu, "
        SQL = SQL & " facturas_variedad.fecfactu, facturas_variedad.numlinea "
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
        While Not Rs.EOF
            Cad = ""
            
            i = i + 1
            
            REFEAN = ""
            REFEAN = DevuelveDesdeBDNew(cAgro, "codigoean", "codigoean", "codclien", Rs1!CodClien, "N", , "codforfait", Rs!codforfait, "T", "codvarie", Rs!codvarie, "N")
            
            '[Monica]15/07/2016: cambiamos antes salia el nombre de la variedad, ahora el nombre de la confeccion
            NomArtic = ""
            NomArtic = DevuelveDesdeBDNew(cAgro, "forfaits", "nomconfe", "codforfait", Rs!codforfait, "T")
            
            TipFac = TipoFacturarForfaits(CStr(Rs!NumAlbar), CStr(Rs!numlinealbar))
            
            Cad = Cad & RellenaABlancos(DBLet(Rs!NumFactu, "N"), True, 15) '1.-nro factura
            Cad = Cad & Format(i, "00000")  '2.-nro linea
            Cad = Cad & RellenaABlancos(REFEAN, True, 17) '3.-referencia del articulo
            Cad = Cad & Space(35)      '4.-
            Cad = Cad & Space(35)      '5.-
            Cad = Cad & Space(2)       '6.-
            Cad = Cad & Space(14)      '7.-
            Cad = Cad & RellenaABlancos(NomArtic, True, 70)         '8.- Descripcion del articulo
            
            '9.- Cantidad Facturada
            Select Case TipFac
                Case 0 'unidades
                    If DBLet(Rs!Unidades, "N") >= 0 Then
                        Cad = Cad & "+"
                    Else
                        Cad = Cad & "-"
                    End If
                    Cad = Cad & Format(DBLet(Rs!Unidades, "N"), "0000000000.000")
                    Cad = Cad & Space(10) '10.-
                    Cad = Cad & "PCE" '11.-
                Case 1 'kilos
                    If DBLet(Rs!cantreal, "N") >= 0 Then
                        Cad = Cad & "+"
                    Else
                        Cad = Cad & "-"
                    End If
                    Cad = Cad & Format(DBLet(Rs!cantreal, "N"), "0000000000.000")
                    Cad = Cad & Space(10) '10.-
                    Cad = Cad & "KGM" '11.-
            End Select
            
            '12.- Precio Bruto
            If DBLet(Rs!precibru, "N") >= 0 Then
                Cad = Cad & "+"
            Else
                Cad = Cad & "-"
            End If
            Cad = Cad & Format(DBLet(Rs!precibru, "N"), "0000000000.000")
            
            '[Monica]07/07/2016: si es lidl dejamos a cero la posicion del precio neto
            If DBLet(Rs1!CodClien, "N") = 104 Then
                Cad = Cad & Space(15)
            Else
                '13.- Precio Neto
                If DBLet(Rs!precibru, "N") >= 0 Then
                    Cad = Cad & "+"
                Else
                    Cad = Cad & "-"
                End If
                Cad = Cad & Format(DBLet(Rs!precibru, "N"), "0000000000.000")
            End If
            
'--monica: si hay descuentos de cabecera los precios son iguales y coinciden son el bruto.
'            '13.- Precio Neto
'            If DBLet(RS!precinet, "N") >= 0 Then
'                cad = cad & "+"
'            Else
'                cad = cad & "-"
'            End If
'            cad = cad & Format(DBLet(RS!precinet, "N"), "0000000000.000")
            
            '[Monica]05/07/2016: en el caso de lidl es obligatorio
            If DBLet(Rs1!CodClien, "N") = 104 Then
                Cad = Cad & "VAT" '14.- tipo de impuesto
                
                '15.- porcentaje de iva
                PorcIva = 0
                cadAux = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", Rs!Codigiva, "N")
                If cadAux <> "" Then PorcIva = ImporteSinFormato(cadAux)
                
                If DBLet(PorcIva, "N") >= 0 Then
                    Cad = Cad & "+"
                Else
                    Cad = Cad & "-"
                End If
                Cad = Cad & Format(DBLet(PorcIva, "N"), "000.000")
            Else
                Cad = Cad & Space(3) '14.-
                Cad = Cad & Space(8) '15.-
            End If
            
            Cad = Cad & Space(15) '16.-
            Cad = Cad & Space(3) '17.-
            Cad = Cad & Space(8) '18.-
            Cad = Cad & Space(15) '19.-
            Cad = Cad & Space(3) '20.-
            Cad = Cad & Space(8) '21.-
            Cad = Cad & Space(15) '22.-
            Cad = Cad & Space(3) '23.-
            Cad = Cad & Space(2) '24.-
            Cad = Cad & Space(3) '25.-
            Cad = Cad & Space(8) '26.-
            Cad = Cad & Space(15) '27.-
            Cad = Cad & Space(3) '28.-
            Cad = Cad & Space(2) '29.-
            Cad = Cad & Space(3) '30.-
            Cad = Cad & Space(8) '31.-
            Cad = Cad & Space(15) '32.-
            Cad = Cad & Space(3) '33.-
            Cad = Cad & Space(2) '34.-
            Cad = Cad & Space(3) '35.-
            Cad = Cad & Space(8) '36.-
            Cad = Cad & Space(15) '37.-
            Cad = Cad & Space(3) '38.-
            Cad = Cad & Space(2) '39.-
            Cad = Cad & Space(3) '40.-
            Cad = Cad & Space(8) '41.-
            Cad = Cad & Space(15) '42.-
            Cad = Cad & Space(15) '43.-
            
            '44.- Importe neto
            '++monica:010708 como tiene que cuadrar linea cantidad por precio bruto ponemos el importe bruto
            If DBLet(Rs!imporbru, "N") >= 0 Then
                Cad = Cad & "+"
            Else
                Cad = Cad & "-"
            End If
            Cad = Cad & Format(DBLet(Rs!imporbru, "N"), "0000000000.000")
            
            Cad = Cad & Space(15) '45.- Punto verde ????????
            Cad = Cad & Space(17) '46.-
            Cad = Cad & Space(17) '47.-
            Cad = Cad & Space(17) '48.-
            Cad = Cad & Space(12) '49.-
            Cad = Cad & Space(12) '50.-
            Cad = Cad & Space(17) '51.-
            Cad = Cad & Space(3) '52.-
            Cad = Cad & Space(15) '53.-
            Cad = Cad & Space(17) '54.-
            Cad = Cad & Space(17) '55.-
            Cad = Cad & Space(15) '56.-
            Cad = Cad & Space(12) '57.-
            Cad = Cad & Space(35) '58.-
            Cad = Cad & Space(35) '59.-
            Cad = Cad & Space(15) '60.-
            Cad = Cad & Space(3) '61.-
            Cad = Cad & Space(12) '62.-
            Cad = Cad & Space(35) '63.-
            Cad = Cad & Space(3) '64.-
        
            Print #NF, Cad
            
            Rs.MoveNext
            
        Wend
    
        Set Rs = Nothing
    
' no hay envases de las facturas que mandamos a edicom
'        '***** INSERTAMOS LAS LINEAS DE ENVASES *****
'        Sql = "select facturas_envases.* from facturas_envases where "
'        Sql = Sql & " facturas_envases.codtipom = " & DBSet(Rs1!codTipoM, "T")
'        Sql = Sql & " and facturas_envases.numfactu = " & DBSet(Rs1!NumFactu, "N")
'        Sql = Sql & " and facturas_envases.fecfactu = " & DBSet(Rs1!FecFactu, "F")
'        Sql = Sql & " order by facturas_envases.codtipom, facturas_envases.numfactu, "
'        Sql = Sql & " facturas_envases.fecfactu, facturas_envases.numlinea "
'
'        Set RS = New ADODB.Recordset
'        RS.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'        While Not RS.EOF
'            cad = ""
'
'            i = i + 1
'
'            RefEAN = ""
'            RefEAN = DevuelveDesdeBDNew(cAgro, "codigoea", "sartic", "codartic", RS!codArtic, "T")
'
'            NomArtic = ""
'            NomArtic = DevuelveDesdeBDNew(cAgro, "nomartic", "sartic", "codartic", RS!codArtic, "T")
'
'            cad = cad & RellenaABlancos(DBLet(RS!NumFactu, "N"), True, 15) '1.-nro factura
'            cad = cad & Format(i, "00000")  '2.-nro linea
'            cad = cad & RellenaABlancos(RefEAN, True, 17) '3.-referencia del articulo
'            cad = cad & Space(35)      '4.-
'            cad = cad & Space(35)      '5.-
'            cad = cad & Space(2)       '6.-
'            cad = cad & Space(14)      '7.-
'            cad = cad & RellenaABlancos(NomArtic, True, 70)         '8.- Descripcion del articulo
'            '9.- Cantidad Facturada
'            If DBLet(RS!cantidad, "N") >= 0 Then
'                cad = cad & "+"
'            Else
'                cad = cad & "-"
'            End If
'            cad = cad & Format(DBLet(RS!cantidad, "N"), "0000000000.000")
'
'            cad = cad & Space(10) '10.-
'            cad = cad & Space(3) '11.-
'
'            '12.- Precio Bruto
'            If DBLet(RS!precioar, "N") >= 0 Then
'                cad = cad & "+"
'            Else
'                cad = cad & "-"
'            End If
'            cad = cad & Format(DBLet(RS!precioar, "N"), "0000000000.000")
'
'            '13.- Precio Neto
'            PrecioNeto = 0
'            If DBLet(RS!cantidad, "N") <> 0 Then
'                PrecioNeto = Round2(DBLet(RS!ImporteL, "N") / DBLet(RS!cantidad, "N"), 3)
'            End If
'            If PrecioNeto >= 0 Then
'                cad = cad & "+"
'            Else
'                cad = cad & "-"
'            End If
'            cad = cad & Format(PrecioNeto, "0000000000.000")
'
'            cad = cad & Space(3) '14.-
'            cad = cad & Space(8) '15.-
'            cad = cad & Space(15) '16.-
'            cad = cad & Space(3) '17.-
'            cad = cad & Space(8) '18.-
'            cad = cad & Space(15) '19.-
'            cad = cad & Space(3) '20.-
'            cad = cad & Space(8) '21.-
'            cad = cad & Space(15) '22.-
'            cad = cad & Space(3) '23.-
'            cad = cad & Space(2) '24.-
'            cad = cad & Space(3) '25.-
'            cad = cad & Space(8) '26.-
'            cad = cad & Space(15) '27.-
'            cad = cad & Space(3) '28.-
'            cad = cad & Space(2) '29.-
'            cad = cad & Space(3) '30.-
'            cad = cad & Space(8) '31.-
'            cad = cad & Space(15) '32.-
'            cad = cad & Space(3) '33.-
'            cad = cad & Space(2) '34.-
'            cad = cad & Space(3) '35.-
'            cad = cad & Space(8) '36.-
'            cad = cad & Space(15) '37.-
'            cad = cad & Space(3) '38.-
'            cad = cad & Space(2) '39.-
'            cad = cad & Space(3) '40.-
'            cad = cad & Space(8) '41.-
'            cad = cad & Space(15) '42.-
'            cad = cad & Space(15) '43.-
'            '44.- Importe neto
'            If DBLet(RS!ImporteL, "N") >= 0 Then
'                cad = cad & "+"
'            Else
'                cad = cad & "-"
'            End If
'            cad = cad & Format(DBLet(RS!ImporteL, "N"), "0000000000.000")
'
'            cad = cad & Space(15) '45.- Punto verde ????????
'            cad = cad & Space(17) '46.-
'            cad = cad & Space(17) '47.-
'            cad = cad & Space(17) '48.-
'            cad = cad & Space(12) '49.-
'            cad = cad & Space(12) '50.-
'            cad = cad & Space(17) '51.-
'            cad = cad & Space(3) '52.-
'            cad = cad & Space(15) '53.-
'            cad = cad & Space(17) '54.-
'            cad = cad & Space(17) '55.-
'            cad = cad & Space(15) '56.-
'            cad = cad & Space(12) '57.-
'            cad = cad & Space(35) '58.-
'            cad = cad & Space(35) '59.-
'            cad = cad & Space(15) '60.-
'            cad = cad & Space(3) '61.-
'            cad = cad & Space(12) '62.-
'            cad = cad & Space(35) '63.-
'            cad = cad & Space(3) '64.-
'
'            Print #nf, cad
'
'            RS.MoveNext
'
'        Wend
    
'        Set RS = Nothing
    
    
        Rs1.MoveNext
    Wend
    
    Rs1.Close
    Set Rs1 = Nothing
    
    Close #NF
    
    GeneraLINFAC = b
    Exit Function
    
eGeneraLINFAC:
    If Err.Number <> 0 Then
        Close #NF
        GeneraLINFAC = False
        Mens = Mens & vbCrLf & Err.Description
    End If
    
'    SQL = "select facturas_variedad.*, facturas.codclien from facturas, facturas_variedad where " & cad
'    SQL = SQL & " and facturas.codtipom = facturas_variedad.codtipom "
'    SQL = SQL & " and facturas.numfactu = facturas_variedad.numfactu "
'    SQL = SQL & " and facturas.fecfactu = facturas_variedad.fecfactu "
'    SQL = SQL & " order by facturas_variedad.codtipom, facturas_variedad.numfactu, "
'    SQL = SQL & " facturas_variedad.fecfactu, facturas_variedad.numlinea "
'
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'    Label2.Caption = "Generando Fichero LINFAC.TXT"
'
'    i = 0
'    While Not RS.EOF
'        cad = ""
'
'        RefEAN = ""
'        RefEAN = DevuelveDesdeBDNew(vagro, "codigoean", "codigoean", "codclien", RS!CodClien, "N", , "codforfait", RS!codforfait, "T", "codvarie", RS!codvarie, "N")
'
'        NomArtic = ""
'        NomArtic = DevuelveDesdeBDNew(vagro, "variedades", "nomvarie", "codvarie", RS!codvarie, "N")
'
'        cad = cad & RellenaABlancos(DBLet(RS!NumFactu, "N"), True, 15) '1.-nro factura
'        cad = cad & Format(DBLet(RS!numlinea, "N"), "00000")  '2.-nro linea
'        cad = cad & RellenaABlancos(RefEAN, True, 17) '3.-referencia del articulo
'        cad = cad & Space(35)      '4.-
'        cad = cad & Space(35)      '5.-
'        cad = cad & Space(2)       '6.-
'        cad = cad & Space(14)      '7.-
'        cad = cad & RellenaABlancos(NomArtic, True, 70)         '8.- Descripcion del articulo
'        '9.- Cantidad Facturada
'        If DBLet(RS!cantfact, "N") >= 0 Then
'            cad = cad & "+"
'        Else
'            cad = cad & "-"
'        End If
'        cad = cad & Format(DBLet(RS!cantfact, "N"), "0000000000.000")
'        cad = cad & Space(10) '10.-
'        cad = cad & "KGM" '11.-
'
'        '12.- Precio Bruto
'        If DBLet(RS!precibru, "N") >= 0 Then
'            cad = cad & "+"
'        Else
'            cad = cad & "-"
'        End If
'        cad = cad & Format(DBLet(RS!precibru, "N"), "0000000000.000")
'
'        '13.- Precio Neto
'        If DBLet(RS!precinet, "N") >= 0 Then
'            cad = cad & "+"
'        Else
'            cad = cad & "-"
'        End If
'        cad = cad & Format(DBLet(RS!precinet, "N"), "0000000000.000")
'        cad = cad & Space(3) '14.-
'        cad = cad & Space(8) '15.-
'        cad = cad & Space(15) '16.-
'        cad = cad & Space(3) '17.-
'        cad = cad & Space(8) '18.-
'        cad = cad & Space(15) '19.-
'        cad = cad & Space(3) '20.-
'        cad = cad & Space(8) '21.-
'        cad = cad & Space(15) '22.-
'        cad = cad & Space(3) '23.-
'        cad = cad & Space(2) '24.-
'        cad = cad & Space(3) '25.-
'        cad = cad & Space(8) '26.-
'        cad = cad & Space(15) '27.-
'        cad = cad & Space(3) '28.-
'        cad = cad & Space(2) '29.-
'        cad = cad & Space(3) '30.-
'        cad = cad & Space(8) '31.-
'        cad = cad & Space(15) '32.-
'        cad = cad & Space(3) '33.-
'        cad = cad & Space(2) '34.-
'        cad = cad & Space(3) '35.-
'        cad = cad & Space(8) '36.-
'        cad = cad & Space(15) '37.-
'        cad = cad & Space(3) '38.-
'        cad = cad & Space(2) '39.-
'        cad = cad & Space(3) '40.-
'        cad = cad & Space(8) '41.-
'        cad = cad & Space(15) '42.-
'        cad = cad & Space(15) '43.-
'        '44.- Importe neto
'        If DBLet(RS!impornet, "N") >= 0 Then
'            cad = cad & "+"
'        Else
'            cad = cad & "-"
'        End If
'        cad = cad & Format(DBLet(RS!impornet, "N"), "0000000000.000")
'
'        cad = cad & Space(15) '45.- Punto verde ????????
'        cad = cad & Space(17) '46.-
'        cad = cad & Space(17) '47.-
'        cad = cad & Space(17) '48.-
'        cad = cad & Space(12) '49.-
'        cad = cad & Space(12) '50.-
'        cad = cad & Space(17) '51.-
'        cad = cad & Space(3) '52.-
'        cad = cad & Space(15) '53.-
'        cad = cad & Space(17) '54.-
'        cad = cad & Space(17) '55.-
'        cad = cad & Space(15) '56.-
'        cad = cad & Space(12) '57.-
'        cad = cad & Space(35) '58.-
'        cad = cad & Space(35) '59.-
'        cad = cad & Space(15) '60.-
'        cad = cad & Space(3) '61.-
'        cad = cad & Space(12) '62.-
'        cad = cad & Space(35) '63.-
'        cad = cad & Space(3) '64.-
'
'        Print #nf, cad
'
'        RS.MoveNext
'    Wend
    
End Function

Private Sub BorrarFicheros()

    If Dir(vParamAplic.PathEdicom & "\cabfac.txt") <> "" Then BorrarArchivo vParamAplic.PathEdicom & "\cabfac.txt"
    If Dir(vParamAplic.PathEdicom & "\obsfac.txt") <> "" Then BorrarArchivo vParamAplic.PathEdicom & "\obsfac.txt"
    If Dir(vParamAplic.PathEdicom & "\dtofac.txt") <> "" Then BorrarArchivo vParamAplic.PathEdicom & "\dtofac.txt"
    If Dir(vParamAplic.PathEdicom & "\impfac.txt") <> "" Then BorrarArchivo vParamAplic.PathEdicom & "\impfac.txt"
    If Dir(vParamAplic.PathEdicom & "\contenedfac.txt") <> "" Then BorrarArchivo vParamAplic.PathEdicom & "\contenedfac.txt"
    If Dir(vParamAplic.PathEdicom & "\linfac.txt") <> "" Then BorrarArchivo vParamAplic.PathEdicom & "\linfac.txt"
    If Dir(vParamAplic.PathEdicom & "\obslfac.txt") <> "" Then BorrarArchivo vParamAplic.PathEdicom & "\obslfac.txt"
    If Dir(vParamAplic.PathEdicom & "\dtolfac.txt") <> "" Then BorrarArchivo vParamAplic.PathEdicom & "\dtolfac.txt"

End Sub


Private Function HayRegistrosEnvases(Cad As String) As Boolean
Dim SQL As String
Dim SQL1 As String

    SQL = "select codtipom, numfactu, fecfactu from facturas where " & Cad
    SQL1 = "select count(*) from facturas_envases where (codtipom, numfactu, fecfactu) = (" & SQL & ")"
    
    HayRegistrosEnvases = (RegistrosAListar(SQL1) <> 0)

End Function



Private Function ComprobarFicheros(cadwhere As String) As Boolean
Dim b As Boolean
Dim SQL As String
Dim Mens As String
        
    On Error GoTo eComprobarFicheros
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    b = True
    If b Then
        Mens = "Comprobando Registros para Cabecera de Factura"
        b = ComprobarCABFAC(cadwhere)
    End If
    
    If b Then
        Mens = "Comprobando Registros para L�neas de Factura"
        b = ComprobarLINFAC(cadwhere)
    End If
    
    ComprobarFicheros = b
    Exit Function
    
eComprobarFicheros:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobando Ficheros"
        ComprobarFicheros = False
    Else
        If Not b Then
            MuestraError Err.Number, "Error en la Comprobaci�n de Ficheros: " & vbCrLf & Mens
            ComprobarFicheros = False
        End If
    End If
End Function

Private Function ComprobarCABFAC(cadwhere As String) As Boolean
Dim b As Boolean
Dim NF As Long
Dim Cad As String
Dim SQL As String
Dim SQL1 As String
Dim i As Integer
Dim Longitud As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim vCliente As CCliente
Dim Neto As Currency
Dim Impuestos As Currency
Dim DiasVto As String
Dim Dias As Integer
Dim FecVto As Date
Dim Mens As String


    On Error GoTo eComprobarCABFAC
    
    b = True
    
    '[Monica] 29/01/2010 enlazo con facturas_variedad para no coger facturas sin lineas
    SQL = "select distinct facturas.* from facturas INNER JOIN facturas_variedad ON "
    SQL = SQL & " facturas.codtipom = facturas_variedad.codtipom "
    SQL = SQL & " and facturas.numfactu = facturas_variedad.numfactu "
    SQL = SQL & " and facturas.fecfactu = facturas_variedad.fecfactu "
    If cadwhere <> "" Then SQL = SQL & " where " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Label2.Caption = "Comprobando Registros para el Fichero CABFAC.TXT"
    Me.Refresh
    
    i = 0
    
    'parametros de empresa
    If vParamAplic.CodigoEdi = "" Then '2.-codigo edi vendedor
        Mens = "No existe codigo edi vendedor"
        SQL = "insert into tmpinformes (codusu, importe1, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(Rs!NumFactu, "N") & "," & DBSet(Mens, "T") & ")"
        conn.Execute SQL
    End If
    If vParam.NombreEmpresa = "" Then  '88.-
        Mens = "No existe nombre de empresa"
        SQL = "insert into tmpinformes (codusu, importe1, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(Rs!NumFactu, "N") & "," & DBSet(Mens, "T") & ")"
        conn.Execute SQL
    End If
    If vParam.DomicilioEmpresa = "" Then  '89.-
        Mens = "No existe domicilio de la empresa"
        SQL = "insert into tmpinformes (codusu, importe1, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(Rs!NumFactu, "N") & "," & DBSet(Mens, "T") & ")"
        conn.Execute SQL
    End If
    If vParam.Poblacion = "" Then  '90.-
        Mens = "No existe poblacion de la empresa"
        SQL = "insert into tmpinformes (codusu, importe1, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(Rs!NumFactu, "N") & "," & DBSet(Mens, "T") & ")"
        conn.Execute SQL
    End If
    If vParam.CPostal = "" Then '91.-
        Mens = "No existe codigo postal de la empresa"
        SQL = "insert into tmpinformes (codusu, importe1, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(Rs!NumFactu, "N") & "," & DBSet(Mens, "T") & ")"
        conn.Execute SQL
    End If
    If vParam.CifEmpresa = "" Then   '92.-
        Mens = "No existe cif empresa"
        SQL = "insert into tmpinformes (codusu, importe1, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(Rs!NumFactu, "N") & "," & DBSet(Mens, "T") & ")"
        conn.Execute SQL
    End If
    If vParamAplic.RegMercantil = "" Then  '93.-
        Mens = "No existe registro mercantil empresa"
        SQL = "insert into tmpinformes (codusu, importe1, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(Rs!NumFactu, "N") & "," & DBSet(Mens, "T") & ")"
        conn.Execute SQL
    End If
    'end parametros empresa
    
    
    While Not Rs.EOF
        Cad = ""
        
        Set vCliente = New CCliente
    
        'si se ha modificado el cliente volver a cargar los datos
        If vCliente.Existe(Rs!CodClien) Then
            If vCliente.LeerDatos(Rs!CodClien) Then
                SQL1 = "select albaran.*, destinos.codigoedi, destinos.nomdesti, destinos.domdesti, destinos.pobdesti, destinos.codpobla "
                SQL1 = SQL1 & " from albaran, facturas_variedad, destinos "
                SQL1 = SQL1 & " where facturas_variedad.codtipom = " & DBSet(Rs!codTipoM, "T")
                SQL1 = SQL1 & " and facturas_variedad.numfactu = " & DBSet(Rs!NumFactu, "N")
                SQL1 = SQL1 & " and facturas_variedad.fecfactu = " & DBSet(Rs!FecFactu, "F")
                SQL1 = SQL1 & " and facturas_variedad.numalbar = albaran.numalbar "
                SQL1 = SQL1 & " and albaran.codclien = destinos.codclien "
                SQL1 = SQL1 & " and albaran.coddesti = destinos.coddesti "
                
                Set Rs1 = New ADODB.Recordset
                Rs1.Open SQL1, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
                If Not Rs1.EOF Then Rs1.MoveFirst
'                While Not Rs1.EOF
                
                    If vCliente.CodigoEdi = "" Then '5.-codigo edi comprador
                        Mens = "No existe codigo edi comprador " & Rs!CodClien
                        SQL = "insert into tmpinformes (codusu, importe1, nombre1) values (" & _
                              vUsu.Codigo & "," & DBSet(Rs!NumFactu, "N") & "," & DBSet(Mens, "T") & ")"
                        conn.Execute SQL
                    End If
                    If DBLet(Rs1!CodigoEdi, "T") = "" Then  '7.-codigo edi receptor--> de la tabla de destinos
                        Mens = "No existe codigo edi del destino "
                        SQL = "insert into tmpinformes (codusu, importe1, nombre1) values (" & _
                              vUsu.Codigo & "," & DBSet(Rs!NumFactu, "N") & "," & DBSet(Mens, "T") & ")"
                        conn.Execute SQL
                    End If
                    If DBLet(Rs1!refclien, "T") = "" Then '10.-nro pedido-->albaran.refclien
                        Mens = "No existe la referencia del albaran"
                        SQL = "insert into tmpinformes (codusu, importe1, nombre1) values (" & _
                              vUsu.Codigo & "," & DBSet(Rs!NumFactu, "N") & "," & DBSet(Mens, "T") & ")"
                        conn.Execute SQL
                    End If
                    If vCliente.DestEDI = 0 Then 'destino de factura es el cliente
                        If vCliente.Nombre = "" Then       '14.-razon social del cliente
                            Mens = "No existe nombre del cliente"
                            SQL = "insert into tmpinformes (codusu, importe1, nombre1) values (" & _
                                  vUsu.Codigo & "," & DBSet(Rs!NumFactu, "N") & "," & DBSet(Mens, "T") & ")"
                            conn.Execute SQL
                        End If
                        If vCliente.Domicilio = "" Then      '15.-domicilio del cliente
                            Mens = "No existe domicilio del cliente"
                            SQL = "insert into tmpinformes (codusu, importe1, nombre1) values (" & _
                                  vUsu.Codigo & "," & DBSet(Rs!NumFactu, "N") & "," & DBSet(Mens, "T") & ")"
                            conn.Execute SQL
                        End If
                        If vCliente.Poblacion = "" Then      '16.-ciudad del cliente
                            Mens = "No existe poblacion del cliente"
                            SQL = "insert into tmpinformes (codusu, importe1, nombre1) values (" & _
                                  vUsu.Codigo & "," & DBSet(Rs!NumFactu, "N") & "," & DBSet(Mens, "T") & ")"
                            conn.Execute SQL
                        End If
                        If vCliente.CPostal = "" Then                '17.-codigo postal del cliente
                            Mens = "No existe codigo postal del cliente"
                            SQL = "insert into tmpinformes (codusu, importe1, nombre1) values (" & _
                                  vUsu.Codigo & "," & DBSet(Rs!NumFactu, "N") & "," & DBSet(Mens, "T") & ")"
                            conn.Execute SQL
                        End If
                    Else ' destinos de factura es el destino
                        If DBLet(Rs1!nomdesti, "T") = "" Then       '14.-nombre del cliente
                            Mens = "No existe nombre del destino"
                            SQL = "insert into tmpinformes (codusu, importe1, nombre1) values (" & _
                                  vUsu.Codigo & "," & DBSet(Rs!NumFactu, "N") & "," & DBSet(Mens, "T") & ")"
                            conn.Execute SQL
                        End If
                        If DBLet(Rs1!domdesti, "T") = "" Then       '15.-domicilio del cliente
                            Mens = "No existe domicilio del destino"
                            SQL = "insert into tmpinformes (codusu, importe1, nombre1) values (" & _
                                  vUsu.Codigo & "," & DBSet(Rs!NumFactu, "N") & "," & DBSet(Mens, "T") & ")"
                            conn.Execute SQL
                        End If
                        If DBLet(Rs1!pobdesti, "T") = "" Then      '16.-ciudad del cliente
                            Mens = "No existe poblacion del destino"
                            SQL = "insert into tmpinformes (codusu, importe1, nombre1) values (" & _
                                  vUsu.Codigo & "," & DBSet(Rs!NumFactu, "N") & "," & DBSet(Mens, "T") & ")"
                            conn.Execute SQL
                        End If
                        If vCliente.CPostal = "" Then                '17.-codigo postal del cliente
                            Mens = "No existe codigo postal del destino"
                            SQL = "insert into tmpinformes (codusu, importe1, nombre1) values (" & _
                                  vUsu.Codigo & "," & DBSet(Rs!NumFactu, "N") & "," & DBSet(Mens, "T") & ")"
                            conn.Execute SQL
                        End If
                    End If
                    If vCliente.NIF = "" Then          '18.-nif cliente
                        Mens = "No existe nif del cliente"
                        SQL = "insert into tmpinformes (codusu, importe1, nombre1) values (" & _
                              vUsu.Codigo & "," & DBSet(Rs!NumFactu, "N") & "," & DBSet(Mens, "T") & ")"
                        conn.Execute SQL
                    End If
                    
'                    Rs1.MoveNext
'                Wend
                Set Rs1 = Nothing
                
            End If
        End If
        
        
        Rs.MoveNext
    Wend
    
    Rs.Close
    Set Rs = Nothing
    
    Set vCliente = Nothing
    
    ComprobarCABFAC = b
    Exit Function
    
eComprobarCABFAC:
    If Err.Number <> 0 Then
        Set vCliente = Nothing
        ComprobarCABFAC = False
    End If
    
End Function






Private Function ComprobarLINFAC(cadwhere As String) As Boolean
Dim b As Boolean
Dim NF As Long
Dim Cad As String
Dim SQL As String
Dim SQL1 As String
Dim i As Integer
Dim Longitud As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim vCliente As CCliente
Dim Neto As Currency
Dim Impuestos As Currency
Dim DiasVto As String
Dim Dias As Integer
Dim FecVto As Date
Dim REFEAN As String
Dim Mens As String


    On Error GoTo eComprobarLINFAC
    
    b = True
    
    '[Monica] 29/01/2010 enlazo con facturas_variedad para no coger facturas sin lineas
    SQL = "select distinct facturas.* from facturas INNER JOIN facturas_variedad ON "
    SQL = SQL & " facturas.codtipom = facturas_variedad.codtipom "
    SQL = SQL & " and facturas.numfactu = facturas_variedad.numfactu "
    SQL = SQL & " and facturas.fecfactu = facturas_variedad.fecfactu "
    If cadwhere <> "" Then SQL = SQL & " where " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Label2.Caption = "Comprobando Registros para el Fichero CABFAC.TXT"
    Me.Refresh
    
    
    While Not Rs.EOF
        '***** INSERTAMOS LAS LINEAS DE VARIEDADES *****
        SQL = "select facturas_variedad.*, albaran_variedad.codforfait, albaran_variedad.codvarie from facturas_variedad, albaran_variedad where "
        SQL = SQL & " facturas_variedad.codtipom = " & DBSet(Rs!codTipoM, "T")
        SQL = SQL & " and facturas_variedad.numfactu = " & DBSet(Rs!NumFactu, "N")
        SQL = SQL & " and facturas_variedad.fecfactu = " & DBSet(Rs!FecFactu, "F")
        SQL = SQL & " and facturas_variedad.numalbar = albaran_variedad.numalbar "
        SQL = SQL & " and facturas_variedad.numlinealbar = albaran_variedad.numlinea "
        SQL = SQL & " order by facturas_variedad.codtipom, facturas_variedad.numfactu, "
        SQL = SQL & " facturas_variedad.fecfactu, facturas_variedad.numlinea "
        
        Set Rs1 = New ADODB.Recordset
        Rs1.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
        While Not Rs1.EOF
            REFEAN = ""
            REFEAN = DevuelveDesdeBDNew(cAgro, "codigoean", "codigoean", "codclien", Rs!CodClien, "N", , "codforfait", Rs1!codforfait, "T", "codvarie", Rs1!codvarie, "N")
            
            If REFEAN = "" Then
                Mens = "No existe referencia C" & Format(Rs!CodClien, "000000") & "-F" & Trim(Rs1!codforfait) & "-V" & Format(Rs1!codvarie, "0000")
                SQL = "insert into tmpinformes (codusu, importe1, nombre1) values (" & _
                      vUsu.Codigo & "," & DBSet(Rs1!NumFactu, "N") & "," & DBSet(Mens, "T") & ")"
                conn.Execute SQL
            End If
            Rs1.MoveNext
        Wend
        
        Set Rs1 = Nothing
        Rs.MoveNext
    Wend
    
    Rs.Close
    Set Rs = Nothing
    
    ComprobarLINFAC = b
    Exit Function
    
eComprobarLINFAC:
    If Err.Number <> 0 Then
        ComprobarLINFAC = False
    End If
End Function


