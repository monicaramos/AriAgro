VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmVtasListDif 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   7545
   Icon            =   "frmVtasListDif.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   7545
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
      Height          =   6015
      Left            =   45
      TabIndex        =   8
      Top             =   0
      Width           =   7365
      Begin VB.CheckBox Check3 
         Caption         =   "Resumido"
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
         Left            =   450
         TabIndex        =   21
         Top             =   4080
         Width           =   2355
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Incluir Valoración"
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
         Left            =   3720
         TabIndex        =   5
         Top             =   3735
         Width           =   2355
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Desglosar por Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   450
         TabIndex        =   4
         Top             =   3690
         Width           =   2355
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
         Index           =   5
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text5"
         Top             =   1920
         Width           =   4215
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
         Index           =   4
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   1500
         Width           =   4215
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
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1935
         Width           =   830
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
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   0
         Top             =   1500
         Width           =   830
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3090
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2685
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
         Left            =   5790
         TabIndex        =   7
         Top             =   5310
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
         Left            =   4605
         TabIndex        =   6
         Top             =   5310
         Width           =   1065
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   420
         TabIndex        =   19
         Top             =   4845
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Cargando tabla temporal"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   27
         Left            =   450
         TabIndex        =   20
         Top             =   5085
         Width           =   3390
      End
      Begin VB.Label Label2 
         Caption         =   "Antes era Conectando a la base de datos. Proceso muy lento .... Un momento por favor."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   555
         Left            =   450
         TabIndex        =   18
         Top             =   4860
         Width           =   4740
      End
      Begin VB.Label Label1 
         Caption         =   "Diferencias de Kilos Entrados-Salidos"
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
         TabIndex        =   17
         Top             =   450
         Width           =   5655
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1485
         MouseIcon       =   "frmVtasListDif.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1485
         MouseIcon       =   "frmVtasListDif.frx":015E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   1500
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
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
         TabIndex        =   16
         Top             =   1125
         Width           =   855
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
         Left            =   735
         TabIndex        =   15
         Top             =   1920
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
         Left            =   735
         TabIndex        =   14
         Top             =   1500
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
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
         Left            =   510
         TabIndex        =   11
         Top             =   2340
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
         Left            =   735
         TabIndex        =   10
         Top             =   2685
         Width           =   645
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
         Left            =   735
         TabIndex        =   9
         Top             =   3090
         Width           =   600
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1485
         Picture         =   "frmVtasListDif.frx":02B0
         ToolTipText     =   "Buscar fecha"
         Top             =   2685
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1485
         Picture         =   "frmVtasListDif.frx":033B
         ToolTipText     =   "Buscar fecha"
         Top             =   3090
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmVtasListDif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MANOLO +-+-
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
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmMensVariedad As frmMensajes 'mensajes
Attribute frmMensVariedad.VB_VarHelpID = -1

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
Dim Variedades As String


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Check2_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Check3_Click()
    If Check3.Value = 1 Then
        Check2.Value = 0
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
Dim vsqlVariedad As String
InicializarVbles
    
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H Variedad
    cDesde = Trim(txtCodigo(4).Text)
    cHasta = Trim(txtCodigo(5).Text)
    nDesde = txtNombre(4).Text
    nHasta = txtNombre(5).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{tmpinfventas.numalbar}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad= """) Then Exit Sub
    End If
    
    '[Monica]15/12/2014: Añadido el que puedan seleccionar mas de una variedad
    
    Variedades = ""
    
    vsqlVariedad = ""
    If txtCodigo(4).Text <> "" Then vsqlVariedad = vsqlVariedad & " and variedades.codvarie >= " & DBSet(txtCodigo(4).Text, "N")
    If txtCodigo(5).Text <> "" Then vsqlVariedad = vsqlVariedad & " and variedades.codvarie <= " & DBSet(txtCodigo(5).Text, "N")
    
    
    If vsqlVariedad <> "" And txtCodigo(4).Text <> txtCodigo(5).Text Then
        Set frmMensVariedad = New frmMensajes
    
        frmMensVariedad.OpcionMensaje = 21
        frmMensVariedad.Label5 = "Variedades"
        frmMensVariedad.cadwhere = vsqlVariedad
        frmMensVariedad.Show vbModal
    
        Set frmMensVariedad = Nothing
    End If
    
    'D/H Fecha
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{tmpinfventas.fecalbar}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    cadTABLA = tabla & " INNER JOIN albaran_variedad ON albaran.numalbar = albaran_variedad.numalbar "
    
    ' si desglosamos el informe o no por fechas
    cadParam = cadParam & "pFecha=" & Check1.Value & "|"
    numParam = numParam + 1
    
    If Not AnyadirAFormula(cadFormula, "{tmpinfventas.codusu} = " & vUsu.Codigo) Then Exit Sub
    
    If Check2.Value = 1 Then
        If CargarTablaTemporalValorado Then
            If HayRegParaInforme("tmpinfventas", "codusu=" & vUsu.Codigo & " and " & cadselect) Then
                cadTitulo = "Diferencias entre kilos entrados-salidos Valorado"
                cadNombreRPT = "rVtasListDifVal.rpt"
                LlamarImprimir
            End If
        End If
    Else
        If CargarTablaTemporal Then
            If HayRegParaInforme("tmpinfventas", "codusu=" & vUsu.Codigo & " and " & cadselect) Then
                cadTitulo = "Diferencias entre kilos entrados-salidos"
                If Check3.Value = 1 Then
                    cadNombreRPT = "rVtasListDifRes.rpt"
                Else
                    cadNombreRPT = "rVtasListDif.rpt"
                End If
                LlamarImprimir
            End If
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(4)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me
    
    Label2.visible = False
    Me.Refresh
    DoEvents

   'IMAGES para busqueda
    Me.imgBuscar(4).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Me.imgBuscar(5).Picture = frmPpal.imgListImages16.ListImages(1).Picture

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, H, W
    indFrame = 5
    tabla = "albaran"
    
    Check2.visible = (vParamAplic.Cooperativa = 9)
    Check2.Enabled = (vParamAplic.Cooperativa = 9)
    
    '[Monica]15/12/2014: solo para el caso de Picassent
    Check3.visible = (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16)
    Check3.Enabled = (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16)
    
    Label4(27).visible = False
    Me.Pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CerrarConexionMultibase
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmMensVariedad_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        Sql = " {albaran_variedad.codvarie} in (" & CadenaSeleccion & ")"
        Variedades = " in (" & CadenaSeleccion & ")"
    Else
        Variedades = " in (-1) "
    End If
'    If Not AnyadirAFormula(cadselect, Sql) Then Exit Sub
'    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Variedades
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
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
    If txtCodigo(Index).Text <> "" Then frmC.NovaData = txtCodigo(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtCodigo(CByte(imgFec(2).Tag))
    ' ***************************
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 4, 5 'VARIEDADES
            AbrirFrmVariedades (Index)
        
    End Select
    PonerFoco txtCodigo(indCodigo)
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
            Case 4: KEYBusqueda KeyAscii, 4 'variedad desde
            Case 5: KEYBusqueda KeyAscii, 5 'variedad hasta
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
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
        Case 2, 3 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 4, 5 'VARIEDAD
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
            
                        
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 6015
        Me.FrameCobros.Width = 7365
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
        .EnvioEMail = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .Opcion = 0
        .Show vbModal
    End With
End Sub

Private Sub AbrirFrmVariedades(indice As Integer)
    indCodigo = indice
    Set frmVar = New frmManVariedad
    frmVar.DatosADevolverBusqueda = "0|1|"
    frmVar.DeConsulta = True
    frmVar.CodigoActual = txtCodigo(indCodigo)
    frmVar.Show vbModal
    Set frmVar = Nothing
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

Private Function CargarTablaTemporal() As Boolean
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim Sql As String
Dim SQL1 As String
Dim Sql2 As String
Dim Calidad As Integer
Dim Destrio As Long

    On Error GoTo eCargarTablaTemporal

    CargarTablaTemporal = False
    
    ' borramos los registros de la tabla temporal
    Sql = "delete from tmpinfventas where codusu = " & DBSet(vUsu.Codigo, "N")
    conn.Execute Sql
    
    
    Screen.MousePointer = vbHourglass
     
'    Label2.visible = True
    Me.Refresh
    DoEvents


'26-05-2009: quitamos la conexion a multibase, ahora miramos la rhisfruta
'    If Not AbrirConexionMultibase("mb", "mb") Then
'        Screen.MousePointer = vbDefault
'        Exit Function
'    End If
'
'    SQL = "CREATE TEMPORARY TABLE tmp ( "
'    SQL = SQL & "codigo1 SMALLINT(3) UNSIGNED  DEFAULT '0' NOT NULL, "
'    SQL = SQL & "codigo2 smallint(3) UNSIGNED  DEFAULT '0' NOT NULL, "
'    SQL = SQL & "fecalbar Date, "
'    SQL = SQL & "neto1 int(7), "
'    SQL = SQL & "neto3 int(7))" ' columna de destrio
'
'    conn.Execute SQL
'
'    SQL = "select a.codprodu , a.codvarie , a.fecalbar , a.kilosnet, "
'    SQL = SQL & " a.kilosn01, a.kilosn02, a.kilosn03, a.kilosn04, a.kilosn05, a.kilosn06, "
'    SQL = SQL & " a.kilosn07, a.kilosn08, a.kilosn09, a.kilosn10, a.kilosn11, a.kilosn12, "
'    SQL = SQL & " a.kilosn13, a.kilosn14,  a.kilosn15, a.kilosn16, a.kilosn17, a.kilosn18, "
'    SQL = SQL & " a.kilosn19, a.kilosn20 "
'    SQL = SQL & " from shifru a where 1=1 "
''    SQL = "select a.codprodu , a.codvarie , a.fecalbar , sum(a.kilosnet) "
''    SQL = SQL & " from shifru a where 1=1 "
'
'    If txtCodigo(2).Text <> "" Then SQL = SQL & " and a.fecalbar >= """ & Trim(txtCodigo(2).Text) & """"
'    If txtCodigo(3).Text <> "" Then SQL = SQL & " and a.fecalbar <= """ & Trim(txtCodigo(3).Text) & """"
'
'    If txtCodigo(4).Text <> "" Then
'        SQL = SQL & " and a.codprodu >= " & DBSet(Mid(txtCodigo(4).Text, 3, 2), "N")
''27/10/2008: quito la variedad pq pueden ponerme productos diferentes
''        SQL = SQL & " and a.codvarie >= " & DBSet(Mid(txtCodigo(4).Text, 5, 2), "N")
'    End If
'    If txtCodigo(5).Text <> "" Then
'        SQL = SQL & " and a.codprodu <= " & DBSet(Mid(txtCodigo(5).Text, 3, 2), "N")
''27/10/2008: quito la variedad pq pueden ponerme productos diferentes
''        SQL = SQL & " and a.codvarie <= " & DBSet(Mid(txtCodigo(5).Text, 5, 2), "N")
'    End If
'    If txtCodigo(4).Text <> "" And txtCodigo(5).Text <> "" Then
'        If Mid(txtCodigo(4).Text, 3, 2) = Mid(txtCodigo(5).Text, 3, 2) Then
'            SQL = SQL & " and a.codvarie >= " & DBSet(Mid(txtCodigo(4).Text, 5, 2), "N")
'            SQL = SQL & " and a.codvarie <= " & DBSet(Mid(txtCodigo(5).Text, 5, 2), "N")
'        End If
'    End If
'
''    SQL = SQL & " group by 1, 2, 3"
''    SQL = SQL & " order by 1, 2, 3"
'
'
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, ConnMB, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    Sql1 = ""
'    While Not RS.EOF
'        '++que calidad es de destrio
'        Sql2 = "select codcalir from scalir where codprodu = " & DBLet(RS.Fields(0).Value, "N")
'        Sql2 = Sql2 & " and codvarie = " & DBLet(RS.Fields(1).Value, "N") & " and esdestri = ""S"""
'
'        Set Rs1 = New ADODB.Recordset
'        Rs1.Open Sql2, ConnMB, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'        If Not Rs1.EOF Then
'            Calidad = DBLet(Rs1.Fields(0).Value, "N")
'        End If
'
'        Sql1 = Sql1 & "(" ' & vUsu.Codigo & ","
'        Sql1 = Sql1 & DBSet(RS.Fields(0).Value, "N") & "," & DBSet(RS.Fields(1).Value, "N") & ","
'        Sql1 = Sql1 & DBSet(RS.Fields(3).Value, "N") & "," & DBSet(RS.Fields(2), "F") & ","
'        Sql1 = Sql1 & DBSet(RS.Fields(3 + Calidad).Value, "N") & "),"
'
'
'
'        RS.MoveNext
'    Wend
'
'    Set RS = Nothing
'
'    CerrarConexionMultibase
'
'    If Sql1 <> "" Then
'        ' quitamos la ultima coma
'        Sql1 = Mid(Sql1, 1, Len(Sql1) - 1)
'
'        ' insertamos los registros en la tabla temporal
'        SQL = "insert into tmp (codigo1, codigo2, neto1, fecalbar, neto3) values "
'        SQL = SQL & Sql1
'
'        conn.Execute SQL
'
''        SQL = "insert into tmpinfventas (codusu, codigo1, codigo2, neto1, fecalbar) values "
''        SQL = SQL & Sql1
''
''        Conn.Execute SQL
'    End If
'
'    SQL = "insert into tmpinfventas (codusu, numalbar, fecalbar, neto1, neto3) "
'    SQL = SQL & " select " & vUsu.Codigo & ", concat(right(concat('00',codigo1),2),right(concat('00',codigo2),2) ), fecalbar, sum(neto1), sum(neto3) "
'    SQL = SQL & " from tmp "
'    SQL = SQL & " group by 1,2,3 "
'    SQL = SQL & " order by 1,2,3 "
'
'    conn.Execute SQL
'
'    'Borrar la tabla temporal
'    SQL = " DROP TABLE IF EXISTS tmp;"
'    conn.Execute SQL
'
'
'
''    ' ponemos en numalbar el codigo de la variedad, concatenando codigo1 y codigo2 del MB.
''    SQL = "update tmpinfventas set numalbar = concat(right(concat('00',codigo1),2),right(concat('00',codigo2),2) ) "
''    SQL = SQL & " where codusu = " & vUsu.Codigo
''
''    Conn.Execute SQL
    
    
    ' nuevo: los kilos de entrada de la rhisfruta en lugar de MB
    Sql = "select rhisfruta.codvarie,rhisfruta.fecalbar, sum(rhisfruta.kilosnet)"
    Sql = Sql & " from rhisfruta where 1=1 "
    If txtCodigo(2).Text <> "" Then Sql = Sql & " and rhisfruta.fecalbar >= " & DBSet(txtCodigo(2).Text, "F")
    If txtCodigo(3).Text <> "" Then Sql = Sql & " and rhisfruta.fecalbar <= " & DBSet(txtCodigo(3).Text, "F")
    
    If txtCodigo(4).Text <> "" Then Sql = Sql & " and rhisfruta.codvarie >= " & DBSet(txtCodigo(4).Text, "N")
    If txtCodigo(5).Text <> "" Then Sql = Sql & " and rhisfruta.codvarie <= " & DBSet(txtCodigo(5).Text, "N")
    
    '[Monica]15/12/2014: punteo de variedades
    If Variedades <> "" Then
        Sql = Sql & " and rhisfruta.codvarie " & Variedades
    End If
    
    Sql = Sql & " group by 1,2 "
    Sql = Sql & " order by 1,2 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SQL1 = ""
    While Not Rs.EOF
        SQL1 = "select count(*) from tmpinfventas where codusu = " & vUsu.Codigo & " and "
        SQL1 = SQL1 & " numalbar = " & DBSet(Rs.Fields(0).Value, "N") & " and "
        SQL1 = SQL1 & " fecalbar = " & DBSet(Rs.Fields(1).Value, "F")
        
        If TotalRegistros(SQL1) = 0 Then
            ' si no existe registro lo insertamos
            SQL1 = "insert into tmpinfventas (codusu, numalbar, fecalbar, neto1) values ("
            SQL1 = SQL1 & DBSet(vUsu.Codigo, "N") & "," & DBSet(Rs.Fields(0).Value, "N") & ","
            SQL1 = SQL1 & DBSet(Rs.Fields(1).Value, "F") & "," & DBSet(Rs.Fields(2).Value, "N") & ")"
        Else
            ' si existe el registro lo updateamos
            SQL1 = "update tmpinfventas set neto1 = " & DBSet(Rs.Fields(2).Value, "N")
            SQL1 = SQL1 & " where numalbar = " & DBSet(Rs.Fields(0).Value, "N")
            SQL1 = SQL1 & " and fecalbar = " & DBSet(Rs.Fields(1).Value, "F")
            SQL1 = SQL1 & " and codusu = " & DBSet(vUsu.Codigo, "N")
        End If
        
        conn.Execute SQL1
        
        ' para esta variedad, fecha cuales son los kilos de destrio
        SQL1 = "select sum(rhisfruta_clasif.kilosnet) from rhisfruta_clasif inner join rhisfruta on rhisfruta_clasif.numalbar = rhisfruta.numalbar "
        SQL1 = SQL1 & " where rhisfruta_clasif.codvarie = " & DBSet(Rs.Fields(0).Value, "N")
        SQL1 = SQL1 & " and rhisfruta.fecalbar = " & DBSet(Rs.Fields(1).Value, "F") & " and codcalid in (select codcalid "
        SQL1 = SQL1 & " from rcalidad where tipcalid = 1 and codvarie = " & DBSet(Rs.Fields(0).Value, "N") & ")"
        
        Destrio = DevuelveValor(SQL1)
    
        SQL1 = "update tmpinfventas set neto3 = " & DBSet(Destrio, "N")
        SQL1 = SQL1 & " where numalbar = " & DBSet(Rs.Fields(0).Value, "N")
        SQL1 = SQL1 & " and fecalbar = " & DBSet(Rs.Fields(1).Value, "F")
        SQL1 = SQL1 & " and codusu = " & DBSet(vUsu.Codigo, "N")
    
        conn.Execute SQL1
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    '[Monica]15/12/2014: miramos si hay algo en entradas y en la clasificacion
    If Check3.Value = 1 Then
        Sql = "select rentradas.codvarie, rentradas.fechaent, sum(rentradas.kilosnet)"
        Sql = Sql & " from rentradas where 1=1 "
        If txtCodigo(2).Text <> "" Then Sql = Sql & " and rentradas.fechaent >= " & DBSet(txtCodigo(2).Text, "F")
        If txtCodigo(3).Text <> "" Then Sql = Sql & " and rentradas.fechaent <= " & DBSet(txtCodigo(3).Text, "F")
        
        If txtCodigo(4).Text <> "" Then Sql = Sql & " and rentradas.codvarie >= " & DBSet(txtCodigo(4).Text, "N")
        If txtCodigo(5).Text <> "" Then Sql = Sql & " and rentradas.codvarie <= " & DBSet(txtCodigo(5).Text, "N")
        
        '[Monica]15/12/2014: punteo de variedades
        If Variedades <> "" Then
            Sql = Sql & " and rentradas.codvarie " & Variedades
        End If
        
        Sql = Sql & " group by 1,2 "
        Sql = Sql & " order by 1,2 "
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        SQL1 = ""
        While Not Rs.EOF
            SQL1 = "select count(*) from tmpinfventas where codusu = " & vUsu.Codigo & " and "
            SQL1 = SQL1 & " numalbar = " & DBSet(Rs.Fields(0).Value, "N") & " and "
            SQL1 = SQL1 & " fecalbar = " & DBSet(Rs.Fields(1).Value, "F")
            
            If TotalRegistros(SQL1) = 0 Then
                ' si no existe registro lo insertamos
                SQL1 = "insert into tmpinfventas (codusu, numalbar, fecalbar, neto1) values ("
                SQL1 = SQL1 & DBSet(vUsu.Codigo, "N") & "," & DBSet(Rs.Fields(0).Value, "N") & ","
                SQL1 = SQL1 & DBSet(Rs.Fields(1).Value, "F") & "," & DBSet(Rs.Fields(2).Value, "N") & ")"
            Else
                ' si existe el registro lo updateamos
                SQL1 = "update tmpinfventas set neto1 = " & DBSet(Rs.Fields(2).Value, "N")
                SQL1 = SQL1 & " where numalbar = " & DBSet(Rs.Fields(0).Value, "N")
                SQL1 = SQL1 & " and fecalbar = " & DBSet(Rs.Fields(1).Value, "F")
                SQL1 = SQL1 & " and codusu = " & DBSet(vUsu.Codigo, "N")
            End If
            
            conn.Execute SQL1
            
            Rs.MoveNext
        Wend
        
        Set Rs = Nothing
        
        
        Sql = "select rclasifica.codvarie,rclasifica.fechaent, sum(rclasifica.kilosnet)"
        Sql = Sql & " from rclasifica where 1=1 "
        If txtCodigo(2).Text <> "" Then Sql = Sql & " and rclasifica.fechaent >= " & DBSet(txtCodigo(2).Text, "F")
        If txtCodigo(3).Text <> "" Then Sql = Sql & " and rclasifica.fechaent <= " & DBSet(txtCodigo(3).Text, "F")
        
        If txtCodigo(4).Text <> "" Then Sql = Sql & " and rclasifica.codvarie >= " & DBSet(txtCodigo(4).Text, "N")
        If txtCodigo(5).Text <> "" Then Sql = Sql & " and rclasifica.codvarie <= " & DBSet(txtCodigo(5).Text, "N")
        
        '[Monica]15/12/2014: punteo de variedades
        If Variedades <> "" Then
            Sql = Sql & " and rclasifica.codvarie " & Variedades
        End If
        
        Sql = Sql & " group by 1,2 "
        Sql = Sql & " order by 1,2 "
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        SQL1 = ""
        While Not Rs.EOF
            SQL1 = "select count(*) from tmpinfventas where codusu = " & vUsu.Codigo & " and "
            SQL1 = SQL1 & " numalbar = " & DBSet(Rs.Fields(0).Value, "N") & " and "
            SQL1 = SQL1 & " fecalbar = " & DBSet(Rs.Fields(1).Value, "F")
            
            If TotalRegistros(SQL1) = 0 Then
                ' si no existe registro lo insertamos
                SQL1 = "insert into tmpinfventas (codusu, numalbar, fecalbar, neto1) values ("
                SQL1 = SQL1 & DBSet(vUsu.Codigo, "N") & "," & DBSet(Rs.Fields(0).Value, "N") & ","
                SQL1 = SQL1 & DBSet(Rs.Fields(1).Value, "F") & "," & DBSet(Rs.Fields(2).Value, "N") & ")"
            Else
                ' si existe el registro lo updateamos
                SQL1 = "update tmpinfventas set neto1 = " & DBSet(Rs.Fields(2).Value, "N")
                SQL1 = SQL1 & " where numalbar = " & DBSet(Rs.Fields(0).Value, "N")
                SQL1 = SQL1 & " and fecalbar = " & DBSet(Rs.Fields(1).Value, "F")
                SQL1 = SQL1 & " and codusu = " & DBSet(vUsu.Codigo, "N")
            End If
            
            conn.Execute SQL1
            
            Rs.MoveNext
        Wend
        
        Set Rs = Nothing
    End If
    
    
    
    ' metemos los kilos de salida de la tabla albaran_variedad
    Sql = "select albaran_variedad.codvarie, albaran.fechaalb, sum(albaran_variedad.pesoneto)"
    Sql = Sql & " from albaran inner join albaran_variedad on albaran.numalbar = albaran_variedad.numalbar "
    
    If txtCodigo(2).Text <> "" Then Sql = Sql & " and albaran.fechaalb >= " & DBSet(txtCodigo(2).Text, "F")
    If txtCodigo(3).Text <> "" Then Sql = Sql & " and albaran.fechaalb <= " & DBSet(txtCodigo(3).Text, "F")
    
    If txtCodigo(4).Text <> "" Then Sql = Sql & " and albaran_variedad.codvarie >= " & DBSet(txtCodigo(4).Text, "N")
    If txtCodigo(5).Text <> "" Then Sql = Sql & " and albaran_variedad.codvarie <= " & DBSet(txtCodigo(5).Text, "N")
    
    '[Monica]15/12/2014: punteo de variedades
    If Variedades <> "" Then
        Sql = Sql & " and albaran_variedad.codvarie " & Variedades
    End If
    
    Sql = Sql & " group by 1,2 "
    Sql = Sql & " order by 1,2 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SQL1 = ""
    While Not Rs.EOF
        SQL1 = "select count(*) from tmpinfventas where codusu = " & vUsu.Codigo & " and "
        SQL1 = SQL1 & " numalbar = " & DBSet(Rs.Fields(0).Value, "N") & " and "
        SQL1 = SQL1 & " fecalbar = " & DBSet(Rs.Fields(1).Value, "F")
        
        If TotalRegistros(SQL1) = 0 Then
            ' si no existe registro lo insertamos
            SQL1 = "insert into tmpinfventas (codusu, numalbar, fecalbar, neto2) values ("
            SQL1 = SQL1 & DBSet(vUsu.Codigo, "N") & "," & DBSet(Rs.Fields(0).Value, "N") & ","
            SQL1 = SQL1 & DBSet(Rs.Fields(1).Value, "F") & "," & DBSet(Rs.Fields(2).Value, "N") & ")"
        Else
            ' si existe el registro lo updateamos
            SQL1 = "update tmpinfventas set neto2 = " & DBSet(Rs.Fields(2).Value, "N")
            SQL1 = SQL1 & " where numalbar = " & DBSet(Rs.Fields(0).Value, "N")
            SQL1 = SQL1 & " and fecalbar = " & DBSet(Rs.Fields(1).Value, "F")
        End If
        
        conn.Execute SQL1
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    CargarTablaTemporal = True
    
    Label2.visible = False
    Me.Refresh
    DoEvents
    
    Screen.MousePointer = vbDefault
    
    Exit Function
    
eCargarTablaTemporal:
    MuestraError Err.Number, "Cargando tabla temporal" & " " & Err.Description
    
    CerrarConexionMultibase
    Screen.MousePointer = vbDefault
    Label2.visible = False
    Me.Refresh
    DoEvents
    
    'Borrar la tabla temporal
    Sql = " DROP TABLE IF EXISTS tmp;"
    conn.Execute Sql
    
End Function




Private Function CargarTablaTemporalValorado() As Boolean
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim Sql As String
Dim SQL1 As String
Dim Sql2 As String
Dim Calidad As Integer
Dim Destrio As Long
Dim ImporteFacturado As Currency

Dim Nregs As Long


    On Error GoTo eCargarTablaTemporalValorado

    CargarTablaTemporalValorado = False
    
    ' borramos los registros de la tabla temporal
    Sql = "delete from tmpinfventas where codusu = " & DBSet(vUsu.Codigo, "N")
    conn.Execute Sql
    
    Screen.MousePointer = vbHourglass
     
'    Label2.visible = True
    Me.Refresh
    DoEvents


    Sql = "select rhisfruta.numalbar, rhisfruta.codvarie, rhisfruta.fecalbar, rhisfruta.kilosnet, rhisfruta.prestimado "
    Sql = Sql & " from rhisfruta where 1=1 "
    If txtCodigo(2).Text <> "" Then Sql = Sql & " and rhisfruta.fecalbar >= " & DBSet(txtCodigo(2).Text, "F")
    If txtCodigo(3).Text <> "" Then Sql = Sql & " and rhisfruta.fecalbar <= " & DBSet(txtCodigo(3).Text, "F")
    
    If txtCodigo(4).Text <> "" Then Sql = Sql & " and rhisfruta.codvarie >= " & DBSet(txtCodigo(4).Text, "N")
    If txtCodigo(5).Text <> "" Then Sql = Sql & " and rhisfruta.codvarie <= " & DBSet(txtCodigo(5).Text, "N")
    
    '[Monica]15/12/2014: punteo de variedades
    If Variedades <> "" Then
        Sql = Sql & " and rhisfruta.codvarie " & Variedades
    End If
    
    Nregs = TotalRegistrosConsulta(Sql)
    
    Sql = "select albaran_variedad.numalbar, albaran_variedad.numlinea, albaran_variedad.codvarie, albaran.fechaalb, albaran_variedad.pesoneto, albaran_variedad.preciopro, albaran_variedad.preciodef "
    Sql = Sql & " from albaran inner join albaran_variedad on albaran.numalbar = albaran_variedad.numalbar "
    
    If txtCodigo(2).Text <> "" Then Sql = Sql & " and albaran.fechaalb >= " & DBSet(txtCodigo(2).Text, "F")
    If txtCodigo(3).Text <> "" Then Sql = Sql & " and albaran.fechaalb <= " & DBSet(txtCodigo(3).Text, "F")
    
    If txtCodigo(4).Text <> "" Then Sql = Sql & " and albaran_variedad.codvarie >= " & DBSet(txtCodigo(4).Text, "N")
    If txtCodigo(5).Text <> "" Then Sql = Sql & " and albaran_variedad.codvarie <= " & DBSet(txtCodigo(5).Text, "N")
    
    '[Monica]15/12/2014: punteo de variedades
    If Variedades <> "" Then
        Sql = Sql & " and albaran_variedad.codvarie " & Variedades
    End If
    
    
    Nregs = Nregs + TotalRegistrosConsulta(Sql)



    Label4(27).visible = True
    Pb1.visible = True
    Label4(27).Caption = "Cargando tabla temporal: entradas "
    DoEvents
    Pb1.Max = Nregs
    Pb1.Value = 0
    


    ' DATOS DE ENTRADA
    ' kilos de entrada de la rhisfruta y lo que se ha pagado
    Sql = "select rhisfruta.numalbar, rhisfruta.codvarie, rhisfruta.fecalbar, rhisfruta.kilosnet, rhisfruta.prestimado "
    Sql = Sql & " from rhisfruta where 1=1 "
    If txtCodigo(2).Text <> "" Then Sql = Sql & " and rhisfruta.fecalbar >= " & DBSet(txtCodigo(2).Text, "F")
    If txtCodigo(3).Text <> "" Then Sql = Sql & " and rhisfruta.fecalbar <= " & DBSet(txtCodigo(3).Text, "F")
    
    If txtCodigo(4).Text <> "" Then Sql = Sql & " and rhisfruta.codvarie >= " & DBSet(txtCodigo(4).Text, "N")
    If txtCodigo(5).Text <> "" Then Sql = Sql & " and rhisfruta.codvarie <= " & DBSet(txtCodigo(5).Text, "N")
    
    '[Monica]15/12/2014: punteo de
    If Variedades <> "" Then
        Sql = Sql & " and rhisfruta.codvarie " & Variedades
    End If
    Sql = Sql & " order by 1,2 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SQL1 = ""
    While Not Rs.EOF
        IncrementarProgresNew Pb1, 1
    
        ImporteFacturado = ImporteAlbaranFacturadoaSocio(CStr(Rs!NumAlbar), CStr(Rs!codvarie))
        
        If ImporteFacturado = 0 Then
            ImporteFacturado = Round2(DBLet(Rs!KilosNet, "N") * DBLet(Rs!prestimado, "N"), 2)
        End If
    
        SQL1 = "select count(*) from tmpinfventas where codusu = " & vUsu.Codigo & " and "
        SQL1 = SQL1 & " numalbar = " & DBSet(Rs.Fields(1).Value, "N") & " and "
        SQL1 = SQL1 & " fecalbar = " & DBSet(Rs.Fields(2).Value, "F")
        
        If TotalRegistros(SQL1) = 0 Then
            ' si no existe registro lo insertamos
            SQL1 = "insert into tmpinfventas (codusu, numalbar, fecalbar, neto1, gastos1, neto2, gastos2) values ("
            SQL1 = SQL1 & DBSet(vUsu.Codigo, "N") & "," & DBSet(Rs.Fields(1).Value, "N") & ","
            SQL1 = SQL1 & DBSet(Rs.Fields(2).Value, "F") & "," & DBSet(Rs.Fields(3).Value, "N") & ","
            SQL1 = SQL1 & DBSet(ImporteFacturado, "N") & ",0,0)"
        Else
            ' si existe el registro lo updateamos
            SQL1 = "update tmpinfventas set neto1 = neto1 + " & DBSet(Rs.Fields(3).Value, "N")
            SQL1 = SQL1 & ", gastos1 = gastos1 + " & DBSet(ImporteFacturado, "N")
            SQL1 = SQL1 & " where numalbar = " & DBSet(Rs.Fields(1).Value, "N")
            SQL1 = SQL1 & " and fecalbar = " & DBSet(Rs.Fields(2).Value, "F")
            SQL1 = SQL1 & " and codusu = " & DBSet(vUsu.Codigo, "N")
        End If
        
        conn.Execute SQL1
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    
    ' DATOS DE VENTA
    ' kilos de salida de albaranes de venta y lo que se ha facturado
    
    Label4(27).Caption = "Cargando tabla temporal: salidas "
    DoEvents
    
    
    ' metemos los kilos de salida de la tabla albaran_variedad
    Sql = "select albaran_variedad.numalbar, albaran_variedad.numlinea, albaran_variedad.codvarie, albaran.fechaalb, albaran_variedad.pesoneto, albaran_variedad.preciopro, albaran_variedad.preciodef "
    Sql = Sql & " from albaran inner join albaran_variedad on albaran.numalbar = albaran_variedad.numalbar "
    
    If txtCodigo(2).Text <> "" Then Sql = Sql & " and albaran.fechaalb >= " & DBSet(txtCodigo(2).Text, "F")
    If txtCodigo(3).Text <> "" Then Sql = Sql & " and albaran.fechaalb <= " & DBSet(txtCodigo(3).Text, "F")
    
    If txtCodigo(4).Text <> "" Then Sql = Sql & " and albaran_variedad.codvarie >= " & DBSet(txtCodigo(4).Text, "N")
    If txtCodigo(5).Text <> "" Then Sql = Sql & " and albaran_variedad.codvarie <= " & DBSet(txtCodigo(5).Text, "N")
    
    '[Monica]15/12/2014: punteo de variedades
    If Variedades <> "" Then
        Sql = Sql & " and albaran_variedad.codvarie " & Variedades
    End If
    
    Sql = Sql & " order by 1,2 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SQL1 = ""
    While Not Rs.EOF
        IncrementarProgresNew Pb1, 1
    
        Sql2 = "select sum(facturas_variedad.impornet) from facturas_variedad where numalbar = " & DBSet(Rs!NumAlbar, "N")
        Sql2 = Sql2 & " and numlinealbar = " & DBSet(Rs!NumLinea, "N")
        
        ImporteFacturado = DevuelveValor(Sql2)
        
        If ImporteFacturado = 0 Then
            If DBLet(Rs!preciodef, "N") = 0 Then
                ImporteFacturado = Round2(DBLet(Rs!Pesoneto, "N") * DBLet(Rs!preciopro, "N"), 2)
            Else
                ImporteFacturado = Round2(DBLet(Rs!Pesoneto, "N") * DBLet(Rs!preciodef, "N"), 2)
            End If
        End If
        
        '[Monica]12/06/2013: al importe facturado he de eliminarle los gastos
        ImporteFacturado = ImporteFacturado - CalcularGastos(Rs!NumAlbar, Rs!NumLinea)
        
        
        SQL1 = "select count(*) from tmpinfventas where codusu = " & vUsu.Codigo & " and "
        SQL1 = SQL1 & " numalbar = " & DBSet(Rs.Fields(2).Value, "N") & " and "
        SQL1 = SQL1 & " fecalbar = " & DBSet(Rs.Fields(3).Value, "F")
        
        If TotalRegistros(SQL1) = 0 Then
            ' si no existe registro lo insertamos
            SQL1 = "insert into tmpinfventas (codusu, numalbar, fecalbar, neto2, gastos2, neto1, gastos1) values ("
            SQL1 = SQL1 & DBSet(vUsu.Codigo, "N") & "," & DBSet(Rs.Fields(2).Value, "N") & ","
            SQL1 = SQL1 & DBSet(Rs.Fields(3).Value, "F") & "," & DBSet(Rs.Fields(4).Value, "N") & ","
            SQL1 = SQL1 & DBSet(ImporteFacturado, "N") & ",0,0)"
        Else
            ' si existe el registro lo updateamos
            SQL1 = "update tmpinfventas set neto2 = neto2 + " & DBSet(Rs.Fields(4).Value, "N")
            SQL1 = SQL1 & ", gastos2 = gastos2 + " & DBSet(ImporteFacturado, "N")
            SQL1 = SQL1 & " where numalbar = " & DBSet(Rs.Fields(2).Value, "N")
            SQL1 = SQL1 & " and fecalbar = " & DBSet(Rs.Fields(3).Value, "F")
            SQL1 = SQL1 & " and codusu = " & DBSet(vUsu.Codigo, "N")
        End If
        
        conn.Execute SQL1
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    CargarTablaTemporalValorado = True
    
    Label2.visible = False
    Me.Refresh
    DoEvents
    
    Screen.MousePointer = vbDefault
    
    Exit Function
    
eCargarTablaTemporalValorado:
    MuestraError Err.Number, "Cargando tabla temporal" & " " & Err.Description
    
    Screen.MousePointer = vbDefault
    Label2.visible = False
    Me.Refresh
    DoEvents
    
    'Borrar la tabla temporal
    Sql = " DROP TABLE IF EXISTS tmp;"
    conn.Execute Sql
    
End Function



'Private Function ImporteAlbaranFacturadoaSocio(numalbar As String) As Currency
'Dim Sql As String
'Dim Importe As Currency
'Dim importe2 As Currency
'
'    Sql = "select sum(importel) from rlifter where numalbar = " & DBSet(numalbar, "N")
'    Importe = DevuelveValor(Sql)
'
'    Sql = "select sum(if(importe is null,0,importe) - if(imporgasto is null,0,imporgasto)) from rfactsoc_albaran where numalbar = " & DBSet(numalbar, "N") & " and codtipom in (select codtipom from usuarios.stipom where tipodocu = 2)"
'    Sql = Sql & " and not (rfactsoc_albaran.codtipom, rfactsoc_albaran.numfactu, rfactsoc_albaran.fecfactu) in (select rectif_codtipom,rectif_numfactu,rectif_fecfactu from rfactsoc "
'    Sql = Sql & " where not rectif_codtipom is null and not numfactu is null and not rectif_fecfactu is null)"
'
'    importe2 = DevuelveValor(Sql)
'
'    '[Monica]06/11/2013: si no esta liquidado cogemos todo lo que haya anticipado
'    If Not AlbaranLiquidado(numalbar) Then
'        Sql = "select sum(if(importe is null,0,importe) - if(imporgasto is null,0,imporgasto)) from rfactsoc_albaran where numalbar = " & DBSet(numalbar, "N") & " and codtipom in (select codtipom from usuarios.stipom where tipodocu = 1)"
'        Sql = Sql & " and not (rfactsoc_albaran.codtipom, rfactsoc_albaran.numfactu, rfactsoc_albaran.fecfactu) in (select rectif_codtipom,rectif_numfactu,rectif_fecfactu from rfactsoc "
'        Sql = Sql & " where not rectif_codtipom is null and not numfactu is null and not rectif_fecfactu is null)"
'
'        importe2 = DevuelveValor(Sql)
'    End If
'
'    ImporteAlbaranFacturadoaSocio = Importe + importe2
'
'End Function
'
'Private Function AlbaranLiquidado(numalbar As String) As Boolean
'Dim Sql As String
'
'    Sql = "select count(*) from rfactsoc_albaran where numalbar = " & DBSet(numalbar, "N") & " and codtipom in (select codtipom from usuarios.stipom where tipodocu = 2)"
'
'    AlbaranLiquidado = (TotalRegistros(Sql) <> 0)
'
'End Function

Private Function ImporteAlbaranFacturadoaSocio(NumAlbar As String, Variedad As String) As Currency
Dim Sql As String
Dim Importe As Currency
Dim importe2 As Currency

    Sql = "select sum(importel) from rlifter where numalbar = " & DBSet(NumAlbar, "N") & " and codvarie = " & DBSet(Variedad, "N")
    If txtCodigo(2).Text <> "" Then Sql = Sql & " and fechaalb >= " & DBSet(txtCodigo(2).Text, "F")
    If txtCodigo(3).Text <> "" Then Sql = Sql & " and fechaalb <= " & DBSet(txtCodigo(3).Text, "F")
    Importe = DevuelveValor(Sql)
    
    Sql = "select sum(if(importe is null,0,importe) - if(imporgasto is null,0,imporgasto)) from rfactsoc_albaran where numalbar = " & DBSet(NumAlbar, "N") & " and codvarie = " & DBSet(Variedad, "N") & " and codtipom in (select codtipom from usuarios.stipom where tipodocu = 2)"
    If txtCodigo(2).Text <> "" Then Sql = Sql & " and fecalbar >= " & DBSet(txtCodigo(2).Text, "F")
    If txtCodigo(3).Text <> "" Then Sql = Sql & " and fecalbar <= " & DBSet(txtCodigo(3).Text, "F")
    
    Sql = Sql & " and not (rfactsoc_albaran.codtipom, rfactsoc_albaran.numfactu, rfactsoc_albaran.fecfactu) in (select rectif_codtipom,rectif_numfactu,rectif_fecfactu from rfactsoc "
    Sql = Sql & " where not rectif_codtipom is null and not numfactu is null and not rectif_fecfactu is null)"
    
    importe2 = DevuelveValor(Sql)
    
    '[Monica]06/11/2013: si no esta liquidado cogemos todo lo que haya anticipado
    If Not AlbaranLiquidado(NumAlbar, Variedad) Then
        Sql = "select sum(if(importe is null,0,importe) - if(imporgasto is null,0,imporgasto)) from rfactsoc_albaran where numalbar = " & DBSet(NumAlbar, "N") & " and codvarie = " & DBSet(Variedad, "N") & " and codtipom in (select codtipom from usuarios.stipom where tipodocu = 1)"
        
        If txtCodigo(2).Text <> "" Then Sql = Sql & " and fecalbar >= " & DBSet(txtCodigo(2).Text, "F")
        If txtCodigo(3).Text <> "" Then Sql = Sql & " and fecalbar <= " & DBSet(txtCodigo(3).Text, "F")
        
        Sql = Sql & " and not (rfactsoc_albaran.codtipom, rfactsoc_albaran.numfactu, rfactsoc_albaran.fecfactu) in (select rectif_codtipom,rectif_numfactu,rectif_fecfactu from rfactsoc "
        Sql = Sql & " where not rectif_codtipom is null and not numfactu is null and not rectif_fecfactu is null)"
        
        importe2 = DevuelveValor(Sql)
    End If

    ImporteAlbaranFacturadoaSocio = Importe + importe2
    
End Function

Private Function AlbaranLiquidado(NumAlbar As String, Variedad As String) As Boolean
Dim Sql As String

    Sql = "select count(*) from rfactsoc_albaran where numalbar = " & DBSet(NumAlbar, "N") & " and codvarie = " & DBSet(Variedad, "N") & " and codtipom in (select codtipom from usuarios.stipom where tipodocu = 2)"
    If txtCodigo(2).Text <> "" Then Sql = Sql & " and fecalbar >= " & DBSet(txtCodigo(2).Text, "F")
    If txtCodigo(3).Text <> "" Then Sql = Sql & " and fecalbar <= " & DBSet(txtCodigo(3).Text, "F")
    

    AlbaranLiquidado = (TotalRegistros(Sql) <> 0)

End Function


Private Function CalcularGastos(Albaran As Long, Linea As Integer) As Currency
Dim Sql As String

    CalcularGastos = 0
    
    CalcularGastos = CalcularGastos + CCur(ImporteSinFormato(TotalCostesEnvases(Albaran, Linea, 1)))
    CalcularGastos = CalcularGastos + CCur(ImporteSinFormato(TotalCostesEnvases(Albaran, Linea, 4)))
    CalcularGastos = CalcularGastos + CCur(ImporteSinFormato(TotalCostesEnvases(Albaran, Linea, 2)))
    CalcularGastos = CalcularGastos + CCur(ImporteSinFormato(TotalCostesEnvases(Albaran, Linea, 0)))
    CalcularGastos = CalcularGastos + CCur(ImporteSinFormato(TotalCostesEnvases(Albaran, Linea, 3)))

End Function
