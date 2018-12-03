VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListadoPed 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6825
   Icon            =   "frmListadoPed.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameGenAlbaran 
      Height          =   4110
      Left            =   45
      TabIndex        =   5
      Top             =   90
      Width           =   6435
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
         Left            =   2025
         MaxLength       =   3
         TabIndex        =   1
         Top             =   2715
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text5"
         Top             =   2715
         Width           =   3330
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   25
         Left            =   2025
         MaxLength       =   10
         TabIndex        =   0
         Top             =   2085
         Width           =   1215
      End
      Begin VB.CheckBox chkImpAlbaran 
         Caption         =   "Imprimir Albaran"
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
         Height          =   255
         Left            =   540
         TabIndex        =   2
         Top             =   3285
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdAceptarGenAlb 
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
         Left            =   3480
         TabIndex        =   3
         Top             =   3420
         Width           =   975
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
         Index           =   3
         Left            =   4560
         TabIndex        =   4
         Top             =   3420
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Incidencia"
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
         Index           =   12
         Left            =   540
         TabIndex        =   10
         Top             =   2715
         Width           =   1005
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   0
         Left            =   1755
         Top             =   2745
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Albaran"
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
         Index           =   32
         Left            =   540
         TabIndex        =   8
         Top             =   2085
         Width           =   1410
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1620
         Picture         =   "frmListadoPed.frx":000C
         Top             =   2085
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Pasar Pedido a Albaran"
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
         Left            =   600
         TabIndex        =   7
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Introduzca los siguiente campos para el Albaran: "
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
         Height          =   240
         Index           =   14
         Left            =   600
         TabIndex        =   6
         Top             =   1200
         Width           =   4980
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6030
      Top             =   4365
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmListadoPed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DatoSeleccionado(CadenaSeleccion As String)

Public Opcionlistado As Integer
'(ver opciones en frmListado)
      
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir


Public CodClien As String 'Para seleccionar inicialmente las ofertas del Proveedor

'#Laura 14/11/2006 Recuperar facturas Alzira
Public EstaRecupFact As Boolean ' si esta recuperando facturas (para albaranes de mostrador)


'Private HaDevueltoDatos As Boolean
Private NomTabla As String
Private NomTablaLin As String

Private WithEvents frmInc As frmManInciden
Attribute frmInc.VB_VarHelpID = -1


'Private WithEvents frmB As frmBuscaGrid  'Busquedas
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1

'----- Variables para el INforme ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String
Private numParam As Byte
Private cadselect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private Titulo As String 'Titulo informe que se pasa a frmImprimir
Private nomRPT As String 'nombre del fichero .rpt a imprimir
Private conSubRPT As Boolean 'si tiene subinformes para enlazarlos a las tablas correctas
'-------------------------------------


Dim indCodigo As Integer 'indice para txtCodigo

Dim PrimeraVez As Boolean


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub chkImpAlbaran_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub chkImpAlbaran_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub cmdAceptarGenAlb_Click()
'Solicitar datos para Generar Albaran a partir de un Pedido
Dim Cad As String

    'DAVID
    'Comprobar que me han puesto algun dato
    '-------------------------------------------------------------------
    If txtCodigo(25).Text = "" Then
        MsgBox "Campos obligatorios", vbExclamation
        Exit Sub
    End If
    
    Cad = Cad & txtCodigo(25).Text & "|"
    Cad = Cad & Me.chkImpAlbaran.Value & "|"
    Cad = Cad & txtCodigo(5).Text & "|"  'codigo de incidencia
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub


Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case Opcionlistado
            Case 43 '43: Generar Albaran desde Pedido (NO IMPRIME LISTADO)
                txtCodigo(5).Text = 0
                txtNombre(5) = PonerNombreDeCod(txtCodigo(5), "inciden", "nomincid", "codincid", "N")
                PonerFoco txtCodigo(25)
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim indFrame As Single

    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    limpiar Me

    'Ocultar todos los Frames de Formulario
    Me.FrameGenAlbaran.visible = False
    
    CommitConexion
    
    NomTabla = "scaped"
    NomTablaLin = "sliped"
        
    Select Case Opcionlistado
        'LISTADOS DE FACTURACION
        '-----------------------
            
        Case 43 '43: Generar Albaran desde Pedido (NO IMPRIME LISTADO)
            W = 6515
            H = 5415
            PonerFrameVisible Me.FrameGenAlbaran, True, H, W
            txtCodigo(25).Text = Format(Now, "dd/mm/yyyy")
            indFrame = 3
            Me.imgBuscarOfer(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture

            If NumCod = "REP" Then
                Label3.Caption = "Pasar Reparación a Albaran"
            Else
                Label3.Caption = "Pasar Pedido a Albaran"
            End If
        
    End Select
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
        
End Sub



Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtCodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub



Private Sub frmInc_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Incidencias
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscarOfer_Click(Index As Integer)
    Select Case Index
        Case 0 ' codigo de incidencia
            indCodigo = 5
            Set frmInc = New frmManInciden
            frmInc.DatosADevolverBusqueda = "0|1|"
            If Not IsNumeric(txtCodigo(indCodigo).Text) Then txtCodigo(indCodigo).Text = ""
            frmInc.Show vbModal
            Set frmInc = Nothing
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub


Private Sub imgFecha_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   
   '++monica
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim obj As Object

    Set frmF = New frmCal
    
    esq = imgFecha(Index).Left
    dalt = imgFecha(Index).Top
    
    Set obj = imgFecha(Index).Container

    While imgFecha(Index).Parent.Name <> obj.Name
        esq = esq + obj.Left
        dalt = dalt + obj.Top
        Set obj = obj.Container
    Wend
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    frmF.Left = esq + imgFecha(Index).Parent.Left + 30
    frmF.Top = dalt + imgFecha(Index).Parent.Top + imgFecha(Index).Height + menu - 40
   
   frmF.NovaData = Now
   
   Select Case Index
        Case 0 'Frame Pasar Pedido -> Albaran
            indCodigo = 25
   End Select
   
   PonerFormatoFecha txtCodigo(indCodigo)
   If txtCodigo(indCodigo).Text <> "" Then frmF.NovaData = CDate(txtCodigo(indCodigo).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtCodigo(indCodigo)
End Sub


Private Sub txtCodigo_GotFocus(Index As Integer)
    If Index <> 11 Then ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtCodigo_LostFocus(Index As Integer)
Dim devuelve As String
Dim codCampo As String, nomCampo As String
Dim tabla As String
      
    Select Case Index
        Case 5 ' codigo de incidencia
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "inciden", "nomincid", "codincid", "N")
                If txtCodigo(Index).Text <> "" And txtNombre(Index).Text <> "" Then
                    txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
                Else
                    PonerFoco txtCodigo(Index)
                End If
            Else
                txtNombre(Index).Text = ""
            End If
            
        
        'FECHA Desde Hasta
        Case 25
            If txtCodigo(Index).Text <> "" Then
                PonerFormatoFecha txtCodigo(Index)
                If Index = 34 Then _
                    txtCodigo(39).Text = Format(CDate(txtCodigo(34).Text) - 1, "dd/mm/yyyy")
            End If
           
    End Select
End Sub


Private Function AnyadirParametroDH(Cad As String, indD As Byte, indH As Byte) As String
On Error Resume Next

    If txtCodigo(indD).Text <> "" And txtCodigo(indH).Text <> "" Then
        If txtCodigo(indD).Text = txtCodigo(indH).Text Then
            Cad = Cad & txtCodigo(indD).Text
            If txtNombre(indD).Text <> "" Then Cad = Cad & " - " & txtNombre(indD).Text
            AnyadirParametroDH = Cad
            Exit Function
        End If
    End If
    
    If txtCodigo(indD).Text <> "" Then
        Cad = Cad & "desde " & txtCodigo(indD).Text
        If txtNombre(indD).Text <> "" Then Cad = Cad & " - " & txtNombre(indD).Text
    End If
    If txtCodigo(indH).Text <> "" Then
        Cad = Cad & "  hasta " & txtCodigo(indH).Text
        If txtNombre(indH).Text <> "" Then Cad = Cad & " - " & txtNombre(indH).Text
    End If
    AnyadirParametroDH = Cad
End Function


Private Function PonerDesdeHasta(campo As String, Tipo As String, indD As Byte, indH As Byte, param As String) As Boolean
Dim devuelve As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    If Tipo <> "F" Then
        If Not AnyadirAFormula(cadselect, devuelve) Then Exit Function
    End If
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            cadParam = cadParam & AnyadirParametroDH(param, indD, indH) & """|"
            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function


Private Sub InicializarVbles()
    cadFormula = ""
    cadselect = ""
    cadParam = ""
    numParam = 0
End Sub


Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = Opcionlistado
        .Titulo = Titulo
        .ConSubInforme = conSubRPT
        .NombreRPT = nomRPT  'nombre del informe
        .Show vbModal
    End With
End Sub


