VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOperaEnv 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Operaciones Globales sobre Confecciones"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6615
   Icon            =   "frmOperaEnv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   4095
      TabIndex        =   5
      Top             =   4545
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
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
      Left            =   5400
      TabIndex        =   6
      Top             =   4545
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Confecciones"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4200
      Left            =   90
      TabIndex        =   7
      Top             =   135
      Width           =   6405
      Begin MSComctlLib.ListView ListView1 
         Height          =   3375
         Left            =   90
         TabIndex        =   8
         Top             =   540
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   5953
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         Index           =   0
         Left            =   135
         Picture         =   "frmOperaEnv.frx":000C
         ToolTipText     =   "Marcar todos"
         Top             =   225
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   375
         Picture         =   "frmOperaEnv.frx":685E
         ToolTipText     =   "Desmarcar todos"
         Top             =   225
         Width           =   240
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   4245
      Left            =   90
      TabIndex        =   9
      Top             =   90
      Width           =   6495
      Begin VB.TextBox Text1 
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
         Left            =   1170
         MaxLength       =   16
         TabIndex        =   3
         Tag             =   "Codigo Articulo|T|N|||forfaits_envases|codartic||N|"
         Top             =   2070
         Width           =   1815
      End
      Begin VB.TextBox Text1 
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
         Left            =   1170
         MaxLength       =   6
         TabIndex        =   4
         Top             =   2790
         Width           =   780
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
         Left            =   3015
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2070
         Width           =   3240
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipo de Operación"
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
         Height          =   960
         Left            =   270
         TabIndex        =   10
         Top             =   270
         Width           =   5910
         Begin VB.OptionButton Option1 
            Caption         =   "Variación"
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
            Index           =   2
            Left            =   4005
            TabIndex        =   2
            Top             =   315
            Width           =   1500
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Borrado"
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
            Left            =   2250
            TabIndex        =   1
            Top             =   315
            Width           =   1500
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Inserción"
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
            Left            =   540
            TabIndex        =   0
            Top             =   315
            Width           =   1455
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
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
         Index           =   0
         Left            =   360
         TabIndex        =   13
         Top             =   2520
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Envase"
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
         Index           =   3
         Left            =   360
         TabIndex        =   12
         Top             =   1800
         Width           =   705
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   900
         MouseIcon       =   "frmOperaEnv.frx":7260
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   2070
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmOperaEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event Actualizar(vValor As Integer)


Public CodigoActual As String
Public Event DatoSeleccionado(CadenaSeleccion As String)
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados

Private WithEvents frmArt As frmManArtic 'articulos
Attribute frmArt.VB_VarHelpID = -1

Dim Modo As Byte ' 0 = pedir datos  1 = seleccionar confecciones

Private Sub cmdAceptar_Click()
Dim b As Boolean

    On Error Resume Next
    
    If DatosOk Then
        
        Screen.MousePointer = vbHourglass
        
        If Modo = 1 And Option1(0).Value Then
            Modo = 0
            Screen.MousePointer = vbDefault
            ActivarFrameDatos Modo
            Exit Sub
        End If
        
        b = False
    
        If Option1(0).Value And Modo = 0 Then b = InsercionMasiva
        If Option1(1).Value Then b = BorradoMasivo
        If Option1(2).Value Then b = VariacionMasiva
        
        If b Then MsgBox "El proceso se ha realizado correctamente.", vbExclamation
        
        Screen.MousePointer = vbDefault
        Unload Me
        
        If Err.Number <> 0 Then Err.Clear
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Modo = 1
    
    Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture

    ActivarFrameDatos Modo
    
    Me.Option1(0).Value = True
    
End Sub

Private Sub CargarListView()
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem

    On Error GoTo ECargar

    'Los encabezados
    ListView1.ColumnHeaders.Clear
    Me.ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Add , , "Código", 2000
    ListView1.ColumnHeaders.Add , , "Nombre", 4000
    
    
    Sql = "select forfaits.codforfait, forfaits.nomconfe "
    Sql = Sql & "from forfaits "
    Sql = Sql & " WHERE forfaits.codforfait not in (select distinct forfaits_envases.codforfait from forfaits_envases where codartic = " & DBSet(Text1(0).Text, "T") & ")"
    Sql = Sql & " GROUP BY forfaits.codforfait"
    Sql = Sql & " ORDER BY forfaits.codforfait"
    
    Set Rs = New ADODB.Recordset
    
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Set ItmX = ListView1.ListItems.Add
        
        ItmX.Checked = False
        ItmX.Text = DBLet(Rs.Fields(0).Value, "T")
        ItmX.SubItems(1) = DBLet(Rs.Fields(1).Value, "T")
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
ECargar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar confecciones.", Err.Description
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Articulos
    Text1(0).Text = RecuperaValor(CadenaSeleccion, 1) 'codartic
    txtNombre(0).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub Image1_Click(Index As Integer)
Dim i As Integer

    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 0
            For i = 1 To ListView1.ListItems.Count
                ListView1.ListItems(i).Checked = True
            Next i
        Case 1
            For i = 1 To ListView1.ListItems.Count
                ListView1.ListItems(i).Checked = False
            Next i

    End Select
    Screen.MousePointer = vbDefault

End Sub

Private Sub imgBuscar_Click(Index As Integer)
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Codigo Ariculo
            Set frmArt = New frmManArtic
            frmArt.DatosADevolverBusqueda = "0|1|"
            frmArt.Show vbModal
            Set frmArt = Nothing
    End Select
    PonerFoco Text1(Index)
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0
            Label2(0).Caption = "Cantidad"
            Text1(4).Text = ""
            Text1(4).visible = True
            Text1(4).Enabled = True
        Case 1
            Label2(0).Caption = ""
            Text1(4).visible = False
            Text1(4).Enabled = False
        Case 2
            Label2(0).Caption = "Porcentaje"
            Text1(4).Text = ""
            Text1(4).visible = True
            Text1(4).Enabled = True
    End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub ActivarFrameDatos(Modo As Byte)
' modo = 1 se visualiza el frame de pedir datos
' modo = 0 se visualiza el frame de pedir confeccioness
Dim b As Boolean

    If Modo = 0 Then CargarListView

    b = (Modo = 1)
    Me.Frame3.visible = b
    Me.Frame3.Enabled = b
    
    Me.Frame1.visible = Not b
    Me.Frame1.Enabled = Not b
End Sub


Private Function InsercionMasiva() As Boolean
Dim i As Integer
Dim Sql As String
Dim Sql2 As String

    On Error GoTo eInsercionMasiva

    InsercionMasiva = False
    i = 1
    Sql = ""
    While (i <= Me.ListView1.ListItems.Count)
        If Me.ListView1.ListItems(i).Checked Then
            Sql = Sql & "('" & Me.ListView1.ListItems(i).Text & "'," & DBSet(Text1(0).Text, "T")
            Sql = Sql & "," & DBSet(Text1(4).Text, "N") & "),"
        End If
        i = i + 1
    Wend

    'quitamos la ultima coma
    Sql = Mid(Sql, 1, Len(Sql) - 1)

    Sql2 = "insert into forfaits_envases (codforfait, codartic, cantidad) values " & Sql
    
    conn.Execute Sql2
    InsercionMasiva = True
    Exit Function
    
eInsercionMasiva:
    MuestraError Err.Number, "Error en la insercion masiva" & vbCrLf & Err.Description
End Function


Private Function BorradoMasivo() As Boolean
Dim Sql As String

    On Error GoTo eBorradoMasivo
    BorradoMasivo = False
    Sql = "delete from forfaits_envases where codartic = " & DBSet(Text1(0).Text, "T")
    
    conn.Execute Sql
    BorradoMasivo = True
    Exit Function
    
eBorradoMasivo:
    MuestraError Err.Number, "Error en el borrado masivo" & vbCrLf & Err.Description
End Function


Private Function VariacionMasiva() As Boolean
Dim Sql As String

    On Error GoTo eVariacionMasiva

    VariacionMasiva = False
    Sql = "update forfaits_envases set cantidad = cantidad + round(cantidad * " & DBSet(Text1(4).Text, "N") & " / 100,4) "
    Sql = Sql & " where codartic = " & DBSet(Text1(0).Text, "T")
    
    conn.Execute Sql
    VariacionMasiva = True
    Exit Function
    
eVariacionMasiva:
    MuestraError Err.Number, "Error en la variación masiva" & vbCrLf & Err.Description
End Function


Private Sub Text1_LostFocus(Index As Integer)
Dim tabla As String
Dim codCampo As String, nomCampo As String
Dim TipCampo As String, Formato As String
Dim Titulo As String
Dim EsNomCod As Boolean 'Si es campo Cod-Descripcion llama a PonerNombreDeCod
Dim cadMen As String


    'Quitar espacios en blanco por los lados
    Text1(Index).Text = Trim(Text1(Index).Text)

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    EsNomCod = False
        
    Select Case Index
        Case 0  'Cod. ARTICULO
'            txtNombre(Index).Text = PonerNombreDeCod(text1(Index), "sartic", "nomartic", "codartic", "T")
'[Monica]09/09/2013: error es un varchar
'             If PonerFormatoEntero(Text1(Index)) Then
                txtNombre(Index).Text = ""
                txtNombre(Index).Text = PonerNombreDeCod(Text1(Index), "sartic", "nomartic")
                If txtNombre(Index).Text = "" Then
                    cadMen = "No existe el Envase: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmArt = New frmManArtic
                        frmArt.DatosADevolverBusqueda = "0|1|"
                        frmArt.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        frmArt.Show vbModal
                        Set frmArt = Nothing
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
'            Else
'                txtNombre(Index).Text = ""
'            End If
        
        Case 4 ' cantidad o porcentaje de variacion
            If Option1(0).Value = True Then
                If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 7
            Else
                If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 2
            End If
    End Select

End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean

    b = True
    If Text1(0).Text = "" Then
        MsgBox "Debe introducir un Código de Envase obligatoriamente.", vbExclamation
        b = False
        PonerFoco Text1(0)
    Else
        If txtNombre(0).Text = "" Then
            MsgBox "Código de envase no existe. Revise.", vbExclamation
            b = False
            PonerFoco Text1(0)
        End If
    End If
    
    If b And Me.Option1(0).Value Then
        If Text1(4).Text = "" Then
            If MsgBox("La cantidad a insertar de envase es 0. Desea continuar", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                b = False
                PonerFoco Text1(4)
            End If
        End If
    End If
    If b And Me.Option1(2).Value Then
        If Text1(4).Text = "" Then
            If MsgBox("El porcentaje de aumento es 0. Desea continuar", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                b = False
                PonerFoco Text1(4)
            End If
        End If
    End If
    DatosOk = b
    
End Function
