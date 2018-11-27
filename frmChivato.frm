VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChivato 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportaciones a Chivato (Trazatec)"
   ClientHeight    =   5655
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   6855
   Icon            =   "frmChivato.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Operación"
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
      Height          =   735
      Left            =   540
      TabIndex        =   13
      Top             =   3840
      Width           =   5625
      Begin VB.OptionButton Option1 
         Caption         =   "Modificar"
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
         Left            =   3690
         TabIndex        =   15
         Top             =   330
         Width           =   1515
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Insertar"
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
         Left            =   1440
         TabIndex        =   14
         Top             =   330
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   395
      Left            =   5085
      TabIndex        =   3
      Top             =   5040
      Width           =   1065
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
      Left            =   2730
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1005
      Width           =   3360
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
      Left            =   2730
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   630
      Width           =   3360
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
      Left            =   1785
      MaxLength       =   6
      TabIndex        =   1
      Top             =   1005
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
      Index           =   0
      Left            =   1785
      MaxLength       =   6
      TabIndex        =   0
      Top             =   630
      Width           =   915
   End
   Begin VB.CommandButton cmdPrueba 
      Caption         =   "Pruebas"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1260
      TabIndex        =   5
      Top             =   3840
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton cmdTodos 
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
      Height          =   395
      Left            =   3870
      TabIndex        =   2
      Top             =   5040
      Width           =   1065
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2325
      Index           =   0
      Left            =   1800
      TabIndex        =   11
      Top             =   1410
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   4101
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   3
      Left            =   1485
      Picture         =   "frmChivato.frx":6852
      ToolTipText     =   "Desmarcar todos"
      Top             =   1410
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   2
      Left            =   1245
      Picture         =   "frmChivato.frx":7254
      ToolTipText     =   "Marcar todos"
      Top             =   1410
      Width           =   240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Opción"
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
      Index           =   12
      Left            =   510
      TabIndex        =   12
      Top             =   1410
      Width           =   675
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   1485
      MouseIcon       =   "frmChivato.frx":DAA6
      MousePointer    =   4  'Icon
      ToolTipText     =   "Buscar socio"
      Top             =   1005
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   1485
      MouseIcon       =   "frmChivato.frx":DBF8
      MousePointer    =   4  'Icon
      ToolTipText     =   "Buscar socio"
      Top             =   630
      Width           =   240
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Socio"
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
      Left            =   540
      TabIndex        =   10
      Top             =   345
      Width           =   540
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
      Left            =   810
      TabIndex        =   9
      Top             =   1005
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
      Index           =   13
      Left            =   810
      TabIndex        =   8
      Top             =   630
      Width           =   690
   End
   Begin VB.Label lblInf 
      Height          =   255
      Left            =   570
      TabIndex        =   4
      Top             =   4620
      Width           =   5565
   End
End
Attribute VB_Name = "frmChivato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim indCodigo As Integer


Private WithEvents frmSoc As frmBasico
Attribute frmSoc.VB_VarHelpID = -1

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrueba_Click()
    'CargarTodosLosCampos
    'CargarUnCampo 601, "D"
    'CargarUnaPoblacion 0, "I"
    'CargarUnSocio 1141, "I"
    'CargarUnaCuadrilla 15, "I"
    CargarUnaPartida 431, "I"
    CargarUnVehiculo 1, "I"
    'CargarUnProducto 1, "I"
    CargarUnaVariedad 119, "I"
    CargarUnCampo 601, "I"
    MsgBox "Proceso finalizado"

End Sub

Private Sub cmdTodos_Click()
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim i As Integer
    Dim jj As Integer
    Dim vCadena As String
    Dim Produ As Long
    
'    '-- Poblaciones
'    i = 0
'    sql = "select * from rpueblos"
'    Set rs = dbAriagro.cursor(sql)
'    If Not rs.EOF Then
'        rs.MoveFirst
'        While Not rs.EOF
'            i = i + 1
'            lblInf.Caption = "Poblaciones --> " & CStr(i)
'            lblInf.Refresh
'            DoEvents
'            CargarUnaPoblacion rs!codpobla, "I"
'            rs.MoveNext
'        Wend
'    End If

    For jj = 1 To ListView1(0).ListItems.Count
         If ListView1(0).ListItems(jj).Checked = True Then
            Select Case jj
                Case 1

                    '-- Socios
                    i = 0
                    Sql = "select * from rsocios where 1=1 "
                    If txtCodigo(0).Text <> "" Then Sql = Sql & " and codsocio >= " & DBSet(txtCodigo(0).Text, "N")
                    If txtCodigo(1).Text <> "" Then Sql = Sql & " and codsocio <= " & DBSet(txtCodigo(1).Text, "N")
                    
                    
                    Set Rs = dbAriagro.cursor(Sql)
                    If Not Rs.EOF Then
                        Rs.MoveFirst
                        While Not Rs.EOF
                            i = i + 1
                            lblInf.Caption = "Socios --> " & CStr(i)
                            lblInf.Refresh
                            DoEvents
                            
                            If option1(0).Value Then
                                CargarUnSocio Rs!CodSocio, "I"
                            Else
                                CargarUnSocio Rs!CodSocio, "U"
                            End If
                            Rs.MoveNext
                        Wend
                    End If
    
                Case 2
                    '-- Cuadrillas
                    i = 0
                    Sql = "select * from rcapataz"
                    Set Rs = dbAriagro.cursor(Sql)
                    If Not Rs.EOF Then
                        Rs.MoveFirst
                        While Not Rs.EOF
                            i = i + 1
                            lblInf.Caption = "Cuadrillas --> " & CStr(i)
                            lblInf.Refresh
                            DoEvents
                            If option1(0).Value Then
                                CargarUnaCuadrilla Rs!codcapat, "I"
                            Else
                                CargarUnaCuadrilla Rs!codcapat, "U"
                            End If
                            Rs.MoveNext
                        Wend
                    End If
    
                Case 3
                    '-- Partidas
                    i = 0
                    Sql = "select * from rpartida"
                    Set Rs = dbAriagro.cursor(Sql)
                    If Not Rs.EOF Then
                        Rs.MoveFirst
                        While Not Rs.EOF
                            i = i + 1
                            lblInf.Caption = "Partidas --> " & CStr(i)
                            lblInf.Refresh
                            DoEvents
                            If option1(0).Value Then
                                CargarUnaPartida Rs!codparti, "I"
                            Else
                                CargarUnaPartida Rs!codparti, "U"
                            End If
                            Rs.MoveNext
                        Wend
                    End If
    
                Case 4
                    '-- Vehiculos
                    i = 0
                    Sql = "select * from rtransporte"
                    Set Rs = dbAriagro.cursor(Sql)
                    If Not Rs.EOF Then
                        Rs.MoveFirst
                        While Not Rs.EOF
                            i = i + 1
                            lblInf.Caption = "Vehiculos --> " & CStr(i)
                            lblInf.Refresh
                            DoEvents
                            If option1(0).Value Then
                                CargarUnVehiculo Rs!codTrans, "I"
                            Else
                                CargarUnVehiculo Rs!codTrans, "U"
                            End If
                            Rs.MoveNext
                        Wend
                    End If
                Case 5
                    '-- Productos
                    i = 0
                    Sql = "select * from productos"
                    Set Rs = dbAriagro.cursor(Sql)
                    If Not Rs.EOF Then
                        Rs.MoveFirst
                        While Not Rs.EOF
                            i = i + 1
                            lblInf.Caption = "Productos --> " & CStr(i)
                            lblInf.Refresh
                            DoEvents
                            If option1(0).Value Then
                                CargarUnProducto Rs!codprodu, "I"
                            Else
                                CargarUnProducto Rs!codprodu, "U"
                            End If
                            Rs.MoveNext
                        Wend
                    End If
                Case 6
                    '-- Variedades
                    i = 0
                    Sql = "select * from variedades"
                    Set Rs = dbAriagro.cursor(Sql)
                    If Not Rs.EOF Then
                        Rs.MoveFirst
                        While Not Rs.EOF
                            i = i + 1
                            lblInf.Caption = "Variedades --> " & CStr(i)
                            lblInf.Refresh
                            DoEvents
                            If option1(0).Value Then
                                CargarUnaVariedad Rs!codvarie, "I"
                            Else
                            
                                '[Monica]18/09/2013: si estamos actualizando variedad en Picassent el claveant en 'PP&VVVV'
                                vCadena = DBLet(Rs!codprodu, "N") & "&" & DBLet(Rs!codvarie, "N")
                            
                                CargarUnaVariedad Rs!codvarie, "U", vCadena
                            End If
                            Rs.MoveNext
                        Wend
                    End If
                Case 7
                    '-- Campos
                    i = 0
                    Sql = "select * from rcampos where (1=1) "
                    If txtCodigo(0).Text <> "" Then Sql = Sql & " and codsocio >= " & DBSet(txtCodigo(0).Text, "N")
                    If txtCodigo(1).Text <> "" Then Sql = Sql & " and codsocio <= " & DBSet(txtCodigo(1).Text, "N")
                    
                    Set Rs = dbAriagro.cursor(Sql)
                    If Not Rs.EOF Then
                        Rs.MoveFirst
                        While Not Rs.EOF
                            i = i + 1
                            lblInf.Caption = "Campos --> " & CStr(i)
                            lblInf.Refresh
                            DoEvents
                            If option1(0).Value Then
                                CargarUnCampo Rs!codCampo, "I"
                            Else
                               '[Monica]17/09/2013:en el campo ant en picassent ponemos otra cosa
                                Produ = DevuelveValor("select codprodu from variedades where codvarie = " & DBLet(Rs!codvarie, "N"))
                                vCadena = DBLet(Rs!CodSocio, "N") & "&" & DBLet(Rs!codCampo, "N") & "&" & Produ & "&" & DBLet(Rs!codvarie, "N")

                                CargarUnCampo Rs!codCampo, "U", vCadena
                            End If
                            Rs.MoveNext
                        Wend
                    End If
            End Select
        End If
    Next jj
        
    MsgBox "Proceso finalizado", vbExclamation
    Unload Me

End Sub


Private Sub Form_Activate()
    PonerFoco txtCodigo(0)
End Sub

Private Sub Form_Load()
Dim i As Integer

    For i = 0 To 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i

    CargarListView (0)
    Image1_Click (2)
    Set dbAriagro = New BaseDatos
    dbAriagro.abrir_MYSQL vConfig.SERVER, vUsu.CadenaConexion, "root", "aritel"

    option1(0).Value = 1

    
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub Image1_Click(Index As Integer)
Dim i As Integer

    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 2
            For i = 1 To ListView1(0).ListItems.Count
                ListView1(0).ListItems(i).Checked = True
            Next i
        Case 3
            For i = 1 To ListView1(0).ListItems.Count
                ListView1(0).ListItems(i).Checked = False
            Next i
    End Select
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1 'Socios
        
            indCodigo = Index
        
            Set frmSoc = New frmBasico
            
            AyudaSocios frmSoc, 0
            
            Set frmSoc = Nothing
            
    End Select
    
    PonerFoco txtCodigo(indCodigo)

End Sub


Private Sub CargarListView(Index As Integer)
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String

    On Error GoTo ECargarList

    'Los encabezados
    ListView1(Index).ColumnHeaders.Clear

 '   ListView1.ColumnHeaders.Add , , "Tipo", 650
    ListView1(Index).ColumnHeaders.Add , , "Tablas", 3250
    
    Set ItmX = ListView1(Index).ListItems.Add
    ItmX.Text = "Socios"
    ItmX.Key = "a"
    
    Set ItmX = ListView1(Index).ListItems.Add
    ItmX.Text = "Cuadrillas"
    ItmX.Key = "b"
    
    Set ItmX = ListView1(Index).ListItems.Add
    ItmX.Text = "Partidas"
    ItmX.Key = "c"

    Set ItmX = ListView1(Index).ListItems.Add
    ItmX.Text = "Vehículos"
    ItmX.Key = "d"
    
    Set ItmX = ListView1(Index).ListItems.Add
    ItmX.Text = "Productos"
    ItmX.Key = "e"
    
    Set ItmX = ListView1(Index).ListItems.Add
    ItmX.Text = "Variedades"
    ItmX.Key = "f"
    
    Set ItmX = ListView1(Index).ListItems.Add
    ItmX.Text = "Campos"
    ItmX.Key = "g"
    
    
ECargarList:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargar Tablas.", Err.Description
    End If
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

Private Sub KEYpress(KeyAscii As Integer)
    Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub txtCodigo_LostFocus(Index As Integer)
Dim tabla As String
Dim codCampo As String, nomCampo As String
Dim TipCampo As String, Formato As String
Dim Titulo As String
Dim EsNomCod As Boolean 'Si es campo Cod-Descripcion llama a PonerNombreDeCod


    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    EsNomCod = False
        
    Select Case Index
        Case 0, 1  'SOCIOS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rsocios", "nomsocio", "codsocio", "N")
        
    End Select
    
End Sub

