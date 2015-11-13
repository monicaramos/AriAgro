VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
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
         Name            =   "Tahoma"
         Size            =   8.25
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
      Top             =   3750
      Width           =   5625
      Begin VB.OptionButton Option1 
         Caption         =   "Modificar"
         Height          =   195
         Index           =   1
         Left            =   3690
         TabIndex        =   15
         Top             =   330
         Width           =   1065
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Insertar"
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
      Height          =   495
      Left            =   4770
      TabIndex        =   3
      Top             =   4860
      Width           =   1335
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   2685
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1005
      Width           =   3360
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   2685
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   630
      Width           =   3360
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   1785
      MaxLength       =   6
      TabIndex        =   1
      Top             =   1005
      Width           =   830
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   1785
      MaxLength       =   6
      TabIndex        =   0
      Top             =   630
      Width           =   830
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
      Height          =   495
      Left            =   3270
      TabIndex        =   2
      Top             =   4860
      Width           =   1335
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2325
      Index           =   0
      Left            =   1800
      TabIndex        =   11
      Top             =   1410
      Width           =   3375
      _ExtentX        =   5953
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
      Left            =   1440
      Picture         =   "frmChivato.frx":6852
      ToolTipText     =   "Desmarcar todos"
      Top             =   1410
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   2
      Left            =   1200
      Picture         =   "frmChivato.frx":7254
      ToolTipText     =   "Marcar todos"
      Top             =   1410
      Width           =   240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Opción"
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
      Index           =   12
      Left            =   510
      TabIndex        =   12
      Top             =   1410
      Width           =   495
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
      Index           =   11
      Left            =   540
      TabIndex        =   10
      Top             =   390
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "Hasta"
      Height          =   195
      Index           =   12
      Left            =   900
      TabIndex        =   9
      Top             =   1005
      Width           =   420
   End
   Begin VB.Label Label4 
      Caption         =   "Desde"
      Height          =   195
      Index           =   13
      Left            =   900
      TabIndex        =   8
      Top             =   630
      Width           =   465
   End
   Begin VB.Label lblInf 
      Height          =   255
      Left            =   570
      TabIndex        =   4
      Top             =   4530
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

Private Sub CmdCancel_Click()
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
    Dim SQL As String
    Dim RS As ADODB.Recordset
    Dim I As Integer
    Dim JJ As Integer
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

    For JJ = 1 To ListView1(0).ListItems.Count
         If ListView1(0).ListItems(JJ).Checked = True Then
            Select Case JJ
                Case 1

                    '-- Socios
                    I = 0
                    SQL = "select * from rsocios where 1=1 "
                    If txtCodigo(0).Text <> "" Then SQL = SQL & " and codsocio >= " & DBSet(txtCodigo(0).Text, "N")
                    If txtCodigo(1).Text <> "" Then SQL = SQL & " and codsocio <= " & DBSet(txtCodigo(1).Text, "N")
                    
                    
                    Set RS = dbAriagro.cursor(SQL)
                    If Not RS.EOF Then
                        RS.MoveFirst
                        While Not RS.EOF
                            I = I + 1
                            lblInf.Caption = "Socios --> " & CStr(I)
                            lblInf.Refresh
                            DoEvents
                            
                            If Option1(0).Value Then
                                CargarUnSocio RS!CodSocio, "I"
                            Else
                                CargarUnSocio RS!CodSocio, "U"
                            End If
                            RS.MoveNext
                        Wend
                    End If
    
                Case 2
                    '-- Cuadrillas
                    I = 0
                    SQL = "select * from rcapataz"
                    Set RS = dbAriagro.cursor(SQL)
                    If Not RS.EOF Then
                        RS.MoveFirst
                        While Not RS.EOF
                            I = I + 1
                            lblInf.Caption = "Cuadrillas --> " & CStr(I)
                            lblInf.Refresh
                            DoEvents
                            If Option1(0).Value Then
                                CargarUnaCuadrilla RS!codcapat, "I"
                            Else
                                CargarUnaCuadrilla RS!codcapat, "U"
                            End If
                            RS.MoveNext
                        Wend
                    End If
    
                Case 3
                    '-- Partidas
                    I = 0
                    SQL = "select * from rpartida"
                    Set RS = dbAriagro.cursor(SQL)
                    If Not RS.EOF Then
                        RS.MoveFirst
                        While Not RS.EOF
                            I = I + 1
                            lblInf.Caption = "Partidas --> " & CStr(I)
                            lblInf.Refresh
                            DoEvents
                            If Option1(0).Value Then
                                CargarUnaPartida RS!codparti, "I"
                            Else
                                CargarUnaPartida RS!codparti, "U"
                            End If
                            RS.MoveNext
                        Wend
                    End If
    
                Case 4
                    '-- Vehiculos
                    I = 0
                    SQL = "select * from rtransporte"
                    Set RS = dbAriagro.cursor(SQL)
                    If Not RS.EOF Then
                        RS.MoveFirst
                        While Not RS.EOF
                            I = I + 1
                            lblInf.Caption = "Vehiculos --> " & CStr(I)
                            lblInf.Refresh
                            DoEvents
                            If Option1(0).Value Then
                                CargarUnVehiculo RS!codTrans, "I"
                            Else
                                CargarUnVehiculo RS!codTrans, "U"
                            End If
                            RS.MoveNext
                        Wend
                    End If
                Case 5
                    '-- Productos
                    I = 0
                    SQL = "select * from productos"
                    Set RS = dbAriagro.cursor(SQL)
                    If Not RS.EOF Then
                        RS.MoveFirst
                        While Not RS.EOF
                            I = I + 1
                            lblInf.Caption = "Productos --> " & CStr(I)
                            lblInf.Refresh
                            DoEvents
                            If Option1(0).Value Then
                                CargarUnProducto RS!codprodu, "I"
                            Else
                                CargarUnProducto RS!codprodu, "U"
                            End If
                            RS.MoveNext
                        Wend
                    End If
                Case 6
                    '-- Variedades
                    I = 0
                    SQL = "select * from variedades"
                    Set RS = dbAriagro.cursor(SQL)
                    If Not RS.EOF Then
                        RS.MoveFirst
                        While Not RS.EOF
                            I = I + 1
                            lblInf.Caption = "Variedades --> " & CStr(I)
                            lblInf.Refresh
                            DoEvents
                            If Option1(0).Value Then
                                CargarUnaVariedad RS!codvarie, "I"
                            Else
                            
                                '[Monica]18/09/2013: si estamos actualizando variedad en Picassent el claveant en 'PP&VVVV'
                                vCadena = DBLet(RS!codprodu, "N") & "&" & DBLet(RS!codvarie, "N")
                            
                                CargarUnaVariedad RS!codvarie, "U", vCadena
                            End If
                            RS.MoveNext
                        Wend
                    End If
                Case 7
                    '-- Campos
                    I = 0
                    SQL = "select * from rcampos where (1=1) "
                    If txtCodigo(0).Text <> "" Then SQL = SQL & " and codsocio >= " & DBSet(txtCodigo(0).Text, "N")
                    If txtCodigo(1).Text <> "" Then SQL = SQL & " and codsocio <= " & DBSet(txtCodigo(1).Text, "N")
                    
                    Set RS = dbAriagro.cursor(SQL)
                    If Not RS.EOF Then
                        RS.MoveFirst
                        While Not RS.EOF
                            I = I + 1
                            lblInf.Caption = "Campos --> " & CStr(I)
                            lblInf.Refresh
                            DoEvents
                            If Option1(0).Value Then
                                CargarUnCampo RS!CodCampo, "I"
                            Else
                               '[Monica]17/09/2013:en el campo ant en picassent ponemos otra cosa
                                Produ = DevuelveValor("select codprodu from variedades where codvarie = " & DBLet(RS!codvarie, "N"))
                                vCadena = DBLet(RS!CodSocio, "N") & "&" & DBLet(RS!CodCampo, "N") & "&" & Produ & "&" & DBLet(RS!codvarie, "N")

                                CargarUnCampo RS!CodCampo, "U", vCadena
                            End If
                            RS.MoveNext
                        Wend
                    End If
            End Select
        End If
    Next JJ
        
    MsgBox "Proceso finalizado", vbExclamation
    Unload Me

End Sub


Private Sub Form_Activate()
    PonerFoco txtCodigo(0)
End Sub

Private Sub Form_Load()
Dim I As Integer

    For I = 0 To 1
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I

    CargarListView (0)
    Image1_Click (2)
    Set dbAriagro = New BaseDatos
    dbAriagro.abrir_MYSQL vConfig.SERVER, vUsu.CadenaConexion, "root", "aritel"

    Option1(0).Value = 1

    
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub Image1_Click(Index As Integer)
Dim I As Integer

    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 2
            For I = 1 To ListView1(0).ListItems.Count
                ListView1(0).ListItems(I).Checked = True
            Next I
        Case 3
            For I = 1 To ListView1(0).ListItems.Count
                ListView1(0).ListItems(I).Checked = False
            Next I
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
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

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
Dim Tabla As String
Dim CodCampo As String, nomCampo As String
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

