VERSION 5.00
Begin VB.Form frmCaracteresMB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revisión Caracteres Multibase"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6090
   Icon            =   "frmCaracteresMB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameMultibase 
      BorderStyle     =   0  'None
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.CommandButton cmdMultibase2 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   20
         Top             =   5160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chkRoot 
         Alignment       =   1  'Right Justify
         Caption         =   "Tablas"
         Height          =   255
         Left            =   4710
         TabIndex        =   19
         Top             =   4770
         Width           =   975
      End
      Begin VB.Frame FrameTablas 
         Height          =   3375
         Left            =   210
         TabIndex        =   14
         Top             =   1260
         Visible         =   0   'False
         Width           =   5475
         Begin VB.ComboBox cboTablas 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   480
            Width           =   2895
         End
         Begin VB.ComboBox cboCampos 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   1560
            Width           =   2895
         End
         Begin VB.Label Label6 
            Caption         =   "TABLAS"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   480
            Width           =   3255
         End
         Begin VB.Label Label7 
            Caption         =   "TABLAS"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   1560
            Width           =   3255
         End
      End
      Begin VB.CheckBox chkMultibase 
         Caption         =   "Trabajadores"
         Height          =   255
         Index           =   4
         Left            =   2700
         TabIndex        =   13
         Top             =   3150
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkMultibase 
         Caption         =   "Proveedores"
         Height          =   255
         Index           =   3
         Left            =   2700
         TabIndex        =   12
         Top             =   2610
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CommandButton cmdMultiBase 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   3480
         TabIndex        =   5
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton cmdMultiBase 
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   4680
         TabIndex        =   4
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CheckBox chkMultibase 
         Caption         =   "Clientes"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   3
         Top             =   2640
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkMultibase 
         Caption         =   "Destinos"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   2
         Top             =   3180
         Value           =   1  'Checked
         Width           =   2145
      End
      Begin VB.CheckBox chkMultibase 
         Caption         =   "Artículos"
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   1
         Top             =   3720
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.Label lblMultibase 
         Height          =   375
         Left            =   210
         TabIndex        =   21
         Top             =   5100
         Width           =   2895
      End
      Begin VB.Label Label29 
         Caption         =   "Revisión caracteres multibase"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   720
         TabIndex        =   11
         Top             =   120
         Width           =   4935
      End
      Begin VB.Label Label30 
         Caption         =   "Utlidad para revisar los caracteres especiales que puedan quedar al realizar integraciones. "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   720
         Width           =   5775
      End
      Begin VB.Label Label31 
         Caption         =   "No debe trabajar nadie en la aplicación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   1320
         Width           =   4815
      End
      Begin VB.Label Label32 
         Caption         =   "A este proceso le puede costar mucho tiempo."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   1680
         Width           =   4815
      End
      Begin VB.Label Label33 
         Caption         =   "Datos a revisar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2280
         Width           =   4815
      End
      Begin VB.Label Label34 
         Caption         =   "Label34"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   4800
         Width           =   5535
      End
      Begin VB.Line Line5 
         X1              =   240
         X2              =   5640
         Y1              =   4140
         Y2              =   4140
      End
   End
End
Attribute VB_Name = "frmCaracteresMB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String

Dim devuelve As String
Dim miSQL As String

Private Sub cboTablas_Click()
    cboCampos.Clear
    If cboTablas.ListIndex < 0 Then Exit Sub
    CargarCamposTabla
End Sub


Private Sub CargarCamposTabla()
'Dim Cad As String
'Dim Aux As String
Dim RS As ADODB.Recordset
Dim i As Integer
Dim TieneClaves As Boolean

    
    miSQL = "Select * from " & Me.cboTablas.List(cboTablas.ListIndex) & " LIMIT 1,1"
    Set RS = New ADODB.Recordset
    RS.Open miSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
 
        TieneClaves = False
        For i = 0 To RS.Fields.Count - 1
           
            
            
            'SOLO TEXTOS
            If RS.Fields(i).Type = 129 Or RS.Fields(i).Type = 200 Or RS.Fields(i).Type = adVarChar Then
    
       
  
                If RS.Fields(i).Properties(18).Value Then
                    'NO HACEMOS NADA. Es campo clave
                
                Else
                    cboCampos.AddItem RS.Fields(i).Name
                End If
                
            End If
            
            'Para saber si tiene claves
            If RS.Fields(i).Properties(18).Value Then TieneClaves = True
            
        Next i
        
        
        
    RS.Close
    Set RS = Nothing

    If cboCampos.ListCount > 0 And Not TieneClaves Then
        MsgBox "No tiene campos clave", vbInformation
        Me.cboCampos.Clear
    End If
End Sub




Private Sub chkRoot_Click()
    Me.FrameTablas.visible = Me.chkRoot.Value = 1
    cmdMultibase2.visible = Me.chkRoot.Value = 1
    cmdMultiBase(0).visible = Me.chkRoot.Value <> 1
    If Me.chkRoot.Value = 1 Then
        If Me.cboTablas.ListCount = 0 Then
            Screen.MousePointer = vbHourglass
            Me.lblMultibase.Caption = "Cargando datos"
            Me.lblMultibase.Refresh
            
            CargaTablasCambio
            
            Screen.MousePointer = vbDefault
            Me.lblMultibase.Caption = ""
        End If
    End If
End Sub


Private Sub CargaTablasCambio()


    Set miRsAux = New ADODB.Recordset
    
    miRsAux.Open "show tables", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Me.cboTablas.AddItem miRsAux.Fields(0)
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing


End Sub





Private Sub cmdMultiBase_Click(Index As Integer)
Dim i As Integer
    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
    
    
    'Comprobamos k ha selecionado algun nivel
    NE = 0
    For i = 0 To Me.chkMultibase.Count - 1
        If Me.chkMultibase(i).Value = 1 Then NE = NE + 1
    Next i
    If NE = 0 Then
        MsgBox "Seleccione donde se van a realizar los cambios", vbExclamation
        Exit Sub
    End If
    
    'Comprobacion si hay alguien trabajando
    If UsuariosConectados Then Exit Sub
    
    SQL = "Seguro que desea continuar con el proceso"
    If MsgBox(SQL, vbCritical + vbYesNoCancel) <> vbYes Then Exit Sub
    
'   'BLOQUEAMOS LA BD
'   If Not Bloquear_DesbloquearBD(True) Then
'        MsgBox "No se ha podido bloquea a nivel de BD.", vbExclamation
'        Exit Sub
'    End If
'
    
    Screen.MousePointer = vbHourglass
    NumRegElim = 0
    For i = 0 To Me.chkMultibase.Count - 1
        If Me.chkMultibase(i).Value = 1 Then
            'Hacemos los cambios para ese valor
            HacerCambios i
        End If
    Next i
'    Bloquear_DesbloquearBD False
    Screen.MousePointer = vbDefault
    Label34.Caption = ""
    SQL = "Proceso finalizado" & vbCrLf & "Se han realizado: " & NumRegElim & " cambio(s)."
    MsgBox SQL, vbInformation
End Sub

Private Sub cmdMultibase2_Click()
    If cboTablas.ListIndex < 0 Then Exit Sub
    
    If MsgBox("Va a buscar en los campos seleccionados. ¿Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    UpdatearTablaRoot
    
    cadFrom = ""
    Me.lblMultibase.Caption = ""
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
End Sub


Private Sub UpdatearTablaRoot()
Dim i As Integer
Dim TienDatos As Boolean

    On Error GoTo EUpdatearTablaRoot
    
    devuelve = Me.cboTablas.List(cboTablas.ListIndex)
    miSQL = "Select " & Me.cboCampos.List(cboCampos.ListIndex) & "," & devuelve & ".* from " & devuelve

    miRsAux.Open miSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cadFrom = ""
    miSQL = ""
    While Not miRsAux.EOF
        If IsNull(miRsAux.Fields(0)) Then
            Me.lblMultibase.Caption = ""
            Me.lblMultibase.Refresh
        Else
            miSQL = miRsAux.Fields(0)
            Me.lblMultibase.Caption = miSQL
            Me.lblMultibase.Refresh
            devuelve = RevisaCaracterMultibase(miSQL)
            
            If miSQL <> devuelve Then
                    'La clave
                    cadFrom = ""
                    For i = 0 To miRsAux.Fields.Count - 1
                        If miRsAux.Fields(i).Properties(18).Value Then
                            Select Case miRsAux.Fields(i).Type
                            Case 133
                                campo = CStr(miRsAux.Fields(i))
                                campo = "'" & Format(campo, "yyyy-mm-dd") & "'"
            
                            Case 135 'Fecha/Hora
                                campo = DBSet(miRsAux.Fields(i), "FH", "S")
                            'Numero normal, sin decimales
                            Case 2, 3, 16 To 19
                                campo = miRsAux.Fields(i)
                            Case 129, 200
                                campo = DBSet(miRsAux.Fields(i), "T")
                            Case Else
                                MsgBox "No tratado: " & miRsAux.Fields(i).Type, vbExclamation
                                Exit Sub
                                
                            End Select
                            cadFrom = cadFrom & " AND " & miRsAux.Fields(i).Name & " = " & campo
                        End If
                    Next i
                    cadFrom = Mid(cadFrom, 6)
                    devuelve = DevNombreSQL(devuelve)
                    miSQL = "UPDATE " & Me.cboTablas.List(cboTablas.ListIndex) & " SET " & Me.cboCampos.List(cboCampos.ListIndex)
                    miSQL = miSQL & " = '" & devuelve & "' WHERE " & cadFrom
                    Conn.Execute miSQL
            End If 'DEl campo <>
        End If 'de ISNULL
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'If miSQL <> "" Then
        MsgBox "Proceso finalizado", vbInformation
    'Else
    '    MsgBox "No hay registros", vbInformation
    'End If
    Exit Sub
EUpdatearTablaRoot:
    MuestraError Err.Number, Err.Description
End Sub






Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
    Else
        Screen.MousePointer = vbDefault
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim W, H
    PrimeraVez = True
    
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    Me.frameMultibase.visible = False
    
    'MULTIBASE
    Me.Caption = "Sustitución caracteres multibase"
    W = Me.frameMultibase.Width
    H = Me.frameMultibase.Height + 300
    Me.frameMultibase.visible = True
    Me.FrameTablas.visible = False
    Label34.Caption = ""
    cmdMultiBase(1).Cancel = True
    Me.Width = W + 120
    Me.Height = H + 120
End Sub

Private Sub HacerCambios(ByVal Tabla As Integer)
Dim Cambio As String
Dim Inicio As Integer
Dim Fin As Integer
Dim Cad As String

    'RevisaCaracterMultibase
    Select Case Tabla
    Case 0
        'Clientes
        SQL = "Select codclien, nomclien, domclien, pobclien, proclien"
        SQL = SQL & " FROM clientes"
        Inicio = 1 'k es dos
        Fin = 4
    Case 1
        'Destinos
        SQL = "Select codclien, coddesti, nomdesti, pobdesti, prodesti  from destinos"
        Inicio = 3
        Fin = 1
    Case 2
        'Artículos
        SQL = "Select codartic, nomartic FROM sartic "
        Cad = ""
        Inicio = 1
        Fin = 1
    Case 3
        'Proveedores
        SQL = "Select codprove, nomprove, nomcomer, domprove, pobprove, proprove FROM proveedor "
        Cad = ""
        Inicio = 1
        Fin = 5
    Case 4
        'Trabajadores
        SQL = "Select codtraba, nomtraba, domtraba, pobtraba, protraba FROM straba "
        Cad = ""
        Inicio = 1
        Fin = 4
    End Select
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        
        While Not RS.EOF
            Label34.Caption = RS.Fields(0) & " - " & RS.Fields(1)
            Label34.Refresh
            Cambio = ""
            
            For i = Inicio To Fin
                'Campo no nulo
                If Not IsNull(RS.Fields(i)) Then
                    SQL = RS.Fields(i)
                    Cad = RevisaCaracterMultibase(SQL)
                    If SQL <> Cad Then
                        'Han habido cambios
                        If Cambio <> "" Then Cambio = Cambio & ","
'                        Sql = NombreSQL(Cad)
                        SQL = DevNombreSQL(Cad)
                        NumRegElim = NumRegElim + 1
                        Cambio = Cambio & RS.Fields(i).Name & " = '" & SQL & "'"
                    End If
                End If
            Next i
            If Cambio <> "" Then
                'OK HAY K CAMBIAR, k updatear
                Select Case Tabla
                Case 0 'clientes
                    SQL = "UPDATE clientes SET " & Cambio & " WHERE codclien =" & RS.Fields(0)
            
                Case 1 'destinos
                    SQL = "UPDATE destinos"
                    SQL = SQL & " SET " & Cambio & " WHERE codclien = " & DBSet(RS.Fields(0).Value, "N")
                    SQL = SQL & " and coddesti = " & DBSet(RS.Fields(1).Value, "N")
                
                Case 2 'articulos
                    SQL = "UPDATE sartic SET " & Cambio & " WHERE codartic =" & DBSet(RS.Fields(0).Value, "N")
                
                Case 3 'proveedor
                    SQL = "UPDATE proveedor SET " & Cambio & " WHERE codprove =" & DBSet(RS.Fields(0).Value, "N")
                
                Case 4 'trabajadores
                    SQL = "UPDATE straba SET " & Cambio & " WHERE codtraba =" & DBSet(RS.Fields(0).Value, "N")
                End Select
                
                'Ejecutamos
                Conn.Execute SQL
            End If
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing
            
End Sub

