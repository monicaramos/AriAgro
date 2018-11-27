VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTraerCalib 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Busqueda Calibres de Variedades"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7350
   Icon            =   "frmTraerCalib.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1080
      MaxLength       =   12
      TabIndex        =   0
      Tag             =   "año del Folleto|N|N|||follviaj|anyfovia|||"
      Text            =   "123456789012"
      Top             =   495
      Width           =   1620
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1350
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2381
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
      Left            =   4920
      TabIndex        =   1
      Top             =   2865
      Width           =   1065
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
      Left            =   6180
      TabIndex        =   2
      Top             =   2865
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "Datos de la Variedad"
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
      Height          =   255
      Index           =   2
      Left            =   135
      TabIndex        =   6
      Top             =   1035
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Calibre"
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
      Height          =   255
      Index           =   16
      Left            =   135
      TabIndex        =   5
      Top             =   480
      Width           =   1020
   End
   Begin VB.Label Label1 
      Caption         =   "Seleccione el calibre  que desee buscar:"
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
      Left            =   150
      TabIndex        =   3
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmTraerCalib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event Actualizar(vValor As String)


Public CodigoActual As String
Public Event DatoSeleccionado(CadenaSeleccion As String)
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados

Private Sub cmdAceptar_Click()
Dim Cad As Integer, cadAux As String
Dim i As Integer
Dim NumF As Long
Dim J As Integer
Dim Aux As String
Dim CADENA As String
    On Error Resume Next
    
    If Text1(0).Text = "" Then
        MsgBox "Debe introducir un valor en el calibre.", vbExclamation
        PonerFoco Text1(0)
        Exit Sub
    End If
    
    If Me.ListView1.ListItems.Count = 0 Then
        MsgBox "No existe este calibre en ninguna variedad.", vbExclamation
        PonerFoco Text1(0)
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    i = 1
    CADENA = ""
    While (i <= Me.ListView1.ListItems.Count)
        If Me.ListView1.ListItems(i).Checked Then
            Cad = Me.ListView1.ListItems(i)
            CADENA = CADENA & " codvarie = " & Me.ListView1.ListItems(i) & " or "
        End If
        i = i + 1
    Wend
    
    If CADENA = "" Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    CADENA = "(" & Mid(CADENA, 1, Len(CADENA) - 3) & ")"
    
    
    RaiseEvent Actualizar(CADENA)
    
    Screen.MousePointer = vbDefault
    Unload Me
    
    If Err.Number <> 0 Then Err.Clear
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Text1(0).Text = ""
    PonerFoco Text1(0)
End Sub

Private Sub Form_Load()
'    Text1(0).Text = ""
'    PonerFoco Text1(0)
End Sub

Private Sub CargarListView()
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem

    On Error GoTo ECargar

    'Los encabezados
    ListView1.ColumnHeaders.Clear
    Me.ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Add , , "Código", 1000
    ListView1.ColumnHeaders.Add , , "Nombre de la Variedad", 3000
    ListView1.ColumnHeaders.Add , , "Nombre del Calibre", 3000
    
    Sql = "select calibres.codvarie, variedades.nomvarie, calibres.nomcalib from variedades, calibres"
    Sql = Sql & " WHERE calibres.nomcalib=" & DBSet(Text1(0).Text, "T") & " AND calibres.codvarie = variedades.codvarie"
    Sql = Sql & " ORDER BY codvarie "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Set ItmX = ListView1.ListItems.Add
        ItmX.Text = Format(Rs.Fields(0).Value, "000000")
        
        ItmX.Checked = False
        
        ItmX.SubItems(1) = DBLet(Rs.Fields(1).Value, "T")
        ItmX.SubItems(2) = DBLet(Rs.Fields(2).Value, "T")
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
ECargar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar calibres.", Err.Description
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    If Text1(Index).Text <> "" Then
'       text1(0).Text = Format(text1(0).Text, "000000")
       CargarListView
       
    End If
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub
