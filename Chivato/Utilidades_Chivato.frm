VERSION 5.00
Begin VB.Form frmUtExport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportaciones a Chivato (Trazatec)"
   ClientHeight    =   3090
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   4680
   Icon            =   "Utilidades_Chivato.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2580
      TabIndex        =   3
      Tag             =   "Clasificación|N|N|0|1|variedades|tipoclasifica|0|N|"
      Text            =   "Combo1"
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmdPrueba 
      Caption         =   "Pruebas"
      Enabled         =   0   'False
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton cmdTodos 
      Caption         =   "Traspasar todos los datos"
      Height          =   495
      Left            =   450
      TabIndex        =   0
      Top             =   1800
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Selecciona Base de Datos :"
      Height          =   255
      Index           =   19
      Left            =   480
      TabIndex        =   4
      Top             =   390
      Width           =   2340
   End
   Begin VB.Label lblInf 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   4335
   End
End
Attribute VB_Name = "frmUtExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
    '-- Socios
'    i = 0
'    Sql = "select * from rsocios"
'    Set Rs = dbAriagro.cursor(Sql)
'    If Not Rs.EOF Then
'        Rs.MoveFirst
'        While Not Rs.EOF
'            i = i + 1
'            lblInf.Caption = "Socios --> " & CStr(i)
'            lblInf.Refresh
'            DoEvents
'            CargarUnSocio Rs!codsocio, "I"
'            Rs.MoveNext
'        Wend
'    End If
'    '-- Cuadrillas
'    i = 0
'    Sql = "select * from rcapataz"
'    Set Rs = dbAriagro.cursor(Sql)
'    If Not Rs.EOF Then
'        Rs.MoveFirst
'        While Not Rs.EOF
'            i = i + 1
'            lblInf.Caption = "Cuadrillas --> " & CStr(i)
'            lblInf.Refresh
'            DoEvents
'            CargarUnaCuadrilla Rs!codcapat, "I"
'            Rs.MoveNext
'        Wend
'    End If
'    '-- Partidas
'    i = 0
'    Sql = "select * from rpartida"
'    Set Rs = dbAriagro.cursor(Sql)
'    If Not Rs.EOF Then
'        Rs.MoveFirst
'        While Not Rs.EOF
'            i = i + 1
'            lblInf.Caption = "Partidas --> " & CStr(i)
'            lblInf.Refresh
'            DoEvents
'            CargarUnaPartida Rs!codparti, "I"
'            Rs.MoveNext
'        Wend
'    End If
'    '-- Vehiculos
'    i = 0
'    Sql = "select * from rtransporte"
'    Set Rs = dbAriagro.cursor(Sql)
'    If Not Rs.EOF Then
'        Rs.MoveFirst
'        While Not Rs.EOF
'            i = i + 1
'            lblInf.Caption = "Vehiculos --> " & CStr(i)
'            lblInf.Refresh
'            DoEvents
'            CargarUnVehiculo Rs!codtrans, "I"
'            Rs.MoveNext
'        Wend
'    End If
'    '-- Productos
'    i = 0
'    Sql = "select * from productos"
'    Set Rs = dbAriagro.cursor(Sql)
'    If Not Rs.EOF Then
'        Rs.MoveFirst
'        While Not Rs.EOF
'            i = i + 1
'            lblInf.Caption = "Productos --> " & CStr(i)
'            lblInf.Refresh
'            DoEvents
'            CargarUnProducto Rs!codprodu, "I"
'            Rs.MoveNext
'        Wend
'    End If
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
            CargarUnaVariedad Rs!codvarie, "U"
            Rs.MoveNext
        Wend
    End If
    '-- Campos
    i = 0
    Sql = "select * from rcampos"
    Set Rs = dbAriagro.cursor(Sql)
    If Not Rs.EOF Then
        Rs.MoveFirst
        While Not Rs.EOF
            i = i + 1
            lblInf.Caption = "Campos --> " & CStr(i)
            lblInf.Refresh
            DoEvents
            CargarUnCampo Rs!codcampo, "U"
            Rs.MoveNext
        Wend
    End If
    MsgBox "Proceso finalizado", vbExclamation
    Unload Me

End Sub

Private Sub CargaCombo()
Dim cad As String
Dim Rs As ADODB.Recordset
Dim i As Integer

    On Error GoTo ErrCarga
    Combo1.Clear
    'Conceptos
    
    cad = "SELECT ariagro FROM usuarios.empresasariagro ORDER BY ariagro"
    Set Rs = dbAriagro.cursor(cad)
    
    i = 0
    While Not Rs.EOF
        Combo1.AddItem Rs!ariagro
        Combo1.ItemData(Combo1.NewIndex) = i
        Rs.MoveNext
        '.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    Exit Sub
    
ErrCarga:
    If Err.Number <> 0 Then
        MsgBox "Error Cargar datos combo." & Err.Description, vbExclamation
    End If
End Sub

Private Sub Combo1_LostFocus()
    dbAriagro.abrir_MYSQL vConfig.SERVER, Combo1.Text, "root", "aritel"
End Sub

Private Sub Form_Load()
    
    CargaCombo
    Combo1.ListIndex = 0
    
End Sub
