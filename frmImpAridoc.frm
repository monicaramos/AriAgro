VERSION 5.00
Begin VB.Form frmImpAridoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportar datos a AriDoc"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1140
      Left            =   135
      TabIndex        =   7
      Top             =   1665
      Width           =   6585
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
         Index           =   1
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Tag             =   "Tipo|N|N|||straba|codsecci||N|"
         Top             =   540
         Width           =   2745
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
         Left            =   225
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   540
         Width           =   1350
      End
      Begin VB.Label Label8 
         Caption         =   "Secci�n "
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
         Height          =   260
         Left            =   2520
         TabIndex        =   11
         Top             =   225
         Width           =   1695
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1620
         Picture         =   "frmImpAridoc.frx":0000
         ToolTipText     =   "Buscar fecha"
         Top             =   225
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Recibo"
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
         Height          =   260
         Index           =   0
         Left            =   225
         TabIndex        =   9
         Top             =   225
         Width           =   1320
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Sobre Horas Productivas"
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
      Left            =   360
      TabIndex        =   17
      Top             =   3000
      Width           =   2940
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
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
      Left            =   5625
      TabIndex        =   2
      Top             =   3900
      Width           =   1065
   End
   Begin VB.CommandButton cmdAceptar 
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
      Left            =   4455
      TabIndex        =   1
      Top             =   3900
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Caption         =   "Carpeta de destino"
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
      Height          =   1215
      Left            =   135
      TabIndex        =   3
      Top             =   360
      Width           =   6600
      Begin VB.TextBox txtCarp 
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
         Left            =   1290
         TabIndex        =   5
         Top             =   600
         Width           =   5160
      End
      Begin VB.TextBox txtCarp 
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
         Left            =   225
         TabIndex        =   0
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "C�digo"
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
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   315
         Width           =   660
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   1140
      Left            =   135
      TabIndex        =   12
      Top             =   1620
      Width           =   6585
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
         Left            =   2865
         MaxLength       =   10
         TabIndex        =   14
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   720
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
         Index           =   0
         Left            =   2865
         MaxLength       =   10
         TabIndex        =   13
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   330
         Width           =   1350
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
         Height          =   195
         Index           =   16
         Left            =   180
         TabIndex        =   18
         Top             =   270
         Width           =   1500
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
         Left            =   1755
         TabIndex        =   16
         Top             =   360
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
         Left            =   1755
         TabIndex        =   15
         Top             =   720
         Width           =   735
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   2550
         Picture         =   "frmImpAridoc.frx":008B
         ToolTipText     =   "Buscar fecha"
         Top             =   360
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   2550
         Picture         =   "frmImpAridoc.frx":0116
         ToolTipText     =   "Buscar fecha"
         Top             =   720
         Width           =   240
      End
   End
   Begin VB.Label lblInf 
      Alignment       =   2  'Center
      Caption         =   "Informaci�n del proceso"
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
      Left            =   240
      TabIndex        =   6
      Top             =   3375
      Width           =   6465
   End
End
Attribute VB_Name = "frmImpAridoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Tipo As Byte
    'Tipo:  0 Impresion de albaranes
    '       1 Impresion de facturas de ventas
    '       2 Recibos n�mina ---> AHORA ESTA EN RECOLECCION

Dim DesdeFecha As Date
Dim Hastafecha As Date
Dim frmVis As frmVisReport
Dim impor As ArdImportador

Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1

Private Sub cmdAceptar_Click()
    If Not DatosOk() Then Exit Sub
    '-- Cargar facturas de gasolinera entre las fechas seleccionadas
    Select Case Tipo
        Case 0 ' albaranes de venta
            CargaAlbaranes DesdeFecha, Hastafecha
            MsgBox "Proceso finalizado", vbInformation
        Case 1 ' facturas de venta
            CargaFacturas DesdeFecha, Hastafecha
            MsgBox "Proceso finalizado", vbInformation
'        Case 2 ' recibos n�mina
'            CargaRecibos DesdeFecha, Hastafecha
'            MsgBox "Proceso finalizado", vbInformation
    End Select
    cmdSalir_Click
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Function DatosOk() As Boolean
    DesdeFecha = CDate(txtCodigo(0).Text)
    Hastafecha = CDate(txtCodigo(1).Text)
    If DesdeFecha > Hastafecha Then
        MsgBox "La fecha desde debe ser menor que la fecha hasta", vbInformation
        Exit Function
    End If
    If txtCarp(1) = "" Then
        MsgBox "Debe seleccionar una carpeta de importaci�n.", vbInformation
        Exit Function
    End If
    If Tipo = 2 Then
        If txtCodigo(2).Text = "" Then
            MsgBox "Debe introducir la fecha de Recibo.", vbInformation
            Exit Function
        End If
    End If
    DatosOk = True
End Function


Private Sub Combo1_LostFocus(Index As Integer)
    If Tipo = 2 Then
        Select Case Combo1(1).ListIndex
            Case 0
                Me.txtCarp(0).Text = vParamAplic.CarpetaRecCampo
            Case 1
                Me.txtCarp(0).Text = vParamAplic.CarpetaRecAlmacen
        End Select
        txtCarp_LostFocus (0)
    End If
End Sub

Private Sub Form_Load()
    txtCodigo(0).Text = Date
    txtCodigo(1).Text = Date
    txtCodigo(2).Text = Date
    Set impor = New ArdImportador
    
    Set ardDB = New BaseDatos
    ardDB.Tipo = "MYSQL"
    ardDB.abrir "Aridoc", "root", "aritel"
    
    Frame2.Enabled = (Tipo = 2)
    Frame2.visible = (Tipo = 2)
    
    Frame3.Enabled = (Tipo <> 2)
    Frame3.visible = (Tipo <> 2)
    
    CargaCombo
    
    Combo1(1).ListIndex = 1
    Check1(1).Enabled = False
    Check1(1).visible = False
    
    Select Case Tipo
        Case 0:
            Me.txtCarp(0).Text = vParamAplic.CarpetaAlb
        Case 1:
            Me.txtCarp(0).Text = vParamAplic.CarpetaFac
        Case 2
            Select Case Combo1(1).ListIndex
                Case 0
                    Me.txtCarp(0).Text = vParamAplic.CarpetaRecCampo
                Case 1
                    Me.txtCarp(0).Text = vParamAplic.CarpetaRecAlmacen
            End Select
            
            Check1(1).Enabled = True
            Check1(1).visible = True

    End Select
    txtCarp_LostFocus (0)
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
    imgFec(0).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtCodigo(Index).Text <> "" Then frmC.NovaData = txtCodigo(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtCodigo(CByte(imgFec(0).Tag))
    ' ***************************
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub txtCarp_GotFocus(Index As Integer)
    ConseguirFoco txtCarp(Index), 3
End Sub

Private Sub txtCarp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCarp_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCarp_LostFocus(Index As Integer)
Dim Cad As String

    If Index = 0 Then
        If txtCarp(0) <> "" Then 'txtCarp(1) = impor.nombreCarpeta(CLng(txtCarp(0)))
            Cad = CargaPath(txtCarp(Index))
            txtCarp(1).Text = Mid(Cad, 2, Len(Cad))
        End If
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
            Case 0: KEYFecha KeyAscii, 0 'fecha desde
            Case 1: KEYFecha KeyAscii, 1 'fecha hasta
            Case 2: KEYFecha KeyAscii, 1 'fecha recibo
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
        Case 0, 1, 2 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
    End Select
End Sub


Private Sub CargaFacturas(DFecha As Date, HFecha As Date)
    Dim db As BaseDatos
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim Rs2 As ADODB.Recordset
    Dim i As Long
    Dim FicheroPDF As String
    Dim C1 As String
    Dim c2 As String
    Dim c3 As String
    Dim c4 As String
    Dim f1 As Date
    Dim f3 As Date
    Dim i1 As Currency
    Dim fr As frmVisReport
On Error GoTo err_CargaFacturas
    
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim numParam As Byte
Dim cadParam As String

    Set db = New BaseDatos
    db.Tipo = "MYSQL"
    
    db.abrir_MYSQL vConfig.SERVER, vUsu.CadenaConexion, "root", "aritel"


'    db.abrir "accArigasol", "", ""
    Sql = "select facturas.*, stipom.letraser" & _
            " from facturas, usuarios.stipom stipom where fecfactu >= " & db.fecha(CDate(txtCodigo(0).Text)) & _
            " and fecfactu <= " & db.fecha(CDate(txtCodigo(1).Text)) & _
            " and facturas.codtipom = stipom.codtipom " & _
            " and facturas.pasaridoc = 0"
            
    Set Rs = db.cursor(Sql)
    
    
    If Not Rs.EOF Then
        Rs.MoveFirst
        While Not Rs.EOF
            i = i + 1
            lblInf.Caption = "Procesando registro " & CStr(i)
            lblInf.Refresh
            '-- Creamos el pdf
            FicheroPDF = App.path & "\ExpAriDoc.pdf"

'18/02/2010: lo quito para que prueben
'            If Not IntentaMatar(FicheroPDF) Then Err.Raise 53
            
            
            Set fr = New frmVisReport
            
            '++monica: seleccionamos que rpt se ha de ejecutar
            cadParam = "pEmpresa=""Ariagro""|"
            numParam = 1
            indRPT = 12 'Impresion de Factura
            If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu, True) Then Exit Sub
            '++
            fr.NumeroParametros = numParam
            fr.OtrosParametros = cadParam
            fr.ConSubInforme = True
            fr.Informe = App.path & "\Informes\" & nomDocu
            fr.FormulaSeleccion = "{facturas.codtipom} = '" & Rs!codTipoM & "' and " & _
                                            "{facturas.numfactu} =" & CStr(Rs!NumFactu) & " and " & _
                                            "{facturas.fecfactu} = Date(" & Format(Rs!FecFactu, "yyyy") & _
                                                                    "," & Format(Rs!FecFactu, "mm") & _
                                                                    "," & Format(Rs!FecFactu, "dd") & ")"
            fr.FicheroPDF = FicheroPDF
            Load fr 'trabaja sin mostrar el formulario
            Screen.MousePointer = vbDefault
'--monica
'            sql = "select * from clientes where codclien = " & db.numero(RS!CodClien)
'            Set Rs2 = db.cursor(sql)
'            c1 = Rs2!nomclien
'            c2 = Format(RS!numfactu, "0000000") & "-" & RS!letraser
'            c3 = "ARIAGRO"
'            c4 = RS!CodClien
'++monica: c1 a c4 esta parametrizado
            Sql = "select * from clientes where codclien = " & db.numero(Rs!CodClien)
            Set Rs2 = db.cursor(Sql)
            C1 = CargaParametroFac(vParamAplic.C1Factura, Rs, Rs2)
            c2 = CargaParametroFac(vParamAplic.C2Factura, Rs, Rs2)
            c3 = CargaParametroFac(vParamAplic.C3Factura, Rs, Rs2)
            c4 = CargaParametroFac(vParamAplic.C4Factura, Rs, Rs2)
            
            f1 = Rs!FecFactu
            i1 = Rs!TotalFac
            f3 = Now
            If impor.importaFicheroPDF(FicheroPDF, CLng(txtCarp(0)), C1, c2, c3, c4, f1, f3, i1) Then
                'actualizamos el pasaridoc de facturas
                Sql = "update facturas set pasaridoc = 1 where codtipom = " & DBSet(Rs!codTipoM, "T")
                Sql = Sql & " and numfactu = " & DBSet(Rs!NumFactu, "N") & " and fecfactu = " & DBSet(Rs!FecFactu, "F")
                db.ejecutar Sql
            End If
            
            Unload fr
            Set fr = Nothing
            
            Rs.MoveNext
        Wend
    End If
    Exit Sub
err_CargaFacturas:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "CargaFacturas"
    End If
End Sub




Private Sub CargaAlbaranes(DFecha As Date, HFecha As Date)
    Dim db As BaseDatos
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim Rs2 As ADODB.Recordset
    Dim i As Long
    Dim FicheroPDF As String
    Dim C1 As String
    Dim c2 As String
    Dim c3 As String
    Dim c4 As String
    Dim f1 As Date
    Dim f3 As Date
    Dim i1 As Currency
    Dim fr As frmVisReport
    
On Error GoTo err_CargaAlbaranes
    
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim numParam As Byte
Dim cadParam As String

    Set db = New BaseDatos
    db.Tipo = "MYSQL"
    
    db.abrir_MYSQL vConfig.SERVER, vUsu.CadenaConexion, "root", "aritel"
    
'    db.abrir "accArigasol", "", ""
    Sql = "select albaran.*" & _
            " from albaran where fechaalb >= " & db.fecha(CDate(txtCodigo(0).Text)) & _
            " and fechaalb <= " & db.fecha(CDate(txtCodigo(1).Text)) & _
            " and pasaridoc = 0 "
    Set Rs = db.cursor(Sql)
    If Not Rs.EOF Then
        Rs.MoveFirst
        While Not Rs.EOF
            i = i + 1
            lblInf.Caption = "Procesando registro " & CStr(i)
            lblInf.Refresh
            '-- Creamos el pdf
            FicheroPDF = App.path & "\ExpAriDoc.pdf"

'18/02/2010: lo quito para que prueben
'            If Not IntentaMatar(FicheroPDF) Then Err.Raise 53
            
            Set fr = New frmVisReport
            
            '++monica: seleccionamos que rpt se ha de ejecutar
            cadParam = "pEmpresa=""Ariagro""|"
            numParam = 1
            indRPT = 9 'Impresion de Albaran
            If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu, True) Then Exit Sub
            '++
            fr.NumeroParametros = numParam
            fr.OtrosParametros = cadParam
            fr.ConSubInforme = True
            fr.Informe = App.path & "\Informes\" & nomDocu
            fr.FormulaSeleccion = "{albaran.numalbar} =" & CStr(Rs!NumAlbar) & " and " & _
                                            "{albaran.fechaalb} = Date(" & Format(Rs!FechaAlb, "yyyy") & _
                                                                    "," & Format(Rs!FechaAlb, "mm") & _
                                                                    "," & Format(Rs!FechaAlb, "dd") & ")"
            fr.FicheroPDF = FicheroPDF
            Load fr 'trabaja sin mostrar el formulario
            Screen.MousePointer = vbDefault

            Sql = "select * from clientes where codclien = " & db.numero(Rs!CodClien)
            Set Rs2 = db.cursor(Sql)
'            c1 = RS2!nomclien
'            c2 = Format(RS!numalbar, "000000")
'            c3 = "ARIAGRO"
'            c4 = RS!CodClien
            C1 = CargaParametroAlb(vParamAplic.C1Albaran, Rs, Rs2)
            c2 = CargaParametroAlb(vParamAplic.C2Albaran, Rs, Rs2)
            c3 = CargaParametroAlb(vParamAplic.C3Albaran, Rs, Rs2)
            c4 = CargaParametroAlb(vParamAplic.C4Albaran, Rs, Rs2)

            f1 = Rs!FechaAlb
            f3 = Now
            i1 = 0
            
            If impor.importaFicheroPDF(FicheroPDF, CLng(txtCarp(0)), C1, c2, c3, c4, f1, f3, i1) Then
                'actualizamos el pasaridoc de albaranes
                Sql = "update albaran set pasaridoc = 1 where numalbar = " & DBSet(Rs!NumAlbar, "N")
                db.ejecutar Sql
            End If
            
            Unload fr
            Set fr = Nothing
            
            Rs.MoveNext
        Wend
    End If
    Exit Sub
err_CargaAlbaranes:
    If Err.Number Then
        MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "CargaAlbaranes"
    End If
End Sub


Private Sub CargaRecibos(DFecha As Date, HFecha As Date)
    Dim db As BaseDatos
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    Dim Rs2 As ADODB.Recordset
    Dim i As Long
    Dim FicheroPDF As String
    Dim C1 As String
    Dim c2 As String
    Dim c3 As String
    Dim c4 As String
    Dim f1 As Date
    Dim f3 As Date
    Dim i1 As Currency
    Dim fr As frmVisReport
On Error GoTo err_CargaRecibos
    
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim numParam As Byte
Dim cadParam As String

    Set db = New BaseDatos
    db.Tipo = "MYSQL"
    
    db.abrir_MYSQL vConfig.SERVER, vUsu.CadenaConexion, "root", "aritel"


'    db.abrir "accArigasol", "", ""
    Sql = "select horas.codtraba " & _
            " from horas where fecharec = " & db.fecha(CDate(txtCodigo(2).Text)) & _
            " and horas.pasaridoc = 0 " & _
            " and codtraba in (select codtraba from straba where codsecci = " & Combo1(1).ListIndex & ")" & _
            " group by codtraba "
            
    Set Rs = db.cursor(Sql)
    
    
    If Not Rs.EOF Then
        Rs.MoveFirst
        While Not Rs.EOF
            i = i + 1
            lblInf.Caption = "Procesando registro " & CStr(i)
            lblInf.Refresh
            '-- Creamos el pdf
            FicheroPDF = App.path & "\ExpAriDoc.pdf"
            
            If Not IntentaMatar(FicheroPDF) Then Err.Raise 53
            
            Set fr = New frmVisReport
            
            '++monica: seleccionamos que rpt se ha de ejecutar
            cadParam = "pEmpresa=""Ariagro""|"
            numParam = 1
            indRPT = 13 'Impresion de Factura
            If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu, True) Then Exit Sub
            '++
            cadParam = cadParam & "|pFecha=""" & txtCodigo(2).Text & """|"
            numParam = numParam + 1
            cadParam = cadParam & "|pTitulo=""" & "Recibo Horas " & Combo1(1).Text & """|"
            numParam = numParam + 1
            cadParam = cadParam & "|pHProductivas=" & Check1(1).Value & "|"
            numParam = numParam + 1
            
            
            
            fr.NumeroParametros = numParam
            fr.OtrosParametros = cadParam
            fr.ConSubInforme = False
            fr.Informe = App.path & "\Informes\" & nomDocu
            fr.FormulaSeleccion = "{horas.codtraba} = " & Rs!codtraba & " and " & _
                                           "{horas.fecharec} = Date(" & Format(CDate(txtCodigo(2).Text), "yyyy") & _
                                                                    "," & Format(CDate(txtCodigo(2).Text), "mm") & _
                                                                    "," & Format(CDate(txtCodigo(2).Text), "dd") & ") and " & _
                                           "{horas.pasaridoc} = 0 "
                                                                    
            fr.FicheroPDF = FicheroPDF
            Load fr 'trabaja sin mostrar el formulario
            Screen.MousePointer = vbDefault
'--monica
'            sql = "select * from clientes where codclien = " & db.numero(RS!CodClien)
'            Set Rs2 = db.cursor(sql)
'            c1 = Rs2!nomclien
'            c2 = Format(RS!numfactu, "0000000") & "-" & RS!letraser
'            c3 = "ARIAGRO"
'            c4 = RS!CodClien
'++monica: c1 a c4 esta parametrizado
            Sql = "select * from straba where codtraba = " & db.numero(Rs!codtraba)
            Set Rs2 = db.cursor(Sql)
            C1 = CargaParametroRec(vParamAplic.C1Recibo, Rs, Rs2)
            c2 = CargaParametroRec(vParamAplic.C2Recibo, Rs, Rs2)
            c3 = CargaParametroRec(vParamAplic.C3Recibo, Rs, Rs2)
            c4 = CargaParametroRec(vParamAplic.C4Recibo, Rs, Rs2)
            
'            f1 = RS!fechahora
'            i1 = RS!TotalFac
            f1 = CDate(txtCodigo(2).Text)
            i1 = 0
            f3 = Now
            If impor.importaFicheroPDF(FicheroPDF, CLng(txtCarp(0)), C1, c2, c3, c4, f1, f3, i1) Then
                'actualizamos el pasaridoc de facturas
                Sql = "update horas set pasaridoc = 1 where codtraba = " & DBSet(Rs!codtraba, "N")
    '            SQL = SQL & " and fechahora = " & DBSet(RS!fechahora, "F")
                db.ejecutar Sql
            End If
            
            Unload fr
            Set fr = Nothing
            
            Rs.MoveNext
        Wend
    End If
    Exit Sub
err_CargaRecibos:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "CargaRecibos"
    End If
End Sub




Private Function CargaParametroFac(param As Byte, ByRef Rs As ADODB.Recordset, ByRef Rs2 As ADODB.Recordset) As String
    Select Case param
        Case 0 'facturas
            CargaParametroFac = Format(Rs!NumFactu, "0000000") & "-" & Rs!letraser
        Case 1 'codigo cliente
            CargaParametroFac = Rs!CodClien
        Case 2 'nombre cliente
            CargaParametroFac = Rs2!Nomclien
        Case 3 'procedencia
            CargaParametroFac = "ARIAGRO"
        Case Else
            CargaParametroFac = ""
    End Select

End Function

Private Function CargaParametroAlb(param As Byte, ByRef Rs As ADODB.Recordset, ByRef Rs2 As ADODB.Recordset) As String
Dim Sql As String
Dim RS3 As ADODB.Recordset
Dim db As BaseDatos
    
    Set db = New BaseDatos
    db.Tipo = "MYSQL"
    
    db.abrir_MYSQL vConfig.SERVER, vUsu.CadenaConexion, "root", "aritel"

    Select Case param
        Case 0 'albaran
            CargaParametroAlb = Format(Rs!NumAlbar, "0000000") & "-" & Rs!letraser
        Case 1 'codigo cliente
            CargaParametroAlb = Rs!CodClien
        Case 2 'nombre cliente
            CargaParametroAlb = Rs2!Nomclien
        Case 3 'destino
            Sql = "select * from destinos where codclien = " & db.numero(Rs!CodClien)
            Sql = Sql & " and coddesti = " & db.numero(Rs!coddesti)
            Set RS3 = db.cursor(Sql)
            
            CargaParametroAlb = RS3!nomdesti
        Case 4 'procedencia
            CargaParametroAlb = "ARIAGRO"
        Case Else
            CargaParametroAlb = ""
    End Select
End Function

Private Function CargaParametroRec(param As Byte, ByRef Rs As ADODB.Recordset, ByRef Rs2 As ADODB.Recordset) As String
    Select Case param
        Case 0 'facturas
'            CargaParametroRec = Format(RS!numfactu, "0000000") & "-" & RS!letraser
            CargaParametroRec = Rs!codtraba
        Case 1 'codigo trabajador
            CargaParametroRec = Rs2!NomTraba
        Case 2 'nombre trabajador
            CargaParametroRec = "ARIAGRO"
        Case 3 'procedencia
            CargaParametroRec = "ARIAGRO"
        Case Else
            CargaParametroRec = ""
    End Select

End Function

Private Sub CargaCombo()
Dim Ini As Integer
Dim fin As Integer
Dim i As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
'    For I = 0 To Combo1.Count - 1
'        Combo1(I).Clear
'    Next I

    Combo1(1).Clear
    
    Combo1(1).AddItem "Campo"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Almac�n"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    
    
End Sub


Private Function CargaPath(Codigo As Integer) As String
Dim Nod As Node
Dim J As Integer
Dim i As Integer
Dim C As String
Dim campo1 As String
Dim padre As String
Dim A As String

Dim Sql As String
Dim Rs As ADODB.Recordset

    'distinto del cargapath de parametros de aplicacion

    Sql = "select nombre, padre from carpetas where codcarpeta = " & DBSet(Codigo, "N")
    Set Rs = ardDB.cursor(Sql)

    If Not Rs.EOF Then
        C = "\" & Rs!Nombre
        If Rs!padre > 0 Then
            C = CargaPath(CInt(Rs!padre)) & C
        End If
    End If
    
    CargaPath = C
End Function

Private Function IntentaMatar(FicheroPDF As String) As Boolean
Dim i As Integer

    On Error Resume Next
    i = 1
    IntentaMatar = False
    Do
        If Dir(FicheroPDF, vbArchive) <> "" Then
            Kill FicheroPDF
            If Err.Number <> 0 Then
                Err.Clear
                i = i + 1
            Else
                IntentaMatar = True
                i = 6
            End If
        Else
            IntentaMatar = True
            i = 6
        End If
    Loop Until i < 5 Or IntentaMatar = True
    
    
End Function

