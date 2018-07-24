VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListPaletsConf 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6660
   Icon            =   "frmListPaletsConf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameInfArticulos 
      Height          =   4950
      Left            =   -60
      TabIndex        =   0
      Top             =   0
      Width           =   6465
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
         Left            =   1725
         MaxLength       =   6
         TabIndex        =   4
         Top             =   2820
         Width           =   870
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
         Index           =   3
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Text5"
         Top             =   2850
         Width           =   3135
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
         Index           =   2
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "Text5"
         Top             =   2460
         Width           =   3135
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
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   3
         Top             =   2460
         Width           =   870
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo de Informe"
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
         Height          =   1470
         Left            =   360
         TabIndex        =   10
         Top             =   3210
         Width           =   2055
         Begin VB.OptionButton Opcion 
            Caption         =   "Todos"
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
            Index           =   2
            Left            =   225
            TabIndex        =   14
            Top             =   990
            Width           =   1560
         End
         Begin VB.OptionButton Opcion 
            Caption         =   "Asignados"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   0
            Left            =   225
            TabIndex        =   12
            Top             =   270
            Width           =   1650
         End
         Begin VB.OptionButton Opcion 
            Caption         =   "Pendientes"
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
            Index           =   1
            Left            =   225
            TabIndex        =   11
            Top             =   675
            Width           =   1560
         End
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
         Index           =   1
         Left            =   1710
         MaxLength       =   16
         TabIndex        =   2
         Top             =   1800
         Width           =   1350
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
         Index           =   0
         Left            =   1710
         MaxLength       =   16
         TabIndex        =   1
         Top             =   1440
         Width           =   1350
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
         Left            =   3795
         TabIndex        =   5
         Top             =   4305
         Width           =   1065
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
         Index           =   0
         Left            =   4965
         TabIndex        =   6
         Top             =   4305
         Width           =   1065
      End
      Begin VB.Label Label3 
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
         Index           =   1
         Left            =   765
         TabIndex        =   19
         Top             =   2460
         Width           =   690
      End
      Begin VB.Label Label3 
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
         Index           =   0
         Left            =   765
         TabIndex        =   18
         Top             =   2775
         Width           =   645
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1410
         MouseIcon       =   "frmListPaletsConf.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   2820
         Width           =   240
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   0
         Left            =   5820
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   2430
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1425
         MouseIcon       =   "frmListPaletsConf.frx":015E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   2460
         Width           =   240
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
         Left            =   480
         TabIndex        =   16
         Top             =   2190
         Width           =   675
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   1440
         Picture         =   "frmListPaletsConf.frx":02B0
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1440
         Picture         =   "frmListPaletsConf.frx":033B
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Informe de Palets"
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
         TabIndex        =   13
         Top             =   495
         Width           =   3900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   38
         Left            =   510
         TabIndex        =   9
         Top             =   1170
         Width           =   600
      End
      Begin VB.Label Label3 
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
         Index           =   54
         Left            =   735
         TabIndex        =   8
         Top             =   1800
         Width           =   645
      End
      Begin VB.Label Label3 
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
         Index           =   51
         Left            =   735
         TabIndex        =   7
         Top             =   1440
         Width           =   690
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   8730
      Top             =   5580
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmListPaletsConf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmConf As frmManForfaits
Attribute frmConf.VB_VarHelpID = -1

Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmCli As frmClientes 'Clientes
Attribute frmCli.VB_VarHelpID = -1

'---- Variables para el INFORME ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadselect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para el frmImprimir
Private conSubRPT As Boolean 'Si el informe tiene subreports
Private cadNombreRPT As String 'Nombre del informe
'-----------------------------------

Dim TipCod As String
Dim indCodigo As Integer 'indice para txtCodigo

Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean
Dim indFrame As Single


Private Sub KEYpress(KeyAscii As Integer)
    Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cmdAceptar_Click()
'Listado de Articulos
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim campo As String
Dim Opcion As Byte, numOp As Byte

    InicializarVbles
    
    cadNombreRPT = "rPaletsConf.rpt"  'Nombre fichero .rpt a Imprimir
    cadTABLA = "palets"
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    '====================================================
    '================= FORMULA ==========================
    
    'Cadena para seleccion D/H Fecha Fin
    '--------------------------------------------
    cDesde = Trim(txtCodigo(0).Text)
    cHasta = Trim(txtCodigo(1).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & cadTABLA & ".fechafin}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha= """) Then Exit Sub
    End If

    
    cadParam = cadParam & "pDesFec= """ & cDesde & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pHasFec= """ & cHasta & """|"
    numParam = numParam + 1
    
    'Obtener el parametro con el Orden del Informe
    '---------------------------------------------
    numOp = PonerGrupo(1, "TipMercan")
   
    'Parametro Orden del Informe
    If Me.Opcion(0).Value Then Opcion = 0
    If Me.Opcion(1).Value Then Opcion = 1
    If Me.Opcion(2).Value Then Opcion = 2
    
    cadTitulo = "Listado de Palets"
    Select Case Opcion
        Case 0
            cadTitulo = cadTitulo & " Asignados"
            AnyadirAFormula cadFormula, "not isnull({palets.numpedid})"
            AnyadirAFormula cadselect, "numpedid is not null"
                        
            ' cliente asignado al pedido
            If txtCodigo(2).Text <> "" Or txtCodigo(3).Text <> "" Then
                If InsertarTemporal(txtCodigo(2).Text) Then
                    cadNombreRPT = "rPaletsConf2.rpt"
                End If
            End If
            
        Case 1
            cadNombreRPT = "rPaletsConf1.rpt"
            '[Monica]06/05/2015: para Catadau ordenado por palet en lugar de por variedad
            If vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Then cadNombreRPT = "rPaletsConf.rpt"
            
            cadTitulo = cadTitulo & " Pendientes Asignar"
            AnyadirAFormula cadFormula, "isnull({palets.numpedid})"
            AnyadirAFormula cadselect, "numpedid is null "
        Case 2
    End Select
    
    campo = "pTipo=" & Opcion
    cadParam = cadParam & campo & "|"
    numParam = numParam + 1
    
    '[Monica]27/10/2015:cargamos una tabla temporal para ponerlo todo agrupado por variedad,confeccion y calidad
    cadParam = cadParam & "pUsu=" & vUsu.Codigo & "|"
    numParam = numParam + 1
    If Not CargarTemporal Then Exit Sub
    
    If HayRegParaInforme(cadTABLA, cadselect) Then
       LlamarImprimir
    End If
    
End Sub

Private Function CargarTemporal() As Boolean
Dim Sql As String
    
    On Error GoTo eCargarTemporal
    
    CargarTemporal = False
    
    Sql = "delete from tmpinformes2 where codusu = " & vUsu.Codigo
    conn.Execute Sql
    
    Sql = "insert into tmpinformes2 (codusu, importe1, nombre1, campo1, importe3, importe4) "
    
    Sql = Sql & " select " & vUsu.Codigo & ", palets_calibre.codvarie, palets_variedad.codforfait, palets_calibre.codcalib, sum(coalesce(palets_calibre.numcajas,0)), count(distinct palets.numpalet) "
    
    If Opcion(0).Value Then
        Sql = Sql & " from pedidos, palets, palets_variedad, palets_calibre "
        Sql = Sql & " where pedidos.numpedid = palets.numpedid and palets.numpalet = palets_variedad.numpalet and palets_variedad.numpalet = palets_calibre.numpalet and palets_variedad.numlinea = palets_calibre.numlinea "
    End If
    
    If Opcion(1).Value Then
        Sql = Sql & " from palets, palets_variedad, palets_calibre "
        Sql = Sql & " where palets.numpalet = palets_variedad.numpalet and palets_variedad.numpalet = palets_calibre.numpalet and palets_variedad.numlinea = palets_calibre.numlinea "
        Sql = Sql & " and (palets.numpedid is null or palets.numpedid = 0)"
    End If
    
    If Opcion(2).Value Then
        Sql = Sql & " from ((palets left join pedidos on pedidos.numpedid = palets.numpedid) inner join palets_variedad on palets.numpalet = palets_variedad.numpalet) inner join  palets_calibre on palets_variedad.numpalet = palets_calibre.numpalet and palets_variedad.numlinea = palets_calibre.numlinea "
        Sql = Sql & " where (1=1) "
    End If
    
    If txtCodigo(0).Text <> "" Then Sql = Sql & " and palets.fechafin >= " & DBSet(txtCodigo(0).Text, "F")
    If txtCodigo(1).Text <> "" Then Sql = Sql & " and palets.fechafin <= " & DBSet(txtCodigo(1).Text, "F")
    
    If Opcion(0).Value Or Opcion(2).Value Then
        If txtCodigo(2).Text <> "" Then Sql = Sql & " and pedidos.codclien >= " & DBSet(txtCodigo(2).Text, "N")
        If txtCodigo(3).Text <> "" Then Sql = Sql & " and pedidos.codclien <= " & DBSet(txtCodigo(3).Text, "N")
    End If
    
    Sql = Sql & " group by 1,2,3,4"
    Sql = Sql & " order by 1,2,3,4"
    
    conn.Execute Sql
                            
                            
    Sql = "delete from tmpliquidacion where codusu = " & vUsu.Codigo
    conn.Execute Sql
    
    Sql = "insert into tmpliquidacion  (codusu, codvarie, codsocio, nomvarie, codcampo, kilosnet, importe, gastos) "
    Sql = Sql & " select " & vUsu.Codigo & ", palets_variedad.codvarie, palets_variedad.codvarco, palets_variedad.codforfait,sum(coalesce(palets_variedad.pesobrut,0)), sum(coalesce(palets_variedad.pesoneto,0)), sum(coalesce(palets_variedad.numcajas,0)), count(distinct palets.numpalet) "
    
    If Opcion(0).Value Then
        Sql = Sql & " from pedidos, palets, palets_variedad "
        Sql = Sql & " where pedidos.numpedid = palets.numpedid and palets.numpalet = palets_variedad.numpalet "
    End If
    If Opcion(1).Value Then
        Sql = Sql & " from palets, palets_variedad "
        Sql = Sql & " where palets.numpalet = palets_variedad.numpalet and (palets.numpedid is null or palets.numpedid = 0) "
    End If
    If Opcion(2).Value Then
        Sql = Sql & " from (palets left join pedidos  on pedidos.numpedid = palets.numpedid) inner join palets_variedad on palets.numpalet = palets_variedad.numpalet "
        Sql = Sql & " where (1=1) "
    End If
    
    If txtCodigo(0).Text <> "" Then Sql = Sql & " and palets.fechafin >= " & DBSet(txtCodigo(0).Text, "F")
    If txtCodigo(1).Text <> "" Then Sql = Sql & " and palets.fechafin <= " & DBSet(txtCodigo(1).Text, "F")
    
    If Opcion(0).Value Or Opcion(2).Value Then
        If txtCodigo(2).Text <> "" Then Sql = Sql & " and pedidos.codclien >= " & DBSet(txtCodigo(2).Text, "N")
        If txtCodigo(3).Text <> "" Then Sql = Sql & " and pedidos.codclien <= " & DBSet(txtCodigo(3).Text, "N")
    End If
    
    Sql = Sql & " group by 1,2,3,4"
    Sql = Sql & " order by 1,2,3,4"
    
    conn.Execute Sql
                            
                            
    CargarTemporal = True
    Exit Function
                            
eCargarTemporal:
    MuestraError Err.Number, "Cargar Temporal", Err.Description
End Function



Private Function InsertarTemporal(Cliente As String) As Boolean
Dim Sql As String
    
    On Error GoTo eInsertarTemporal
    
    InsertarTemporal = False
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
    
    Sql = "insert into tmpinformes (codusu, importe1, importe2) "
    Sql = Sql & " select " & vUsu.Codigo & ", palets.numpalet, pedidos.codclien "
    Sql = Sql & " from pedidos, palets "
    Sql = Sql & " where pedidos.numpedid = palets.numpedid"
    
    If txtCodigo(0).Text <> "" Then Sql = Sql & " and palets.fechafin >= " & DBSet(txtCodigo(0).Text, "F")
    If txtCodigo(1).Text <> "" Then Sql = Sql & " and palets.fechafin <= " & DBSet(txtCodigo(1).Text, "F")
    
    If txtCodigo(2).Text <> "" Then Sql = Sql & " and pedidos.codclien >= " & DBSet(txtCodigo(2).Text, "N")
    If txtCodigo(3).Text <> "" Then Sql = Sql & " and pedidos.codclien <= " & DBSet(txtCodigo(3).Text, "N")
        
    conn.Execute Sql
                            
    InsertarTemporal = True
    Exit Function
                            
eInsertarTemporal:
    MuestraError Err.Number, "Insertar Temporal", Err.Description
End Function


Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub



Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(0)
        Me.Opcion(0).Value = True
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim i As Integer
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
     For H = 0 To imgBuscar.Count - 1
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Next H

    'Ocultar todos los Frames de Formulario
    Me.FrameInfArticulos.visible = False
    
    CommitConexion
    
    cadTitulo = ""
    cadNombreRPT = ""
    
    ListadosAlmacen H, W
    
    imgAyuda(0).Picture = frmPpal.ImageListB.ListImages(10).Picture
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(0).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub


Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtCodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String

    Select Case Index
        Case 0
           ' "____________________________________________________________"
            vCadena = "Sólo se utiliza si el listado es para el Tipo de Informe:" & vbCrLf & _
                      "Asignados, y metemos un valor en desde/hasta " & vbCrLf & _
                      "cliente." & vbCrLf & vbCrLf
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"
    
End Sub

Private Sub AbrirFrmClientes(indice As Integer)
    indCodigo = indice
    Set frmCli = New frmClientes
    frmCli.DatosADevolverBusqueda = "0|2|"
    frmCli.Show vbModal
    Set frmCli = Nothing
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1 'CLIENTE
            AbrirFrmClientes (Index)
        
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub imgFecha_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim obj As Object

    Screen.MousePointer = vbHourglass

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


   imgFecha(0).Tag = Index
'   Set frmF = New frmCal
   frmF.NovaData = Now
   
   indCodigo = Index
   
   PonerFormatoFecha txtCodigo(indCodigo)
   If txtCodigo(indCodigo).Text <> "" Then frmF.NovaData = CDate(txtCodigo(indCodigo).Text)
   
   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtCodigo(indCodigo)
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
        Case 0, 1 ' fechas
        
            If txtCodigo(Index).Text <> "" Then
                 PonerFormatoFecha txtCodigo(Index)
            End If
    
        Case 2, 3 ' clientes
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "clientes", "nomclien", "codclien", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
        
    End Select
    
End Sub

Private Sub ponerFrameArticulosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el informe de Articulos, de tabla: sartic
Dim b As Boolean

    b = True
    H = 4950 '4350
    W = 6465
    
    PonerFrameVisible Me.FrameInfArticulos, visible, H, W

End Sub



Private Sub InicializarVbles()
    cadFormula = ""
    cadselect = ""
    cadParam = ""
    numParam = 0
    conSubRPT = False
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
        .Titulo = cadTitulo
        .ConSubInforme = True
        .EnvioEMail = False
        .NombreRPT = cadNombreRPT
        .Opcion = 0 'Opcion
        .Show vbModal
    End With
End Sub


Private Function PonerGrupo(numGrupo As Byte, cadgrupo As String) As Byte
Dim campo As String
Dim nomCampo As String

    campo = "pGroup" & numGrupo & "="
    nomCampo = "pGroup" & numGrupo & "Name="
    PonerGrupo = 0
    
    Select Case cadgrupo
        Case "TipMercan"
            cadParam = cadParam & campo & "{palets.tipmercan}" & "|"
            numParam = numParam + 1
    End Select

End Function



Private Sub ListadosAlmacen(H As Integer, W As Integer)
   'Listado de Artículo
    ponerFrameArticulosVisible True, H, W
End Sub


