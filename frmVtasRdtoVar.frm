VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmVtasRdtoVar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6795
   Icon            =   "frmVtasRdtoVar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCobros 
      Height          =   6255
      Left            =   45
      TabIndex        =   10
      Top             =   0
      Width           =   6690
      Begin VB.Frame FrameEntradas 
         BorderStyle     =   0  'None
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
         Height          =   1575
         Left            =   2700
         TabIndex        =   25
         Top             =   3420
         Width           =   3495
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   11
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   7
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   840
            Width           =   1050
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   10
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   6
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   510
            Width           =   1050
         End
         Begin VB.Image imgAyuda 
            Height          =   240
            Index           =   0
            Left            =   3000
            MousePointer    =   4  'Icon
            Tag             =   "-1"
            ToolTipText     =   "Ayuda"
            Top             =   300
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Albarán"
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
            Height          =   255
            Index           =   4
            Left            =   270
            TabIndex        =   28
            Top             =   270
            Width           =   1155
         End
         Begin VB.Label Label4 
            Caption         =   "Desde"
            Height          =   195
            Index           =   3
            Left            =   930
            TabIndex        =   27
            Top             =   525
            Width           =   465
         End
         Begin VB.Label Label4 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   1
            Left            =   930
            TabIndex        =   26
            Top             =   840
            Width           =   420
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   3
            Left            =   1485
            Picture         =   "frmVtasRdtoVar.frx":000C
            ToolTipText     =   "Buscar fecha"
            Top             =   825
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   2
            Left            =   1485
            Picture         =   "frmVtasRdtoVar.frx":0097
            ToolTipText     =   "Buscar fecha"
            Top             =   510
            Width           =   240
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text5"
         Top             =   3060
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "Text5"
         Top             =   2745
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   1545
         MaxLength       =   3
         TabIndex        =   4
         Top             =   3060
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   1545
         MaxLength       =   3
         TabIndex        =   3
         Top             =   2745
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1545
         MaxLength       =   6
         TabIndex        =   5
         Top             =   3915
         Width           =   830
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4095
         TabIndex        =   8
         Top             =   5595
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5265
         TabIndex        =   9
         Top             =   5595
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1785
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2100
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1575
         MaxLength       =   6
         TabIndex        =   0
         Top             =   1245
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2475
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   1245
         Width           =   3720
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   390
         TabIndex        =   12
         Top             =   5055
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   13
         Left            =   1260
         MouseIcon       =   "frmVtasRdtoVar.frx":0122
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar mercado"
         Top             =   3060
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   12
         Left            =   1260
         MouseIcon       =   "frmVtasRdtoVar.frx":0274
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar mercado"
         Top             =   2745
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Mercado"
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
         Index           =   21
         Left            =   360
         TabIndex        =   24
         Top             =   2460
         Width           =   1185
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   22
         Left            =   660
         TabIndex        =   23
         Top             =   3060
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   23
         Left            =   660
         TabIndex        =   22
         Top             =   2745
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "%Incremento sobre Gasto"
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
         Index           =   0
         Left            =   315
         TabIndex        =   19
         Top             =   3645
         Width           =   1905
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   1275
         Picture         =   "frmVtasRdtoVar.frx":03C6
         ToolTipText     =   "Buscar fecha"
         Top             =   2100
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1275
         Picture         =   "frmVtasRdtoVar.frx":0451
         ToolTipText     =   "Buscar fecha"
         Top             =   1785
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   720
         TabIndex        =   18
         Top             =   2100
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   720
         TabIndex        =   17
         Top             =   1785
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
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
         Height          =   255
         Index           =   16
         Left            =   360
         TabIndex        =   16
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
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
         Index           =   2
         Left            =   360
         TabIndex        =   15
         Top             =   1050
         Width           =   630
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1275
         MouseIcon       =   "frmVtasRdtoVar.frx":04DC
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   1245
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Rendimiento por Variedad"
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
         Left            =   360
         TabIndex        =   14
         Top             =   495
         Width           =   5205
      End
      Begin VB.Label Label4 
         Caption         =   "Cargando tabla temporal"
         Height          =   195
         Index           =   27
         Left            =   390
         TabIndex        =   13
         Top             =   5325
         Width           =   3390
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6030
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmVtasRdtoVar"
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

Private WithEvents frmCla As frmManClases 'Clases
Attribute frmCla.VB_VarHelpID = -1
Private WithEvents frmVar As frmManVariedad 'Variedad
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmCli As frmClientes 'Clientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmDes As frmDestCli 'Destinos de Clientes
Attribute frmDes.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmFor As frmManForfaits 'Forfaits
Attribute frmFor.VB_VarHelpID = -1
Private WithEvents frmMar As frmManMarcas 'Marcas
Attribute frmMar.VB_VarHelpID = -1
Private WithEvents frmTMe As frmManTipMerc 'Tipos de Mercado
Attribute frmTMe.VB_VarHelpID = -1
Private WithEvents frmPais As frmManPaises 'Paises
Attribute frmPais.VB_VarHelpID = -1
Private WithEvents frmMensMercado As frmMensajes 'mensajes
Attribute frmMensMercado.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim Tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean

Dim GastosEnvases As Currency
Dim GastosPortes As Currency


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadtabla As String, cOrden As String
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim SQL1 As String
Dim vSqlMercado As String

InicializarVbles
    
    If txtCodigo(2).Text = "" Then
        MsgBox "Debe introducir una variedad. Reintroduzca.", vbExclamation
        Exit Sub
    End If
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
    numParam = numParam + 1
    
    
     '======== FORMULA  ====================================
    'Seleccionar registros de la empresa conectada
'    Codigo = "{" & tabla & ".codempre}=" & vEmpresa.codEmpre
'    If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
'    If Not AnyadirAFormula(cadSelect, Codigo) Then Exit Sub
    
    txtNombre(2).Text = PonerNombreDeCod(txtCodigo(2), "variedades", "nomvarie", "codvarie", "N")
    AnyadirAFormula cadSelect, "{albaran_variedad.codvarie} = " & DBSet(txtCodigo(2).Text, "N")
    
    'D/H Fecha albaran
    cDesde = Trim(txtCodigo(16).Text)
    cHasta = Trim(txtCodigo(17).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{albaran.fechaalb}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    'D/H Tipo de Mercado
    cDesde = Trim(txtCodigo(12).Text)
    cHasta = Trim(txtCodigo(13).Text)
    nDesde = txtNombre(12).Text
    nHasta = txtNombre(13).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{albaran.codtimer}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHMercado= """) Then Exit Sub
    End If
    vSqlMercado = ""
    If txtCodigo(12).Text <> "" Then vSqlMercado = vSqlMercado & " and tipomer.codtimer >= " & DBSet(txtCodigo(12).Text, "N")
    If txtCodigo(13).Text <> "" Then vSqlMercado = vSqlMercado & " and tipomer.codtimer <= " & DBSet(txtCodigo(13).Text, "N")

    
    If txtCodigo(0).Text <> "" Then
        cadParam = cadParam & "pGastosPor= " & TransformaComasPuntos(ImporteSinFormato(txtCodigo(0).Text)) & "|"
    Else
        cadParam = cadParam & "pGastosPor= 0|"
    End If
    numParam = numParam + 1
    
    ' Para punteo del tipo de mercado
    If vSqlMercado <> "" And txtCodigo(12).Text <> txtCodigo(13).Text Then
        Set frmMensMercado = New frmMensajes
    
        frmMensMercado.OpcionMensaje = 21
        frmMensMercado.Label5 = "Mercados"
        frmMensMercado.cadwhere = vSqlMercado
        frmMensMercado.Show vbModal
    
        Set frmMensMercado = Nothing
    End If
    
    
    cadtabla = Tabla & " INNER JOIN albaran_variedad ON albaran.numalbar = albaran_variedad.numalbar "
    cadtabla = "(" & cadtabla & ") INNER JOIN variedades ON albaran_variedad.codvarie = variedades.codvarie "
    
    ' solo los facturados
    cadtabla = "(" & cadtabla & ") LEFT JOIN facturas_variedad ON albaran_variedad.numalbar = facturas_variedad.numalbar and albaran_variedad.numlinea = facturas_variedad.numlinealbar "
    
    cadFormula = "{tmpinfventas.codusu} = " & vUsu.Codigo
    
    If HayRegistros(cadtabla, cadSelect) Then
        GastosEnvases = 0
        GastosPortes = 0
        If ProcesarCambios(cadtabla, cadSelect) Then
        
            If vParamAplic.Cooperativa = 9 Then
                If Not ProcesarCambios2 Then Exit Sub
                cadParam = cadParam & "pEntradas=1|"
                numParam = numParam + 1
            Else
                cadParam = cadParam & "pEntradas=0|"
                numParam = numParam + 1
            End If
        
              'Nombre fichero .rpt a Imprimir
              
            cadTitulo = "Rendimiento por Variedad"
            cadNombreRPT = "rRdtoVariedad1.rpt"
            cadParam = cadParam & "pOrden=""Variedad - Fecha""|"
            numParam = numParam + 1
            
            cadParam = cadParam & "pGastosEnv=" & DBSet(GastosEnvases, "N") & "|"
            numParam = numParam + 1
            
            cadParam = cadParam & "pGastosPortes=" & DBSet(GastosPortes, "N") & "|"
            numParam = numParam + 1
            
            LlamarImprimir
      End If
    End If
    
End Sub

Private Function ProcesarCambios(cadtabla, cadwhere As String) As Boolean
Dim SQL As String
Dim SQL1 As String
Dim Sql2 As String
Dim i As Integer
Dim HayReg As Long
Dim b As Boolean
Dim RS As ADODB.Recordset
Dim Rsx As ADODB.Recordset
Dim TotalGastos As Currency
Dim PesoCaja As Currency
Dim PesoReal As Currency
Dim ImpVenta As Currency
Dim Facturado As Byte
Dim Cobrado As Byte
Dim cadTabla2 As String

Dim Coste1 As Integer
Dim Coste2 As Integer
Dim Coste3 As Integer
Dim Coste4 As Integer

Dim Gasto1 As Currency
Dim Gasto2 As Currency
Dim Gasto3 As Currency
Dim Gasto4 As Currency
Dim Costes As Integer
'Dim GastosEnvases As Currency
'Dim GastosPortes As Currency

Dim A(100) As Currency
Dim Sql3 As String

On Error GoTo eProcesarCambios

    HayReg = 0
    
    ProcesarCambios = False
    
    conn.Execute "delete from tmpinfventas where codusu = " & DBSet(vUsu.Codigo, "N")
        
    If vParamAplic.Cooperativa = 9 Then
        conn.Execute "delete from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
    End If
        
    If cadwhere <> "" Then
        cadwhere = QuitarCaracterACadena(cadwhere, "{")
        cadwhere = QuitarCaracterACadena(cadwhere, "}")
        cadwhere = QuitarCaracterACadena(cadwhere, "_1")
    End If
        
    SQL1 = "select albaran.fechaalb, albaran.numalbar, albaran_variedad.numlinea, "
    SQL1 = SQL1 & "albaran_variedad.numcajas, albaran_variedad.pesoneto, albaran_variedad.preciopro, "
    SQL1 = SQL1 & "sum(facturas_variedad.impornet), "
    SQL1 = SQL1 & "albaran_variedad.preciodef from " & cadtabla
    SQL1 = SQL1 & " where (1 = 1) "
    If cadwhere <> "" Then SQL1 = SQL1 & " and " & cadwhere
    SQL1 = SQL1 & " group by 1, 2, 3, 4, 5, 6"
    SQL1 = SQL1 & " order by 1, 2, 3, 4, 5, 6"
        
    Set RS = New ADODB.Recordset
    RS.Open SQL1, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Label4(27).visible = True
    Pb1.visible = True
        
    HayReg = TotalRegistrosConsulta(SQL1)
    
    Pb1.Max = HayReg
    Pb1.Value = 0
    
    SQL1 = "select sum(impcoste) from (" & cadtabla
    SQL1 = SQL1 & ") INNER JOIN albaran_costes on albaran_variedad.numalbar = albaran_costes.numalbar and albaran_variedad.numlinea = albaran_costes.numlinea "
    SQL1 = SQL1 & " where albaran_costes.tipogasto = 1 "
    If cadwhere <> "" Then SQL1 = SQL1 & " and " & cadwhere
    GastosEnvases = DevuelveValor(SQL1)
    
    
    '[Monica] 15/06/2010 añadido costes paletizacion
    SQL1 = "select sum(impcoste) from (" & cadtabla
    SQL1 = SQL1 & ") INNER JOIN albaran_costes on albaran_variedad.numalbar = albaran_costes.numalbar and albaran_variedad.numlinea = albaran_costes.numlinea "
    SQL1 = SQL1 & " where albaran_costes.tipogasto = 4 "
    If cadwhere <> "" Then SQL1 = SQL1 & " and " & cadwhere
    GastosEnvases = GastosEnvases + DevuelveValor(SQL1)
    
    
    SQL1 = "select sum(impcoste) from (" & cadtabla
    SQL1 = SQL1 & ") INNER JOIN albaran_costes on albaran_variedad.numalbar = albaran_costes.numalbar and albaran_variedad.numlinea = albaran_costes.numlinea "
    SQL1 = SQL1 & " where albaran_costes.tipogasto = 2 "
    If cadwhere <> "" Then SQL1 = SQL1 & " and " & cadwhere
    GastosPortes = DevuelveValor(SQL1)
              
    
    Coste1 = -1
    Coste2 = -1
    Coste3 = -1
    Coste4 = -1
    
    ' tmpinfcostes: tabla donde insertaremos los costes
    Sql2 = "delete from tmpinfcostes where codusu = " & vUsu.Codigo
    conn.Execute Sql2
    
    Sql2 = "insert into tmpinfcostes(codusu, codcoste, denominacion, importe) select " & vUsu.Codigo
    Sql2 = Sql2 & ", codcoste, denominacion, 0 from nombcoste "
    conn.Execute Sql2
            
    For i = 0 To 10
        A(i) = 0
    Next i
    
    SQL = ""
    
    While Not RS.EOF
        IncrementarProgresNew Pb1, 1
    
        Sql2 = "select sum(impcoste) from albaran_costes where numalbar = "
        Sql2 = Sql2 & DBSet(RS.Fields(1).Value, "N") & " and numlinea = "
        Sql2 = Sql2 & DBSet(RS.Fields(2).Value, "N")
        
        TotalGastos = DevuelveValor(Sql2)
        
        
        ImpVenta = 0
        If Not IsNull(RS.Fields(6).Value) Then
            ImpVenta = RS.Fields(6).Value
            Facturado = 1
        Else
            '[Monica]07/06/2013: he añadido que si hay precio definitivo se calcula el importe con el `recio definitivo
            '                    y si no lo hay se calcula con el precio provisional
            If DBLet(RS.Fields(7).Value, "N") <> 0 Then
                ImpVenta = Round2(DBLet(RS.Fields(4).Value, "N") * DBLet(RS.Fields(7).Value, "N"), 2)
            Else
                ImpVenta = Round2(DBLet(RS.Fields(4).Value, "N") * DBLet(RS.Fields(5).Value, "N"), 2)
            End If
            Facturado = 0
        End If
        
        Gasto1 = 0
        Gasto2 = 0
        Gasto3 = 0
        Gasto4 = 0
        
        Sql2 = "select codcoste, impcoste from albaran_costes where albaran_costes.numalbar = " & DBSet(RS.Fields(1).Value, "N")
        Sql2 = Sql2 & " and albaran_costes.numlinea = " & DBSet(RS.Fields(2).Value, "N")
        Sql2 = Sql2 & " and albaran_costes.tipogasto = 0 "
        
        Set Rsx = New ADODB.Recordset
        Rsx.Open Sql2, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        Gasto2 = 0
        
        While Not Rsx.EOF
            Coste1 = DBLet(Rsx.Fields(0).Value, "N")
            Gasto1 = Round2(DBLet(Rsx.Fields(1).Value, "N"), 2)
            
            Gasto2 = Gasto2 + Gasto1
            
           A(Coste1) = A(Coste1) + Gasto1
            
           Rsx.MoveNext
        Wend
        
        SQL = SQL & "(" & DBSet(vUsu.Codigo, "N") & ","
        SQL = SQL & DBSet(RS.Fields(0).Value, "F") & "," & DBSet(RS.Fields(1).Value, "N") & "," & DBSet(RS.Fields(2).Value, "N") & ","
        SQL = SQL & DBSet(RS.Fields(3).Value, "N") & "," 'numero de cajas
        SQL = SQL & "0," & DBSet(RS.Fields(4).Value, "N") & "," 'peso neto
        SQL = SQL & DBSet(TotalGastos, "N") & "," & DBSet(ImpVenta, "N") & "," ' importe de venta
        SQL = SQL & DBSet(Facturado, "N") & ","  'facturado o no
        SQL = SQL & "0," 'cobrado o no
        SQL = SQL & DBSet(Coste1, "N") & "," & DBSet(Gasto2, "N") & "," 'coste1 gasto1
        SQL = SQL & "0,0,0,0,0,0,"
        SQL = SQL & "0,"  ' gastos portes
        SQL = SQL & "0)," ' gastos envases
      
        RS.MoveNext
    Wend
    
'++monica: rapidez un solo insert
    If SQL <> "" Then ' quitamos la ultima coma
        SQL = Mid(SQL, 1, Len(SQL) - 1)
        
        Sql3 = "insert into tmpinfventas (codusu, fecalbar, numalbar, numlinea, numcajas, pesoreal, pesoneto, gastos, impventa, facturado, cobrado, "
        Sql3 = Sql3 & " codigo1, gastos1, codigo2, gastos2, codigo3, gastos3, codigo4, gastos4, gastosportes, gastosenvases) values "
        Sql3 = Sql3 & SQL
   
       conn.Execute Sql3
    End If
        
'++monica:rapidez
    For i = 0 To 10
        If A(i) <> 0 Then
            Sql2 = "update tmpinfcostes set importe =  " & DBSet(A(i), "N")
            Sql2 = Sql2 & " where codusu = " & vUsu.Codigo & " and codcoste = " & DBSet(i, "N")

            conn.Execute Sql2
        End If
    Next i
    
    
    ProcesarCambios = True

    Label4(27).visible = False
    Pb1.visible = False
    
eProcesarCambios:
    If Err.Number <> 0 Then
        ProcesarCambios = False
    End If
End Function


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(2)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
     For h = 2 To 2 'imgBuscar.Count - 1
        Me.imgBuscar(h).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Next h
     For h = 12 To 13 'imgBuscar.Count - 1
        Me.imgBuscar(h).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Next h
    
    imgAyuda(0).Picture = frmPpal.ImageListB.ListImages(10).Picture

    '###Descomentar
'    CommitConexion
         
    Me.FrameEntradas.visible = (vParamAplic.Cooperativa = 9)
    Me.FrameEntradas.Enabled = (vParamAplic.Cooperativa = 9)
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    Tabla = "albaran"
    
    Label4(27).visible = False
    Pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = w + 70
    Me.Height = h + 350
End Sub


Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmTMe_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Tipo de Mercado
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Variedades
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           ' "____________________________________________________________"
            vCadena = "Si se quieren visualizar los datos de entradas debe indicar el rango de  " & vbCrLf & _
                      "Fechas de los albaranes de entrada de Recolección. " & vbCrLf & vbCrLf

    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"
    

End Sub

Private Sub imgFec_Click(Index As Integer)
Dim Indice As Integer

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

    Select Case Index
        Case 0, 1
            Indice = Index + 16
        Case 2, 3
            Indice = Index + 8
    End Select
    ' ***canviar l'index de imgFec pel 1r index de les imagens de buscar data***
    imgFec(0).Tag = Indice 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtCodigo(Indice).Text <> "" Then frmC.NovaData = txtCodigo(Indice).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtCodigo(CByte(imgFec(0).Tag))
    ' ***************************
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 2 'VARIEDADES
            AbrirFrmVariedades (Index)
        
        Case 12, 13 'TIPO DE MERCADO
            AbrirFrmMercados (Index)

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
            Case 2: KEYBusqueda KeyAscii, 2 'variedad desde
            Case 16: KEYFecha KeyAscii, 0 'fecha desde
            Case 17: KEYFecha KeyAscii, 1 'fecha hasta
            Case 12: KEYBusqueda KeyAscii, 12 'tipo de mercado desde
            Case 13: KEYBusqueda KeyAscii, 13 'tipo de mercado hasta
            Case 10: KEYFecha KeyAscii, 2 'fecha desde
            Case 11: KEYFecha KeyAscii, 3 'fecha hasta
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (Indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFec_Click (Indice)
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0 ' PORCENTAJE DE INCREMENTO DE GASTO SOBRE IMPORTE
            If txtCodigo(Index).Text <> "" Then PonerFormatoDecimal txtCodigo(Index), 4
            
        Case 2 'VARIEDAD
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            If txtCodigo(Index).Text <> "" Then
                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
                If txtNombre(Index).Text = "" Then
                    MsgBox "La Variedad no existe. Reintroduzca.", vbExclamation
                    txtCodigo(Index).Text = ""
                    PonerFoco txtCodigo(Index)
                End If
            End If
            
        Case 12, 13 'TIPO DE MERCADO
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "tipomer", "nomtimer", "codtimer", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            
        Case 16, 17, 10, 11 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 6060
        Me.FrameCobros.Width = 6885
        w = Me.FrameCobros.Width
        h = Me.FrameCobros.Height
    End If
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
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
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        devuelve2 = CadenaDesdeHastaBD(codD, codH, Codigo, TipCod)
        If devuelve2 = "Error" Then Exit Function
        If Not AnyadirAFormula(cadSelect, devuelve2) Then Exit Function
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
        .EnvioEMail = False
        .NombreRPT = cadNombreRPT
        .ConSubInforme = True
        .Opcion = 1
        .Show vbModal
    End With
End Sub

Private Sub AbrirFrmVariedades(Indice As Integer)
    indCodigo = Indice
    Set frmVar = New frmManVariedad
    frmVar.DatosADevolverBusqueda = "0|1|"
    frmVar.DeConsulta = True
    frmVar.CodigoActual = txtCodigo(indCodigo)
    frmVar.Show vbModal
    Set frmVar = Nothing
End Sub

Private Sub AbrirFrmMercados(Indice As Integer)
    indCodigo = Indice
    Set frmTMe = New frmManTipMerc
    frmTMe.DatosADevolverBusqueda = "0|1|"
    frmTMe.DeConsulta = True
    frmTMe.CodigoActual = txtCodigo(indCodigo)
    frmTMe.Show vbModal
    Set frmTMe = Nothing
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


Private Function HayRegistros(cTabla As String, cwhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
Dim RS As ADODB.Recordset

    SQL = "Select * FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cwhere <> "" Then
        cwhere = QuitarCaracterACadena(cwhere, "{")
        cwhere = QuitarCaracterACadena(cwhere, "}")
        cwhere = QuitarCaracterACadena(cwhere, "_1")
        SQL = SQL & " WHERE " & cwhere
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If RS.EOF Then
        MsgBox "No hay datos para mostrar en el Informe.", vbInformation
        HayRegistros = False
    Else
        HayRegistros = True
    End If

End Function

Private Function ProcesarCambios2() As Boolean
Dim SQL As String
Dim SQL1 As String
Dim Sql2 As String
Dim i As Integer
Dim HayReg As Long
Dim b As Boolean
Dim RS As ADODB.Recordset
Dim Rsx As ADODB.Recordset
Dim TotalGastos As Currency
Dim PesoCaja As Currency
Dim PesoReal As Currency
Dim ImpVenta As Currency
Dim Facturado As Byte
Dim Cobrado As Byte
Dim cadTabla2 As String

Dim Coste1 As Integer
Dim Coste2 As Integer
Dim Coste3 As Integer
Dim Coste4 As Integer

Dim Gasto1 As Currency
Dim Gasto2 As Currency
Dim Gasto3 As Currency
Dim Gasto4 As Currency
Dim Costes As Integer
'Dim GastosEnvases As Currency
'Dim GastosPortes As Currency

Dim A(100) As Currency
Dim Sql3 As String
Dim ImporteFacturado As Currency

On Error GoTo eProcesarCambios

    HayReg = 0
    
    ProcesarCambios2 = False
    
    conn.Execute "delete from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
        
    
    Label4(27).visible = True
    Label4(27).Caption = "Calculando datos de Entradas"
    
    Pb1.visible = True
        
    SQL = "insert into tmpinformes (codusu, importe1, importe2, precio1, importe3, importe4) "
    SQL = SQL & "select " & vUsu.Codigo & ", numalbar, kilosnet, prestimado, round(kilosnet * if(prestimado is null, 0, prestimado),2), 0 from rhisfruta where codvarie = " & DBSet(txtCodigo(2).Text, "N")
    If txtCodigo(10).Text <> "" Then SQL = SQL & " and fecalbar >= " & DBSet(txtCodigo(10).Text, "F")
    If txtCodigo(11).Text <> "" Then SQL = SQL & " and fecalbar <= " & DBSet(txtCodigo(11).Text, "F")
    
    conn.Execute SQL
    
    SQL1 = "select * from tmpinformes where codusu = " & vUsu.Codigo
        
    HayReg = TotalRegistrosConsulta(SQL1)
    
    If HayReg = 0 Then
        ProcesarCambios2 = True
    
        Label4(27).visible = False
        Pb1.visible = False
        Exit Function
    End If
    
    Pb1.Max = HayReg
    Pb1.Value = 0
    
    Set RS = New ADODB.Recordset
    RS.Open SQL1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    While Not RS.EOF
        IncrementarProgresNew Pb1, 1
    
        '
'        If CInt(RS!importe1) = 21541 Or CInt(RS!importe1) = 22221 Or CInt(RS!importe1) = 22241 Or CInt(RS!importe1) = 22370 Then
'            MsgBox CInt(RS!importe1), vbExclamation
'        End If
'
    
        ImporteFacturado = ImporteAlbaranFacturadoaSocio(CStr(RS!importe1))
        
        If ImporteFacturado <> 0 Then
            SQL = "update tmpinformes set importe3 = 0, importe4 = " & DBSet(ImporteFacturado, "N")
            SQL = SQL & " where codusu = " & vUsu.Codigo & " and importe1 = " & DBSet(RS!importe1, "N")
            
            conn.Execute SQL
        End If
        
        RS.MoveNext
    Wend
    
    
    ProcesarCambios2 = True

    Label4(27).visible = False
    Pb1.visible = False
    
eProcesarCambios:
    If Err.Number <> 0 Then
        ProcesarCambios2 = False
    End If
End Function

Private Function ImporteAlbaranFacturadoaSocio(NumAlbar As String) As Currency
Dim SQL As String
Dim Importe As Currency
Dim importe2 As Currency

    SQL = "select sum(importel) from rlifter where numalbar = " & DBSet(NumAlbar, "N") & " and codvarie = " & DBSet(txtCodigo(2).Text, "N")
    If txtCodigo(10).Text <> "" Then SQL = SQL & " and fechaalb >= " & DBSet(txtCodigo(10).Text, "F")
    If txtCodigo(11).Text <> "" Then SQL = SQL & " and fechaalb <= " & DBSet(txtCodigo(11).Text, "F")
    Importe = DevuelveValor(SQL)
    
    SQL = "select sum(if(importe is null,0,importe) - if(imporgasto is null,0,imporgasto)) from rfactsoc_albaran where numalbar = " & DBSet(NumAlbar, "N") & " and codvarie = " & DBSet(txtCodigo(2).Text, "N") & " and codtipom in (select codtipom from usuarios.stipom where tipodocu = 2)"
    If txtCodigo(10).Text <> "" Then SQL = SQL & " and fecalbar >= " & DBSet(txtCodigo(10).Text, "F")
    If txtCodigo(11).Text <> "" Then SQL = SQL & " and fecalbar <= " & DBSet(txtCodigo(11).Text, "F")
    
    SQL = SQL & " and not (rfactsoc_albaran.codtipom, rfactsoc_albaran.numfactu, rfactsoc_albaran.fecfactu) in (select rectif_codtipom,rectif_numfactu,rectif_fecfactu from rfactsoc "
    SQL = SQL & " where not rectif_codtipom is null and not numfactu is null and not rectif_fecfactu is null)"
    
    importe2 = DevuelveValor(SQL)
    
    '[Monica]06/11/2013: si no esta liquidado cogemos todo lo que haya anticipado
    If Not AlbaranLiquidado(NumAlbar) Then
        SQL = "select sum(if(importe is null,0,importe) - if(imporgasto is null,0,imporgasto)) from rfactsoc_albaran where numalbar = " & DBSet(NumAlbar, "N") & " and codvarie = " & DBSet(txtCodigo(2).Text, "N") & " and codtipom in (select codtipom from usuarios.stipom where tipodocu = 1)"
        
        If txtCodigo(10).Text <> "" Then SQL = SQL & " and fecalbar >= " & DBSet(txtCodigo(10).Text, "F")
        If txtCodigo(11).Text <> "" Then SQL = SQL & " and fecalbar <= " & DBSet(txtCodigo(11).Text, "F")
        
        SQL = SQL & " and not (rfactsoc_albaran.codtipom, rfactsoc_albaran.numfactu, rfactsoc_albaran.fecfactu) in (select rectif_codtipom,rectif_numfactu,rectif_fecfactu from rfactsoc "
        SQL = SQL & " where not rectif_codtipom is null and not numfactu is null and not rectif_fecfactu is null)"
        
        importe2 = DevuelveValor(SQL)
    End If

    ImporteAlbaranFacturadoaSocio = Importe + importe2
    
End Function

Private Function AlbaranLiquidado(NumAlbar As String) As Boolean
Dim SQL As String

    SQL = "select count(*) from rfactsoc_albaran where numalbar = " & DBSet(NumAlbar, "N") & " and codvarie = " & DBSet(txtCodigo(2).Text, "N") & " and codtipom in (select codtipom from usuarios.stipom where tipodocu = 2)"
    If txtCodigo(10).Text <> "" Then SQL = SQL & " and fecalbar >= " & DBSet(txtCodigo(10).Text, "F")
    If txtCodigo(11).Text <> "" Then SQL = SQL & " and fecalbar <= " & DBSet(txtCodigo(11).Text, "F")
    

    AlbaranLiquidado = (TotalRegistros(SQL) <> 0)

End Function
