VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmVtasInfGra 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6630
   Icon            =   "frmVtasInfGra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCobros 
      Height          =   5760
      Left            =   45
      TabIndex        =   5
      Top             =   0
      Width           =   6510
      Begin VB.Frame FrameOrdenar 
         Caption         =   "Agrupado Por:"
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
         Height          =   1845
         Left            =   315
         TabIndex        =   12
         Top             =   2655
         Width           =   2850
         Begin VB.OptionButton optList1 
            Caption         =   "Países"
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
            Index           =   3
            Left            =   270
            TabIndex        =   16
            Top             =   1395
            Width           =   2235
         End
         Begin VB.OptionButton optList1 
            Caption         =   "Forfaits"
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
            Left            =   270
            TabIndex        =   15
            Top             =   1035
            Width           =   2235
         End
         Begin VB.OptionButton optList1 
            Caption         =   "Variedades"
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
            Left            =   270
            TabIndex        =   14
            Top             =   315
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton optList1 
            Caption         =   "Marcas"
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
            Left            =   270
            TabIndex        =   13
            Top             =   675
            Width           =   2235
         End
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
         Left            =   1620
         MaxLength       =   3
         TabIndex        =   0
         Top             =   930
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
         Index           =   0
         Left            =   2505
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text5"
         Top             =   930
         Width           =   3810
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
         Index           =   17
         Left            =   1605
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2070
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
         Index           =   16
         Left            =   1605
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1665
         Width           =   1350
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
         Left            =   5265
         TabIndex        =   4
         Top             =   5130
         Width           =   1065
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
         Left            =   4095
         TabIndex        =   3
         Top             =   5130
         Width           =   1065
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   315
         TabIndex        =   17
         Top             =   4725
         Width           =   6060
         _ExtentX        =   10689
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Cargando tabla temporal"
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
         Index           =   27
         Left            =   360
         TabIndex        =   18
         Top             =   5040
         Width           =   3390
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Producto"
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
         Index           =   6
         Left            =   405
         TabIndex        =   11
         Top             =   735
         Width           =   885
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1350
         MouseIcon       =   "frmVtasInfGra.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar producto"
         Top             =   945
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Gráfico Informes de Ventas "
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
         TabIndex        =   9
         Top             =   225
         Width           =   4800
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Index           =   16
         Left            =   405
         TabIndex        =   8
         Top             =   1350
         Width           =   1815
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
         Left            =   630
         TabIndex        =   7
         Top             =   1665
         Width           =   690
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
         Left            =   630
         TabIndex        =   6
         Top             =   2070
         Width           =   645
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1350
         Picture         =   "frmVtasInfGra.frx":015E
         ToolTipText     =   "Buscar fecha"
         Top             =   1665
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   1350
         Picture         =   "frmVtasInfGra.frx":01E9
         ToolTipText     =   "Buscar fecha"
         Top             =   2070
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7920
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmVtasInfGra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MONICA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmPro As frmManProductos 'Productos
Attribute frmPro.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadselect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String

InicializarVbles

    If txtCodigo(0).Text = "" Then
        MsgBox "Debe introducir un Producto.", vbExclamation
        Exit Sub
    End If
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
     '======== FORMULA  ====================================
    'Seleccionar registros de la empresa conectada
'    Codigo = "{" & tabla & ".codempre}=" & vEmpresa.codEmpre
'    If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
'    If Not AnyadirAFormula(cadSelect, Codigo) Then Exit Sub
    
    'D/H Fecha albaran
    cDesde = Trim(txtCodigo(16).Text)
    cHasta = Trim(txtCodigo(17).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{albaran.fechaalb}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    cadTABLA = tabla & " INNER JOIN albaran_variedad ON albaran.numalbar = albaran_variedad.numalbar "
    cadTABLA = "(" & cadTABLA & ") INNER JOIN variedades ON albaran_variedad.codvarie = variedades.codvarie "
    cadTABLA = "(" & cadTABLA & ") INNER JOIN productos ON variedades.codprodu = productos.codprodu "
    cadTABLA = "(" & cadTABLA & ") INNER JOIN destinos ON albaran.codclien = destinos.codclien and albaran.coddesti = destinos.coddesti "
    cadTABLA = "(" & cadTABLA & ") LEFT JOIN facturas_variedad ON albaran_variedad.numalbar = facturas_variedad.numalbar and albaran_variedad.numlinea = facturas_variedad.numlinealbar "

    If Not AnyadirAFormula(cadselect, "{variedades.codprodu} = " & txtCodigo(0).Text) Then Exit Sub
    
    
    cadParam = cadParam & "pProducto= ""Producto: " & txtCodigo(0).Text & " " & txtNombre(0).Text & """|"
    numParam = numParam + 1
    
    cadFormula = "{tmpinfventas.codusu} = " & vUsu.Codigo
    
    If HayRegistros(cadTABLA, cadselect) Then
        If ProcesarCambios(cadTABLA, cadselect) Then
              'Nombre fichero .rpt a Imprimir
            cadTitulo = "Albaranes de Venta"
            If optList1(0).Value Then
                 cadNombreRPT = "rInfVtasGra1.rpt"
'                 cadParam = cadParam & "pGrupo1={albaran_variedad.codvarie}|"
'                 numParam = numParam + 1
            End If
            If optList1(1).Value Then
                 cadNombreRPT = "rInfVtasGra2.rpt"
                 cadParam = cadParam & "pGrupo1={albaran_variedad.codmarca}|"
                 numParam = numParam + 1
            End If
            If optList1(2).Value Then
                cadNombreRPT = "rInfVtasGra3.rpt"
                cadParam = cadParam & "pGrupo1={albaran_variedad.codforfait}|"
                numParam = numParam + 1
            End If
            If optList1(3).Value Then
                cadNombreRPT = "rInfVtasGra4.rpt"
                cadParam = cadParam & "pGrupo1={destinos.codpaise}|"
                numParam = numParam + 1
            End If
            LlamarImprimir
       End If
    End If
    
End Sub

Private Function ProcesarCambios(cadTABLA, cadwhere As String) As Boolean
Dim Sql As String
Dim SQL1 As String
Dim Sql2 As String
Dim i As Integer
Dim HayReg As Long
Dim b As Boolean
Dim Rs As ADODB.Recordset
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
Dim GastosEnvases As Currency
Dim GastosPortes As Currency

On Error GoTo eProcesarCambios

    HayReg = 0
    
    ProcesarCambios = False
    
    conn.Execute "delete from tmpinfventas where codusu = " & DBSet(vUsu.Codigo, "N")
        
    If cadwhere <> "" Then
        cadwhere = QuitarCaracterACadena(cadwhere, "{")
        cadwhere = QuitarCaracterACadena(cadwhere, "}")
        cadwhere = QuitarCaracterACadena(cadwhere, "_1")
    End If
        
        
    SQL1 = "select albaran.fechaalb, albaran.numalbar, albaran_variedad.numlinea, "
    SQL1 = SQL1 & "albaran_variedad.numcajas, albaran_variedad.pesoneto, albaran_variedad.preciopro, "
    SQL1 = SQL1 & "sum(facturas_variedad.impornet) from " & cadTABLA
    SQL1 = SQL1 & " where (1 = 1) "
    If cadwhere <> "" Then SQL1 = SQL1 & " and " & cadwhere
    SQL1 = SQL1 & " group by 1, 2, 3, 4, 5"
    SQL1 = SQL1 & " order by 1, 2, 3, 4, 5"
        
    Set Rs = New ADODB.Recordset
    Rs.Open SQL1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Label4(27).visible = True
    Pb1.visible = True
        
    HayReg = TotalRegistrosConsulta(SQL1)
    
    Pb1.Max = HayReg
    Pb1.Value = 0
    
    Coste1 = -1
    Coste2 = -1
    Coste3 = -1
    Coste4 = -1
    While Not Rs.EOF
        IncrementarProgresNew Pb1, 1
    
        Sql2 = "select sum(impcoste) from albaran_costes where numalbar = "
        Sql2 = Sql2 & DBSet(Rs.Fields(1).Value, "N") & " and numlinea = "
        Sql2 = Sql2 & DBSet(Rs.Fields(2).Value, "N")
        
        TotalGastos = DevuelveValor(Sql2)
        
        Sql2 = "select forfaits.kiloscaj from forfaits, albaran_variedad where "
        Sql2 = Sql2 & " albaran_variedad.numalbar = " & DBSet(Rs.Fields(1).Value, "N")
        Sql2 = Sql2 & " and albaran_variedad.numlinea = " & DBSet(Rs.Fields(2).Value, "N")
        Sql2 = Sql2 & " and albaran_variedad.codforfait = forfaits.codforfait "
        
        PesoCaja = DevuelveValor(Sql2)
        PesoReal = Round2(PesoCaja * DBLet(Rs.Fields(3).Value, "N"), 2)
        
        ImpVenta = 0
        If Not IsNull(Rs.Fields(6).Value) Then
            ImpVenta = Rs.Fields(6).Value
            Facturado = 1
            ' solo en este caso miro si esta o no cobrada en tesoreria
            
            '[Monica]16/04/2010:antes FacturaCobradaTesoreria
'            Cobrado = FacturaCobradaTesoreria(DBLet(Rs.Fields(1).Value, "N"), DBLet(Rs.Fields(2).Value, "N"))
            Cobrado = AlbaranCobradoTesoreria(DBLet(Rs.Fields(1).Value, "N"), DBLet(Rs.Fields(2).Value, "N"))
        Else
            ImpVenta = Round2(DBLet(Rs.Fields(4).Value, "N") * DBLet(Rs.Fields(5).Value, "N"), 2)
            Facturado = 0
            Cobrado = 0
        End If
        
        Gasto1 = 0
        Gasto2 = 0
        Gasto3 = 0
        Gasto4 = 0
        
        cadTabla2 = "(" & cadTABLA & ") inner join albaran_costes on albaran_variedad.numalbar = albaran_costes.numalbar "
        cadTabla2 = cadTabla2 & " and albaran_variedad.numlinea = albaran_costes.numlinea "
        
        Sql2 = "select count(distinct albaran_costes.codcoste) from " & cadTabla2
        Sql2 = Sql2 & cadwhere
        
        Costes = DevuelveValor(Sql2)
        If CCur(Costes) > 4 Then
            MsgBox "El numero de costes distintos es superior a cuatro y no cabe en el listado", vbExclamation
            ProcesarCambios = False
            Exit Function
        End If
        
        Sql2 = "select codcoste, impcoste from albaran_costes where albaran_costes.numalbar = " & DBSet(Rs.Fields(1).Value, "N")
        Sql2 = Sql2 & " and albaran_costes.numlinea = " & DBSet(Rs.Fields(2).Value, "N")
        Sql2 = Sql2 & " and albaran_costes.tipogasto = 0 "
        
        Set Rsx = New ADODB.Recordset
        Rsx.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rsx.EOF
            If Coste1 = -1 Or Coste1 = DBLet(Rsx.Fields(0).Value, "N") Then
                Coste1 = DBLet(Rsx.Fields(0).Value, "N")
                Gasto1 = DBLet(Rsx.Fields(1).Value, "N")
            Else
                If Coste2 = -1 Or Coste2 = DBLet(Rsx.Fields(0).Value, "N") Then
                    Coste2 = DBLet(Rsx.Fields(0).Value, "N")
                    Gasto2 = DBLet(Rsx.Fields(1).Value, "N")
                Else
                    If Coste3 = -1 Or Coste3 = DBLet(Rsx.Fields(0).Value, "N") Then
                        Coste3 = DBLet(Rsx.Fields(0).Value, "N")
                        Gasto3 = DBLet(Rsx.Fields(1).Value, "N")
                    Else
                        If Coste4 = -1 Or Coste4 = DBLet(Rsx.Fields(0).Value, "N") Then
                            Coste4 = DBLet(Rsx.Fields(0).Value, "N")
                            Gasto4 = DBLet(Rsx.Fields(1).Value, "N")
                        End If
                    End If
                End If
           End If
           Rsx.MoveNext
        Wend
        
        Sql2 = "select sum(impcoste) from albaran_costes where albaran_costes.numalbar = " & DBSet(Rs.Fields(1).Value, "N")
        Sql2 = Sql2 & " and albaran_costes.numlinea = " & DBSet(Rs.Fields(2).Value, "N")
        Sql2 = Sql2 & " and albaran_costes.tipogasto = 1 "
        GastosEnvases = DevuelveValor(Sql2)
        
        '[Monica] 15/06/2010 añadido costes paletizacion
        Sql2 = "select sum(impcoste) from albaran_costes where albaran_costes.numalbar = " & DBSet(Rs.Fields(1).Value, "N")
        Sql2 = Sql2 & " and albaran_costes.numlinea = " & DBSet(Rs.Fields(2).Value, "N")
        Sql2 = Sql2 & " and albaran_costes.tipogasto = 4 "
        GastosEnvases = GastosEnvases + DevuelveValor(Sql2)
        
        Sql2 = "select impcoste from albaran_costes where albaran_costes.numalbar = " & DBSet(Rs.Fields(1).Value, "N")
        Sql2 = Sql2 & " and albaran_costes.numlinea = " & DBSet(Rs.Fields(2).Value, "N")
        Sql2 = Sql2 & " and albaran_costes.tipogasto = 2 "
        GastosPortes = DevuelveValor(Sql2)
        
        
        Sql = "insert into tmpinfventas (codusu, fecalbar, numalbar, numlinea, numcajas, pesoreal, pesoneto, gastos, impventa, facturado, cobrado, "
        Sql = Sql & " codigo1, gastos1, codigo2, gastos2, codigo3, gastos3, codigo4, gastos4, gastosportes, gastosenvases) values (" & DBSet(vUsu.Codigo, "N") & ","
        Sql = Sql & DBSet(Rs.Fields(0).Value, "F") & "," & DBSet(Rs.Fields(1).Value, "N") & "," & DBSet(Rs.Fields(2).Value, "N") & ","
        Sql = Sql & DBSet(Rs.Fields(3).Value, "N") & "," 'numero de cajas
        Sql = Sql & DBSet(PesoReal, "N") & "," & DBSet(Rs.Fields(4).Value, "N") & "," 'peso neto
        Sql = Sql & DBSet(TotalGastos, "N") & "," & DBSet(ImpVenta, "N") & "," ' importe de venta
        Sql = Sql & DBSet(Facturado, "N") & ","  'facturado o no
        Sql = Sql & DBSet(Cobrado, "N") & "," 'cobrado o no
        Sql = Sql & DBSet(Coste1, "N") & "," & DBSet(Gasto1, "N") & "," 'coste1 gasto1
        Sql = Sql & DBSet(Coste2, "N") & "," & DBSet(Gasto2, "N") & "," 'coste2 gasto2
        Sql = Sql & DBSet(Coste3, "N") & "," & DBSet(Gasto3, "N") & "," 'coste3 gasto3
        Sql = Sql & DBSet(Coste4, "N") & "," & DBSet(Gasto4, "N") & "," 'coste4 gasto4
        Sql = Sql & DBSet(GastosPortes, "N") & "," ' gastos portes
        Sql = Sql & DBSet(GastosEnvases, "N") & ")" ' gastos envases
        
        conn.Execute Sql
      
        Rs.MoveNext
    Wend
    
    ProcesarCambios = True

    Label4(27).visible = False
    Pb1.visible = False
    
eProcesarCambios:
    If Err.Number <> 0 Then
        ProcesarCambios = False
    End If
End Function


Private Function ProcesarCambiosCalibres(cadTABLA, cadwhere As String) As Boolean
Dim Sql As String
Dim SQL1 As String
Dim Sql2 As String
Dim i As Integer
Dim HayReg As Long
Dim b As Boolean
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim PesoCaja As Currency
Dim PesoReal As Currency
Dim ImpVenta As Currency
Dim Facturado As Byte
Dim cadTabla2 As String
Dim Incluido As Byte ' me indica si he podido incluir el calibre

Dim VariedadAnt As Integer

Dim Calibre, Neto

On Error GoTo eProcesarCambiosCalibres

    HayReg = 0
    
    conn.Execute "delete from tmpinfventas where codusu = " & DBSet(vUsu.Codigo, "N")
        
    If cadwhere <> "" Then
        cadwhere = QuitarCaracterACadena(cadwhere, "{")
        cadwhere = QuitarCaracterACadena(cadwhere, "}")
        cadwhere = QuitarCaracterACadena(cadwhere, "_1")
    End If
        
    SQL1 = "select albaran_variedad.codvarie, albaran.fechaalb, albaran.numalbar, albaran_variedad.numlinea, "
    SQL1 = SQL1 & "albaran_variedad.numcajas, albaran_variedad.pesoneto from " & cadTABLA
    SQL1 = SQL1 & cadwhere
    SQL1 = SQL1 & " order by 1, 2, 3, 4, 5, 6"
        
    Set Rs = New ADODB.Recordset
    Rs.Open SQL1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Label4(27).visible = True
    Pb1.visible = True
        
    HayReg = TotalRegistrosConsulta(SQL1)
    Pb1.Max = HayReg
            
    If Not Rs.EOF Then
        Calibre = Array(-1, -1, -1, -1, -1, -1, -1, -1, -1)
        VariedadAnt = DBLet(Rs.Fields(0).Value, "N")
    End If
            
    While Not Rs.EOF
        If VariedadAnt <> DBLet(Rs.Fields(0).Value, "N") Then
            Calibre = Array(-1, -1, -1, -1, -1, -1, -1, -1, -1)
            VariedadAnt = DBLet(Rs.Fields(0).Value, "N")
        End If
        IncrementarProgresNew Pb1, 1
    
        Sql2 = "select codcalib, sum(pesoneto) from albaran_calibre where numalbar = "
        Sql2 = Sql2 & DBSet(Rs.Fields(2).Value, "N") & " and numlinea = "
        Sql2 = Sql2 & DBSet(Rs.Fields(3).Value, "N")
        Sql2 = Sql2 & " group by 1 "
        Sql2 = Sql2 & " order by 1 "
        
        Set Rs1 = New ADODB.Recordset
        Rs1.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
        Neto = Array(0, 0, 0, 0, 0, 0, 0, 0, 0)
        
        While Not Rs1.EOF
            Incluido = 0
            For i = 0 To 8
                If Calibre(i) = -1 Or Calibre(i) = DBLet(Rs1.Fields(0).Value, "N") Then
                    Calibre(i) = DBLet(Rs1.Fields(0).Value, "N")
                    Neto(i) = Neto(i) + DBLet(Rs1.Fields(1).Value, "N")
                    Incluido = 1
                    Exit For
                End If
            Next i
            
            Rs1.MoveNext
        Wend
        Set Rs1 = Nothing
        
        Sql = "insert into tmpinfventas (codusu, fecalbar, numalbar, numlinea, numcajas, pesoneto, "
        Sql = Sql & " calibre1, neto1, calibre2, neto2, calibre3, neto3, calibre4, neto4, calibre5, neto5, "
        Sql = Sql & " calibre6, neto6, calibre7, neto7, calibre8, neto8, calibre9, neto9, impcalibres) values (" & DBSet(vUsu.Codigo, "N") & ","
        Sql = Sql & DBSet(Rs.Fields(1).Value, "F") & "," & DBSet(Rs.Fields(2).Value, "N") & "," & DBSet(Rs.Fields(3).Value, "N") & ","
        Sql = Sql & DBSet(Rs.Fields(4).Value, "N") & "," 'numero de cajas
        Sql = Sql & DBSet(Rs.Fields(5).Value, "N") & "," 'peso neto
        Sql = Sql & DBSet(Calibre(0), "N") & "," & DBSet(Neto(0), "N") & "," ' calibre 1
        Sql = Sql & DBSet(Calibre(1), "N") & "," & DBSet(Neto(1), "N") & "," ' calibre 2
        Sql = Sql & DBSet(Calibre(2), "N") & "," & DBSet(Neto(2), "N") & "," ' calibre 3
        Sql = Sql & DBSet(Calibre(3), "N") & "," & DBSet(Neto(3), "N") & "," ' calibre 4
        Sql = Sql & DBSet(Calibre(4), "N") & "," & DBSet(Neto(4), "N") & "," ' calibre 5
        Sql = Sql & DBSet(Calibre(5), "N") & "," & DBSet(Neto(5), "N") & "," ' calibre 6
        Sql = Sql & DBSet(Calibre(6), "N") & "," & DBSet(Neto(6), "N") & "," ' calibre 7
        Sql = Sql & DBSet(Calibre(7), "N") & "," & DBSet(Neto(7), "N") & "," ' calibre 8
        Sql = Sql & DBSet(Calibre(8), "N") & "," & DBSet(Neto(8), "N") & "," ' calibre 9
        Sql = Sql & DBSet(Incluido, "N") & ")" ' si se han podido incluir todos los calibres
        
        conn.Execute Sql
      
        Rs.MoveNext
    Wend
    
    Label4(27).visible = False
    Pb1.visible = False
    
    ProcesarCambiosCalibres = HayRegistros("tmpinfventas", "codusu = " & vUsu.Codigo)
    
    Exit Function
    
eProcesarCambiosCalibres:
    If Err.Number <> 0 Then
        ProcesarCambiosCalibres = False
    End If
End Function


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(0)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
     For H = 0 To imgBuscar.Count - 1
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Next H

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, H, W
    indFrame = 5
    tabla = "albaran"
    
    Label4(27).visible = False
    Pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = W + 70
    Me.Height = H + 70
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub


Private Sub frmPro_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Variedades
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
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
    imgFec(0).Tag = Index + 16 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtCodigo(Index + 16).Text <> "" Then frmC.NovaData = txtCodigo(Index + 16).Text

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
        Case 0 'PRODUCTOS
            AbrirFrmProductos (Index)
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
            Case 0: KEYBusqueda KeyAscii, 0 'producto
            Case 16: KEYFecha KeyAscii, 0 'fecha desde
            Case 17: KEYFecha KeyAscii, 1 'fecha hasta
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
            
        Case 0  'PRODUCTOS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "productos", "nomprodu", "codprodu", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
        
        Case 16, 17 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
                        
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 5760 '5400
        Me.FrameCobros.Width = 6720
        W = Me.FrameCobros.Width
        H = Me.FrameCobros.Height
    End If
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadselect = ""
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
        .EnvioEMail = False
        .NombreRPT = cadNombreRPT
        .Opcion = 1
        .Show vbModal
    End With
End Sub

Private Sub AbrirFrmProductos(indice As Integer)
    indCodigo = indice
    Set frmPro = New frmManProductos
    frmPro.DatosADevolverBusqueda = "0|1|"
    frmPro.Show vbModal
    Set frmPro = Nothing
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


Private Function HayRegistros(cTabla As String, cWhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
Dim Rs As ADODB.Recordset

    Sql = "Select * FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Rs.EOF Then
        MsgBox "No hay datos para mostrar en el Informe.", vbInformation
        HayRegistros = False
    Else
        HayRegistros = True
    End If

End Function

