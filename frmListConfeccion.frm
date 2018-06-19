VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListConfeccion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   8160
   Icon            =   "frmListConfeccion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   8730
      Top             =   5580
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameInfConfecciones 
      Height          =   6375
      Left            =   45
      TabIndex        =   0
      Top             =   -15
      Width           =   8040
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
         Index           =   4
         Left            =   1710
         MaxLength       =   16
         TabIndex        =   4
         Top             =   2430
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
         Index           =   3
         Left            =   1710
         MaxLength       =   16
         TabIndex        =   3
         Top             =   2055
         Width           =   1350
      End
      Begin VB.Frame Frame1 
         Caption         =   "Ordenado por"
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
         Height          =   705
         Left            =   405
         TabIndex        =   18
         Top             =   4815
         Width           =   4815
         Begin VB.OptionButton Opcion 
            Caption         =   "Alfabético"
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
            Index           =   5
            Left            =   2610
            TabIndex        =   20
            Top             =   270
            Width           =   1545
         End
         Begin VB.OptionButton Opcion 
            Caption         =   "Código"
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
            Index           =   4
            Left            =   720
            TabIndex        =   19
            Top             =   270
            Width           =   1605
         End
      End
      Begin VB.Frame FrameStockMaxMin 
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
         Height          =   1830
         Left            =   405
         TabIndex        =   12
         Top             =   2925
         Width           =   4800
         Begin VB.OptionButton Opcion 
            Caption         =   "Costes por Confección detallada en línea"
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
            Index           =   6
            Left            =   135
            TabIndex        =   21
            Top             =   1170
            Width           =   4560
         End
         Begin VB.OptionButton Opcion 
            Caption         =   "Costes por Confección detallada"
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
            Left            =   135
            TabIndex        =   17
            Top             =   870
            Width           =   3900
         End
         Begin VB.OptionButton Opcion 
            Caption         =   "Confecciones completas"
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
            Left            =   135
            TabIndex        =   16
            Top             =   1470
            Width           =   3405
         End
         Begin VB.OptionButton Opcion 
            Caption         =   "Envases por Confección "
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
            Left            =   135
            TabIndex        =   14
            Top             =   270
            Width           =   3360
         End
         Begin VB.OptionButton Opcion 
            Caption         =   "Costes por Confección"
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
            Left            =   135
            TabIndex        =   13
            Top             =   570
            Width           =   3360
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Index           =   71
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Text5"
         Top             =   1485
         Width           =   4485
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Index           =   70
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "Text5"
         Top             =   1125
         Width           =   4485
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
         Index           =   71
         Left            =   1710
         MaxLength       =   16
         TabIndex        =   2
         Top             =   1485
         Width           =   1545
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
         Index           =   70
         Left            =   1710
         MaxLength       =   16
         TabIndex        =   1
         Top             =   1125
         Width           =   1545
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
         Left            =   5490
         TabIndex        =   5
         Top             =   5715
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
         Index           =   11
         Left            =   6660
         TabIndex        =   6
         Top             =   5715
         Width           =   1065
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1440
         Picture         =   "frmListConfeccion.frx":000C
         Top             =   2055
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   1440
         Picture         =   "frmListConfeccion.frx":0097
         Top             =   2430
         Width           =   240
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
         Index           =   3
         Left            =   780
         TabIndex        =   36
         Top             =   2085
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
         Index           =   2
         Left            =   780
         TabIndex        =   35
         Top             =   2445
         Width           =   645
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
         Index           =   2
         Left            =   510
         TabIndex        =   34
         Top             =   1785
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Informe de Confecciones"
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
         TabIndex        =   15
         Top             =   285
         Width           =   6735
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   28
         Left            =   1425
         Top             =   1485
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   27
         Left            =   1425
         Top             =   1125
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Confección"
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
         TabIndex        =   11
         Top             =   810
         Width           =   1095
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
         Left            =   780
         TabIndex        =   10
         Top             =   1485
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
         Left            =   780
         TabIndex        =   9
         Top             =   1125
         Width           =   690
      End
   End
   Begin VB.Frame FrameDuplicarConf 
      Height          =   4350
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   8040
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   6570
         TabIndex        =   26
         Top             =   3645
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepDuplicar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5400
         TabIndex        =   25
         Top             =   3645
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   0
         Left            =   2010
         Locked          =   -1  'True
         MaxLength       =   16
         TabIndex        =   28
         Top             =   1380
         Width           =   1455
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   1
         Left            =   1980
         MaxLength       =   16
         TabIndex        =   23
         Top             =   2175
         Width           =   1455
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   3495
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "Text5"
         Top             =   1380
         Width           =   4305
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   2
         Left            =   1980
         MaxLength       =   40
         TabIndex        =   24
         Text            =   "Text5"
         Top             =   2550
         Width           =   4305
      End
      Begin VB.Label Label3 
         Caption         =   "Descripción"
         Height          =   195
         Index           =   1
         Left            =   990
         TabIndex        =   33
         Top             =   2580
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Código"
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   32
         Top             =   2220
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nueva Confección"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   1
         Left            =   510
         TabIndex        =   31
         Top             =   1890
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Confección Origen"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   0
         Left            =   510
         TabIndex        =   30
         Top             =   1410
         Width           =   1320
      End
      Begin VB.Label Label2 
         Caption         =   "Duplicar Confección"
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
         Left            =   510
         TabIndex        =   29
         Top             =   495
         Width           =   6735
      End
   End
End
Attribute VB_Name = "frmListConfeccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public Opcionlistado As Integer
    '==== Listados BASICOS ====
    '=============================
    '0 - Listado de confecciones
    '1 - Duplicar una confeccion de origen


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

Dim CambioHorientacionPapel As Boolean ' indicamos si se va a imprimir en landscape

Dim PrimeraVez As Boolean
Dim indFrame As Single


Private Sub KEYpress(KeyAscii As Integer)
    Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub CmdAcepDuplicar_Click()

    Confeccion = ""
    If DatosOk Then
        If ProcesoDuplicarConfeccion Then
            MsgBox "Se ha realizado el proceso correctamente.", vbExclamation
            Confeccion = txtCodigo(1).Text
            cmdCancel_Click (0)
        End If
    End If
    
End Sub

Private Function DatosOk()
Dim b As Boolean
Dim SQL As String

    DatosOk = False
    b = True
    
    Select Case Opcionlistado
        Case 0 ' Informe de confecciones
        
        Case 1 ' Duplicar Confeccion
            ' Comprobamos que me hayan metido la confeccion origen y que exista
            If txtCodigo(0).Text = "" Then
                MsgBox "Debe introducir la confección de origen. Revise.", vbExclamation
                b = False
                PonerFoco txtCodigo(0)
            Else
                SQL = "select count(*) from forfaits where codforfait = " & DBSet(txtCodigo(0).Text, "T")
                If TotalRegistros(SQL) = 0 Then
                    MsgBox "No existe la confección de origen. Reintroduzca.", vbExclamation
                    b = False
                    PonerFoco txtCodigo(0)
                End If
            End If
            
            ' comprobamos que me hayan metido la nueva confeccion y que no esté ya dada de alta
            If b Then
                If txtCodigo(1).Text = "" Or txtCodigo(2).Text = "" Then
                    MsgBox "Debe introducir la nueva confección. Revise.", vbExclamation
                    b = False
                    PonerFoco txtCodigo(1)
                Else
                    SQL = "select count(*) from forfaits where codforfait = " & DBSet(txtCodigo(1).Text, "T")
                    If TotalRegistros(SQL) > 0 Then
                        MsgBox "La nueva confección ya existe. Revise.", vbExclamation
                        b = False
                        PonerFoco txtCodigo(1)
                    End If
                End If
            End If
    End Select


    DatosOk = b

End Function


Private Function ProcesoDuplicarConfeccion() As Boolean
Dim SQL As String
    
    On Error GoTo eProcesoDuplicarConfeccion

    conn.BeginTrans

    ' tabla de cabecera: forfaits
    SQL = "insert into forfaits (codforfait,nomconfe,observac,cajakilo,facturar,kiloscaj,kilosuni,codvarie,codtipen,"
    SQL = SQL & "codcapac,codmedid,codtipco,codprese,codmarca,codpalet,pesocaja,cajaspalet,preciokilonom)  "
    SQL = SQL & " select " & DBSet(txtCodigo(1).Text, "T") & "," & DBSet(txtCodigo(2).Text, "T") & ","
    SQL = SQL & " observac,cajakilo,facturar,kiloscaj,kilosuni,codvarie,codtipen,"
    SQL = SQL & " codcapac,codmedid,codtipco,codprese,codmarca,codpalet,pesocaja,cajaspalet,preciokilonom "
    SQL = SQL & " from forfaits where codforfait = " & DBSet(txtCodigo(0).Text, "T")
    
    conn.Execute SQL
    
    ' tabla de lineas envases: forfaits_envases
    SQL = "insert into forfaits_envases (codforfait,numlinea,codartic,cantidad) "
    SQL = SQL & " select " & DBSet(txtCodigo(1).Text, "T") & ","
    SQL = SQL & " numlinea,codartic,cantidad "
    SQL = SQL & " from forfaits_envases where codforfait = " & DBSet(txtCodigo(0).Text, "T")
    
    conn.Execute SQL
    
    ' tabla de lineas de costes: forfaits_costes
    SQL = "insert into forfaits_costes (codforfait,codcoste,importes) "
    SQL = SQL & " select " & DBSet(txtCodigo(1).Text, "T") & ","
    SQL = SQL & " codcoste,importes "
    SQL = SQL & " from forfaits_costes where codforfait = " & DBSet(txtCodigo(0).Text, "T")
    
    conn.Execute SQL
    
    ProcesoDuplicarConfeccion = True
    conn.CommitTrans
    Exit Function

eProcesoDuplicarConfeccion:
    conn.RollbackTrans
    MuestraError Err.Number, "Proceso Duplicar Confección", Err.Description
End Function


Private Sub cmdAceptar_Click()
'Listado de Articulos
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim campo As String
Dim Opcion As Byte, numOp As Byte

Dim indRPT As Byte
Dim nomDocu As String

    InicializarVbles
    
    cadNombreRPT = "rConfeccion.rpt"  'Nombre fichero .rpt a Imprimir
    cadTABLA = "forfaits"
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    '====================================================
    '================= FORMULA ==========================
    
    'Cadena para seleccion D/H Confeccion
    '--------------------------------------------
    cDesde = Trim(txtCodigo(70).Text)
    cHasta = Trim(txtCodigo(71).Text)
    nDesde = txtNombre(70).Text
    nHasta = txtNombre(71).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & cadTABLA & ".codforfait}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHConfeccion= """) Then Exit Sub
    End If

    
    'Obtener el parametro con el Orden del Informe
    '---------------------------------------------
    
    If Me.Opcion(4).Value = True Then numOp = PonerGrupo(1, "Forfait")
    If Me.Opcion(5).Value = True Then numOp = PonerGrupo(1, "NomForfait")
   
    'Parametro Orden del Informe
    If Me.Opcion(0).Value Then Opcion = 0
    If Me.Opcion(1).Value Then Opcion = 1
    If Me.Opcion(2).Value Then Opcion = 2
    If Me.Opcion(3).Value Then Opcion = 3
    If Me.Opcion(6).Value Then Opcion = 4
    
    CambioHorientacionPapel = (Opcion <> 0)
    
    campo = "pTipo=" & Opcion
    cadParam = cadParam & campo & "|"
    numParam = numParam + 1
    
    cadTitulo = "Listado de Confecciones"

    If Opcion = 0 Then
        cadTABLA = "forfaits_envases"
        cadselect = Replace(cadselect, "forfaits", "forfaits_envases")
    End If
    If HayRegParaInforme(cadTABLA, cadselect) Then
        '[Monica]13/12/2010: caso de costes de confeccion por linea
        If Me.Opcion(6).Value Then
            If CargarParametros Then
            
                indRPT = 108 ' Personalizacion del informe de confecciones
                If PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then
                    CambioHorientacionPapel = False
                    cadNombreRPT = nomDocu '"rConfeccion1.rpt"  'Nombre fichero .rpt a Imprimir
                    cadTABLA = "forfaits"
                    
                    If vParamAplic.Cooperativa = 0 Then
                        If txtCodigo(3).Text <> "" Then
                            cadParam = cadParam & "pDesFec=Date(" & Year(txtCodigo(3).Text) & "," & Month(txtCodigo(3).Text) & "," & Day(txtCodigo(3).Text) & ")" & "|"
                            numParam = numParam + 1
                        End If
                        If txtCodigo(4).Text <> "" Then
                            cadParam = cadParam & "pDesFec=Date(" & Year(txtCodigo(4).Text) & "," & Month(txtCodigo(4).Text) & "," & Day(txtCodigo(4).Text) & ")" & "|"
                            numParam = numParam + 1
                        End If
                    End If
                    
                End If
            End If
        End If
       LlamarImprimir
    End If
    
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub



Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        
        Select Case Opcionlistado
            Case 0
                PonerFoco txtCodigo(70)
                Me.Opcion(0).Value = True
                Me.Opcion(4).Value = True
            
            Case 1
                PonerFoco txtCodigo(1)
        
        End Select
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



    For i = 27 To 28
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i

    'Ocultar todos los Frames de Formulario
    Me.FrameInfConfecciones.visible = False
    Me.FrameDuplicarConf.visible = False
    
    '###Descomentar
'    CommitConexion
    
    Select Case Opcionlistado
    'LISTADOS DE MANTENIMIENTOS BASICOS
    '---------------------
    Case 0 '0: Listado de Confecciones
        FrameInfConfeccionesVisible True, H, W
    
    Case 1 '1: Duplicar Confeccion
        FrameDuplicarConfeccionesVisible True, H, W
        
        txtCodigo(0).Text = NumCod
        txtNombre(0).Text = DevuelveValor("select nomconfe from forfaits where codforfait = " & DBSet(txtCodigo(0).Text, "T"))
    
    End Select
    
    
    CommitConexion
    
    cadTitulo = ""
    cadNombreRPT = ""
    
'    ListadosAlmacen H, W
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(indFrame).Cancel = True
'    Me.Width = W + 70
'    Me.Height = H + 350
End Sub

Private Sub FrameInfConfeccionesVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de clientes
    Me.FrameInfConfecciones.visible = visible
    If visible = True Then
        Me.FrameInfConfecciones.Top = -90
        Me.FrameInfConfecciones.Left = 0
        Me.FrameInfConfecciones.Height = 4650
        Me.FrameInfConfecciones.Width = 8240
        W = Me.FrameInfConfecciones.Width
        H = Me.FrameInfConfecciones.Height
    End If
End Sub


Private Sub FrameDuplicarConfeccionesVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de clientes
    Me.FrameDuplicarConf.visible = visible
    If visible = True Then
        Me.FrameDuplicarConf.Top = -90
        Me.FrameDuplicarConf.Left = 0
        Me.FrameDuplicarConf.Height = 4650
        Me.FrameDuplicarConf.Width = 8240
        W = Me.FrameDuplicarConf.Width
        H = Me.FrameDuplicarConf.Height
    End If
End Sub



Private Sub frmConf_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de confecciones
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtCodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgBuscar_Click(Index As Integer)
'Buscar general: cada index llama a una tabla
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 27, 28 'cod. de confeccion
            indCodigo = Index + 43
            Set frmConf = New frmManForfaits
            frmConf.DatosADevolverBusqueda = "0|1|" 'Abrimos en Modo Busqueda
            frmConf.DeConsulta = True
            frmConf.Show vbModal
            Set frmConf = Nothing
            
        Case 0 ' Confeccion de origen
            indCodigo = Index
            Set frmConf = New frmManForfaits
            frmConf.DatosADevolverBusqueda = "0|1|" 'Abrimos en Modo Busqueda
            frmConf.DeConsulta = True
            frmConf.Show vbModal
            Set frmConf = Nothing
    
    End Select
    PonerFoco txtCodigo(indCodigo)
    Screen.MousePointer = vbDefault
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
   
   indCodigo = Index + 3
   
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
Dim Tabla As String
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
        Case 0 ' confeccion de origen
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "forfaits", "nomconfe", "codforfait", "T")
        
        Case 70, 71  'Cod. confeccion
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "forfaits", "nomconfe", "codforfait", "T")
            
        Case 3, 4 ' fechas
            If txtCodigo(Index).Text <> "" Then
                 PonerFormatoFecha txtCodigo(Index)
            End If
            
    End Select
    
End Sub


Private Sub ponerFrameConfeccionesVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el informe de Articulos, de tabla: sartic
Dim b As Boolean

    b = True
    H = 4950
    W = 8250
    
    PonerFrameVisible Me.FrameInfConfecciones, visible, H, W

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
        .EnvioEMail = False
        .NombreRPT = cadNombreRPT
        .ConSubInforme = True
        .Opcion = 0 'Opcion
        .CambioHorientacionPapel = CambioHorientacionPapel
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
        Case "Forfait"
            cadParam = cadParam & campo & "{forfaits.codforfait}" & "|"
'            If numGrupo = 1 Then
'                cadParam = cadParam & nomCampo & "|"
'            End If
            numParam = numParam + 1
            
        Case "NomForfait"
            cadParam = cadParam & campo & "{forfaits.nomconfe}" & "|"
            numParam = numParam + 1
    End Select

End Function


Private Function ComprobarFechasConta(ind As Integer) As Boolean
'comprobar que el periodo de fechas a contabilizar esta dentro del
'periodo de fechas del ejercicio de la contabilidad
Dim FechaIni As String, FechaFin As String
Dim Cad As String
Dim Rs As ADODB.Recordset
    
On Error GoTo EComprobar

    ComprobarFechasConta = False
    
    If txtCodigo(ind).Text <> "" Then
        FechaIni = "Select fechaini,fechafin From parametros"
        Set Rs = New ADODB.Recordset
        Rs.Open FechaIni, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        If Not Rs.EOF Then
            FechaIni = DBLet(Rs!FechaIni, "F")
            FechaFin = DBLet(Rs!FechaFin, "F") + 365
            'nos guardamos los valores
            Orden1 = FechaIni
            Orden2 = FechaFin
        
            If Not EntreFechas(FechaIni, txtCodigo(ind).Text, FechaFin) Then
                 Cad = "El período de contabilización debe estar dentro del ejercicio:" & vbCrLf & vbCrLf
                 Cad = Cad & "    Desde: " & FechaIni & vbCrLf
                 Cad = Cad & "    Hasta: " & FechaFin
                 MsgBox Cad, vbExclamation
                 txtCodigo(ind).Text = ""
            Else
                ComprobarFechasConta = True
            End If
        End If
        Rs.Close
        Set Rs = Nothing
    Else
        ComprobarFechasConta = True
    End If
    
EComprobar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Fechas", Err.Description
End Function

Private Sub ListadosAlmacen(H As Integer, W As Integer)
   'Listado de Artículo
    ponerFrameConfeccionesVisible True, H, W
    Codigo = "{sartic"
    indFrame = 11
    cadTitulo = "Listado de Confecciones"
End Sub

Private Function CargarParametros() As Boolean
Dim SQL As String
Dim Sql2 As String
Dim SqlValues As String
Dim Costes As Integer
Dim Coste1 As Integer
Dim Coste2 As Integer
Dim Coste3 As Integer
Dim Coste4 As Integer
Dim Coste5 As Integer
Dim nCoste1 As String
Dim nCoste2 As String
Dim nCoste3 As String
Dim nCoste4 As String
Dim nCoste5 As String


Dim Rsx As ADODB.Recordset


    On Error GoTo eCargarParametros

    CargarParametros = False

    
    Sql2 = "select count(distinct nombcoste.codcoste) from nombcoste"
    
    Costes = DevuelveValor(Sql2)
    If CCur(Costes) > 5 Then
        MsgBox "El numero de costes distintos es superior a cinco y no cabe en el listado", vbExclamation
        CargarParametros = False
        Exit Function
    End If
    
    Sql2 = "select codcoste, denominacion from nombcoste "
    
    
    Set Rsx = New ADODB.Recordset
    Rsx.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Coste1 = -1
    Coste2 = -1
    Coste3 = -1
    Coste4 = -1
    Coste5 = -1
    nCoste1 = ""
    nCoste2 = ""
    nCoste3 = ""
    nCoste4 = ""
    nCoste5 = ""
    
    
    While Not Rsx.EOF
        If Coste1 = -1 Or Coste1 = DBLet(Rsx.Fields(0).Value, "N") Then
            Coste1 = DBLet(Rsx.Fields(0).Value, "N")
            nCoste1 = DBLet(Rsx.Fields(1).Value, "T")
        Else
            If Coste2 = -1 Or Coste2 = DBLet(Rsx.Fields(0).Value, "N") Then
                Coste2 = DBLet(Rsx.Fields(0).Value, "N")
                nCoste2 = DBLet(Rsx.Fields(1).Value, "T")
            Else
                If Coste3 = -1 Or Coste3 = DBLet(Rsx.Fields(0).Value, "N") Then
                    Coste3 = DBLet(Rsx.Fields(0).Value, "N")
                    nCoste3 = DBLet(Rsx.Fields(1).Value, "T")
                Else
                    If Coste4 = -1 Or Coste4 = DBLet(Rsx.Fields(0).Value, "N") Then
                        Coste4 = DBLet(Rsx.Fields(0).Value, "N")
                        nCoste4 = DBLet(Rsx.Fields(1).Value, "T")
                    Else
                        If Coste5 = -1 Or Coste5 = DBLet(Rsx.Fields(0).Value, "N") Then
                            Coste5 = DBLet(Rsx.Fields(0).Value, "N")
                            nCoste5 = DBLet(Rsx.Fields(1).Value, "T")
                        End If
                    End If
                End If
            End If
       End If
       Rsx.MoveNext
    Wend
    
    cadParam = cadParam & "pCoste1=" & Coste1 & "|"
    numParam = numParam + 1
    
    cadParam = cadParam & "pCoste2=" & Coste2 & "|"
    numParam = numParam + 1
    
    cadParam = cadParam & "pCoste3=" & Coste3 & "|"
    numParam = numParam + 1
    
    cadParam = cadParam & "pCoste4=" & Coste4 & "|"
    numParam = numParam + 1
    
    cadParam = cadParam & "pCoste5=" & Coste5 & "|"
    numParam = numParam + 1
    
    cadParam = cadParam & "nCoste1=""" & nCoste1 & """|"
    numParam = numParam + 1
    
    cadParam = cadParam & "nCoste2=""" & nCoste2 & """|"
    numParam = numParam + 1
    
    cadParam = cadParam & "nCoste3=""" & nCoste3 & """|"
    numParam = numParam + 1
    
    cadParam = cadParam & "nCoste4=""" & nCoste4 & """|"
    numParam = numParam + 1
    
    cadParam = cadParam & "nCoste5=""" & nCoste5 & """|"
    numParam = numParam + 1

    CargarParametros = True
    Exit Function
    
eCargarParametros:
    MuestraError Err.Number, "Cargar Parámetros", Err.Description
End Function

